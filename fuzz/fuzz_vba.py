#!/usr/bin/env python3
"""
visi vs. Microsoft Excel Differential Fuzzing: VBA execution
============================================================
Phase 1 of `docs/vba-macro-support.md`. Generates a VBA procedure, runs it in
`visi`'s interpreter and in real Excel, and compares three things:

    the returned value    as CStr() renders it
    the returned subtype  as TypeName() reports it
    the error number      as Err.Number reports it, when it raised

The subtype is not decoration. `1 + 1` is an `Integer` and `1 / 1` is a
`Double`; an interpreter that computes the right number with the wrong
subtype has a real bug that only shows up later, when something overflows
that should not have. `Err.Number` is likewise compared exactly -- it is a
small enumerated set, and "which error" is as much a result as "what value".

**Why the generated code is wrapped.** The procedure under test is invoked
through a `Harness` function whose `On Error GoTo` turns any runtime error
into a returned string. Without it an untrapped error pops a modal dialog
that `set display alerts to false` does not suppress, `osascript` never
returns, and Excel has to be SIGKILLed -- so a fuzzer generating random VBA
would stall on most iterations. This was established in issue #46 and the
wrapper is verified to catch errors from arbitrary call depth.

**Why the generator emits a typed AST rather than text.** The lesson already
recorded in this directory's README: a generator that splices strings
produces files Excel silently rejects, and the failure gets attributed to the
engine. Generating from a small typed grammar guarantees by construction that
every variable is `Dim`'d, no name collides with a VBA keyword, every loop
terminates, and no expression can recurse without bound.

**What is deliberately not generated.** Anything non-deterministic (`Now`,
`Rnd`, `Timer`) -- the two engines would differ by construction and every
iteration would be noise. Anything touching the workbook, since Phase 1 has
no host object model and both sides would simply agree that it is
unsupported. And division by an expression that could be zero is *allowed*:
error 11 is a result worth comparing.

Usage:
    python3 fuzz/fuzz_vba.py --driver mock --iterations 50   # visi only, no Excel
    python3 fuzz/fuzz_vba.py --iterations 20 --seed 1
    python3 fuzz/fuzz_vba.py --iterations 100 --batch 25     # 25 cases per Excel round trip
"""

import argparse
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

EXCEL_APP = "Microsoft Excel"

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required: source fuzz/venv/bin/activate && pip install -r fuzz/requirements.txt")

try:
    import visi_core
except ImportError:
    sys.exit(
        "the visi_core bindings are required: "
        "maturin develop -m visi-python/Cargo.toml --release"
    )


# -----------------------------------------------------------------------------
# 1. Generation
# -----------------------------------------------------------------------------

# Variable names chosen to avoid VBA's keywords entirely -- `name`, `date`,
# `line`, `error` and friends are all keywords in some position.
VAR_NAMES = ["va", "vb", "vc", "vd", "ve"]

# Literals spanning every subtype boundary the Variant model has to get right:
# the Integer/Long edge at 32767, the Long/Double edge, and the suffixes.
NUM_LITERALS = [
    "0", "1", "2", "3", "7", "10", "-1", "-7", "255",
    "32767", "32768", "-32768", "100000", "2147483647",
    "1.5", "-2.5", "0.1", "3.75", "1E3", "0.0001",
    "1&", "1%", "2!", "3#", "&HFF", "&O17",
]
STR_LITERALS = ['""', '"a"', '"abc"', '"1"', '"12"', '"1.5"', '"  3  "', '"Z"']
BOOL_LITERALS = ["True", "False"]
SPECIAL_LITERALS = ["Empty", "Null"]

ARITH_OPS = ["+", "-", "*", "/", "\\", "Mod", "^"]
COMPARE_OPS = ["=", "<>", "<", ">", "<=", ">="]
LOGICAL_OPS = ["And", "Or", "Xor", "Eqv", "Imp"]

# One-argument intrinsics that are total enough to be worth generating.
#
# `Len` is absent on purpose. `Len(False)` -- Len applied to a *Boolean
# literal* -- is a compile error in Excel, which surfaces as the modal-dialog
# hang and poisons the whole batch. It is another instance of the Phase 0
# boundary: a check that needs the argument's static type, which the parser
# does not track. `Len` is still covered below, wrapped in `CStr` so its
# argument is always a string.
UNARY_FUNCS = [
    "CStr", "CInt", "CLng", "CDbl", "CSng", "CBool", "Abs", "Sgn", "Int",
    "Fix", "UCase", "LCase", "Trim", "TypeName", "IsNumeric",
    "IsEmpty", "IsNull", "Val", "StrReverse",
]


class VbaGenerator:
    """Generates VBA expressions and statements from a small typed grammar.

    `depth` bounds expression nesting, which is what keeps generated source
    from growing without limit and keeps the two engines comparing the same
    thing rather than one of them giving up.
    """

    def __init__(self, seed=None):
        self.rng = random.Random(seed)

    # ---- expressions ----------------------------------------------------

    def literal(self):
        bucket = self.rng.random()
        if bucket < 0.55:
            return self.rng.choice(NUM_LITERALS)
        if bucket < 0.78:
            return self.rng.choice(STR_LITERALS)
        if bucket < 0.92:
            return self.rng.choice(BOOL_LITERALS)
        return self.rng.choice(SPECIAL_LITERALS)

    def expr(self, depth, vars_in_scope):
        if depth <= 0:
            if vars_in_scope and self.rng.random() < 0.35:
                return self.rng.choice(vars_in_scope)
            return self.literal()

        kind = self.rng.random()
        if kind < 0.34:
            op = self.rng.choice(ARITH_OPS)
            return f"({self.expr(depth - 1, vars_in_scope)} {op} {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.46:
            op = self.rng.choice(COMPARE_OPS)
            return f"({self.expr(depth - 1, vars_in_scope)} {op} {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.56:
            op = self.rng.choice(LOGICAL_OPS)
            return f"({self.expr(depth - 1, vars_in_scope)} {op} {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.64:
            return f"({self.expr(depth - 1, vars_in_scope)} & {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.72:
            return f"(Not {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.80:
            return f"(-{self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.90:
            fn = self.rng.choice(UNARY_FUNCS)
            return f"{fn}({self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.94:
            # See UNARY_FUNCS: Len needs a string argument to stay compilable.
            return f"Len(CStr({self.expr(depth - 1, vars_in_scope)}))"
        # Two-argument string intrinsics, which need a sane count argument.
        fn = self.rng.choice(["Left", "Right", "String", "Space", "InStr", "Mid"])
        if fn in ("Space",):
            return f"Space({self.rng.randint(0, 4)})"
        if fn == "String":
            return f'String({self.rng.randint(0, 4)}, "x")'
        if fn == "InStr":
            return f"InStr({self.expr(depth - 1, vars_in_scope)}, {self.rng.choice(STR_LITERALS)})"
        if fn == "Mid":
            return f"Mid({self.expr(depth - 1, vars_in_scope)}, {self.rng.randint(1, 4)}, {self.rng.randint(0, 4)})"
        return f"{fn}({self.expr(depth - 1, vars_in_scope)}, {self.rng.randint(0, 4)})"

    # ---- statements -----------------------------------------------------

    def statements(self, count, vars_in_scope, depth):
        out = []
        for _ in range(count):
            out.extend(self.statement(vars_in_scope, depth))
        return out

    def statement(self, vars_in_scope, depth):
        kind = self.rng.random()
        target = self.rng.choice(vars_in_scope)

        if kind < 0.40:
            return [f"{target} = {self.expr(depth, vars_in_scope)}"]

        if kind < 0.55:
            return [
                f"If {self.expr(depth - 1, vars_in_scope)} Then",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "Else",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "End If",
            ]

        if kind < 0.70:
            # Bounds are literal so the loop provably terminates; a generated
            # expression could produce a limit that never ends.
            lo, hi = 1, self.rng.randint(1, 4)
            step = self.rng.choice(["", " Step 2", " Step -1"])
            if step == " Step -1":
                lo, hi = hi, 1
            return [
                f"For vi = {lo} To {hi}{step}",
                f"    {target} = {self.expr(depth - 1, vars_in_scope + ['vi'])}",
                "Next vi",
            ]

        if kind < 0.80:
            return [
                "vn = 0",
                f"Do While vn < {self.rng.randint(1, 3)}",
                "    vn = vn + 1",
                f"    {target} = {self.expr(depth - 1, vars_in_scope + ['vn'])}",
                "Loop",
            ]

        if kind < 0.90:
            subject = self.expr(depth - 1, vars_in_scope)
            return [
                f"Select Case {subject}",
                "Case 0, 1",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "Case 2 To 5",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "Case Else",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "End Select",
            ]

        # A locally-trapped error, so `Err.Number` gets compared through the
        # path a real macro uses rather than only through the outer harness.
        return [
            "On Error Resume Next",
            f"{target} = {self.expr(depth, vars_in_scope)}",
            "On Error GoTo 0",
        ]

    def module(self, index):
        depth = self.rng.randint(1, 3)
        n_vars = self.rng.randint(2, len(VAR_NAMES))
        vars_in_scope = VAR_NAMES[:n_vars]
        body = self.statements(self.rng.randint(1, 4), vars_in_scope, depth)
        result = self.rng.choice(vars_in_scope)

        lines = [f"Private Function Gen{index}()"]
        lines.append("    Dim " + ", ".join(vars_in_scope) + ", vi, vn")
        for v in vars_in_scope:
            lines.append(f"    {v} = {self.rng.choice(NUM_LITERALS)}")
        lines.extend("    " + s for s in body)
        lines.append(f"    Gen{index} = {result}")
        lines.append("End Function")
        return "\n".join(lines)


HARNESS_TEMPLATE = """Public Function Harness{i}() As String
    On Error GoTo Failed
    Dim r
    r = Gen{i}()
    Harness{i} = "OK|" & TypeName(r) & "|" & CStr(r)
    Exit Function
Failed:
    Harness{i} = "ERR|" & CStr(Err.Number)
End Function"""


def build_module(cases):
    """One module holding every case in a batch, plus its harnesses.

    Batching matters: the AppleScript round trip dominates the cost by three
    orders of magnitude, so 25 cases in one workbook run in roughly the time
    one case would.
    """
    parts = ['Attribute VB_Name = "M"']
    for i, src in cases:
        parts.append(src)
        parts.append(HARNESS_TEMPLATE.format(i=i))
    return "\n\n".join(parts) + "\n"


# -----------------------------------------------------------------------------
# 2. The two engines
# -----------------------------------------------------------------------------


def visi_result(source, proc):
    """`OK|TypeName|CStr` or `ERR|number`, the same shape the harness returns."""
    try:
        type_name, value = visi_core.run_macro(source, proc)
    except visi_core.VbaRuntimeError as e:
        return f"ERR|{getattr(e, 'number', '?')}"
    except visi_core.VbaSyntaxError as e:
        return f"SYNTAX|{e}"
    except visi_core.VisiError as e:
        return f"SYNTAX|{type(e).__name__}: {e}"
    if value is None:
        # CStr(Null) is itself error 94 in VBA, which is what the harness
        # would have reported.
        return "ERR|94"
    return f"OK|{type_name}|{value}"


class ExcelDriver:
    def __init__(self, excel_path=None, driver_type="auto", timeout=60):
        self.excel_path = excel_path
        self.timeout = timeout
        self.driver_type = driver_type
        if driver_type == "auto":
            self.driver_type = "applescript" if sys.platform == "darwin" else "mock"
        self.restarts = 0

    def app_name(self):
        name = self.excel_path or EXCEL_APP
        if name.endswith(".app"):
            name = os.path.splitext(os.path.basename(name))[0]
        return name

    def restart(self):
        self.restarts += 1
        subprocess.run(["killall", EXCEL_APP], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.0)
        pgrep = subprocess.run(["pgrep", "-x", EXCEL_APP], stdout=subprocess.PIPE, text=True)
        for pid in pgrep.stdout.split():
            subprocess.run(["kill", "-9", pid], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.0)
        subprocess.run(["open", "-a", EXCEL_APP], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(4.0)

    def run_batch(self, xlsm, indices):
        """Returns {index: result-string}, or {} if Excel could not be asked."""
        if self.driver_type == "mock":
            return {}
        # `linefeed`, not "\n": AppleScript string literals do not carry a
        # newline escape, and the whole batch silently returned nothing until
        # this was fixed.
        calls = "\n".join(
            f'    set acc to acc & "{i}=" & (run VB macro "Harness{i}") & linefeed'
            for i in indices
        )
        script = "\n".join([
            f'tell application "{self.app_name()}"',
            "    set display alerts to false",
            "    try",
            "        close workbooks saving no",
            "    end try",
            f'    open POSIX file "{os.path.abspath(xlsm)}"',
            "    set wb to active workbook",
            '    set acc to ""',
            calls,
            "    close wb saving no",
            "    return acc",
            "end tell",
        ])
        for attempt in range(2):
            try:
                res = subprocess.run(
                    ["osascript", "-e", script],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, timeout=self.timeout,
                )
            except subprocess.TimeoutExpired:
                # A hang here is a compile error in generated source (Excel
                # goes modal) or session degradation. Either way the batch is
                # lost; restart and retry once.
                self.restart()
                if attempt == 0:
                    continue
                return {}
            if res.returncode == 0:
                out = {}
                for line in res.stdout.splitlines():
                    if "=" in line:
                        k, _, v = line.partition("=")
                        if k.strip().isdigit():
                            out[int(k.strip())] = v.strip()
                return out
            self.restart()
            if attempt == 0:
                continue
            return {}
        return {}


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--excel-path")
    ap.add_argument("--driver", choices=["auto", "applescript", "mock"], default="auto")
    ap.add_argument("--iterations", type=int, default=20)
    ap.add_argument("--batch", type=int, default=20,
                    help="Cases per Excel round trip (default 20). The round trip dominates cost.")
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument("--output-dir", default="./fuzz_results")
    args = ap.parse_args()

    gen = VbaGenerator(args.seed)
    driver = ExcelDriver(args.excel_path, args.driver, args.timeout)
    failures_dir = os.path.join(args.output_dir, "failures")
    os.makedirs(failures_dir, exist_ok=True)

    print("=" * 69)
    print("    visi vs. Microsoft Excel VBA Execution Differential Fuzzer    ".center(69))
    print("=" * 69)
    print(f" Cases       : {args.iterations} in batches of {args.batch}")
    print(f" Excel driver: {driver.driver_type} ({args.excel_path or 'default'})")
    if driver.driver_type == "mock":
        print(" MOCK DRIVER -- visi runs, Excel is not consulted, nothing is compared.")
    print("=" * 69 + "\n")

    passed = failed = skipped = 0
    workdir = tempfile.mkdtemp(prefix="vba_exec_fuzz_")
    start = time.time()

    try:
        for batch_start in range(0, args.iterations, args.batch):
            n = min(args.batch, args.iterations - batch_start)
            cases = [(batch_start + i + 1, gen.module(batch_start + i + 1)) for i in range(n)]
            source = build_module(cases)
            indices = [i for i, _ in cases]

            excel = {}
            if driver.driver_type != "mock":
                base = os.path.join(workdir, "base.xlsx")
                xlsm = os.path.join(workdir, f"batch_{batch_start}.xlsm")
                openpyxl.Workbook().save(base)
                wb = visi_core.Workbook.load(base)
                wb.add_macro("M", source)
                wb.save(xlsm)
                excel = driver.run_batch(xlsm, indices)

            for i, _ in cases:
                mine = visi_result(source, f"Gen{i}")
                theirs = excel.get(i)
                if theirs is None:
                    skipped += 1
                    continue
                if mine == theirs:
                    passed += 1
                    continue
                failed += 1
                print(f" case {i:<5} [MISMATCH]")
                print(f"   visi : {mine}")
                print(f"   excel: {theirs}")
                out = os.path.join(failures_dir, f"vba_exec_case_{i}")
                os.makedirs(out, exist_ok=True)
                with open(os.path.join(out, "source.bas"), "w") as f:
                    f.write(source)
                with open(os.path.join(out, "verdicts.txt"), "w") as f:
                    f.write(f"procedure: Gen{i}\nvisi:  {mine}\nexcel: {theirs}\n")
                print(f"   saved: {out}")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    total = passed + failed
    print("\n" + "=" * 69)
    print(f" Completed in {time.time() - start:.1f}s ({driver.restarts} Excel restarts)")
    print(f" Agreed   : {passed}/{total}" if total else " Agreed   : n/a")
    print(f" Mismatch : {failed}")
    if skipped:
        print(f" Skipped  : {skipped} (Excel gave no answer)")
    print("=" * 69)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
