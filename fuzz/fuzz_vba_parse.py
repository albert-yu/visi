#!/usr/bin/env python3
"""
visi vs. Microsoft Excel Differential Fuzzing: VBA syntax
=========================================================
Asks one question, of both engines, about generated VBA source: **does this
compile?** `visi_core.check_syntax` answers it by parsing (Phase 0 of
`docs/vba-macro-support.md`); real Excel answers it by being asked to run a
procedure containing the source. Where they disagree, one of them is wrong.

There is no interpreter behind `check_syntax` yet, so nothing here compares
*values* -- that is Phase 1's harness. What this catches is the two failure
modes a syntax checker has:

    visi rejects, Excel accepts   a false positive -- `visi macro check`
                                  calls working code broken. The worse of
                                  the two: it makes the command untrustworthy.
    visi accepts, Excel rejects   a false negative -- `macro check` passes
                                  source Excel will refuse to compile, which
                                  is exactly what the command exists to catch.

Getting Excel's verdict is the whole difficulty, and three empirical findings
shape the design (each verified against Excel 16.112, macOS 26.6.1):

1. **A compile error is only observable as a hang.** Excel raises it as a
   modal dialog that `set display alerts to false` does not suppress; the
   `osascript` call never returns and Excel has to be SIGKILLed. So "Excel
   rejected it" is read from a timeout, not from an error message. The
   `On Error` wrapper that makes *runtime* errors reportable does not help --
   a compile error is not a trappable runtime error.

2. **VBA compiles lazily, per procedure -- strictly.** A first attempt put a
   trivial `ProbeCompile` procedure in the same module as broken code and
   called that; it returned "ok" happily, because Excel never compiled the
   broken procedure. Nor does *referencing* the broken procedure from a dead
   branch of the one being called force it. Only invoking a procedure
   compiles it, which is why `--corpus` below cannot get an Excel verdict for
   an arbitrary module: there is no way to compile its procedures without
   running them.

3. **`If False Then ... End If` compiles its body without running it.** This
   is what makes (2) usable: the generated source goes inside a dead branch
   of the procedure being called, so Excel must compile it, and cannot
   execute it. Verified with an `Err.Raise 5` inside the dead branch, which
   returned "OK" rather than the trapped error -- proof the body compiled and
   did not run.

Together: the harness never executes a line of generated VBA. That removes
runtime errors, infinite loops, and side effects from the picture entirely,
leaving only the compile verdict -- which is all Phase 0 is about.

Cost note: a *valid* case costs one fast AppleScript round trip (~1s). An
*invalid* one costs the full timeout plus an Excel restart (~15s), since a
hang is the signal. Runs are therefore dominated by however many invalid
cases the generator produces; `--timeout` trades confidence for speed.

Usage:
    python3 fuzz/fuzz_vba_parse.py --driver mock --iterations 20   # no Excel; parser-only smoke test
    python3 fuzz/fuzz_vba_parse.py --iterations 10 --seed 1
    python3 fuzz/fuzz_vba_parse.py --corpus visi-core/fuzz/seeds/vba_parse
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

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
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

# Fragments that are valid VBA *statements*, safe to drop inside a dead branch.
VALID_FRAGMENTS = [
    "Dim x As Long",
    "Dim s As String, t As Double",
    "x = 1 + 2 * 3",
    "x = -2 ^ 2",
    "x = 2 ^ 3 ^ 2",
    "s = \"a\" & \"b\" & CStr(1)",
    "x = 10 \\ 3 Mod 2",
    "b = Not 1 = 0 And True",
    "b = True Xor False Eqv True Imp False",
    "If x > 1 Then x = 2 Else x = 3",
    "If x > 1 Then\n    x = 2\nElseIf x = 1 Then\n    x = 3\nElse\n    x = 4\nEnd If",
    "For i = 1 To 10 Step 2\n    x = i\nNext i",
    "For Each c In rng\n    x = 1\nNext",
    "Do While x < 10\n    x = x + 1\nLoop",
    "Do\n    x = x + 1\nLoop Until x > 10",
    "While x < 3\n    x = x + 1\nWend",
    "Select Case x\nCase 1, 2\n    y = 1\nCase 3 To 5\n    y = 2\nCase Is >= 6\n    y = 3\nCase Else\n    y = 4\nEnd Select",
    "With obj\n    .a = .b(1)\nEnd With",
    "Set o = New Collection",
    "o.Method arg1, arg2",
    "o.Method Key:=1, Value:=2",
    "x = Helper(1, , 3)",
    "x = arr(1, 2)",
    "ReDim Preserve arr(1 To 5)",
    "Erase arr",
    "On Error Resume Next",
    "x = ws.Range(\"A1\").Value",
    "x = rs!Field",
    "d = #1/1/2000#",
    "x = &HFF + &O17 + 1.5e-3",
    "x = a _\n    + b",
    "x = 1: y = 2",
    "Static counter As Long",
    "b = TypeOf o Is Worksheet",
]

# Tokens a mutation can splice in. Most produce invalid source; a few do not,
# which is the point -- the harness must not assume mutation implies invalid.
MUTATION_TOKENS = [
    "Then", "End", "Next", "Loop", "Wend", "Case", "Else", "As", "To", "In",
    "=", "(", ")", ",", "&", "^", ".", ":", "\"", "#", "_", "Mod", "Not",
    "Dim", "Sub", "Function", "1", "x",
]


class VbaSourceGenerator:
    """Builds VBA statement bodies, valid by construction, then optionally
    mutates them.

    Generating from fragments rather than raw characters is deliberate: the
    interesting disagreements live near the boundary of the grammar, not out
    in random bytes, and a generator that mostly emits garbage would spend
    every iteration paying the invalid-case timeout for no signal.
    """

    def __init__(self, seed=None):
        self.rng = random.Random(seed)

    def body(self, mutate):
        lines = [self.rng.choice(VALID_FRAGMENTS) for _ in range(self.rng.randint(1, 4))]
        src = "\n".join(lines)
        if mutate:
            src = self.mutate(src)
        return src

    def mutate(self, src):
        """One structural edit, chosen to be the kind a human typo produces."""
        toks = src.split(" ")
        if not toks:
            return src
        kind = self.rng.choice(["delete", "insert", "duplicate", "swap", "truncate"])
        i = self.rng.randrange(len(toks))
        if kind == "delete":
            del toks[i]
        elif kind == "insert":
            toks.insert(i, self.rng.choice(MUTATION_TOKENS))
        elif kind == "duplicate":
            toks.insert(i, toks[i])
        elif kind == "swap" and len(toks) > 1:
            j = self.rng.randrange(len(toks))
            toks[i], toks[j] = toks[j], toks[i]
        elif kind == "truncate":
            toks = toks[: max(1, i)]
        return " ".join(toks)


# The module template. `Gen`'s body never runs -- see this file's docstring
# for why the dead branch is load-bearing rather than decorative.
#
# `Helper` exists because of the first real disagreement this harness found.
# `x = f(1, , 3)` -- an omitted middle argument -- parses fine and *is* valid
# VBA, but Excel rejects it at compile time when it cannot resolve `f` to a
# procedure with an `Optional` parameter in that position. That is a semantic
# check requiring name resolution, which Phase 0 deliberately does not do (a
# `Call` node cannot even tell a procedure call from an array index without a
# symbol table). So the divergence is a boundary of what parse-only checking
# can see, not a parser bug -- and declaring a real callee keeps the
# omitted-argument syntax under test instead of dropping the coverage.
MODULE_TEMPLATE = """Attribute VB_Name = "M"
Private Function Helper(Optional a, Optional b, Optional c)
End Function

Public Sub Gen({sig})
    If False Then
{body}
    End If
End Sub

Public Function Harness() As String
    On Error GoTo Failed
    Gen{args}
    Harness = "OK"
    Exit Function
Failed:
    Harness = "ERR|" & CStr(Err.Number)
End Function
"""


def build_module(body, sig="", args=""):
    """The module source wrapping one statement body.

    `sig`/`args` give `Gen` a parameter list: `build_module(src, "ByVal x
    As Long", " 1")` declares `Public Sub Gen(ByVal x As Long)` and calls
    it as `Gen 1`. Both default to empty, the parameterless wrapper
    everything else uses. `args` is spliced straight after `Gen`, so it
    carries its own leading space.

    They come as a pair because the dead branch is not optional: Excel
    compiles a procedure only when something invokes it (point 2 above),
    so `Gen` has to stay reachable from `Harness` however it is declared.
    A caller supplying a whole procedure of its own would have to
    re-establish that, which is why the signature is threaded through here
    instead.
    """
    indented = "\n".join("        " + line for line in body.split("\n"))
    return MODULE_TEMPLATE.format(body=indented, sig=sig, args=args)


# -----------------------------------------------------------------------------
# 2. The two verdicts
# -----------------------------------------------------------------------------


def visi_verdict(source):
    """(accepted, detail) from visi's parser."""
    try:
        visi_core.check_syntax(source)
        return True, ""
    except visi_core.VbaSyntaxError as e:
        return False, str(e)
    except visi_core.VisiError as e:
        return False, f"{type(e).__name__}: {e}"


# The child process `ExcelVerdictDriver._win32com_verdict` launches (see
# that method's docstring for why an in-process win32com call can't be
# timed out directly). `argv[1]` is the .xlsm path; prints the `Harness`
# result ("OK" or "ERR|n") to stdout on success. A compile error leaves
# Excel hung showing the modal dialog, so nothing is printed and the
# process itself hangs -- observable only as the parent's subprocess
# timeout firing, the exact same signal the AppleScript path reads.
_WIN32COM_VERDICT_RUNNER = """
import sys
import win32com.client

xlsm_path = sys.argv[1]

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1
try:
    wb = excel.Workbooks.Open(xlsm_path)
    try:
        result = excel.Run("Harness")
        print(result)
    finally:
        wb.Close(False)
finally:
    excel.Quit()
"""


class ExcelVerdictDriver:
    """Excel's verdict, read from whether `run VB macro "Harness"` returns.

    Returning at all -- with "OK" or a trapped runtime "ERR|n" -- means the
    module compiled. Only a timeout means it did not.
    """

    def __init__(self, excel_path=None, driver_type="auto", timeout=15):
        self.excel_path = excel_path
        self.timeout = timeout
        self.driver_type = driver_type
        if driver_type == "auto":
            if sys.platform == "darwin":
                self.driver_type = "applescript"
            elif sys.platform == "win32":
                self.driver_type = "win32com"
            else:
                self.driver_type = "mock"
        self.restarts = 0

    def app_name(self):
        name = self.excel_path or EXCEL_APP
        if name.endswith(".app"):
            name = os.path.splitext(os.path.basename(name))[0]
        return name

    def script(self, path):
        return "\n".join([
            f'tell application "{self.app_name()}"',
            "    set display alerts to false",
            "    try",
            "        close workbooks saving no",
            "    end try",
            f'    open POSIX file "{path}"',
            "    set wb to active workbook",
            '    set r to run VB macro "Harness"',
            "    close wb saving no",
            "    return r",
            "end tell",
        ])

    def restart_excel(self):
        """SIGKILL by PID -- `killall` alone can leave Excel running, since it
        may intercept SIGTERM to run its own quit handshake (see
        fuzz_pivot.py::_restart_excel, where this was first needed)."""
        self.restarts += 1
        subprocess.run(["killall", EXCEL_APP], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.0)
        pgrep = subprocess.run(["pgrep", "-x", EXCEL_APP], stdout=subprocess.PIPE, text=True)
        for pid in pgrep.stdout.split():
            subprocess.run(["kill", "-9", pid], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1.0)
        subprocess.run(["open", "-a", EXCEL_APP], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(4.0)

    def restart_windows(self):
        """`taskkill` every EXCEL.EXE. Nothing to relaunch -- the next
        verdict's `gencache.EnsureDispatch("Excel.Application")` starts a
        fresh one."""
        self.restarts += 1
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        time.sleep(1.0)

    def _win32com_verdict(self, xlsm_path):
        """Runs the win32com verdict check in a *child process*.

        A compile error hangs `excel.Run("Harness")` exactly the way it
        hangs the AppleScript `run VB macro` call -- that is the whole
        signal this driver reads (see the class docstring and this
        module's point 1). A bare in-process win32com call has no way to
        be timed out from within the same process (COM calls block the
        calling thread; there's no clean cross-thread interrupt for one),
        so the actual COM work happens in a child `python -u -c` process,
        and a *subprocess* timeout is what can still kill it out from
        under a hung Excel.
        """
        return subprocess.run(
            [sys.executable, "-u", "-c", _WIN32COM_VERDICT_RUNNER, os.path.abspath(xlsm_path)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
            timeout=self.timeout,
        )

    def verdict(self, xlsm_path):
        """(accepted, detail). `accepted is None` means "could not tell"."""
        if self.driver_type == "mock":
            return None, "mock driver: Excel not invoked"

        if self.driver_type == "win32com":
            for attempt in range(2):
                try:
                    res = self._win32com_verdict(xlsm_path)
                except subprocess.TimeoutExpired:
                    self.restart_windows()
                    if attempt == 0:
                        continue
                    return False, "compile error (Excel went modal and had to be killed)"
                if res.returncode == 0:
                    return True, res.stdout.strip()
                self.restart_windows()
                if attempt == 0:
                    continue
                return None, f"win32com error: {res.stderr.strip()}"
            return None, "indeterminate"

        for attempt in range(2):
            try:
                res = subprocess.run(
                    ["osascript", "-e", self.script(os.path.abspath(xlsm_path))],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, timeout=self.timeout,
                )
            except subprocess.TimeoutExpired:
                # A hang is the compile-error signal -- but it is also what
                # session degradation looks like,
                # so confirm it survives a restart before believing it.
                self.restart_excel()
                if attempt == 0:
                    continue
                return False, "compile error (Excel went modal and had to be killed)"
            if res.returncode == 0:
                return True, res.stdout.strip()
            # A non-timeout AppleScript failure this late is degradation, not
            # a verdict; restart and try once more before giving up.
            self.restart_excel()
            if attempt == 0:
                continue
            return None, f"AppleScript error: {res.stderr.strip()}"
        return None, "indeterminate"


# -----------------------------------------------------------------------------
# 3. Comparison
# -----------------------------------------------------------------------------


def classify(visi_ok, excel_ok):
    if excel_ok is None:
        return "SKIPPED"
    if visi_ok == excel_ok:
        return "PASSED"
    return "FALSE_POSITIVE" if excel_ok else "FALSE_NEGATIVE"


VERDICT_BLURB = {
    "FALSE_POSITIVE": "visi rejected source Excel compiles -- `macro check` would report a bug that isn't there",
    "FALSE_NEGATIVE": "visi accepted source Excel refuses to compile -- `macro check` missed a real error",
}


def run_corpus(cases, path):
    """Parser-only regression check over real `.bas` files.

    Not differential, and deliberately so: Excel compiles a procedure only
    when it is invoked, and invoking one runs it. For an arbitrary module
    there is no way to ask "does this compile?" without also asking it to do
    whatever it does, which is not a thing a test harness should do to code it
    did not write.
    """
    print("=" * 69)
    print("        visi VBA parser: real-world corpus regression check       ".center(69))
    print("=" * 69)
    print(f" Corpus : {path} ({len(cases)} files)")
    print(" Excel  : not consulted -- see --corpus help and finding 2")
    print("=" * 69 + "\n")

    failed = 0
    for name, source, _ in cases:
        ok, detail = visi_verdict(source)
        procs = len(visi_core.check_syntax(source)) if ok else 0
        print(f" {name:<24} [{'OK' if ok else 'REJECTED'}]"
              f"{f' ({procs} procedure{"" if procs == 1 else "s"})' if ok else ''}")
        if not ok:
            print(f"   {detail}")
            failed += 1

    print("\n" + "=" * 69)
    print(f" Accepted: {len(cases) - failed}/{len(cases)}")
    print("=" * 69)
    return 1 if failed else 0


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--excel-path", help="Path to Microsoft Excel binary or application bundle.")
    ap.add_argument("--driver", choices=["auto", "applescript", "win32com", "mock"], default="auto")
    ap.add_argument("--iterations", type=int, default=10)
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument(
        "--mutation-rate", type=float, default=0.5,
        help="Fraction of iterations that get a mutation applied (default 0.5). "
             "Mutated cases are usually invalid, and invalid cases are the slow ones.",
    )
    ap.add_argument(
        "--corpus",
        help="Instead of generating, run visi's parser over every .bas file in this "
             "directory and require it to accept all of them -- a regression check "
             "against real-world VBA. Excel is NOT consulted: getting its verdict "
             "means invoking a procedure, and invoking one runs it (see finding 2).",
    )
    ap.add_argument("--timeout", type=int, default=15,
                    help="Seconds to wait before calling a hang a compile error (default 15).")
    ap.add_argument("--output-dir", default="./fuzz_results")
    args = ap.parse_args()

    driver = ExcelVerdictDriver(args.excel_path, args.driver, args.timeout)
    failures_dir = os.path.join(args.output_dir, "failures")
    os.makedirs(failures_dir, exist_ok=True)

    cases = []
    if args.corpus:
        for name in sorted(os.listdir(args.corpus)):
            if name.endswith(".bas"):
                with open(os.path.join(args.corpus, name)) as f:
                    cases.append((name, f.read(), False))
        return run_corpus(cases, args.corpus)
    else:
        # Resolve the seed rather than passing None straight through:
        # `random.Random(None)` seeds from entropy, which would make a
        # failure found here impossible to reproduce -- the one thing a
        # saved reproduction is for.
        seed = args.seed if args.seed is not None else random.randrange(1_000_000)
        gen = VbaSourceGenerator(seed)
        for i in range(1, args.iterations + 1):
            mutate = gen.rng.random() < args.mutation_rate
            cases.append((f"iter_{i}", build_module(gen.body(mutate)), mutate))

    print("=" * 69)
    print("     visi vs. Microsoft Excel VBA Syntax Differential Fuzzer     ".center(69))
    print("=" * 69)
    print(f" Cases       : {len(cases)}")
    # The seed goes in every saved failure's directory name. Without it two
    # runs both save to `vba_parse_iter_7` and the second silently destroys
    # the first's reproduction.
    run_tag = "" if args.corpus else f"_seed_{seed}"
    print(f" Source      : {args.corpus or f'generated (seed {seed})'}")
    print(f" Excel driver: {driver.driver_type} ({args.excel_path or 'default'})")
    print(f" Timeout     : {args.timeout}s (a hang is how Excel reports a compile error)")
    if driver.driver_type == "mock":
        print(" MOCK DRIVER -- parser runs, Excel is not consulted, nothing is compared.")
    print("=" * 69 + "\n")

    tally = {"PASSED": 0, "FALSE_POSITIVE": 0, "FALSE_NEGATIVE": 0, "SKIPPED": 0}
    start = time.time()
    workdir = tempfile.mkdtemp(prefix="vba_parse_fuzz_")

    try:
        for label, source, mutated in cases:
            visi_ok, visi_detail = visi_verdict(source)

            excel_ok, excel_detail = None, "not run"
            if driver.driver_type != "mock":
                base = os.path.join(workdir, "base.xlsx")
                xlsm = os.path.join(workdir, f"{label}.xlsm")
                openpyxl.Workbook().save(base)
                wb = visi_core.Workbook.load(base)
                # `add_macro` writes source verbatim -- deliberately including
                # source that does not parse, which is the whole point here.
                wb.add_macro("M", source)
                wb.save(xlsm)
                excel_ok, excel_detail = driver.verdict(xlsm)

            result = classify(visi_ok, excel_ok)
            tally[result] += 1

            flag = " (mutated)" if mutated else ""
            print(f" {label:<12} [{result}]{flag} visi={'accept' if visi_ok else 'reject'}"
                  f" excel={'accept' if excel_ok else 'reject' if excel_ok is False else 'n/a'}")
            if result in VERDICT_BLURB:
                print(f"   {VERDICT_BLURB[result]}")
                print(f"   visi : {visi_detail or 'accepted'}")
                print(f"   excel: {excel_detail}")
                out = os.path.join(failures_dir, f"vba_parse_{label}{run_tag}")
                os.makedirs(out, exist_ok=True)
                with open(os.path.join(out, "source.bas"), "w") as f:
                    f.write(source)
                with open(os.path.join(out, "verdicts.txt"), "w") as f:
                    f.write(f"result: {result}\nvisi: {visi_detail or 'accepted'}\nexcel: {excel_detail}\n")
                print(f"   saved: {out}")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    elapsed = time.time() - start
    print("\n" + "=" * 69)
    print(f" Completed in {elapsed:.1f}s ({driver.restarts} Excel restarts)")
    print(f" Agreed         : {tally['PASSED']}/{len(cases)}")
    print(f" False positives: {tally['FALSE_POSITIVE']}  (visi rejects, Excel compiles)")
    print(f" False negatives: {tally['FALSE_NEGATIVE']}  (visi accepts, Excel refuses)")
    if tally["SKIPPED"]:
        print(f" Skipped        : {tally['SKIPPED']}")
    print("=" * 69)

    return 1 if tally["FALSE_POSITIVE"] or tally["FALSE_NEGATIVE"] else 0


if __name__ == "__main__":
    sys.exit(main())
