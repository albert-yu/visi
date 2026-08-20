#!/usr/bin/env python3
"""Grammar-based differential fuzzer for VBA parser syntax.

This is the valid-by-construction companion to ``fuzz_vba_parse.py``'s
fragment/mutation fuzzer. It builds whole modules from a small VBA grammar:
module declarations, procedure signatures, property Get/Let/Set members,
labels/GoTo/GoSub, line continuations, comments, contextual keyword member
names, and optional/named arguments.

The Excel verdict uses the same lazy-compile trick as ``fuzz_vba_parse.py``:
a generated standard ``Harness`` procedure references every generated member
inside ``If False Then``. Excel compiles that dead branch, but does not run it.
"""

import argparse
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fuzz_vba_parse import ExcelVerdictDriver, classify

try:
    import visi_core
except ImportError:
    sys.exit(
        "the visi_core bindings are required: "
        "maturin develop -m visi-python/Cargo.toml --release"
    )


@dataclass
class ModuleSource:
    name: str
    kind: str
    source: str


class VbaGrammarGenerator:
    """Produces small, valid VBA projects from grammar-shaped pieces."""

    SCALAR_TYPES = ["Long", "String", "Double", "Boolean", "Variant"]
    CONTEXTUAL_NAMES = ["Name", "Line", "Get", "Width", "Value", "Item"]

    def __init__(self, seed=None):
        self.rng = random.Random(seed)
        self.counter = 0

    def ident(self, prefix="v"):
        self.counter += 1
        if self.rng.random() < 0.2:
            return self.rng.choice(self.CONTEXTUAL_NAMES) + str(self.counter)
        return f"{prefix}{self.counter}"

    def type_name(self):
        return self.rng.choice(self.SCALAR_TYPES)

    def param(self, optional=False, paramarray=False):
        if paramarray:
            return f"ParamArray {self.ident('args')}() As Variant"
        by = self.rng.choice(["", "ByVal ", "ByRef "])
        opt = "Optional " if optional else ""
        default = ""
        typ = self.type_name()
        if optional:
            if typ == "String":
                default = ' = "d"'
            elif typ == "Boolean":
                default = " = False"
            elif typ != "Variant":
                default = " = 1"
        return f"{opt}{by}{self.ident('p')} As {typ}{default}"

    def params(self, allow_paramarray=False):
        kind = self.rng.choice(["none", "one", "optional", "two"])
        if allow_paramarray and self.rng.random() < 0.25:
            return self.param(paramarray=True)
        if kind == "none":
            return ""
        if kind == "one":
            return self.param()
        if kind == "optional":
            return f"{self.param()}, {self.param(optional=True)}"
        return f"{self.param()}, {self.param()}"

    def declarations(self):
        enum_a = self.ident("EnumA")
        enum_b = self.ident("EnumB")
        return "\n".join(
            [
                "Option Explicit",
                f"Private Const {self.ident('MAX')} As Long = 10",
                f"Private {self.ident('moduleValue')} As String, {self.ident('flags')}(1 To 3) As Boolean",
                "Private Type Point",
                "    X As Long",
                "    Y As Long",
                "End Type",
                "Public Enum GeneratedColor",
                f"    {enum_a} = 1",
                f"    {enum_b}",
                "End Enum",
            ]
        )

    def statement_block(self):
        label = self.ident("Done")
        sub_label = self.ident("SubPath")
        contextual = self.rng.choice(self.CONTEXTUAL_NAMES)
        return "\n".join(
            [
                "    ' generated grammar case",
                "    Dim x As Long, text As String",
                f"    Dim {contextual} As Variant",
                "    Dim arr() As Variant",
                "    x = 1 _",
                "        + 2",
                "    text = \"a\" & _",
                "        \"b\"",
                f"    GoSub {sub_label}",
                f"    If x > 0 Then GoTo {label}",
                f"{sub_label}:",
                "    x = x + 1: Return",
                f"{label}:",
                f"    {contextual} = x",
                "    Call TakesOptional(1, , 3)",
                "    TakesOptional first:=1, third:=3",
                "    ReDim Preserve arr(1 To 2)",
                "    Erase arr",
            ]
        )

    def standard_module(self):
        prop_name = "Answer"
        source = f'''Attribute VB_Name = "GrammarM"
{self.declarations()}

Private Sub TakesOptional(Optional first, Optional second, Optional third)
End Sub

Public Sub GeneratedSub(Optional ByVal first As Variant, Optional ByVal second As Variant)
{self.statement_block()}
End Sub

Private Function Compute(ByVal value As Long, Optional suffix As String = "x") As String
    Compute = CStr(value) & suffix
End Function

Public Property Get {prop_name}() As Long
    {prop_name} = 1
End Property

Public Property Let {prop_name}(ByVal v As Long)
End Property
'''
        return ModuleSource("GrammarM", "standard", source)

    def class_module(self):
        # Kept simple enough that the harness can reference it from dead code
        # without needing any runtime object model. Events are class-module
        # syntax in real Excel, so they live here rather than in the standard
        # harness module.
        source = '''Attribute VB_Name = "GrammarC"
Option Explicit
Public Event Changed(ByVal value As Long)
Private currentValue As Long

Public Property Get Value() As Long
    Value = currentValue
End Property

Public Property Let Value(ByVal v As Long)
    currentValue = v
    RaiseEvent Changed(v)
End Property

Public Sub Touch(Optional ByVal amount As Long = 1)
    currentValue = currentValue + amount
End Sub
'''
        return ModuleSource("GrammarC", "class", source)

    def harness_module(self, include_class):
        class_lines = ""
        if include_class:
            class_lines = """
        Dim c As GrammarC
        Set c = New GrammarC
        c.Value = Answer
        c.Touch amount:=2
"""
        source = f'''Attribute VB_Name = "HarnessM"
Option Explicit

Public Function Harness() As String
    On Error GoTo Failed
    If False Then
        GeneratedSub 1, 2
        Call GeneratedSub(1, 2)
        Dim n As Long
        n = Answer
{class_lines}    End If
    Harness = "OK"
    Exit Function
Failed:
    Harness = "ERR|" & CStr(Err.Number)
End Function
'''
        return ModuleSource("HarnessM", "standard", source)

    def project(self, include_class=False):
        modules = [self.standard_module()]
        if include_class:
            modules.append(self.class_module())
        modules.append(self.harness_module(include_class))
        return modules


class VbaProjectMutator:
    TOKENS = ["End", "Then", "Property", "As", "(", ")", ",", "_", "#", '"']

    def __init__(self, rng):
        self.rng = rng

    def mutate(self, modules):
        modules = [ModuleSource(m.name, m.kind, m.source) for m in modules]
        target = self.rng.randrange(len(modules))
        text = modules[target].source
        kind = self.rng.choice(["delete_line", "insert_token", "truncate"])
        lines = text.splitlines()
        if kind == "delete_line" and len(lines) > 3:
            del lines[self.rng.randrange(1, len(lines))]
            text = "\n".join(lines) + "\n"
        elif kind == "truncate" and len(lines) > 3:
            text = "\n".join(lines[: self.rng.randrange(2, len(lines))]) + "\n"
        else:
            i = self.rng.randrange(len(text))
            text = text[:i] + self.rng.choice(self.TOKENS) + text[i:]
        modules[target].source = text
        return modules


def resolve_visi_binary(binary_path=None):
    if binary_path and os.path.exists(binary_path):
        return binary_path
    candidates = []
    if binary_path:
        candidates.extend([binary_path, binary_path + ".exe"])
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for flavor in ("release", "debug"):
        base = os.path.join(repo, "target", flavor, "visi")
        candidates.extend([base, base + ".exe"])
    existing = [p for p in candidates if os.path.exists(p)]
    return max(existing, key=os.path.getmtime) if existing else binary_path


def write_project(path, modules):
    openpyxl.Workbook().save(path)
    wb = visi_core.Workbook.load(path)
    for module in modules:
        wb.add_macro(module.name, module.source, kind=module.kind)
    xlsm = os.path.splitext(path)[0] + ".xlsm"
    wb.save(xlsm)
    return xlsm


def visi_project_verdict(modules, path, visi_binary):
    xlsm = write_project(path, modules)
    res = subprocess.run(
        [visi_binary, "macro", "check", xlsm, "--json", "--quiet"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=30,
    )
    if res.returncode != 0:
        return False, f"visi macro check failed: {res.stderr.strip() or res.stdout.strip()}"
    checked = json.loads(res.stdout)
    errors = [f"{m['module']}: {m['error']}" for m in checked if m.get("error")]
    return (not errors, "; ".join(errors))


def excel_project_verdict(driver, modules, path):
    if driver.driver_type == "mock":
        return None, "mock driver: Excel not invoked"
    xlsm = write_project(path, modules)
    return driver.verdict(xlsm)


def save_failure(failures_dir, label, seed, modules, result, visi_detail, excel_detail):
    out = os.path.join(failures_dir, f"vba_parse_grammar_{label}_seed_{seed}")
    os.makedirs(out, exist_ok=True)
    for module in modules:
        with open(os.path.join(out, f"{module.name}.bas"), "w", encoding="utf-8") as f:
            f.write(module.source)
    with open(os.path.join(out, "verdicts.txt"), "w", encoding="utf-8") as f:
        f.write(f"result: {result}\nvisi: {visi_detail or 'accepted'}\nexcel: {excel_detail}\n")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel-path", help="Path to Microsoft Excel binary or application bundle.")
    ap.add_argument("--driver", choices=["auto", "applescript", "win32com", "mock"], default="auto")
    ap.add_argument("--iterations", type=int, default=10)
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--mutation-rate", type=float, default=0.0)
    ap.add_argument(
        "--include-classes",
        action="store_true",
        help="Also generate class modules with events/properties. Off by default because Excel's lazy compile path for class modules is slower and stricter.",
    )
    ap.add_argument("--timeout", type=int, default=15)
    ap.add_argument("--output-dir", default="./fuzz_results")
    ap.add_argument("--visi-path", default=None, help="Path to the visi CLI binary for project syntax checks.")
    args = ap.parse_args()

    seed = args.seed if args.seed is not None else random.randrange(1_000_000)
    gen = VbaGrammarGenerator(seed)
    mutator = VbaProjectMutator(gen.rng)
    driver = ExcelVerdictDriver(args.excel_path, args.driver, args.timeout)
    visi_binary = resolve_visi_binary(args.visi_path)
    if not visi_binary or not os.path.exists(visi_binary):
        sys.exit("visi binary not found; run cargo build or pass --visi-path")
    failures_dir = os.path.join(args.output_dir, "failures")
    os.makedirs(failures_dir, exist_ok=True)

    print("=" * 69)
    print("   visi vs. Microsoft Excel VBA Grammar Parser Fuzzer   ".center(69))
    print("=" * 69)
    print(f" Cases       : {args.iterations}")
    print(f" Source      : generated grammar projects (seed {seed})")
    print(f" Visi        : {visi_binary}")
    print(f" Excel driver: {driver.driver_type} ({args.excel_path or 'default'})")
    print(f" Timeout     : {args.timeout}s")
    if driver.driver_type == "mock":
        print(" MOCK DRIVER -- parser runs, Excel is not consulted, nothing is compared.")
    print("=" * 69 + "\n")

    tally = {"PASSED": 0, "FALSE_POSITIVE": 0, "FALSE_NEGATIVE": 0, "SKIPPED": 0}
    workdir = tempfile.mkdtemp(prefix="vba_parse_grammar_")
    start = time.time()
    try:
        for i in range(1, args.iterations + 1):
            modules = gen.project(include_class=args.include_classes)
            mutated = gen.rng.random() < args.mutation_rate
            if mutated:
                modules = mutator.mutate(modules)

            visi_ok, visi_detail = visi_project_verdict(
                modules, os.path.join(workdir, f"case_{i}_visi.xlsx"), visi_binary
            )
            excel_ok, excel_detail = excel_project_verdict(
                driver, modules, os.path.join(workdir, f"case_{i}_excel.xlsx")
            )
            result = classify(visi_ok, excel_ok)
            tally[result] += 1
            flag = " (mutated)" if mutated else ""
            print(f" iter_{i:<7} [{result}]{flag} visi={'accept' if visi_ok else 'reject'} "
                  f"excel={'accept' if excel_ok else 'reject' if excel_ok is False else 'n/a'}")
            if result in ("FALSE_POSITIVE", "FALSE_NEGATIVE"):
                out = save_failure(failures_dir, f"iter_{i}", seed, modules, result, visi_detail, excel_detail)
                print(f"   visi : {visi_detail or 'accepted'}")
                print(f"   excel: {excel_detail}")
                print(f"   saved: {out}")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    elapsed = time.time() - start
    print("\n" + "=" * 69)
    print(f" Completed in {elapsed:.1f}s ({driver.restarts} Excel restarts)")
    print(f" Agreed         : {tally['PASSED']}/{args.iterations}")
    print(f" False positives: {tally['FALSE_POSITIVE']}  (visi rejects, Excel compiles)")
    print(f" False negatives: {tally['FALSE_NEGATIVE']}  (visi accepts, Excel refuses)")
    if tally["SKIPPED"]:
        print(f" Skipped        : {tally['SKIPPED']}")
    print("=" * 69)
    return 1 if tally["FALSE_POSITIVE"] or tally["FALSE_NEGATIVE"] else 0


if __name__ == "__main__":
    sys.exit(main())
