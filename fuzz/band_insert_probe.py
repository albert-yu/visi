#!/usr/bin/env python3
"""
What does a *partial* insert do to formulas?
============================================
`ListRows.Add` is not a row insert. Measured with `vba_table_probe.py`:
adding a row to a table at `A1:C4` moves `A8` down to `A9` but leaves `E2`
alone -- only the table's own columns shift. Excel calls this
`Insert Shift:=xlDown` over a column band.

The engine has no such operation (`Sheet::insert_row` moves the whole row),
and building one means deciding what happens to every formula reference that
touches the band. Some of those have no obvious answer:

* a reference *inside* the band, below the insert point -- surely shifts
* a reference *outside* the band -- surely does not
* a range wholly inside the band -- shifts or grows, as for a row insert
* **a range that straddles the band's edge** -- it cannot both shift and not
  shift, so Excel has to do something arbitrary, and that is the case worth
  knowing before writing any code
* a range that *contains* the whole band

Each case is a formula placed outside the affected area so it survives to be
read, and every case runs in its own Excel round trip because each one
mutates the sheet's shape.

    source fuzz/venv/bin/activate
    python fuzz/band_insert_probe.py
"""

import argparse
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

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

from fuzz_vba import HARNESS_TEMPLATE, ExcelDriver

# The band is A:C and the insert point is row 2, i.e.
# `ws.Range("A2:C2").Insert Shift:=xlDown`. Formulas live in H, which is
# outside the band, so they neither move nor get overwritten.
PREAMBLE = [
    "Dim ws As Worksheet",
    'Set ws = ThisWorkbook.Worksheets("Sheet1")',
    "Dim v, s",
]

# (label, the formula put in H1, what to read back afterwards)
CASES = [
    ("ref inside the band, below the insert", "=A5"),
    ("ref inside the band, above the insert", "=A1"),
    ("ref at the insert point itself", "=A2"),
    ("ref outside the band", "=E5"),
    ("range wholly inside the band, below", "=SUM(A5:A6)"),
    ("range wholly inside the band, spanning the insert", "=SUM(A1:A6)"),
    ("range inside the band, all three columns", "=SUM(A5:C6)"),
    ("range straddling the band edge (A:E)", "=SUM(A5:E5)"),
    ("range straddling, multi-row", "=SUM(A5:E6)"),
    ("range containing the whole band and more", "=SUM(A1:E10)"),
    ("range wholly outside the band", "=SUM(E5:F6)"),
    ("whole-column ref inside the band", "=SUM(A:A)"),
    ("whole-column ref outside the band", "=SUM(E:E)"),
]


def build_module(cases):
    parts = ['Attribute VB_Name = "B"']
    for i, (_, formula) in enumerate(cases, start=1):
        body = "\n".join(f"    {s}" for s in PREAMBLE)
        parts.append(
            f"Private Function Gen{i}()\n{body}\n"
            f'    ws.Range("H1").Formula = "{formula}"\n'
            '    ws.Range("A2:C2").Insert Shift:=xlDown\n'
            f'    Gen{i} = ws.Range("H1").Formula\n'
            "End Function"
        )
        parts.append(HARNESS_TEMPLATE.format(i=i))
    return "\n\n".join(parts) + "\n"


def build_workbook(path):
    """A grid wide enough that a formula can sit outside the A:C band."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r in range(1, 11):
        for c in range(1, 7):
            ws.cell(row=r, column=c, value=r * 10 + c)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--excel-path")
    ap.add_argument("--driver", choices=["auto", "applescript", "mock"], default="auto")
    ap.add_argument("--timeout", type=int, default=180)
    args = ap.parse_args()

    driver = ExcelDriver(args.excel_path, args.driver, args.timeout)
    workdir = tempfile.mkdtemp(prefix="band_insert_probe_")
    base = os.path.join(workdir, "base.xlsx")
    xlsm = os.path.join(workdir, "probe.xlsm")
    build_workbook(base)
    wbk = visi_core.Workbook.load(base)
    wbk.add_macro("B", build_module(CASES))
    wbk.save(xlsm)

    label_w = max(len(l) for l, _ in CASES)
    formula_w = max(len(f) for _, f in CASES)
    print('band = A:C, insert at row 2 (`ws.Range("A2:C2").Insert Shift:=xlDown`)\n')
    print(f"{'case':<{label_w}}  {'before':<{formula_w}}  after")
    # One case per round trip: each changes the sheet's shape.
    for i, (label, formula) in enumerate(CASES, start=1):
        got = driver.run_batch(xlsm, [i])
        answer = got.get(i, "<nothing -- compile error>")
        if answer.startswith("OK|String|"):
            answer = answer[len("OK|String|"):]
        print(f"{label:<{label_w}}  {formula:<{formula_w}}  {answer}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
