#!/usr/bin/env python3
"""
What does Excel's VBA `PivotTables` surface actually do?
=======================================================
The last surface in Phase 3 of `docs/vba-macro-support.md`:
`PivotTables(name|index)`, `.RefreshTable`, `.PivotFields(...)`, and the
`CurrentPage` that issue #58 singled out as needing a decision.

The file-format half is settled by `pivot_filter_probe.py`; this is the object
model on top of it. Two questions matter most and neither is guessable:

* What does `.CurrentPage` **read** as -- when nothing is filtered, when one
  item is selected, and when several are? Excel shows `(All)` and
  `(Multiple Items)` in the cell, but a property need not return the same
  string.
* Does assigning `.CurrentPage` re-render the grid on its own, or does it
  need `.RefreshTable`? `CLAUDE.md` records that nothing recomputes a pivot
  in visi without an explicit refresh; if Excel differs, a macro that filters
  and then reads cells sees different things in the two engines.

The pivot is built by Excel itself (`EnsurePivot`, idempotent so each case can
call it), because visi cannot create one Excel will treat as live.

    source fuzz/venv/bin/activate
    maturin develop -m visi-python/Cargo.toml --release
    python fuzz/vba_pivot_probe.py
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

# Idempotent so every case can open with it -- a batch shares one session and
# a second `CreatePivotTable` with the same name is a run-time error.
HELPERS = '''Public Sub EnsurePivot()
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable
    Set ws = ThisWorkbook.Sheets("Sheet1")
    On Error Resume Next
    Set pt = ws.PivotTables("P1")
    On Error GoTo 0
    If Not pt Is Nothing Then Exit Sub
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ws.Range("A1:C7"))
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("F1"), TableName:="P1")
    pt.PivotFields("Region").Orientation = xlRowField
    With pt.PivotFields("Amount")
        .Orientation = xlDataField
        .Function = xlSum
    End With
    pt.PivotFields("Product").Orientation = xlPageField
End Sub'''

PREAMBLE = [
    "Dim ws As Worksheet, pt As PivotTable, pf As PivotField",
    "EnsurePivot",
    'Set ws = ThisWorkbook.Sheets("Sheet1")',
    'Set pt = ws.PivotTables("P1")',
    "Dim v, s",
]

# Each case runs in its own round trip: several mutate the filter, and a
# batch would have them measure each other.
CASES = [
    # --- the objects
    "TypeName(ws.PivotTables)",
    "TypeName(ws.PivotTables(1))",
    'TypeName(ws.PivotTables("P1"))',
    "CStr(ws.PivotTables.Count)",
    "pt.Name",
    'TypeName(pt.PivotFields("Product"))',
    "CStr(pt.PivotFields.Count)",
    # --- where the pivot sits
    "pt.TableRange1.Address",
    "pt.TableRange2.Address",
    # --- CurrentPage, unfiltered
    'pt.PivotFields("Product").CurrentPage',
    'TypeName(pt.PivotFields("Product").CurrentPage)',
    # --- a field that is not a page field
    'pt.PivotFields("Region").CurrentPage',
    # --- orientation, as a way to tell the areas apart
    'CStr(pt.PivotFields("Product").Orientation)',
    'CStr(pt.PivotFields("Region").Orientation)',
    'CStr(pt.PivotFields("Amount").Orientation)',
    # --- errors
    'ws.PivotTables("nope").Name',
    "ws.PivotTables(5).Name",
    'pt.PivotFields("nope").Orientation',
    # --- setting CurrentPage, and whether the grid follows without a refresh
    'pt.PivotFields("Product").CurrentPage = "Widget" :: pt.PivotFields("Product").CurrentPage',
    'pt.PivotFields("Product").CurrentPage = "Widget" :: CStr(ws.Range("G5").Value)',
    'pt.PivotFields("Product").CurrentPage = "Widget" :: CStr(ws.Range("G1").Value)',
    # ...and after an explicit refresh, to see whether it changes anything
    'pt.PivotFields("Product").CurrentPage = "Widget"\\npt.RefreshTable :: CStr(ws.Range("G5").Value)',
    # --- setting it to a value that does not exist
    'pt.PivotFields("Product").CurrentPage = "Nonesuch" :: "no error"',
    # --- clearing back to everything
    'pt.PivotFields("Product").CurrentPage = "Widget"\\npt.PivotFields("Product").CurrentPage = "(All)" :: pt.PivotFields("Product").CurrentPage & "/" & CStr(ws.Range("G5").Value)',
    # --- what CurrentPage reads as under a multi-item selection
    'pt.PivotFields("Product").EnableMultiplePageItems = True\\npt.PivotFields("Product").PivotItems("Gadget").Visible = False :: pt.PivotFields("Product").CurrentPage',
    # --- RefreshTable's return value
    "TypeName(pt.RefreshTable)",
    "CStr(pt.RefreshTable)",
]


def parse_case(text):
    setup, sep, expr = text.rpartition("::")
    if not sep:
        return [], text.strip()
    return [s.strip() for s in setup.split("\\n") if s.strip()], expr.strip()


def build_module(cases):
    parts = ['Attribute VB_Name = "V"', HELPERS]
    for i, (setup, expr) in enumerate(cases, start=1):
        body = "\n".join(f"    {s}" for s in PREAMBLE + setup)
        parts.append(f"Private Function Gen{i}()\n{body}\n    Gen{i} = {expr}\nEnd Function")
        parts.append(HARNESS_TEMPLATE.format(i=i))
    return "\n\n".join(parts) + "\n"


def build_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    rows = [
        ("Region", "Product", "Amount"),
        ("East", "Widget", 10),
        ("East", "Gadget", 5),
        ("West", "Widget", 30),
        ("West", "Gadget", 40),
        ("North", "Doohickey", 7),
        ("North", "Widget", 3),
    ]
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--excel-path")
    ap.add_argument("--driver", choices=["auto", "applescript", "mock"], default="auto")
    ap.add_argument("--timeout", type=int, default=180)
    ap.add_argument("-k", "--filter", default="", help="only cases containing this")
    args = ap.parse_args()

    selected = [c for c in CASES if args.filter in c]
    cases = [parse_case(c) for c in selected]
    driver = ExcelDriver(args.excel_path, args.driver, args.timeout)
    workdir = tempfile.mkdtemp(prefix="vba_pivot_probe_")
    base = os.path.join(workdir, "base.xlsx")
    xlsm = os.path.join(workdir, "probe.xlsm")
    build_workbook(base)
    wbk = visi_core.Workbook.load(base)
    wbk.add_macro("V", build_module(cases))
    wbk.save(xlsm)

    width = min(max(len(c) for c in selected), 76)
    for i, text in enumerate(selected, start=1):
        got = driver.run_batch(xlsm, [i])
        answer = got.get(i, "<nothing -- compile error>")
        print(f"{text:<{width}}  {answer}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
