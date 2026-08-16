#!/usr/bin/env python3
"""
How does Excel encode a pivot filter selection in the file?
==========================================================
`PivotFilterField::selected_values` is deliberately *not* reconstructed on
import today -- `pivot_xlsx.rs` resets it to "all", because restoring it would
mean trusting index-based item references against source data that may since
have changed. Issue #58 forces the question: a macro can set a filter, save,
and carry on, so `PivotFields(...).CurrentPage` cannot be exposed on top of a
gap that silently drops what the macro just did.

Closing the gap means writing and reading real `<sharedItems>` values, and
`CLAUDE.md` warns that visi's pivot XML was only ever validated against
`openpyxl` -- never real Excel, which accepts a malformed pivot part silently
because `refreshOnLoad="1"` lets it rebuild the cache. So a mistake here does
not announce itself.

This asks Excel directly: it *builds* the pivot, sets a filter, saves, and the
XML it wrote is dumped. That is the ground truth to parse and emit against,
and it is independent of anything visi does.

    source fuzz/venv/bin/activate
    maturin develop -m visi-python/Cargo.toml --release
    python fuzz/pivot_filter_probe.py                 # both variants
    python fuzz/pivot_filter_probe.py --variant page  # just CurrentPage

Two variants, because Excel encodes them differently and a macro can reach
both:

* `multi`  -- `EnableMultiplePageItems = True` and individual `PivotItems(x).Visible = False`
* `page`   -- a single `.CurrentPage = "Widget"`

`--variant visi` is the other direction: can Excel open a pivot table *visi*
wrote? Exits non-zero if not, so it works as a check and not only as an
exploration.
"""

import argparse
import os
import re
import subprocess
import sys
import tempfile
import zipfile

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

from fuzz_vba import ExcelDriver

BUILD = """Attribute VB_Name = "P"
Public Sub Build()
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable, pf As PivotField
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ws.Range("A1:C7"))
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("F1"), TableName:="P1")
    pt.PivotFields("Region").Orientation = xlRowField
    With pt.PivotFields("Amount")
        .Orientation = xlDataField
        .Function = xlSum
    End With
    Set pf = pt.PivotFields("Product")
    pf.Orientation = xlPageField
{selection}
End Sub
"""

# A pivot that *visi* wrote, opened and re-saved by Excel. This is the
# check `CLAUDE.md` asks for and never got: visi's pivot XML was validated
# against openpyxl, which does not rebuild the cache, so an `<item x="N"/>`
# index pointing at the wrong `sharedItems` entry looks fine there and only
# misbehaves in Excel.
VISI_VARIANT = "visi"

VARIANTS = {
    # Hide one of three items, leaving two selected.
    "multi": "    pf.EnableMultiplePageItems = True\n"
             '    pf.PivotItems("Gadget").Visible = False',
    # The single-selection form, which is what `.CurrentPage` sets.
    "page": '    pf.CurrentPage = "Widget"',
}

# Parts worth seeing. The cache definition is where item *values* live; the
# table part is where the selection is recorded against them.
PARTS = ("xl/pivotCache/pivotCacheDefinition", "xl/pivotTables/pivotTable")


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


def run_and_save(driver, xlsm, out_path, macro="Build"):
    """Opens the workbook, optionally runs a macro, and saves.

    `macro=None` is the "just let Excel rewrite the file" case, which is what
    checking visi's own output needs -- and it avoids `run VB macro`, whose
    sporadic "Parameter error (-50)" is a known Mac Excel bridge fault rather
    than anything to do with the workbook (see `fuzz_pivot.py`).
    """
    run_line = [f'    run VB macro "{macro}"'] if macro else []
    script = "\n".join([
        f'tell application "{driver.app_name()}"',
        "    set display alerts to false",
        "    try",
        "        close workbooks saving no",
        "    end try",
        f'    open POSIX file "{os.path.abspath(xlsm)}"',
        "    set wb to active workbook",
        *run_line,
        f'    save wb in POSIX file "{os.path.abspath(out_path)}"',
        "    close wb saving no",
        '    return "ok"',
        "end tell",
    ])
    try:
        res = subprocess.run(["osascript", "-e", script], stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE, text=True, timeout=driver.timeout)
    except subprocess.TimeoutExpired:
        subprocess.run(["killall", "Microsoft Excel"], stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        raise RuntimeError("Excel did not respond; killed it")
    if res.returncode != 0:
        raise RuntimeError(f"AppleScript failed: {res.stderr.strip()}")
    return res.stdout.strip()


def pretty(xml, keep):
    """The elements worth reading, one per line, with the noise dropped."""
    out = []
    for m in re.finditer(r"<[^>]+>", xml):
        tag = m.group(0)
        local = re.match(r"</?([A-Za-z0-9_]+)", tag)
        if local and local.group(1) in keep:
            out.append(tag)
    return "\n".join(out)


def can_excel_open(driver, path, timeout=60):
    """Whether Excel opens the workbook at all.

    Returns (ok, detail). A timeout means a modal dialog -- Excel offering to
    *repair* a file it considers damaged -- which is indistinguishable from a
    hang over the AppleScript bridge, so it is reported as its own outcome.
    """
    script = "\n".join([
        f'tell application "{driver.app_name()}"',
        "    set display alerts to false",
        "    try",
        "        close workbooks saving no",
        "    end try",
        f'    open POSIX file "{os.path.abspath(path)}"',
        "    set n to name of active workbook",
        "    close active workbook saving no",
        "    return n",
        "end tell",
    ])
    try:
        res = subprocess.run(["osascript", "-e", script], stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        subprocess.run(["killall", "Microsoft Excel"], stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        return False, "timed out -- a modal dialog, most likely a repair prompt"
    if res.returncode != 0:
        return False, res.stderr.strip().split(": ")[-1]
    return True, res.stdout.strip()


def visi_written(driver, full):
    """Can real Excel open a pivot table that visi wrote?

    `CLAUDE.md` records that visi's pivot XML was validated against openpyxl
    and never against Excel, because the automation grant could not be
    completed at the time. openpyxl is a strict *reader* but it does not
    resolve `<item x="N"/>` against `<sharedItems>`, so an index pointing into
    an empty list reads fine there.

    Two controls make the answer unambiguous: the same workbook before visi
    touches it, and a visi round trip with no pivot in it.
    """
    workdir = tempfile.mkdtemp(prefix="pivot_filter_visi_")
    base = os.path.join(workdir, "base.xlsx")
    nopivot = os.path.join(workdir, "nopivot.xlsx")
    made = os.path.join(workdir, "pivot.xlsx")
    build_workbook(base)

    visi_core.Workbook.load(base).save(nopivot)

    wbk = visi_core.Workbook.load(base)
    wbk.add_pivot_from_range(name="P1", source_sheet=None, start_row=0, start_col=0,
                             end_row=6, end_col=2, dest_sheet=None, dest_row=0,
                             dest_col=5, grand_totals_row=True, grand_totals_col=True)
    wbk.add_pivot_field("P1", "row", "Region")
    wbk.add_pivot_field("P1", "value", "Amount", agg="sum")
    wbk.add_pivot_field("P1", "filter", "Product")
    wbk.set_pivot_filter("P1", "Product", ["Widget"])
    wbk.refresh_pivot("P1")
    wbk.save(made)

    print("can Excel open it?\n")
    failures = 0
    for label, path in (
        ("the source workbook (control)", base),
        ("visi round trip, no pivot (control)", nopivot),
        ("visi round trip WITH a pivot", made),
    ):
        ok, detail = can_excel_open(driver, path)
        print(f"  {'yes' if ok else 'NO ':<4} {label:<38} {detail}")
        if not ok:
            failures += 1

    if failures:
        print("\nvisi's own pivot XML, for comparison with what Excel writes above:")
        with zipfile.ZipFile(made) as z:
            for name in sorted(z.namelist()):
                if not name.startswith(PARTS):
                    continue
                xml = z.read(name).decode()
                print(f"\n--- {name} ---")
                print(xml if full else pretty(xml, {
                    "cacheField", "sharedItems", "s", "n",
                    "pivotField", "item", "items", "pageField", "pageFields",
                }))
    return 1 if failures else 0


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--excel-path")
    ap.add_argument("--driver", choices=["auto", "applescript", "mock"], default="auto")
    ap.add_argument("--timeout", type=int, default=180)
    ap.add_argument("--variant", choices=sorted(VARIANTS) + [VISI_VARIANT],
                    help="only this variant")
    ap.add_argument("--full", action="store_true", help="dump the parts verbatim")
    args = ap.parse_args()

    driver = ExcelDriver(args.excel_path, args.driver, args.timeout)
    if args.variant == VISI_VARIANT:
        return visi_written(driver, args.full)
    wanted = [args.variant] if args.variant else sorted(VARIANTS)

    for variant in wanted:
        workdir = tempfile.mkdtemp(prefix=f"pivot_filter_{variant}_")
        base = os.path.join(workdir, "base.xlsx")
        xlsm = os.path.join(workdir, "probe.xlsm")
        saved = os.path.join(workdir, "built.xlsm")
        build_workbook(base)
        wbk = visi_core.Workbook.load(base)
        wbk.add_macro("P", BUILD.format(selection=VARIANTS[variant]))
        wbk.save(xlsm)

        print(f"\n{'=' * 70}\n{variant}: {VARIANTS[variant].strip()}\n{'=' * 70}")
        try:
            run_and_save(driver, xlsm, saved)
        except RuntimeError as exc:
            print(f"  failed: {exc}")
            continue

        with zipfile.ZipFile(saved) as z:
            for name in sorted(z.namelist()):
                if not name.startswith(PARTS):
                    continue
                xml = z.read(name).decode()
                print(f"\n--- {name} ---")
                if args.full:
                    print(xml)
                else:
                    print(pretty(xml, {
                        "cacheField", "sharedItems", "s", "n", "b", "m",
                        "pivotField", "item", "items", "pageField", "pageFields",
                    }))
    return 0


if __name__ == "__main__":
    sys.exit(main())
