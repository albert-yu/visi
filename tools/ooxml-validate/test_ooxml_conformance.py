from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import visi_core
from validate_ooxml import validate_ooxml_file

XSD_DIR = Path(os.environ["OOXML_XSD_DIR"]) if os.environ.get("OOXML_XSD_DIR") else None
REQUIRE_XSD = os.environ.get("OOXML_REQUIRE_XSD", "") == "1"


def assert_ooxml_valid(path: Path) -> None:
    errors = validate_ooxml_file(path, xsd_dir=XSD_DIR, require_xsd=REQUIRE_XSD)
    assert not errors, f"OOXML validation failed for {path}:\n" + "\n".join(errors)


def test_fresh_export_values_and_formulas(tmp_path: Path) -> None:
    out = tmp_path / "fresh.xlsx"
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "10")
    wb.set_cell(0, 1, "20")
    wb.set_cell(0, 2, "=SUM(A1:B1)")
    wb.set_cell(1, 0, "  leading and trailing  ")
    wb.set_cell(1, 1, "=TRUE()")
    wb.set_cell(1, 2, "=1/0")
    wb.set_cell(2, 0, "=XLOOKUP(10, A1:A1, B1:B1)")
    wb.evaluate()
    wb.save(str(out))

    assert_ooxml_valid(out)

    reloaded = visi_core.Workbook.load(str(out))
    assert reloaded.get_cell(0, 2) == 30
    assert reloaded.get_cell(1, 0) == "  leading and trailing  "
    assert reloaded.get_cell(1, 1) is True
    assert isinstance(reloaded.get_cell(1, 2), visi_core.CellError)


def test_surgical_patch_round_trip(tmp_path: Path) -> None:
    src = tmp_path / "base.xlsx"
    patched = tmp_path / "patched.xlsx"

    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "100")
    wb.set_cell(0, 1, "=A1*2")
    wb.evaluate()
    wb.save(str(src))
    assert_ooxml_valid(src)

    loaded = visi_core.Workbook.load(str(src))
    loaded.set_cell(0, 0, "250")
    loaded.set_cell(1, 0, "patched text")
    loaded.evaluate()
    loaded.save(str(patched))
    assert_ooxml_valid(patched)

    verify = visi_core.Workbook.load(str(patched))
    assert verify.get_cell(0, 1) == 500
    assert verify.get_cell(1, 0) == "patched text"


def test_pivots_and_charts_conformance(tmp_path: Path) -> None:
    out = tmp_path / "objects.xlsx"
    wb = visi_core.Workbook()
    rows = [
        ("Region", "Sales", "Units"),
        ("East", "120", "4"),
        ("West", "200", "6"),
        ("East", "80", "2"),
    ]
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            wb.set_cell(r, c, val)
    wb.evaluate()

    wb.add_pivot_from_range(
        "SalesPivot",
        start_row=0,
        start_col=0,
        end_row=3,
        end_col=2,
        dest_row=0,
        dest_col=5,
    )
    wb.add_pivot_field("SalesPivot", "row", "Region")
    wb.add_pivot_field("SalesPivot", "value", "Sales", agg="sum")
    wb.refresh_pivot("SalesPivot")
    wb.add_chart("Sheet1", "column", "Sheet1!A1:B4", title="Sales by Region")
    wb.save(str(out))

    assert_ooxml_valid(out)


def test_xlsm_macro_conformance(tmp_path: Path) -> None:
    out = tmp_path / "macro.xlsm"
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "42")
    wb.add_macro(
        "Mod1",
        'Attribute VB_Name = "Mod1"\nPublic Sub Test()\n    Range("A2").Value = 99\nEnd Sub\n',
    )
    wb.save(str(out))

    assert_ooxml_valid(out)
