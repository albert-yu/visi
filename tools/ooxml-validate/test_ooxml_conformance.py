from __future__ import annotations

import io
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fetch_xsds import ensure_xsds
from validate_ooxml import validate_ooxml_bytes, validate_ooxml_file

PROJECT_ROOT = Path(__file__).resolve().parents[2]
XSD_DIR = ensure_xsds()

EXPECTED_FIXTURES = [
    "01_values_and_types.xlsx",
    "01_values_and_types_patched.xlsx",
    "02_styles_and_formatting.xlsx",
    "02_styles_and_formatting_patched.xlsx",
    "03_sheet_structure.xlsx",
    "03_sheet_structure_patched.xlsx",
    "04_formulas_and_arrays.xlsx",
    "04_formulas_and_arrays_patched.xlsx",
    "05_comments_hyperlinks_validation_cf.xlsx",
    "05_comments_hyperlinks_validation_cf_patched.xlsx",
    "06_tables.xlsx",
    "06_tables_patched.xlsx",
    "07_pivots.xlsx",
    "07_pivots_patched.xlsx",
    "08_charts.xlsx",
    "08_charts_patched.xlsx",
    "09_structural_edits_patched.xlsx",
    "10_vba_macros.xlsm",
    "10_vba_macros_patched.xlsm",
]


@pytest.fixture(scope="session")
def emitted_fixture_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    out_dir = tmp_path_factory.mktemp("ooxml_fixtures")
    res = subprocess.run(
        [
            "cargo",
            "run",
            "--quiet",
            "-p",
            "visi-core",
            "--example",
            "emit_ooxml_fixtures",
            "--",
            str(out_dir),
        ],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode == 0, f"emit_ooxml_fixtures failed:\n{res.stderr}"
    return out_dir


@pytest.mark.parametrize("fixture_name", EXPECTED_FIXTURES)
def test_emitted_fixture_passes_ecma376_xsd_and_invariants(
    emitted_fixture_dir: Path, fixture_name: str
) -> None:
    path = emitted_fixture_dir / fixture_name
    assert path.is_file(), f"Missing fixture {fixture_name}"
    errors = validate_ooxml_file(path, xsd_dir=XSD_DIR, require_xsd=True)
    assert not errors, f"OOXML validation failed for {fixture_name}:\n" + "\n".join(errors)


def _zip_bytes(entries: dict[str, str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    return buf.getvalue()


def test_validator_rejects_xsd_and_opc_violations() -> None:
    bad_pkg = _zip_bytes(
        {
            "[Content_Types].xml": (
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                "</Types>"
            ),
            "_rels/.rels": (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>'
                "</Relationships>"
            ),
            "xl/_rels/workbook.xml.rels": (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing.xml"/>'
                "</Relationships>"
            ),
            "xl/workbook.xml": (
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/>'
                '<sheet name="sheet1" sheetId="1" r:id="rId99"/></sheets>'
                "</workbook>"
            ),
            "xl/worksheets/sheet1.xml": (
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="2"><c r="B2" t="b"><v>yes</v></c>'
                '<c r="A2" t="e"><v>#BOGUS!</v></c></row><row r="1"/></sheetData>'
                '<cols/><mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>'
                "</worksheet>"
            ),
        }
    )
    errors = validate_ooxml_bytes(bad_pkg, xsd_dir=XSD_DIR, require_xsd=True)
    joined = "\n".join(errors)
    assert "TargetMode='External'" in joined
    assert "dangling relationship" in joined
    assert "duplicate sheet name" in joined
    assert "violates XSD sequence" in joined
    assert "sml.xsd" in joined
    assert "overlapping mergeCells" in joined
