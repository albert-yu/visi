import io
import zipfile

from fuzz_excel import DifferentialComparator, XLSXEvaluatedReader
from fuzz_structural_edit import normalize_formula_text


def test_text_whitespace_is_significant():
    comparator = DifferentialComparator()

    assert not comparator.values_equal(" a", "a")
    assert not comparator.values_equal("a ", "a")
    assert not comparator.values_equal(" a ", "a")


def test_blank_equivalence_still_allows_missing_whitespace_only_cell():
    comparator = DifferentialComparator()

    assert comparator.values_equal(None, "")
    assert comparator.values_equal(None, "   ")


def test_numeric_looking_text_is_not_equal_to_numbers():
    comparator = DifferentialComparator()

    assert not comparator.values_equal("08", 8)
    assert not comparator.values_equal(8, "08")
    assert not comparator.values_equal("1", 1)
    assert not comparator.values_equal(1, "1")
    assert not comparator.values_equal(".0394", 0.0394)
    assert not comparator.values_equal(0.0394, ".0394")


def test_evaluated_reader_keys_cells_by_actual_sheet_name():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Data &amp; &quot;Sheet&quot;" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row></sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/worksheets/sheet2.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="2"><c r="B2"><v>42</v></c></row></sheetData>
</worksheet>
""",
        )

    cells = XLSXEvaluatedReader.read_evaluated_cells_bytes(buf.getvalue())

    assert cells[("Summary", "A1")]["val"] == "ok"
    assert cells[("Data & \"Sheet\"", "B2")]["val"] == 42
    assert ("sheet1", "A1") not in cells
    assert ("sheet2", "B2") not in cells


def test_structural_formula_normalization_preserves_string_literal_whitespace():
    assert normalize_formula_text('=" a "&A1') == '=" a "&A1'
    assert normalize_formula_text('=" a " & A1') == '=" a "&A1'
    assert normalize_formula_text('=" a " & a1') == '=" a "&A1'

    assert normalize_formula_text('=" a "&A1') != normalize_formula_text('="a"&A1')
    assert normalize_formula_text('=" a "&A1') != normalize_formula_text('="a "&A1')
    assert normalize_formula_text('=" a "&A1') != normalize_formula_text('=" a"&A1')
    assert normalize_formula_text('=" a "&A1') != normalize_formula_text('="  a  "&A1')


def test_structural_formula_normalization_preserves_string_literal_casing():
    assert normalize_formula_text('="abc" & A1') != normalize_formula_text('="ABC" & A1')
    assert normalize_formula_text('=SUM(a1:a2) & "abc"') == '=SUM(A1:A2)&"abc"'


def test_structural_formula_normalization_handles_escaped_quotes_and_ref_errors():
    assert normalize_formula_text('="He said ""hello world""!" & B2') == '="He said ""hello world""!"&B2'
    assert normalize_formula_text('="He said ""hello world""!" & B2') == normalize_formula_text('="He said ""hello world""!"   &  b2')
    assert normalize_formula_text('=""') == '=""'
    assert normalize_formula_text('=""""') == '=""""'
    assert normalize_formula_text(None) is None

    # Cross-sheet #REF! in non-string parts is normalized, but string literals are untouched
    assert normalize_formula_text('=Data!#REF! + 1') == '=#REF!+1'
    assert normalize_formula_text('=\'Data\'!#REF! + 1') == '=#REF!+1'
    assert normalize_formula_text('="Data!#REF!"') == '="Data!#REF!"'
