from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

visi_core = pytest.importorskip(
    "visi_core",
    reason="build it with `maturin develop -m visi-python/Cargo.toml --release`",
)

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fetch_xsds import ensure_xsds
from validate_ooxml import validate_ooxml_bytes

XSD_DIR = ensure_xsds()


def _zip_bytes(entries: dict[str, str | bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    return buf.getvalue()


def _read_zip_map(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        return {info.filename: zf.read(info.filename) for info in zf.infolist() if not info.is_dir()}


def _build_excel_wild_workbook() -> bytes:
    return _zip_bytes(
        {
            "[Content_Types].xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
                '<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>'
                '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>'
                '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>'
                '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                "</Types>"
            ),
            "_rels/.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                "</Relationships>"
            ),
            "docProps/core.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/">'
                "<dc:creator>WildExcelUser</dc:creator>"
                "<dc:title>Q4 Forecast</dc:title>"
                "</cp:coreProperties>"
            ),
            "docProps/app.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                "<Application>Microsoft Excel</Application>"
                "</Properties>"
            ),
            "xl/_rels/workbook.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
                '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>'
                '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
                "</Relationships>"
            ),
            "xl/workbook.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                "<workbookProtection lockStructure=\"1\"/>"
                '<sheets><sheet name="Forecast" sheetId="1" r:id="rId1"/></sheets>'
                '<extLst><ext uri="{140A7094-0417-4885-8D1C-4012BE97E9F7}"><wbCustom flag="keep"/></ext></extLst>'
                "</workbook>"
            ),
            "xl/styles.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
                '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
                "</styleSheet>"
            ),
            "xl/sharedStrings.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
                "<si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space=\"preserve\"> Title</t></r></si>"
                "</sst>"
            ),
            "xl/theme/theme1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
                '<a:themeElements><a:clrScheme name="Office">'
                '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
                '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
                '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
                '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
                '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
                '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
                '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
                '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
                '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
                '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
                '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
                '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
                "</a:clrScheme></a:themeElements></a:theme>"
            ),
            "xl/calcChain.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<c r="B2" i="1"/>'
                "</calcChain>"
            ),
            "xl/slicers/slicer1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
                '<slicer name="RegionSlicer" cache="Slicer_Region"/>'
                "</slicers>"
            ),
            "xl/threadedComments/threadedComment1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
                '<threadedComment ref="A1" id="{11111111-1111-1111-1111-111111111111}"><text>Check Q4</text></threadedComment>'
                "</ThreadedComments>"
            ),
            "xl/drawings/vmlDrawing1.vml": "<xml><v:shape id=\"_x0000_s1025\"/></xml>",
            "xl/externalLinks/externalLink1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
                "</externalLink>"
            ),
            "xl/externalLinks/_rels/externalLink1.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///source.xlsx" TargetMode="External"/>'
                "</Relationships>"
            ),
            "xl/worksheets/_rels/sheet1.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
                "</Relationships>"
            ),
            "xl/worksheets/sheet1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
                'mc:Ignorable="x14">'
                "<sheetData>"
                '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
                '<row r="2"><c r="A2"><v>10</v></c><c r="B2"><f>A2*2</f><v>20</v></c></row>'
                "</sheetData>"
                '<sheetProtection sheet="1" objects="1"/>'
                '<autoFilter ref="A1:B2"/>'
                '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                '<pageSetup orientation="landscape"/>'
                '<headerFooter><oddHeader>&amp;CConfidential</oddHeader></headerFooter>'
                '<legacyDrawing r:id="rId3"/>'
                "<extLst>"
                '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext>'
                '<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"><x14:slicerList><x14:slicer r:id="rId1"/></x14:slicerList></ext>'
                "</extLst>"
                "</worksheet>"
            ),
        }
    )


def _build_strict_ooxml_workbook() -> bytes:
    return _zip_bytes(
        {
            "[Content_Types].xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                "</Types>"
            ),
            "_rels/.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>'
                "</Relationships>"
            ),
            "xl/_rels/workbook.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/>'
                "</Relationships>"
            ),
            "xl/workbook.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" '
                'xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">'
                '<sheets><sheet name="StrictSheet" sheetId="1" r:id="rId1"/></sheets>'
                "</workbook>"
            ),
            "xl/styles.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<styleSheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">'
                '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
                '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
                "</styleSheet>"
            ),
            "xl/worksheets/sheet1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">'
                "<sheetData>"
                '<row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>A1+5</f><v>12</v></c></row>'
                "</sheetData>"
                '<extLst><ext uri="{STRICT-KEEP}"><marker v="1"/></ext></extLst>'
                "</worksheet>"
            ),
        }
    )


def _build_prefixed_and_absolute_rels_workbook() -> bytes:
    return _zip_bytes(
        {
            "[Content_Types].xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                "</Types>"
            ),
            "_rels/.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml"/>'
                "</Relationships>"
            ),
            "xl/_rels/workbook.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="/xl/styles.xml"/>'
                "</Relationships>"
            ),
            "xl/workbook.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<x:sheets><x:sheet name="POISheet" sheetId="1" r:id="rId1"/></x:sheets>'
                "</x:workbook>"
            ),
            "xl/styles.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<x:fonts count="1"><x:font><x:sz val="11"/><x:name val="Calibri"/></x:font></x:fonts>'
                '<x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="gray125"/></x:fill></x:fills>'
                '<x:borders count="1"><x:border><x:left/><x:right/><x:top/><x:bottom/><x:diagonal/></x:border></x:borders>'
                '<x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellStyleXfs>'
                '<x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></x:cellXfs>'
                "</x:styleSheet>"
            ),
            "xl/worksheets/sheet1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                "<x:sheetData>"
                '<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Inline Hello</x:t></x:is></x:c>'
                '<x:c r="B1"><x:v>10</x:v></x:c></x:row>'
                '<x:row r="2"><x:c r="B2"><x:f t="shared" ref="B2:B3" si="0">B1*2</x:f><x:v>20</x:v></x:c></x:row>'
                '<x:row r="3"><x:c r="B3"><x:f t="shared" si="0"/><x:v>40</x:v></x:c></x:row>'
                "</x:sheetData>"
                '<x:pageSetup orientation="portrait"/>'
                "</x:worksheet>"
            ),
        }
    )


def test_wild_excel_workbook_preserves_unmodeled_parts_and_extensions_on_patch() -> None:
    raw = _build_excel_wild_workbook()
    wb = visi_core.Workbook.load_bytes(raw)
    assert wb.sheet_names == ["Forecast"]
    assert wb.get_cell(0, 0) == "Bold Title"
    assert wb.get_cell(1, 1) == 20.0

    wb.set_cell(2, 0, "Added note")
    patched = wb.save_bytes()

    errors = validate_ooxml_bytes(patched, xsd_dir=XSD_DIR, require_xsd=True)
    assert not errors, f"Patched wild Excel workbook failed validation:\n" + "\n".join(errors)

    parts = _read_zip_map(patched)
    for required_part in (
        "docProps/core.xml",
        "docProps/app.xml",
        "xl/theme/theme1.xml",
        "xl/calcChain.xml",
        "xl/slicers/slicer1.xml",
        "xl/threadedComments/threadedComment1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/externalLinks/externalLink1.xml",
        "xl/externalLinks/_rels/externalLink1.xml.rels",
    ):
        assert required_part in parts, f"Surgical patch dropped {required_part}"

    ws_xml = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    for snippet in (
        "sheetProtection",
        "autoFilter",
        "pageSetup",
        "headerFooter",
        "legacyDrawing",
        "sparklineGroups",
        "slicerList",
    ):
        assert snippet in ws_xml, f"Surgical patch dropped <{snippet}>"

    wb_xml = parts["xl/workbook.xml"].decode("utf-8")
    assert "workbookProtection" in wb_xml
    assert "wbCustom" in wb_xml

    sst_xml = parts["xl/sharedStrings.xml"].decode("utf-8")
    assert "<r><rPr><b/></rPr><t>Bold</t></r>" in sst_xml


def test_wild_excel_workbook_invalidates_calc_chain_on_formula_patch() -> None:
    raw = _build_excel_wild_workbook()
    wb = visi_core.Workbook.load_bytes(raw)
    wb.set_cell(1, 1, "=A2*5")
    wb.evaluate()
    assert wb.get_cell(1, 1) == 50.0

    patched = wb.save_bytes()
    errors = validate_ooxml_bytes(patched, xsd_dir=XSD_DIR, require_xsd=True)
    assert not errors, f"Formula-patched wild workbook failed validation:\n" + "\n".join(errors)

    parts = _read_zip_map(patched)
    assert "xl/calcChain.xml" not in parts
    assert "xl/slicers/slicer1.xml" in parts
    assert "docProps/core.xml" in parts


def test_strict_ooxml_workbook_survives_surgical_patch() -> None:
    raw = _build_strict_ooxml_workbook()
    wb = visi_core.Workbook.load_bytes(raw)
    assert wb.sheet_names == ["StrictSheet"]
    assert wb.get_cell(0, 1) == 12.0

    wb.set_cell(0, 0, "10")
    wb.evaluate()
    assert wb.get_cell(0, 1) == 15.0

    patched = wb.save_bytes()
    errors = validate_ooxml_bytes(patched, xsd_dir=None, require_xsd=False)
    assert not errors, f"Patched Strict OOXML workbook failed OPC/semantic validation:\n" + "\n".join(errors)

    parts = _read_zip_map(patched)
    ws_xml = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert "http://purl.oclc.org/ooxml/spreadsheetml/main" in ws_xml
    assert "STRICT-KEEP" in ws_xml


def test_prefixed_xml_absolute_rels_inline_str_and_shared_formulas_round_trip() -> None:
    raw = _build_prefixed_and_absolute_rels_workbook()
    wb = visi_core.Workbook.load_bytes(raw)
    assert wb.sheet_names == ["POISheet"]
    assert wb.get_cell(0, 0) == "Inline Hello"
    assert wb.get_cell(1, 1) == 20.0
    assert wb.get_cell(2, 1) == 40.0

    wb.set_cell(3, 0, "Patched")
    patched = wb.save_bytes()
    errors = validate_ooxml_bytes(patched, xsd_dir=XSD_DIR, require_xsd=True)
    assert not errors, f"Patched prefixed/POI workbook failed validation:\n" + "\n".join(errors)

    parts = _read_zip_map(patched)
    ws_xml = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert "pageSetup" in ws_xml

    reloaded = visi_core.Workbook.load_bytes(patched)
    assert reloaded.get_cell(0, 0) == "Inline Hello"
    assert reloaded.get_cell(2, 1) == 40.0
    assert reloaded.get_cell(3, 0) == "Patched"
