_The following content is LLM-generated._

# OOXML Specification Conformance Matrix

Tracks `visi-core`'s SpreadsheetML (`ECMA-376 Part 1 §18`, `Part 2 OPC`, `Part 1 §21 DrawingML`, and `[MS-XLSX]`) compatibility, separating **Modeled**, **Pass-Through Preserved**, **Partial**, and **Unsupported** features along with their verification coverage and known gaps.

VBA macro conformance (`[MS-OVBA]` / `[MS-VBAL]`) is tracked separately in [tools/vba-validate/CONFORMANCE.md](../vba-validate/CONFORMANCE.md).

## Status Legend

- **Modeled** — Read, write, and edit supported in `visi-core`; emitted on both fresh export and patch.
- **Pass-Through (Patch)** — Preserved verbatim when surgical patching applies (`OoxmlSourceState::can_patch_surgically` in `visi-core/src/core/xlsx/patch.rs`); not modeled in memory and dropped if a mutation triggers full-export fallback.
- **Partial** — Subset modeled or read-only/write-only; see Notes for exact boundaries.
- **Unsupported** — Not modeled and not generated (may still survive surgical patch if stored as an opaque part or sibling element).

## Verification Layers

| Code | Layer | Implementation |
| --- | --- | --- |
| `U` | Rust unit / round-trip tests | `visi-core/src/core/xlsx/*.rs`, `pivot_xlsx.rs`, `table.rs`, `pivot.rs` |
| `V` | Pure-Rust OPC / sequence / semantic validator | `visi-core/src/core/xlsx/validate.rs` |
| `F` | Emitted + patched fixture suite | `visi-core/examples/emit_ooxml_fixtures.rs` |
| `X` | ECMA-376 Transitional XSD + Python semantic validator | `tools/ooxml-validate/validate_ooxml.py`, `test_ooxml_conformance.py` |
| `C` | Wild corpus patch-preservation tests | _Planned_ |
| `S` | Microsoft Open XML SDK `OpenXmlValidator` (`[MS-XLSX]`) | _Planned_ |
| `O` | Real-application oracle (LibreOffice / Microsoft Excel) | _Planned_ |

---

## 1. OPC Package & Namespaces (`ECMA-376 Part 2` & `Part 1 §18.2`)

| Area | Spec Reference | Status | Verification | Notes & Gaps |
| --- | --- | --- | --- | --- |
| ZIP container, `[Content_Types].xml`, `_rels/.rels`, part `.rels` | `ECMA-376 Part 2 §9–10` | Modeled | `U, V, F, X` | Validated for part-name normalization, case-insensitive collisions, content-type coverage, and dangling/external relationships |
| Transitional namespaces (`schemas.openxmlformats.org`) | `ECMA-376 Part 4` | Modeled | `U, V, F, X` | Default export flavor (`NamespaceFlavor::Transitional`) |
| Strict namespaces (`purl.oclc.org/ooxml`) | `ECMA-376 Part 1` | Partial | `U` | Detected on import and preserved on patch (`NamespaceFlavor::Strict`); fresh export always uses Transitional; `validate_ooxml.py` XSD checks currently load Transitional XSDs only |
| Macro-enabled `.xlsm` content type & `xl/vbaProject.bin` | `[MS-XLSX] §2.1.7.61` | Modeled | `U, V, F, X` | Covered by `10_vba_macros.xlsm` and `10_vba_macros_patched.xlsm` |
| Package properties (`docProps/core.xml`, `docProps/app.xml`, `custom.xml`) | `ECMA-376 Part 2 §11` | Pass-Through (Patch) | — | Preserved on surgical patch; not emitted on fresh export; needs `C` coverage |
| Encrypted packages (`EncryptedPackage` / Agile & Standard) | `[MS-OFFCRYPTO]` | Unsupported | — | Not supported |
| Package digital signatures (`_xmlsignatures`) | `ECMA-376 Part 2 §13` | Unsupported | — | Not modeled |

---

## 2. Workbook & Worksheet Structure (`ECMA-376 Part 1 §18.2–18.3`)

| Area | Spec Reference | Status | Verification | Notes & Gaps |
| --- | --- | --- | --- | --- |
| Sheets (`<sheets>`, names, `sheetId`, `r:id`) | `§18.2.19–20` | Modeled | `U, V, F, X` | Enforces unique case-insensitive names, `<= 31` chars, forbidden char validation; structural add/delete/rename covered by `09_structural_edits_patched.xlsx` |
| Sheet visibility (`visible`, `hidden`, `veryHidden`) | `§18.18.68` | Modeled | `U, V, F, X` | Covered by `03_sheet_structure.xlsx` |
| 1900 / 1904 date system (`workbookPr/@date1904`) | `§18.2.28` | Modeled | `U, V, F, X` | Covered by `03_sheet_structure.xlsx` and surgical patch |
| Defined names (`<definedNames>`, global & `localSheetId`, `hidden`, `comment`) | `§18.2.5–6` | Modeled | `U, V, F, X` | Covered by `04_formulas_and_arrays.xlsx` and surgical patch |
| Cell values & types (`n`, `b`, `e`, `s`, `str`, `inlineStr`) | `§18.3.1.4` | Modeled | `U, V, F, X` | Covered by `01_values_and_types.xlsx`; validates ascending rows/cols, boolean `0`/`1`, and ECMA-376 error literals |
| Shared Strings (`xl/sharedStrings.xml`, `_xHHHH_`, `xml:space="preserve"`) | `§18.4` | Partial | `U, V, F, X` | Plain strings fully modeled; existing rich-text (`<r>`) and phonetic (`<rPh>`) entries survive untouched on patch, but newly written/edited strings are plain `<t>` |
| Column widths, hidden cols, row heights, hidden rows | `§18.3.1.13`, `§18.3.1.73` | Modeled | `U, V, F, X` | Covered by `03_sheet_structure.xlsx` and surgical patch |
| Merged cells (`<mergeCells>`) | `§18.3.1.55` | Modeled | `U, V, F, X` | Overlap validation in `validate.rs` and `validate_ooxml.py`; covered by `03_sheet_structure.xlsx` |
| Freeze panes (`<sheetViews>/<sheetView>/<pane>`) | `§18.3.1.66` | Modeled | `U, V, F, X` | Row, column, and row+col freeze panes covered by `03_sheet_structure.xlsx` |
| Worksheet `<autoFilter>` & `<sortState>` | `§18.3.1.2`, `§18.3.1.92` | Pass-Through (Patch) | — | Preserved on surgical patch (table `<autoFilter>` is modeled in `table.rs`); standalone sheet autoFilter not modeled |
| Page setup, margins, print options, headers/footers, page breaks | `§18.3.1.62–65`, `§18.3.1.46`, `§18.3.1.74` | Pass-Through (Patch) | `U` | Preserved in XSD sequence order on surgical patch via `dom.rs` |
| Sheet & workbook protection (`sheetProtection`, `workbookProtection`, `protectedRanges`) | `§18.3.1.85`, `§18.2.29` | Pass-Through (Patch) | — | Preserved on surgical patch; not modeled |
| Extension lists (`<extLst>`) & `mc:AlternateContent` | `§18.2.10`, `Part 3 §10` | Pass-Through (Patch) | `U` | Preserved on surgical patch via `dom.rs`; dropped on full-export fallback; needs `C` coverage |

---

## 3. Styles, Themes & Annotations (`ECMA-376 Part 1 §18.3, §18.7–18.8`)

| Area | Spec Reference | Status | Verification | Notes & Gaps |
| --- | --- | --- | --- | --- |
| Cell styles (`fonts`, `fills`, `borders`, `alignment`, `numFmts`, `cellXfs`) | `§18.8` | Modeled | `U, V, F, X` | Bold/italic/underline/strike, font family/size/color, solid & pattern fills, side + diagonal borders, alignment/wrap/rotation/indent/shrink, custom number formats (`02_styles_and_formatting.xlsx`) |
| Theme colors (`xl/theme/theme1.xml`) | `§20.1.6.9` | Partial | `U` | Theme palette parsed on import to resolve `theme`/`tint` colors; original `theme1.xml` preserved on surgical patch; not emitted on fresh export |
| Conditional formatting (`<conditionalFormatting>` + `<dxfs>`) | `§18.3.1.18`, `§18.8.14` | Partial | `U, V, F, X` | `cellIs` and `expression` rules with `dxf` styles modeled (`05_comments_hyperlinks_validation_cf.xlsx`); `colorScale`, `dataBar`, `iconSet`, and `x14:conditionalFormattings` are pass-through on surgical patch only |
| Data validations (`<dataValidations>`) | `§18.3.1.33` | Modeled | `U, V, F, X` | `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`, operators, input/error messages (`05_comments_hyperlinks_validation_cf.xlsx`); `x14:dataValidations` cross-sheet list refs pass-through only |
| Hyperlinks (`<hyperlinks>`, external & internal) | `§18.3.1.48` | Modeled | `U, V, F, X` | External `TargetMode="External"` relationships + internal `location`, display, tooltip (`05_comments_hyperlinks_validation_cf.xlsx`) |
| Legacy cell comments (`xl/commentsN.xml`) | `§18.7` | Partial | `U, V, F, X` | Authors and comment text modeled (`05_comments_hyperlinks_validation_cf.xlsx`); legacy VML drawing (`xl/drawings/vmlDrawingN.vml`) is not emitted on fresh export (Excel renders comments using VML for note box geometry) |
| Threaded comments (`xl/threadedComments/`, `xl/persons/`) | `[MS-XLSX] §2.1.7.53` | Pass-Through (Patch) | — | Preserved on surgical patch; not modeled |

---

## 4. Tables, Pivot Tables & Charts (`ECMA-376 Part 1 §18.5, §18.10, §21.2`)

| Area | Spec Reference | Status | Verification | Notes & Gaps |
| --- | --- | --- | --- | --- |
| Tables (`xl/tables/tableN.xml`, `tableParts`, `tableColumns`, `tableStyleInfo`) | `§18.5.1` | Modeled | `U, V, F, X` | Covered by `06_tables.xlsx` and `06_tables_patched.xlsx` (including row insert inside table) |
| Pivot tables (`pivotTableN.xml`, `pivotCacheDefinitionN.xml`, `pivotCacheRecordsN.xml`) | `§18.10`, `§12.3.12` | Modeled | `U, V, F, X` | Row/column/filter/value fields, aggregations, subtotals, grand totals, and `xl/pivotTables/_rels/pivotTableN.xml.rels` -> `pivotCacheDefinitionN.xml` modeled (`07_pivots.xlsx`) |
| Charts & drawings (`xl/charts/chartN.xml`, `xl/drawings/drawingN.xml`) | `§21.2`, `§20.5` | Partial | `U, V, F, X` | Column, Bar, Line, Pie, Scatter, Area with two-cell anchors, titles, axis titles, legend (`08_charts.xlsx`); combo/3D/ChartEx (`cx:`) charts not modeled |
| Images, shapes, sparklines, slicers, timelines, OLE/ActiveX controls | `§20.5`, `[MS-XLSX] §2.1.7` | Pass-Through (Patch) | — | Preserved on surgical patch; dropped if full-export fallback occurs |

---

## 5. Formula Serialization & Calculation (`ECMA-376 Part 1 §18.17` & `[MS-XLSX]`)

| Area | Spec Reference | Status | Verification | Notes & Gaps |
| --- | --- | --- | --- | --- |
| Standard A1 formulas & cached values (`<f>`, `<v>`) | `§18.3.1.40`, `§18.17` | Modeled | `U, V, F, X` | Leading `=` stripped on write; cached values written alongside `<f>` |
| Future function prefixes (`_xlfn.`, `_xlfn._xlws.`) | `[MS-XLSX] §2.2.2`, `§2.5.198.3` | Modeled | `U, V, F, X` | `ooxml_function_prefix` in `visi-core/src/core/xlsx/formula.rs` prefixes post-2007 functions (`XLOOKUP`, `SEQUENCE`, `LET`, `LAMBDA`, etc.) and `_xlfn._xlws.` (`FILTER`, `SORT`) without touching string literals or already-prefixed calls |
| Spill range operator (`A1#` -> `_xlfn.ANCHORARRAY(A1)`) | `[MS-XLSX] §2.2.2` | Modeled | `U, V, F, X` | Encoded by `encode_xlsx_spill_references` in `visi-core/src/core/parser.rs` |
| `LET` / `LAMBDA` parameter prefix (`_xlpm.`) | `[MS-XLSX] §2.2.2` | Modeled | `U, V, F, X` | `prefix_xlpm_parameters` in `visi-core/src/core/xlsx/formula.rs` prefixes bound `LET`/`LAMBDA` parameters and lexical references with `_xlpm.`; evaluator strips `_xlpm.` and `_xlfn.LAMBDA` on import |
| Dynamic array formulas (`t="array"`, `cm="1"`, `xl/metadata.xml` `XLDAPR`) | `[MS-XLSX] §2.1.7.34`, `§2.6.208` | Modeled | `U, V, F, X` | Emits `xl/metadata.xml` with `xda:dynamicArrayProperties` and links via `workbook.xml.rels` (`04_formulas_and_arrays.xlsx`) |
| Legacy CSE array formulas (`t="array"`, `ref`) | `§18.3.1.40` | Modeled | `U, V, F, X` | Covered by `04_formulas_and_arrays.xlsx` |
| Shared formulas (`t="shared"`, `si`, `ref`) | `§18.3.1.40` | Partial | `U` | Expanded into individual formulas on import (`worksheet.rs`); not re-compressed into shared formulas on export; untouched shared formulas in unedited cells survive surgical patch |
| Structured table references (`Sales[Price]`, `[@Qty]`, `[#Headers]`) | `§18.17.2.5` | Modeled | `U, V, F, X` | Parsed, evaluated, and serialized (`06_tables.xlsx`) |
| 3D sheet references (`Sheet1:Sheet3!A1`) | `§18.17.2.3` | **Gap (Unsupported)** | — | Not yet parsed/evaluated/serialized as 3D references |
| External workbook references (`[1]Sheet1!A1`, `xl/externalLinks/`) | `§18.14`, `§18.17.2.3` | **Gap (Pass-Through)** | — | `xl/externalLinks/` parts survive surgical patch; formula references to external workbooks are not modeled |
| Calculation chain invalidation (`xl/calcChain.xml`) | `§18.6` | Modeled | `U, V` | `invalidate_calc_chain` in `patch.rs` strips `xl/calcChain.xml`, its `workbook.xml.rels` relationship, and its `[Content_Types].xml` override whenever any formula cell or defined name changes on surgical patch; preserves `calcChain.xml` when only non-formula cells change |

---

## 6. Surgical Patch Preservation Boundaries (`patch.rs`)

Surgical patching (`OoxmlSourceState::patch_to_bytes`) preserves unmodeled parts, relationships, content types, `<extLst>`, and sibling XML elements **only when `can_patch_surgically` returns `true`**.

Currently, mutating any of the following forces a **full-export fallback** (which drops unmodeled parts/extensions from the source workbook):
- Adding, removing, or reordering sheets (`sheets.len()` or `sheet.id` mismatch)
- Modifying tables, comments, hyperlinks, data validations, or conditional formats on any sheet
- Modifying charts, pivot tables, or VBA modules

---

## Priority Backlog to Reach Full Compatibility

1. **Build a wild-file interoperability corpus** (`C`) testing surgical patch preservation of `extLst`, `mc:AlternateContent`, `docProps`, themes, drawings, slicers, sparklines, and Strict OOXML files from Excel, LibreOffice, Google Sheets, and `openpyxl`.
2. **Add `OpenXmlValidator` (`S`) and Excel/LibreOffice oracle (`O`)** verification stages to CI/release workflows.
3. **Narrow full-export fallback in `patch.rs`** so editing comments, hyperlinks, validations, tables, charts, or pivots patches surgically instead of regenerating the whole package.
4. **Add 3D sheet reference support** (`Sheet1:Sheet3!A1`) in parser, evaluator, and serializer.
