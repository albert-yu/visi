_The following content is LLM-generated._

## Recommended direction

Build a **native, full-rebuild XLSB writer** in `visi-core`, with strict feature admission: if an input contains unsupported semantic features, reject it rather than silently dropping them.

Use Excel-driven differential fuzzing as the primary compatibility gate.

## 1. Add an XLSB-specific fuzz pipeline first

Create `fuzz/fuzz_xlsb.py`, reusing `ExcelFuzzGenerator` and `DifferentialComparator` from `fuzz/fuzz_excel.py`.

For every seed:

1. Generate `source.xlsx`.
2. Ask Excel to save it as `source.xlsb` using file format `xlExcel12`/`50`.
3. Load and evaluate `source.xlsb` with visi.
4. Save both:
   - `visi_semantic.xlsx`, isolating XLSB import/evaluation.
   - `visi_out.xlsb`, testing the new writer.
5. Have Excel recalculate:
   - `source.xlsb` → `excel_oracle.xlsx`
   - `visi_out.xlsb` → `excel_visi_out.xlsx`
6. Compare:
   - `visi_semantic.xlsx` vs `excel_oracle.xlsx`
   - `excel_visi_out.xlsx` vs `excel_oracle.xlsx`
7. Reload `visi_out.xlsb` with visi and compare formulas, types, cached results, and metadata.

Failure artifacts should include:

```text
source.xlsx
source.xlsb
visi_semantic.xlsx
visi_out.xlsb
visi_reload.xlsx
excel_oracle.xlsx
excel_visi_out.xlsx
seed.json
```

Extend `ExcelDriver` with explicit conversion and calculation modes. Before relying on “no recalculation” conversion for cached-value checks, add a small macOS/Windows probe proving Excel preserves XLSB formula caches in manual-calculation mode.

## 2. Introduce feature profiles

The current generator includes tables, structured references, dynamic arrays, and modern functions. Native XLSB support should advance through explicit profiles:

- `basic`: scalar values, dates, errors, multiple sheets.
- `formula-core`: arithmetic, references, ranges, standard functions.
- `formula-modern`: `_xlfn`, dynamic arrays, spill and implicit intersection.
- `styles-layout`: cell styles, number formats, widths, heights.
- `tables`
- `charts`
- `pivots`
- `vba`
- `full`: existing unrestricted generator.

A profile may only graduate when a fresh Excel run passes cleanly and every prior failure has an Excel-free regression test.

## 3. Separate file-format dispatch from XLSX implementation

Refactor the import/export boundary rather than expanding `xlsx.rs` further:

```text
visi-core/src/core/excel.rs
visi-core/src/core/xlsb/
    mod.rs
    record.rs
    package.rs
    workbook.rs
    worksheet.rs
    formula.rs
    styles.rs
    table.rs
    chart.rs
    pivot.rs
    vba.rs
```

Key API changes:

- Content-detect XLSX versus XLSB in `WorkbookManager::load_bytes`.
- Add a public `WorkbookFormat` and `save_bytes_as(format)`.
- Preserve `save_bytes()` as XLSX for backward compatibility.
- Add `Error::Xlsb(String)`.
- Ensure any new public type is deliberately added to the curated re-export list.

The CLI should infer output format from `.xlsx` or `.xlsb`. Python `Workbook.save(path)` should do the same, while `save_bytes(format="xlsx")` remains explicit and defaults to XLSX.

## 4. Implement XLSB import

Use `calamine::Xlsb` or `open_workbook_auto_from_rs` for values and formulas. Calamine 0.36 already supports both for XLSB.

Do not run XLSX-specific helpers against XLSB:

- `import_cell_styles`
- `import_column_widths`
- `import_row_heights`
- `import_tables_from_zip`
- `import_charts_from_zip`
- `pivot_xlsx::import_pivot_tables`
- `vba_xlsx::import_vba_project`

Add binary readers for represented metadata as each feature profile is enabled.

Because the chosen policy is reject-on-loss, validate XLSB parts and records during load. Reject known unsupported features with a precise diagnostic before materializing a supposedly round-trippable workbook.

## 5. Build the BIFF12 writer in layers

### Package and records

Implement:

- BIFF12 variable-length record headers.
- OPC parts, content types, and relationships.
- `workbook.bin`
- `worksheet*.bin`
- `sharedStrings.bin`
- `styles.bin`
- document properties and theme parts.

Start with values only. Excel must open every generated file without repair warnings.

### Formula encoding

This is the critical path. Existing third-party Rust writers do not currently offer sufficient formula support.

Use `parser::parse_excel_formula()` and recursively serialize its `Expr` tree into postfix BIFF12 `Ptg` tokens:

- literals and errors
- unary/binary operators
- local and 3D references
- ranges
- fixed and variable-arity functions
- names and structured references
- modern formula extensions

Emit the appropriate cached-result record:

- `BrtFmlaNum`
- `BrtFmlaString`
- `BrtFmlaBool`
- `BrtFmlaError`

Maintain one spec-derived function-ID registry shared by tests and the encoder.

### Styles and layout

Deduplicate and serialize:

- fonts
- fills
- number formats
- cell XFs
- row heights
- column widths

Map exactly from `CellStyle`, rather than attempting to support formatting absent from the model.

## 6. Add represented workbook features incrementally

Implement in this order:

1. Excel Tables and structured references.
2. VBA by reusing the existing `vbaProject.bin` synthesis, with XLSB relationships and workbook code names.
3. Charts.
4. Pivot definitions and caches.

`pivot_xlsx.rs` and much of the chart code are OOXML-specific, so share the model and pivot-grid logic, not the serializers.

Unsupported items such as external links, data connections, conditional formatting, merged cells, or defined names should remain explicit load errors until represented by visi’s model.

## 7. Regression and robustness testing

For each minimized Excel mismatch:

- Check in a small `.xlsb` fixture for import bugs.
- Add Rust writer tests that generate XLSB and reopen it with calamine.
- Add round-trip assertions covering formulas, cached types, styles, and relationships.
- Add `cargo-fuzz` targets for:
  - BIFF12 record decoding.
  - Formula token decoding.
  - XLSB package import.

Also extend `fuzz/test_backend_parity.py` so bindings and CLI select the same format and produce equivalent XLSB output.

## Milestone gates

| Milestone | Exit criterion |
|---|---|
| Import | XLSX-generated and Excel-converted XLSB inputs produce identical visi snapshots |
| Basic writer | Excel opens generated XLSB without repair; scalar cells round-trip |
| Formula writer | Excel recalculation of `visi_out.xlsb` matches the source oracle |
| Styles/layout | Existing style/date-format fuzz cases survive XLSB round-trip |
| Tables | Structured-reference and table fuzz profile passes |
| Advanced features | XLSB variants of chart, pivot, and VBA harnesses pass |
| Release | Fresh full-profile run passes on Windows Excel; macOS run passes as secondary validation |

## Dependency recommendation

Use the Microsoft [MS-XLSB specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/acc8aa92-1f02-4167-99f5-84f9f676b95a) as authoritative.

Existing crates can provide implementation references, but should not initially become core dependencies:

- `xlsb-writer`: no formulas and Arrow-oriented.
- `rxlsb`: formula support is still planned.
- `justybase-spreadsheet`: explicitly preview-quality and lacks the required workbook fidelity.

The highest-risk work is formula-token serialization, followed by binary tables and pivots. The harness should therefore be built before the writer so every added BIFF12 record family receives immediate Excel validation.