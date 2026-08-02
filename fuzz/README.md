# Differential Fuzzing Harness: `visi` vs. Microsoft Excel

This directory contains a differential fuzzing test harness designed to generate `.xlsx` spreadsheets containing arbitrary data and formulas, execute them in both **`visi`** and actual **Microsoft Excel**, and verify that evaluated results match cell-for-cell with complete feature parity.

---

## Architecture Overview

```
                        ┌───────────────────────────────┐
                        │   fuzz_excel.py (Generator)   │
                        └───────────────┬───────────────┘
                                        │
                                Creates source.xlsx
                                        │
                    ┌───────────────────┴───────────────────┐
                    ▼                                       ▼
    ┌───────────────────────────────┐       ┌───────────────────────────────┐
    │          visi eval            │       │        Microsoft Excel        │
    │ (Updates cached <v> XML tags) │       │ (AppleScript / COM Automation)│
    └───────────────┬───────────────┘       └───────────────┬───────────────┘
                    │                                       │
            Produces visi_out.xlsx                  Produces excel_out.xlsx
                    │                                       │
                    └───────────────────┬───────────────────┘
                                        │
                                        ▼
                        ┌───────────────────────────────┐
                        │    DifferentialComparator     │
                        │ (OpenXML Evaluated Reader)    │
                        └───────────────────────────────┘
                                        │
                        ┌───────────────┴───────────────┐
                        ▼                               ▼
                 [MATCH PASSED]                  [MISMATCH FAILED]
                                            (Saves reproducing files
                                             in fuzz_results/failures/)
```

---

## Setup & Requirements

### 1. Requirements
- Python 3.8+
- `openpyxl` (for generating `.xlsx` test files):
  ```bash
  pip install -r fuzz/requirements.txt
  ```
- Compiled `visi` binary (`cargo build --release`)
- Microsoft Excel (macOS or Windows) for actual Excel execution.

---

## Usage

### Run Differential Fuzzing

```bash
# macOS with Microsoft Excel installed in /Applications
python3 fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20

# Windows
python3 fuzz/fuzz_excel.py --driver win32com --iterations 50

# Custom CLI or script runner
python3 fuzz/fuzz_excel.py --driver cli --excel-path "/usr/local/bin/excel_runner" --iterations 10

# Mock Mode (runs test pipeline without invoking Excel binary)
python3 fuzz/fuzz_excel.py --driver mock --iterations 5
```

---

## Pivot Table Fuzzing (`fuzz_pivot.py`)

Pivot tables get their own script rather than a mode inside `fuzz_excel.py`, because the generation/execution pipeline is fundamentally different: Excel must *actively construct* a live PivotTable via its object model (there's no XML shortcut the way `calculate` recalculates existing formulas), and `openpyxl` cannot write pivot tables at all -- its `TableDefinition` is an 87-parameter raw XML mirror with no builder API, so authoring one by hand would mean re-implementing most of `pivot_xlsx.rs` in Python. `fuzz_pivot.py` instead:

1. Uses `openpyxl` (which handles plain data + Excel Tables fine) to generate a random source workbook with columns chosen to exercise grouping -- low-cardinality categories (with occasional blanks and same-value-different-case duplicates), a numeric-looking-text column, and numeric columns for aggregation.
2. Picks a random pivot configuration (row/col/value/filter fields, subtotal toggles, grand totals) as a plain dict.
3. Builds a matching pivot table in `visi` via the `visi pivot` CLI (`create` + repeated `add-field`/`filter`), and in Excel by invoking a one-time human-authored VBA macro (`fuzz/BuildFuzzPivot.bas`) through `run VB macro` on macOS, or driving Excel's PivotTable object model directly via COM on Windows -- `ExcelPivotDriver` in `fuzz_pivot.py`. See "AppleScript can't create pivot caches" below for why macOS needs the macro indirection.
4. Reuses `XLSXEvaluatedReader`/`DifferentialComparator` from `fuzz_excel.py` unchanged to compare the two engines' materialized output cells, since both write plain literal values into the destination range.

```bash
python3 fuzz/fuzz_pivot.py --driver mock --iterations 5                      # smoke-test the pipeline, no Excel needed
python3 fuzz/fuzz_pivot.py --excel-path "/Applications/Microsoft Excel.app" --iterations 1 --seed 1
```

### AppleScript can't create pivot caches -- confirmed, worked around

`make new pivot cache at wb` is declared in `Microsoft Excel.app/Contents/Resources/Excel.sdef` (extracted directly from the app bundle, since the `sdef` CLI tool needs a full Xcode install), but fails with a generic "Parameter error (-50)" against real Excel for every variant tried: bare, with properties, range object vs. text source data, different container forms. Reading (`count of pivot caches of wb`) works fine, only creation fails -- a real functional gap in Mac Excel's AppleScript support, not a syntax mistake. (Several other names *were* wrong on the first pass and are now fixed against the real dictionary: the per-field property is `pivot field orientation` not `orientation`, its enum values are `orient as row field` etc., there's no table-wide `RowAxisLayout`/`SubtotalLocation` -- those are per-field `layout form`/`layout subtotal location` -- subtotals toggle one function at a time via the `set subtotals` command, a ListObject's range is `range object` not `range`, and labeled command parameters use **space**-separated syntax, not colons.)

The workaround: `fuzz/BuildFuzzPivot.bas` is a VBA macro (using the same well-documented `PivotCaches.Create`/`CreatePivotTable`/`PivotFields` object model the win32com path uses directly) that a human pastes once into a macro-enabled workbook, `fuzz/pivot_macro_template.xlsm` (gitignored -- a compiled binary asset each developer creates locally, not generated). The AppleScript driver copies random data into that template via `openpyxl`'s `keep_vba=True`, then invokes the macro with `run VB macro "BuildFuzzPivot" arg1 ... arg9 ...`, which *is* a working AppleScript command. See `fuzz_pivot.py`'s module docstring for the exact one-time setup steps.

### Known caveats and open findings

- **visi didn't render filter/page fields as header rows -- FIXED.** Piloted against real Excel: Excel always shows a `FieldName | (All)`-or-`(Multiple Items)` row plus a blank spacer row above the grid when a filter field exists (confirmed empirically -- even a single selected value out of several still shows "(Multiple Items)", never the value's own name, since that's specific to the classic single-select page-field mode Excel no longer defaults to). `libvisi/src/core/pivot.rs`'s `compute_pivot`/`PivotGrid` now computes a `filter_rows: Vec<(String, String)>` the same way, `PivotGrid::grid_row_offset()`/`height()` account for it, `visi/src/engine.rs`'s `refresh_pivot_table` materializes it above the grid, and `pivot_xlsx.rs`'s native `<location>` (which Excel's own convention keeps scoped to just the row/col header + data grid, documented via `rowPageCount`/`colPageCount`) shifts accordingly on both export and re-import. Regression tests: `pivot.rs`'s `test_filter_field_state_label_all_vs_multiple_items`/`test_no_filter_fields_means_no_reserved_rows`, and `cli_tests.rs`'s `test_pivot_filter_field_materializes_as_header_row_above_grid`. Verified fixed via an isolated real-Excel pilot (no row/col fields, filter field only): the filter-row cell mismatches are gone.
- **Excel's row/col header caption is literally "Row Labels"/"Column Labels", not the field name -- FIXED, on visi's side.** `BuildFuzzPivot.bas` sets `PivotField.LayoutForm = xlTabular` per field, but the exported `pivotTable1.xml`'s `compact` attribute stays at its default regardless -- confirmed genuinely ineffective on Mac Excel. The standard fix, `PivotTable`'s table-wide `RowAxisLayout`/`ColumnAxisLayout`/`SubtotalLocation` methods, **hangs Mac Excel outright** when called via `run VB macro` (confirmed: wrapped each call in its own `On Error Resume Next` and it still hung, requiring `killall "Microsoft Excel"` to recover). Since Excel's own side can't be fixed here, `libvisi/src/core/pivot.rs`'s `compute_pivot` was reworked to match Excel's compact-form display directly: the outermost row field's header caption becomes the literal text "Row Labels" (deeper row fields keep their real name), a column-fields-non-empty pivot gets an extra prepended header row captioned "Column Labels", a grand-total column with multiple value fields shows "Total <value label>" instead of repeating "Grand Total" per value field, repeated adjacent header labels merge (blank after the first, matching Excel's merged-cell convention), and the corner/label-column reservation rules for the various row-fields-empty/col-fields-empty/multi-value-field combinations were derived and unit-tested against real Excel output for each shape. See the doc comments on `row_label_width` and the header-row-building block in `compute_pivot` for the exact rules, and `pivot.rs`'s `test_row_labels_caption_replaces_outermost_row_field_name`/`test_column_labels_row_prepended_and_deeper_row_field_keeps_its_name`/`test_grand_total_column_shows_total_prefixed_value_label_with_multiple_value_fields`/`test_flat_pivot_with_no_row_or_col_fields_has_no_reserved_label_column`/`test_no_row_fields_with_multiple_value_fields_has_no_reserved_label_column_either`/`test_multiple_value_fields_with_no_column_fields_share_one_header_row` for the regression coverage. This was the single biggest driver of the "wide-grid"/large-mismatch-count iterations previously reported here.
- **A field used as both a row/col field and a filter field -- FIXED.** Real Excel only lets a field occupy one PivotTable area at a time (setting `PivotField.Orientation` a second time *moves* the field, it doesn't duplicate it), but `WorkbookManager::add_pivot_field` (`visi/src/engine.rs`) used to just push onto the target area's `Vec` without evicting the field from wherever else it already was -- so a column added as a col field and *then* as a filter field ended up grouped by in both, a shape real Excel can't produce. Fixed to evict from Row/Column/Filter (in any combination) before adding, mirroring Excel; Value is the intentional exception, since Excel does allow the same source column to back multiple value fields simultaneously. Regression test: `cli_tests.rs`'s `test_pivot_field_area_reassignment_evicts_from_previous_area`.
- **Per-field subtotal toggles and pivot-wide grand-total toggles were silently discarded by the CLI -- FIXED.** Two separate bugs: (1) `visi pivot add-field` had no flag to turn a field's subtotal off at all, and `visi pivot create` had no flags for the grand-total row/column toggles either -- both are now exposed as `--no-subtotal`, `--no-grand-totals-row`, `--no-grand-totals-col`. (2) Even with the flag, `pivot_xlsx.rs`'s exporter used to suppress the `<item t="default"/>` subtotal marker for whichever field was *currently* the innermost field of its axis, regardless of its actual `subtotal` setting -- since every `visi pivot` CLI command round-trips the whole workbook through xlsx in a fresh process, a field's `subtotal: false` was reliably lost the moment it was (even briefly) the sole/innermost field on its axis, which is exactly the state it's in right after being added and before a second field joins it. Now written and read back unconditionally per field. Regression tests: `pivot.rs`'s `test_fuzz_pivot_random_invariants` (subtotal now asserted to round-trip exactly, not "reset to true"), `xlsx.rs`'s `test_xlsx_pivot_outer_field_subtotal_off_survives_round_trip`.
- **Case-variant text values (e.g. "East"/"east") grouped separately instead of merging -- FIXED.** `fuzz_pivot.py`'s generator deliberately mixes casings in its `Mixed` column to probe this. Two layered bugs: grouping was case-sensitive at all (fixed by comparing case-insensitively in `build_group_tree`), and even after that fix the casing a merged group displayed under was decided independently per branch of the *other* axis (so the same value could show as "EAST" nested under one Group and "east" under another) instead of using one canonical spelling field-wide the way Excel's pivot cache does -- fixed by canonicalizing case globally, per field, before grouping ever happens. Regression test: `pivot.rs`'s `test_case_variant_values_merge_using_globally_first_seen_casing`.
- **A blank group didn't sort last among otherwise-numeric-looking values, and numeric-looking *text* was sorted as if it were real numbers -- FIXED.** `NumStr`, the generator's forced-text numeric-looking column, exists specifically to probe this ambiguity. `(blank)` now always sorts last regardless of the field's numeric/text order (previously a single blank among numeric values forced the *whole* field into alphabetical order, and even then blank still landed first, not last). Separately, `sort_group_entries` used to guess numeric-vs-text per field by trying to `f64::parse` the string keys, which can't distinguish a real number from text that merely looks numeric -- fixed by deciding this from the original `ResultData` variants (`field_is_numeric`) before values get collapsed to strings for grouping. Regression tests: `pivot.rs`'s `test_blank_group_sorts_last_even_among_numeric_siblings`, plus `field_is_numeric`'s use throughout `compute_pivot`.
- **A sparse row/column intersection with zero underlying records rendered as a computed 0 or `#DIV/0!` instead of blank -- FIXED.** Excel shows a genuinely blank cell for a row-group/col-group combination that simply has no matching records, for every aggregation kind (even `Sum`/`Count`, which have an "obvious" zero answer) -- `aggregate()` now returns `ResultData::None` immediately when its input list is empty, before dispatching to any aggregation-specific logic. Regression test: `pivot.rs`'s `test_empty_row_col_intersection_renders_blank_not_zero_or_error`.
- **Still open, and *not* root-causable from visi's side: reusing the same source column across two value fields is non-deterministic in real Excel.** `BuildFuzzPivot.bas`'s `ApplyValueFields` reuses `pt.PivotFields(name)` + `.Orientation = xlDataField` in a loop (not the `AddDataField` API Excel actually documents for this), and confirmed via direct A/B testing that this specific pattern doesn't reliably create two distinct data fields in real Excel: sometimes it does, and the second one's underlying field gets auto-renamed "`<Column>2`" (which `pivot.rs`'s `value_field_labels` now replicates when it happens, e.g. "Min of Amount2"); sometimes the second call instead silently overwrites the first field's aggregation function, or drops the requested one for something else entirely. Replaying the *exact* same field/config sequence through `visi`'s own CLI is stable and correct every time (verified directly), so the non-determinism is upstream in Excel's handling of this macro pattern, not in `visi`. Whoever picks this up next: switch `ApplyValueFields` to `pt.AddDataField(pt.PivotFields(name), label, function)`.
- **Still open, and heavily investigated without a root cause: real Excel's rendered `rowGrandTotals`/`colGrandTotals` occasionally comes out the *opposite* of what the config dict (and `BuildFuzzPivot.bas`'s straightforward `pt.RowGrand`/`pt.ColumnGrand` assignment from it) asked for.** Seen 3 times across unrelated configs (varying row/col field counts, with and without a filter field, 1 or 2 value fields) -- in each case Excel's own exported `pivotTable1.xml` had `rowGrandTotals`/`colGrandTotals` swapped relative to what the same run's config dict specified, while `visi`'s side (built from the identical dict) was correct. Ruled out as a `visi` bug via extensive direct A/B testing: replaying each failing config's *exact* field/create sequence by hand through `visi pivot create`/`add-field` (including the 2-column-field, filter-field, and multi-value-field shapes each failure used) reliably produces the correct `rowGrandTotals`/`colGrandTotals` every time -- and `fuzz_pivot.py`'s own `VisiPivotDriver.run()` (the code that actually issues `--no-grand-totals-row`/`--no-grand-totals-col`) was re-read character-by-character and matches. So the swap happens somewhere in Excel's own handling of `pt.ColumnGrand =`/`pt.RowGrand =` (or possibly `pt.RefreshTable`) when driven via this VBA/AppleScript path, not in anything this repo controls. Whoever picks this up next: try isolating it with the win32com driver on Windows, or add a `Debug.Print pt.RowGrand, pt.ColumnGrand` immediately after each assignment in `BuildFuzzPivot.bas` (this environment couldn't capture that output directly, since AppleScript's `run VB macro` doesn't surface the Immediate window).
- **One tiny-edge-case config threw an outright Excel-side error.** A column used as both a `col_field` and the `filter_field`, with the filter selecting zero values, over a single-row source, made the macro fail with "Parameter error (-50)". Not yet root-caused.
- **Aggregation function mapping.** `visi`'s `Count` (counts any non-blank value) maps to Excel's `xlCount` (its *default* for text fields); `visi`'s `CountNumbers` (numeric-only) maps to `xlCountNums`. There is no separate `xlCountA` member in Excel's `XlConsolidationFunction` enum at all (confirmed via `Excel.sdef`) -- an early draft of this mapping assumed one existed and was wrong.
- **A pre-existing, non-pivot-specific `openpyxl`/`calamine` interop gap**: `openpyxl` writes worksheet-to-table relationship `Target` attributes as absolute package paths (`/xl/tables/table1.xml`), which is valid per the OPC spec but which `calamine` (the crate `visi`'s xlsx importer uses) doesn't resolve -- it only special-cases `../`-relative targets, so an `openpyxl`-authored Excel Table silently imports as zero tables (confirmed independent of pivot tables: plain `visi table list` on an untouched `openpyxl` file reports none). `fuzz_pivot.py`'s generator works around this itself (`_fix_openpyxl_table_rels`, rewriting the relationship XML to a relative path after `openpyxl` saves) so `--source-mode table` iterations aren't silently broken, but the underlying gap is real and out of this harness's scope to fix in `visi` itself.
- **A calamine crash on zero-data-row tables -- FIXED.** Exporting an Excel Table with a header row but zero data rows and reimporting it used to panic inside `calamine::Xlsx::table_by_name` ("invalid range bounds"). Found via `libvisi/src/core/pivot.rs`'s Rust-side invariant fuzzer (`test_fuzz_pivot_random_invariants`, which special-cased Table-sourced configs to `num_rows >= 1` to work around it). Confirmed still present in calamine's latest release (0.36.1), not just the pinned 0.26 -- a version bump wouldn't have fixed it -- so `calamine` is now vendored and patched at `vendor/calamine/` (see `vendor/calamine/PATCHES.md` for the root cause and fix, wired in via `[patch.crates-io]` in the workspace root `Cargo.toml`). Fixing the panic alone surfaced a second issue: calamine's empty `Range` has no `start()`/`end()`, which made `xlsx.rs`'s table-import code silently *drop* the table instead of crashing -- fixed by preferring the table's own declared `ref` XML attribute (already parsed for header/totals flags) as the source of truth for a table's position, rather than reconstructing bounds from calamine's data range at all. The invariant fuzzer's workaround was removed (`num_rows` now covers 0 for Table-sourced configs too), and `xlsx.rs`'s `test_xlsx_zero_data_row_table_import_export_cycle` is a permanent regression test for the exact repro.

Failure artifacts land under `fuzz_results/failures/pivot_fail_iter_<N>_seed_<SEED>/` (same shape as the formula fuzzer's, see below, with a `pivot_` prefix so the two don't collide).

---

## Key Considerations for Excel Parity

When building full feature parity between `visi` and Microsoft Excel for formula evaluation and file import/export, several critical edge cases and subtle behaviors must be addressed:

### 1. Cached Formula Values (`<v>` openxml tags)
Excel relies on cached formula evaluation results written inside sheet XML tags (`<v>15.0</v>`).
- When programmatic libraries write `.xlsx` files, formulas are written to `<f>` tags without cached `<v>` values.
- When `visi eval` runs, it must compute formulas and save updated `<v>` tags.
- When Microsoft Excel recalculates and saves, it writes computed `<v>` tags.
- The fuzz harness extracts `<v>` values directly from the zipped XML structures to compare evaluated results cleanly.

### 2. OpenXML Structural Metadata vs. Semantic Content
Comparing raw `.xlsx` zip files or XML structures directly byte-for-byte will **always fail** because:
- Excel injects OS local timestamps, printer settings, recalculation IDs (`calcPr calcId="..."`), sheet relationship IDs, and custom namespace prefixes (`r:id`).
- Content comparison **must be semantic**: checking cell values, error codes (`#DIV/0!`, `#VALUE!`, `#REF!`, `#N/A`, `#NUM!`, `#NAME?`, `#NULL!`), booleans, and float tolerances.

### 3. Floating-Point Precision & Rounding (IEEE 754)
- Excel uses IEEE 754 double precision but applies 15-digit display precision heuristics.
- Minor floating-point bit differences (e.g. `0.30000000000000004` vs `0.3`) must be compared using relative and absolute tolerances (`math.isclose(v1, v2, rel_tol=1e-7, abs_tol=1e-7)`).

### 4. Excel Date System & The 1900 Leap Year Bug
- Excel represents dates as serial numbers (e.g. `45000.5`).
- Excel intentionally inherited Lotus 1-2-3's bug where `1900-02-29` (serial number `60`) is treated as a valid leap year date, shifting all pre-March 1, 1900 dates by 1 day.

### 5. Dynamic Arrays vs. Legacy Implicit Intersection
- Modern Excel (Excel 365 / 2021+) supports dynamic arrays (`=A1:A5 * 2`) which spill into adjacent cells.
- Modern functions store formula names prefixed with `_xlfn.` (e.g., `_xlfn.XLOOKUP`, `_xlfn.ANCHORARRAY`).

### 6. Headless Automation & Modal Dialogs
- In automated testing, Excel dialogs (e.g., Circular Reference Warnings, Privacy Alerts, File Corruption Warnings) can block headless execution.
- The harness suppresses alerts (`display alerts to false` in AppleScript / `DisplayAlerts = False` in COM).

---

## Reproducing & Debugging Failures

When a test iteration fails, the harness automatically creates a reproduction folder under `fuzz_results/failures/fail_iter_<N>_seed_<SEED>/`:

```
fuzz_results/failures/fail_iter_3_seed_48291/
├── source.xlsx      # Original generated workbook before evaluation
├── visi_out.xlsx    # Workbook evaluated and saved by visi
└── excel_out.xlsx   # Workbook evaluated and saved by Microsoft Excel
```

You can inspect `source.xlsx` or run `visi eval source.xlsx --output test.xlsx` to debug formula evaluation discrepancies directly.
