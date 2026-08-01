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
- **A separate, still-open layout gap surfaced while verifying the fix above: Excel's row/col header caption is literally "Row Labels"/"Column Labels", not the field name -- unlike visi's grid, which always shows the actual field name.** Investigated in depth: `BuildFuzzPivot.bas` sets `PivotField.LayoutForm = xlTabular` per field, and inspecting the actual exported `pivotTable1.xml` afterward confirms this genuinely has no effect -- the `<pivotField>`'s `compact` attribute stays at its (omitted-means-true) default regardless. The standard fix is `PivotTable`'s table-wide `RowAxisLayout`/`ColumnAxisLayout`/`SubtotalLocation` methods (which the win32com driver already uses directly) -- but calling these from VBA via `run VB macro` **hangs Mac Excel outright** (confirmed: wrapped each call individually in its own `On Error Resume Next`, and it still hung rather than raising a catchable error, requiring `killall "Microsoft Excel"` to recover; this environment also lacks the Screen Recording/Accessibility permissions that would let an agent see what's actually on screen during the hang to diagnose it further). Given that, `BuildFuzzPivot.bas` deliberately keeps the (known ineffective, but safe) per-field `LayoutForm` calls rather than the table-wide ones. This remains the actual cause behind most of the "wide-grid"/large-mismatch-count iterations previously attributed to an unknown cause -- any config with row or column fields hits it. Whoever picks this up next: either find a working Mac-VBA-safe way to trigger Tabular Form, or test the equivalent win32com path on Windows instead, where `RowAxisLayout` is already wired up and untested.
- **One tiny-edge-case config threw an outright Excel-side error.** A column used as both a `col_field` and the `filter_field`, with the filter selecting zero values, over a single-row source, made the macro fail with "Parameter error (-50)". Not yet root-caused.
- **Aggregation function mapping.** `visi`'s `Count` (counts any non-blank value) maps to Excel's `xlCount` (its *default* for text fields); `visi`'s `CountNumbers` (numeric-only) maps to `xlCountNums`. There is no separate `xlCountA` member in Excel's `XlConsolidationFunction` enum at all (confirmed via `Excel.sdef`) -- an early draft of this mapping assumed one existed and was wrong.
- **A pre-existing, non-pivot-specific `openpyxl`/`calamine` interop gap**: `openpyxl` writes worksheet-to-table relationship `Target` attributes as absolute package paths (`/xl/tables/table1.xml`), which is valid per the OPC spec but which `calamine` (the crate `visi`'s xlsx importer uses) doesn't resolve -- it only special-cases `../`-relative targets, so an `openpyxl`-authored Excel Table silently imports as zero tables (confirmed independent of pivot tables: plain `visi table list` on an untouched `openpyxl` file reports none). `fuzz_pivot.py`'s generator works around this itself (`_fix_openpyxl_table_rels`, rewriting the relationship XML to a relative path after `openpyxl` saves) so `--source-mode table` iterations aren't silently broken, but the underlying gap is real and out of this harness's scope to fix in `visi` itself.
- **A pre-existing calamine crash on zero-data-row tables**: exporting an Excel Table with a header row but zero data rows and reimporting it panics inside `calamine::Xlsx::table_by_name` ("invalid range bounds"). This was found via `libvisi/src/core/pivot.rs`'s Rust-side invariant fuzzer (`test_fuzz_pivot_random_invariants`), which works around it by keeping Table-sourced fuzz configs at >=1 data row (Range-sourced configs still cover the zero-row case directly against `compute_pivot`). Also out of scope to fix here.

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
