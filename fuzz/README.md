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

- **Still open, and *not* root-causable from visi's side: reusing the same source column across two value fields is non-deterministic in real Excel.** `BuildFuzzPivot.bas`'s `ApplyValueFields` reuses `pt.PivotFields(name)` + `.Orientation = xlDataField` in a loop (not the `AddDataField` API Excel actually documents for this), and confirmed via direct A/B testing that this specific pattern doesn't reliably create two distinct data fields in real Excel: sometimes it does, and the second one's underlying field gets auto-renamed "`<Column>2`" (which `pivot.rs`'s `value_field_labels` now replicates when it happens, e.g. "Min of Amount2"); sometimes the second call instead silently overwrites the first field's aggregation function, or drops the requested one for something else entirely. Replaying the *exact* same field/config sequence through `visi`'s own CLI is stable and correct every time (verified directly), so the non-determinism is upstream in Excel's handling of this macro pattern, not in `visi`. Whoever picks this up next: switch `ApplyValueFields` to `pt.AddDataField(pt.PivotFields(name), label, function)`.
- **Still open, and heavily investigated without a root cause: real Excel's rendered `rowGrandTotals`/`colGrandTotals` occasionally comes out the *opposite* of what the config dict (and `BuildFuzzPivot.bas`'s straightforward `pt.RowGrand`/`pt.ColumnGrand` assignment from it) asked for.** Seen 3 times across unrelated configs (varying row/col field counts, with and without a filter field, 1 or 2 value fields) -- in each case Excel's own exported `pivotTable1.xml` had `rowGrandTotals`/`colGrandTotals` swapped relative to what the same run's config dict specified, while `visi`'s side (built from the identical dict) was correct. Ruled out as a `visi` bug via extensive direct A/B testing: replaying each failing config's *exact* field/create sequence by hand through `visi pivot create`/`add-field` (including the 2-column-field, filter-field, and multi-value-field shapes each failure used) reliably produces the correct `rowGrandTotals`/`colGrandTotals` every time -- and `fuzz_pivot.py`'s own `VisiPivotDriver.run()` (the code that actually issues `--no-grand-totals-row`/`--no-grand-totals-col`) was re-read character-by-character and matches. So the swap happens somewhere in Excel's own handling of `pt.ColumnGrand =`/`pt.RowGrand =` (or possibly `pt.RefreshTable`) when driven via this VBA/AppleScript path, not in anything this repo controls. Whoever picks this up next: try isolating it with the win32com driver on Windows, or add a `Debug.Print pt.RowGrand, pt.ColumnGrand` immediately after each assignment in `BuildFuzzPivot.bas` (this environment couldn't capture that output directly, since AppleScript's `run VB macro` doesn't surface the Immediate window).
- **One tiny-edge-case config threw an outright Excel-side error.** A column used as both a `col_field` and the `filter_field`, with the filter selecting zero values, over a single-row source, made the macro fail with "Parameter error (-50)". Not yet root-caused.
- **Aggregation function mapping.** `visi`'s `Count` (counts any non-blank value) maps to Excel's `xlCount` (its *default* for text fields); `visi`'s `CountNumbers` (numeric-only) maps to `xlCountNums`. There is no separate `xlCountA` member in Excel's `XlConsolidationFunction` enum at all (confirmed via `Excel.sdef`) -- an early draft of this mapping assumed one existed and was wrong.
- **A pre-existing, non-pivot-specific `openpyxl`/`calamine` interop gap**: `openpyxl` writes worksheet-to-table relationship `Target` attributes as absolute package paths (`/xl/tables/table1.xml`), which is valid per the OPC spec but which `calamine` (the crate `visi`'s xlsx importer uses) doesn't resolve -- it only special-cases `../`-relative targets, so an `openpyxl`-authored Excel Table silently imports as zero tables (confirmed independent of pivot tables: plain `visi table list` on an untouched `openpyxl` file reports none). `fuzz_pivot.py`'s generator works around this itself (`_fix_openpyxl_table_rels`, rewriting the relationship XML to a relative path after `openpyxl` saves) so `--source-mode table` iterations aren't silently broken, but the underlying gap is real and out of this harness's scope to fix in `visi` itself.

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
