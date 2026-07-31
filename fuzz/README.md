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
3. Builds a matching pivot table in `visi` via the `visi pivot` CLI (`create` + repeated `add-field`/`filter`) and in Excel by driving its native PivotTable object model (AppleScript on macOS, COM on Windows) -- `ExcelPivotDriver` in `fuzz_pivot.py`.
4. Reuses `XLSXEvaluatedReader`/`DifferentialComparator` from `fuzz_excel.py` unchanged to compare the two engines' materialized output cells, since both write plain literal values into the destination range.

```bash
python3 fuzz/fuzz_pivot.py --driver mock --iterations 5                      # smoke-test the pipeline, no Excel needed
python3 fuzz/fuzz_pivot.py --excel-path "/Applications/Microsoft Excel.app" --iterations 1 --seed 1
```

### Known caveats specific to this harness

- **AppleScript pivot construction is unverified.** Building a live PivotTable via AppleScript (as opposed to `fuzz_excel.py`'s driver, which only needs `calculate` over cells `visi` already computed) has not been piloted against real Excel in this environment -- see the note below on why. `ExcelPivotDriver`'s AppleScript path uses Excel's standard VBA-name-mirroring convention (`RefreshTable` -> `refresh table`, etc.), which is the right convention but unconfirmed against a live dictionary. **Before trusting results beyond a single-iteration pilot run**, inspect one generated case's `visi_out.xlsx` vs `excel_out.xlsx` by hand (e.g. via `openpyxl`) to confirm cell layout actually lines up -- see "Layout alignment" below. The `win32com` path uses the standard, well-documented VBA object model directly and should be far more reliable.
- **Layout alignment.** Excel's default new-pivot-table layout ("Compact Form", which merges row-field labels into a single indented column) does not structurally match `visi`'s flat one-column-per-field tabular grid (`PivotGrid.body_rows[i].row_labels`). The driver explicitly sets Tabular Form, subtotals-at-bottom, and disables label repetition to align the two -- if a real pilot run shows structural mismatches unrelated to actual values, check these settings first.
- **Aggregation function mapping.** `visi`'s `Count` (counts any non-blank value) maps to Excel's `xlCountA`; `visi`'s `CountNumbers` (numeric-only) maps to Excel's `xlCount`. Get this backwards and every count-based comparison will show a spurious, systematic mismatch.
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
