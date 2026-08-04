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

### Financial functions (`ExcelFuzzGenerator.generate_financial_formula`)

The 27 TVM/depreciation Financial functions (`PV`/`FV`/`PMT`/`RATE`/`IRR`/
`XIRR`/`DDB`/etc. -- see `libvisi/src/core/finance.rs`) get their own
generator method rather than feeding into `generate_formula`'s recursive
`gen_expr` tree: their arguments have specific meanings (a rate has to
stay small and positive, a period has to stay within `[1, nper]`) that
arbitrary sub-expression substitution would violate far too often to be a
useful test. `create_fuzz_workbook` lays out a small "financial data"
block (a cashflow column, an aligned ascending-date column, a small-rates
"schedule" column) for the array-argument functions
(`NPV`/`IRR`/`MIRR`/`XNPV`/`XIRR`/`FVSCHEDULE`) to reference, since visi's
parser has no `{...}` array-literal syntax.

`_fin_rate()`'s range is deliberately realistic (0.1%-3% per period), not
wide -- an earlier version generated up to 20% per period, and at a high
per-period rate compounded over hundreds of periods, `(1+rate)^nper`
overflows into territory where computing the amortizing payment loses
essentially all `f64` precision (verified against arbitrary-precision
decimal arithmetic while chasing a real mismatch this generator found).
That's a real floating-point floor no formula rewrite escapes, so the fix
was to stop generating inputs no real financial instrument would have.

`PDURATION`/`RRI` (added in Excel 2013) are written with an explicit
`_xlfn.` prefix (`_xlfn.RRI(...)`) -- real Excel's own OOXML writer always
prefixes post-2007 functions this way, and a plain `RRI(...)` with no
prefix (what `openpyxl` writes by default) reads back as `#NAME?` in real
Excel even though the function exists. `evaluate_function` already strips
a leading `_xlfn.` before dispatch, so writing the prefix here is what a
real xlsx producer would do and keeps both sides consistent.

`IRR`/`XIRR`/`RATE` are Newton-Raphson root finds and a known residual
source of differential fuzzer failures even after all of the above: they
can return `#NUM!` on inputs real Excel's own (undocumented) solver
happens to converge on, and occasionally the reverse. `rate()` rejects
solutions that converge to the degenerate `-100%` boundary (Excel does
too); `irr()`/`xirr()` retry once from a `0.0` guess if the caller's guess
fails, a narrow fallback verified not to also introduce false positives
(finding a root Excel's own solver doesn't bother to find) the way an
earlier, broader multi-guess version did. Closing the rest of this gap
would mean reverse-engineering Excel's exact iterative algorithm, which is
out of scope for now.

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

- **Aggregation function mapping.** `visi`'s `Count` (counts any non-blank value) maps to Excel's `xlCount` (its *default* for text fields); `visi`'s `CountNumbers` (numeric-only) maps to `xlCountNums`. There is no separate `xlCountA` member in Excel's `XlConsolidationFunction` enum at all (confirmed via `Excel.sdef`) -- an early draft of this mapping assumed one existed and was wrong.
- **The "Parameter error (-50)" from issue #15 was a session-degradation artifact of the AppleScript driver, not a pivot-config bug.** The originally-suspected shape (a column used as both a `col_field` and the `filter_field`, filter selecting zero values, single-row source) was isolated and driven through real Excel a dozen-plus ways (same/different column, zero/non-zero selection, single/multi row, table/range source) early in a fresh Excel session -- every variant passed. The failure only appeared, 100% reproducibly and independent of pivot shape (even the simplest single-col-field, no-filter config triggered it), after roughly 20-30 consecutive `run VB macro` calls against one long-lived Excel process in the same fuzzing run; force-quitting and relaunching Excel made the identical failing config succeed immediately. `ExcelPivotDriver._restart_excel` now does that quit/relaunch automatically on any non-timeout AppleScript failure, not just on a hung/timed-out call.

Failure artifacts land under `fuzz_results/failures/pivot_fail_iter_<N>_seed_<SEED>/` (same shape as the formula fuzzer's, see below, with a `pivot_` prefix so the two don't collide).

---

## Chart Fuzzing (`fuzz_chart.py`)

Charts get their own script for the same reason pivot tables do: Excel must *actively construct* a chart object via its object model -- there's no XML shortcut the way `calculate` recalculates existing formulas -- and `fuzz_excel.py`'s `XLSXEvaluatedReader`/`DifferentialComparator` only understand cell values, so they're structurally blind to charts entirely. `fuzz_chart.py` instead:

1. Uses `openpyxl` to generate a small source data grid (one category column, one numeric column).
2. Picks a random chart configuration -- type, title, axis labels, legend visibility -- for an initial `visi chart add`, and a second, usually-different configuration for a follow-up `visi chart edit` (so `chart edit`, added alongside this fuzzer, gets differential coverage against real Excel too, not just `add`).
3. Builds the chart in `visi` via the actual `visi chart add`/`visi chart edit` CLI (`VisiChartDriver`), and in Excel by driving its chart object model directly via AppleScript on macOS or COM on Windows (`ExcelChartDriver`) -- built to the *final* target state in one call, since Excel has no separate "edit" step to mirror; only the resulting xlsx structure is compared.
4. Compares the two engines' resulting chart structure -- type, category/value ranges, title, axis labels, legend -- via `chart_xlsx_reader.read_charts`, a new `openpyxl`-based reader module (`ChartComparator`, not the cell-based comparator above).

```bash
python3 fuzz/fuzz_chart.py --driver mock --iterations 5                      # smoke-test the pipeline, no Excel needed
python3 fuzz/fuzz_chart.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20 --seed 1
```

### Unlike pivot caches, AppleScript chart creation works natively

`make new pivot cache at wb` fails outright against real Excel (see above), but the equivalent for charts does work -- with one undocumented quirk. `make new chart object at end of chart objects of <sheet>` (the form you'd expect from `Excel.sdef`, which lists `chart object` as an element of `sheet`) fails with the same generic "Parameter error (-50)"; the working form, found by manual trial, is `make new chart object at <sheet>` -- omitting `at end of chart objects of` entirely. Once the chart object exists, `chart wizard` (the AppleScript exposure of VBA's `Chart.ChartWizard`) reliably sets source data, chart type, title, axis titles, and legend in one call -- for every chart type except Pie, where passing `category title`/`value title` also raises -50 (pie charts have no axes), so those two parameters are omitted whenever the target type is Pie. No VBA-macro-template workaround (`BuildFuzzPivot.bas`'s approach) is needed for charts at all.

### Known limitations

- **One chart per iteration.** The comparator assumes exactly one chart per file to avoid chart-matching/ordering ambiguity; this matches `Chart`'s own single-series-per-chart model, so it isn't a meaningful coverage gap today.
- **Pie and Area charts never get randomly generated axis labels.** Beyond the Pie/`chart wizard` restriction above, a manual spike found Area chart axis titles set via `chart wizard` don't reliably read back through openpyxl's `_charts` in this Excel/openpyxl version, for reasons not investigated further. Rather than chase that gap, the generator (`ChartFuzzGenerator.AXIS_LABEL_TYPES`) only ever assigns xlabel/ylabel to column/bar/line/scatter charts.
- **openpyxl's Bar/Column distinction.** openpyxl represents both as a `BarChart`; `chart_xlsx_reader.py` disambiguates via `BarChart.type` (`"col"`/`"bar"`), mirroring `parse_chart_xml`'s own `<c:barDir>` handling.
- **`Chart.id`/`Chart.name` don't round-trip through xlsx** (pre-existing, unrelated to this fuzzer -- `import_xlsx_data` always regenerates a fresh id and name-by-position on import). `VisiChartDriver` looks the id up fresh via `chart list --json` immediately before editing rather than assuming it's stable across a save/reload it didn't itself trigger.

Failure artifacts land under `fuzz_results/failures/chart_fail_iter_<N>_seed_<SEED>/` (same shape as the formula fuzzer's, see below).

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
