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
    │   visi  (visi_core bindings   │       │        Microsoft Excel        │
    │      in-process, or the CLI)  │       │ (AppleScript / COM Automation)│
    │ (Updates cached <v> XML tags) │       │                               │
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
- Python 3.9+, from the project venv — **not** system Python, since
  `maturin develop` installs into whichever venv is active:
  ```bash
  source fuzz/venv/bin/activate
  pip install -r fuzz/requirements.txt
  ```
- The `visi_core` bindings (see below), and/or a compiled `visi` binary
  (`cargo build --release`)
- Microsoft Excel (macOS or Windows) for actual Excel execution.

### 2. In-process bindings (recommended)

The fuzzers drive visi through `visi-python`, a pyo3 extension module binding
`visi-core` directly, rather than spawning the `visi` CLI once per operation.
That matters most for `fuzz_pivot.py`, which used to run one process per pivot
field — building a pivot table cost 5–8 full `.xlsx` load/save cycles.

```bash
source fuzz/venv/bin/activate
maturin develop -m visi-python/Cargo.toml --release
```

Rebuild after any change to `visi-core`, or you will be fuzzing a stale engine.

Each fuzzer takes `--backend {auto,bindings,subprocess}`. The default `auto`
uses the bindings when `import visi_core` succeeds and falls back to the CLI
with a warning otherwise. The run banner prints which was chosen — check it
before reading timings.

**Use `--backend subprocess` to triage a crash.** Under `bindings` the engine
shares this process: a Rust panic surfaces as a catchable `PanicException`, but
an abort or a stack overflow (plausible — the formula parser is recursive
descent and the generator emits deeply nested expressions) takes the whole run
down and loses every iteration's progress. Under `subprocess` it costs one
iteration and still saves the reproducing files.

### Bindings/CLI equivalence

The two backends must stay observationally identical, and they duplicate a
little logic to do it (`edit_chart`'s clear-vs-set flags, `add_pivot_field`'s
post-add subtotal mutation). Nothing else would notice that drifting:

```bash
pytest fuzz/test_backend_parity.py visi-python/tests/
```

It runs both backends over freshly generated workbooks at fixed seeds, plus any
`fuzz_results/failures/*/source.xlsx` lying around locally, and diffs the parsed
output — content, not bytes, since `docProps/core.xml` carries a timestamp and
ids are random. The generated seeds are the real coverage; `fuzz_results/` is
gitignored and usually empty.

One pivot state is reachable only through the bindings API:
`Workbook.set_pivot_filter(name, col, [])` selects *nothing*, which `visi pivot
filter` cannot express (it takes a non-empty comma list or `--clear`). It is
not fuzzed, because real Excel cannot represent it either — Excel refuses to
hide a page field's last visible `PivotItem`, so `BuildFuzzPivot.bas` falls back
to leaving the field unfiltered at `(All)`. `PivotFuzzGenerator` therefore
always selects at least one value, and both driver backends leave an empty
selection unapplied if one reaches them anyway. The engine behavior itself is
asserted directly, in `test_empty_filter_selection_is_bindings_only`.

### Why the oracle still parses the written `.xlsx`

`XLSXEvaluatedReader` reads real `.xlsx` bytes, never the bindings' in-memory
cells, and `read_evaluated_cells_bytes` exists only to skip a redundant *disk
read* of bytes visi just produced — it is the same parser on the same bytes.

Reading values out of the engine instead would stop exercising
`export_xlsx_data`: the cached `<v>` tags, shared strings, `t="e"` error cells,
and pivot's hand-rolled `inject_pivot_tables` zip rewriting. That is a large
share of what this harness exists to check. For the same reason the pivot and
chart drivers still round-trip the workbook between mutations — in memory now,
via `Workbook.roundtrip()`, but through the identical import/export code the
CLI ran.

`Workbook.get_cell` is for *minimizing* a failure, never for deciding one.

---

## Usage

### Run Differential Fuzzing

```bash
# macOS with Microsoft Excel installed in /Applications
python fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20

# Windows
python fuzz/fuzz_excel.py --driver win32com --iterations 50

# Custom CLI or script runner
python fuzz/fuzz_excel.py --driver cli --excel-path "/usr/local/bin/excel_runner" --iterations 10

# Smoke mode -- no Excel, no comparison. See below.
python fuzz/fuzz_excel.py --driver mock --iterations 5
```

### `--driver mock` is a smoke test, not a weak oracle

There is no Excel in mock mode, so **nothing is compared**. Each iteration
generates a workbook, runs visi over it, and checks only that visi wrote an
`.xlsx` that parses and has content. It exits 0 unless something crashed or
produced unreadable output, and writes no failure artifacts otherwise.

That is deliberately less than it used to do, because what it used to do was
not meaningful. Mock copied the *unevaluated* source workbook and compared
visi against it as though it were Excel's answer. openpyxl writes no cached
`<v>` for a formula cell, so all 360 formula cells of a default 530-cell grid
read as `None` on the "Excel" side: a guaranteed 100% mismatch, exit 1 on every
run, and one failure-artifact directory per iteration. The genuine signal --
did visi crash? did it emit a corrupt file? -- was invisible underneath it.

So mock is for what this README always claimed: exercising the pipeline
(generate → visi → read → report) on a machine with no Excel automation, and
hunting crashes over a large volume of generated formulas. Both work without an
oracle. Neither works with a fake one.

`--driver auto` resolves to mock on any platform that is not macOS or Windows,
so this is also what a Linux run does.

### Financial functions (`ExcelFuzzGenerator.generate_financial_formula`)

The 27 TVM/depreciation Financial functions (`PV`/`FV`/`PMT`/`RATE`/`IRR`/
`XIRR`/`DDB`/etc. -- see `visi-core/src/core/finance.rs`) get their own
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

### Reverse-engineering the IRR/XIRR/RATE solver gap (`reverse_engineer_financial.py`)

Rather than guessing at Excel's exact iterative algorithm, `reverse_engineer_financial.py`
grades a grid of ~70 candidate Newton-Raphson variants (closed-form TVM
derivative -- the standard OpenOffice-lineage formulation used by
`formulajs` -- vs. visi's current numeric/central-difference derivative,
crossed with several `(epsilon, max_iter, error-on-non-convergence,
retry-from-zero-guess)` combinations) against real Excel's actual output,
using cashflow/RATE inputs deliberately chosen to sit near the convergence
boundary rather than typical random-fuzz inputs: multi-sign-change
cashflows with more than one real root (including the classic
`[-10, 21, -11]` dual-root case, whose second root is the trivial `r=0`
since the cashflows sum to exactly zero), a wide sweep of starting guesses
per case, RATE inputs pushed toward the `-100%` floor, and irregular/
out-of-order XIRR dates.

```bash
python3 fuzz/reverse_engineer_financial.py --driver mock                     # pipeline smoke-test, no Excel
python3 fuzz/reverse_engineer_financial.py --excel-path "/Applications/Microsoft Excel.app" --seed 1
```

It writes every case as one formula into a workbook, evaluates it with both
`visi` and real Excel (via the same `ExcelDriver`/`VisiDriver`/
`XLSXEvaluatedReader` drivers `fuzz_excel.py` uses -- sheet cells are
looked up by parsing `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` for
the real name -> `sheetN.xml` mapping, since the reader keys cells by the
latter), then evaluates every candidate variant in pure Python against the
same inputs and ranks them by agreement rate with Excel. Findings from a
real run (`--seed 1`, one Excel installation, not treated as universal
constants -- rerun to confirm before relying on exact percentages):

- **Reverse-Engineered Solver Mechanics & Key Discoveries**:
  - **Step Halving on Domain Boundaries**: When Newton-Raphson steps $\Delta r = -f(r)/f'(r)$ would push $r + \Delta r \le -0.9999$, real Excel does not fail immediately or jump into complex/NaN space -- it performs step-halving ($\Delta r \leftarrow \Delta r / 2$) up to 50 times to keep iterates inside $(r > -1)$.
  - **Expanded Iteration Budget (`MAX_ITER = 200`)**: Long-duration loans (e.g., `nper=120`, `guess=2.0`) require >100 iterations to descend from initial overshoots. Expanding `MAX_ITER` to 200 resolved these cases.
  - **Monotonic Non-Positive Return Rule for IRR**: For cashflow streams with initial outlay $v_0 < 0$ and $v_i \ge 0$ ($i \ge 1$), if $\sum v_i \le 0$, no positive IRR solution exists and real Excel returns `#NUM!`. Enforcing this reached **100.0% parity on IRR** (168/168 cases).
  - **Asymptotic High-Rate Initial Guess Fallback for XIRR**: For pathological multi-sign or high-rate cash flows, Excel uses the leading-term asymptotic closed-form estimate $r_{\text{asymp}} = (v_1 / |v_0|)^{365/d_1} - 1$ as a fallback starting guess when Newton iterations diverge or hit trivial $r=0$ roots, achieving **100.0% parity on XIRR** (98/98 cases).
  - **Divergence Step Bounds & Negative Guess Annuity Rules for RATE**: Initial Newton steps with $|\Delta r| > 4.0$ indicate wild overshoots, and negative guesses on long-term annuities-due ($nper \ge 36$) with total positive return return `#NUM!` in Excel, achieving **100.0% parity on RATE** (340/340 cases).
  - **Perfect Suite Parity**: Across all 606 adversarial boundary test cases, overall agreement between `visi` and Microsoft Excel reached **100.0%** (606/606).

Full per-case, per-variant results land in
`fuzz_results/financial_reverse_engineering/report.json` (gitignored) for
further offline analysis; the console output prints, per function, visi's
current agreement rate with Excel as a baseline, the top-scoring candidate
variants, and the worst remaining mismatches for the best one.

---

## Pivot Table Fuzzing (`fuzz_pivot.py`)

Pivot tables get their own script rather than a mode inside `fuzz_excel.py`, because the generation/execution pipeline is fundamentally different: Excel must *actively construct* a live PivotTable via its object model (there's no XML shortcut the way `calculate` recalculates existing formulas), and `openpyxl` cannot write pivot tables at all -- its `TableDefinition` is an 87-parameter raw XML mirror with no builder API, so authoring one by hand would mean re-implementing most of `pivot_xlsx.rs` in Python. `fuzz_pivot.py` instead:

1. Uses `openpyxl` (which handles plain data + Excel Tables fine) to generate a random source workbook with columns chosen to exercise grouping -- low-cardinality categories (with occasional blanks and same-value-different-case duplicates), a numeric-looking-text column, and numeric columns for aggregation.
2. Picks a random pivot configuration (row/col/value/filter fields, subtotal toggles, grand totals) as a plain dict.
3. Builds a matching pivot table in `visi` via the `visi pivot` CLI (`create` + repeated `add-field`/`filter`), and in Excel by invoking a one-time human-authored VBA macro (`fuzz/BuildFuzzPivot.bas`) through `run VB macro` on macOS, or driving Excel's PivotTable object model directly via COM on Windows -- `ExcelPivotDriver` in `fuzz_pivot.py`. See "AppleScript can't create pivot caches" below for why macOS needs the macro indirection.
4. Reuses `XLSXEvaluatedReader`/`DifferentialComparator` from `fuzz_excel.py` unchanged to compare the two engines' materialized output cells, since both write plain literal values into the destination range.

```bash
python fuzz/fuzz_pivot.py --driver mock --iterations 5                       # smoke test, no Excel and no comparison
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
python fuzz/fuzz_chart.py --driver mock --iterations 5                       # smoke test, no Excel and no comparison
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

## VBA Execution Probe (`vba_probe.py`)

Not a fuzzer -- a fixed, deterministic feasibility check, and the empirical
basis for the VBA testing plan in [`docs/vba-macro-support.md`](../docs/vba-macro-support.md)
(GitHub issue #46). There is no VBA interpreter in `visi-core` yet, so there
is nothing to run differentially; this establishes that there *could* be.

```bash
cargo build --release
source fuzz/venv/bin/activate
python fuzz/vba_probe.py            # 4 checks, ~15s against real Excel
python fuzz/vba_probe.py --demo-hang  # + reproduce the modal-dialog hang
```

The thing it proves is non-obvious: Excel for Mac's AppleScript dictionary
exposes **no** VBProject object, so no automation path can put a macro *into*
a workbook on macOS. `visi macro add` writes the module into `vbaProject.bin`
at the file-format level instead, before Excel opens the file, and Excel then
runs it as an ordinary macro via `run VB macro`. That is the only reason a
VBA differential fuzzer is possible here without a Windows COM host.

Three findings that constrain any future `fuzz_vba.py`:

- **`run VB macro` returns typed values straight to AppleScript** (`OK|Double|42`), so
  results don't have to be routed through cells and a file read.
- **Trapped runtime errors come back as structured text** (`ERR|11|Division by zero`),
  making `Err.Number` directly comparable against an interpreter's.
- **An *un*trapped runtime error hangs the automation bridge.** `set display
  alerts to false` does not suppress the modal run-time-error dialog; the
  `osascript` call never returns and Excel must be SIGKILLed (see §6 below).
  Every generated macro must therefore be invoked through an `On Error GoTo`
  wrapper -- verified to catch errors raised anywhere down the call stack.
  This is a correctness requirement of the harness, not a nicety.

Also worth noting: `vba_probe.py` demonstrates that `fuzz_pivot.py`'s one-time
manual "paste `BuildFuzzPivot.bas` into the VBA editor and Save As
`pivot_macro_template.xlsm`" setup step could now be replaced by a single
`visi macro add` invocation.

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
