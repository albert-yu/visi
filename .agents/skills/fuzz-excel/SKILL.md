---
name: fuzz-excel
description: Run a differential fuzz harness from fuzz/ against a real copy of Microsoft Excel, repeating until a full run of N iterations passes clean, and turning every mismatch into a regression test that needs no Excel. Use when asked to fuzz, run the fuzz tests against Excel, chase down a fuzz failure, or fuzz until clean.
---

# Fuzzing against real Excel

The harnesses in `fuzz/` drive **real Microsoft Excel** and compare it against
`visi-core` cell-for-cell. This skill is the loop around them: run N iterations,
and for each failure, find the root cause, decide *which engine is right*, fix
or document, and leave behind a **Rust unit test that reproduces the case
without Excel** — because CI has no Excel and a fuzz finding that only lives in
`fuzz_results/` is a finding that comes back.

The loop does not end at "I explained the failure". It ends when a **fresh full
run of N iterations passes** with no failures.

## Ask first, if not already given

- **Which harness** (table below). If the user just said "the fuzz tests",
  default to `fuzz_excel.py` — the formula-evaluation one.
- **How many iterations** must pass. Default 20 for `fuzz_excel.py`, 200 for the
  VBA ones (their cases are much cheaper).

## Preflight

Every one of these matters; skipping one produces a confusing failure, not an
obvious one.

First check which host platform you are on, then choose the Excel driver from
that. On Windows, use `--driver win32com` and do **not** pass the macOS Excel
path. On macOS, use AppleScript and pass the application path explicitly with
`--excel-path "/Applications/Microsoft Excel.app"`.

```bash
uname -s                                             # MINGW*/MSYS*/CYGWIN* = Windows, Darwin = macOS
source fuzz/venv/bin/activate                       # macOS/Linux venv; never system python
source fuzz/venv/Scripts/activate                   # Windows/Git Bash venv; never system python
pip install -r fuzz/requirements.txt                # first run only
maturin develop -m visi-python/Cargo.toml --release # rebuild after ANY visi-core change
cargo build --release                               # only needed for --backend subprocess

# Confirm the oracle/driver for this host:
python - <<'PY'
import platform
print(platform.system())
PY
# Windows: use --driver win32com
# macOS:  ls "/Applications/Microsoft Excel.app"
```

`maturin develop` installs into the **active venv**, and the bindings backend is
what the harness uses by default — after every fix to `visi-core`, rebuild them
before re-running or you will be re-measuring the old engine and chasing a
failure you already fixed. If the Python venv architecture and Rust default
target differ on Windows, pass the Python architecture explicitly, e.g.
`maturin develop -m visi-python/Cargo.toml --release --target x86_64-pc-windows-msvc`.

A first-ever AppleScript run on macOS needs a one-time interactive
automation-permission grant. If Excel never responds and no permission dialog is
visible, ask the user to run the command themselves with a leading `!` so they
can click through it.

## The harnesses

| Script | What it compares | Failure artifacts |
| --- | --- | --- |
| `fuzz_excel.py` | formula evaluation, cell for cell | `fuzz_results/failures/fail_iter_<N>_seed_<SEED>/` — `source.xlsx`, `visi_out.xlsx`, `excel_out.xlsx` |
| `fuzz_chart.py` | chart definitions round-tripped | `chart_fail_iter_<N>_seed_<SEED>/` |
| `fuzz_pivot.py` | pivot grids and the hand-rolled pivot XML | `pivot_fail_iter_<N>_seed_<SEED>/` |
| `fuzz_vba.py` | VBA execution + the cells a macro wrote | `vba_exec_case_<N>/` — `source.bas`, `verdicts.txt` |
| `fuzz_vba_parse.py` | does visi's parser accept what Excel compiles | `vba_parse_<label>/` |

All of them exit non-zero when anything failed and take `--iterations`,
`--seed`, `--excel-path`, `--driver` and `--output-dir`. **Always record the
reproduction handle**, but note it differs by harness: `fuzz_excel.py`,
`fuzz_chart.py` and `fuzz_pivot.py` print a per-iteration seed next to each
verdict, while the two VBA fuzzers seed the whole *run* (`--seed`, random when
omitted) and identify failures by case number. So a VBA failure is reproduced
by re-running with that run's seed, and the artifact directory is the durable
record — the case number alone means nothing against a fresh seed.

```bash
# Windows:
python fuzz/fuzz_excel.py --driver win32com --iterations 20

# macOS:
python fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20

python fuzz/fuzz_excel.py --seed 48291 --iterations 1     # reproduce one iteration exactly; add the same driver/path as above
python fuzz/fuzz_excel.py --driver mock --iterations 5    # pipeline smoke test, NOT an oracle
```

`--driver mock` compares nothing. Use it to check the harness still runs after
editing it; never report a mock run as a passing fuzz run.

## The loop

1. Run the harness for the requested iteration count against real Excel.
2. If it exits 0 with `Failed : 0`, report and stop.
3. Otherwise take the **first** failing seed and work it end to end (triage
   below). Fix or document, add the Excel-free test, and confirm
   `cargo test -p visi-core` passes.
4. Re-run **the full N iterations from scratch** — new seeds, not just the one
   that failed. A fix that changes shared coercion or rounding code routinely
   moves a different family, and re-running only the old seed hides that.
5. Repeat until a full clean run. Report each round's seed and what it was.

Rebuild the bindings (`maturin develop ... --release`) between step 3 and step 4.

If the same root cause keeps resurfacing across rounds and cannot be fixed
inside the session's scope, say so explicitly with the seeds rather than
lowering the iteration count to get a green run.

## Triage: which engine is wrong?

**Excel is not automatically right.** `docs/excel-discrepancies.md` already lists
cases where visi is measurably more accurate, and "fixing" visi to match Excel
there is a regression. Work in this order:

1. **Check `docs/excel-discrepancies.md` first.** If the case is already listed,
   the harness is supposed to be excluding it — the finding is that the
   exclusion has a hole, so tighten the generator's exclusion (they live as
   inline comments next to the function lists in `fuzz_excel.py`) rather than
   touching the engine.
2. **Minimize.** Reduce to the smallest grid or expression that still differs.
   - formulas: cut the grid down and re-check with `visi eval`
   - VBA: `python fuzz/vba_expr_probe.py -e 'a = 1 :: a + 1'` runs one
     expression through both engines side by side — the fastest reducer here
   - structural/style/pivot questions have dedicated probes
     (`grid_edit_probe.py`, `band_insert_probe.py`, `vba_style_probe.py`,
     `vba_range_tracking_probe.py`, `vba_table_probe.py`, `vba_pivot_probe.py`,
     `pivot_filter_probe.py`) — **re-run the probe rather than reasoning from
     memory about what Excel does.**
3. **Arbitrate with a third reference**, not with the two disagreeing engines: a
   high-precision `decimal`/`mpmath` evaluation, the documented spec, or the day
   count/financial definition. Whoever matches it is right.
4. **Then act:**
   - *Excel is right* → fix `visi-core`, add the regression test.
   - *visi is right* → do **not** change the engine. Add a numbered section to
     `docs/excel-discrepancies.md` (state which kind: "Excel is wrong" / "visi
     gap" / "no stable answer"), exclude the case in the generator with a
     comment pointing at that section, and add a test pinning visi against the
     independent reference — the same shape as
     `test_besselj_stays_accurate_where_excel_does_not`.
   - *No stable answer* (Excel is internally inconsistent or heuristic) →
     document and exclude; do not encode Excel's coin flip.

### Windows Excel wins

Where **Windows Excel and macOS Excel disagree, Windows is authoritative.**
visi's behaviour is pinned to Windows; the macOS result is then a platform note,
never the thing a test asserts.

- Reach for a Windows result whenever the case smells platform-specific: VBA
  `Err.Number` values, object-model error numbers, locale/date rendering,
  chart and drawing XML details, anything Mac-only in the AppleScript bridge.
  Section 17 of `docs/excel-discrepancies.md` is an existing instance —
  Excel for Mac's error number there is not even reproducible run to run.
- On Windows, run the harnesses with `--driver win32com`. On macOS, run them
  with `--excel-path "/Applications/Microsoft Excel.app"` (AppleScript). If no
  Windows machine is available in this session for a platform-sensitive case,
  **say so** and do not silently pin the engine to the macOS answer: either
  leave the case documented as awaiting Windows confirmation, or ask the user
  whether they can run the reduced case on Windows.
- When a discrepancy entry records a platform split, name both results and mark
  which one visi implements.

## The test is the deliverable

Every fixed or documented failure leaves behind a Rust test that **runs in CI
with no Excel installed**. Put it where the case actually lives:

| Case | Where |
| --- | --- |
| formula evaluation | `visi-core/src/core/engine/tests/{aggregate,logical,math,math_trig,rounding,stats,text,text_fn,extended,new_functions}.rs` |
| VBA host object model | inline `mod tests` in `visi-core/src/core/vba/host.rs` |
| VBA `Variant` semantics | inline tests in `visi-core/src/core/vba/value.rs` |
| tables / pivots / xlsx round-trip | inline `mod tests` in `table.rs`, `pivot.rs`, `xlsx.rs` (the pivot-XML round trip is tested from `xlsx.rs`, not `pivot_xlsx.rs`) |
| structural edits | `grid_edit.rs` tests |
| anything that needs a real file round trip | `visi/tests/cli_tests.rs` |

Engine tests follow the harvested-fuzz-case convention: a literal grid handed to
the local `create_sheet` helper, `sheet.commit(None).unwrap()`, then one
assertion on one cell, with the expected value in the panic message. Name it
after the behaviour (`test_fuzz_<what>_<condition>`) and keep the minimized grid,
not the original 10×5 one.

For a VBA host test, read the expected string **off a probe run** and paste it —
that is why those tests assert exact strings.

Then:

```bash
cargo test -p visi-core
cargo fmt && cargo clippy --workspace --all-targets -- -D warnings
```

## Things that will bite

- **A run that produces no output at all is a modal Excel dialog, not a slow
  run.** A VBA *compile* error (undefined name, duplicate `Dim`) is not
  catchable by the `On Error` harness, so Excel goes modal and `osascript` never
  returns. `killall "Microsoft Excel"` and read the generated source.
- Editing a shared VBA source constant (`HARNESS_TEMPLATE` in `fuzz_vba.py`, which
  both probe scripts splice in) breaks the importers as a **hang**, not a test
  failure. Keep it self-contained.
- Triaging a crash: use `--backend subprocess` (on `fuzz_excel.py`,
  `fuzz_chart.py`, `fuzz_pivot.py` — the two VBA fuzzers have no such flag and
  always run in process). Under the default bindings backend the engine shares
  the harness process, so a Rust panic or stack overflow takes the whole run
  down instead of one iteration.
- `fuzz_excel.py` tolerates cells where both engines errored with *different*
  error classes (documented divergence); `--strict-error-class` makes those
  failures. Don't chase them unless the user asked for that mode.
- Long runs: `fuzz_vba.py` takes `--restart-every` because Excel degrades over
  batches. If mismatches suddenly appear in bulk late in a run, restart Excel
  and re-run those seeds before believing them.

## Report

Say which harness ran, the iteration count, how many rounds it took, and for
each failure: the seed, the reduced case, which engine was right and on what
evidence, the fix or the discrepancy section, and the test that now covers it.
State plainly if the final run was clean — and if it wasn't, say what is still
failing rather than reporting completion.
