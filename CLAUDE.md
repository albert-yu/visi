# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
cargo build --workspace                 # dev binary at target/debug/visi
cargo build --release --workspace       # target/release/visi
cargo test --workspace                  # all unit + integration tests
cargo clippy --workspace                # lints (code carries #[allow(clippy::...)] in places)
cargo fmt

# A single test / module (engine tests live inside visi-core's lib target)
cargo test -p visi-core test_fuzz_cell_reference_zero_coercion
cargo test -p visi-core --lib core::engine::tests::rounding
cargo test -p visi --test cli_tests
```

Differential fuzzing against real Microsoft Excel (Python, `fuzz/`). Use the venv,
not system Python — `maturin develop` installs into whichever venv is active:

```bash
source fuzz/venv/bin/activate
pip install -r fuzz/requirements.txt
maturin develop -m visi-python/Cargo.toml --release      # the in-process bindings
cargo build --release                                   # the CLI, for --backend subprocess

python fuzz/fuzz_excel.py --driver mock --iterations 5   # no Excel needed; exercises the pipeline
python fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20
python fuzz/fuzz_excel.py --seed 48291 --iterations 1    # reproduce a specific failure

pytest fuzz/test_backend_parity.py visi-python/tests/    # bindings must match the CLI
```

Each fuzzer takes `--backend {auto,bindings,subprocess}`. `auto` prefers the
bindings and falls back to the CLI with a warning. Reach for `subprocess` when
triaging a crash: under `bindings` the engine shares the harness process, so a
Rust abort or stack overflow takes the whole run down instead of one iteration.

`cargo build --workspace` / `cargo test --workspace` include `visi-python`, which
needs a `python3` on PATH to link. Without one, pass `--exclude visi-python` or set
`PYO3_NO_PYTHON=1`; a bare `cargo build` / `cargo test` already skips it via
`default-members`.

Failures land in `fuzz_results/failures/fail_iter_<N>_seed_<SEED>/` as `source.xlsx` / `visi_out.xlsx` / `excel_out.xlsx`. See `fuzz/README.md` for the Excel-parity edge cases the harness is built around (cached `<v>` values, 1900 leap-year bug, `_xlfn.` prefixes, float tolerance).

Crash/panic fuzzing of the VBA import path (Rust, `visi-core/fuzz/`, separate from the Python differential harness above):

```bash
cargo install cargo-fuzz                                          # needs a nightly toolchain
cd visi-core && cargo +nightly fuzz run ovba_decompress
mkdir -p fuzz/corpus/vba_import
cargo +nightly fuzz run vba_import fuzz/corpus/vba_import fuzz/seeds/vba_import   # seed corpus gets past the CFB-magic-bytes gate
```

See `visi-core/fuzz/README.md`. `core::ovba`'s roundtrip/never-panics properties are also covered by `proptest` cases in `cargo test -p visi-core`, no nightly needed.

## Architecture

Cargo workspace, edition 2024:

- **`visi-core`** — the engine, published to crates.io as `visi-core` (the directory matches). Plain `rlib` — it kept a `cdylib` crate-type for a while without a single `extern "C"` symbol behind it; embedding for another language belongs in a separate crate (as `visi-python` now does), since `crate-type` can't be feature-gated. Still meant to stay embeddable in the sense that matters (no CLI/IO assumptions in `core`). Uses `web-time` instead of `std::time` and `getrandom` for IDs so it can target wasm — the browser JS backend is behind the **`wasm` feature** (`getrandom/js`), off by default because a library must not force a global getrandom backend on its consumers.
- **`visi`** — clap-based CLI. `cli.rs` is the arg surface, `main.rs` holds one `handle_*` fn per subcommand, `engine.rs` wraps everything in `WorkbookManager`.
  - be sure to follow [Command Line Interface Guidelines](https://clig.dev) when making changes to the CLI
  - the CLI keeps its own `Result<_, String>` style internally and converts at the boundary (`exit_with_error` takes `impl Display`)
- **`visi-python`** — pyo3 bindings over `visi-core`, exposed to Python as the module `visi_core`. `crate-type = ["cdylib"]`, `abi3-py39`, `publish = false`; built with maturin for `fuzz/`, which drives the engine in-process instead of spawning the CLI per operation. Three things about it are load-bearing:
  - It depends on **`visi-core` only**, never on `visi`. Where it has to mirror CLI behavior — `edit_chart`'s `--title`/`--clear-title` pair, and `add_pivot_field`'s post-add subtotal/label mutation (`visi/src/main.rs`'s AddField arm) — that mirroring is duplicated logic, and `fuzz/test_backend_parity.py` is the only thing that will notice it drifting.
  - `extension-module` is **not** a default cargo feature. Turning it on breaks `cargo test --workspace`'s link step with an undefined `_PyModule_Create2`, whose message points nowhere near the cause. maturin enables it via `pyproject.toml`.
  - The Python module is named `visi_core`, not `visi`: the repo root holds a `visi/` directory with no `__init__.py`, which PEP 420 makes an implicit namespace package, so `import visi` from the root resolves to the CLI crate's source directory.

  Adding a binding is not a reason to widen `visi-core`'s public API — reach through the existing public fields, as `get_cell` does.

The two crates version independently: `visi` is at the workspace version, `visi-core` pins its own (`0.1.0`) since it is newer to crates.io.

### Public API surface

Not everything in `core` is public. The modules implementing Excel's function library — `stats`, `math_trig`, `text`, `date_fn`, `date`, `engineering`, `finance`, `extended_fn`, `ets`, `xml`, `vba_synth`, `pivot_xlsx`, `parser`, `formula`, `actions`, `shared_vec` — are `pub(crate)`; their types reach users only through the curated `pub use` list at the bottom of `core/mod.rs`. `ovba` and `vba_xlsx` are `#[doc(hidden)] pub` because `visi-core/fuzz` and the `dump_vba_fuzz_seeds` example need them, not because they are supported.

**When adding a public item, ask whether it belongs in that re-export list.** Anything reachable from `core`'s `pub use` is a semver commitment.

`lib.rs` carries `#![warn(missing_docs)]`, so a new public item without a doc comment warns. Note the lint's blind spot: it fires on the item's *definition site*, so it says nothing about a `pub` item inside a `pub(crate)` module even when a `pub use` re-exports it into the public API. Sealing a module therefore silences the lint without actually shrinking the surface — check with `cargo doc` (and `-W unnameable_types`, which catches a type that stays reachable through a public field or variant but can no longer be named).

Fallible public API returns `crate::Error` (`src/error.rs`), not `String`: an `#[non_exhaustive]` enum with `NotFound`/`AlreadyExists`/`NameTaken`/`InvalidName` carrying an `ObjectKind` so callers can branch without parsing text. Lower layers (Excel Table and pivot internals) still produce `String` and are wrapped in `Error::InvalidArgument` at the `workbook.rs` boundary — carve real variants out of it as those layers get typed. Formula-evaluation internals deliberately keep `Result<_, String>`, where the string is an Excel error code like `#VALUE!`, not a Rust error.

### Data model (`visi-core/src/core/engine/`)

A `Sheet` is **column-oriented**: `columns: Vec<DataColumn>`, each with parallel per-row vectors:

- `src: SharedVec<String>` — the raw user text (`"10"`, `"=SUM(A1:A2)"`, `"\"literal text\""`)
- `data: ColumnData` — computed values, stored as a typed column (`Integer`/`Float` + validity `Bitmask`, or `Any(Vec<ResultData>)`). Writing a mismatched type auto-promotes `Integer → Float` or demotes to `Any`.
- `compiled_src: SharedVec<CompiledFormula>` — cached compile output
- `dirty_indices` — recompute queue

Everything internal is **0-based `(row, col)`**; A1 notation exists only at the parser and CLI boundaries (`parser::col_idx_to_letters`, `visi/src/utils.rs`). `src`, `data`, `compiled_src`, and `styles` must stay the same length, and every column of a sheet must have the same number of rows (`Sheet::row_count` reads only the first and assumes the rest match).

Those four vectors and `Sheet::columns` are `pub(crate)`; outside the crate they are reachable read-only through `DataColumn`'s accessors (`len`, `src`, `value`, `compiled`, `style`) and `Sheet::columns()`. **Change a column's length only through `DataColumn`'s paired operations** — `push_row`, `insert_row`, `remove_row`, `drain_rows`, `resize_rows`, `rebuild_after_load` — which touch all four vectors together. Hand-maintaining them is what let `extend` and `delete` silently skip `styles`, which made a row added by `extend` unstylable and shifted every style below a deleted range onto the wrong row. `dirty_indices` is deliberately *not* part of the invariant: it is a recompute queue `commit` drains, and the paired operations rebase it for you.

`ResultData` is the value type (`None`/`Boolean`/`Integer`/`Float`/`String`/`List`/`Dict`/`Error`). `result_data::format_excel_number` reproduces Excel's 15-significant-digit display rules — change it only with fuzz evidence.

### Dates are numbers with a format (`core/date.rs`)

There is deliberately **no date value type**. As in Excel, a date cell holds a plain numeric serial and the notation it was typed in lives on the cell, as `CellStyle::num_format` (an Excel format code like `m/d/yy`). So `6/22/26` is `Float(46195)` — `ISNUMBER` is true, `SUM` counts it, every numeric path works untouched — and only rendering consults the format. A `ResultData::Date` variant was considered and rejected: the ~200 sites that match on `Float` all have catch-all arms, so any missed one would silently treat a date as non-numeric.

- `commit` recognizes a literal via `date::parse_date` and records `DateFormat::to_format_code()`.
- `Sheet::get_display_string` is the **only** place that renders a serial back to a date — show values through it, not by formatting `ResultData` directly.
- `Sheet::inherited_date_format` gives `=A1+1` its operand's format. The rule keys off the *operator*, not the dependency count, because those come apart: `=YEAR(A1)` reads one date cell and returns a year. Only a bare cell ref and `+`/`-` with exactly one date side inherit; `=A1-A2` (a day count) and `=SUM(...)` deliberately do not.
- `date::render_date_code` is shared with `TEXT()` so there is one date formatter. It scans token *runs* in one pass — successive string replacement corrupts month names, since `December` contains an `m` and `May` a `y`.
- A format code cannot carry month-name casing, so `22-JUN-2026` round-trips through `format_date` but comes back title-cased through a worksheet — matching Excel. Zero-padding of a numeric month/day is likewise not recorded (`06/22/2026` → `6/22/2026`).
- Text that merely looks like a date must be quoted to stay text (`xlsx::text_cell_src` does this for imported string cells).

### Formula pipeline

Formula text goes through **two distinct representations**, which is the single most important thing to know before touching `parser.rs`:

1. `compile_formula(src, &sheets)` → `CompiledFormula`: splits text into `FormulaPart`s where every reference is stored by **`sheet_id` / `col_id` (u64), not by name**. This is what makes sheet/table/column renames non-destructive.
2. `serialize_formula(&compiled, &sheets)` → A1 text again, rendered with the *current* names.
3. `parse_excel_formula(text)` → `Expr` AST (via `lex_eval`).
4. `Sheet::evaluate_ast` / `evaluate_function` walk the AST and return `(ResultData, Vec<Dependency>)`.

`Sheet::commit()` runs all four per dirty cell — compile, re-serialize, then evaluate the re-serialized string. Non-formula cells (no leading `=`) are parsed as literals right there in `commit`, which is why importing text that *looks* numeric requires quoting (see `xlsx::text_cell_src`).

`evaluate_function` dispatches on the uppercased name after stripping a leading `_xlfn.`. Alongside Excel functions it implements engine-specific ones (`GET`, `GET_COL`, `GET_COL_IDX`, `SLICE`, `STR`) — don't assume every name maps to Excel.

### Recalculation and dependencies

`Sheet::commit(context)` is a BFS over a dirty queue, maintaining both directions of the dependency graph (`dependencies: Dependency → dependents`, `dependencies_rev: cell → its providers`). `Dependency` distinguishes `Local`/`LocalColumn` from `Remote`/`RemoteColumn` (cross-sheet, keyed by sheet *name*).

**`commit` only propagates local dependencies.** Cross-sheet propagation is handled a level up by `WorkbookManager::evaluate()`, which marks every sheet dirty and runs **3 fixed passes** over all sheets, rebuilding a `Context` (name → `&Sheet`) for each target sheet via `split_at_mut`. Deep cross-sheet chains can therefore need more passes than exist. Circular references are bounded by `max_ops` inside `commit`, not detected properly.

Cross-sheet evaluation always needs a `Context`; without one, remote refs error out.

### Excel Tables vs sheets (naming trap)

A `Sheet` is informally called a "table" throughout this codebase (`Sheet::new` defaults to `"table_1"`, `Context::add_table`). An **`ExcelTable`** (`core/table.rs`) is a different thing: a ListObject — a named rectangular sub-range *on* a sheet with a header row, optional totals row, and named columns.

Structured references (`Sales[Amount]`, `[@Amount]`, `Table[#Headers]`) resolve in `evaluate_ast`'s `Expr::StructuredRef` arm: first look for a real `ExcelTable` by name (this sheet, then any sheet in the `Context`), and only if none exists fall back to the legacy behavior of treating the leading name as a *sheet* name with the whole sheet as an implicit table. Both paths must keep working.

Table names are unique **workbook-wide** (enforced in `WorkbookManager`), and lookups are case-insensitive. Renaming a table or a table column cascades into formula *text* across the whole workbook via `parser::rewrite_structured_table_reference` (called from `WorkbookManager::rewrite_table_references`, then re-evaluated) — mirroring Excel. `parser::render_structured_ref_text` is shared by `serialize_formula` and the rename rewriter so the canonical bracket syntax stays in sync between them.

### Pivot tables (`visi-core/src/core/pivot.rs`, `pivot_xlsx.rs`)

A `PivotTable` is workbook-level (like `Chart`), not sheet-scoped like `ExcelTable`, since its source and destination ranges can live on different sheets; `WorkbookManager.pivot_tables: Vec<PivotTable>` holds them. `PivotSource` is either an `ExcelTable` name (re-resolved by name on every refresh, so table renames/resizes are picked up automatically) or a raw sheet range.

`pivot::compute_pivot(sheets, &pivot)` is a pure function: reads source rows, applies `filter_fields`, groups nested `row_fields`/`col_fields` (with per-field subtotal toggles and grand totals), aggregates `value_fields` (`Sum`/`Count`/`CountNumbers`/`Average`/`Max`/`Min`), and returns a `PivotGrid` — display-ready header/body rows plus the underlying `row_axis`/`col_axis` (`PivotAxisItem`s) that the xlsx writer needs to reconstruct native `rowItems`/`colItems`. `WorkbookManager::refresh_pivot_table` (`visi/src/engine.rs`) is the only thing that writes the grid into cells, as literal values via `set_cell`/`ensure_capacity` — **like Excel, nothing recomputes a pivot table automatically**; every CRUD op (`add_pivot_field`, `remove_pivot_field`, `set_pivot_filter`) explicitly calls refresh afterward.

Neither xlsx library used here has pivot table support: calamine doesn't expose `xl/pivotCache/*` or `xl/pivotTables/*`, and rust_xlsxwriter has no writer for them at all. `pivot_xlsx.rs` hand-rolls both directions:

- **Export** (`inject_pivot_tables`) post-processes the zip `export_xlsx_data` already produced — rust_xlsxwriter has no hook for extra parts — by re-opening it with the `zip` crate, editing `[Content_Types].xml` / `xl/workbook.xml` (`<pivotCaches>`) / `xl/_rels/workbook.xml.rels` / the destination worksheet's `.rels`, and writing new `pivotCacheDefinition`/`pivotCacheRecords`/`pivotTable` parts, then rewriting the whole zip. `rowItems`/`colItems` encode Excel's leading-field repeat suppression (`<i r="N">` = "the first N fields are unchanged from the previous row") plus `t="default"`/`t="grand"` subtotal/grand-total markers — get this wrong and Excel still opens the file (`refreshOnLoad="1"` lets it silently rebuild the cache) but may misrender the grid. Verified against `openpyxl` (a strict independent OOXML reader) rather than real Excel, since driving Excel via AppleScript needs a one-time interactive automation-permission grant this environment couldn't complete — re-verify with the `fuzz/` Excel driver or manually in Excel before trusting further pivot XML changes.
- **Import** (`import_pivot_tables`) reconstructs each `PivotTable`'s source/destination/row/col/value fields from that XML, including per-field subtotal toggles (recovered from whether the field's `<item t="default"/>` placeholder is present) — this is not just fidelity, it's load-bearing: the CLI is a fresh process per invocation, so a pivot table definition only survives `pivot add-field` in a later command because it round-trips through this xlsx parsing. Filter-field *selections* are the one thing not reconstructed (they reset to "all"), since that would require trusting index-based item references against data that may have changed. **Anything that mutates a filter selection must therefore be the last step before saving** — a round trip past that point silently drops it.

### xlsx I/O (`visi-core/src/core/xlsx.rs`)

Import uses **calamine**, export uses **rust_xlsxwriter** — two different libraries with different models, so round-tripping is asymmetric and worth checking after changes.

- Export writes each formula with its **cached result** (`Formula::set_result`), so Excel/openpyxl/the fuzzer can read values without recalculating. Dropping this breaks the differential harness.
- Table header/totals flags aren't exposed by calamine, so `read_table_row_flags` parses `xl/tables/table*.xml` out of the zip directly.
- Charts are parsed by walking the drawing/chart rels and XML by hand with `quick-xml`.
- Worksheet names get truncated to 31 chars and de-duplicated on export; `orig_to_assigned_name` maps original → assigned so table definitions reattach to the right sheet.

### CLI conventions (`visi/`)

Follows clig.dev. `-` means stdin/stdout for the file argument. Writes require either `--output <path>` or `--in-place`/`-i` (`resolve_output_path` in `main.rs`). User-facing row/column indices are **1-based** (or letters for columns) and converted in `utils.rs`; `set` accepts repeated `-S A1=100` pairs. `--quiet` suppresses informational stderr; most commands also take `--json`/`--format`.

## Tests

- `visi-core/src/core/engine/tests/unit.rs` — hand-written engine tests.
- `visi-core/src/core/engine/tests/{aggregate,logical,math,rounding,text}.rs` — **regression cases harvested from the differential fuzzer**, each a literal grid fed to the local `create_sheet` helper plus an assertion on one cell. When the Python harness finds an Excel mismatch, minimize it and add it here.
- `visi-core/src/core/table.rs`, `pivot.rs`, and `xlsx.rs` have inline `#[cfg(test)] mod tests` for table CRUD, pivot computation/grouping, and xlsx round-tripping (including a pivot table round-trip through the hand-rolled OOXML).
- `visi/tests/cli_tests.rs` — integration tests that drive `WorkbookManager` (the same API the CLI handlers call) through real file round-trips.
