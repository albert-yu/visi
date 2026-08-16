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

CI (`.github/workflows/ci.yml`) runs `cargo fmt --check`, `cargo clippy --workspace
--all-targets -- -D warnings`, `cargo test --workspace --exclude visi-python` on
Linux and macOS, and the pytest suites (`fuzz/test_backend_parity.py`,
`visi-python/tests/`) on every PR. Everything requiring a real Excel — the `fuzz/`
differential harness — and the nightly-only cargo-fuzz targets stay local.

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

python fuzz/grid_edit_probe.py                           # what a row/col insert/delete does to formula text
python fuzz/fuzz_vba.py --iterations 200 --seed 909      # VBA execution + the cells a macro wrote
python fuzz/vba_host_probe.py                            # what Excel's object model actually does
python fuzz/vba_range_tracking_probe.py                  # does a held Range follow a row/col edit (one case per round trip)
python fuzz/vba_expr_probe.py -e 'a = 1 :: a + 1'        # one expression, both engines, side by side

pytest fuzz/test_backend_parity.py visi-python/tests/    # bindings must match the CLI
```

Anything in `fuzz/` that becomes VBA *source* has a trap worth knowing: an
undefined name or a duplicate `Dim` is a **compile** error, which the `On Error`
harness cannot catch, so Excel goes modal and `osascript` never returns. A run
that produces no output at all is a compile error, not a slow run -- `killall
"Microsoft Excel"` and look at the generated source. `HARNESS_TEMPLATE` in
`fuzz_vba.py` is imported and spliced into modules by both probe scripts, so it
has to stay self-contained; `fuzz_vba.GRID_HARNESS_TEMPLATE` is the one that may
depend on that file's own helpers.

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
mkdir -p fuzz/corpus/vba_parse
cargo +nightly fuzz run vba_parse fuzz/corpus/vba_parse fuzz/seeds/vba_parse       # VBA source text, not a binary
```

See `visi-core/fuzz/README.md`. `core::ovba`'s roundtrip/never-panics properties are also covered by `proptest` cases in `cargo test -p visi-core`, no nightly needed.

## Architecture

Cargo workspace, edition 2024:

- **`visi-core`** — the engine, published to crates.io as `visi-core` (the directory matches). Plain `rlib` — it kept a `cdylib` crate-type for a while without a single `extern "C"` symbol behind it; embedding for another language belongs in a separate crate (as `visi-python` now does), since `crate-type` can't be feature-gated. Still meant to stay embeddable in the sense that matters (no CLI/IO assumptions in `core`). Uses `web-time` instead of `std::time` and `getrandom` for IDs so it can target wasm — the browser JS backend is behind the **`wasm` feature** (`getrandom/js`), off by default because a library must not force a global getrandom backend on its consumers.
- **`visi`** — clap-based CLI. `cli.rs` is the arg surface, `main.rs` holds one `handle_*` fn per subcommand, `engine.rs` wraps everything in `WorkbookManager`.
  - be sure to follow [Command Line Interface Guidelines](https://clig.dev) when making changes to the CLI
  - the CLI keeps its own `Result<_, String>` style internally and converts at the boundary (`exit_with_error` takes `impl Display`)
- **`visi-python`** — pyo3 bindings over `visi-core`, exposed to Python as the module `visi_core`. `crate-type = ["cdylib"]`, `abi3-py39`, `publish = false`; built with maturin for `fuzz/`, which drives the engine in-process instead of spawning the CLI per operation. Three things about it are load-bearing:
  - It depends on **`visi-core` only**, never on `visi`. Where it has to mirror CLI behavior — `edit_chart`'s `--title`/`--clear-title` pair, `add_pivot_field`'s post-add subtotal/label mutation (`visi/src/main.rs`'s AddField arm), and `add_macro`'s sheet-name-to-id resolution with its `ThisWorkbook` exemption (the Add arm of `handle_macro`) — that mirroring is duplicated logic, and `fuzz/test_backend_parity.py` is the only thing that will notice it drifting.
  - `extension-module` is **not** a default cargo feature. Turning it on breaks `cargo test --workspace`'s link step with an undefined `_PyModule_Create2`, whose message points nowhere near the cause. maturin enables it via `pyproject.toml`.
  - The Python module is named `visi_core`, not `visi`: the repo root holds a `visi/` directory with no `__init__.py`, which PEP 420 makes an implicit namespace package, so `import visi` from the root resolves to the CLI crate's source directory.

  Adding a binding is not a reason to widen `visi-core`'s public API — reach through the existing public fields, as `get_cell` does.

The two crates version independently: `visi` is at the workspace version, `visi-core` pins its own (`0.1.0`) since it is newer to crates.io.

### Public API surface

Not everything in `core` is public. The modules implementing Excel's function library — `stats`, `math_trig`, `text`, `date_fn`, `date`, `engineering`, `finance`, `extended_fn`, `ets`, `xml`, `vba_synth`, `pivot_xlsx`, `parser`, `formula`, `actions`, `shared_vec` — are `pub(crate)`; their types reach users only through the curated `pub use` list at the bottom of `core/mod.rs`. `ovba`, `vba_xlsx`, and `vba`'s `ast`/`lexer`/`parser` submodules are `#[doc(hidden)] pub` because `visi-core/fuzz` and the `dump_vba_fuzz_seeds` example need them, not because they are supported — the VBA syntax layer's supported surface is `check_syntax`/`ModuleSyntax`, and the AST's shape is deliberately not a semver commitment until the interpreter phases need it.

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

### Structural edits (`visi-core/src/core/grid_edit.rs`)

Inserting or deleting a row/column does not just move cells — every formula in the **whole workbook** has to be rewritten so its references follow, or `=A3` keeps pointing at row 3 after the value it meant slid to row 4. `WorkbookManager::{insert,delete}_{row,col}` go through `apply_grid_edit`, which is deliberately **three phases**:

1. compile every formula *before* the edit (compiling needs the grid the text was written against),
2. apply the edit, and move `ExcelTable` extents and `PivotSource::Range`/destination coordinates with it,
3. serialize the shifted formulas back to text *after* the edit, at wherever each formula's own cell moved to.

Phase 3 cannot be folded into phase 1: a whole-column reference is held by `col_id` and renders as the column's *current* letter, so serializing `=SUM(B:B)` before a column is inserted to its left writes `B:B` into a cell where `B` now names a different column — and `src` is what the next recompile reads, so the wrong text wins.

The shift rules were **measured against real Excel** via `fuzz/grid_edit_probe.py` (15 cases, all agreeing), not taken from documentation. The counterintuitive ones: **`$` does not pin a reference against a structural edit** (`$A$3` shifts exactly as `A3` does); inserting at a range's *first* row moves the range while inserting one row lower grows it; deleting part of a range shrinks it but deleting all of it is `#REF!`; and `#REF!` replaces the *reference*, not the formula (`=A3+1` becomes `=#REF!+1`). Re-run the probe rather than "fixing" one from memory.

`shift_span` takes a real index — the whole-column sentinel (`end_row: usize::MAX`, what `A:C` compiles to) must be screened out first or `end + 1` overflows. `parser::lex_eval` grew an `EvalToken::Error` over the closed `EXCEL_ERROR_CODES` set for this, since a formula could not previously hold a literal `#REF!` at all.

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

### VBA (`visi-core/src/core/vba/`)

Two layers that share a directory but not much else:

- **Storage** (`mod.rs`, plus `ovba.rs`/`vba_xlsx.rs`/`vba_synth.rs` outside it) — `VbaProject`/`VbaModule` and the `vbaProject.bin` round trip. Workbook-level like `Chart`, not sheet-scoped. Read those files' own doc comments before touching them; the p-code prefix and MODULECOOKIE handling are both load-bearing in ways that are not guessable.
- **Syntax** (`lexer.rs`, `ast.rs`, `parser.rs`) — Phase 0 of `docs/vba-macro-support.md`. Parses only: no name resolution, no types, no evaluation, which is why a `Call` node cannot distinguish a procedure call from an array index.
- **Execution** (`value.rs`, `interp.rs`, `builtins.rs`) — Phase 1. A tree-walking interpreter over the AST. Driven by `visi macro run`, which is opt-in per invocation and never implicit in `eval`.
- **Host object model** (`host.rs`) — Phase 2. Binds the interpreter to a `WorkbookManager` so a macro can read and write cells, walk sheets, and call `Application.WorksheetFunction`. Anything **outside** its allow-list still raises 438 naming what it was, rather than being skipped — the refusal is the feature, and widening the list is a decision. `core::run_macro` (source text, no workbook) and `WorkbookManager::run_macro` (workbook-bound) are the two entry points; only the latter can mutate anything, and `visi macro run` demands `--output`/`--in-place` when it does.

Things that will bite:

- **VBA's operator precedence differs from the formula language's and was pinned against real Excel**, not documentation. `^` is **left**-associative (`2 ^ 3 ^ 2` = 64, not 512) and binds tighter than unary minus (`-2 ^ 2` = -4); `Eqv` binds tighter than `Imp`, `Xor` tighter than `Eqv`, `Not` looser than comparison. The table and its confirming cases are at the top of `parser.rs`, each with a unit test naming the Excel result. Do not "fix" one of these from memory.
- **Keywords are not reserved.** The lexer emits every one as a plain `Ident` and the parser matches case-insensitively, because VBA's keyword set is contextual — `Name`, `Line`, `Get`, and `Width` are all statements in one position and ordinary property names in another. The `Stmt::Opaque` guards in `parser.rs` are narrow for exactly this reason.
- **Excel compiles VBA lazily, per invoked procedure.** Nothing short of calling a procedure compiles it — not a probe in the same module, not a reference from a dead branch. This is why `fuzz/fuzz_vba_parse.py` wraps generated source in `If False Then ... End If` inside the procedure it calls, and why there is no way to ask Excel whether an arbitrary module compiles without also running something.
- A **compile** error in Excel hangs the AppleScript bridge and, unlike a runtime error, is *not* catchable by an `On Error` wrapper.
- **An object is a handle, not a pointer.** `ObjRef` holds ids and never borrows the workbook, which is what lets the interpreter hold `&mut WorkbookManager` for a whole run. `Is` compares an identity *token*, not the coordinates: `ws.Range("A1") Is ws.Range("A1")` is **False** in Excel (each call builds a fresh object) while `ws Is wb.Worksheets(1)` is True (worksheets are cached). Measured; the obvious tuple comparison is wrong.
- **A `Range` tracks structural edits, so its coordinates live in `Host::ranges`, not in the `ObjRef`.** `Set r = ws.Range("A5")` then `ws.Rows(1).Insert` leaves `r` reading `$A$6` *and* still holding what was in `A5` — and a copy of `r` taken **before** the edit moves too, which is what forced interning over a by-value `Range`. The geometry turned out to be exactly `core::grid_edit`'s, case for case, so `Rows.Insert` and a formula's range reference share `shift_span` rather than two hand-written rules. A range whose every cell is deleted becomes `RangeState::Dead`: **not** `Nothing`, still `TypeName` `"Range"`, but every member access raises. All measured with `fuzz/vba_range_tracking_probe.py`; the one deliberate divergence is the error *number* (Excel for Mac's is not reproducible run to run — see `docs/excel-discrepancies.md` #17).
- **`Insert`/`Delete` are refused on a partial range.** Excel accepts one and picks the shift direction from the range's shape — measured, `Range("A2:A3").Insert` shifts *right*, not down. Guessing silently moves a macro's data sideways, so only a whole-row or whole-column band is accepted.
- **A plain `=` reads an object's default member and `Set` does not**, and the parser cannot tell them apart. Everything wanting a scalar goes through `Interpreter::scalar`; `Set`, `Is`, `TypeName`, a `With` subject and a user procedure's arguments deliberately skip it. Getting this wrong does not raise — it silently produces the wrong kind of value.
- **A write marks the workbook stale; the next read that could observe it recalculates.** Excel recalculates per assignment and it is observable (`A1 = 5` then reading a `D1` holding `=A1*2`), but `WorkbookManager::evaluate` is three passes over every sheet, so doing it literally would be unaffordable in a write loop. Same behaviour, one recalculation per run of consecutive writes.
- **`Application.WorksheetFunction.X` raises where `Application.X` returns.** A failing `WorksheetFunction.VLookup` is a trappable 1004; `Application.VLookup` returns an error `Variant` that `IsError` detects. Two call paths, one implementation — `Sheet::call_worksheet_function`, a `pub(crate)` entry onto `evaluate_function`. The engine's own non-Excel functions (`GET`, `GET_COL`, `GET_COL_IDX`, `SLICE`, `STR`) are deliberately not exposed: a macro using one would work here and fail in Excel, the one direction the differential harness cannot catch.
- **Every `Variant` rule in `value.rs` was measured against real Excel**, not taken from documentation, and each cites the `fuzz/vba_variant_probe.bas` case it came from. Re-measure rather than "correcting" one from memory — a careful hand-probe already got one backwards. **Overflow promotes at runtime but not between literals**: `32767 + 1` written with two literals is error 6, but `a = 32767 : a + 1` is the `Long` 32768. `value::ArithMode` carries which applies. Also non-obvious: `"1" + 1` is a `Double` but `"1" + "2"` is a `String`; `7.6 \ 2` is `4` and typed `Long`; every conversion is banker's rounding; `CStr(-0.0)` is `"-0"`; and `""` is *not* a zero (`"" = 0` is error 13) while `Empty` is.

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
- `visi-core/src/core/vba/host.rs` has inline tests over the VBA host object model, each asserting the exact string `fuzz/vba_host_probe.py` got back from real Excel for the same expression. Read one off a probe run rather than reasoning about it.
- `visi/tests/cli_tests.rs` — integration tests that drive `WorkbookManager` (the same API the CLI handlers call) through real file round-trips.
