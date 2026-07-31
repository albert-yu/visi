# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
cargo build --workspace                 # dev binary at target/debug/visi
cargo build --release --workspace       # target/release/visi
cargo test --workspace                  # all unit + integration tests
cargo clippy --workspace                # lints (code carries #[allow(clippy::...)] in places)
cargo fmt

# A single test / module (engine tests live inside libvisi's lib target)
cargo test -p libvisi test_fuzz_cell_reference_zero_coercion
cargo test -p libvisi --lib core::engine::tests_fuzz::rounding
cargo test -p visi --test cli_tests
```

Differential fuzzing against real Microsoft Excel (Python, `fuzz/`):

```bash
pip install -r fuzz/requirements.txt
cargo build --release                                     # harness defaults to ./target/release/visi
python3 fuzz/fuzz_excel.py --driver mock --iterations 5    # no Excel needed; exercises the pipeline
python3 fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20
python3 fuzz/fuzz_excel.py --seed 48291 --iterations 1     # reproduce a specific failure
```

Failures land in `fuzz_results/failures/fail_iter_<N>_seed_<SEED>/` as `source.xlsx` / `visi_out.xlsx` / `excel_out.xlsx`. See `fuzz/README.md` for the Excel-parity edge cases the harness is built around (cached `<v>` values, 1900 leap-year bug, `_xlfn.` prefixes, float tolerance).

## Architecture

Cargo workspace, edition 2024:

- **`libvisi`** — the engine. Built as `rlib` **and `cdylib`**, so it is meant to stay embeddable (no CLI/IO assumptions in `core`). Uses `web-time` instead of `std::time` and `getrandom` for IDs so it can target wasm.
- **`visi`** — clap-based CLI. `cli.rs` is the arg surface, `main.rs` holds one `handle_*` fn per subcommand, `engine.rs` wraps everything in `WorkbookManager`.

### Data model (`libvisi/src/core/engine/`)

A `Sheet` is **column-oriented**: `columns: Vec<DataColumn>`, each with parallel per-row vectors:

- `src: SharedVec<String>` — the raw user text (`"10"`, `"=SUM(A1:A2)"`, `"\"literal text\""`)
- `data: ColumnData` — computed values, stored as a typed column (`Integer`/`Float` + validity `Bitmask`, or `Any(Vec<ResultData>)`). Writing a mismatched type auto-promotes `Integer → Float` or demotes to `Any`.
- `compiled_src: SharedVec<CompiledFormula>` — cached compile output
- `dirty_indices` — recompute queue

Everything internal is **0-based `(row, col)`**; A1 notation exists only at the parser and CLI boundaries (`parser::col_idx_to_letters`, `visi/src/utils.rs`). `src`, `data`, and `compiled_src` must stay the same length — row/col insert/delete paths in `sheet.rs` maintain that invariant by hand.

`ResultData` is the value type (`None`/`Boolean`/`Integer`/`Float`/`String`/`List`/`Dict`/`Plot`/`Error`). `result_data::format_excel_number` reproduces Excel's 15-significant-digit display rules — change it only with fuzz evidence.

### Formula pipeline

Formula text goes through **two distinct representations**, which is the single most important thing to know before touching `parser.rs`:

1. `compile_formula(src, &sheets)` → `CompiledFormula`: splits text into `FormulaPart`s where every reference is stored by **`sheet_id` / `col_id` (u64), not by name**. This is what makes sheet/table/column renames non-destructive.
2. `serialize_formula(&compiled, &sheets)` → A1 text again, rendered with the *current* names.
3. `parse_excel_formula(text)` → `Expr` AST (via `lex_eval`).
4. `Sheet::evaluate_ast` / `evaluate_function` walk the AST and return `(ResultData, Vec<Dependency>)`.

`Sheet::commit()` runs all four per dirty cell — compile, re-serialize, then evaluate the re-serialized string. Non-formula cells (no leading `=`) are parsed as literals right there in `commit`, which is why importing text that *looks* numeric requires quoting (see `xlsx::text_cell_src`).

`evaluate_function` dispatches on the uppercased name after stripping a leading `_xlfn.`. Alongside Excel functions it implements engine-specific ones (`PLOT`, `GET`, `GET_COL`, `GET_COL_IDX`, `SLICE`, `STR`) — don't assume every name maps to Excel.

### Recalculation and dependencies

`Sheet::commit(context)` is a BFS over a dirty queue, maintaining both directions of the dependency graph (`dependencies: Dependency → dependents`, `dependencies_rev: cell → its providers`). `Dependency` distinguishes `Local`/`LocalColumn` from `Remote`/`RemoteColumn` (cross-sheet, keyed by sheet *name*).

**`commit` only propagates local dependencies.** Cross-sheet propagation is handled a level up by `WorkbookManager::evaluate()`, which marks every sheet dirty and runs **3 fixed passes** over all sheets, rebuilding a `Context` (name → `&Sheet`) for each target sheet via `split_at_mut`. Deep cross-sheet chains can therefore need more passes than exist. Circular references are bounded by `max_ops` inside `commit`, not detected properly.

Cross-sheet evaluation always needs a `Context`; without one, remote refs error out.

### Excel Tables vs sheets (naming trap)

A `Sheet` is informally called a "table" throughout this codebase (`Sheet::new` defaults to `"table_1"`, `Context::add_table`). An **`ExcelTable`** (`core/table.rs`) is a different thing: a ListObject — a named rectangular sub-range *on* a sheet with a header row, optional totals row, and named columns.

Structured references (`Sales[Amount]`, `[@Amount]`, `Table[#Headers]`) resolve in `evaluate_ast`'s `Expr::StructuredRef` arm: first look for a real `ExcelTable` by name (this sheet, then any sheet in the `Context`), and only if none exists fall back to the legacy behavior of treating the leading name as a *sheet* name with the whole sheet as an implicit table. Both paths must keep working.

Table names are unique **workbook-wide** (enforced in `WorkbookManager`), and lookups are case-insensitive. Renaming a table or a table column cascades into formula *text* across the whole workbook via `parser::rewrite_structured_table_reference` (called from `WorkbookManager::rewrite_table_references`, then re-evaluated) — mirroring Excel. `parser::render_structured_ref_text` is shared by `serialize_formula` and the rename rewriter so the canonical bracket syntax stays in sync between them.

### xlsx I/O (`libvisi/src/core/xlsx.rs`)

Import uses **calamine**, export uses **rust_xlsxwriter** — two different libraries with different models, so round-tripping is asymmetric and worth checking after changes.

- Export writes each formula with its **cached result** (`Formula::set_result`), so Excel/openpyxl/the fuzzer can read values without recalculating. Dropping this breaks the differential harness.
- Table header/totals flags aren't exposed by calamine, so `read_table_row_flags` parses `xl/tables/table*.xml` out of the zip directly.
- Charts are parsed by walking the drawing/chart rels and XML by hand with `quick-xml`.
- Worksheet names get truncated to 31 chars and de-duplicated on export; `orig_to_assigned_name` maps original → assigned so table definitions reattach to the right sheet.

### CLI conventions (`visi/`)

Follows clig.dev. `-` means stdin/stdout for the file argument. Writes require either `--output <path>` or `--in-place`/`-i` (`resolve_output_path` in `main.rs`). User-facing row/column indices are **1-based** (or letters for columns) and converted in `utils.rs`; `set` accepts repeated `-S A1=100` pairs. `--quiet` suppresses informational stderr; most commands also take `--json`/`--format`.

## Tests

- `libvisi/src/core/engine/tests.rs` — hand-written engine tests.
- `libvisi/src/core/engine/tests_fuzz/{aggregate,logical,math,rounding,text}.rs` — **regression cases harvested from the differential fuzzer**, each a literal grid fed to the local `create_sheet` helper plus an assertion on one cell. When the Python harness finds an Excel mismatch, minimize it and add it here.
- `libvisi/src/core/table.rs` and `xlsx.rs` have inline `#[cfg(test)] mod tests` for table CRUD and xlsx round-tripping.
- `visi/tests/cli_tests.rs` — integration tests that drive `WorkbookManager` (the same API the CLI handlers call) through real file round-trips.
