# AGENTS.md

This file provides guidance to agents when working with code in this repository.

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
`fuzz/test_comparator.py`, `visi-python/tests/`) on every PR. Everything requiring a real Excel — the `fuzz/`
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
python fuzz/vba_style_probe.py --paint                   # BGR vs RGB, checked against the saved file not the object model
python fuzz/vba_table_probe.py                           # ListObjects; --empty is the zero-data-row case, one round trip per case
python fuzz/band_insert_probe.py                         # what a partial (column-band) insert does to formulas
python fuzz/pivot_filter_probe.py --variant visi         # can Excel open a pivot visi wrote? (exits non-zero if not)
python fuzz/vba_pivot_probe.py                           # the VBA PivotTables object model
python fuzz/vba_expr_probe.py -e 'a = 1 :: a + 1'        # one expression, both engines, side by side

pytest fuzz/test_backend_parity.py fuzz/test_comparator.py visi-python/tests/    # bindings must match the CLI
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
- **`visi`** — usage-rs-based CLI. `cli.rs` is the arg surface, `main.rs` holds one `handle_*` fn per subcommand, `engine.rs` wraps everything in `WorkbookManager`.
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

### Data model (`visi-core/src/core/engine/`)

A `Sheet` is **column-oriented**: `columns: Vec<DataColumn>`, each with parallel per-row vectors:

### Excel Tables vs sheets (naming trap)

A `Sheet` is informally called a "table" throughout this codebase (`Sheet::new` defaults to `"table_1"`, `Context::add_table`). An **`ExcelTable`** (`core/table.rs`) is a different thing: a ListObject — a named rectangular sub-range *on* a sheet with a header row, optional totals row, and named columns.

Structured references (`Sales[Amount]`, `[@Amount]`, `Table[#Headers]`) resolve in `evaluate_ast`'s `Expr::StructuredRef` arm: first look for a real `ExcelTable` by name (this sheet, then any sheet in the `Context`), and only if none exists fall back to the legacy behavior of treating the leading name as a *sheet* name with the whole sheet as an implicit table. Both paths must keep working.

Table names are unique **workbook-wide** (enforced in `WorkbookManager`), and lookups are case-insensitive. Renaming a table or a table column cascades into formula *text* across the whole workbook via `parser::rewrite_structured_table_reference` (called from `WorkbookManager::rewrite_table_references`, then re-evaluated) — mirroring Excel. `parser::render_structured_ref_text` is shared by `serialize_formula` and the rename rewriter so the canonical bracket syntax stays in sync between them.

### CLI conventions (`visi/`)

Follows clig.dev. `-` means stdin/stdout for the file argument. Writes require either `--output <path>` or `--in-place`/`-i` (`resolve_output_path` in `main.rs`). User-facing row/column indices are **1-based** (or letters for columns) and converted in `utils.rs`; `set` accepts repeated `-S A1=100` pairs. `--quiet` suppresses informational stderr; most commands also take `--json`/`--format`.

## Tests

- `visi-core/src/core/engine/tests/unit.rs` — hand-written engine tests.
- `visi-core/src/core/engine/tests/{aggregate,logical,math,rounding,text}.rs` — **regression cases harvested from the differential fuzzer**, each a literal grid fed to the local `create_sheet` helper plus an assertion on one cell. When the Python harness finds an Excel mismatch, minimize it and add it here.
- `visi-core/src/core/table.rs`, `pivot.rs`, and `xlsx.rs` have inline `#[cfg(test)] mod tests` for table CRUD, pivot computation/grouping, and xlsx round-tripping (including a pivot table round-trip through the hand-rolled OOXML).
- `visi-core/src/core/vba/host.rs` has inline tests over the VBA host object model, each asserting the exact string `fuzz/vba_host_probe.py` got back from real Excel for the same expression. Read one off a probe run rather than reasoning about it.
- `visi/tests/cli_tests.rs` — integration tests that drive `WorkbookManager` (the same API the CLI handlers call) through real file round-trips.

## Comments and prose

When writing comments, documentation, or prose, clearly indicate that it was written by an LLM agent.
