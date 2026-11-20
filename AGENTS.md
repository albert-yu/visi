# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Commands

```bash
cargo build --workspace                 # dev binary at target/debug/visi
cargo build --release --workspace       # target/release/visi
cargo test --workspace                  # all unit + integration tests
cargo clippy --workspace                # lints (code carries #[allow(clippy::...)] in places)
cargo fmt

cargo test -p visi-core test_fuzz_cell_reference_zero_coercion
cargo test -p visi-core --lib core::engine::tests::rounding
cargo test -p visi --test cli_tests
```

## OOXML Spec Validation & Fuzzing

The ECMA-376 / `[MS-XLSX]` specifications are the source of truth.

```bash
python -m venv .venv
source .venv/bin/activate
pip install "lxml>=5.0" "maturin>=1.7" "pytest>=8.0"
maturin develop -m visi-python/Cargo.toml --release

# validate exported/patched workbooks against OPC rules, semantic invariants, and optional ECMA-376 XSDs
pytest visi-python/tests/ tools/ooxml-validate/
python tools/ooxml-validate/validate_ooxml.py path/to/file.xlsx --xsd-dir path/to/ecma376-xsds --require-xsd
```

### Rust fuzzing

Crash/panic fuzzing for Rust code lives in `visi-core/fuzz/`.

```bash
# needs a nightly toolchain
cargo install cargo-fuzz
cd visi-core && cargo +nightly fuzz run ovba_decompress
cargo +nightly fuzz run xlsx_import
mkdir -p fuzz/corpus/vba_import
cargo +nightly fuzz run vba_import fuzz/corpus/vba_import fuzz/seeds/vba_import   # seed corpus gets past the CFB-magic-bytes gate
mkdir -p fuzz/corpus/vba_parse
cargo +nightly fuzz run vba_parse fuzz/corpus/vba_parse fuzz/seeds/vba_parse       # VBA source text, not a binary
```

## Architecture

Cargo workspace, edition 2024:

- **`visi-core`** — the engine, published to crates.io as `visi-core` (the directory matches). Plain `rlib` — it kept a `cdylib` crate-type for a while without a single `extern "C"` symbol behind it; embedding for another language belongs in a separate crate (as `visi-python` now does), since `crate-type` can't be feature-gated. Still meant to stay embeddable in the sense that matters (no CLI/IO assumptions in `core`). Uses `web-time` instead of `std::time` and `getrandom` for IDs so it can target wasm — the browser JS backend is behind the **`wasm` feature** (`getrandom/js`), off by default because a library must not force a global getrandom backend on its consumers.
- **`visi`** — usage-rs-based CLI. `cli.rs` is the arg surface, `main.rs` holds one `handle_*` fn per subcommand, `engine.rs` wraps everything in `WorkbookManager`.
  - be sure to follow [Command Line Interface Guidelines](https://clig.dev) when making changes to the CLI
  - the CLI keeps its own `Result<_, String>` style internally and converts at the boundary (`exit_with_error` takes `impl Display`)
- **`visi-python`** — pyo3 bindings over `visi-core`, exposed to Python as the module `visi_core`. `crate-type = ["cdylib"]`, `abi3-py39`, `publish = false`; built with maturin. Three things about it are load-bearing:
  - It depends on **`visi-core` only**, never on `visi`. Where it has to mirror CLI behavior — `edit_chart`'s `--title`/`--clear-title` pair, `add_pivot_field`'s post-add subtotal/label mutation (`visi/src/main.rs`'s AddField arm), and `add_macro`'s sheet-name-to-id resolution with its `ThisWorkbook` exemption (the Add arm of `handle_macro`) — that mirroring is duplicated logic, and `visi-python/tests/test_cli_parity.py` is the only thing that will notice it drifting.
  - `extension-module` is **not** a default cargo feature. Turning it on breaks `cargo test --workspace`'s link step with an undefined `_PyModule_Create2`, whose message points nowhere near the cause. maturin enables it via `pyproject.toml`.
  - The Python module is named `visi_core`, not `visi`: the repo root holds a `visi/` directory with no `__init__.py`, which PEP 420 makes an implicit namespace package, so `import visi` from the root resolves to the CLI crate's source directory.

  Adding a binding is not a reason to widen `visi-core`'s public API — reach through the existing public fields, as `get_cell` does.

The two crates version independently: `visi` is at the workspace version, `visi-core` pins its own (`0.1.0`) since it is newer to crates.io.

### Public API surface

Not everything in `core` is public. The modules implementing Excel's function library — `stats`, `math_trig`, `text`, `date_fn`, `date`, `engineering`, `finance`, `extended_fn`, `ets`, `xml`, `vba_synth`, `pivot_xlsx`, `parser`, `formula`, `actions`, `shared_vec` — are `pub(crate)`; their types reach users only through the curated `pub use` list at the bottom of `core/mod.rs`. `ovba`, `vba_xlsx`, and `vba`'s `ast`/`lexer`/`parser` submodules are `#[doc(hidden)] pub` because `visi-core/fuzz` and the `dump_vba_fuzz_seeds` example need them, not because they are supported — the VBA syntax layer's supported surface is `check_syntax`/`ModuleSyntax`, and the AST's shape is deliberately not a semver commitment until the interpreter phases need it.

**When adding a public item, ask whether it belongs in that re-export list.** Anything reachable from `core`'s `pub use` is a semver commitment.

### CLI conventions (`visi/`)

Follows clig.dev. `-` means stdin/stdout for the file argument. Writes require either `--output <path>` or `--in-place`/`-i` (`resolve_output_path` in `main.rs`). User-facing row/column indices are **1-based** (or letters for columns) and converted in `utils.rs`; `set` accepts repeated `-S A1=100` pairs. `--quiet` suppresses informational stderr; most commands also take `--json`/`--format`.

## Tests

- `visi-core/src/core/engine/tests/unit.rs` — hand-written engine tests.
- `visi-core/src/core/engine/tests/{aggregate,logical,math,rounding,text}.rs` — spec/regression test cases, each a literal grid fed to the local `create_sheet` helper plus an assertion on one cell.
- `visi-core/src/core/xlsx/validate.rs` — pure-Rust OOXML OPC/sequence/semantic invariant validator and round-trip conformance tests.
- `visi-core/src/core/table.rs`, `pivot.rs`, and `xlsx/` have inline `#[cfg(test)] mod tests` for table CRUD, pivot computation/grouping, and xlsx round-tripping.
- `visi-core/src/core/vba/host.rs` has inline tests over the VBA host object model.
- `visi/tests/cli_tests.rs` — integration tests that drive `WorkbookManager` (the same API the CLI handlers call) through real file round-trips.

## Source Code Comments

Never add comments to source code.

## Documentation/Prose

When writing documentation, PR descriptions, or review comments, prefix
with the standalone line:

_The following content is LLM-generated._
