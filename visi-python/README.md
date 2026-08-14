# visi-python

Python bindings for [`visi-core`](../visi-core), exposed as the module
`visi_core`. **Development-only** — not published, no stability promise. It
exists so the differential fuzz harness in [`fuzz/`](../fuzz) can drive the
engine in process instead of spawning the `visi` CLI once per operation.

## Building

Into the project venv (`maturin develop` installs into whichever venv is
active, so activate it first):

```bash
source ../fuzz/venv/bin/activate
pip install -r ../fuzz/requirements.txt
maturin develop -m Cargo.toml --release
pytest tests/
```

## Usage

```python
import visi_core

wb = visi_core.Workbook.load("book.xlsx")
wb.set_cell(0, 0, "=SUM(Sheet2!A1:A10)")   # 0-based (row, col)
wb.evaluate()
wb.save("out.xlsx")
```

Formula failures are *values*, not exceptions — `evaluate()` almost never
raises:

```python
wb.set_cell(0, 0, "=1/0"); wb.evaluate()
v = wb.get_cell(0, 0)
isinstance(v, visi_core.CellError)   # True
v.code                               # '#DIV/0!'
```

`CellError` compares equal to its code string for convenience, which means a
cell holding the *text* `#DIV/0!` also compares equal to it. The type is what
distinguishes them; use `isinstance`, not `==`.

Engine failures raise a hierarchy under `VisiError`, carrying structured
payload rather than only a message:

```python
try:
    wb.refresh_pivot("nope")
except visi_core.NotFoundError as e:
    e.kind, e.name, e.available     # ('pivot table', 'nope', [...])
```

## Notes for maintainers

Three things here are load-bearing and easy to undo by accident:

- **`extension-module` is not a default cargo feature.** Enabling it by default
  breaks `cargo test --workspace`'s link step with an undefined
  `_PyModule_Create2`, an error that points nowhere near the cause. maturin
  turns it on through `pyproject.toml`. This is *not* what `maturin new`
  generates, so anyone tidying the manifest is likely to reintroduce it.
- **The module is `visi_core`, not `visi`.** The repo root holds a `visi/`
  directory with no `__init__.py`, which PEP 420 makes an implicit namespace
  package — `import visi` from the root resolves to the CLI crate's source
  directory. The `visi-core` dependency is aliased to `visi_engine` in
  `Cargo.toml` for the same reason: this crate's own lib is named `visi_core`.
- **Some behavior is mirrored from the CLI, not shared with it.** This crate
  depends on `visi-core` only, so `edit_chart`'s clear-vs-set flags and
  `add_pivot_field`'s post-add subtotal/label mutation are reimplementations of
  what `visi/src/main.rs` does. `fuzz/test_backend_parity.py` is the only thing
  that will catch them drifting apart.

Adding a binding is not a reason to widen `visi-core`'s public API. `get_cell`
reaches through the already-public `sheets` field rather than adding a method
to `WorkbookManager`, deliberately: `set_cell` writes source text, and the
plausible reads (`get_result_data`, `get_display_string`, `get_src_str`) are
three different things, so there is no single `get_cell` to bless.
