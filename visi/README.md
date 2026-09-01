# visi

CLI tool for authoring _and_ executing Excel workbooks. The execution
layer should match Excel's behavior 100%.

Libraries such as `openpyxl` can author Excel workbooks,
but they cannot evaluate formulas. `visi` enables headless shell-scripting
and LLM automation of Excel workflows.

---

## Building from source

### Requirements

- [Rust](https://www.rust-lang.org/) (2024 edition supported)

### Build Binary

```bash
# dev
cargo build --workspace

# release binary
cargo build --release --workspace
```
The compiled CLI executable will be located at `target/release/visi` (or `target/debug/visi`).

---

## Testing

Run all unit and integration tests across `visi-core` and `visi`:

```bash
cargo test --workspace
```

---
