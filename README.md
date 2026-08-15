# visi

[![CI](https://github.com/albert-yu/visi/actions/workflows/ci.yml/badge.svg)](https://github.com/albert-yu/visi/actions/workflows/ci.yml)

A spreadsheet engine for editing and evaluating Excel (`.xlsx`) files.

My goals with this project are:

1. Match Excel's execution behavior 100% (or, as much as possible without a UI)
2. Prioritize performance, making it possible to handle large workloads

`visi` is structured follows:
- **[`visi-core`](visi-core/)**: embeddedable spreadsheet engine providing Excel parsing, AST formula compilation, dependency resolution, execution engine, date calculations, chart metadata, and Excel (`.xlsx`) import/export, see [its README](visi-core/README.md).
- **[`visi`](visi/)**: Command-line application using `visi-core` which can edit and execute Excel files headlessly

`visi` aims for parity by using [fuzz testing](fuzz/README.md) and throwing
LLM tokens at it.

---

## Testing

Run all unit and integration tests across `visi-core` and `visi`:

```bash
cargo test --workspace
```

---

## License

Dual-licensed:

- [Apache License, Version 2.0](LICENSE-APACHE)
- [MIT License](LICENSE-MIT)
