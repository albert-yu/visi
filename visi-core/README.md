# visi-core

[![crates.io](https://img.shields.io/crates/v/visi-core.svg)](https://crates.io/crates/visi-core)
[![docs.rs](https://docs.rs/visi-core/badge.svg)](https://docs.rs/visi-core)

An embeddable spreadsheet engine: Excel formula compilation and evaluation,
a dependency-tracking recalculation engine, and `.xlsx` import/export.

`visi-core` is the engine behind the [`visi`](https://crates.io/crates/visi)
command-line tool, published separately so it can be embedded directly. It
makes no CLI or filesystem assumptions — everything is driven through byte
buffers — and it uses `web-time` and `getrandom` rather than `std::time` so
it can target wasm.

## Usage

```toml
[dependencies]
visi-core = "0.1"
```

[`WorkbookManager`] is the entry point. It owns a workbook's sheets, charts,
pivot tables and VBA project, and is the layer that makes cross-sheet formulas
behave correctly — prefer it over reaching for `core::engine::Sheet` directly.

```rust,no_run
use visi_core::WorkbookManager;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read("book.xlsx")?;
    let mut wb = WorkbookManager::load_bytes(&bytes)?;

    // 0-based (row, col); A1 notation is a boundary concern.
    wb.set_cell(0, 0, 0, "=SUM(Sheet2!A1:A10)".to_string());
    wb.evaluate()?;

    std::fs::write("out.xlsx", wb.save_bytes()?)?;
    Ok(())
}
```

[`WorkbookManager`]: https://docs.rs/visi-core/latest/visi_core/struct.WorkbookManager.html

Fallible calls return `visi_core::Error`, which implements `std::error::Error`.
Failures that name a workbook object carry an `ObjectKind`, so you can react
without parsing message text:

```rust,no_run
use visi_core::{Error, ObjectKind, WorkbookManager};

fn f(wb: &mut WorkbookManager) -> Result<(), Box<dyn std::error::Error>> {
    match wb.rename_sheet("Sheet1", "Data") {
        Err(Error::NotFound { kind: ObjectKind::Sheet, name, .. }) => {
            eprintln!("no sheet called {name}");
        }
        other => other?,
    }
    Ok(())
}
```

## What it does

- **Formulas** — 500+ Excel functions across math/trig, statistics, text,
  date/time, financial, engineering, lookup, and information families.
  Formula text compiles to a representation that stores references by
  sheet/column *id* rather than name, so renames are non-destructive.
- **Recalculation** — a dirty-queue BFS maintaining both directions of the
  dependency graph, with cross-sheet propagation handled at the workbook level.
- **Excel fidelity** — number formatting reproduces Excel's 15-significant-digit
  display rules, and divergences are tracked deliberately (see
  [`docs/excel-discrepancies.md`](https://github.com/albert-yu/visi/blob/main/docs/excel-discrepancies.md))
  against a differential fuzzer that runs generated workbooks through real
  Microsoft Excel.
- **Excel objects** — tables (ListObjects) with structured references, pivot
  tables, charts, cell styles, and VBA project round-tripping.

## Features

| Feature | Default | Description |
| --- | --- | --- |
| `wasm` | no | Routes `getrandom` through the browser crypto API for `wasm32-unknown-unknown`. Enable only from a final binary, never a library. |

## Stability

Pre-1.0: the public API is still moving. Pin a minor version.

The modules implementing Excel's function library (statistics, text, financial,
engineering, date/time, the formula parser and AST) are crate-private — what is
re-exported from the crate root and from `core` is the intended surface.
`Error` is `#[non_exhaustive]`, so match it with a `_` arm.
