//! `visi-core`: an embeddable spreadsheet engine.
//!
//! Excel formula compilation and evaluation, dependency-tracked
//! recalculation, and `.xlsx` import/export. Makes no CLI or filesystem
//! assumptions -- everything is driven through byte buffers -- and uses
//! `web-time` and `getrandom` rather than `std::time` so it can target wasm.
//!
//! # Getting started
//!
//! [`WorkbookManager`] is the entry point. It owns a workbook's sheets,
//! charts, pivot tables and VBA project, and is the layer that makes
//! cross-sheet formulas and pivot tables behave correctly -- see its
//! documentation before reaching for [`core::engine::Sheet`] directly.
//!
//! ```no_run
//! use visi_core::WorkbookManager;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let bytes = std::fs::read("book.xlsx")?;
//! let mut wb = WorkbookManager::load_bytes(&bytes)?;
//!
//! // 0-based (row, col); A1 notation is a boundary concern.
//! wb.set_cell(0, 0, 0, "=SUM(Sheet2!A1:A10)".to_string());
//! wb.evaluate()?;
//!
//! std::fs::write("out.xlsx", wb.save_bytes()?)?;
//! # Ok(())
//! # }
//! ```
//!
//! # Errors
//!
//! Fallible calls return [`Error`], which implements [`std::error::Error`]
//! and so composes with `anyhow`, `eyre`, or `Box<dyn Error>`. Failures that
//! name a workbook object carry an [`ObjectKind`] rather than only a message,
//! so callers can react without parsing text:
//!
//! ```no_run
//! use visi_core::{Error, ObjectKind, WorkbookManager};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! # let mut wb = WorkbookManager::new_empty()?;
//! match wb.rename_sheet("Sheet1", "Data") {
//!     Err(Error::NotFound { kind: ObjectKind::Sheet, name, .. }) => {
//!         eprintln!("no sheet called {name}");
//!     }
//!     other => other?,
//! }
//! # Ok(())
//! # }
//! ```
//!
//! # Stability
//!
//! Pre-1.0; the public API is still moving. Modules that exist to implement
//! Excel's function library are crate-private -- what is re-exported here and
//! from [`core`] is the intended surface.

pub mod core;
mod error;

pub use core::workbook::{SheetSummary, WorkbookManager, WorkbookSummary};
pub use core::xlsx::{export_xlsx_data, import_xlsx_data};
pub use error::{Error, ObjectKind, Result};
