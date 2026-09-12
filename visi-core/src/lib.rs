#![warn(missing_docs)]
#![doc = "Core spreadsheet engine, workbook model, and file-format helpers for Visi."]

/// The engine's modules: sheets and cells, Excel Tables, pivot tables,
/// charts, styling, VBA, and `.xlsx` I/O.
///
/// The curated re-exports at this module's root are the intended surface.
/// Modules implementing Excel's function library are crate-private.
pub mod core;
mod error;

pub use core::workbook::{SheetSummary, WorkbookManager, WorkbookSummary};
pub use core::xlsx::{export_xlsx_data, import_xlsx_data};
pub use error::{Error, ObjectKind, Result};
