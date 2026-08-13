//! `visi-core`: Core logic for the visi spreadsheet application.
//! Designed for embedding in other applications (C, C++, Python, NodeJS, etc.)
//! and consumption by the `visi` CLI binary.
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
//! # fn main() -> Result<(), String> {
//! let bytes = std::fs::read("book.xlsx").map_err(|e| e.to_string())?;
//! let mut wb = WorkbookManager::load_bytes(&bytes)?;
//!
//! // 0-based (row, col); A1 notation is a boundary concern.
//! wb.set_cell(0, 0, 0, "=SUM(Sheet2!A1:A10)".to_string());
//! wb.evaluate()?;
//!
//! std::fs::write("out.xlsx", wb.save_bytes()?).map_err(|e| e.to_string())?;
//! # Ok(())
//! # }
//! ```

pub mod core;

pub use core::workbook::{SheetSummary, WorkbookManager, WorkbookSummary};
pub use core::xlsx::{export_xlsx_data, import_xlsx_data};
