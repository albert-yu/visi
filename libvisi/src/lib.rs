//! `libvisi`: Core logic for the visi spreadsheet application.
//! Designed for embedding in other applications (C, C++, Python, NodeJS, etc.)
//! and consumption by the `visi` CLI binary.

pub mod core;

pub use core::xlsx::{export_xlsx_data, import_xlsx_data};
