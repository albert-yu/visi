//! Filesystem and stdio access for [`WorkbookManager`].
//!
//! The workbook model itself lives in `visi-core` and is deliberately
//! byte-oriented, so that it stays usable when embedded somewhere without a
//! filesystem (wasm, for instance). The path- and stdio-based conventions
//! this CLI needs -- including clig.dev's `-` meaning stdin/stdout -- are
//! layered on here instead, as an extension trait.
//!
//! [`WorkbookManager`] is re-exported so `visi::engine::WorkbookManager`
//! keeps resolving for existing callers.

pub use visi_core::{SheetSummary, WorkbookManager, WorkbookSummary};

use std::fs;
use std::io::{self, Read, Write};
use std::path::Path;

/// Path- and stdio-based loading and saving, layered over `visi-core`'s
/// byte-oriented [`WorkbookManager::load_bytes`]/[`WorkbookManager::save_bytes`].
///
/// A `path_str` of `-` means stdin (loading) or stdout (saving).
pub trait WorkbookFile: Sized {
    /// Load an Excel workbook from a file path, or stdin for `-`.
    fn load_file(path_str: &str) -> Result<Self, String>;

    /// Load an Excel workbook from a file path, creating a new empty
    /// workbook if the file does not exist.
    fn load_file_or_create(path_str: &str) -> Result<Self, String>;

    /// Save an Excel workbook to a file path, or stdout for `-`. Missing
    /// parent directories are created.
    fn save_file(&self, path_str: &str) -> Result<(), String>;
}

impl WorkbookFile for WorkbookManager {
    fn load_file(path_str: &str) -> Result<Self, String> {
        let buffer = if path_str == "-" {
            let mut stdin_bytes = Vec::new();
            io::stdin()
                .read_to_end(&mut stdin_bytes)
                .map_err(|e| format!("Failed to read stdin: {}", e))?;
            stdin_bytes
        } else {
            fs::read(path_str).map_err(|e| format!("Failed to read file '{}': {}", path_str, e))?
        };

        Self::load_bytes(&buffer)
    }

    fn load_file_or_create(path_str: &str) -> Result<Self, String> {
        if path_str != "-" && !Path::new(path_str).exists() {
            Self::new_empty()
        } else {
            Self::load_file(path_str)
        }
    }

    fn save_file(&self, path_str: &str) -> Result<(), String> {
        let bytes = self.save_bytes()?;

        if path_str == "-" {
            io::stdout()
                .write_all(&bytes)
                .map_err(|e| format!("Failed to write to stdout: {}", e))?;
            io::stdout()
                .flush()
                .map_err(|e| format!("Failed to flush stdout: {}", e))?;
        } else {
            let path = Path::new(path_str);
            if let Some(parent) = path.parent()
                && !parent.as_os_str().is_empty()
            {
                fs::create_dir_all(parent)
                    .map_err(|e| format!("Failed to create parent directories: {}", e))?;
            }
            fs::write(path_str, bytes)
                .map_err(|e| format!("Failed to save file to '{}': {}", path_str, e))?;
        }
        Ok(())
    }
}
