pub use visi_core::{SheetSummary, WorkbookManager, WorkbookSummary};

use std::fs;
use std::io::{self, Read, Write};
use std::path::Path;

/// Path- and stdio-based loading and saving
pub trait WorkbookFile: Sized {
    fn load_file(path_str: &str) -> Result<Self, String>;

    fn load_file_or_create(path_str: &str) -> Result<Self, String>;

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

        Self::load_bytes(&buffer).map_err(|e| e.to_string())
    }

    fn load_file_or_create(path_str: &str) -> Result<Self, String> {
        if path_str != "-" && !Path::new(path_str).exists() {
            Self::new_empty().map_err(|e| e.to_string())
        } else {
            Self::load_file(path_str)
        }
    }

    fn save_file(&self, path_str: &str) -> Result<(), String> {
        let bytes = self.save_bytes().map_err(|e| e.to_string())?;

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
