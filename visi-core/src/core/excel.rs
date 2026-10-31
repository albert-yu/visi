use std::io::Cursor;
use std::path::Path;
use std::str::FromStr;

use crate::core::{Sheet, chart::Chart, pivot::PivotTable, vba::VbaProject};

#[doc = "A worksheet materialized by an Excel file-format importer."]
#[derive(Debug, serde::Serialize, serde::Deserialize)]
pub struct ImportedSheet {
    #[doc = "The imported worksheet model."]
    pub sheet: Sheet,
}

#[doc = "Workbook data shared by the XLSX and XLSB import boundaries."]
pub type ImportedWorkbookData = (
    Vec<ImportedSheet>,
    Vec<Chart>,
    Vec<PivotTable>,
    Option<VbaProject>,
);

#[doc = "An Excel workbook package format supported for output."]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum WorkbookFormat {
    #[doc = "The XML-based `.xlsx` package format."]
    #[default]
    Xlsx,
    #[doc = "The binary BIFF12 `.xlsb` package format."]
    Xlsb,
}

impl WorkbookFormat {
    #[doc = "Infers a workbook format from an `.xlsx`, `.xlsm`, or `.xlsb` path."]
    pub fn from_path(path: impl AsRef<Path>) -> Option<Self> {
        match path
            .as_ref()
            .extension()
            .and_then(|extension| extension.to_str())
            .map(str::to_ascii_lowercase)
            .as_deref()
        {
            Some("xlsb") => Some(Self::Xlsb),
            Some("xlsx" | "xlsm") => Some(Self::Xlsx),
            _ => None,
        }
    }

    #[doc = "Returns the conventional lowercase file extension for this format."]
    pub fn extension(self) -> &'static str {
        match self {
            Self::Xlsx => "xlsx",
            Self::Xlsb => "xlsb",
        }
    }
}

impl std::fmt::Display for WorkbookFormat {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.extension())
    }
}

impl FromStr for WorkbookFormat {
    type Err = crate::Error;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value.trim().to_ascii_lowercase().as_str() {
            "xlsx" => Ok(Self::Xlsx),
            "xlsb" => Ok(Self::Xlsb),
            other => Err(crate::Error::InvalidArgument(format!(
                "unsupported workbook format '{other}'; expected xlsx or xlsb"
            ))),
        }
    }
}

#[doc = "Detects XLSX versus XLSB from the workbook part inside an OPC package."]
pub fn detect_workbook_format(buffer: &[u8]) -> crate::Result<WorkbookFormat> {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer))
        .map_err(|error| crate::Error::Xlsx(format!("failed to open Excel package: {error}")))?;
    let mut xlsx = false;
    let mut xlsb = false;
    for index in 0..archive.len() {
        let name = archive
            .by_index(index)
            .map_err(|error| crate::Error::Xlsx(error.to_string()))?
            .name()
            .to_ascii_lowercase();
        xlsx |= name == "xl/workbook.xml";
        xlsb |= name == "xl/workbook.bin";
    }
    match (xlsx, xlsb) {
        (true, false) => Ok(WorkbookFormat::Xlsx),
        (false, true) => Ok(WorkbookFormat::Xlsb),
        (false, false) => Err(crate::Error::Xlsx(
            "Excel package contains neither xl/workbook.xml nor xl/workbook.bin".to_string(),
        )),
        (true, true) => Err(crate::Error::Xlsx(
            "Excel package ambiguously contains both XLSX and XLSB workbook parts".to_string(),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_inference_is_case_insensitive() {
        assert_eq!(
            WorkbookFormat::from_path("book.XLSB"),
            Some(WorkbookFormat::Xlsb)
        );
        assert_eq!(
            WorkbookFormat::from_path("book.xlsm"),
            Some(WorkbookFormat::Xlsx)
        );
        assert_eq!(WorkbookFormat::from_path("book.csv"), None);
    }
}
