//! The error type returned by `visi-core`'s public API.

use crate::core::engine::EngineError;

/// The kind of workbook object an [`Error`] refers to.
///
/// Used by the [`Error::NotFound`] / [`Error::AlreadyExists`] /
/// [`Error::NameTaken`] variants so callers can distinguish "no such sheet"
/// from "no such table" without parsing the message text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ObjectKind {
    /// A worksheet.
    Sheet,
    /// An Excel Table (ListObject) -- a named range with a header row, not a
    /// worksheet. See [`crate::core::ExcelTable`].
    Table,
    /// A column within an Excel Table.
    TableColumn,
    /// A pivot table.
    PivotTable,
    /// A field within a pivot table.
    PivotField,
    /// A chart.
    Chart,
    /// A VBA module.
    VbaModule,
}

impl ObjectKind {
    /// The human-readable name used in error messages ("sheet", "table", ...).
    pub fn as_str(self) -> &'static str {
        match self {
            ObjectKind::Sheet => "sheet",
            ObjectKind::Table => "table",
            ObjectKind::TableColumn => "table column",
            ObjectKind::PivotTable => "pivot table",
            ObjectKind::PivotField => "pivot field",
            ObjectKind::Chart => "chart",
            ObjectKind::VbaModule => "VBA module",
        }
    }
}

impl std::fmt::Display for ObjectKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.as_str())
    }
}

/// Errors returned by `visi-core`'s public API.
///
/// This enum is `#[non_exhaustive]`: match with a `_` arm, since new variants
/// may be added in a minor release.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Error {
    /// No object of this kind goes by this name (or id, for charts).
    NotFound {
        /// What was being looked up.
        kind: ObjectKind,
        /// The name that was not found.
        name: String,
        /// The names that *do* exist, when the call can supply them cheaply,
        /// so callers can render a "did you mean" hint. Often empty.
        available: Vec<String>,
    },
    /// An object of this kind already goes by this name, so it cannot be added.
    AlreadyExists {
        /// What was being added.
        kind: ObjectKind,
        /// The name that collided.
        name: String,
    },
    /// A rename was rejected because the new name is already in use.
    ///
    /// Distinct from [`Error::AlreadyExists`], which is raised when *creating*.
    NameTaken {
        /// What was being renamed.
        kind: ObjectKind,
        /// The requested new name.
        name: String,
    },
    /// A name was rejected as structurally invalid, independent of collisions.
    InvalidName {
        /// What was being named.
        kind: ObjectKind,
        /// The rejected name.
        name: String,
        /// Why it was rejected.
        reason: String,
    },
    /// A row or column index fell outside the sheet.
    OutOfBounds {
        /// What was being indexed ("row" or "column").
        what: &'static str,
        /// The offending 0-based index.
        index: usize,
        /// The number of rows/columns that exist.
        len: usize,
    },
    /// A cell range was malformed -- for example, an end before its start.
    InvalidRange(String),
    /// The operation needs at least one sheet and the workbook has none.
    EmptyWorkbook,
    /// The last remaining sheet cannot be deleted; a workbook needs one.
    LastSheetInWorkbook,
    /// A worksheet can carry only one bound VBA document module.
    DocumentModuleExists,
    /// The operation was rejected by a lower layer that does not yet report a
    /// typed error -- currently the Excel Table and pivot internals.
    ///
    /// Carries message text only. Do not match on the string; variants will be
    /// carved out of this one as those layers are typed, which is why [`Error`]
    /// is `#[non_exhaustive]`.
    InvalidArgument(String),
    /// Reading or writing the `.xlsx` container failed.
    Xlsx(String),
    /// Reading or writing the VBA project failed.
    Vba(String),
    /// Formula evaluation failed.
    Eval(EngineError),
}

impl Error {
    /// A [`Error::NotFound`] with no "did you mean" candidates.
    pub fn not_found(kind: ObjectKind, name: impl Into<String>) -> Self {
        Error::NotFound {
            kind,
            name: name.into(),
            available: Vec::new(),
        }
    }

    /// A [`Error::NotFound`] that also carries the names that do exist.
    pub fn not_found_among(
        kind: ObjectKind,
        name: impl Into<String>,
        available: Vec<String>,
    ) -> Self {
        Error::NotFound {
            kind,
            name: name.into(),
            available,
        }
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::NotFound {
                kind,
                name,
                available,
            } => {
                write!(f, "{kind} '{name}' not found")?;
                if !available.is_empty() {
                    write!(f, ". Available {kind}s: {}", available.join(", "))?;
                }
                Ok(())
            }
            Error::AlreadyExists { kind, name } => write!(f, "{kind} '{name}' already exists"),
            Error::NameTaken { kind, name } => {
                write!(f, "{kind} name '{name}' is already taken")
            }
            Error::InvalidName { kind, name, reason } => {
                write!(f, "invalid {kind} name '{name}': {reason}")
            }
            Error::OutOfBounds { what, index, len } => {
                write!(f, "{what} index {index} is out of bounds (sheet has {len})")
            }
            Error::InvalidRange(msg) => write!(f, "invalid range: {msg}"),
            Error::EmptyWorkbook => f.write_str("workbook contains no sheets"),
            Error::LastSheetInWorkbook => {
                f.write_str("cannot delete the only sheet in the workbook")
            }
            Error::DocumentModuleExists => {
                f.write_str("that sheet already has a bound document module")
            }
            Error::InvalidArgument(msg) => f.write_str(msg),
            Error::Xlsx(msg) => write!(f, "xlsx error: {msg}"),
            Error::Vba(msg) => write!(f, "VBA error: {msg}"),
            Error::Eval(err) => write!(f, "evaluation error: {err}"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Eval(err) => Some(err),
            _ => None,
        }
    }
}

impl From<EngineError> for Error {
    fn from(err: EngineError) -> Self {
        Error::Eval(err)
    }
}

/// A `Result` whose error type is [`Error`].
pub type Result<T> = std::result::Result<T, Error>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn display_reads_naturally() {
        let e = Error::not_found(ObjectKind::PivotTable, "Sales");
        assert_eq!(e.to_string(), "pivot table 'Sales' not found");

        let e = Error::NameTaken {
            kind: ObjectKind::Sheet,
            name: "Data".into(),
        };
        assert_eq!(e.to_string(), "sheet name 'Data' is already taken");
    }

    #[test]
    fn is_a_std_error() {
        fn assert_std_error<E: std::error::Error>(_: &E) {}
        assert_std_error(&Error::EmptyWorkbook);
        let boxed: Box<dyn std::error::Error> = Box::new(Error::EmptyWorkbook);
        assert_eq!(boxed.to_string(), "workbook contains no sheets");
    }

    #[test]
    fn callers_can_match_on_kind_without_parsing_text() {
        let e = Error::not_found(ObjectKind::Table, "Q1");
        assert!(matches!(
            e,
            Error::NotFound {
                kind: ObjectKind::Table,
                ..
            }
        ));
    }
}
