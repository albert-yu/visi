#![allow(missing_docs)]
use crate::core::engine::EngineError;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ObjectKind {
    Sheet,

    Table,

    TableColumn,

    PivotTable,

    PivotField,

    Chart,

    VbaModule,
}

impl ObjectKind {
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

#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Error {
    NotFound {
        kind: ObjectKind,

        name: String,

        available: Vec<String>,
    },

    AlreadyExists {
        kind: ObjectKind,

        name: String,
    },

    NameTaken {
        kind: ObjectKind,

        name: String,
    },

    InvalidName {
        kind: ObjectKind,

        name: String,

        reason: String,
    },

    OutOfBounds {
        what: &'static str,

        index: usize,

        len: usize,
    },

    InvalidRange(String),

    EmptyWorkbook,

    LastSheetInWorkbook,

    DocumentModuleExists,

    InvalidArgument(String),

    Xlsx(String),

    Vba(String),

    VbaSyntax {
        message: String,

        module: Option<String>,

        line: u32,

        column: u32,
    },

    VbaRuntime {
        message: String,

        number: i32,
    },

    Eval(EngineError),
}

impl Error {
    pub fn not_found(kind: ObjectKind, name: impl Into<String>) -> Self {
        Error::NotFound {
            kind,
            name: name.into(),
            available: Vec::new(),
        }
    }

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
            Error::VbaRuntime { message, number } => {
                write!(f, "run-time error {number}: {message}")
            }
            Error::VbaSyntax {
                message,
                module,
                line,
                column,
            } => match module {
                Some(m) => write!(f, "{m}({line},{column}): {message}"),
                None => write!(f, "line {line}, column {column}: {message}"),
            },
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
