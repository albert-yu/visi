//! `visi_core::Error` as a Python exception hierarchy.
//!
//! Every exception derives from [`VisiError`], so `except VisiError` catches
//! anything the engine raises. The subclass says *what kind* of failure it was
//! without parsing message text, and the structured payload each variant
//! carries (`ObjectKind`, the offending name, the available names) is exposed
//! as instance attributes.
//!
//! Two invariants worth keeping:
//!
//! - **`args` is always a 1-tuple of the message.** That makes `str(exc)`
//!   equal to `Error`'s `Display`, which is exactly what the CLI prints to
//!   stderr minus its `"Error: "` prefix -- so `fuzz/test_backend_parity.py`
//!   can compare error text across the two backends. Putting the payload in
//!   `args` instead would make `str(exc)` render a tuple.
//! - **`visi_core::Error` is `#[non_exhaustive]`**, so the match below has a
//!   `_` arm that maps to the base `VisiError`. A variant added upstream must
//!   widen to the base class, never land in a wrong subclass.

use pyo3::create_exception;
use pyo3::exceptions::PyException;
use pyo3::prelude::*;
use visi_engine::Error as CoreError;

create_exception!(
    visi_core,
    VisiError,
    PyException,
    "Base class for every error raised by visi-core."
);
create_exception!(
    visi_core,
    NotFoundError,
    VisiError,
    "No object of the requested kind goes by that name. Carries `kind`, `name`, `available`."
);
create_exception!(
    visi_core,
    AlreadyExistsError,
    VisiError,
    "An object of that kind already has this name, so it cannot be created. Carries `kind`, `name`."
);
create_exception!(
    visi_core,
    NameTakenError,
    VisiError,
    "A rename was rejected because the new name is in use. Carries `kind`, `name`."
);
create_exception!(
    visi_core,
    InvalidNameError,
    VisiError,
    "A name was structurally invalid, independent of collisions. Carries `kind`, `name`, `reason`."
);
create_exception!(
    visi_core,
    OutOfBoundsError,
    VisiError,
    "A row or column index fell outside the sheet. Carries `what`, `index`, `length`."
);
create_exception!(
    visi_core,
    InvalidRangeError,
    VisiError,
    "A cell range was malformed -- for example, an end before its start."
);
create_exception!(
    visi_core,
    InvalidArgumentError,
    VisiError,
    "Rejected by a layer that does not yet report a typed error, or by an argument check in these bindings."
);
create_exception!(
    visi_core,
    XlsxError,
    VisiError,
    "Reading or writing the .xlsx container failed."
);
create_exception!(
    visi_core,
    VbaError,
    VisiError,
    "Reading or writing the VBA project failed."
);
create_exception!(
    visi_core,
    VbaSyntaxError,
    VbaError,
    "A VBA module failed to parse. Carries `module`, `line`, `column`."
);
create_exception!(
    visi_core,
    EvaluationError,
    VisiError,
    "Formula evaluation failed. Note that per-cell formula errors do NOT raise this -- they arrive as `CellError` values in cells."
);
create_exception!(
    visi_core,
    EmptyWorkbookError,
    VisiError,
    "The operation needs at least one sheet and the workbook has none."
);
create_exception!(
    visi_core,
    LastSheetError,
    VisiError,
    "The last remaining sheet cannot be deleted; a workbook needs one."
);
create_exception!(
    visi_core,
    DocumentModuleExistsError,
    VisiError,
    "A worksheet can carry only one bound VBA document module."
);

/// Newtype around [`CoreError`] so `?` works inside `#[pymethods]`.
///
/// Needed because both `visi_core::Error` and `PyErr` are foreign to this
/// crate, so the orphan rule forbids `impl From<CoreError> for PyErr`.
pub struct Wrapped(pub CoreError);

impl From<CoreError> for Wrapped {
    fn from(e: CoreError) -> Self {
        Wrapped(e)
    }
}

impl From<Wrapped> for PyErr {
    fn from(w: Wrapped) -> PyErr {
        let e = w.0;
        let msg = e.to_string();

        let err = match &e {
            CoreError::NotFound { .. } => PyErr::new::<NotFoundError, _>((msg,)),
            CoreError::AlreadyExists { .. } => PyErr::new::<AlreadyExistsError, _>((msg,)),
            CoreError::NameTaken { .. } => PyErr::new::<NameTakenError, _>((msg,)),
            CoreError::InvalidName { .. } => PyErr::new::<InvalidNameError, _>((msg,)),
            CoreError::OutOfBounds { .. } => PyErr::new::<OutOfBoundsError, _>((msg,)),
            CoreError::InvalidRange(_) => PyErr::new::<InvalidRangeError, _>((msg,)),
            CoreError::InvalidArgument(_) => PyErr::new::<InvalidArgumentError, _>((msg,)),
            CoreError::Xlsx(_) => PyErr::new::<XlsxError, _>((msg,)),
            CoreError::Vba(_) => PyErr::new::<VbaError, _>((msg,)),
            // A subclass of VbaError, so `except VbaError` still catches it.
            CoreError::VbaSyntax { .. } => PyErr::new::<VbaSyntaxError, _>((msg,)),
            CoreError::Eval(_) => PyErr::new::<EvaluationError, _>((msg,)),
            CoreError::EmptyWorkbook => PyErr::new::<EmptyWorkbookError, _>((msg,)),
            CoreError::LastSheetInWorkbook => PyErr::new::<LastSheetError, _>((msg,)),
            CoreError::DocumentModuleExists => PyErr::new::<DocumentModuleExistsError, _>((msg,)),
            // `Error` is #[non_exhaustive]: widen to the base class rather than
            // guessing at a subclass or panicking.
            _ => PyErr::new::<VisiError, _>((msg,)),
        };

        // The structured payload goes on attributes, leaving `args` as
        // `(message,)`. Attribute writes are best-effort: failing to attach a
        // "did you mean" hint must not replace the real error with a different
        // one.
        Python::attach(|py| {
            let v = err.value(py);
            match &e {
                CoreError::NotFound {
                    kind,
                    name,
                    available,
                } => {
                    let _ = v.setattr("kind", kind.as_str());
                    let _ = v.setattr("name", name.as_str());
                    let _ = v.setattr("available", available.clone());
                }
                CoreError::AlreadyExists { kind, name } | CoreError::NameTaken { kind, name } => {
                    let _ = v.setattr("kind", kind.as_str());
                    let _ = v.setattr("name", name.as_str());
                }
                CoreError::InvalidName { kind, name, reason } => {
                    let _ = v.setattr("kind", kind.as_str());
                    let _ = v.setattr("name", name.as_str());
                    let _ = v.setattr("reason", reason.as_str());
                }
                CoreError::VbaSyntax {
                    module,
                    line,
                    column,
                    ..
                } => {
                    let _ = v.setattr("module", module.clone());
                    let _ = v.setattr("line", *line);
                    let _ = v.setattr("column", *column);
                }
                CoreError::OutOfBounds { what, index, len } => {
                    let _ = v.setattr("what", *what);
                    let _ = v.setattr("index", *index);
                    // Not `len`: shadowing the builtin on an exception object
                    // reads badly at a REPL, and `len(exc)` is not a thing.
                    let _ = v.setattr("length", *len);
                }
                _ => {}
            }
        });

        err
    }
}

/// An argument these bindings rejected before it reached visi-core.
///
/// Separate from [`Wrapped`] so argument validation and engine failures do not
/// have to be spelled the same way at every call site.
pub fn invalid_argument(msg: impl std::fmt::Display) -> PyErr {
    PyErr::new::<InvalidArgumentError, _>((msg.to_string(),))
}

/// Registers every exception class on the module.
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = m.py();
    m.add("VisiError", py.get_type::<VisiError>())?;
    m.add("NotFoundError", py.get_type::<NotFoundError>())?;
    m.add("AlreadyExistsError", py.get_type::<AlreadyExistsError>())?;
    m.add("NameTakenError", py.get_type::<NameTakenError>())?;
    m.add("InvalidNameError", py.get_type::<InvalidNameError>())?;
    m.add("OutOfBoundsError", py.get_type::<OutOfBoundsError>())?;
    m.add("InvalidRangeError", py.get_type::<InvalidRangeError>())?;
    m.add(
        "InvalidArgumentError",
        py.get_type::<InvalidArgumentError>(),
    )?;
    m.add("XlsxError", py.get_type::<XlsxError>())?;
    m.add("VbaError", py.get_type::<VbaError>())?;
    m.add("VbaSyntaxError", py.get_type::<VbaSyntaxError>())?;
    m.add("EvaluationError", py.get_type::<EvaluationError>())?;
    m.add("EmptyWorkbookError", py.get_type::<EmptyWorkbookError>())?;
    m.add("LastSheetError", py.get_type::<LastSheetError>())?;
    m.add(
        "DocumentModuleExistsError",
        py.get_type::<DocumentModuleExistsError>(),
    )?;
    Ok(())
}
