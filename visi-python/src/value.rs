//! `ResultData` <-> Python conversion.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use visi_engine::core::ResultData;

/// An Excel error *value* (`#DIV/0!`, `#VALUE!`, `#N/A`) sitting in a cell.
///
/// A distinct type rather than a plain `str`, because a cell can legitimately
/// hold the *text* `#DIV/0!` -- `=CONCATENATE("#DIV/0!")` produces exactly
/// that -- and returning both as `str` would erase the difference. **The type
/// is what carries the distinction: use `isinstance(v, CellError)`, not `==`.**
///
/// `__eq__` compares equal to the bare code string, so
/// `wb.get_cell(0, 0) == "#DIV/0!"` reads true for a real error. That
/// convenience necessarily also makes it true for a text cell holding those
/// characters -- the two are equal *as values*, and only their types differ.
/// This matches the harness's existing convention: `XLSXEvaluatedReader`
/// normalizes a `t="e"` cell to the upper-cased code string, i.e. the oracle
/// has always compared errors and text by value. `CellError` adds information
/// on top of that; it does not change the comparison.
// `from_py_object` is opted into deliberately: `__eq__` extracts a `CellError`
// from its argument to compare two of them.
#[pyclass(module = "visi_core", frozen, from_py_object)]
#[derive(Clone)]
pub struct CellError {
    /// The Excel error code, e.g. `"#DIV/0!"`.
    #[pyo3(get)]
    pub code: String,
}

#[pymethods]
impl CellError {
    #[new]
    fn new(code: String) -> Self {
        Self { code }
    }

    fn __str__(&self) -> &str {
        &self.code
    }

    fn __repr__(&self) -> String {
        format!("CellError({:?})", self.code)
    }

    // Hand-written rather than `#[pyclass(eq)]`: comparing equal to a bare
    // `str` is the point, and the derived version only compares against its
    // own type.
    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        if let Ok(s) = other.extract::<String>() {
            return self.code == s;
        }
        other
            .extract::<CellError>()
            .map(|o| o.code == self.code)
            .unwrap_or(false)
    }

    // Must agree with `__eq__`: a CellError and its code string compare equal,
    // so they have to hash equal too, or dict/set membership disagrees with
    // `==`.
    fn __hash__(&self, py: Python<'_>) -> PyResult<isize> {
        self.code.as_str().into_pyobject(py)?.hash()
    }
}

/// Converts one [`ResultData`] into a Python object.
///
/// The match is deliberately total -- no `_` arm -- so that a new `ResultData`
/// variant fails to compile here rather than silently converting to `None`.
pub fn result_to_py<'py>(py: Python<'py>, v: &ResultData) -> PyResult<Bound<'py, PyAny>> {
    Ok(match v {
        ResultData::None => py.None().into_bound(py),
        ResultData::Boolean(b) => b.into_pyobject(py)?.to_owned().into_any(),
        ResultData::Integer(i) => i.into_pyobject(py)?.into_any(),
        ResultData::Float(f) => f.into_pyobject(py)?.into_any(),
        ResultData::String(s) => s.into_pyobject(py)?.into_any(),
        ResultData::List(items) => {
            let converted = items
                .iter()
                .map(|x| result_to_py(py, x))
                .collect::<PyResult<Vec<_>>>()?;
            PyList::new(py, converted)?.into_any()
        }
        ResultData::Dict(pairs) => {
            let d = PyDict::new(py);
            for (k, val) in pairs {
                d.set_item(hashable_key(py, k)?, result_to_py(py, val)?)?;
            }
            d.into_any()
        }
        ResultData::Error(code) => Bound::new(py, CellError { code: code.clone() })?.into_any(),
    })
}

/// Converts a value being used as a dict *key*.
///
/// Same as [`result_to_py`] except that a `List` becomes a tuple, since a
/// Python list is unhashable and would make the whole dict unbuildable.
fn hashable_key<'py>(py: Python<'py>, v: &ResultData) -> PyResult<Bound<'py, PyAny>> {
    match v {
        ResultData::List(items) => {
            let converted = items
                .iter()
                .map(|x| hashable_key(py, x))
                .collect::<PyResult<Vec<_>>>()?;
            Ok(PyTuple::new(py, converted)?.into_any())
        }
        other => result_to_py(py, other),
    }
}
