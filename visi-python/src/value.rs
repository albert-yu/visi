use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use visi_engine::core::ResultData;

#[pyclass(module = "visi_core", frozen, from_py_object)]
#[derive(Clone)]
pub struct CellError {
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

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        if let Ok(s) = other.extract::<String>() {
            return self.code == s;
        }
        other
            .extract::<CellError>()
            .map(|o| o.code == self.code)
            .unwrap_or(false)
    }

    fn __hash__(&self, py: Python<'_>) -> PyResult<isize> {
        self.code.as_str().into_pyobject(py)?.hash()
    }
}

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
