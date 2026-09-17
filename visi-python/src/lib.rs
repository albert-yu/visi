use pyo3::prelude::*;
use visi_engine::WorkbookManager;

mod enums;
mod errors;
mod value;
mod workbook;

use errors::Wrapped;

#[pyfunction]
fn eval_file(input: std::path::PathBuf, output: std::path::PathBuf) -> PyResult<()> {
    let bytes = std::fs::read(&input)?;
    let mut wb = WorkbookManager::load_bytes(&bytes).map_err(Wrapped)?;
    wb.evaluate().map_err(Wrapped)?;
    std::fs::write(&output, wb.save_bytes().map_err(Wrapped)?)?;
    Ok(())
}

#[pyfunction]
fn check_syntax(source: &str) -> PyResult<Vec<String>> {
    Ok(visi_engine::core::check_syntax(source)
        .map_err(Wrapped)?
        .procedures)
}

#[pyfunction]
#[pyo3(signature = (source, procedure, args=None))]
fn run_macro(
    source: &str,
    procedure: &str,
    args: Option<Vec<String>>,
) -> PyResult<(String, Option<String>)> {
    let args = args.unwrap_or_default();
    let refs: Vec<&str> = args.iter().map(|s| s.as_str()).collect();
    let out = visi_engine::core::run_macro(source, procedure, &refs).map_err(Wrapped)?;
    Ok((out.type_name, out.value))
}

#[pymodule]
fn visi_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<workbook::Workbook>()?;
    m.add_class::<value::CellError>()?;
    m.add_function(wrap_pyfunction!(eval_file, m)?)?;
    m.add_function(wrap_pyfunction!(check_syntax, m)?)?;
    m.add_function(wrap_pyfunction!(run_macro, m)?)?;
    errors::register(m)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}
