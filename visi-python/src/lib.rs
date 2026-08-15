//! Python bindings for `visi-core`, built with pyo3 and maturin.
//!
//! Development-only: this exists so the differential fuzz harness in `fuzz/`
//! can drive the engine in-process instead of spawning the `visi` CLI once per
//! operation. It is not published and carries no stability promise.
//!
//! It lives in its own crate because `crate-type` cannot be feature-gated, so
//! a `cdylib` in `visi-core` would make every Rust consumer pay for a shared
//! object they never link (see the commit that dropped `visi-core`'s cdylib).
//!
//! Depends on `visi-core` only, never on the `visi` CLI crate. Where a binding
//! has to mirror CLI behavior -- `edit_chart`'s clear-vs-set flags,
//! `add_pivot_field`'s post-add subtotal/label mutation, and `add_macro`'s
//! sheet-name-to-id resolution with its `ThisWorkbook` exemption -- that
//! mirroring is a contract enforced by `fuzz/test_backend_parity.py`, not an
//! accident.

use pyo3::prelude::*;
use visi_engine::WorkbookManager;

mod enums;
mod errors;
mod value;
mod workbook;

use errors::Wrapped;

/// Reads a workbook, recalculates every formula, and writes the result.
///
/// The whole of what the formula fuzzer needs, and the exact equivalent of
/// `visi eval <input> --output <output>`.
#[pyfunction]
fn eval_file(input: std::path::PathBuf, output: std::path::PathBuf) -> PyResult<()> {
    let bytes = std::fs::read(&input)?;
    let mut wb = WorkbookManager::load_bytes(&bytes).map_err(Wrapped)?;
    wb.evaluate().map_err(Wrapped)?;
    std::fs::write(&output, wb.save_bytes().map_err(Wrapped)?)?;
    Ok(())
}

/// Checks VBA source for syntax errors, returning the procedure names it
/// declares.
///
/// Raises `VbaSyntaxError` (carrying `line` and `column`) if it does not
/// parse. The exact equivalent of `visi macro check` over a `.bas` file, and
/// what `fuzz/fuzz_vba_parse.py` compares against real Excel's verdict.
#[pyfunction]
fn check_syntax(source: &str) -> PyResult<Vec<String>> {
    Ok(visi_engine::core::check_syntax(source)
        .map_err(Wrapped)?
        .procedures)
}

/// The Python module, `visi_core`.
#[pymodule]
fn visi_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<workbook::Workbook>()?;
    m.add_class::<value::CellError>()?;
    m.add_function(wrap_pyfunction!(eval_file, m)?)?;
    m.add_function(wrap_pyfunction!(check_syntax, m)?)?;
    errors::register(m)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}
