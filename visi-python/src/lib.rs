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

/// Runs a VBA procedure from loose source text and returns
/// `(type_name, value)`.
///
/// Expressions, control flow, `Sub`/`Function` and `On Error`, with **no
/// workbook**: this takes source text, not a file, so `Range`, `Worksheets`
/// and `ThisWorkbook` have nothing to resolve against and report so. Use
/// `Workbook.run_macro` for a run that can touch cells.
///
/// Raises `VbaRuntimeError` (carrying `number`) for a run-time error and
/// `VbaSyntaxError` if the source does not parse.
///
/// `value` is `None` where VBA itself cannot stringify the result, which in
/// practice means `Null`. What `fuzz/fuzz_vba.py` compares against Excel.
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

/// The Python module, `visi_core`.
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
