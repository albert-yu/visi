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
//! has to mirror CLI behavior -- `edit_chart`'s clear-vs-set flags, and
//! `add_pivot_field`'s post-add subtotal/label mutation -- that mirroring is a
//! contract enforced by `fuzz/test_backend_parity.py`, not an accident.

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

/// The Python module, `visi_core`.
#[pymodule]
fn visi_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<workbook::Workbook>()?;
    m.add_class::<value::CellError>()?;
    m.add_function(wrap_pyfunction!(eval_file, m)?)?;
    errors::register(m)?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}
