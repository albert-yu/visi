//! The `Workbook` class: a thin wrapper over [`WorkbookManager`].
//!
//! Deliberately the *only* class exposed. `Sheet` and `Context` stay internal:
//! `Context<'a>` borrows its sheets and so can never be `'static` as pyo3
//! requires, and a `Sheet` handed out on its own would lose the cross-sheet
//! propagation and pivot refresh that only exist at the manager level.
//!
//! All row/column coordinates are 0-based, matching visi-core. A1 notation is
//! not parsed here -- callers pass indices -- so there is no second A1 parser
//! to drift out of sync with `visi/src/utils.rs`.

use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};
use std::path::PathBuf;

use visi_engine::WorkbookManager;
use visi_engine::core::chart::ChartType;
use visi_engine::core::{
    CellRef, PivotArea, VbaModuleKind, col_idx_to_letters, value_field_labels,
};

use crate::enums::{
    parse_cell_type, parse_chart_type, parse_pivot_agg, parse_pivot_area, parse_vba_module_kind,
};
use crate::errors::{Wrapped, invalid_argument};
use crate::value::result_to_py;

/// A workbook: sheets, charts, pivot tables and any VBA project.
///
/// ```python
/// import visi_core
/// wb = visi_core.Workbook.load("book.xlsx")
/// wb.evaluate()
/// wb.save("out.xlsx")
/// ```
#[pyclass(module = "visi_core")]
pub struct Workbook {
    inner: WorkbookManager,
}

impl Workbook {
    fn sheet_idx(&self, sheet: Option<&str>) -> PyResult<usize> {
        Ok(self.inner.find_sheet_index(sheet).map_err(Wrapped)?)
    }
}

#[pymethods]
impl Workbook {
    /// An empty workbook with one sheet.
    #[new]
    fn new() -> PyResult<Self> {
        Ok(Self {
            inner: WorkbookManager::new_empty().map_err(Wrapped)?,
        })
    }

    /// Reads an `.xlsx` from disk.
    #[staticmethod]
    fn load(path: PathBuf) -> PyResult<Self> {
        let bytes = std::fs::read(&path)?;
        Ok(Self {
            inner: WorkbookManager::load_bytes(&bytes).map_err(Wrapped)?,
        })
    }

    /// Reads an `.xlsx` from an in-memory buffer.
    #[staticmethod]
    fn load_bytes(data: &[u8]) -> PyResult<Self> {
        Ok(Self {
            inner: WorkbookManager::load_bytes(data).map_err(Wrapped)?,
        })
    }

    /// Serializes to `.xlsx` bytes.
    fn save_bytes<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let data = self.inner.save_bytes().map_err(Wrapped)?;
        Ok(PyBytes::new(py, &data))
    }

    /// Writes an `.xlsx` to disk, creating parent directories as needed.
    fn save(&self, path: PathBuf) -> PyResult<()> {
        let data = self.inner.save_bytes().map_err(Wrapped)?;
        if let Some(parent) = path.parent()
            && !parent.as_os_str().is_empty()
        {
            std::fs::create_dir_all(parent)?;
        }
        std::fs::write(&path, data)?;
        Ok(())
    }

    /// A fresh `Workbook` round-tripped through the `.xlsx` format.
    ///
    /// Equivalent to `Workbook.load_bytes(wb.save_bytes())`, and to what one
    /// in-place (`-i`) CLI invocation did between two edits. The fuzz harness
    /// uses it to keep exercising `export_xlsx_data` / `import_xlsx_data`
    /// between mutations now that it no longer spawns a process per step.
    ///
    /// Not everything survives: chart ids are re-derived from sheet name and
    /// position, and a pivot filter's selected values reset to "all". See
    /// `add_chart` and `set_pivot_filter`.
    fn roundtrip(&self) -> PyResult<Self> {
        let data = self.inner.save_bytes().map_err(Wrapped)?;
        Ok(Self {
            inner: WorkbookManager::load_bytes(&data).map_err(Wrapped)?,
        })
    }

    /// Recalculates every formula in every sheet.
    ///
    /// Almost never raises: visi-core discards per-sheet commit errors, so a
    /// formula that fails shows up as a `CellError` *value* in the cell rather
    /// than as an exception here. Check cells, not exceptions.
    fn evaluate(&mut self) -> PyResult<()> {
        self.inner.evaluate().map_err(Wrapped)?;
        Ok(())
    }

    /// The worksheet names, in workbook order.
    #[getter]
    fn sheet_names(&self) -> Vec<String> {
        self.inner.sheets.iter().map(|s| s.name.clone()).collect()
    }

    /// The index of a sheet by name, or of the first sheet when `name` is
    /// `None`.
    #[pyo3(signature = (name=None))]
    fn sheet_index(&self, name: Option<&str>) -> PyResult<usize> {
        self.sheet_idx(name)
    }

    /// `(rows, cols)` for a sheet.
    #[pyo3(signature = (sheet=None))]
    fn dimensions(&self, sheet: Option<&str>) -> PyResult<(usize, usize)> {
        let idx = self.sheet_idx(sheet)?;
        let s = &self.inner.sheets[idx];
        Ok((s.row_count(), s.col_count()))
    }

    /// Writes a cell's source text -- a literal (`"10"`) or a formula
    /// (`"=SUM(A1:A2)"`). Call `evaluate()` afterwards to recompute.
    #[pyo3(signature = (row, col, value, sheet=None, cell_type=None))]
    fn set_cell(
        &mut self,
        row: usize,
        col: usize,
        value: String,
        sheet: Option<&str>,
        cell_type: Option<&str>,
    ) -> PyResult<()> {
        let idx = self.sheet_idx(sheet)?;
        self.inner.ensure_capacity(idx, row, col);
        if let Some(ct_str) = cell_type {
            let ct = parse_cell_type(ct_str)?;
            self.inner.set_cell_with_type(idx, row, col, value, ct);
        } else {
            self.inner.set_cell(idx, row, col, value);
        }
        Ok(())
    }

    /// Sets the intrinsic data type of a cell at (row, col).
    #[pyo3(signature = (row, col, cell_type, sheet=None))]
    fn set_cell_type(
        &mut self,
        row: usize,
        col: usize,
        cell_type: &str,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        let idx = self.sheet_idx(sheet)?;
        let ct = parse_cell_type(cell_type)?;
        self.inner.set_cell_type(idx, row, col, ct);
        Ok(())
    }

    /// A cell's computed value.
    ///
    /// Excel error values come back as `CellError`, not as `str`, so a cell
    /// holding the *text* `#DIV/0!` stays distinguishable from one that
    /// evaluated to that error. A date is a plain number here -- use
    /// `get_display` for the rendered form.
    #[pyo3(signature = (row, col, sheet=None))]
    fn get_cell<'py>(
        &self,
        py: Python<'py>,
        row: usize,
        col: usize,
        sheet: Option<&str>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let idx = self.sheet_idx(sheet)?;
        let v = self.inner.sheets[idx].get_result_data(&CellRef::new(row, col));
        result_to_py(py, &v)
    }

    /// A cell's value rendered the way it would be shown, honoring the cell's
    /// number format. This is the only correct way to render a date.
    #[pyo3(signature = (row, col, sheet=None))]
    fn get_display(&self, row: usize, col: usize, sheet: Option<&str>) -> PyResult<String> {
        let idx = self.sheet_idx(sheet)?;
        Ok(self.inner.sheets[idx].get_display_string(&CellRef::new(row, col)))
    }

    /// The cell's intrinsic type ("auto", "empty", "number", "string", "boolean", "error", "formula").
    #[pyo3(signature = (row, col, sheet=None))]
    fn get_cell_type(&self, row: usize, col: usize, sheet: Option<&str>) -> PyResult<String> {
        let idx = self.sheet_idx(sheet)?;
        let t = self.inner.get_cell_type(idx, row, col);
        let s = match t {
            visi_engine::core::CellType::Auto => "auto",
            visi_engine::core::CellType::Empty => "empty",
            visi_engine::core::CellType::Number => "number",
            visi_engine::core::CellType::String => "string",
            visi_engine::core::CellType::Boolean => "boolean",
            visi_engine::core::CellType::Error => "error",
            visi_engine::core::CellType::Formula => "formula",
        };
        Ok(s.to_string())
    }

    /// A cell's source text, as typed.
    #[pyo3(signature = (row, col, sheet=None))]
    fn get_src(&self, row: usize, col: usize, sheet: Option<&str>) -> PyResult<String> {
        let idx = self.sheet_idx(sheet)?;
        Ok(self.inner.sheets[idx].get_src_str(&CellRef::new(row, col)))
    }

    // ---- charts ---------------------------------------------------------

    /// Adds a chart over `range` (A1 with a sheet prefix, e.g. `"Sheet1!A1:B10"`).
    ///
    /// Returns the new chart's id. **That id is valid only for this in-memory
    /// workbook.** `import_xlsx_data` re-derives chart ids from the sheet name
    /// and the chart's position within that sheet, so after `save`/`load` or
    /// `roundtrip()` the id will differ -- re-read it from `charts()`.
    #[pyo3(signature = (sheet, chart_type, range, title=None, anchor=None))]
    fn add_chart(
        &mut self,
        sheet: &str,
        chart_type: &str,
        range: String,
        title: Option<String>,
        anchor: Option<(usize, usize)>,
    ) -> PyResult<u64> {
        let ct = parse_chart_type(chart_type)?;
        Ok(self
            .inner
            .add_chart(sheet, ct, range, title, anchor)
            .map_err(Wrapped)?)
    }

    /// Edits a chart.
    ///
    /// Each of title/xlabel/ylabel is three-state -- leave alone, set, or
    /// clear -- which a single argument cannot express, since pyo3 maps Python
    /// `None` onto "not supplied". So each has a paired `clear_*` flag,
    /// mirroring the CLI's `--title` / `--clear-title`. Passing both a value
    /// and its `clear_*` flag is an error, as it is on the CLI.
    #[allow(clippy::too_many_arguments)]
    #[pyo3(signature = (
        chart_id, *, name=None, chart_type=None, range=None,
        title=None, clear_title=false,
        xlabel=None, clear_xlabel=false,
        ylabel=None, clear_ylabel=false,
        show_legend=None, anchor=None,
    ))]
    fn edit_chart(
        &mut self,
        chart_id: u64,
        name: Option<String>,
        chart_type: Option<&str>,
        range: Option<String>,
        title: Option<String>,
        clear_title: bool,
        xlabel: Option<String>,
        clear_xlabel: bool,
        ylabel: Option<String>,
        clear_ylabel: bool,
        show_legend: Option<bool>,
        anchor: Option<(usize, usize)>,
    ) -> PyResult<()> {
        fn tri(value: Option<String>, clear: bool, what: &str) -> PyResult<Option<Option<String>>> {
            match (value, clear) {
                (Some(_), true) => Err(invalid_argument(format!(
                    "pass either {what} or clear_{what}, not both"
                ))),
                (Some(v), false) => Ok(Some(Some(v))),
                (None, true) => Ok(Some(None)),
                (None, false) => Ok(None),
            }
        }

        let title = tri(title, clear_title, "title")?;
        let xlabel = tri(xlabel, clear_xlabel, "xlabel")?;
        let ylabel = tri(ylabel, clear_ylabel, "ylabel")?;
        let ct: Option<ChartType> = chart_type.map(parse_chart_type).transpose()?;

        self.inner
            .edit_chart(
                chart_id,
                name,
                ct,
                range,
                title,
                xlabel,
                ylabel,
                show_legend,
                anchor,
            )
            .map_err(Wrapped)?;
        Ok(())
    }

    /// Deletes a chart by id.
    fn delete_chart(&mut self, chart_id: u64) -> PyResult<()> {
        self.inner.delete_chart(chart_id).map_err(Wrapped)?;
        Ok(())
    }

    /// The charts, as dicts.
    ///
    /// Carries every key `visi chart list --json` emits (`id`, `name`, `type`,
    /// `data_range`, `title`, `anchor`) so the two are directly comparable,
    /// plus `xlabel`, `ylabel` and `show_legend`, which that command does not
    /// report.
    fn charts<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let mut out = Vec::with_capacity(self.inner.charts.len());
        for c in &self.inner.charts {
            let d = PyDict::new(py);
            d.set_item("id", c.id)?;
            d.set_item("name", &c.name)?;
            d.set_item("type", format!("{:?}", c.chart_type))?;
            d.set_item("data_range", &c.data_range)?;
            d.set_item("title", c.title.clone())?;
            d.set_item(
                "anchor",
                format!("{}{}", col_idx_to_letters(c.anchor_col), c.anchor_row + 1),
            )?;
            d.set_item("xlabel", c.xlabel.clone())?;
            d.set_item("ylabel", c.ylabel.clone())?;
            d.set_item("show_legend", c.show_legend)?;
            out.push(d);
        }
        PyList::new(py, out)
    }

    // ---- pivot tables ---------------------------------------------------

    /// Creates a pivot table sourced from a named Excel Table.
    #[allow(clippy::too_many_arguments)]
    #[pyo3(signature = (
        name, source_table, *, dest_sheet=None, dest_row=0, dest_col=0,
        grand_totals_row=true, grand_totals_col=true,
    ))]
    fn add_pivot_from_table(
        &mut self,
        name: &str,
        source_table: &str,
        dest_sheet: Option<&str>,
        dest_row: usize,
        dest_col: usize,
        grand_totals_row: bool,
        grand_totals_col: bool,
    ) -> PyResult<u64> {
        Ok(self
            .inner
            .add_pivot_table_from_table(
                name,
                source_table,
                dest_sheet,
                dest_row,
                dest_col,
                grand_totals_row,
                grand_totals_col,
            )
            .map_err(Wrapped)?)
    }

    /// Creates a pivot table sourced from a raw sheet range, given as 0-based
    /// inclusive bounds.
    #[allow(clippy::too_many_arguments)]
    #[pyo3(signature = (
        name, *, source_sheet=None, start_row, start_col, end_row, end_col,
        dest_sheet=None, dest_row=0, dest_col=0,
        grand_totals_row=true, grand_totals_col=true,
    ))]
    fn add_pivot_from_range(
        &mut self,
        name: &str,
        source_sheet: Option<&str>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        dest_sheet: Option<&str>,
        dest_row: usize,
        dest_col: usize,
        grand_totals_row: bool,
        grand_totals_col: bool,
    ) -> PyResult<u64> {
        Ok(self
            .inner
            .add_pivot_table_from_range(
                name,
                source_sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                dest_sheet,
                dest_row,
                dest_col,
                grand_totals_row,
                grand_totals_col,
            )
            .map_err(Wrapped)?)
    }

    /// Adds a field to a pivot area (`"row"`, `"column"`, `"value"`,
    /// `"filter"`), then refreshes if anything changed.
    ///
    /// `subtotal=False` and `label` are applied by mutating the field after
    /// the add, because `WorkbookManager` has no "add a field with subtotals
    /// off" entry point. This mirrors `handle_pivot`'s AddField arm exactly;
    /// the two must stay in step, which is what `fuzz/test_backend_parity.py`
    /// checks.
    #[pyo3(signature = (pivot, area, column, *, agg=None, subtotal=true, label=None))]
    fn add_pivot_field(
        &mut self,
        pivot: &str,
        area: &str,
        column: &str,
        agg: Option<&str>,
        subtotal: bool,
        label: Option<String>,
    ) -> PyResult<()> {
        let area = parse_pivot_area(area)?;
        let agg = agg.map(parse_pivot_agg).transpose()?;
        if matches!(area, PivotArea::Value) && agg.is_none() {
            return Err(invalid_argument(
                "a value field needs an aggregation; pass agg=\"sum\" (or count, count-numbers, average, max, min)",
            ));
        }

        self.inner
            .add_pivot_field(pivot, area, column, agg)
            .map_err(Wrapped)?;

        let idx = self
            .inner
            .pivot_tables
            .iter()
            .position(|p| p.name.eq_ignore_ascii_case(pivot));
        let mut needs_refresh = false;
        if let Some(i) = idx {
            if !subtotal {
                let pt = &mut self.inner.pivot_tables[i];
                if let Some(pf) = pt
                    .row_fields
                    .iter_mut()
                    .chain(pt.col_fields.iter_mut())
                    .rev()
                    .find(|f| f.column.eq_ignore_ascii_case(column))
                {
                    pf.subtotal = false;
                    needs_refresh = true;
                }
            }
            if let Some(l) = label
                && let Some(vf) = self.inner.pivot_tables[i]
                    .value_fields
                    .iter_mut()
                    .rev()
                    .find(|f| f.column.eq_ignore_ascii_case(column))
            {
                vf.custom_name = Some(l);
                needs_refresh = true;
            }
        }
        if needs_refresh {
            self.inner.refresh_pivot_table(pivot).map_err(Wrapped)?;
        }
        Ok(())
    }

    /// Removes a field from a pivot area.
    fn remove_pivot_field(&mut self, pivot: &str, area: &str, column: &str) -> PyResult<()> {
        let area = parse_pivot_area(area)?;
        self.inner
            .remove_pivot_field(pivot, area, column)
            .map_err(Wrapped)?;
        Ok(())
    }

    /// Sets which values of a filter field take part.
    ///
    /// `None` clears the filter (every value allowed); a list restricts to
    /// those values, and an **empty list selects nothing** -- a state the
    /// `visi pivot filter` command cannot express, since it requires either a
    /// value list or `--clear`.
    ///
    /// A selection survives `roundtrip()` -- it is written as indices into
    /// the pivot cache's shared items and resolved back to values on import.
    /// Two cases still cannot: selecting *every* value marks nothing hidden
    /// and so reads back as no filter (the grid is identical either way), and
    /// a filter on a column that is also a row or column field has nowhere to
    /// be recorded, since a pivot field carries one orientation.
    #[pyo3(signature = (pivot, column, values))]
    fn set_pivot_filter(
        &mut self,
        pivot: &str,
        column: &str,
        values: Option<Vec<String>>,
    ) -> PyResult<()> {
        self.inner
            .set_pivot_filter(pivot, column, values)
            .map_err(Wrapped)?;
        Ok(())
    }

    /// Recomputes a pivot table's grid and writes it into its destination
    /// cells. Nothing does this implicitly.
    fn refresh_pivot(&mut self, pivot: &str) -> PyResult<()> {
        self.inner.refresh_pivot_table(pivot).map_err(Wrapped)?;
        Ok(())
    }

    /// Deletes a pivot table by name.
    fn delete_pivot(&mut self, pivot: &str) -> PyResult<()> {
        self.inner.delete_pivot_table(pivot).map_err(Wrapped)?;
        Ok(())
    }

    /// The pivot tables, as dicts.
    ///
    /// Carries every key `visi pivot list --json` emits (`id`, `name`,
    /// `row_fields`, `col_fields`, `value_fields`, `filter_fields`), plus
    /// `subtotals` (per row/column field) and `filter_selections`, which that
    /// command does not report but which are exactly the state a round trip
    /// can lose.
    fn pivots<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let mut out = Vec::with_capacity(self.inner.pivot_tables.len());
        for p in self.inner.list_pivot_tables() {
            let d = PyDict::new(py);
            d.set_item("id", p.id)?;
            d.set_item("name", &p.name)?;
            d.set_item(
                "row_fields",
                p.row_fields.iter().map(|f| &f.column).collect::<Vec<_>>(),
            )?;
            d.set_item(
                "col_fields",
                p.col_fields.iter().map(|f| &f.column).collect::<Vec<_>>(),
            )?;
            d.set_item("value_fields", value_field_labels(&p.value_fields))?;
            d.set_item(
                "filter_fields",
                p.filter_fields
                    .iter()
                    .map(|f| &f.column)
                    .collect::<Vec<_>>(),
            )?;

            let subtotals = PyDict::new(py);
            for f in p.row_fields.iter().chain(p.col_fields.iter()) {
                subtotals.set_item(&f.column, f.subtotal)?;
            }
            d.set_item("subtotals", subtotals)?;

            let selections = PyDict::new(py);
            for f in &p.filter_fields {
                selections.set_item(&f.column, f.selected_values.clone())?;
            }
            d.set_item("filter_selections", selections)?;

            out.push(d);
        }
        PyList::new(py, out)
    }

    // ---- VBA macro modules ----------------------------------------------

    /// Whether the workbook carries a VBA project at all.
    fn has_macros(&self) -> bool {
        self.inner.has_vba_project()
    }

    /// Adds a VBA module, mirroring `visi macro add`.
    ///
    /// `kind` is `"standard"`, `"class"` or `"document"`. `sheet` names the
    /// sheet a document module binds to and is required for `"document"` --
    /// except for `ThisWorkbook`, which isn't tied to a specific sheet. This
    /// resolve-sheet-name-to-id-then-special-case-ThisWorkbook step is the
    /// CLI behaviour (`visi/src/main.rs`'s Add arm) duplicated here, in the
    /// same way `edit_chart` and `add_pivot_field` are; `visi-core` takes the
    /// id, not the name.
    ///
    /// `source` is written verbatim -- callers include their own
    /// `Attribute VB_Name = "..."` line, matching how real Excel-authored
    /// module streams are shaped, and nothing reconciles it against `name`.
    ///
    /// Excel only loads macros from a `.xlsm`, so save the result with that
    /// extension. Unlike the CLI, which refuses any other extension outright,
    /// this does not police the filename -- `save_bytes` has no filename to
    /// police, and a fuzz harness writing to a temp path shouldn't have to
    /// care.
    #[pyo3(signature = (name, source, *, kind="standard", sheet=None))]
    fn add_macro(
        &mut self,
        name: &str,
        source: &str,
        kind: &str,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        let kind = parse_vba_module_kind(kind)?;
        let bound_sheet_id = match (kind, sheet) {
            (VbaModuleKind::Document, Some(sheet_name)) => {
                let idx = self.sheet_idx(Some(sheet_name))?;
                Some(self.inner.sheets[idx].id)
            }
            (VbaModuleKind::Document, None) if name == "ThisWorkbook" => None,
            (VbaModuleKind::Document, None) => {
                return Err(invalid_argument(
                    "kind='document' requires sheet=... (except for 'ThisWorkbook')",
                ));
            }
            _ => None,
        };
        self.inner
            .add_vba_module(name.to_string(), kind, source.to_string(), bound_sheet_id)
            .map_err(Wrapped)?;
        Ok(())
    }

    /// Removes a VBA module by name. Mirrors `visi macro remove`.
    fn remove_macro(&mut self, name: &str) -> PyResult<()> {
        self.inner.remove_vba_module(name).map_err(Wrapped)?;
        Ok(())
    }

    /// Renames a VBA module. Mirrors `visi macro rename`.
    fn rename_macro(&mut self, old: &str, new: &str) -> PyResult<()> {
        self.inner.rename_vba_module(old, new).map_err(Wrapped)?;
        Ok(())
    }

    /// Runs one of this workbook's macros **against** this workbook, mirroring
    /// `visi macro run FILE`.
    ///
    /// Returns `(type_name, value, mutated)`. Unlike the module-level
    /// `visi_core.run_macro`, which takes loose source text and has no
    /// workbook to touch, this gives the macro the Phase 2 host object model:
    /// it can read and write cells, walk the sheets, and call
    /// `Application.WorksheetFunction`. Anything it changes is in this
    /// `Workbook` afterwards, and `save`/`save_bytes` is what persists it --
    /// there is no implicit write, exactly as in the CLI.
    ///
    /// `module` picks which module to take the procedure from; omitted, every
    /// module is searched for one declaring it. That resolution lives in
    /// `visi-core` rather than here precisely so this and the CLI cannot
    /// drift apart -- `edit_chart` and `add_pivot_field` show what the other
    /// choice costs.
    ///
    /// **This executes code the workbook's author wrote.** Nothing else in
    /// these bindings does: not `load`, not `evaluate`, not `roundtrip`.
    #[pyo3(signature = (procedure, *, module=None, args=None))]
    fn run_macro(
        &mut self,
        procedure: &str,
        module: Option<&str>,
        args: Option<Vec<String>>,
    ) -> PyResult<(String, Option<String>, bool)> {
        let args = args.unwrap_or_default();
        let refs: Vec<&str> = args.iter().map(|s| s.as_str()).collect();
        let out = self
            .inner
            .run_macro(module, procedure, &refs)
            .map_err(Wrapped)?;
        Ok((out.type_name, out.value, out.mutated))
    }

    /// Runs startup macro events (`Workbook_Open` in `ThisWorkbook` then `Auto_Open` in standard modules).
    fn run_open_events(&mut self) -> PyResult<(String, Option<String>, bool)> {
        let out = self.inner.run_open_events().map_err(Wrapped)?;
        Ok((out.type_name, out.value, out.mutated))
    }

    /// Replaces a module's source text. Mirrors `visi macro set-source`.
    fn set_macro_source(&mut self, name: &str, source: &str) -> PyResult<()> {
        self.inner
            .set_vba_module_source(name, source.to_string())
            .map_err(Wrapped)?;
        Ok(())
    }

    /// The VBA modules, as dicts.
    ///
    /// Carries the keys `visi macro list --json` emits (`name`, `kind`,
    /// `source_lines`), plus `source` itself and `bound_sheet_id` -- the two
    /// things that command does not report but that a round trip can lose.
    fn macros<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let modules = self.inner.list_vba_modules();
        let mut out = Vec::with_capacity(modules.len());
        for m in modules {
            let d = PyDict::new(py);
            d.set_item("name", &m.name)?;
            d.set_item("kind", format!("{:?}", m.kind))?;
            d.set_item("source", &m.source)?;
            d.set_item("source_lines", m.source.lines().count())?;
            d.set_item("bound_sheet_id", m.bound_sheet_id)?;
            out.push(d);
        }
        PyList::new(py, out)
    }
}
