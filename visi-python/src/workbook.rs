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
    #[new]
    #[pyo3(signature = (locale=None))]
    fn new(locale: Option<&str>) -> PyResult<Self> {
        let mut inner = WorkbookManager::new_empty().map_err(Wrapped)?;
        if let Some(loc_str) = locale {
            let loc = visi_engine::core::Locale::from_code(loc_str)
                .ok_or_else(|| invalid_argument(format!("unrecognized locale '{loc_str}'")))?;
            inner.set_locale(loc);
        }
        Ok(Self { inner })
    }

    #[getter]
    fn locale(&self) -> String {
        self.inner.locale.code.clone()
    }

    #[setter]
    fn set_locale(&mut self, locale_str: &str) -> PyResult<()> {
        let loc = visi_engine::core::Locale::from_code(locale_str)
            .ok_or_else(|| invalid_argument(format!("unrecognized locale '{locale_str}'")))?;
        self.inner.set_locale(loc);
        Ok(())
    }

    #[staticmethod]
    fn load(path: PathBuf) -> PyResult<Self> {
        let bytes = std::fs::read(&path)?;
        Ok(Self {
            inner: WorkbookManager::load_bytes(&bytes).map_err(Wrapped)?,
        })
    }

    #[staticmethod]
    fn load_bytes(data: &[u8]) -> PyResult<Self> {
        Ok(Self {
            inner: WorkbookManager::load_bytes(data).map_err(Wrapped)?,
        })
    }

    fn save_bytes<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let data = self.inner.save_bytes().map_err(Wrapped)?;
        Ok(PyBytes::new(py, &data))
    }

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

    fn roundtrip(&self) -> PyResult<Self> {
        let data = self.inner.save_bytes().map_err(Wrapped)?;
        Ok(Self {
            inner: WorkbookManager::load_bytes(&data).map_err(Wrapped)?,
        })
    }

    fn evaluate(&mut self) -> PyResult<()> {
        self.inner.evaluate().map_err(Wrapped)?;
        Ok(())
    }

    #[getter]
    fn sheet_names(&self) -> Vec<String> {
        self.inner.sheets.iter().map(|s| s.name.clone()).collect()
    }

    #[pyo3(signature = (name=None))]
    fn sheet_index(&self, name: Option<&str>) -> PyResult<usize> {
        self.sheet_idx(name)
    }

    #[pyo3(signature = (sheet=None))]
    fn dimensions(&self, sheet: Option<&str>) -> PyResult<(usize, usize)> {
        let idx = self.sheet_idx(sheet)?;
        let s = &self.inner.sheets[idx];
        Ok((s.row_count(), s.col_count()))
    }

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

    #[pyo3(signature = (row, col, sheet=None))]
    fn get_display(&self, row: usize, col: usize, sheet: Option<&str>) -> PyResult<String> {
        let idx = self.sheet_idx(sheet)?;
        Ok(self.inner.sheets[idx].get_display_string(&CellRef::new(row, col)))
    }

    #[pyo3(signature = (row, col, sheet=None))]
    fn get_cell_type(&self, row: usize, col: usize, sheet: Option<&str>) -> PyResult<String> {
        let idx = self.sheet_idx(sheet)?;
        let t = self.inner.get_cell_type(idx, row, col);
        Ok(t.as_str().to_string())
    }

    #[pyo3(signature = (row, col, sheet=None))]
    fn get_src(&self, row: usize, col: usize, sheet: Option<&str>) -> PyResult<String> {
        let idx = self.sheet_idx(sheet)?;
        Ok(self.inner.sheets[idx].get_src_str(&CellRef::new(row, col)))
    }

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

    fn delete_chart(&mut self, chart_id: u64) -> PyResult<()> {
        self.inner.delete_chart(chart_id).map_err(Wrapped)?;
        Ok(())
    }

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

    fn remove_pivot_field(&mut self, pivot: &str, area: &str, column: &str) -> PyResult<()> {
        let area = parse_pivot_area(area)?;
        self.inner
            .remove_pivot_field(pivot, area, column)
            .map_err(Wrapped)?;
        Ok(())
    }

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

    fn refresh_pivot(&mut self, pivot: &str) -> PyResult<()> {
        self.inner.refresh_pivot_table(pivot).map_err(Wrapped)?;
        Ok(())
    }

    fn delete_pivot(&mut self, pivot: &str) -> PyResult<()> {
        self.inner.delete_pivot_table(pivot).map_err(Wrapped)?;
        Ok(())
    }

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

    fn has_macros(&self) -> bool {
        self.inner.has_vba_project()
    }

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

    fn remove_macro(&mut self, name: &str) -> PyResult<()> {
        self.inner.remove_vba_module(name).map_err(Wrapped)?;
        Ok(())
    }

    fn rename_macro(&mut self, old: &str, new: &str) -> PyResult<()> {
        self.inner.rename_vba_module(old, new).map_err(Wrapped)?;
        Ok(())
    }

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

    fn run_open_events(&mut self) -> PyResult<(String, Option<String>, bool)> {
        let out = self.inner.run_open_events().map_err(Wrapped)?;
        Ok((out.type_name, out.value, out.mutated))
    }

    fn set_macro_source(&mut self, name: &str, source: &str) -> PyResult<()> {
        self.inner
            .set_vba_module_source(name, source.to_string())
            .map_err(Wrapped)?;
        Ok(())
    }

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
