use std::collections::HashMap;

use js_sys::{Array, Object, Reflect};
use serde::{Deserialize, Serialize};
use visi_engine::core::chart::ChartType;
use visi_engine::core::typecheck::{
    FormulaDiagnosticKind, FormulaDiagnosticSeverity, check_formula,
};
use visi_engine::core::{
    CellRef, CellStyle, Context, PivotArea, VbaModuleKind, col_idx_to_letters, parse_cell_ref,
    parse_range_ref, value_field_labels,
};
use visi_engine::{WorkbookFormat, WorkbookManager};
use wasm_bindgen::prelude::*;

use crate::enums::{
    parse_cell_type, parse_chart_type, parse_pivot_agg, parse_pivot_area, parse_vba_module_kind,
    parse_workbook_format,
};
use crate::errors::{WasmResult, Wrapped, invalid_argument};
use crate::value::result_to_js;

pub fn to_js_value<T: Serialize>(val: &T) -> WasmResult<JsValue> {
    val.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
        .map_err(|e| invalid_argument(e.to_string()))
}

fn from_js_opt<T: for<'de> Deserialize<'de> + Default>(val: Option<JsValue>) -> WasmResult<T> {
    match val {
        Some(v) if !v.is_undefined() && !v.is_null() => {
            serde_wasm_bindgen::from_value(v).map_err(|e| invalid_argument(e.to_string()))
        }
        _ => Ok(T::default()),
    }
}

fn from_js_req<T: for<'de> Deserialize<'de>>(val: JsValue) -> WasmResult<T> {
    serde_wasm_bindgen::from_value(val).map_err(|e| invalid_argument(e.to_string()))
}

fn diagnostic_severity(severity: &FormulaDiagnosticSeverity) -> &'static str {
    match severity {
        FormulaDiagnosticSeverity::Error => "error",
        FormulaDiagnosticSeverity::Warning => "warning",
    }
}

fn diagnostic_kind(kind: &FormulaDiagnosticKind) -> &'static str {
    match kind {
        FormulaDiagnosticKind::Syntax => "syntax",
        FormulaDiagnosticKind::Reference => "reference",
        FormulaDiagnosticKind::Function => "function",
        FormulaDiagnosticKind::Type => "type",
    }
}

#[derive(Serialize, Deserialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct JsCellStyle {
    #[serde(alias = "font_color", skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(alias = "bg_color", skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(alias = "font_family", skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(alias = "font_size", skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(alias = "num_format", skip_serializing_if = "Option::is_none")]
    pub num_format: Option<String>,
}

impl From<JsCellStyle> for CellStyle {
    fn from(s: JsCellStyle) -> Self {
        Self {
            font_color: s.font_color,
            bg_color: s.bg_color,
            bold: s.bold,
            italic: s.italic,
            underline: s.underline,
            font_family: s.font_family,
            font_size: s.font_size,
            num_format: s.num_format,
            ..Default::default()
        }
    }
}

impl From<CellStyle> for JsCellStyle {
    fn from(s: CellStyle) -> Self {
        Self {
            font_color: s.font_color,
            bg_color: s.bg_color,
            bold: s.bold,
            italic: s.italic,
            underline: s.underline,
            font_family: s.font_family,
            font_size: s.font_size,
            num_format: s.num_format,
        }
    }
}

#[derive(Deserialize)]
#[serde(untagged)]
enum CellSelector {
    Coords((usize, usize)),
    A1(String),
}

#[derive(Deserialize)]
#[serde(untagged)]
enum RangeSelector {
    Bounds((usize, usize, usize, usize)),
    A1(String),
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct TypecheckOptions {
    sheet: Option<String>,
    cell: Option<CellSelector>,
    range: Option<RangeSelector>,
}

#[derive(Serialize)]
struct DiagnosticEntry {
    sheet: String,
    cell: String,
    row: usize,
    col: usize,
    severity: &'static str,
    kind: &'static str,
    message: String,
    formula: String,
}

#[derive(Serialize)]
struct TypecheckReport {
    error_count: usize,
    warning_count: usize,
    #[serde(rename = "errorCount")]
    error_count_camel: usize,
    #[serde(rename = "warningCount")]
    warning_count_camel: usize,
    diagnostics: Vec<DiagnosticEntry>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct JsSheetSummary {
    name: String,
    row_count: usize,
    col_count: usize,
    formula_count: usize,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct JsWorkbookSummary {
    file_name: String,
    sheet_count: usize,
    chart_count: usize,
    sheets: Vec<JsSheetSummary>,
}

#[derive(Serialize)]
struct JsDimensions {
    rows: usize,
    cols: usize,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsCellUpdate {
    row: usize,
    col: usize,
    value: String,
    sheet: Option<String>,
    #[serde(alias = "cell_type")]
    cell_type: Option<String>,
    style: Option<JsCellStyle>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsAddTableOptions {
    name: String,
    sheet: Option<String>,
    #[serde(alias = "start_row")]
    start_row: usize,
    #[serde(alias = "start_col")]
    start_col: usize,
    #[serde(alias = "end_row")]
    end_row: usize,
    #[serde(alias = "end_col")]
    end_col: usize,
    #[serde(alias = "has_header_row", default = "default_true")]
    has_header_row: bool,
    #[serde(alias = "has_totals_row", default)]
    has_totals_row: bool,
    #[serde(alias = "style_name")]
    style_name: Option<String>,
}

fn default_true() -> bool {
    true
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct JsTableInfo {
    id: u64,
    name: String,
    sheet: String,
    sheet_id: u64,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    has_header_row: bool,
    has_totals_row: bool,
    columns: Vec<String>,
    style_name: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsPivotDestOptions {
    #[serde(alias = "dest_sheet")]
    dest_sheet: Option<String>,
    #[serde(alias = "dest_row", default)]
    dest_row: usize,
    #[serde(alias = "dest_col", default)]
    dest_col: usize,
    #[serde(alias = "grand_totals_row", default = "default_true")]
    grand_totals_row: bool,
    #[serde(alias = "grand_totals_col", default = "default_true")]
    grand_totals_col: bool,
}

impl Default for JsPivotDestOptions {
    fn default() -> Self {
        Self {
            dest_sheet: None,
            dest_row: 0,
            dest_col: 0,
            grand_totals_row: true,
            grand_totals_col: true,
        }
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsPivotRangeOptions {
    #[serde(alias = "source_sheet")]
    source_sheet: Option<String>,
    #[serde(alias = "start_row")]
    start_row: usize,
    #[serde(alias = "start_col")]
    start_col: usize,
    #[serde(alias = "end_row")]
    end_row: usize,
    #[serde(alias = "end_col")]
    end_col: usize,
    #[serde(alias = "dest_sheet")]
    dest_sheet: Option<String>,
    #[serde(alias = "dest_row", default)]
    dest_row: usize,
    #[serde(alias = "dest_col", default)]
    dest_col: usize,
    #[serde(alias = "grand_totals_row", default = "default_true")]
    grand_totals_row: bool,
    #[serde(alias = "grand_totals_col", default = "default_true")]
    grand_totals_col: bool,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsPivotFieldOptions {
    agg: Option<String>,
    #[serde(default = "default_true")]
    subtotal: bool,
    label: Option<String>,
}

impl Default for JsPivotFieldOptions {
    fn default() -> Self {
        Self {
            agg: None,
            subtotal: true,
            label: None,
        }
    }
}

#[derive(Serialize)]
struct JsPivotInfo {
    id: u64,
    name: String,
    row_fields: Vec<String>,
    col_fields: Vec<String>,
    value_fields: Vec<String>,
    filter_fields: Vec<String>,
    #[serde(rename = "rowFields")]
    row_fields_camel: Vec<String>,
    #[serde(rename = "colFields")]
    col_fields_camel: Vec<String>,
    #[serde(rename = "valueFields")]
    value_fields_camel: Vec<String>,
    #[serde(rename = "filterFields")]
    filter_fields_camel: Vec<String>,
    subtotals: HashMap<String, bool>,
    filter_selections: HashMap<String, Option<Vec<String>>>,
    #[serde(rename = "filterSelections")]
    filter_selections_camel: HashMap<String, Option<Vec<String>>>,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct JsAddChartOptions {
    title: Option<String>,
    anchor: Option<(usize, usize)>,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct JsEditChartOptions {
    name: Option<String>,
    #[serde(alias = "chart_type")]
    chart_type: Option<String>,
    range: Option<String>,
    title: Option<String>,
    #[serde(alias = "clear_title", default)]
    clear_title: bool,
    xlabel: Option<String>,
    #[serde(alias = "clear_xlabel", default)]
    clear_xlabel: bool,
    ylabel: Option<String>,
    #[serde(alias = "clear_ylabel", default)]
    clear_ylabel: bool,
    #[serde(alias = "show_legend")]
    show_legend: Option<bool>,
    anchor: Option<(usize, usize)>,
}

#[derive(Serialize)]
struct JsChartInfo {
    id: u64,
    name: String,
    #[serde(rename = "type")]
    chart_type: String,
    data_range: String,
    #[serde(rename = "dataRange")]
    data_range_camel: String,
    title: Option<String>,
    anchor: String,
    xlabel: Option<String>,
    ylabel: Option<String>,
    show_legend: bool,
    #[serde(rename = "showLegend")]
    show_legend_camel: bool,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct JsAddMacroOptions {
    kind: Option<String>,
    sheet: Option<String>,
}

#[derive(Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct JsRunMacroOptions {
    module: Option<String>,
    args: Option<Vec<String>>,
}

#[derive(Serialize)]
struct JsVbaModuleInfo {
    name: String,
    kind: String,
    source: String,
    source_lines: usize,
    #[serde(rename = "sourceLines")]
    source_lines_camel: usize,
    bound_sheet_id: Option<u64>,
    #[serde(rename = "boundSheetId")]
    bound_sheet_id_camel: Option<u64>,
}

#[derive(Serialize)]
pub struct JsMacroOutcome {
    #[serde(rename = "type")]
    pub type_short: String,
    pub type_name: String,
    #[serde(rename = "typeName")]
    pub type_name_camel: String,
    pub value: Option<String>,
    pub mutated: bool,
}

#[wasm_bindgen]
pub struct Workbook {
    inner: WorkbookManager,
}

impl Workbook {
    fn sheet_idx(&self, sheet: Option<&str>) -> WasmResult<usize> {
        Ok(self.inner.find_sheet_index(sheet).map_err(Wrapped)?)
    }
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new(locale: Option<String>) -> WasmResult<Workbook> {
        let mut inner = WorkbookManager::new_empty().map_err(Wrapped)?;
        if let Some(loc_str) = locale {
            let loc = visi_engine::core::Locale::from_code(&loc_str)
                .ok_or_else(|| invalid_argument(format!("unrecognized locale '{loc_str}'")))?;
            inner.set_locale(loc);
        }
        Ok(Self { inner })
    }

    #[wasm_bindgen(js_name = fromBytes)]
    pub fn from_bytes(data: &[u8], locale: Option<String>) -> WasmResult<Workbook> {
        let mut inner = WorkbookManager::load_bytes(data).map_err(Wrapped)?;
        if let Some(loc_str) = locale {
            let loc = visi_engine::core::Locale::from_code(&loc_str)
                .ok_or_else(|| invalid_argument(format!("unrecognized locale '{loc_str}'")))?;
            inner.set_locale(loc);
        }
        Ok(Self { inner })
    }

    #[wasm_bindgen(js_name = loadBytes)]
    pub fn load_bytes(data: &[u8]) -> WasmResult<Workbook> {
        Self::from_bytes(data, None)
    }

    #[wasm_bindgen(js_name = fromCsvBytes)]
    pub fn from_csv_bytes(
        data: &[u8],
        sheet_name: Option<String>,
        locale: Option<String>,
    ) -> WasmResult<Workbook> {
        let name = sheet_name.as_deref().unwrap_or("Sheet1");
        let mut inner = WorkbookManager::from_csv_bytes(data, name).map_err(Wrapped)?;
        if let Some(loc_str) = locale {
            let loc = visi_engine::core::Locale::from_code(&loc_str)
                .ok_or_else(|| invalid_argument(format!("unrecognized locale '{loc_str}'")))?;
            inner.set_locale(loc);
        }
        Ok(Self { inner })
    }

    #[wasm_bindgen(js_name = importCsvBytes)]
    pub fn import_csv_bytes(
        &mut self,
        data: &[u8],
        sheet_name: &str,
        replace_existing: Option<bool>,
    ) -> WasmResult<usize> {
        Ok(self
            .inner
            .import_csv_bytes(data, sheet_name, replace_existing.unwrap_or(false))
            .map_err(Wrapped)?)
    }

    #[wasm_bindgen(js_name = toBytes)]
    pub fn to_bytes(&self, format: Option<String>) -> WasmResult<Vec<u8>> {
        let fmt = match format.as_deref() {
            Some(s) => parse_workbook_format(s).map_err(Wrapped)?,
            None => WorkbookFormat::Xlsx,
        };
        Ok(self.inner.save_bytes_as(fmt).map_err(Wrapped)?)
    }

    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&self, format: Option<String>) -> WasmResult<Vec<u8>> {
        self.to_bytes(format)
    }

    pub fn roundtrip(&self) -> WasmResult<Workbook> {
        let data = self.inner.save_bytes().map_err(Wrapped)?;
        Ok(Self {
            inner: WorkbookManager::load_bytes(&data).map_err(Wrapped)?,
        })
    }

    pub fn summary(&self, file_name: Option<String>) -> WasmResult<JsValue> {
        let s = self
            .inner
            .get_summary(file_name.as_deref().unwrap_or("workbook.xlsx"));
        let out = JsWorkbookSummary {
            file_name: s.file_name,
            sheet_count: s.sheet_count,
            chart_count: s.chart_count,
            sheets: s
                .sheets
                .into_iter()
                .map(|sh| JsSheetSummary {
                    name: sh.name,
                    row_count: sh.row_count,
                    col_count: sh.col_count,
                    formula_count: sh.formula_count,
                })
                .collect(),
        };
        to_js_value(&out)
    }

    #[wasm_bindgen(getter)]
    pub fn locale(&self) -> String {
        self.inner.locale.code.clone()
    }

    #[wasm_bindgen(setter)]
    pub fn set_locale(&mut self, locale_str: &str) -> WasmResult<()> {
        let loc = visi_engine::core::Locale::from_code(locale_str)
            .ok_or_else(|| invalid_argument(format!("unrecognized locale '{locale_str}'")))?;
        self.inner.set_locale(loc);
        Ok(())
    }

    pub fn evaluate(&mut self) -> WasmResult<()> {
        self.inner.evaluate().map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = detectCircularReferences)]
    pub fn detect_circular_references(&self) -> WasmResult<JsValue> {
        let cycles = self.inner.detect_circular_references();
        to_js_value(&cycles)
    }

    pub fn typecheck(&self, options: Option<JsValue>) -> WasmResult<JsValue> {
        let opts: TypecheckOptions = from_js_opt(options)?;
        if opts.cell.is_some() && opts.range.is_some() {
            return Err(invalid_argument("pass either cell or range, not both"));
        }

        let mut sheet_override = opts.sheet;
        let resolved_cell = match opts.cell {
            Some(CellSelector::Coords(coords)) => Some(coords),
            Some(CellSelector::A1(s)) => {
                let (sh, r, c) = parse_cell_ref(&s).map_err(invalid_argument)?;
                if sh.is_some() {
                    sheet_override = sh;
                }
                Some((r, c))
            }
            None => None,
        };

        let resolved_range = match opts.range {
            Some(RangeSelector::Bounds(bounds)) => {
                let (sr, sc, er, ec) = bounds;
                if sr > er || sc > ec {
                    return Err(invalid_argument("range end must not precede its start"));
                }
                Some(bounds)
            }
            Some(RangeSelector::A1(s)) => {
                let (sh, sr, sc, er, ec) = parse_range_ref(&s).map_err(invalid_argument)?;
                if sh.is_some() {
                    sheet_override = sh;
                }
                Some((sr, sc, er, ec))
            }
            None => None,
        };

        let sheet_indices =
            if resolved_cell.is_some() || resolved_range.is_some() || sheet_override.is_some() {
                vec![self.sheet_idx(sheet_override.as_deref())?]
            } else {
                (0..self.inner.sheets.len()).collect()
            };

        let mut context = Context::new();
        for workbook_sheet in &self.inner.sheets {
            context.add_table(workbook_sheet.name.clone(), workbook_sheet);
        }
        context.pivot_tables = &self.inner.pivot_tables;
        context.sheet_order = self
            .inner
            .sheets
            .iter()
            .map(|workbook_sheet| workbook_sheet.name.clone())
            .collect();

        let mut diagnostics = Vec::new();
        let mut error_count = 0usize;
        let mut warning_count = 0usize;

        for sheet_idx in sheet_indices {
            let workbook_sheet = &self.inner.sheets[sheet_idx];
            let (start_row, start_col, end_row, end_col) = if let Some((row, col)) = resolved_cell {
                (row, col, row, col)
            } else if let Some(bounds) = resolved_range {
                bounds
            } else {
                (
                    0,
                    0,
                    workbook_sheet.row_count().saturating_sub(1),
                    workbook_sheet.col_count().saturating_sub(1),
                )
            };

            let last_row = end_row.min(workbook_sheet.row_count().saturating_sub(1));
            let last_col = end_col.min(workbook_sheet.col_count().saturating_sub(1));
            for row in start_row..=last_row {
                for col in start_col..=last_col {
                    let src = workbook_sheet
                        .get_src(&CellRef::new(row, col))
                        .cloned()
                        .unwrap_or_default();
                    if !src.starts_with('=') {
                        continue;
                    }
                    for diagnostic in
                        check_formula(&src, workbook_sheet, Some(&context), Some(row), Some(col))
                    {
                        match &diagnostic.severity {
                            FormulaDiagnosticSeverity::Error => error_count += 1,
                            FormulaDiagnosticSeverity::Warning => warning_count += 1,
                        }
                        diagnostics.push(DiagnosticEntry {
                            sheet: workbook_sheet.name.clone(),
                            cell: format!("{}{}", col_idx_to_letters(col), row + 1),
                            row,
                            col,
                            severity: diagnostic_severity(&diagnostic.severity),
                            kind: diagnostic_kind(&diagnostic.kind),
                            message: diagnostic.message,
                            formula: src.clone(),
                        });
                    }
                }
            }
        }

        to_js_value(&TypecheckReport {
            error_count,
            warning_count,
            error_count_camel: error_count,
            warning_count_camel: warning_count,
            diagnostics,
        })
    }

    #[wasm_bindgen(getter, js_name = sheetNames)]
    pub fn sheet_names(&self) -> Vec<String> {
        self.inner.sheets.iter().map(|s| s.name.clone()).collect()
    }

    #[wasm_bindgen(js_name = sheetIndex)]
    pub fn sheet_index(&self, name: Option<String>) -> WasmResult<usize> {
        self.sheet_idx(name.as_deref())
    }

    pub fn dimensions(&self, sheet: Option<String>) -> WasmResult<JsValue> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        let s = &self.inner.sheets[idx];
        to_js_value(&JsDimensions {
            rows: s.row_count(),
            cols: s.col_count(),
        })
    }

    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&mut self, name: &str) -> WasmResult<()> {
        self.inner.add_sheet(name).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = deleteSheet)]
    pub fn delete_sheet(&mut self, name: &str) -> WasmResult<()> {
        self.inner.delete_sheet(name).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = renameSheet)]
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> WasmResult<()> {
        self.inner
            .rename_sheet(old_name, new_name)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = insertRow)]
    pub fn insert_row(&mut self, row: usize, sheet: Option<String>) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner.insert_row(idx, row).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = deleteRow)]
    pub fn delete_row(&mut self, row: usize, sheet: Option<String>) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner.delete_row(idx, row).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = insertCol)]
    pub fn insert_col(&mut self, col: usize, sheet: Option<String>) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner.insert_col(idx, col).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = deleteCol)]
    pub fn delete_col(&mut self, col: usize, sheet: Option<String>) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner.delete_col(idx, col).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = insertCellsShiftDown)]
    pub fn insert_cells_shift_down(
        &mut self,
        row: usize,
        first_col: usize,
        last_col: usize,
        count: usize,
        sheet: Option<String>,
    ) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner
            .insert_cells_shift_down(idx, row, first_col, last_col, count)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = deleteCellsShiftUp)]
    pub fn delete_cells_shift_up(
        &mut self,
        row: usize,
        first_col: usize,
        last_col: usize,
        count: usize,
        sheet: Option<String>,
    ) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner
            .delete_cells_shift_up(idx, row, first_col, last_col, count)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = setCell)]
    pub fn set_cell(
        &mut self,
        row: usize,
        col: usize,
        value: String,
        sheet: Option<String>,
        cell_type: Option<String>,
    ) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        self.inner.ensure_capacity(idx, row, col);
        if let Some(ct_str) = cell_type {
            let ct = parse_cell_type(&ct_str).map_err(Wrapped)?;
            self.inner.set_cell_with_type(idx, row, col, value, ct);
        } else {
            self.inner.set_cell(idx, row, col, value);
        }
        Ok(())
    }

    #[wasm_bindgen(js_name = setCellType)]
    pub fn set_cell_type(
        &mut self,
        row: usize,
        col: usize,
        cell_type: &str,
        sheet: Option<String>,
    ) -> WasmResult<()> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        let ct = parse_cell_type(cell_type).map_err(Wrapped)?;
        self.inner.set_cell_type(idx, row, col, ct);
        Ok(())
    }

    #[wasm_bindgen(js_name = getCell)]
    pub fn get_cell(&self, row: usize, col: usize, sheet: Option<String>) -> WasmResult<JsValue> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        let v = self.inner.sheets[idx].get_result_data(&CellRef::new(row, col));
        Ok(result_to_js(&v))
    }

    #[wasm_bindgen(js_name = getDisplay)]
    pub fn get_display(&self, row: usize, col: usize, sheet: Option<String>) -> WasmResult<String> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        Ok(self.inner.sheets[idx].get_display_string(&CellRef::new(row, col)))
    }

    #[wasm_bindgen(js_name = getCellType)]
    pub fn get_cell_type(
        &self,
        row: usize,
        col: usize,
        sheet: Option<String>,
    ) -> WasmResult<String> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        let t = self.inner.get_cell_type(idx, row, col);
        Ok(t.as_str().to_string())
    }

    #[wasm_bindgen(js_name = getSrc)]
    pub fn get_src(&self, row: usize, col: usize, sheet: Option<String>) -> WasmResult<String> {
        let idx = self.sheet_idx(sheet.as_deref())?;
        Ok(self.inner.sheets[idx].get_src_str(&CellRef::new(row, col)))
    }

    #[wasm_bindgen(js_name = getRange)]
    pub fn get_range(
        &self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        sheet: Option<String>,
    ) -> WasmResult<JsValue> {
        if start_row > end_row || start_col > end_col {
            return Err(Wrapped(visi_engine::Error::InvalidRange(
                "range end must not precede its start".to_string(),
            ))
            .into());
        }
        let idx = self.sheet_idx(sheet.as_deref())?;
        let s = &self.inner.sheets[idx];
        let rows = Array::new();
        for r in start_row..=end_row {
            let row_arr = Array::new();
            for c in start_col..=end_col {
                let cell_ref = CellRef::new(r, c);
                let obj = Object::new();
                let val = result_to_js(&s.get_result_data(&cell_ref));
                let display = s.get_display_string(&cell_ref);
                let src = s.get_src_str(&cell_ref);
                let cell_type = s.get_cell_type(&cell_ref).as_str();
                let style_val = match s.get_cell_style(r, c).cloned() {
                    Some(st) => to_js_value(&JsCellStyle::from(st))?,
                    None => JsValue::NULL,
                };
                let _ = Reflect::set(&obj, &JsValue::from_str("value"), &val);
                let _ = Reflect::set(
                    &obj,
                    &JsValue::from_str("display"),
                    &JsValue::from_str(&display),
                );
                let _ = Reflect::set(&obj, &JsValue::from_str("src"), &JsValue::from_str(&src));
                let _ = Reflect::set(
                    &obj,
                    &JsValue::from_str("cellType"),
                    &JsValue::from_str(cell_type),
                );
                let _ = Reflect::set(&obj, &JsValue::from_str("style"), &style_val);
                row_arr.push(&obj);
            }
            rows.push(&row_arr);
        }
        Ok(rows.into())
    }

    #[wasm_bindgen(js_name = setCells)]
    pub fn set_cells(&mut self, updates: JsValue, evaluate: Option<bool>) -> WasmResult<()> {
        let items: Vec<JsCellUpdate> = from_js_req(updates)?;
        for item in items {
            let idx = self.sheet_idx(item.sheet.as_deref())?;
            if let Some(ct_str) = item.cell_type {
                let ct = parse_cell_type(&ct_str).map_err(Wrapped)?;
                self.inner
                    .set_cell_with_type(idx, item.row, item.col, item.value, ct);
            } else {
                self.inner.set_cell(idx, item.row, item.col, item.value);
            }
            if let Some(st) = item.style {
                let core_style: CellStyle = st.into();
                self.inner.sheets[idx].update_cell_style(item.row, item.col, |existing| {
                    existing.merge(&core_style);
                });
            }
        }
        if evaluate.unwrap_or(false) {
            self.inner.evaluate().map_err(Wrapped)?;
        }
        Ok(())
    }

    #[wasm_bindgen(js_name = getCellStyle)]
    pub fn get_cell_style(
        &self,
        row: usize,
        col: usize,
        sheet: Option<String>,
    ) -> WasmResult<JsValue> {
        let style = self
            .inner
            .get_cell_style(sheet.as_deref(), row, col)
            .map_err(Wrapped)?;
        match style {
            Some(st) => to_js_value(&JsCellStyle::from(st)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = setCellStyle)]
    pub fn set_cell_style(
        &mut self,
        row: usize,
        col: usize,
        style: JsValue,
        sheet: Option<String>,
    ) -> WasmResult<()> {
        let js_style: JsCellStyle = from_js_req(style)?;
        self.inner
            .set_cell_style(sheet.as_deref(), row, col, js_style.into())
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = setRangeStyle)]
    pub fn set_range_style(
        &mut self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        style: JsValue,
        sheet: Option<String>,
    ) -> WasmResult<()> {
        let js_style: JsCellStyle = from_js_req(style)?;
        self.inner
            .set_range_style(
                sheet.as_deref(),
                start_row,
                start_col,
                end_row,
                end_col,
                js_style.into(),
            )
            .map_err(Wrapped)?;
        Ok(())
    }

    pub fn tables(&self) -> WasmResult<JsValue> {
        let list: Vec<JsTableInfo> = self
            .inner
            .list_tables()
            .into_iter()
            .map(|(sheet_name, t)| JsTableInfo {
                id: t.id,
                name: t.name.clone(),
                sheet: sheet_name.to_string(),
                sheet_id: t.sheet_id,
                start_row: t.start_row,
                start_col: t.start_col,
                end_row: t.end_row,
                end_col: t.end_col,
                has_header_row: t.has_header_row,
                has_totals_row: t.has_totals_row,
                columns: t.columns.clone(),
                style_name: t.style_name.clone(),
            })
            .collect();
        to_js_value(&list)
    }

    #[wasm_bindgen(js_name = addTable)]
    pub fn add_table(&mut self, options: JsValue) -> WasmResult<f64> {
        let opts: JsAddTableOptions = from_js_req(options)?;
        let id = self
            .inner
            .add_table(
                opts.sheet.as_deref(),
                &opts.name,
                opts.start_row,
                opts.start_col,
                opts.end_row,
                opts.end_col,
                opts.has_header_row,
                opts.has_totals_row,
            )
            .map_err(Wrapped)?;
        if let Some(style_name) = opts.style_name {
            self.inner
                .set_table_style(&opts.name, &style_name)
                .map_err(Wrapped)?;
        }
        Ok(id as f64)
    }

    #[wasm_bindgen(js_name = deleteTable)]
    pub fn delete_table(&mut self, name: &str) -> WasmResult<()> {
        self.inner.delete_table(name).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = renameTable)]
    pub fn rename_table(&mut self, old_name: &str, new_name: &str) -> WasmResult<()> {
        self.inner
            .rename_table(old_name, new_name)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = resizeTable)]
    pub fn resize_table(
        &mut self,
        name: &str,
        new_end_row: usize,
        new_end_col: usize,
    ) -> WasmResult<()> {
        self.inner
            .resize_table(name, new_end_row, new_end_col)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = renameTableColumn)]
    pub fn rename_table_column(
        &mut self,
        table_name: &str,
        col_index: usize,
        new_name: &str,
    ) -> WasmResult<()> {
        self.inner
            .rename_table_column(table_name, col_index, new_name)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = getTableStyle)]
    pub fn get_table_style(&self, table_name: &str) -> WasmResult<Option<String>> {
        Ok(self.inner.get_table_style(table_name).map_err(Wrapped)?)
    }

    #[wasm_bindgen(js_name = setTableStyle)]
    pub fn set_table_style(&mut self, table_name: &str, style_name: &str) -> WasmResult<()> {
        self.inner
            .set_table_style(table_name, style_name)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = addChart)]
    pub fn add_chart(
        &mut self,
        sheet: &str,
        chart_type: &str,
        range: String,
        options: Option<JsValue>,
    ) -> WasmResult<f64> {
        let ct = parse_chart_type(chart_type).map_err(Wrapped)?;
        let opts: JsAddChartOptions = from_js_opt(options)?;
        let id = self
            .inner
            .add_chart(sheet, ct, range, opts.title, opts.anchor)
            .map_err(Wrapped)?;
        Ok(id as f64)
    }

    #[wasm_bindgen(js_name = editChart)]
    pub fn edit_chart(&mut self, chart_id: f64, options: JsValue) -> WasmResult<()> {
        fn tri(
            value: Option<String>,
            clear: bool,
            what: &str,
        ) -> WasmResult<Option<Option<String>>> {
            match (value, clear) {
                (Some(_), true) => Err(invalid_argument(format!(
                    "pass either {what} or clear_{what}, not both"
                ))),
                (Some(v), false) => Ok(Some(Some(v))),
                (None, true) => Ok(Some(None)),
                (None, false) => Ok(None),
            }
        }

        let opts: JsEditChartOptions = from_js_req(options)?;
        let title = tri(opts.title, opts.clear_title, "title")?;
        let xlabel = tri(opts.xlabel, opts.clear_xlabel, "xlabel")?;
        let ylabel = tri(opts.ylabel, opts.clear_ylabel, "ylabel")?;
        let ct: Option<ChartType> = opts
            .chart_type
            .as_deref()
            .map(parse_chart_type)
            .transpose()
            .map_err(Wrapped)?;

        self.inner
            .edit_chart(
                chart_id as u64,
                opts.name,
                ct,
                opts.range,
                title,
                xlabel,
                ylabel,
                opts.show_legend,
                opts.anchor,
            )
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = deleteChart)]
    pub fn delete_chart(&mut self, chart_id: f64) -> WasmResult<()> {
        self.inner.delete_chart(chart_id as u64).map_err(Wrapped)?;
        Ok(())
    }

    pub fn charts(&self) -> WasmResult<JsValue> {
        let out: Vec<JsChartInfo> = self
            .inner
            .charts
            .iter()
            .map(|c| JsChartInfo {
                id: c.id,
                name: c.name.clone(),
                chart_type: format!("{:?}", c.chart_type),
                data_range: c.data_range.clone(),
                data_range_camel: c.data_range.clone(),
                title: c.title.clone(),
                anchor: format!("{}{}", col_idx_to_letters(c.anchor_col), c.anchor_row + 1),
                xlabel: c.xlabel.clone(),
                ylabel: c.ylabel.clone(),
                show_legend: c.show_legend,
                show_legend_camel: c.show_legend,
            })
            .collect();
        to_js_value(&out)
    }

    #[wasm_bindgen(js_name = addPivotFromTable)]
    pub fn add_pivot_from_table(
        &mut self,
        name: &str,
        source_table: &str,
        options: Option<JsValue>,
    ) -> WasmResult<f64> {
        let opts: JsPivotDestOptions = from_js_opt(options)?;
        let id = self
            .inner
            .add_pivot_table_from_table(
                name,
                source_table,
                opts.dest_sheet.as_deref(),
                opts.dest_row,
                opts.dest_col,
                opts.grand_totals_row,
                opts.grand_totals_col,
            )
            .map_err(Wrapped)?;
        Ok(id as f64)
    }

    #[wasm_bindgen(js_name = addPivotFromRange)]
    pub fn add_pivot_from_range(&mut self, name: &str, options: JsValue) -> WasmResult<f64> {
        let opts: JsPivotRangeOptions = from_js_req(options)?;
        let id = self
            .inner
            .add_pivot_table_from_range(
                name,
                opts.source_sheet.as_deref(),
                opts.start_row,
                opts.start_col,
                opts.end_row,
                opts.end_col,
                opts.dest_sheet.as_deref(),
                opts.dest_row,
                opts.dest_col,
                opts.grand_totals_row,
                opts.grand_totals_col,
            )
            .map_err(Wrapped)?;
        Ok(id as f64)
    }

    #[wasm_bindgen(js_name = addPivotField)]
    pub fn add_pivot_field(
        &mut self,
        pivot: &str,
        area: &str,
        column: &str,
        options: Option<JsValue>,
    ) -> WasmResult<()> {
        let area = parse_pivot_area(area).map_err(Wrapped)?;
        let opts: JsPivotFieldOptions = from_js_opt(options)?;
        let agg = opts
            .agg
            .as_deref()
            .map(parse_pivot_agg)
            .transpose()
            .map_err(Wrapped)?;
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
            if !opts.subtotal {
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
            if let Some(l) = opts.label
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

    #[wasm_bindgen(js_name = removePivotField)]
    pub fn remove_pivot_field(&mut self, pivot: &str, area: &str, column: &str) -> WasmResult<()> {
        let area = parse_pivot_area(area).map_err(Wrapped)?;
        self.inner
            .remove_pivot_field(pivot, area, column)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = setPivotFilter)]
    pub fn set_pivot_filter(
        &mut self,
        pivot: &str,
        column: &str,
        values: JsValue,
    ) -> WasmResult<()> {
        let parsed: Option<Vec<String>> = if values.is_null() || values.is_undefined() {
            None
        } else {
            Some(from_js_req(values)?)
        };
        self.inner
            .set_pivot_filter(pivot, column, parsed)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = refreshPivot)]
    pub fn refresh_pivot(&mut self, pivot: &str) -> WasmResult<()> {
        self.inner.refresh_pivot_table(pivot).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = renamePivot)]
    pub fn rename_pivot(&mut self, old_name: &str, new_name: &str) -> WasmResult<()> {
        self.inner
            .rename_pivot_table(old_name, new_name)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = deletePivot)]
    pub fn delete_pivot(&mut self, pivot: &str) -> WasmResult<()> {
        self.inner.delete_pivot_table(pivot).map_err(Wrapped)?;
        Ok(())
    }

    pub fn pivots(&self) -> WasmResult<JsValue> {
        let mut out = Vec::with_capacity(self.inner.pivot_tables.len());
        for p in self.inner.list_pivot_tables() {
            let row_fields: Vec<String> = p.row_fields.iter().map(|f| f.column.clone()).collect();
            let col_fields: Vec<String> = p.col_fields.iter().map(|f| f.column.clone()).collect();
            let value_fields = value_field_labels(&p.value_fields);
            let filter_fields: Vec<String> =
                p.filter_fields.iter().map(|f| f.column.clone()).collect();

            let mut subtotals = HashMap::new();
            for f in p.row_fields.iter().chain(p.col_fields.iter()) {
                subtotals.insert(f.column.clone(), f.subtotal);
            }

            let mut filter_selections = HashMap::new();
            for f in &p.filter_fields {
                filter_selections.insert(f.column.clone(), f.selected_values.clone());
            }

            out.push(JsPivotInfo {
                id: p.id,
                name: p.name.clone(),
                row_fields_camel: row_fields.clone(),
                col_fields_camel: col_fields.clone(),
                value_fields_camel: value_fields.clone(),
                filter_fields_camel: filter_fields.clone(),
                row_fields,
                col_fields,
                value_fields,
                filter_fields,
                subtotals,
                filter_selections_camel: filter_selections.clone(),
                filter_selections,
            });
        }
        to_js_value(&out)
    }

    #[wasm_bindgen(js_name = hasMacros)]
    pub fn has_macros(&self) -> bool {
        self.inner.has_vba_project()
    }

    #[wasm_bindgen(js_name = addMacro)]
    pub fn add_macro(
        &mut self,
        name: &str,
        source: &str,
        options: Option<JsValue>,
    ) -> WasmResult<()> {
        let opts: JsAddMacroOptions = from_js_opt(options)?;
        let kind_str = opts.kind.as_deref().unwrap_or("standard");
        let kind = parse_vba_module_kind(kind_str).map_err(Wrapped)?;
        let bound_sheet_id = match (kind, opts.sheet.as_deref()) {
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

    #[wasm_bindgen(js_name = removeMacro)]
    pub fn remove_macro(&mut self, name: &str) -> WasmResult<()> {
        self.inner.remove_vba_module(name).map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = renameMacro)]
    pub fn rename_macro(&mut self, old_name: &str, new_name: &str) -> WasmResult<()> {
        self.inner
            .rename_vba_module(old_name, new_name)
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = setMacroSource)]
    pub fn set_macro_source(&mut self, name: &str, source: &str) -> WasmResult<()> {
        self.inner
            .set_vba_module_source(name, source.to_string())
            .map_err(Wrapped)?;
        Ok(())
    }

    #[wasm_bindgen(js_name = runMacro)]
    pub fn run_macro(&mut self, procedure: &str, options: Option<JsValue>) -> WasmResult<JsValue> {
        let opts: JsRunMacroOptions = from_js_opt(options)?;
        let args = opts.args.unwrap_or_default();
        let refs: Vec<&str> = args.iter().map(|s| s.as_str()).collect();
        let out = self
            .inner
            .run_macro(opts.module.as_deref(), procedure, &refs)
            .map_err(Wrapped)?;
        to_js_value(&JsMacroOutcome {
            type_short: out.type_name.clone(),
            type_name: out.type_name.clone(),
            type_name_camel: out.type_name,
            value: out.value,
            mutated: out.mutated,
        })
    }

    #[wasm_bindgen(js_name = runOpenEvents)]
    pub fn run_open_events(&mut self) -> WasmResult<JsValue> {
        let out = self.inner.run_open_events().map_err(Wrapped)?;
        to_js_value(&JsMacroOutcome {
            type_short: out.type_name.clone(),
            type_name: out.type_name.clone(),
            type_name_camel: out.type_name,
            value: out.value,
            mutated: out.mutated,
        })
    }

    pub fn macros(&self) -> WasmResult<JsValue> {
        let modules = self.inner.list_vba_modules();
        let mut out = Vec::with_capacity(modules.len());
        for m in modules {
            let lines = m.source.lines().count();
            out.push(JsVbaModuleInfo {
                name: m.name.clone(),
                kind: format!("{:?}", m.kind),
                source: m.source.clone(),
                source_lines: lines,
                source_lines_camel: lines,
                bound_sheet_id: m.bound_sheet_id,
                bound_sheet_id_camel: m.bound_sheet_id,
            });
        }
        to_js_value(&out)
    }
}
