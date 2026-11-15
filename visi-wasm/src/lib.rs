use js_sys::Reflect;
use serde::Serialize;
use visi_engine::core::{
    col_idx_to_letters as core_col_idx_to_letters, col_letters_to_idx as core_col_letters_to_idx,
    parse_cell_ref as core_parse_cell_ref, parse_range_ref as core_parse_range_ref,
};
use visi_engine::{WorkbookManager, detect_workbook_format as core_detect_workbook_format};
use wasm_bindgen::prelude::*;

mod enums;
mod errors;
mod value;
mod workbook;

pub use value::CellError;
pub use workbook::Workbook;

use enums::parse_workbook_format;
use errors::{WasmResult, Wrapped, invalid_argument};
use workbook::{JsMacroOutcome, to_js_value};

#[wasm_bindgen(typescript_custom_section)]
const TS_CUSTOM_TYPES: &'static str = r#"
export type WorkbookFormat = "xlsx" | "xlsm" | "xlsb";
export type CellType =
  | "empty"
  | "int"
  | "float"
  | "string"
  | "bool"
  | "date-time"
  | "date-time-iso"
  | "duration-iso"
  | "error";
export type ChartType = "column" | "bar" | "line" | "pie" | "scatter" | "area";
export type PivotArea = "row" | "column" | "value" | "filter";
export type PivotAggregation = "sum" | "count" | "count-numbers" | "average" | "max" | "min";
export type VbaModuleKind = "standard" | "class" | "document";

export type CellValue =
  | null
  | boolean
  | number
  | string
  | CellError
  | CellValue[]
  | Map<CellValue, CellValue>;

export interface CellStyle {
  fontColor?: string;
  bgColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSize?: number;
  numFormat?: string;
}

export interface CellSnapshot {
  value: CellValue;
  display: string;
  src: string;
  cellType: CellType;
  style: CellStyle | null;
}

export interface CellUpdate {
  row: number;
  col: number;
  value: string;
  sheet?: string;
  cellType?: CellType;
  style?: CellStyle;
}

export interface SheetSummary {
  name: string;
  rowCount: number;
  colCount: number;
  formulaCount: number;
}

export interface WorkbookSummary {
  fileName: string;
  sheetCount: number;
  chartCount: number;
  sheets: SheetSummary[];
}

export interface TypecheckOptions {
  sheet?: string;
  cell?: [number, number] | string;
  range?: [number, number, number, number] | string;
}

export interface FormulaDiagnostic {
  sheet: string;
  cell: string;
  row: number;
  col: number;
  severity: "error" | "warning";
  kind: "syntax" | "reference" | "function" | "type";
  message: string;
  formula: string;
}

export interface TypecheckResult {
  error_count: number;
  warning_count: number;
  errorCount: number;
  warningCount: number;
  diagnostics: FormulaDiagnostic[];
}

export interface AddTableOptions {
  name: string;
  sheet?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  hasHeaderRow?: boolean;
  hasTotalsRow?: boolean;
  styleName?: string;
}

export interface TableInfo {
  id: number;
  name: string;
  sheet: string;
  sheetId: number;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  hasHeaderRow: boolean;
  hasTotalsRow: boolean;
  columns: string[];
  styleName: string | null;
}

export interface AddChartOptions {
  title?: string;
  anchor?: [number, number];
}

export interface EditChartOptions {
  name?: string;
  chartType?: ChartType;
  range?: string;
  title?: string;
  clearTitle?: boolean;
  xlabel?: string;
  clearXlabel?: boolean;
  ylabel?: string;
  clearYlabel?: boolean;
  showLegend?: boolean;
  anchor?: [number, number];
}

export interface ChartInfo {
  id: number;
  name: string;
  type: string;
  data_range: string;
  dataRange: string;
  title: string | null;
  anchor: string;
  xlabel: string | null;
  ylabel: string | null;
  show_legend: boolean;
  showLegend: boolean;
}

export interface PivotDestOptions {
  destSheet?: string;
  destRow?: number;
  destCol?: number;
  grandTotalsRow?: boolean;
  grandTotalsCol?: boolean;
}

export interface PivotRangeOptions extends PivotDestOptions {
  sourceSheet?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface PivotFieldOptions {
  agg?: PivotAggregation;
  subtotal?: boolean;
  label?: string;
}

export interface PivotInfo {
  id: number;
  name: string;
  row_fields: string[];
  col_fields: string[];
  value_fields: string[];
  filter_fields: string[];
  rowFields: string[];
  colFields: string[];
  valueFields: string[];
  filterFields: string[];
  subtotals: Record<string, boolean>;
  filter_selections: Record<string, string[] | null>;
  filterSelections: Record<string, string[] | null>;
}

export interface VbaModuleInfo {
  name: string;
  kind: string;
  source: string;
  source_lines: number;
  sourceLines: number;
  bound_sheet_id: number | null;
  boundSheetId: number | null;
}

export interface MacroRunOutcome {
  type: string;
  type_name: string;
  typeName: string;
  value: string | null;
  mutated: boolean;
}

export interface ParsedCellRef {
  sheet: string | null;
  row: number;
  col: number;
}

export interface ParsedRangeRef {
  sheet: string | null;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
"#;

#[derive(Serialize)]
struct JsParsedCellRef {
    sheet: Option<String>,
    row: usize,
    col: usize,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct JsParsedRangeRef {
    sheet: Option<String>,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
}

#[wasm_bindgen(js_name = evalBytes)]
pub fn eval_bytes(input: &[u8], format: Option<String>) -> WasmResult<Vec<u8>> {
    let mut wb = WorkbookManager::load_bytes(input).map_err(Wrapped)?;
    wb.evaluate().map_err(Wrapped)?;
    let fmt = match format.as_deref() {
        Some(s) => parse_workbook_format(s).map_err(Wrapped)?,
        None => core_detect_workbook_format(input).map_err(Wrapped)?,
    };
    Ok(wb.save_bytes_as(fmt).map_err(Wrapped)?)
}

#[wasm_bindgen(js_name = detectWorkbookFormat)]
pub fn detect_workbook_format(data: &[u8]) -> WasmResult<String> {
    let fmt = core_detect_workbook_format(data).map_err(Wrapped)?;
    Ok(fmt.extension().to_string())
}

#[wasm_bindgen(js_name = checkSyntax)]
pub fn check_syntax(source: &str) -> WasmResult<Vec<String>> {
    Ok(visi_engine::core::check_syntax(source)
        .map_err(Wrapped)?
        .procedures)
}

#[wasm_bindgen(js_name = checkSyntaxPartial)]
pub fn check_syntax_partial(source: &str) -> WasmResult<Vec<String>> {
    Ok(visi_engine::core::check_syntax_partial(source)
        .map_err(Wrapped)?
        .procedures)
}

#[wasm_bindgen(js_name = runMacro)]
pub fn run_macro(source: &str, procedure: &str, args: Option<Vec<String>>) -> WasmResult<JsValue> {
    let args = args.unwrap_or_default();
    let refs: Vec<&str> = args.iter().map(|s| s.as_str()).collect();
    let out = visi_engine::core::run_macro(source, procedure, &refs).map_err(Wrapped)?;
    to_js_value(&JsMacroOutcome {
        type_short: out.type_name.clone(),
        type_name: out.type_name.clone(),
        type_name_camel: out.type_name,
        value: out.value,
        mutated: out.mutated,
    })
}

#[wasm_bindgen(js_name = colIdxToLetters)]
pub fn col_idx_to_letters(col: usize) -> String {
    core_col_idx_to_letters(col)
}

#[wasm_bindgen(js_name = colLettersToIdx)]
pub fn col_letters_to_idx(letters: &str) -> usize {
    core_col_letters_to_idx(letters)
}

#[wasm_bindgen(js_name = parseCellRef)]
pub fn parse_cell_ref(cell_str: &str) -> WasmResult<JsValue> {
    let (sheet, row, col) = core_parse_cell_ref(cell_str).map_err(invalid_argument)?;
    to_js_value(&JsParsedCellRef { sheet, row, col })
}

#[wasm_bindgen(js_name = parseRangeRef)]
pub fn parse_range_ref(range_str: &str) -> WasmResult<JsValue> {
    let (sheet, start_row, start_col, end_row, end_col) =
        core_parse_range_ref(range_str).map_err(invalid_argument)?;
    to_js_value(&JsParsedRangeRef {
        sheet,
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

#[wasm_bindgen(js_name = isCellError)]
pub fn is_cell_error(val: &JsValue) -> bool {
    if !val.is_object() || val.is_null() {
        return false;
    }
    Reflect::get(val, &JsValue::from_str("code"))
        .ok()
        .is_some_and(|v| v.is_string())
        && Reflect::get(val, &JsValue::from_str("toJSON"))
            .ok()
            .is_some_and(|v| v.is_function())
}

#[wasm_bindgen]
pub fn version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}
