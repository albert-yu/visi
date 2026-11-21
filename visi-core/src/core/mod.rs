pub(crate) mod actions;
#[doc = "Chart definitions and workbook chart helpers."]
pub mod chart;
pub(crate) mod csv_import;
pub(crate) mod date;
pub(crate) mod date_fn;
#[doc = "Spreadsheet engine types for cells, columns, sheets, and evaluation."]
pub mod engine;
pub(crate) mod engineering;
pub(crate) mod ets;
#[doc = "Excel package format detection and output format selection."]
pub mod excel;
pub(crate) mod extended_fn;
pub(crate) mod finance;
pub(crate) mod formula;
pub(crate) mod function_registry;
pub(crate) mod grid_edit;
#[doc = "Locale settings used for date parsing and formatting."]
pub mod locale;
pub(crate) mod math_trig;
pub(crate) mod opc;
#[doc(hidden)]
pub mod ovba;
pub(crate) mod parser;
#[doc = "Pivot table definitions and computation helpers."]
pub mod pivot;
pub(crate) mod pivot_xlsx;
pub(crate) mod shared_vec;
pub(crate) mod stats;
#[doc = "Cell style data used by sheets and file I/O."]
pub mod style;
#[doc = "Excel table metadata and structured-reference helpers."]
pub mod table;
pub(crate) mod text;
#[doc = "Type checking of cell values and formula inputs"]
pub mod typecheck;
#[doc = "VBA project, module, syntax-checking, and macro execution support."]
pub mod vba;
pub(crate) mod vba_synth;
pub(crate) mod vba_validate;
#[doc(hidden)]
pub mod vba_xlsx;
#[doc = "Workbook-level management APIs over sheets and workbook objects."]
pub mod workbook;
#[doc = "Import and export support for `.xlsb` workbooks."]
pub mod xlsb;
#[doc = "Import and export support for `.xlsx` workbooks."]
pub mod xlsx;
pub(crate) mod xml;

pub use actions::SheetAction;
pub use date::{
    DateFormat, EXCEL_1904_EPOCH_OFFSET, SimpleDate, StringCase, date_to_excel_serial,
    date_to_excel_serial_with_epoch, excel_serial_to_date, excel_serial_to_date_with_epoch,
    format_date, is_date_code, parse_date, parse_date_with_locale, render_date_code,
};
pub use engine::{
    ArrayFormulaEntry, ArrayFormulaKind, Bitmask, CellComment, CellHyperlink, CellRef, CellType,
    ColumnData, ConditionalFormatRule, Context, DataColumn, DataValidationKind,
    DataValidationOperator, DataValidationRule, Dependency, Direction, EngineError, EvalError,
    FreezePane, MergedCellRange, RefType, ResultData, Sheet, SheetInit, SheetVisibility,
    TextCellRef, generate_unique_id, get_word_boundaries_from_str,
};
pub use excel::{WorkbookFormat, detect_workbook_format};
pub use formula::{CompiledFormula, FormulaPart, SheetSection};
pub use locale::{DateOrder, Locale};
pub use parser::{
    col_idx_to_letters, col_letters_to_idx, parse_a1_coordinates, parse_cell_ref, parse_range_ref,
};
pub use pivot::{
    PivotAggregation, PivotArea, PivotBodyRow, PivotField, PivotFilterField, PivotGrid,
    PivotSource, PivotTable, PivotValueField, compute_pivot, value_field_labels,
};
pub use shared_vec::SharedVec;
pub use style::{
    BorderSide, BorderStyle, CellAlignment, CellBorders, CellStyle, FillPattern,
    HorizontalAlignment, VerticalAlignment,
};
pub use table::ExcelTable;
pub use vba::{
    ModuleSyntax, RunOutcome, VbaModule, VbaModuleKind, VbaProject, check_syntax,
    check_syntax_partial, run_macro, validate_vba_module_name,
};
pub use workbook::{DefinedName, SheetSummary, WorkbookManager, WorkbookSummary};
