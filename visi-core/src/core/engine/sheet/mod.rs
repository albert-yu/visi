//! The `Sheet` type: a worksheet's cells, its formula evaluation, and the
//! dependency-tracked recalculation over them.

mod edit;
mod functions;

pub use edit::get_word_boundaries_from_str;

use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use web_time::Instant;

use super::cell::{
    CellRef, CellType, Dependency, EngineError, EvalError, TextCellRef, generate_unique_id,
};
use super::column::DataColumn;
use super::result_data::ResultData;
/// Context for evaluating expressions, containing references to other sheets
#[derive(Default)]
pub struct Context<'a> {
    /// Map of sheet names to sheet references for cross-sheet lookups
    pub sheets: HashMap<String, &'a Sheet>,
    /// Every pivot table in the workbook, so `GETPIVOTDATA` can resolve a
    /// rendered pivot's destination cell back to its definition. Pivot
    /// tables are workbook-level (like `Context.sheets`' cross-sheet
    /// lookups), not sheet-scoped, so this lives here rather than on
    /// `Sheet` itself.
    pub pivot_tables: &'a [crate::core::pivot::PivotTable],
    /// Sheet names in true workbook order, so `SHEET()` can report a real
    /// ordinal. `sheets` is an unordered `HashMap`, which is why this is
    /// tracked separately rather than derived from it -- true order only
    /// exists one layer up, in `visi`'s `WorkbookManager::sheets` (a
    /// `Vec`), which populates this when building the context.
    pub sheet_order: Vec<String>,
}

impl<'a> Context<'a> {
    /// Create a new empty context
    pub fn new() -> Self {
        Self {
            sheets: HashMap::new(),
            pivot_tables: &[],
            sheet_order: Vec::new(),
        }
    }

    /// Add a sheet to the context for lookup during evaluation
    pub fn add_table(&mut self, name: String, sheet: &'a Sheet) {
        self.sheets.insert(name, sheet);
    }
}

/// A chain of LET name/value bindings in scope while evaluating a single
/// formula. This is a linked list (not a cloned `HashMap`) because LET
/// binds names one at a time -- each value expression, and the final
/// calculation, must see all *earlier* bindings from the same LET (and any
/// outer LET it's nested inside), and a name can shadow an outer binding of
/// the same spelling. `evaluate_let` builds this chain by recursing one
/// pair at a time rather than mutating a shared map.
enum LetScope<'a> {
    Empty,
    Bound {
        name: &'a str,
        value: &'a ResultData,
        parent: &'a LetScope<'a>,
    },
}

impl<'a> LetScope<'a> {
    fn get(&self, name: &str) -> Option<&ResultData> {
        match self {
            LetScope::Empty => None,
            LetScope::Bound {
                name: n,
                value,
                parent,
            } => {
                if n.eq_ignore_ascii_case(name) {
                    Some(value)
                } else {
                    parent.get(name)
                }
            }
        }
    }
}

/// Which way a fill or selection extends from its anchor cell.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Direction {
    /// No direction; the operation is a no-op.
    None,
    /// Toward row 0.
    Up,
    /// Toward the last row.
    Down,
    /// Toward column 0.
    Left,
    /// Toward the last column.
    Right,
}

/// One worksheet: a grid of cells, the formulas over them, and the dependency
/// graph that keeps them up to date.
///
/// # Coordinates
///
/// Everything here is **0-based `(row, col)`**. A1 notation exists only at the
/// parser and CLI boundaries -- see [`parse_a1_coordinates`] and
/// [`col_idx_to_letters`] to convert.
///
/// # Naming trap
///
/// A `Sheet` is informally called a "table" in places (a new one is named
/// `table_1`, and `Context::add_table` registers one). That is *not* an
/// [`ExcelTable`], which is a ListObject -- a named rectangular range *on* a
/// sheet -- and lives in [`Sheet::tables`].
///
/// # Storage
///
/// Storage is column-oriented: each [`DataColumn`] keeps the raw user text,
/// the computed values and the compiled formulas in three parallel vectors
/// that must stay the same length. The row and column insert/delete paths
/// maintain that invariant by hand, so a new one has to do the same.
///
/// # Recalculation
///
/// [`Sheet::commit`] recomputes the dirty cells and propagates through
/// [`Dependency::Local`] and [`Dependency::LocalColumn`] edges only.
/// Cross-sheet edges are `WorkbookManager::evaluate`'s job, and evaluating a
/// formula with a remote reference requires a [`Context`] -- without one it
/// errors.
///
/// [`parse_a1_coordinates`]: crate::core::parse_a1_coordinates
/// [`col_idx_to_letters`]: crate::core::col_idx_to_letters
/// [`ExcelTable`]: crate::core::table::ExcelTable
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    /// Workbook-unique identifier. Formulas compile references against this
    /// rather than the name, which is what makes a rename non-destructive.
    #[serde(default = "generate_unique_id")]
    pub id: u64,
    /// Display name, as it appears in a cross-sheet reference.
    pub name: String,
    /// The cells, one entry per column. Row `r` of column `c` is
    /// `columns[c]`'s entry `r`.
    ///
    /// Every column has the same number of rows -- [`Sheet::row_count`] reads
    /// only the first and assumes the rest match -- so the `Vec` itself is
    /// crate-private. Read them through [`Sheet::columns`].
    pub(crate) columns: Vec<DataColumn>,
    /// Excel Tables (ListObjects) defined on this sheet.
    #[serde(default)]
    pub tables: Vec<crate::core::table::ExcelTable>,
    /// Forward edges: which cells must be recomputed when a dependency
    /// changes. Rebuilt from the formulas, so not serialized.
    #[serde(skip, default)]
    pub dependencies: HashMap<Dependency, HashSet<CellRef>>,
    /// Reverse edges: what each cell currently reads, so its old edges can be
    /// dropped when its formula changes. Rebuilt, so not serialized.
    #[serde(skip, default)]
    pub dependencies_rev: HashMap<CellRef, HashSet<Dependency>>,
    /// Edits made since the last commit, for callers that want to observe or
    /// replay them.
    #[serde(skip)]
    pub uncommitted_actions: Vec<crate::core::SheetAction>,
    /// Regional locale for date and number parsing.
    #[serde(default)]
    pub locale: crate::core::locale::Locale,
}

/// Arguments for [`Sheet::new`]. [`Default`] gives a 10x5 sheet with a
/// generated id and the name `table_1`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetInit {
    /// Identifier to use; `None` generates a fresh one.
    #[serde(default)]
    pub id: Option<u64>,
    /// Name to use; `None` means `table_1`.
    pub name: Option<String>,
    /// Rows to allocate.
    pub rows: usize,
    /// Columns to allocate.
    pub cols: usize,
}

impl Default for SheetInit {
    fn default() -> Self {
        Self {
            id: None,
            name: None,
            rows: 10,
            cols: 5,
        }
    }
}

/// How a blank cell is treated by the strict numeric flatteners.
#[derive(Clone, Copy, PartialEq, Eq)]
enum BlankPolicy {
    /// Counts as 0 (MULTINOMIAL).
    Zero,
    /// Dropped entirely, shifting later elements (SERIESSUM).
    Skip,
    /// #VALUE!, like text (LINEST/TREND/GROWTH/LOGEST/MMULT).
    Reject,
}

impl Sheet {
    /// Creates a sheet of `args.rows` x `args.cols` empty cells, every one of
    /// them queued as a pending edit so the first [`Sheet::commit`] sees them.
    pub fn new(args: SheetInit) -> Sheet {
        let SheetInit {
            id,
            name,
            rows,
            cols,
        } = args;
        let sheet_id = id.unwrap_or_else(generate_unique_id);
        let sheet_name = name.unwrap_or_else(|| "table_1".to_string());

        let mut columns = Vec::with_capacity(cols);
        for _ in 0..cols {
            columns.push(DataColumn::new(rows));
        }

        let mut uncommitted_actions = Vec::new();
        for c in 0..cols {
            for r in 0..rows {
                uncommitted_actions.push(crate::core::SheetAction::SetCellSrc {
                    sheet_name: sheet_name.clone(),
                    col: c,
                    row: r,
                    src: String::new(),
                });
            }
        }

        Self {
            id: sheet_id,
            name: sheet_name,
            columns,
            tables: Vec::new(),
            dependencies: HashMap::new(),
            dependencies_rev: HashMap::new(),
            uncommitted_actions,
            locale: crate::core::locale::Locale::default(),
        }
    }

    /// Rebuilds what serialization drops.
    ///
    /// Only the raw source text is persisted, so this resizes the value and
    /// compiled-formula vectors back to match it -- restoring the
    /// same-length invariant -- and marks everything dirty. Call it after
    /// deserializing, before [`Sheet::commit`].
    pub fn setup_after_deserialization(&mut self) {
        for col in &mut self.columns {
            col.rebuild_after_load();
        }
        self.mark_all_dirty();
    }

    /// Every sheet a formula on this one could refer to -- this sheet first,
    /// then the rest of `context` -- as the name-to-id lookup table that
    /// `compile_formula` resolves references against.
    ///
    /// "Tables" here means sheets, not [`ExcelTable`]s.
    ///
    /// [`ExcelTable`]: crate::core::table::ExcelTable
    pub(crate) fn get_all_tables_for_compilation(&self, context: Option<&Context>) -> Vec<Sheet> {
        let mut list = vec![self.clone()];
        let mut seen = std::collections::HashSet::new();
        seen.insert(self.id);
        if let Some(ctx) = context {
            for t in ctx.sheets.values() {
                if !seen.contains(&t.id) {
                    seen.insert(t.id);
                    list.push((*t).clone());
                }
            }
        }
        list
    }

    /// Queues every cell for recomputation on the next [`Sheet::commit`].
    ///
    /// This is how cross-sheet staleness is handled: `WorkbookManager` cannot
    /// tell which cells a remote edit reached, so it marks whole sheets.
    pub fn mark_all_dirty(&mut self) {
        for col in &mut self.columns {
            col.dirty_indices.clear();
            col.dirty_indices.extend(0..col.src.len());
        }
    }

    /// Commit all changed src items with a context for sheet lookups
    pub fn commit(&mut self, context: Option<&Context>) -> Result<HashSet<CellRef>, EngineError> {
        let mut queue: VecDeque<CellRef> = VecDeque::new();
        let mut queue_set: HashSet<CellRef> = HashSet::new();
        let mut updated_cells: HashSet<CellRef> = HashSet::new();

        // 1. Collect initial dirty cells
        for (col_idx, col_data) in self.columns.iter_mut().enumerate() {
            for row_idx in &col_data.dirty_indices {
                let cell = CellRef::new(*row_idx, col_idx);
                queue.push_back(cell);
                queue_set.insert(cell);
                updated_cells.insert(cell);
            }
            col_data.dirty_indices.clear();
        }

        let initial_queue_len = queue.len();
        if initial_queue_len == 0 {
            return Ok(updated_cells);
        }

        let start_commit = Instant::now();
        log::info!(
            "Sheet '{}' commit starting for {} dirty cells",
            self.name,
            initial_queue_len
        );
        let max_ops = 10000.max(initial_queue_len * 3);
        let mut ops = 0;

        let mut tables_for_compilation = self.get_all_tables_for_compilation(context);
        let mut last_log_time = Instant::now();

        while let Some(cell_ref) = queue.pop_front() {
            queue_set.remove(&cell_ref);
            ops += 1;
            if ops > max_ops {
                println!("Circular dependency or too many updates detected");
                break;
            }

            if ops % 50000 == 0 {
                log::info!(
                    "Sheet '{}' commit progress: {}/{} cells processed ({:.2?})",
                    self.name,
                    ops,
                    initial_queue_len,
                    last_log_time.elapsed()
                );
                last_log_time = Instant::now();
            }

            let cell_type_hint = self
                .columns
                .get(cell_ref.col)
                .and_then(|c| c.cell_types.get(cell_ref.row).copied())
                .unwrap_or(CellType::Auto);

            // `Some(code)` when this cell's literal was recognized as a date;
            // applied below, since detecting it only has `&self`.
            let mut detected_num_format: Option<String> = None;
            let (result, new_deps, compiled_to_cache, final_cell_type) = {
                let src = self.get_src_str_ref(&cell_ref).unwrap_or("");
                if cell_type_hint == CellType::String {
                    let val = if src.starts_with('"') && src.ends_with('"') && src.len() >= 2 {
                        src[1..src.len() - 1].to_string()
                    } else {
                        src.to_string()
                    };
                    (ResultData::String(val), vec![], None, CellType::String)
                } else if !src.starts_with('=') && cell_type_hint != CellType::Formula {
                    // Numeric text is trimmed before it is parsed, because
                    // that is what entering it does: a cell given `"  3  "`
                    // holds the *number* 3, in Excel and (measured through
                    // `fuzz/fuzz_vba.py`, where a macro assigned exactly that
                    // string) through VBA's `Range.Value` as well.
                    let (res, c_type) = if let Some(stripped) = src.strip_prefix('\'') {
                        (ResultData::String(stripped.to_string()), CellType::String)
                    } else if src.is_empty() {
                        (ResultData::None, CellType::Empty)
                    } else if src.starts_with('"') && src.ends_with('"') && src.len() >= 2 {
                        (
                            ResultData::String(src[1..src.len() - 1].to_string()),
                            CellType::String,
                        )
                    } else if let Ok(i) = src.trim().parse::<i64>() {
                        (ResultData::Integer(i), CellType::Number)
                    } else if let Ok(f) = src.trim().parse::<f64>()
                        // Rust's `f64::from_str` accepts "inf" and "NaN";
                        // Excel has neither, and reports #NUM! for both. A
                        // non-finite literal here would otherwise become a
                        // Float that no formula could have produced -- found
                        // by `fuzz/fuzz_vba.py`, where a macro assigned
                        // `-2.5 ^ 1000` to a cell and this stored `-inf`
                        // where Excel stored `#NUM!`.
                        && f.is_finite()
                    {
                        (ResultData::Float(f), CellType::Number)
                    } else if crate::core::engine::result_data::is_excel_error_code(src) {
                        // Typing an error value into a cell produces the
                        // error, not the text. See `is_excel_error_code`.
                        (ResultData::Error(src.to_uppercase()), CellType::Error)
                    } else if src.eq_ignore_ascii_case("true") {
                        (ResultData::Boolean(true), CellType::Boolean)
                    } else if src.eq_ignore_ascii_case("false") {
                        (ResultData::Boolean(false), CellType::Boolean)
                    } else if let Some((date, format)) = crate::core::date::parse_date_with_locale(
                        src.trim_matches(' '),
                        &self.locale,
                    ) {
                        // Excel stores a typed date as a serial and remembers
                        // the notation as the cell's number format, so `6/22/26`
                        // is a number that happens to display as a date.
                        detected_num_format = Some(format.to_format_code());
                        (
                            ResultData::Float(crate::core::date::date_to_excel_serial(date)),
                            CellType::Number,
                        )
                    } else if let Some(f) = crate::core::date_fn::parse_time_fraction(src) {
                        // Same for a typed time: it is a numeric day fraction,
                        // not text. The time number format is not represented
                        // yet, so for now only the value semantics are matched.
                        (ResultData::Float(f), CellType::Number)
                    } else {
                        (ResultData::String(src.to_string()), CellType::String)
                    };
                    (res, vec![], None, c_type)
                } else {
                    let compiled =
                        crate::core::parser::compile_formula(src, &tables_for_compilation);
                    let eval_src =
                        crate::core::parser::serialize_formula(&compiled, &tables_for_compilation);
                    let (res, deps) = match self.eval_with_row(
                        &eval_src,
                        context,
                        Some(cell_ref.row),
                        Some(cell_ref.col),
                    ) {
                        Ok(r) => r,
                        Err(e) => (ResultData::Error(e.to_string()), vec![]),
                    };
                    let final_res = if let ResultData::None = res {
                        ResultData::Float(0.0)
                    } else {
                        res
                    };
                    (final_res, deps, Some(compiled), CellType::Formula)
                }
            };

            // If a leading apostrophe was used to force string type, strip it from src
            if let Some(src_str) = self.get_src_str_ref(&cell_ref)
                && let Some(stripped) = src_str.strip_prefix('\'')
            {
                let stripped_str = stripped.to_string();
                if let Some(col) = self.columns.get_mut(cell_ref.col)
                    && cell_ref.row < col.src.len()
                {
                    col.src[cell_ref.row] = stripped_str;
                }
            }

            if let Some(col) = self.columns.get_mut(cell_ref.col)
                && cell_ref.row < col.compiled_src.len()
            {
                col.compiled_src[cell_ref.row] = compiled_to_cache.unwrap_or_default();
            }

            if let Some(old_deps) = self.dependencies_rev.remove(&cell_ref) {
                for provider in old_deps {
                    if let Some(dependents) = self.dependencies.get_mut(&provider) {
                        dependents.remove(&cell_ref);
                    }
                }
            }

            // Add new dependencies (only if not empty to save map allocations)
            if !new_deps.is_empty() {
                let mut new_deps_set = HashSet::new();
                for provider in new_deps {
                    new_deps_set.insert(provider.clone());
                    self.dependencies
                        .entry(provider)
                        .or_default()
                        .insert(cell_ref);
                }
                self.dependencies_rev.insert(cell_ref, new_deps_set);
            }

            // A recognized date literal carries its notation onto the cell,
            // and a formula that shifts a date by a number inherits that
            // date's -- `=A1+1` on a date displays as the next day, as in
            // Excel, rather than as a bare serial.
            let inherited = if detected_num_format.is_some()
                || !matches!(result, ResultData::Float(_) | ResultData::Integer(_))
            {
                None
            } else {
                self.get_src_str_ref(&cell_ref)
                    .and_then(|src| src.strip_prefix('='))
                    .and_then(|body| crate::core::parser::parse_excel_formula(body).ok())
                    .and_then(|ast| self.inherited_date_format(&ast))
            };
            // An explicit format the user (or an imported worksheet) already
            // set wins, so re-entering a date does not clobber it.
            if let Some(code) = detected_num_format.or(inherited) {
                let existing = self
                    .get_cell_style(cell_ref.row, cell_ref.col)
                    .and_then(|s| s.num_format.clone());
                if existing.is_none() {
                    self.update_cell_style(cell_ref.row, cell_ref.col, |style| {
                        style.num_format = Some(code);
                    });
                }
            }

            if let Some(col) = self.columns.get_mut(cell_ref.col)
                && cell_ref.row < col.data.len()
            {
                col.cell_types[cell_ref.row] = final_cell_type;
                col.data.set(cell_ref.row, result.clone());
                updated_cells.insert(cell_ref);
            }
            if let Some(comp_sheet) = tables_for_compilation
                .iter_mut()
                .find(|s| s.name == self.name)
                && let Some(col) = comp_sheet.columns.get_mut(cell_ref.col)
                && cell_ref.row < col.data.len()
            {
                col.cell_types[cell_ref.row] = final_cell_type;
                col.data.set(cell_ref.row, result);
            }

            // Propagate to dependents (Local only)
            // If this cell changed, we need to notify anyone who depends on THIS cell (locally).
            // A local dependency is represented as Dependency::Local(this_cell).
            let local_dep_key = Dependency::Local(cell_ref);
            if let Some(dependents) = self.dependencies.get(&local_dep_key) {
                for dependent in dependents {
                    if !queue_set.contains(dependent) {
                        queue.push_back(*dependent);
                        queue_set.insert(*dependent);
                    }
                }
            }

            // Also notify anyone who depends on the whole COLUMN
            let local_col_dep_key = Dependency::LocalColumn(cell_ref.col);
            if let Some(dependents) = self.dependencies.get(&local_col_dep_key) {
                for dependent in dependents {
                    if !queue_set.contains(dependent) {
                        queue.push_back(*dependent);
                        queue_set.insert(*dependent);
                    }
                }
            }
        }
        if initial_queue_len > 0 {
            log::info!(
                "Sheet '{}' commit finished. Processed {} cell updates. Total time: {:.2?}",
                self.name,
                ops,
                start_commit.elapsed()
            );
        }
        Ok(updated_cells)
    }

    /// Evaluates cell source text without storing it, as
    /// [`Sheet::eval`] does, but from the point of view of `(row, col)`.
    ///
    /// The position is what makes relative constructs work -- a structured
    /// reference like `[@Amount]` means "this row", so it needs to know which
    /// row is asking. Pass `None` for both when there is no anchor.
    ///
    /// # Errors
    ///
    /// Returns an [`EngineError`] if the formula cannot be parsed. An *Excel*
    /// error is not a Rust error: `=1/0` succeeds, returning
    /// `ResultData::Error("#DIV/0!")`.
    pub fn eval_with_row(
        &self,
        input: &str,
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
    ) -> Result<(ResultData, Vec<Dependency>), EngineError> {
        if input.is_empty() {
            return Ok((ResultData::None, vec![]));
        }
        if let Some(formula) = input.strip_prefix('=') {
            self.eval_excel(formula, context, row, col)
        } else {
            if let Ok(i) = input.parse::<i64>() {
                Ok((ResultData::Integer(i), vec![]))
            } else if let Ok(f) = input.parse::<f64>() {
                Ok((ResultData::Float(f), vec![]))
            } else if let Ok(b) = input.parse::<bool>() {
                Ok((ResultData::Boolean(b), vec![]))
            } else {
                Ok((ResultData::String(input.to_string()), vec![]))
            }
        }
    }

    /// Evaluates cell source text against this sheet without storing it,
    /// returning the value and the references it read.
    ///
    /// Text with a leading `=` is a formula; anything else is parsed as a
    /// literal. `context` supplies the other sheets, and is required for a
    /// cross-sheet reference to resolve.
    ///
    /// # Errors
    ///
    /// Returns an [`EngineError`] if the formula cannot be parsed. An *Excel*
    /// error is not a Rust error: `=1/0` succeeds, returning
    /// `ResultData::Error("#DIV/0!")`.
    pub fn eval(
        &self,
        input: &str,
        context: Option<&Context>,
    ) -> Result<(ResultData, Vec<Dependency>), EngineError> {
        self.eval_with_row(input, context, None, None)
    }

    fn eval_excel(
        &self,
        code: &str,
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
    ) -> Result<(ResultData, Vec<Dependency>), EngineError> {
        let ast = crate::core::parser::parse_excel_formula(code)
            .map_err(|e| EngineError::EvalError(EvalError::UnknownFunction(e)))?;

        let mut deps = Vec::new();
        let result = match self.evaluate_ast(&ast, context, row, col, &mut deps, &LetScope::Empty) {
            Ok(r) => r,
            Err(EngineError::EvalError(EvalError::UnknownFunction(err_str)))
                if err_str.starts_with('#') =>
            {
                ResultData::Error(err_str)
            }
            Err(e) => return Err(e),
        };
        Ok((result, deps))
    }

    fn evaluate_ast(
        &self,
        ast: &crate::core::parser::Expr,
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::SheetSection;
        use crate::core::parser::Expr;
        use crate::core::parser::Op;

        match ast {
            Expr::Number(n) => Ok(ResultData::Float(*n)),
            Expr::String(s) => Ok(ResultData::String(s.clone())),
            Expr::Boolean(b) => Ok(ResultData::Boolean(*b)),
            Expr::Error(code) => Ok(ResultData::Error(code.to_string())),
            Expr::Identifier(name) => match scope.get(name) {
                Some(val) => Ok(val.clone()),
                None => Ok(ResultData::Error("#NAME?".to_string())),
            },
            Expr::StructuredRef {
                sheet,
                column,
                is_this_row,
                section,
            } => {
                let ref_name = match sheet {
                    Some(name) => name.clone(),
                    None => self.name.clone(),
                };

                // The leading name of a structured reference is first looked up as
                // a real Excel Table (an `ExcelTable` may live on any sheet in
                // scope, and is scoped to its own row/column range). If no such
                // table exists, fall back to the legacy behavior of treating the
                // name as a sheet name and the whole sheet as an implicit table --
                // this keeps existing formulas working for sheets that don't
                // define any explicit table.
                let mut found: Option<(&Sheet, &crate::core::table::ExcelTable)> =
                    self.find_table(&ref_name).map(|t| (self, t));
                if found.is_none()
                    && let Some(ctx) = context
                {
                    for s in ctx.sheets.values() {
                        if let Some(t) = s.find_table(&ref_name) {
                            found = Some((s, t));
                            break;
                        }
                    }
                }

                if let Some((table_sheet, excel_table)) = found {
                    let is_self = table_sheet.name == self.name;
                    let sheet_name = table_sheet.name.clone();

                    // (local index within the table, absolute sheet column index)
                    let col_indices: Vec<(usize, usize)> = if let Some(col_name) = column {
                        let local = excel_table.local_column_index(col_name).ok_or_else(|| {
                            EngineError::EvalError(EvalError::UnknownFunction(format!(
                                "Column not found: {}",
                                col_name
                            )))
                        })?;
                        vec![(local, excel_table.start_col + local)]
                    } else {
                        (0..excel_table.columns.len())
                            .map(|local| (local, excel_table.start_col + local))
                            .collect()
                    };
                    let is_whole_table = column.is_none();

                    match section {
                        SheetSection::Headers => {
                            let names: Vec<ResultData> = col_indices
                                .iter()
                                .map(|&(local, _)| {
                                    ResultData::String(
                                        excel_table.columns.get(local).cloned().unwrap_or_default(),
                                    )
                                })
                                .collect();
                            if is_whole_table {
                                Ok(ResultData::List(names))
                            } else {
                                Ok(names.into_iter().next().unwrap_or(ResultData::None))
                            }
                        }
                        SheetSection::Totals => {
                            if let Some(totals_row) = excel_table.totals_row() {
                                let mut results = Vec::new();
                                for &(_, col_idx) in &col_indices {
                                    let cell_ref = CellRef::new(totals_row, col_idx);
                                    if is_self {
                                        deps.push(Dependency::Local(cell_ref));
                                    } else {
                                        deps.push(Dependency::Remote {
                                            sheet: sheet_name.clone(),
                                            cell: cell_ref,
                                        });
                                    }
                                    results.push(table_sheet.get_result_data(&cell_ref));
                                }
                                if is_whole_table {
                                    Ok(ResultData::List(results))
                                } else {
                                    Ok(results.into_iter().next().unwrap_or(ResultData::None))
                                }
                            } else {
                                Ok(ResultData::None)
                            }
                        }
                        SheetSection::Data | SheetSection::All => {
                            if *is_this_row {
                                let r = row.ok_or_else(|| {
                                    EngineError::EvalError(EvalError::UnknownFunction(
                                        "This row reference cannot be evaluated without row context"
                                            .to_string(),
                                    ))
                                })?;
                                let mut results = Vec::new();
                                for &(_, col_idx) in &col_indices {
                                    let cell_ref = CellRef::new(r, col_idx);
                                    if is_self {
                                        deps.push(Dependency::Local(cell_ref));
                                    } else {
                                        deps.push(Dependency::Remote {
                                            sheet: sheet_name.clone(),
                                            cell: cell_ref,
                                        });
                                    }
                                    results.push(table_sheet.get_result_data(&cell_ref));
                                }
                                if is_whole_table {
                                    Ok(ResultData::List(results))
                                } else {
                                    Ok(results.into_iter().next().unwrap_or(ResultData::None))
                                }
                            } else {
                                // A table's column reference is bounded to
                                // its own data rows, not the whole sheet
                                // column -- so, like any other bounded range
                                // (e.g. A1:A100), each cell in that range
                                // gets its own dependency rather than a
                                // whole-column one. Otherwise a formula
                                // placed in the same column but *outside*
                                // the table (a common layout, since summary
                                // formulas often sit right below or beside
                                // a table) would register a dependency on
                                // its own cell and falsely trip circular-
                                // dependency detection, which real Excel
                                // does not do.
                                let mut results = Vec::new();
                                for &(_, col_idx) in &col_indices {
                                    for r in
                                        excel_table.data_start_row()..=excel_table.data_end_row()
                                    {
                                        let cell_ref = CellRef::new(r, col_idx);
                                        if is_self {
                                            deps.push(Dependency::Local(cell_ref));
                                        } else {
                                            deps.push(Dependency::Remote {
                                                sheet: sheet_name.clone(),
                                                cell: cell_ref,
                                            });
                                        }
                                        results.push(table_sheet.get_result_data(&cell_ref));
                                    }
                                }
                                Ok(ResultData::List(results))
                            }
                        }
                    }
                } else {
                    // Legacy fallback: no explicit ExcelTable found by that name --
                    // resolve `ref_name` as a sheet name and treat the whole sheet
                    // as an implicit table.
                    let sheet_name = ref_name;
                    let is_self = sheet_name == self.name;

                    let target_table = if is_self {
                        self
                    } else if let Some(ctx) = context {
                        if let Some(t) = ctx.sheets.get(&sheet_name) {
                            t
                        } else {
                            return Err(EngineError::EvalError(EvalError::UnknownFunction(
                                format!("Sheet not found: {}", sheet_name),
                            )));
                        }
                    } else {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                            "No context to resolve sheet reference: {}",
                            sheet_name
                        ))));
                    };

                    // `column: None` means the reference spans every column in the
                    // table (e.g. `Table1[#Data]` or `[@]`), rather than a single
                    // named column.
                    let col_indices: Vec<usize> = if let Some(col_name) = column {
                        let pos = target_table
                            .columns
                            .iter()
                            .position(|c| c.name == *col_name)
                            .ok_or_else(|| {
                                EngineError::EvalError(EvalError::UnknownFunction(format!(
                                    "Column not found: {}",
                                    col_name
                                )))
                            })?;
                        vec![pos]
                    } else {
                        (0..target_table.columns.len()).collect()
                    };
                    let is_whole_table = column.is_none();

                    match section {
                        SheetSection::Headers => {
                            let names: Vec<ResultData> = col_indices
                                .iter()
                                .map(|&idx| {
                                    ResultData::String(
                                        target_table
                                            .columns
                                            .get(idx)
                                            .map(|c| c.name.clone())
                                            .unwrap_or_default(),
                                    )
                                })
                                .collect();
                            if is_whole_table {
                                Ok(ResultData::List(names))
                            } else {
                                Ok(names.into_iter().next().unwrap_or(ResultData::None))
                            }
                        }
                        SheetSection::Totals => Ok(ResultData::None),
                        SheetSection::Data | SheetSection::All => {
                            if *is_this_row {
                                let r = row.ok_or_else(|| {
                                    EngineError::EvalError(EvalError::UnknownFunction(
                                        "This row reference cannot be evaluated without row context"
                                            .to_string(),
                                    ))
                                })?;
                                let mut results = Vec::new();
                                for &col_idx in &col_indices {
                                    let cell_ref = CellRef::new(r, col_idx);
                                    if is_self {
                                        deps.push(Dependency::Local(cell_ref));
                                    } else {
                                        deps.push(Dependency::Remote {
                                            sheet: sheet_name.clone(),
                                            cell: cell_ref,
                                        });
                                    }
                                    results.push(target_table.get_result_data(&cell_ref));
                                }
                                if is_whole_table {
                                    Ok(ResultData::List(results))
                                } else {
                                    Ok(results.into_iter().next().unwrap_or(ResultData::None))
                                }
                            } else {
                                let mut results = Vec::new();
                                for &col_idx in &col_indices {
                                    if is_self {
                                        deps.push(Dependency::LocalColumn(col_idx));
                                    } else {
                                        deps.push(Dependency::RemoteColumn {
                                            sheet: sheet_name.clone(),
                                            col: col_idx,
                                        });
                                    }
                                    for r in 0..target_table.row_count() {
                                        let cell_ref = CellRef::new(r, col_idx);
                                        results.push(target_table.get_result_data(&cell_ref));
                                    }
                                }
                                Ok(ResultData::List(results))
                            }
                        }
                    }
                }
            }
            Expr::CellRef {
                sheet,
                row: r_val,
                col,
                ..
            } => {
                let cell_ref = CellRef::new(*r_val, *col);
                let is_self = match sheet {
                    Some(name) => name == &self.name,
                    None => true,
                };

                if is_self {
                    deps.push(Dependency::Local(cell_ref));
                    Ok(self.get_result_data(&cell_ref))
                } else {
                    let name = sheet.as_ref().unwrap().clone();
                    deps.push(Dependency::Remote {
                        sheet: name.clone(),
                        cell: cell_ref,
                    });

                    if let Some(ctx) = context {
                        if let Some(t) = ctx.sheets.get(&name) {
                            Ok(t.get_result_data(&cell_ref))
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                "Sheet not found: {}",
                                name
                            ))))
                        }
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "No context to resolve sheet reference".to_string(),
                        )))
                    }
                }
            }
            Expr::RangeRef {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => {
                let is_self = match sheet {
                    Some(name) => name == &self.name,
                    None => true,
                };

                let target_sheet = if is_self {
                    Some(self)
                } else {
                    context.and_then(|ctx| ctx.sheets.get(sheet.as_ref().unwrap()).copied())
                };

                let actual_end_row = if *end_row == usize::MAX {
                    target_sheet
                        .map(|t| t.row_count().saturating_sub(1))
                        .unwrap_or(0)
                } else {
                    *end_row
                };
                let actual_end_col = if *end_col == usize::MAX {
                    target_sheet
                        .map(|t| t.col_count().saturating_sub(1))
                        .unwrap_or(0)
                } else {
                    *end_col
                };

                let is_col_range = *end_row == usize::MAX;

                // A whole-column range's dependency is scoped to the
                // *column*, not each individual cell, but the loop below
                // still visits every (row, col) pair -- without tracking
                // which columns this range has already registered, the
                // `deps.contains` scan below (needed for correctness
                // against whatever `deps` already held coming in) would
                // run once per *cell* instead of once per *column*, i.e.
                // O(width * height * deps.len()) instead of O(width *
                // deps.len()). For a wide range (e.g. `=C:LL`, 322
                // columns) evaluated repeatedly (e.g. inside a self-
                // referential formula bounded by commit()'s max_ops),
                // that quadratic-in-width blowup was
                // the difference between finishing in under a second and
                // taking tens of seconds to minutes.
                let mut seen_col_deps: HashSet<usize> = HashSet::new();

                let mut results = Vec::new();
                for r in *start_row..=actual_end_row {
                    for c in *start_col..=actual_end_col {
                        let cell_ref = CellRef::new(r, c);
                        if is_self {
                            if is_col_range {
                                if seen_col_deps.insert(c) {
                                    let col_dep = Dependency::LocalColumn(c);
                                    if !deps.contains(&col_dep) {
                                        deps.push(col_dep);
                                    }
                                }
                            } else {
                                deps.push(Dependency::Local(cell_ref));
                            }
                            // A range that includes the very cell this
                            // formula lives in (most commonly a bare,
                            // unaggregated whole-column/whole-row range
                            // like `=C:P` sitting inside columns C..P)
                            // must not read that cell's own currently
                            // stored value back into itself: on every
                            // recompute the stored value *is* this List,
                            // so reading it back would nest a List inside
                            // itself one level deeper each pass (unbounded
                            // growth).
                            // Blank matches this engine's existing
                            // convention for an unresolvable self-read
                            // elsewhere (e.g. ISBLANK(GET(...)) on an
                            // empty cell).
                            if row == Some(r) && col == Some(c) {
                                results.push(ResultData::None);
                            } else {
                                results.push(self.get_result_data(&cell_ref));
                            }
                        } else {
                            let name = sheet.as_ref().unwrap().clone();
                            if is_col_range {
                                if seen_col_deps.insert(c) {
                                    let col_dep = Dependency::RemoteColumn {
                                        sheet: name.clone(),
                                        col: c,
                                    };
                                    if !deps.contains(&col_dep) {
                                        deps.push(col_dep);
                                    }
                                }
                            } else {
                                deps.push(Dependency::Remote {
                                    sheet: name.clone(),
                                    cell: cell_ref,
                                });
                            }
                            if let Some(ctx) = context {
                                if let Some(t) = ctx.sheets.get(&name) {
                                    results.push(t.get_result_data(&cell_ref));
                                } else {
                                    return Err(EngineError::EvalError(
                                        EvalError::UnknownFunction(format!(
                                            "Sheet not found: {}",
                                            name
                                        )),
                                    ));
                                }
                            } else {
                                return Err(EngineError::EvalError(EvalError::UnknownFunction(
                                    "No context to resolve sheet reference".to_string(),
                                )));
                            }
                        }
                    }
                }
                Ok(ResultData::List(results))
            }
            Expr::List(list) => {
                let mut results = Vec::new();
                for item in list {
                    results.push(self.evaluate_ast(item, context, row, col, deps, scope)?);
                }
                Ok(ResultData::List(results))
            }
            Expr::UnaryOp { op, expr } => {
                let val = self.evaluate_ast(expr, context, row, col, deps, scope)?;
                match op {
                    Op::Sub => match val {
                        ResultData::Float(f) => Ok(ResultData::Float(-f)),
                        ResultData::Integer(i) => Ok(ResultData::Integer(-i)),
                        _ => Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "Unary minus expects number".to_string(),
                        ))),
                    },
                    _ => Ok(val),
                }
            }
            Expr::BinaryOp { op, left, right } => {
                let l_val = self.evaluate_ast(left, context, row, col, deps, scope)?;

                match op {
                    Op::Eq | Op::Ne | Op::Lt | Op::Gt | Op::Le | Op::Ge => {
                        if let ResultData::Error(_) = &l_val {
                            return Ok(l_val);
                        }
                        let r_val = self.evaluate_ast(right, context, row, col, deps, scope)?;
                        if let ResultData::Error(_) = &r_val {
                            return Ok(r_val);
                        }
                        let ord = Self::compare_excel_values(&l_val, &r_val);
                        let b = match op {
                            Op::Eq => ord.is_eq(),
                            Op::Ne => !ord.is_eq(),
                            Op::Lt => ord.is_lt(),
                            Op::Gt => ord.is_gt(),
                            Op::Le => ord.is_le(),
                            Op::Ge => ord.is_ge(),
                            _ => unreachable!(),
                        };
                        Ok(ResultData::Boolean(b))
                    }
                    _ => {
                        if let ResultData::Error(_) = &l_val {
                            return Ok(l_val);
                        }
                        let lf = match self.to_f64(&l_val) {
                            Some(f) => f,
                            None => return Ok(ResultData::Error("#VALUE!".to_string())),
                        };
                        let r_val = self.evaluate_ast(right, context, row, col, deps, scope)?;
                        if let ResultData::Error(_) = &r_val {
                            return Ok(r_val);
                        }
                        let rf = match self.to_f64(&r_val) {
                            Some(f) => f,
                            None => return Ok(ResultData::Error("#VALUE!".to_string())),
                        };
                        match op {
                            Op::Add => Ok(ResultData::Float(lf + rf)),
                            Op::Sub => Ok(ResultData::Float(lf - rf)),
                            Op::Mul => Ok(ResultData::Float(lf * rf)),
                            Op::Div => {
                                if rf == 0.0 {
                                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                                }
                                Ok(ResultData::Float(lf / rf))
                            }
                            Op::Exp => {
                                if lf == 0.0 && rf == 0.0 {
                                    return Ok(ResultData::Error("#NUM!".to_string()));
                                }
                                if lf == 0.0 && rf < 0.0 {
                                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                                }
                                if lf < 0.0 {
                                    if rf.fract() != 0.0 || rf.abs() > 1e6 {
                                        return Ok(ResultData::Error("#NUM!".to_string()));
                                    }
                                    let res = lf.powi(rf as i32);
                                    if res.is_nan() || res.is_infinite() {
                                        return Ok(ResultData::Error("#NUM!".to_string()));
                                    }
                                    return Ok(ResultData::Float(res));
                                }
                                let res = lf.powf(rf);
                                if res.is_nan() || res.is_infinite() {
                                    return Ok(ResultData::Error("#NUM!".to_string()));
                                }
                                Ok(ResultData::Float(res))
                            }
                            _ => unreachable!(),
                        }
                    }
                }
            }
            Expr::FunctionCall { name, args } => {
                self.evaluate_function(name, args, context, row, col, deps, scope)
            }
        }
    }

    fn excel_type_rank(val: &ResultData) -> u8 {
        match val {
            ResultData::None => 0,
            ResultData::Integer(_) | ResultData::Float(_) => 1,
            ResultData::String(_) => 2,
            ResultData::Boolean(_) => 3,
            _ => 4,
        }
    }

    fn compare_excel_values(l: &ResultData, r: &ResultData) -> std::cmp::Ordering {
        match (l, r) {
            (ResultData::None, ResultData::None) => return std::cmp::Ordering::Equal,
            (ResultData::None, ResultData::Integer(b)) => {
                return 0.0
                    .partial_cmp(&(*b as f64))
                    .unwrap_or(std::cmp::Ordering::Equal);
            }
            (ResultData::None, ResultData::Float(b)) => {
                return 0.0.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal);
            }
            (ResultData::Integer(a), ResultData::None) => {
                return (*a as f64)
                    .partial_cmp(&0.0)
                    .unwrap_or(std::cmp::Ordering::Equal);
            }
            (ResultData::Float(a), ResultData::None) => {
                return a.partial_cmp(&0.0).unwrap_or(std::cmp::Ordering::Equal);
            }
            (ResultData::None, ResultData::String(b)) => {
                return "".cmp(b.to_lowercase().as_str());
            }
            (ResultData::String(a), ResultData::None) => {
                return a.to_lowercase().as_str().cmp("");
            }
            (ResultData::None, ResultData::Boolean(b)) => {
                return false.cmp(b);
            }
            (ResultData::Boolean(a), ResultData::None) => {
                return a.cmp(&false);
            }
            _ => {}
        }

        let rank_l = Self::excel_type_rank(l);
        let rank_r = Self::excel_type_rank(r);
        if rank_l != rank_r {
            return rank_l.cmp(&rank_r);
        }
        match (l, r) {
            (ResultData::Integer(a), ResultData::Integer(b)) => a.cmp(b),
            (ResultData::Float(a), ResultData::Float(b)) => {
                a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal)
            }
            (ResultData::Integer(a), ResultData::Float(b)) => (*a as f64)
                .partial_cmp(b)
                .unwrap_or(std::cmp::Ordering::Equal),
            (ResultData::Float(a), ResultData::Integer(b)) => a
                .partial_cmp(&(*b as f64))
                .unwrap_or(std::cmp::Ordering::Equal),
            (ResultData::Boolean(a), ResultData::Boolean(b)) => a.cmp(b),
            (ResultData::String(a), ResultData::String(b)) => Self::compare_excel_strings(a, b),
            _ => std::cmp::Ordering::Equal,
        }
    }

    /// `SORT`/`SORTBY`-specific comparator: Microsoft documents that both
    /// functions always place blank cells last, regardless of ascending
    /// vs. descending order -- unlike `compare_excel_values`'s general
    /// blank-coerces-to-0/""/false rule (correct for comparison operators,
    /// MATCH, etc.), which would otherwise rank a blank ahead of every
    /// negative number once descending order reverses the comparison.
    /// E.g. `SORT({-215.8,,-100,-240.97,-88},1,-1)` puts the blank last.
    fn sort_compare_blanks_last(
        l: &ResultData,
        r: &ResultData,
        sort_order: f64,
    ) -> std::cmp::Ordering {
        match (matches!(l, ResultData::None), matches!(r, ResultData::None)) {
            (true, true) => std::cmp::Ordering::Equal,
            (true, false) => std::cmp::Ordering::Greater,
            (false, true) => std::cmp::Ordering::Less,
            (false, false) => {
                let ord = Self::compare_excel_values(l, r);
                if sort_order < 0.0 { ord.reverse() } else { ord }
            }
        }
    }

    fn is_excel_number_str(s: &str) -> bool {
        let s = s.trim();
        if s.is_empty() {
            return false;
        }
        let bytes = s.as_bytes();
        let first = bytes[0];
        if first == b'e' || first == b'E' {
            return false;
        }
        if (first == b'+' || first == b'-') && bytes.len() > 1 {
            let second = bytes[1];
            if second == b'e' || second == b'E' {
                return false;
            }
        }
        true
    }

    fn compare_excel_strings(a: &str, b: &str) -> std::cmp::Ordering {
        let char_weight = |ch: char| -> u32 {
            match ch {
                ' ' => 0,
                '_' => 1,
                '-' => 2,
                ',' => 3,
                ';' => 4,
                ':' => 5,
                '!' => 6,
                '?' => 7,
                '.' => 8,
                '\'' => 9,
                '"' => 10,
                '(' => 11,
                ')' => 12,
                '[' => 13,
                ']' => 14,
                '{' => 15,
                '}' => 16,
                '@' => 17,
                '*' => 18,
                '/' => 19,
                '\\' => 20,
                '&' => 21,
                '#' => 22,
                '%' => 23,
                '`' => 24,
                '^' => 25,
                '+' => 26,
                '<' => 27,
                '=' => 28,
                '>' => 29,
                '|' => 30,
                '~' => 31,
                '$' => 32,
                '0'..='9' => 33 + (ch as u32 - '0' as u32),
                'A'..='Z' => 43 + (ch as u32 - 'A' as u32),
                'a'..='z' => 43 + (ch as u32 - 'a' as u32),
                _ => ch
                    .to_lowercase()
                    .next()
                    .map(|c| c as u32 + 200)
                    .unwrap_or(ch as u32 + 200),
            }
        };

        for (ca, cb) in a.chars().zip(b.chars()) {
            let wa = char_weight(ca);
            let wb = char_weight(cb);
            if wa != wb {
                return wa.cmp(&wb);
            }
        }
        a.len().cmp(&b.len())
    }

    /// Snaps a float to its 15-significant-digit rounding when the two are
    /// within floating-point noise of each other, so accumulated error does
    /// not leak into a result Excel would show as exact.
    ///
    /// Left alone if the rounding moves the value by more than that, and for
    /// zero and non-finite values.
    pub(crate) fn clean_float(val: f64) -> f64 {
        if val == 0.0 || !val.is_finite() {
            return val;
        }
        let abs_val = val.abs();
        let exp = abs_val.log10().floor() as i32;
        let factor = 10.0f64.powi(15 - 1 - exp);
        if factor.is_finite() && factor != 0.0 {
            let rounded = (val * factor).round() / factor;
            if (val - rounded).abs() <= 1e-14 * abs_val {
                return rounded;
            }
        }
        val
    }

    /// Coerces a value to a number the way an Excel arithmetic operator does:
    /// a blank is 0, a boolean is 0 or 1, and text is converted if it reads as
    /// a number or a date (a date becoming its serial).
    ///
    /// `None` for text that is not numeric and for every other value,
    /// including errors -- callers turn that into `#VALUE!`.
    ///
    /// Not every function coerces this way; the stricter families reject text
    /// and booleans outright.
    pub(crate) fn to_f64(&self, val: &ResultData) -> Option<f64> {
        match val {
            ResultData::None => Some(0.0),
            ResultData::Float(f) => Some(*f),
            ResultData::Integer(i) => Some(*i as f64),
            ResultData::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
            ResultData::String(s) => {
                let s_trim = s.trim();
                if Self::is_excel_number_str(s_trim) {
                    if let Ok(f) = s_trim.parse::<f64>() {
                        return Some(f);
                    }
                    if let Some((date, _)) =
                        crate::core::date::parse_date_with_locale(s_trim, &self.locale)
                    {
                        return Some(crate::core::date::date_to_excel_serial(date));
                    }
                    None
                } else if let Some((date, _)) =
                    crate::core::date::parse_date_with_locale(s_trim, &self.locale)
                {
                    Some(crate::core::date::date_to_excel_serial(date))
                } else {
                    None
                }
            }
            _ => None,
        }
    }

    fn to_f64_arg(&self, arg_opt: Option<&ResultData>, fn_name: &str) -> Result<f64, EngineError> {
        let val = arg_opt.ok_or_else(|| {
            EngineError::EvalError(EvalError::UnknownFunction(format!(
                "{} requires argument",
                fn_name
            )))
        })?;
        if let ResultData::Error(e) = val {
            return Err(EngineError::EvalError(EvalError::UnknownFunction(
                e.clone(),
            )));
        }
        self.to_f64(val).ok_or_else(|| {
            EngineError::EvalError(EvalError::UnknownFunction("#VALUE!".to_string()))
        })
    }

    fn find_error_in_args(args: &[ResultData]) -> Option<ResultData> {
        for arg in args {
            match arg {
                ResultData::Error(_) => return Some(arg.clone()),
                ResultData::List(list) => {
                    if let Some(err) = Self::find_error_in_args(list) {
                        return Some(err);
                    }
                }
                _ => {}
            }
        }
        None
    }

    fn check_arg_errors(&self, args: &[ResultData], is_direct: &[bool]) -> Option<ResultData> {
        for (i, arg) in args.iter().enumerate() {
            match arg {
                ResultData::Error(_) => return Some(arg.clone()),
                ResultData::List(list) => {
                    if let Some(err) = self.check_arg_errors(list, &[]) {
                        return Some(err);
                    }
                }
                ResultData::String(_)
                    if is_direct.get(i).copied().unwrap_or(false) && self.to_f64(arg).is_none() =>
                {
                    return Some(ResultData::Error("#VALUE!".to_string()));
                }
                _ => {}
            }
        }
        None
    }

    fn sum_helper(&self, arg: &ResultData, is_direct: bool) -> f64 {
        match arg {
            ResultData::Float(f) => *f,
            ResultData::Integer(i) => *i as f64,
            ResultData::Boolean(b) => {
                if is_direct {
                    if *b { 1.0 } else { 0.0 }
                } else {
                    0.0
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    self.to_f64(arg).unwrap_or(0.0)
                } else {
                    0.0
                }
            }
            ResultData::List(list) => {
                let mut sum = 0.0;
                for item in list {
                    sum += self.sum_helper(item, false);
                }
                sum
            }
            _ => 0.0,
        }
    }

    /// Flattens a single argument (which may be a range/array `List`) into
    /// an ordered `Vec<f64>` for the financial functions that take a
    /// cashflow series (`NPV`, `IRR`, `MIRR`, `XNPV`, `XIRR`, `FVSCHEDULE`).
    /// Mirrors `sum_helper`'s convention: booleans/text only count when
    /// passed directly (not through a range).
    fn flatten_finance_numbers(&self, arg: &ResultData, is_direct: bool) -> Vec<f64> {
        match arg {
            ResultData::Float(f) => vec![*f],
            ResultData::Integer(i) => vec![*i as f64],
            ResultData::Boolean(b) => {
                if is_direct {
                    vec![if *b { 1.0 } else { 0.0 }]
                } else {
                    vec![]
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    self.to_f64(arg).into_iter().collect()
                } else {
                    vec![]
                }
            }
            ResultData::List(list) => list
                .iter()
                .flat_map(|v| self.flatten_finance_numbers(v, false))
                .collect(),
            _ => vec![],
        }
    }

    fn flatten_stat_numbers(&self, arg: &ResultData, is_direct: bool) -> Vec<f64> {
        match arg {
            ResultData::Float(f) => vec![*f],
            ResultData::Integer(i) => vec![*i as f64],
            ResultData::Boolean(b) => {
                if is_direct {
                    vec![if *b { 1.0 } else { 0.0 }]
                } else {
                    vec![]
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    self.to_f64(arg).into_iter().collect()
                } else {
                    vec![]
                }
            }
            ResultData::List(list) => list
                .iter()
                .flat_map(|v| self.flatten_stat_numbers(v, false))
                .collect(),
            _ => vec![],
        }
    }

    /// Flattens one argument positionally: `Some(n)` for a numeric cell,
    /// `None` for anything real Excel excludes from a paired statistical
    /// calculation (text, boolean, blank). Unlike flatten_stat_numbers,
    /// excluded cells still occupy a slot, so two ranges of the same
    /// shape always produce vectors of the same length and element `i` of
    /// one still lines up with element `i` of the other.
    fn flatten_positional(
        &self,
        arg: &ResultData,
        out: &mut Vec<Option<f64>>,
        first_err: &mut Option<String>,
    ) {
        match arg {
            ResultData::List(items) => {
                for item in items {
                    self.flatten_positional(item, out, first_err);
                }
            }
            ResultData::Float(f) => out.push(Some(*f)),
            ResultData::Integer(i) => out.push(Some(*i as f64)),
            ResultData::Error(e) => {
                if first_err.is_none() {
                    *first_err = Some(e.clone());
                }
                out.push(None);
            }
            _ => out.push(None),
        }
    }

    fn positional_numbers(
        &self,
        arg: Option<&ResultData>,
        first_err: &mut Option<String>,
    ) -> Vec<Option<f64>> {
        let mut out = Vec::new();
        if let Some(a) = arg {
            self.flatten_positional(a, &mut out, first_err);
        }
        out
    }

    /// Excel's paired statistical functions (CORREL/PEARSON/COVAR/
    /// COVARIANCE.P/COVARIANCE.S/SLOPE/INTERCEPT/RSQ/STEYX/FORECAST/
    /// TREND/LINEST/GROWTH/LOGEST/T.TEST/SUMX2PY2/SUMXMY2/SUMX2MY2/PROB)
    /// compare the two ranges' *raw* element counts first -- a mismatch
    /// is #N/A regardless of content -- and then drop every (x, y) pair
    /// where either side is non-numeric, keeping what survives aligned.
    ///
    /// Verified directly against real Excel: `COVAR(A1:A4, B1:B4)` with
    /// one text cell in B returns exactly the value of the 3-element
    /// ranges with that whole pair physically removed, and the same holds
    /// for SLOPE/INTERCEPT/RSQ/PEARSON/STEYX/FORECAST/T.TEST/SUMX*.
    /// Booleans and blanks are excluded the same way text is.
    ///
    /// This is deliberately *not* the same as flattening each side
    /// independently (what flatten_stat_numbers does): dropping a
    /// non-numeric from only one side shifts every later element against
    /// its partner, silently correlating the wrong values together.
    /// F.TEST/FTEST is the exception that genuinely does want independent
    /// per-array flattening -- it compares two samples' variances and
    /// doesn't require equal sizes at all (confirmed against real Excel:
    /// `FTEST(4-cell-with-text, ...)` equals `FTEST(full-4-cell, ...)`
    /// against the 3-cell survivor, i.e. each side shrinks on its own).
    fn pair_and_filter(
        xs_raw: Vec<Option<f64>>,
        ys_raw: Vec<Option<f64>>,
    ) -> Result<(Vec<f64>, Vec<f64>), String> {
        if xs_raw.len() != ys_raw.len() {
            return Err("#N/A".to_string());
        }
        let mut xs = Vec::with_capacity(xs_raw.len());
        let mut ys = Vec::with_capacity(ys_raw.len());
        for (x, y) in xs_raw.into_iter().zip(ys_raw) {
            if let (Some(x), Some(y)) = (x, y) {
                xs.push(x);
                ys.push(y);
            }
        }
        Ok((xs, ys))
    }

    /// pair_and_filter over two argument slots.
    fn paired_args(
        &self,
        x_arg: Option<&ResultData>,
        y_arg: Option<&ResultData>,
    ) -> Result<(Vec<f64>, Vec<f64>), String> {
        // The size check has to come *before* propagating any error cell
        // sitting inside either range: real Excel reports #N/A for two
        // differently-sized ranges even when one of them contains a live
        // error (confirmed by probing `CORREL` over a 4-cell and a 3-cell
        // range whose second range held a #DIV/0!, which answers #N/A).
        // These functions are therefore excluded from the generic
        // "any error in an argument short-circuits the call" pre-pass, and
        // re-raise the error here only once the shapes agree.
        // A *scalar* operand carrying an error propagates before any
        // shape logic runs: Excel resolves a 1x1 reference to a plain
        // value first, and an error value in an ordinary operand position
        // short-circuits the call. So `SUMX2PY2(A1:A4, P1:P1)` with a
        // #DIV/0! in P1 is #DIV/0!, even though the two operands are
        // differently sized.
        //
        // An error inside a *multi-cell* range does not get that
        // treatment -- there the size check wins, and
        // `SUMX2PY2(A1:A4, N1:N3)` with an error inside N1:N3 is #N/A.
        // Both confirmed against real Excel, and consistently across
        // CORREL/SLOPE/STEYX/SUMX2PY2.
        for arg in [x_arg, y_arg].into_iter().flatten() {
            // A one-cell *range* evaluates to a one-element List rather
            // than a bare scalar, so both spellings have to be unwrapped
            // here -- matching only the bare form let
            // `STEYX(H6:H6, F2:H2)` report the shape mismatch (#N/A)
            // instead of the error sitting in H6.
            let scalar = match arg {
                ResultData::List(items) if items.len() == 1 => &items[0],
                other => other,
            };
            if let ResultData::Error(e) = scalar {
                return Err(e.clone());
            }
            // A one-cell operand that is *empty* isn't a one-element array,
            // it's a missing operand: Excel answers #VALUE! rather than the
            // #N/A a shape mismatch would give. Note this is specifically
            // about blankness -- a one-cell operand holding text or a
            // boolean still reports #N/A, so it can't be folded into the
            // general non-numeric handling (all three confirmed against
            // real Excel with CORREL against a 4-cell range).
            if Self::is_empty_scalar_operand(arg) {
                return Err("#VALUE!".to_string());
            }
        }
        let mut first_err = None;
        let xs_raw = self.positional_numbers(x_arg, &mut first_err);
        let ys_raw = self.positional_numbers(y_arg, &mut first_err);
        if xs_raw.len() != ys_raw.len() {
            return Err("#N/A".to_string());
        }
        if let Some(e) = first_err {
            return Err(e);
        }
        Self::pair_and_filter(xs_raw, ys_raw)
    }

    /// Like flatten_stat_numbers, but errors instead of silently dropping
    /// a cell real Excel won't accept. Excel's array/matrix-argument
    /// functions don't ignore text the way SUM/AVERAGE-style aggregates
    /// do -- one bad cell makes the whole call #VALUE!.
    ///
    /// `blanks` selects between the three blank-handling behaviours real
    /// Excel actually exhibits here, each established by probing it
    /// directly:
    ///  - `BlankPolicy::Zero` (MULTINOMIAL): a blank counts as 0 and the
    ///    call still succeeds -- `MULTINOMIAL(3, <blank>)` is 1, the
    ///    blank participating as a zero.
    ///  - `BlankPolicy::Skip` (GCD/LCM/SERIESSUM): a blank is dropped
    ///    outright rather than zero-filled. `LCM(1, <blank>)` is 1 (as if
    ///    `LCM(1)`), not `LCM(1, 0)` = 0. For SERIESSUM this also shifts
    ///    every later coefficient down a power:
    ///    `SERIESSUM(0.5, 0, 2, {4, 6, <blank>, 8})` is 6.0 -- exactly the
    ///    3-coefficient answer -- not the 5.625 a zero in that slot gives.
    ///  - `BlankPolicy::Reject` (LINEST/TREND/GROWTH/LOGEST/MMULT): blanks
    ///    are #VALUE! too, same as text and booleans.
    ///
    /// `coerce_text` selects separately whether a numeric-looking string
    /// is accepted (converted the same way `to_f64` would) or rejected
    /// outright as #VALUE! -- this does *not* track the blank policy,
    /// since GCD/LCM (`Skip`) coerce text (`GCD("12", 8)` = 4) while
    /// SERIESSUM (also `Skip`) does not (`SERIESSUM(1.49, 1, 2,
    /// {<blank>, "2", 27, -35})` is #VALUE! in real Excel, not the number
    /// the coerced "2" would give -- fuzz/fuzz_excel.py seed 107768).
    /// Booleans are always rejected regardless of either policy -- `GCD(TRUE,
    /// 8)` is #VALUE! -- which is why this can't just fall through to
    /// `to_f64`, the lenient coercion used for scalar arguments.
    fn flatten_strict_inner(
        &self,
        arg: &ResultData,
        blanks: BlankPolicy,
        coerce_text: bool,
        out: &mut Vec<f64>,
    ) -> Result<(), String> {
        match arg {
            ResultData::List(items) => {
                for item in items {
                    self.flatten_strict_inner(item, blanks, coerce_text, out)?;
                }
                Ok(())
            }
            ResultData::Error(e) => Err(e.clone()),
            ResultData::Float(f) => {
                out.push(*f);
                Ok(())
            }
            ResultData::Integer(i) => {
                out.push(*i as f64);
                Ok(())
            }
            ResultData::None => match blanks {
                BlankPolicy::Zero => {
                    out.push(0.0);
                    Ok(())
                }
                BlankPolicy::Skip => Ok(()),
                BlankPolicy::Reject => Err("#VALUE!".to_string()),
            },
            // Numeric text is coerced, non-numeric text is not, when
            // `coerce_text` is set: real Excel gives GCD("12", 8) = 4,
            // LCM("4", 6) = 12 and MULTINOMIAL("3", 2) = 10, while
            // GCD("x", 8) is #VALUE! either way. Booleans stay rejected
            // regardless -- GCD(TRUE, 8) is #VALUE! -- which is why this
            // can't just fall through to `to_f64`.
            ResultData::String(_) if coerce_text => match self.to_f64(arg) {
                Some(f) => {
                    out.push(f);
                    Ok(())
                }
                None => Err("#VALUE!".to_string()),
            },
            _ => Err("#VALUE!".to_string()),
        }
    }

    fn flatten_strict_numbers(&self, arg: &ResultData) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        self.flatten_strict_inner(arg, BlankPolicy::Zero, true, &mut out)?;
        Ok(out)
    }

    /// flatten_strict_numbers with blanks dropped rather than zero-filled,
    /// for GCD/LCM (which also coerce numeric text, like MULTINOMIAL).
    fn flatten_skipping_blanks(&self, arg: Option<&ResultData>) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        if let Some(a) = arg {
            self.flatten_strict_inner(a, BlankPolicy::Skip, true, &mut out)?;
        }
        Ok(out)
    }

    /// Like `flatten_skipping_blanks`, but a numeric-looking string is
    /// #VALUE! rather than coerced -- SERIESSUM's coefficients, unlike
    /// GCD/LCM's operands, don't accept text at all (measured:
    /// `SERIESSUM(1.49, 1, 2, {<blank>, "2", 27, -35})` is #VALUE! in real
    /// Excel, not the value the coerced "2" would give -- see
    /// `flatten_strict_inner`'s doc comment).
    fn flatten_skipping_blanks_no_text_coercion(
        &self,
        arg: Option<&ResultData>,
    ) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        if let Some(a) = arg {
            self.flatten_strict_inner(a, BlankPolicy::Skip, false, &mut out)?;
        }
        Ok(out)
    }

    /// flatten_strict_numbers with the stricter "a blank is also #VALUE!"
    /// rule the regression-array and matrix functions use.
    fn flatten_numbers_only(&self, arg: &ResultData) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        self.flatten_strict_inner(arg, BlankPolicy::Reject, false, &mut out)?;
        Ok(out)
    }

    /// The value of one cell of a SUMIF/AVERAGEIF/MAXIFS/MINIFS-style
    /// *aggregate* range. Only a real number counts: Excel silently skips
    /// text and booleans in the range being summed/averaged/compared
    /// (confirmed directly -- `SUMIF` over a range holding
    /// `{100, TRUE, 200, "txt", 300}` is 600, and MAXIFS over the same
    /// range is 300, not the boolean coerced to 1). Using the lenient
    /// `to_f64` here instead folded `TRUE` in as a 1, which both shifted
    /// sums/averages and could win a MAX/MIN outright.
    fn aggregate_range_number(val: &ResultData) -> Option<f64> {
        match val {
            ResultData::Float(f) => Some(*f),
            ResultData::Integer(i) => Some(*i as f64),
            _ => None,
        }
    }

    fn flatten_numbers_only_arg(&self, arg: Option<&ResultData>) -> Result<Vec<f64>, String> {
        match arg {
            Some(a) => self.flatten_numbers_only(a),
            None => Ok(vec![]),
        }
    }

    /// `flatten_stat_numbers` across an argument list, applying Excel's rule
    /// for text supplied *directly* as an argument: it is coerced if it
    /// looks numeric, and is `#VALUE!` if it does not. Text reached through
    /// a reference is skipped instead, which is what `flatten_stat_numbers`
    /// already does on its own.
    ///
    /// The split matters because silently skipping uncoercible direct text
    /// turns a wrong formula into a plausible number: `DEVSQ("abc",3,4,5)`
    /// answered 2 (the spread of the remaining three) where Excel answers
    /// `#VALUE!`. Verified against real Excel for SUM, AVERAGE, DEVSQ,
    /// STDEV, VAR, MEDIAN, MAX, MIN, PRODUCT, SUMSQ, GEOMEAN, AVEDEV, SKEW
    /// and KURT. COUNT is the deliberate exception -- it never errors, it
    /// just doesn't count what it can't read -- and does not call this.
    fn flatten_args_stat_numbers(
        &self,
        args: &[ResultData],
        is_direct: &[bool],
    ) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        for (i, arg) in args.iter().enumerate() {
            let direct = is_direct.get(i).copied().unwrap_or(false);
            if direct && matches!(arg, ResultData::String(_)) && self.to_f64(arg).is_none() {
                return Err("#VALUE!".to_string());
            }
            out.extend(self.flatten_stat_numbers(arg, direct));
        }
        Ok(out)
    }

    /// Flatten arguments for the `*A` statistical family (AVERAGEA, MAXA,
    /// MINA, STDEVA, STDEVPA, VARA, VARPA), which count text and booleans
    /// rather than skipping them.
    ///
    /// Text is where the family gets interesting, and the rule depends on
    /// *how* the text arrived. Inside a reference it counts as 0, which is
    /// the documented behaviour everyone knows. Passed directly as an
    /// argument it is coerced instead, and a value that will not coerce is
    /// an error rather than a zero. Against real Excel, with A1 holding the
    /// text "12":
    ///
    /// ```text
    /// AVERAGEA(A1, 3)     = 1.5        text in a reference counts as 0
    /// AVERAGEA("12", 3)   = 7.5        direct text is coerced
    /// AVERAGEA("abc", 3)  = #VALUE!    ... and must coerce
    /// ```
    fn flatten_stat_numbers_a(
        &self,
        arg: &ResultData,
        is_direct: bool,
    ) -> Result<Vec<f64>, String> {
        Ok(match arg {
            ResultData::Float(f) => vec![*f],
            ResultData::Integer(i) => vec![*i as f64],
            ResultData::Boolean(b) => vec![if *b { 1.0 } else { 0.0 }],
            ResultData::String(_) => {
                if is_direct {
                    match self.to_f64(arg) {
                        Some(f) => vec![f],
                        None => return Err("#VALUE!".to_string()),
                    }
                } else {
                    vec![0.0]
                }
            }
            ResultData::Error(e) => return Err(e.clone()),
            // Anything nested is a reference, never a direct argument.
            ResultData::List(list) => {
                let mut out = Vec::new();
                for v in list {
                    out.extend(self.flatten_stat_numbers_a(v, false)?);
                }
                out
            }
            ResultData::None => vec![],
            _ => vec![0.0],
        })
    }

    /// `flatten_stat_numbers_a` over a whole argument list, using the
    /// caller's per-argument direct/reference classification.
    fn flatten_args_stat_numbers_a(
        &self,
        args: &[ResultData],
        is_direct: &[bool],
    ) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        for (i, arg) in args.iter().enumerate() {
            out.extend(
                self.flatten_stat_numbers_a(arg, is_direct.get(i).copied().unwrap_or(false))?,
            );
        }
        Ok(out)
    }

    fn extract_matrix(&self, arg: &ResultData) -> Vec<Vec<f64>> {
        match arg {
            ResultData::List(list) => {
                let mut rows = Vec::new();
                for item in list {
                    match item {
                        ResultData::List(sub_list) => {
                            let row: Vec<f64> =
                                sub_list.iter().flat_map(|v| self.to_f64(v)).collect();
                            if !row.is_empty() {
                                rows.push(row);
                            }
                        }
                        _ => {
                            if let Some(f) = self.to_f64(item) {
                                rows.push(vec![f]);
                            }
                        }
                    }
                }
                rows
            }
            _ => vec![],
        }
    }

    /// Reshapes a range argument's flat evaluated list back into a 2D
    /// row-major matrix using the *reference's* own width.
    ///
    /// A plain rectangular range like `F1:G2` evaluates to a flat
    /// `List` of 4 scalars with no nesting, so extract_matrix (which can
    /// only treat a nested `List` as a row) turned it into a 4x1 column
    /// instead of a 2x2 square -- and every matrix function then reported
    /// #VALUE! on a perfectly valid square range. MMULT already
    /// reconstructed its operands' shapes from the argument expression
    /// this way; this shares that logic with MDETERM/MINVERSE.
    fn matrix_from_arg(
        &self,
        expr: &crate::core::parser::Expr,
        value: &ResultData,
    ) -> Vec<Vec<f64>> {
        // A list of lists already carries its own shape.
        if let ResultData::List(items) = value
            && items.iter().any(|i| matches!(i, ResultData::List(_)))
        {
            return self.extract_matrix(value);
        }
        // Only real numbers: the matrix functions reject text, booleans
        // and blanks alike (all confirmed #VALUE! against real Excel), so
        // a cell that isn't a number collapses the whole matrix rather
        // than being coerced by to_f64.
        fn plain(v: &ResultData) -> Option<f64> {
            match v {
                ResultData::Float(f) => Some(*f),
                ResultData::Integer(i) => Some(*i as f64),
                _ => None,
            }
        }
        let items: Vec<&ResultData> = match value {
            ResultData::List(items) => items.iter().collect(),
            other => vec![other],
        };
        if items.iter().any(|v| plain(v).is_none()) {
            return Vec::new();
        }
        let flat: Vec<f64> = items.iter().filter_map(|v| plain(v)).collect();
        let cols = match Self::range_bounds(expr) {
            Some((_, _, start_col, _, end_col)) => end_col.saturating_sub(start_col) + 1,
            None => flat.len().max(1),
        };
        if cols == 0 || !flat.len().is_multiple_of(cols) {
            return self.extract_matrix(value);
        }
        flat.chunks(cols).map(|c| c.to_vec()).collect()
    }

    /// An optional numeric argument. An *absent* argument falls back to
    /// `default`, but one that is present and non-numeric is #VALUE! --
    /// the `.and_then(to_f64).unwrap_or(default)` shape used in places
    /// conflates the two, so e.g. `LOG(3.14, "E")` quietly computed
    /// base-10 instead of erroring.
    /// `#DIV/0!` when either operand of a paired sum contains no numeric
    /// value at all.
    ///
    /// This is *not* the same as "no pair survived exclusion", which is
    /// simply 0. Real Excel, with a column [53, TRUE] against a row
    /// [TRUE, -10]: every pair is dropped (each holds a boolean), yet the
    /// answer is 0 rather than an error, because each range does hold a
    /// number. Swap in a range that is entirely text or entirely booleans
    /// and it becomes #DIV/0!.
    ///
    /// Fitted against eleven real-Excel cases spanning text, booleans and
    /// mixtures, at one, two and three elements per range.
    fn paired_sum_has_no_numbers(&self, arg: Option<&ResultData>) -> bool {
        let mut ignored = None;
        let slots = self.positional_numbers(arg, &mut ignored);
        slots.iter().all(|v| v.is_none())
    }

    /// True when an argument is a *single-cell* operand that is empty.
    ///
    /// Excel treats that as a missing operand and answers #VALUE!, rather
    /// than as a one-element array of nothing. The distinction is
    /// specifically about a single cell: `SUMPRODUCT(<one blank cell>)` is
    /// #VALUE! while `SUMPRODUCT(<two blank cells>)` is 0, and
    /// `SUMPRODUCT(-50, <blank>)` is #VALUE! too. Same for MULTINOMIAL and
    /// the paired statistical functions.
    ///
    /// A one-cell range evaluates to a one-element `List` rather than a
    /// bare scalar, so both spellings have to be unwrapped. Note this is
    /// about blankness only -- a one-cell operand holding text or a
    /// boolean behaves differently again.
    fn is_empty_scalar_operand(arg: &ResultData) -> bool {
        let scalar = match arg {
            ResultData::List(items) if items.len() == 1 => &items[0],
            other => other,
        };
        matches!(scalar, ResultData::None)
    }

    /// True when the first argument is a boolean and the function is one
    /// of the few that refuse them.
    ///
    /// Excel's numeric coercion is not uniform here. SQRT, FACT, SIGN,
    /// INT, EXP, ROMAN and most of their neighbours take TRUE as 1
    /// without complaint, but ERF, ERFC, FACTDOUBLE and SQRTPI all answer
    /// #VALUE! -- verified one function at a time against real Excel,
    /// because the split does not follow from anything about the
    /// functions themselves.
    fn first_arg_is_boolean(args: &[ResultData]) -> bool {
        matches!(args.first(), Some(ResultData::Boolean(_)))
    }

    fn opt_f64_arg(&self, args: &[ResultData], i: usize, default: f64) -> Result<f64, EngineError> {
        match args.get(i) {
            None => Ok(default),
            // A supplied-but-blank argument is 0, not the default. Excel
            // draws that line sharply: LOG(1, <blank>) is #NUM! because the
            // base is 0, while LOG(1) uses base 10 and returns 0. Same for
            // LEFT("abcd", <blank>) = "" and MROUND(10, <blank>) = 0.
            Some(ResultData::None) => Ok(0.0),
            Some(ResultData::Error(e)) => Err(EngineError::EvalError(EvalError::UnknownFunction(
                e.clone(),
            ))),
            Some(v) => self.to_f64(v).ok_or_else(|| {
                EngineError::EvalError(EvalError::UnknownFunction("#VALUE!".to_string()))
            }),
        }
    }

    fn opt_f64(&self, args: &[ResultData], i: usize, default: f64) -> f64 {
        args.get(i).and_then(|v| self.to_f64(v)).unwrap_or(default)
    }

    fn average_helper(&self, arg: &ResultData, is_direct: bool) -> (f64, usize) {
        match arg {
            ResultData::Float(f) => (*f, 1),
            ResultData::Integer(i) => (*i as f64, 1),
            ResultData::Boolean(b) => {
                if is_direct {
                    (if *b { 1.0 } else { 0.0 }, 1)
                } else {
                    (0.0, 0)
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    if let Some(f) = self.to_f64(arg) {
                        (f, 1)
                    } else {
                        (0.0, 0)
                    }
                } else {
                    (0.0, 0)
                }
            }
            ResultData::List(list) => {
                let mut sum = 0.0;
                let mut count = 0;
                for item in list {
                    let (s, c) = self.average_helper(item, false);
                    sum += s;
                    count += c;
                }
                (sum, count)
            }
            _ => (0.0, 0),
        }
    }

    fn count_helper(&self, arg: &ResultData) -> usize {
        match arg {
            ResultData::Float(_) | ResultData::Integer(_) => 1,
            ResultData::List(list) => {
                let mut count = 0;
                for item in list {
                    count += self.count_helper(item);
                }
                count
            }
            _ => 0,
        }
    }

    fn min_helper(&self, arg: &ResultData, is_direct: bool) -> f64 {
        match arg {
            ResultData::Float(f) => *f,
            ResultData::Integer(i) => *i as f64,
            ResultData::Boolean(b) => {
                if is_direct {
                    if *b { 1.0 } else { 0.0 }
                } else {
                    f64::INFINITY
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    self.to_f64(arg).unwrap_or(f64::INFINITY)
                } else {
                    f64::INFINITY
                }
            }
            ResultData::List(list) => {
                let mut min_val = f64::INFINITY;
                for item in list {
                    min_val = min_val.min(self.min_helper(item, false));
                }
                min_val
            }
            _ => f64::INFINITY,
        }
    }

    fn max_helper(&self, arg: &ResultData, is_direct: bool) -> f64 {
        match arg {
            ResultData::Float(f) => *f,
            ResultData::Integer(i) => *i as f64,
            ResultData::Boolean(b) => {
                if is_direct {
                    if *b { 1.0 } else { 0.0 }
                } else {
                    f64::NEG_INFINITY
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    self.to_f64(arg).unwrap_or(f64::NEG_INFINITY)
                } else {
                    f64::NEG_INFINITY
                }
            }
            ResultData::List(list) => {
                let mut max_val = f64::NEG_INFINITY;
                for item in list {
                    max_val = max_val.max(self.max_helper(item, false));
                }
                max_val
            }
            _ => f64::NEG_INFINITY,
        }
    }

    fn concat_helper(&self, arg: &ResultData, out: &mut String) {
        match arg {
            ResultData::List(list) => {
                for item in list {
                    self.concat_helper(item, out);
                }
            }
            other => {
                out.push_str(&other.to_string());
            }
        }
    }

    fn counta_helper(&self, arg: &ResultData) -> usize {
        match arg {
            ResultData::None => 0,
            // COUNTA counts every non-blank value, and the empty string is a
            // value -- Excel counts both a text cell holding "" and a formula
            // that returned "".
            ResultData::List(list) => {
                let mut count = 0;
                for item in list {
                    count += self.counta_helper(item);
                }
                count
            }
            _ => 1,
        }
    }

    fn product_helper(&self, arg: &ResultData, is_direct: bool) -> (f64, bool) {
        match arg {
            ResultData::Float(f) => (*f, true),
            ResultData::Integer(i) => (*i as f64, true),
            ResultData::Boolean(b) => {
                if is_direct {
                    (if *b { 1.0 } else { 0.0 }, true)
                } else {
                    (1.0, false)
                }
            }
            ResultData::String(_) => {
                if is_direct {
                    if let Some(f) = self.to_f64(arg) {
                        (f, true)
                    } else {
                        (1.0, false)
                    }
                } else {
                    (1.0, false)
                }
            }
            ResultData::List(list) => {
                let mut prod = 1.0;
                let mut has_nums = false;
                for item in list {
                    let (p, h) = self.product_helper(item, false);
                    if h {
                        // Raw here; the 15-significant-digit snap belongs
                        // on the final product only. See the PRODUCT arm.
                        prod *= p;
                        has_nums = true;
                    }
                }
                (prod, has_nums)
            }
            _ => (1.0, false),
        }
    }

    fn to_bool_opt(&self, val: &ResultData) -> Option<bool> {
        match val {
            ResultData::Boolean(b) => Some(*b),
            ResultData::Integer(i) => Some(*i != 0),
            ResultData::Float(f) => Some(*f != 0.0),
            ResultData::String(s) => {
                let s_trim = s.trim();
                if s_trim.eq_ignore_ascii_case("true") {
                    Some(true)
                } else if s_trim.eq_ignore_ascii_case("false") {
                    Some(false)
                } else if let Ok(f) = s_trim.parse::<f64>() {
                    Some(f != 0.0)
                } else {
                    None
                }
            }
            ResultData::None => Some(false),
            _ => None,
        }
    }

    fn to_bool(&self, val: &ResultData) -> bool {
        self.to_bool_opt(val).unwrap_or(false)
    }

    /// Strict "is this a genuine number" check for range-value aggregation
    /// (DCOUNT/DSUM/DAVERAGE/... and friends), as opposed to `to_f64`'s
    /// scalar-arithmetic coercion (which maps blank -> 0 and booleans ->
    /// 1/0). Confirmed against real Excel via the differential fuzzer that
    /// blank and boolean database cells must be excluded here the same
    /// way SUM/COUNT/AVERAGE ignore them within a range argument -- using
    /// `to_f64` instead let a blank row zero out DPRODUCT entirely and
    /// skewed DCOUNT/DSUM/DAVERAGE by counting/summing blanks and
    /// TRUE/FALSE as 0/1.
    fn range_numeric(val: &ResultData) -> Option<f64> {
        match val {
            ResultData::Integer(i) => Some(*i as f64),
            ResultData::Float(f) => Some(*f),
            _ => None,
        }
    }

    /// Exact-match ("match_type 0" / "range_lookup FALSE") comparison for
    /// MATCH/VLOOKUP/HLOOKUP/XLOOKUP.
    ///
    /// A *blank* lookup value is coerced to 0 (Excel's usual empty-cell
    /// coercion) and a blank cell in the searched range never matches
    /// anything. Comparing the two blanks as equal strings instead --
    /// which is what a plain `to_string()` comparison does, since both
    /// render as "" -- made `MATCH(A1, A1:A4, 0)` over a blank A1 report
    /// a hit at position 1 where real Excel reports #N/A.
    fn exact_lookup_matches(lookup: &ResultData, candidate: &ResultData) -> bool {
        if matches!(candidate, ResultData::None) {
            return false;
        }
        let lookup_key = match lookup {
            ResultData::None => "0".to_string(),
            other => other.to_string(),
        };
        candidate.to_string() == lookup_key
    }

    fn wildcard_criteria_matches(pattern: &str, text: &str) -> bool {
        fn rec(pat: &[char], txt: &[char]) -> bool {
            if pat.is_empty() {
                return txt.is_empty();
            }
            match pat[0] {
                '*' => rec(&pat[1..], txt) || (!txt.is_empty() && rec(pat, &txt[1..])),
                '?' => !txt.is_empty() && rec(&pat[1..], &txt[1..]),
                '~' if pat.len() > 1 && matches!(pat[1], '*' | '?' | '~') => {
                    !txt.is_empty() && pat[1] == txt[0] && rec(&pat[2..], &txt[1..])
                }
                ch => !txt.is_empty() && ch == txt[0] && rec(&pat[1..], &txt[1..]),
            }
        }

        let pat = pattern.to_lowercase().chars().collect::<Vec<_>>();
        let txt = text.to_lowercase().chars().collect::<Vec<_>>();
        rec(&pat, &txt)
    }

    fn criteria_text_eq(val: &ResultData, pattern: &str) -> bool {
        let text = val.to_string();
        if pattern.contains('*') || pattern.contains('?') {
            // Excel wildcard criteria are text-pattern matches; numeric and
            // boolean cells are not counted by criteria like "*".
            matches!(val, ResultData::String(_)) && Self::wildcard_criteria_matches(pattern, &text)
        } else {
            text.to_lowercase() == pattern.to_lowercase()
        }
    }

    fn match_criteria(&self, val: &ResultData, criteria: &ResultData) -> bool {
        let crit_str = criteria.to_string();
        if let Some(rest) = crit_str.strip_prefix(">=") {
            // A numeric comparison can only ever be satisfied by a genuine
            // number -- confirmed against real Excel via the differential
            // fuzzer (fuzzing the new database D* functions): blank, text,
            // and boolean cells must all fail ">"/"<" criteria outright,
            // not fall back to comparing as if they were 0.
            let val_f = match Self::range_numeric(val) {
                Some(f) => f,
                None => return false,
            };
            let crit_f = rest.trim().parse::<f64>().unwrap_or(0.0);
            val_f >= crit_f
        } else if let Some(rest) = crit_str.strip_prefix('>') {
            let val_f = match Self::range_numeric(val) {
                Some(f) => f,
                None => return false,
            };
            let crit_f = rest.trim().parse::<f64>().unwrap_or(0.0);
            val_f > crit_f
        } else if let Some(rest) = crit_str.strip_prefix("<>") {
            let remainder = rest.trim();
            !Self::criteria_text_eq(val, remainder)
        } else if let Some(rest) = crit_str.strip_prefix("<=") {
            let val_f = match Self::range_numeric(val) {
                Some(f) => f,
                None => return false,
            };
            let crit_f = rest.trim().parse::<f64>().unwrap_or(0.0);
            val_f <= crit_f
        } else if let Some(rest) = crit_str.strip_prefix('<') {
            let val_f = match Self::range_numeric(val) {
                Some(f) => f,
                None => return false,
            };
            let crit_f = rest.trim().parse::<f64>().unwrap_or(0.0);
            val_f < crit_f
        } else if let Some(rest) = crit_str.strip_prefix('=') {
            let remainder = rest.trim();
            Self::criteria_text_eq(val, remainder)
        } else {
            Self::criteria_text_eq(val, &crit_str)
        }
    }

    /// Resolves an argument `Expr` to its raw `(sheet, start_row, start_col,
    /// end_row, end_col)` range bounds, for functions (like the database
    /// `D*` family below) that need genuine 2D shape and can't work off the
    /// pre-flattened `ResultData::List` every other argument already went
    /// through in `evaluated_args`.
    fn range_bounds(
        expr: &crate::core::parser::Expr,
    ) -> Option<(Option<String>, usize, usize, usize, usize)> {
        use crate::core::parser::Expr;
        match expr {
            Expr::RangeRef {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => Some((sheet.clone(), *start_row, *start_col, *end_row, *end_col)),
            Expr::CellRef {
                sheet, row, col, ..
            } => Some((sheet.clone(), *row, *col, *row, *col)),
            _ => None,
        }
    }

    /// Reads a range's cells into a row-major grid, resolving a whole-column
    /// range's `end_row` sentinel and cross-sheet references via `context`.
    /// Materializing into an owned `Vec<Vec<ResultData>>` (rather than
    /// keeping a live `&Sheet` around) sidesteps the local-vs-remote
    /// lifetime split for the rest of the database-function logic, and
    /// database/criteria ranges are small enough that this is cheap.
    fn materialize_range(
        &self,
        sheet_opt: &Option<String>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        context: Option<&Context>,
    ) -> Option<Vec<Vec<ResultData>>> {
        let is_self = match sheet_opt {
            Some(name) => name == &self.name,
            None => true,
        };
        let source: &Sheet = if is_self {
            self
        } else {
            context?.sheets.get(sheet_opt.as_ref()?)?
        };
        let actual_end_row = if end_row == usize::MAX {
            source.row_count().saturating_sub(1)
        } else {
            end_row
        };
        let actual_end_col = if end_col == usize::MAX {
            source.col_count().saturating_sub(1)
        } else {
            end_col
        };
        if actual_end_row < start_row || actual_end_col < start_col {
            return Some(Vec::new());
        }
        let mut grid = Vec::with_capacity(actual_end_row - start_row + 1);
        for r in start_row..=actual_end_row {
            let mut row = Vec::with_capacity(actual_end_col - start_col + 1);
            for c in start_col..=actual_end_col {
                row.push(source.get_result_data(&CellRef::new(r, c)));
            }
            grid.push(row);
        }
        Some(grid)
    }

    /// Shared implementation for the 12 database `D*` functions
    /// (DAVERAGE/DCOUNT/DCOUNTA/DGET/DMAX/DMIN/DPRODUCT/DSTDEV/DSTDEVP/
    /// DSUM/DVAR/DVARP): each reduces to "match database rows against the
    /// criteria table, then aggregate one field column of the matches" --
    /// they differ only in which aggregation runs at the end.
    ///
    /// `database`/`criteria` are read from the raw `args` AST nodes (not
    /// `evaluated_args`) specifically to recover real row/column bounds;
    /// `field` (name or 1-based index) still comes from `evaluated_args`
    /// since it's a scalar. Criteria semantics match Excel's: multiple
    /// criteria *rows* are OR'd together, multiple non-blank cells within
    /// one criteria row are AND'd, and a blank criteria cell imposes no
    /// constraint on that field.
    fn evaluate_database_function(
        &self,
        func_name: &str,
        args: &[crate::core::parser::Expr],
        evaluated_args: &[ResultData],
        context: Option<&Context>,
    ) -> Result<ResultData, EngineError> {
        if args.len() < 3 || evaluated_args.len() < 3 {
            return Ok(ResultData::Error("#VALUE!".to_string()));
        }
        let (db_sheet, db_sr, db_sc, db_er, db_ec) = match Self::range_bounds(&args[0]) {
            Some(v) => v,
            None => return Ok(ResultData::Error("#VALUE!".to_string())),
        };
        let (crit_sheet, crit_sr, crit_sc, crit_er, crit_ec) = match Self::range_bounds(&args[2]) {
            Some(v) => v,
            None => return Ok(ResultData::Error("#VALUE!".to_string())),
        };
        let db = match self.materialize_range(&db_sheet, db_sr, db_sc, db_er, db_ec, context) {
            Some(g) => g,
            None => return Ok(ResultData::Error("#REF!".to_string())),
        };
        let crit = match self.materialize_range(
            &crit_sheet,
            crit_sr,
            crit_sc,
            crit_er,
            crit_ec,
            context,
        ) {
            Some(g) => g,
            None => return Ok(ResultData::Error("#REF!".to_string())),
        };
        if db.len() < 2 || crit.len() < 2 {
            return Ok(ResultData::Error("#VALUE!".to_string()));
        }

        let db_headers: Vec<String> = db[0].iter().map(|v| v.to_string()).collect();
        let field_idx: usize = match &evaluated_args[1] {
            ResultData::String(s) => {
                match db_headers.iter().position(|h| h.eq_ignore_ascii_case(s)) {
                    Some(idx) => idx,
                    None => return Ok(ResultData::Error("#VALUE!".to_string())),
                }
            }
            other => match self.to_f64(other) {
                Some(n) if n >= 1.0 && (n as usize) <= db_headers.len() => n as usize - 1,
                _ => return Ok(ResultData::Error("#VALUE!".to_string())),
            },
        };

        let crit_headers: Vec<String> = crit[0].iter().map(|v| v.to_string()).collect();
        let crit_to_db: Vec<Option<usize>> = crit_headers
            .iter()
            .map(|h| db_headers.iter().position(|dh| dh.eq_ignore_ascii_case(h)))
            .collect();

        let mut matched: Vec<ResultData> = Vec::new();
        for row in db.iter().skip(1) {
            let row_matches_any_criteria_row = crit.iter().skip(1).any(|crit_row| {
                crit_row.iter().enumerate().all(|(ci, cell)| {
                    if matches!(cell, ResultData::None) {
                        return true;
                    }
                    match crit_to_db.get(ci).copied().flatten() {
                        Some(db_col) => self.match_criteria(&row[db_col], cell),
                        None => false,
                    }
                })
            });
            if row_matches_any_criteria_row {
                matched.push(row[field_idx].clone());
            }
        }

        match func_name {
            "DGET" => match matched.len() {
                0 => Ok(ResultData::Error("#VALUE!".to_string())),
                1 => Ok(matched.into_iter().next().unwrap()),
                _ => Ok(ResultData::Error("#NUM!".to_string())),
            },
            "DCOUNT" => Ok(ResultData::Float(
                matched
                    .iter()
                    .filter(|v| Self::range_numeric(v).is_some())
                    .count() as f64,
            )),
            "DCOUNTA" => Ok(ResultData::Float(
                matched.iter().map(|v| self.counta_helper(v)).sum::<usize>() as f64,
            )),
            _ => {
                let nums: Vec<f64> = matched.iter().filter_map(Self::range_numeric).collect();
                match func_name {
                    "DSUM" => Ok(ResultData::Float(nums.iter().sum())),
                    "DPRODUCT" => Ok(ResultData::Float(if nums.is_empty() {
                        0.0
                    } else {
                        nums.iter().product()
                    })),
                    "DMAX" => {
                        let m = nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
                        Ok(ResultData::Float(if m.is_finite() { m } else { 0.0 }))
                    }
                    "DMIN" => {
                        let m = nums.iter().cloned().fold(f64::INFINITY, f64::min);
                        Ok(ResultData::Float(if m.is_finite() { m } else { 0.0 }))
                    }
                    "DAVERAGE" => {
                        if nums.is_empty() {
                            Ok(ResultData::Error("#DIV/0!".to_string()))
                        } else {
                            Ok(ResultData::Float(
                                nums.iter().sum::<f64>() / nums.len() as f64,
                            ))
                        }
                    }
                    "DSTDEV" => match crate::core::stats::stdev_s(&nums) {
                        Ok(v) => Ok(ResultData::Float(v)),
                        Err(e) => Ok(ResultData::Error(e)),
                    },
                    "DSTDEVP" => match crate::core::stats::stdev_p(&nums) {
                        Ok(v) => Ok(ResultData::Float(v)),
                        Err(e) => Ok(ResultData::Error(e)),
                    },
                    "DVAR" => match crate::core::stats::var_s(&nums) {
                        Ok(v) => Ok(ResultData::Float(v)),
                        Err(e) => Ok(ResultData::Error(e)),
                    },
                    "DVARP" => match crate::core::stats::var_p(&nums) {
                        Ok(v) => Ok(ResultData::Float(v)),
                        Err(e) => Ok(ResultData::Error(e)),
                    },
                    _ => unreachable!(),
                }
            }
        }
    }

    fn proper(&self, s: &str) -> String {
        // Per Microsoft's own definition, PROPER capitalizes a letter
        // preceded by "any character that is not a letter" -- that
        // includes digits, not just punctuation/spacing, which is why
        // PROPER("123abc") is "123Abc": the digits aren't letters, so the
        // 'a' right after them still counts as the start of a new word.
        let mut c_chars = Vec::new();
        let mut capitalize_next = true;
        for c in s.chars() {
            if c.is_alphabetic() {
                if capitalize_next {
                    c_chars.extend(c.to_uppercase());
                } else {
                    c_chars.extend(c.to_lowercase());
                }
                capitalize_next = false;
            } else {
                c_chars.push(c);
                capitalize_next = true;
            }
        }
        c_chars.into_iter().collect()
    }

    fn get_ymd_hms(&self) -> ((i32, u32, u32), (u32, u32, u32)) {
        let now = web_time::SystemTime::now()
            .duration_since(web_time::SystemTime::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let secs_in_day = 86400;
        let days_since_epoch = (now / secs_in_day) as i32;
        let seconds_of_day = (now % secs_in_day) as u32;

        let hour = seconds_of_day / 3600;
        let minute = (seconds_of_day % 3600) / 60;
        let second = seconds_of_day % 60;

        let era = (if days_since_epoch >= -719468 {
            days_since_epoch + 719468
        } else {
            days_since_epoch + 719468 - 146096
        }) / 146097;
        let doe = (days_since_epoch + 719468 - era * 146097) as u32;
        let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
        let y = (yoe as i32) + era * 400;
        let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
        let mp = (5 * doy + 2) / 153;
        let d = doy - (153 * mp + 2) / 5 + 1;
        let m = if mp < 10 { mp + 3 } else { mp - 9 };
        let year = if m <= 2 { y + 1 } else { y };

        ((year, m, d), (hour, minute, second))
    }

    /// Evaluates Excel's LET(name1, value1, [name2, value2, ...],
    /// calculation). Binds each name/value pair in order -- value2 (and
    /// later pairs, and the final calculation) can reference name1, per
    /// Excel's LET semantics -- by recursing one pair at a time so each
    /// level's scope chain only needs to borrow the *previous* level's
    /// binding rather than mutate a shared map (see `LetScope`).
    fn evaluate_let(
        &self,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::parser::Expr;

        if args.is_empty() || args.len().is_multiple_of(2) {
            // Needs one or more name/value pairs followed by a calculation,
            // i.e. an odd number of arguments overall.
            return Ok(ResultData::Error("#VALUE!".to_string()));
        }
        if args.len() == 1 {
            return self.evaluate_ast(&args[0], context, row, col, deps, scope);
        }

        let name = match &args[0] {
            Expr::Identifier(n) => n.as_str(),
            _ => return Ok(ResultData::Error("#VALUE!".to_string())),
        };
        // Excel rejects reusing a name across a single LET's own pairs,
        // rather than letting a later pair silently shadow an earlier one.
        let remaining_pairs = args.len() / 2 - 1;
        let is_duplicate = args[2..]
            .iter()
            .step_by(2)
            .take(remaining_pairs)
            .any(|a| matches!(a, Expr::Identifier(n2) if n2.eq_ignore_ascii_case(name)));
        if is_duplicate {
            return Ok(ResultData::Error("#VALUE!".to_string()));
        }

        let value = self.evaluate_ast(&args[1], context, row, col, deps, scope)?;
        let inner_scope = LetScope::Bound {
            name,
            value: &value,
            parent: scope,
        };
        self.evaluate_let(&args[2..], context, row, col, deps, &inner_scope)
    }

    /// Recognizes `expr` as a `LAMBDA(param1, [param2, ...], body)` call
    /// and, if so, returns its declared parameter names alongside the
    /// (still-unevaluated) body expression. Used by every function below
    /// that takes a lambda argument: the lambda is never evaluated as an
    /// ordinary function call (there's no value a bare LAMBDA could
    /// produce on its own -- see the `#CALC!` case in `evaluate_function`)
    /// -- callers instead inspect its raw AST here and invoke the body
    /// themselves, once per element, via `invoke_lambda`.
    fn extract_lambda(
        expr: &crate::core::parser::Expr,
    ) -> Option<(Vec<&str>, &crate::core::parser::Expr)> {
        use crate::core::parser::Expr;
        let Expr::FunctionCall { name, args } = expr else {
            return None;
        };
        if !name.eq_ignore_ascii_case("LAMBDA") || args.is_empty() {
            return None;
        }
        let (body, params) = args.split_last().unwrap();
        let param_names: Vec<&str> = params
            .iter()
            .filter_map(|p| match p {
                Expr::Identifier(n) => Some(n.as_str()),
                _ => None,
            })
            .collect();
        if param_names.len() != params.len() {
            return None;
        }
        Some((param_names, body))
    }

    /// Evaluates a lambda's body with each of `params` bound (via
    /// `LetScope`) to the corresponding entry of `values`, which must be
    /// the same length. `values` is borrowed rather than consumed so
    /// callers can reuse per-element storage across many invocations
    /// (e.g. MAP calling this once per array element).
    #[allow(clippy::too_many_arguments)]
    fn invoke_lambda<'v>(
        &self,
        params: &[&str],
        values: &'v [ResultData],
        body: &crate::core::parser::Expr,
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'v>,
    ) -> Result<ResultData, EngineError> {
        match (params.split_first(), values.split_first()) {
            (Some((&pname, prest)), Some((vfirst, vrest))) => {
                let inner_scope = LetScope::Bound {
                    name: pname,
                    value: vfirst,
                    parent: scope,
                };
                self.invoke_lambda(prest, vrest, body, context, row, col, deps, &inner_scope)
            }
            _ => self.evaluate_ast(body, context, row, col, deps, scope),
        }
    }

    /// Flattens `expr` (evaluated) into a `Vec<ResultData>`, treating a
    /// scalar as a single-element array -- shared by MAP/REDUCE/SCAN,
    /// which all iterate an "array" argument that might just be one cell.
    fn eval_as_array(
        &self,
        expr: &crate::core::parser::Expr,
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<Vec<ResultData>, EngineError> {
        Ok(
            match self.evaluate_ast(expr, context, row, col, deps, scope)? {
                ResultData::List(items) => Self::flatten_row_major(items).0,
                other => vec![other],
            },
        )
    }

    /// `SEQUENCE`/`MUNIT` (unlike every array-*reshaping* function added
    /// this session) return their 2D result as a genuinely nested
    /// `List(List(row_values), ...)`, one inner list per row, rather than
    /// a flat row-major list -- that's the only place in this engine a
    /// `ResultData::List` still carries real shape. Detect that shape
    /// here and flatten it so downstream consumers (`array_shape`,
    /// `INDEX`, reshape functions) don't need to special-case it; a list
    /// that isn't uniformly nested (the flat convention) passes through
    /// unchanged, with `None` signaling "no shape recovered here".
    fn flatten_row_major(items: Vec<ResultData>) -> (Vec<ResultData>, Option<usize>) {
        if !items.is_empty() && items.iter().all(|v| matches!(v, ResultData::List(_))) {
            let cols = match &items[0] {
                ResultData::List(inner) => inner.len().max(1),
                _ => 1,
            };
            let flat = items
                .into_iter()
                .flat_map(|v| match v {
                    ResultData::List(inner) => inner,
                    other => vec![other],
                })
                .collect();
            (flat, Some(cols))
        } else {
            (items, None)
        }
    }

    /// Infers `(flat_values, num_cols)` for an array-like argument: real
    /// column count from a `RangeRef`/`CellRef` AST node when available,
    /// otherwise treats the flattened result as a single row -- the same
    /// convention `INDEX`'s 3-arg form already uses (see its `num_cols`
    /// match on `args[0]`), since a computed/nested array result (e.g. the
    /// output of another array function) carries no shape of its own in
    /// this engine's flat-`ResultData::List` representation.
    fn array_shape(
        &self,
        expr: &crate::core::parser::Expr,
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<(Vec<ResultData>, usize), EngineError> {
        use crate::core::parser::Expr;
        let items = match self.evaluate_ast(expr, context, row, col, deps, scope)? {
            ResultData::List(items) => items,
            other => vec![other],
        };
        let (flat, nested_cols) = Self::flatten_row_major(items);
        if let Some(cols) = nested_cols {
            return Ok((flat, cols));
        }
        let num_cols = match expr {
            Expr::RangeRef {
                start_col, end_col, ..
            } => (end_col - start_col + 1).max(1),
            Expr::CellRef { .. } => 1,
            Expr::FunctionCall { name, args } => self
                .function_call_cols(name, args, context, row, col, deps, scope)
                .unwrap_or_else(|| flat.len().max(1)),
            _ => flat.len().max(1),
        };
        Ok((flat, num_cols))
    }

    /// Recovers the column count an array-reshaping function call's result
    /// would have, purely from its argument expressions -- needed because
    /// this engine's flat `ResultData::List` carries no shape of its own,
    /// so nesting one of these calls inside another (e.g.
    /// `INDEX(EXPAND(A1:B2,3,3,0),3,3)`) requires recovering the 2D shape.
    /// Returns `None` for anything not in this known set, so callers fall
    /// back to the single-row assumption.
    #[allow(clippy::too_many_arguments)]
    fn function_call_cols(
        &self,
        name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Option<usize> {
        let mut upper = name.to_ascii_uppercase();
        if let Some(rest) = upper.strip_prefix("_XLFN.") {
            upper = rest.to_string();
        }
        if let Some(rest) = upper.strip_prefix("_XLWS.") {
            upper = rest.to_string();
        }
        match upper.as_str() {
            "TRANSPOSE" => {
                let (flat, cols) = self
                    .array_shape(args.first()?, context, row, col, deps, scope)
                    .ok()?;
                Some((flat.len().checked_div(cols).unwrap_or(0)).max(1))
            }
            "HSTACK" => {
                let mut total = 0usize;
                for a in args {
                    total += self.array_shape(a, context, row, col, deps, scope).ok()?.1;
                }
                Some(total)
            }
            "VSTACK" => {
                let mut max_cols = 0usize;
                for a in args {
                    max_cols =
                        max_cols.max(self.array_shape(a, context, row, col, deps, scope).ok()?.1);
                }
                Some(max_cols)
            }
            "CHOOSEROWS" => Some(
                self.array_shape(args.first()?, context, row, col, deps, scope)
                    .ok()?
                    .1,
            ),
            "CHOOSECOLS" => Some(args.len().saturating_sub(1).max(1)),
            "DROP" | "TAKE" => {
                let (_, cols) = self
                    .array_shape(args.first()?, context, row, col, deps, scope)
                    .ok()?;
                let is_take = upper == "TAKE";
                match args.get(2) {
                    Some(e) => {
                        let n = self
                            .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope).ok()?)
                            .unwrap_or(0.0) as isize;
                        let (s, e2) = Self::drop_take_bounds(cols as isize, n, is_take);
                        Some((e2 - s).max(0) as usize)
                    }
                    None => Some(if is_take { cols } else { 0 }),
                }
            }
            "EXPAND" => {
                let (_, cols) = self
                    .array_shape(args.first()?, context, row, col, deps, scope)
                    .ok()?;
                match args.get(2) {
                    Some(e) => Some(
                        self.to_f64(&self.evaluate_ast(e, context, row, col, deps, scope).ok()?)
                            .unwrap_or(cols as f64) as usize,
                    ),
                    None => Some(cols),
                }
            }
            "TOCOL" => Some(1),
            "WRAPROWS" => {
                let n = self
                    .to_f64(
                        &self
                            .evaluate_ast(args.get(1)?, context, row, col, deps, scope)
                            .ok()?,
                    )
                    .unwrap_or(1.0)
                    .max(1.0) as usize;
                Some(n)
            }
            "WRAPCOLS" => {
                let (flat, _) = self
                    .array_shape(args.first()?, context, row, col, deps, scope)
                    .ok()?;
                let wrap = self
                    .to_f64(
                        &self
                            .evaluate_ast(args.get(1)?, context, row, col, deps, scope)
                            .ok()?,
                    )
                    .unwrap_or(1.0)
                    .max(1.0) as usize;
                Some(flat.len().div_ceil(wrap).max(1))
            }
            "UNIQUE" | "SORT" | "SORTBY" | "FILTER" | "TRIMRANGE" => Some(
                self.array_shape(args.first()?, context, row, col, deps, scope)
                    .ok()?
                    .1,
            ),
            "SEQUENCE" => match args.get(1) {
                Some(e) => Some(
                    self.to_f64(&self.evaluate_ast(e, context, row, col, deps, scope).ok()?)
                        .unwrap_or(1.0)
                        .max(1.0) as usize,
                ),
                None => Some(1),
            },
            "MUNIT" => {
                let n = self
                    .to_f64(
                        &self
                            .evaluate_ast(args.first()?, context, row, col, deps, scope)
                            .ok()?,
                    )
                    .unwrap_or(1.0)
                    .max(1.0) as usize;
                Some(n)
            }
            "MAKEARRAY" => match args.get(1) {
                Some(e) => Some(
                    self.to_f64(&self.evaluate_ast(e, context, row, col, deps, scope).ok()?)
                        .unwrap_or(1.0)
                        .max(1.0) as usize,
                ),
                None => Some(1),
            },
            _ => None,
        }
    }

    /// Shared `[start, end)` bound computation for `TAKE`/`DROP`: a
    /// positive count counts from the start, negative from the end;
    /// `is_take` selects which side of that split is kept.
    fn drop_take_bounds(total: isize, n: isize, is_take: bool) -> (isize, isize) {
        let n = n.clamp(-total, total);
        if is_take {
            if n >= 0 { (0, n) } else { (total + n, total) }
        } else if n >= 0 {
            (n, total)
        } else {
            (0, total + n)
        }
    }

    /// Shared implementation for MAP/BYROW/BYCOL/REDUCE/SCAN/MAKEARRAY:
    /// each applies a `LAMBDA` argument to some shape of input (parallel
    /// arrays, rows, columns, an accumulator, or generated row/col
    /// indices) and collects the results -- see each branch for the
    /// specific shape. Dynamic-array results are returned as a flat,
    /// row-major `ResultData::List`, the same convention `SEQUENCE`/
    /// `MUNIT`/etc. already use, since this engine doesn't spill formulas
    /// across cells; callers pull out a single value with `INDEX`.
    #[allow(clippy::too_many_arguments)]
    fn evaluate_lambda_function(
        &self,
        func_name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::parser::Expr;

        match func_name {
            "MAP" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (lambda_expr, array_exprs) = args.split_last().unwrap();
                let Some((params, body)) = Self::extract_lambda(lambda_expr) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                if params.len() != array_exprs.len() {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let arrays: Vec<Vec<ResultData>> = array_exprs
                    .iter()
                    .map(|e| self.eval_as_array(e, context, row, col, deps, scope))
                    .collect::<Result<_, _>>()?;
                let len = arrays.iter().map(|a| a.len()).max().unwrap_or(0);
                let mut results = Vec::with_capacity(len);
                for i in 0..len {
                    let values: Vec<ResultData> = arrays
                        .iter()
                        .map(|a| a.get(i).cloned().unwrap_or(ResultData::None))
                        .collect();
                    results.push(
                        self.invoke_lambda(&params, &values, body, context, row, col, deps, scope)?,
                    );
                }
                Ok(ResultData::List(results))
            }
            "BYROW" | "BYCOL" => {
                if args.len() != 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let Some((params, body)) = Self::extract_lambda(&args[1]) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                if params.len() != 1 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                // Recovers real column count the same way INDEX's 3-arg
                // form does: re-matching the raw AST node, since the
                // already-evaluated array argument is just a flat List.
                let num_cols = match &args[0] {
                    Expr::RangeRef {
                        start_col, end_col, ..
                    } => (end_col - start_col + 1).max(1),
                    _ => 1,
                };
                let flat = self.eval_as_array(&args[0], context, row, col, deps, scope)?;
                let num_rows = if num_cols == 0 {
                    0
                } else {
                    flat.len().div_ceil(num_cols)
                };
                let mut results = Vec::new();
                if func_name == "BYROW" {
                    for r in 0..num_rows {
                        let row_vals: Vec<ResultData> = (0..num_cols)
                            .filter_map(|c| flat.get(r * num_cols + c).cloned())
                            .collect();
                        let arg = vec![ResultData::List(row_vals)];
                        results.push(
                            self.invoke_lambda(
                                &params, &arg, body, context, row, col, deps, scope,
                            )?,
                        );
                    }
                } else {
                    for c in 0..num_cols {
                        let col_vals: Vec<ResultData> = (0..num_rows)
                            .filter_map(|r| flat.get(r * num_cols + c).cloned())
                            .collect();
                        let arg = vec![ResultData::List(col_vals)];
                        results.push(
                            self.invoke_lambda(
                                &params, &arg, body, context, row, col, deps, scope,
                            )?,
                        );
                    }
                }
                Ok(ResultData::List(results))
            }
            "REDUCE" | "SCAN" => {
                // initial_value is optional in real Excel's 3-argument
                // REDUCE/SCAN; since the parser has no dedicated "omitted
                // argument" syntax to express that, this implementation
                // also accepts a plain 2-argument call (array, lambda) as
                // the omitted-initial-value form, seeding the accumulator
                // from the array's own first element and folding over the
                // rest -- rather than only supporting a literal 3rd
                // argument that happens to error out.
                if args.len() != 2 && args.len() != 3 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let lambda_idx = args.len() - 1;
                let array_idx = args.len() - 2;
                let Some((params, body)) = Self::extract_lambda(&args[lambda_idx]) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                if params.len() != 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let array = self.eval_as_array(&args[array_idx], context, row, col, deps, scope)?;
                // SCAN's output has the same length as `array` -- an
                // explicit initial_value (3-arg form) is external to the
                // array and doesn't get its own output entry (every entry
                // is a real fold), whereas the 2-arg fallback's seed *is*
                // the array's own first element, so it does.
                let (mut acc, rest, mut history): (ResultData, &[ResultData], Vec<ResultData>) =
                    if args.len() == 3 {
                        let init = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
                        (init, &array[..], Vec::new())
                    } else {
                        match array.split_first() {
                            Some((first, rest)) => (first.clone(), rest, vec![first.clone()]),
                            None => return Ok(ResultData::Error("#VALUE!".to_string())),
                        }
                    };
                for item in rest {
                    let call_args = [acc.clone(), item.clone()];
                    acc = self
                        .invoke_lambda(&params, &call_args, body, context, row, col, deps, scope)?;
                    history.push(acc.clone());
                }
                if func_name == "REDUCE" {
                    Ok(acc)
                } else {
                    Ok(ResultData::List(history))
                }
            }
            "MAKEARRAY" => {
                if args.len() != 3 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let Some((params, body)) = Self::extract_lambda(&args[2]) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                if params.len() != 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let rows_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
                let cols_val = self.evaluate_ast(&args[1], context, row, col, deps, scope)?;
                let num_rows = self.to_f64(&rows_val).unwrap_or(0.0).max(0.0) as usize;
                let num_cols = self.to_f64(&cols_val).unwrap_or(0.0).max(0.0) as usize;
                let mut results = Vec::with_capacity(num_rows * num_cols);
                for r in 1..=num_rows {
                    for c in 1..=num_cols {
                        let call_args = [ResultData::Float(r as f64), ResultData::Float(c as f64)];
                        results.push(self.invoke_lambda(
                            &params, &call_args, body, context, row, col, deps, scope,
                        )?);
                    }
                }
                Ok(ResultData::List(results))
            }
            _ => unreachable!(),
        }
    }

    /// Minimal A1-notation string parser for `INDIRECT`: `"A1"`,
    /// `"B2:C5"`, `"Sheet1!A1"`, `"Sheet1!A1:B2"`, with optional `$`
    /// absolute markers and an optional `'quoted sheet name'!` prefix.
    /// Deliberately small and local rather than shared with
    /// `visi/src/utils.rs`'s equivalent parser (`parse_cell_ref`/
    /// `parse_range_ref`): `visi-core` cannot depend on the `visi` crate
    /// (the dependency direction is the other way), so this necessarily
    /// duplicates that logic in miniature.
    fn parse_a1_reference(text: &str) -> Option<(Option<String>, usize, usize, usize, usize)> {
        let text = text.trim();
        let (sheet_part, ref_part) = match text.rfind('!') {
            Some(idx) => (Some(&text[..idx]), &text[idx + 1..]),
            None => (None, text),
        };
        let sheet = sheet_part.map(|s| s.trim().trim_matches('\'').to_string());

        fn parse_cell(s: &str) -> Option<(usize, usize)> {
            let s = s.replace('$', "");
            let col_end = s.find(|c: char| c.is_ascii_digit())?;
            let (col_str, row_str) = s.split_at(col_end);
            if col_str.is_empty() || row_str.is_empty() {
                return None;
            }
            let mut col = 0usize;
            for ch in col_str.chars() {
                if !ch.is_ascii_alphabetic() {
                    return None;
                }
                col = col * 26 + (ch.to_ascii_uppercase() as usize - 'A' as usize + 1);
            }
            let row: usize = row_str.parse().ok()?;
            if row == 0 || col == 0 {
                return None;
            }
            Some((row - 1, col - 1))
        }

        if let Some((start, end)) = ref_part.split_once(':') {
            let (r1, c1) = parse_cell(start)?;
            let (r2, c2) = parse_cell(end)?;
            Some((sheet, r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
        } else {
            let (r, c) = parse_cell(ref_part)?;
            Some((sheet, r, c, r, c))
        }
    }

    /// Reads a single cell, registering the appropriate local/remote
    /// dependency -- the same local-vs-remote branch used throughout this
    /// file (see e.g. `evaluate_ast`'s `Expr::CellRef` arm), factored out
    /// since `CELL`/`FORMULATEXT`/`ISFORMULA`/`INDIRECT`/`OFFSET` all need
    /// it for a reference resolved dynamically rather than parsed as an
    /// AST node.
    fn read_cell_with_deps(
        &self,
        sheet_opt: &Option<String>,
        r: usize,
        c: usize,
        context: Option<&Context>,
        deps: &mut Vec<Dependency>,
    ) -> ResultData {
        let is_self = sheet_opt.as_deref().is_none_or(|n| n == self.name);
        if is_self {
            deps.push(Dependency::Local(CellRef::new(r, c)));
            self.get_result_data(&CellRef::new(r, c))
        } else if let Some(ctx) = context {
            let name = sheet_opt.clone().unwrap();
            deps.push(Dependency::Remote {
                sheet: name.clone(),
                cell: CellRef::new(r, c),
            });
            ctx.sheets
                .get(&name)
                .map(|s| s.get_result_data(&CellRef::new(r, c)))
                .unwrap_or(ResultData::None)
        } else {
            ResultData::None
        }
    }

    /// Shared implementation for the range/reference-introspection and
    /// workbook-metadata functions: ROW/ROWS/COLUMN/COLUMNS need the raw
    /// reference's real bounds (not a flattened `evaluated_args` value);
    /// AREAS/ISREF are purely syntactic checks on the argument's AST
    /// shape; FORMULATEXT/ISFORMULA need the cell's raw source text;
    /// INDIRECT/OFFSET build a reference dynamically instead of relying
    /// on one already resolved at parse time; SHEET/SHEETS/CELL/INFO
    /// report workbook/environment metadata.
    #[allow(clippy::too_many_arguments)]
    fn evaluate_range_info_function(
        &self,
        func_name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::parser::Expr;

        match func_name {
            "ROW" => match args.first() {
                // A multi-row reference returns an array of row numbers
                // (one per row spanned), not just the first one -- a
                // single-row reference (including a plain cell, where
                // start_row == end_row) still returns the plain scalar.
                Some(arg) => match Self::range_bounds(arg) {
                    Some((_, start_row, _, end_row, _)) if end_row > start_row => {
                        Ok(ResultData::List(
                            (start_row..=end_row)
                                .map(|r| ResultData::Float((r + 1) as f64))
                                .collect(),
                        ))
                    }
                    Some((_, start_row, _, _, _)) => Ok(ResultData::Float((start_row + 1) as f64)),
                    None => Ok(ResultData::Error("#VALUE!".to_string())),
                },
                None => match row {
                    Some(r) => Ok(ResultData::Float((r + 1) as f64)),
                    None => Ok(ResultData::Error("#VALUE!".to_string())),
                },
            },
            "COLUMN" => match args.first() {
                // Same array-vs-scalar distinction as ROW, but across
                // columns instead of rows.
                Some(arg) => match Self::range_bounds(arg) {
                    Some((_, _, start_col, _, end_col)) if end_col > start_col => {
                        Ok(ResultData::List(
                            (start_col..=end_col)
                                .map(|c| ResultData::Float((c + 1) as f64))
                                .collect(),
                        ))
                    }
                    Some((_, _, start_col, _, _)) => Ok(ResultData::Float((start_col + 1) as f64)),
                    None => Ok(ResultData::Error("#VALUE!".to_string())),
                },
                None => match col {
                    Some(c) => Ok(ResultData::Float((c + 1) as f64)),
                    None => Ok(ResultData::Error("#VALUE!".to_string())),
                },
            },
            "ROWS" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let Some((sheet_opt, start_row, _, end_row, _)) = Self::range_bounds(arg) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let is_self = sheet_opt.as_deref().is_none_or(|n| n == self.name);
                let actual_end_row = if end_row == usize::MAX {
                    if is_self {
                        self.row_count().saturating_sub(1)
                    } else {
                        context
                            .and_then(|ctx| sheet_opt.as_ref().and_then(|n| ctx.sheets.get(n)))
                            .map(|s| s.row_count().saturating_sub(1))
                            .unwrap_or(0)
                    }
                } else {
                    end_row
                };
                Ok(ResultData::Float(
                    (actual_end_row.saturating_sub(start_row) + 1) as f64,
                ))
            }
            "COLUMNS" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                match Self::range_bounds(arg) {
                    Some((_, _, start_col, _, end_col)) => Ok(ResultData::Float(
                        (end_col.saturating_sub(start_col) + 1) as f64,
                    )),
                    None => Ok(ResultData::Error("#VALUE!".to_string())),
                }
            }
            "AREAS" => {
                // This engine's parser has no multi-area (comma-separated
                // union) reference syntax, so every reference is exactly
                // one area.
                if args.is_empty() {
                    Ok(ResultData::Error("#VALUE!".to_string()))
                } else {
                    Ok(ResultData::Float(1.0))
                }
            }
            "ISREF" => Ok(ResultData::Boolean(matches!(
                args.first(),
                Some(Expr::CellRef { .. } | Expr::RangeRef { .. } | Expr::StructuredRef { .. })
            ))),
            "FORMULATEXT" | "ISFORMULA" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let Some((sheet_opt, r, c, _, _)) = Self::range_bounds(arg) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let is_self = sheet_opt.as_deref().is_none_or(|n| n == self.name);
                let src = if is_self {
                    deps.push(Dependency::Local(CellRef::new(r, c)));
                    self.get_src_str(&CellRef::new(r, c))
                } else if let Some(ctx) = context {
                    let name = sheet_opt.unwrap();
                    deps.push(Dependency::Remote {
                        sheet: name.clone(),
                        cell: CellRef::new(r, c),
                    });
                    ctx.sheets
                        .get(&name)
                        .map(|s| s.get_src_str(&CellRef::new(r, c)))
                        .unwrap_or_default()
                } else {
                    String::new()
                };
                let is_formula = src.starts_with('=');
                if func_name == "ISFORMULA" {
                    Ok(ResultData::Boolean(is_formula))
                } else if is_formula {
                    Ok(ResultData::String(src))
                } else {
                    Ok(ResultData::Error("#N/A".to_string()))
                }
            }
            "SHEETS" => Ok(ResultData::Float(
                context.map(|c| c.sheets.len() + 1).unwrap_or(1) as f64,
            )),
            "SHEET" => {
                // With no argument, report this sheet's own ordinal. With
                // a reference argument, report the *referenced* sheet's
                // ordinal (a bare reference with no explicit sheet, e.g.
                // `SHEET(A1)`, means this sheet). Excel also accepts a
                // plain text sheet name, e.g. `SHEET("Sheet2")`.
                let sheet_name = match args.first() {
                    None => Some(self.name.clone()),
                    Some(arg) => match Self::range_bounds(arg) {
                        Some((sheet_opt, ..)) => {
                            Some(sheet_opt.unwrap_or_else(|| self.name.clone()))
                        }
                        None => self
                            .evaluate_ast(arg, context, row, col, deps, scope)
                            .ok()
                            .map(|v| v.to_string()),
                    },
                };

                match sheet_name {
                    Some(name) => {
                        let ordinal = context
                            .and_then(|c| {
                                c.sheet_order
                                    .iter()
                                    .position(|n| n.eq_ignore_ascii_case(&name))
                            })
                            .map(|i| i + 1)
                            // No context (standalone eval outside a
                            // WorkbookManager pass) or the name wasn't
                            // found in workbook order: 1 is the same
                            // approximation this used unconditionally
                            // before.
                            .unwrap_or(1);
                        Ok(ResultData::Float(ordinal as f64))
                    }
                    None => Ok(ResultData::Error("#N/A".to_string())),
                }
            }
            "CELL" => {
                if args.is_empty() {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let info_type = self
                    .evaluate_ast(&args[0], context, row, col, deps, scope)?
                    .to_string()
                    .to_lowercase();
                let bounds = args.get(1).and_then(Self::range_bounds);
                match info_type.as_str() {
                    "row" => match bounds.map(|b| b.1).or(row) {
                        Some(r) => Ok(ResultData::Float((r + 1) as f64)),
                        None => Ok(ResultData::Error("#VALUE!".to_string())),
                    },
                    "col" => match bounds {
                        Some((_, _, c, _, _)) => Ok(ResultData::Float((c + 1) as f64)),
                        None => Ok(ResultData::Error("#VALUE!".to_string())),
                    },
                    "address" => match bounds {
                        Some((_, r, c, _, _)) => Ok(ResultData::String(format!(
                            "${}${}",
                            crate::core::parser::col_idx_to_letters(c),
                            r + 1
                        ))),
                        None => Ok(ResultData::Error("#VALUE!".to_string())),
                    },
                    "contents" => match bounds {
                        Some((sheet_opt, r, c, _, _)) => {
                            Ok(self.read_cell_with_deps(&sheet_opt, r, c, context, deps))
                        }
                        None => Ok(ResultData::Error("#VALUE!".to_string())),
                    },
                    _ => Ok(ResultData::Error("#VALUE!".to_string())),
                }
            }
            "INFO" => {
                if args.is_empty() {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let info_type = self
                    .evaluate_ast(&args[0], context, row, col, deps, scope)?
                    .to_string()
                    .to_lowercase();
                match info_type.as_str() {
                    "numfile" => Ok(ResultData::Float(
                        context.map(|c| c.sheets.len() + 1).unwrap_or(1) as f64,
                    )),
                    "release" => Ok(ResultData::String("16.0".to_string())),
                    "system" => Ok(ResultData::String(
                        if cfg!(target_os = "macos") {
                            "mac"
                        } else {
                            "pcdos"
                        }
                        .to_string(),
                    )),
                    _ => Ok(ResultData::Error("#VALUE!".to_string())),
                }
            }
            "INDIRECT" => {
                if args.is_empty() {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let text = self
                    .evaluate_ast(&args[0], context, row, col, deps, scope)?
                    .to_string();
                let a1_style = match args.get(1) {
                    Some(a) => self.to_bool(&self.evaluate_ast(a, context, row, col, deps, scope)?),
                    None => true,
                };
                if !a1_style {
                    // R1C1-style reference text isn't supported.
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                match Self::parse_a1_reference(&text) {
                    Some((sheet_opt, start_row, start_col, end_row, end_col)) => {
                        if start_row == end_row && start_col == end_col {
                            Ok(self.read_cell_with_deps(
                                &sheet_opt, start_row, start_col, context, deps,
                            ))
                        } else {
                            match self.materialize_range(
                                &sheet_opt, start_row, start_col, end_row, end_col, context,
                            ) {
                                Some(grid) => {
                                    Ok(ResultData::List(grid.into_iter().flatten().collect()))
                                }
                                None => Ok(ResultData::Error("#REF!".to_string())),
                            }
                        }
                    }
                    None => Ok(ResultData::Error("#REF!".to_string())),
                }
            }
            "OFFSET" => {
                if args.len() < 3 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let Some((sheet_opt, base_row, base_col, base_end_row, base_end_col)) =
                    Self::range_bounds(&args[0])
                else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let row_offset = self
                    .to_f64(&self.evaluate_ast(&args[1], context, row, col, deps, scope)?)
                    .unwrap_or(0.0) as isize;
                let col_offset = self
                    .to_f64(&self.evaluate_ast(&args[2], context, row, col, deps, scope)?)
                    .unwrap_or(0.0) as isize;
                let base_height = (base_end_row.saturating_sub(base_row) + 1) as isize;
                let base_width = (base_end_col.saturating_sub(base_col) + 1) as isize;
                let height = match args.get(3) {
                    Some(a) => self
                        .to_f64(&self.evaluate_ast(a, context, row, col, deps, scope)?)
                        .unwrap_or(base_height as f64) as isize,
                    None => base_height,
                };
                let width = match args.get(4) {
                    Some(a) => self
                        .to_f64(&self.evaluate_ast(a, context, row, col, deps, scope)?)
                        .unwrap_or(base_width as f64) as isize,
                    None => base_width,
                };
                let new_row = base_row as isize + row_offset;
                let new_col = base_col as isize + col_offset;
                if new_row < 0 || new_col < 0 || height <= 0 || width <= 0 {
                    return Ok(ResultData::Error("#REF!".to_string()));
                }
                let (start_row, start_col) = (new_row as usize, new_col as usize);
                let (end_row, end_col) = (
                    start_row + (height - 1) as usize,
                    start_col + (width - 1) as usize,
                );
                if start_row == end_row && start_col == end_col {
                    Ok(self.read_cell_with_deps(&sheet_opt, start_row, start_col, context, deps))
                } else {
                    match self.materialize_range(
                        &sheet_opt, start_row, start_col, end_row, end_col, context,
                    ) {
                        Some(grid) => Ok(ResultData::List(grid.into_iter().flatten().collect())),
                        None => Ok(ResultData::Error("#REF!".to_string())),
                    }
                }
            }
            _ => unreachable!(),
        }
    }

    /// `GETPIVOTDATA(data_field, pivot_table_ref, [field, item]...)`.
    /// `pivot_table_ref` must stay an unevaluated cell reference (not a
    /// flattened value) so its sheet/row/col can be matched against
    /// `context.pivot_tables`' rendered destination ranges -- the same
    /// reason `ROW`/`OFFSET`/etc. go through `evaluate_range_info_function`
    /// instead of the generic eagerly-evaluated-args path below.
    fn evaluate_getpivotdata(
        &self,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        if args.len() < 2 || !(args.len() - 2).is_multiple_of(2) {
            return Ok(ResultData::Error("#VALUE!".to_string()));
        }

        let data_field = self
            .evaluate_ast(&args[0], context, row, col, deps, scope)?
            .to_string();

        let (sheet_opt, target_row, target_col, _, _) = match Self::range_bounds(&args[1]) {
            Some(bounds) => bounds,
            None => return Ok(ResultData::Error("#REF!".to_string())),
        };
        // Registers the usual dependency on the referenced cell, mirroring
        // how INDIRECT/OFFSET treat a dynamically resolved reference.
        self.read_cell_with_deps(&sheet_opt, target_row, target_col, context, deps);

        let sheet_id = match &sheet_opt {
            None => self.id,
            Some(name) if name == &self.name => self.id,
            Some(name) => match context.and_then(|c| c.sheets.get(name)) {
                Some(s) => s.id,
                None => return Ok(ResultData::Error("#REF!".to_string())),
            },
        };

        let pivot_tables = context.map(|c| c.pivot_tables).unwrap_or(&[]);
        let pivot = match pivot_tables.iter().find(|p| {
            p.dest_sheet_id == sheet_id
                && p.last_output_end_row
                    .is_some_and(|end| target_row >= p.dest_row && target_row <= end)
                && p.last_output_end_col
                    .is_some_and(|end| target_col >= p.dest_col && target_col <= end)
        }) {
            Some(p) => p,
            None => return Ok(ResultData::Error("#REF!".to_string())),
        };

        let mut criteria: Vec<(String, String)> = Vec::new();
        let mut i = 2;
        while i < args.len() {
            let field = self
                .evaluate_ast(&args[i], context, row, col, deps, scope)?
                .to_string();
            let item = self
                .evaluate_ast(&args[i + 1], context, row, col, deps, scope)?
                .to_string();
            criteria.push((field, item));
            i += 2;
        }

        let mut sheet_refs: Vec<&Sheet> = context
            .map(|c| c.sheets.values().copied().collect())
            .unwrap_or_default();
        sheet_refs.push(self);

        match crate::core::pivot::getpivotdata(&sheet_refs, pivot, &data_field, &criteria) {
            Ok(v) => Ok(v),
            Err(e) => Ok(ResultData::Error(e)),
        }
    }

    /// Shared implementation for the dynamic-array reshaping functions.
    /// All operate on `array_shape`'s `(flat, num_cols)` view and return a
    /// flat, row-major `ResultData::List` -- the same convention
    /// `SEQUENCE`/`MUNIT`/`MAKEARRAY`/etc. already use, since this engine
    /// doesn't spill formulas across cells (a caller pulls out a single
    /// value with `INDEX`, or consumes the whole list with e.g. `SUM`).
    ///
    /// Known simplifications, each accepted given limited fuzzing time
    /// against real Excel for this batch: `UNIQUE`'s `by_col` and `SORT`'s
    /// `by_col` arguments are ignored (both always operate row-wise);
    /// `SORTBY` only supports a single `by_array`/`sort_order` pair, not
    /// the documented repeating list; `XMATCH`'s wildcard match mode and
    /// binary/reverse search modes aren't implemented (falls through to a
    /// forward linear scan).
    #[allow(clippy::too_many_arguments)]
    fn evaluate_array_reshape_function(
        &self,
        func_name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        match func_name {
            "TRANSPOSE" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let (flat, cols) = self.array_shape(arg, context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let mut result = Vec::with_capacity(flat.len());
                for c in 0..cols {
                    for r in 0..rows {
                        result.push(flat[r * cols + c].clone());
                    }
                }
                Ok(ResultData::List(result))
            }
            "HSTACK" | "VSTACK" => {
                if args.is_empty() {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let mut shapes = Vec::with_capacity(args.len());
                for a in args {
                    shapes.push(self.array_shape(a, context, row, col, deps, scope)?);
                }
                let mut result = Vec::new();
                if func_name == "HSTACK" {
                    let max_rows = shapes
                        .iter()
                        .map(|(f, c)| if *c == 0 { 0 } else { f.len() / c })
                        .max()
                        .unwrap_or(0);
                    for r in 0..max_rows {
                        for (flat, cols) in &shapes {
                            let rows = if *cols == 0 { 0 } else { flat.len() / cols };
                            for c in 0..*cols {
                                result.push(if r < rows {
                                    flat[r * cols + c].clone()
                                } else {
                                    ResultData::Error("#N/A".to_string())
                                });
                            }
                        }
                    }
                } else {
                    let max_cols = shapes.iter().map(|(_, c)| *c).max().unwrap_or(0);
                    for (flat, cols) in &shapes {
                        let rows = if *cols == 0 { 0 } else { flat.len() / cols };
                        for r in 0..rows {
                            for c in 0..max_cols {
                                result.push(if c < *cols {
                                    flat[r * cols + c].clone()
                                } else {
                                    ResultData::Error("#N/A".to_string())
                                });
                            }
                        }
                    }
                }
                Ok(ResultData::List(result))
            }
            "CHOOSEROWS" | "CHOOSECOLS" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (flat, cols) = self.array_shape(&args[0], context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let total = if func_name == "CHOOSEROWS" {
                    rows
                } else {
                    cols
                } as isize;
                let mut indices = Vec::with_capacity(args.len() - 1);
                for idx_expr in &args[1..] {
                    let n = self
                        .to_f64(&self.evaluate_ast(idx_expr, context, row, col, deps, scope)?)
                        .unwrap_or(0.0) as isize;
                    let real_idx = if n < 0 { total + n } else { n - 1 };
                    if real_idx < 0 || real_idx >= total {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    indices.push(real_idx as usize);
                }
                let mut result = Vec::new();
                if func_name == "CHOOSEROWS" {
                    for r in indices {
                        for c in 0..cols {
                            result.push(flat[r * cols + c].clone());
                        }
                    }
                } else {
                    for r in 0..rows {
                        for &c in &indices {
                            result.push(flat[r * cols + c].clone());
                        }
                    }
                }
                Ok(ResultData::List(result))
            }
            "DROP" | "TAKE" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (flat, cols) = self.array_shape(&args[0], context, row, col, deps, scope)?;
                let num_rows = flat.len().checked_div(cols).unwrap_or(0) as isize;
                let is_take = func_name == "TAKE";
                let rows_n = self
                    .to_f64(&self.evaluate_ast(&args[1], context, row, col, deps, scope)?)
                    .unwrap_or(0.0) as isize;
                let cols_n = match args.get(2) {
                    Some(e) => self
                        .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope)?)
                        .unwrap_or(0.0) as isize,
                    None => {
                        if is_take {
                            cols as isize
                        } else {
                            0
                        }
                    }
                };
                let (row_start, row_end) = Self::drop_take_bounds(num_rows, rows_n, is_take);
                let (col_start, col_end) = Self::drop_take_bounds(cols as isize, cols_n, is_take);
                if row_start >= row_end || col_start >= col_end {
                    return Ok(ResultData::Error("#CALC!".to_string()));
                }
                let mut result = Vec::new();
                for r in row_start..row_end {
                    for c in col_start..col_end {
                        result.push(flat[(r as usize) * cols + (c as usize)].clone());
                    }
                }
                Ok(ResultData::List(result))
            }
            "EXPAND" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (flat, cols) = self.array_shape(&args[0], context, row, col, deps, scope)?;
                let orig_rows = flat.len().checked_div(cols).unwrap_or(0);
                let new_rows = self
                    .to_f64(&self.evaluate_ast(&args[1], context, row, col, deps, scope)?)
                    .unwrap_or(orig_rows as f64) as usize;
                let new_cols = match args.get(2) {
                    Some(e) => self
                        .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope)?)
                        .unwrap_or(cols as f64) as usize,
                    None => cols,
                };
                let pad = match args.get(3) {
                    Some(e) => self.evaluate_ast(e, context, row, col, deps, scope)?,
                    None => ResultData::Error("#N/A".to_string()),
                };
                if new_rows < orig_rows || new_cols < cols {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let mut result = Vec::with_capacity(new_rows * new_cols);
                for r in 0..new_rows {
                    for c in 0..new_cols {
                        result.push(if r < orig_rows && c < cols {
                            flat[r * cols + c].clone()
                        } else {
                            pad.clone()
                        });
                    }
                }
                Ok(ResultData::List(result))
            }
            "TOCOL" | "TOROW" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let (flat, cols) = self.array_shape(arg, context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let ignore = match args.get(1) {
                    Some(e) => self
                        .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope)?)
                        .unwrap_or(0.0) as i64,
                    None => 0,
                };
                let scan_by_col = match args.get(2) {
                    Some(e) => self.to_bool(&self.evaluate_ast(e, context, row, col, deps, scope)?),
                    None => false,
                };
                let ordered: Vec<ResultData> = if scan_by_col {
                    let mut v = Vec::with_capacity(flat.len());
                    for c in 0..cols {
                        for r in 0..rows {
                            v.push(flat[r * cols + c].clone());
                        }
                    }
                    v
                } else {
                    flat
                };
                let filtered: Vec<ResultData> = ordered
                    .into_iter()
                    .filter(|v| match ignore {
                        1 => !matches!(v, ResultData::None),
                        2 => !matches!(v, ResultData::Error(_)),
                        3 => !matches!(v, ResultData::None | ResultData::Error(_)),
                        _ => true,
                    })
                    .collect();
                Ok(ResultData::List(filtered))
            }
            "WRAPROWS" | "WRAPCOLS" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (flat, _cols) = self.array_shape(&args[0], context, row, col, deps, scope)?;
                let wrap = self
                    .to_f64(&self.evaluate_ast(&args[1], context, row, col, deps, scope)?)
                    .unwrap_or(1.0)
                    .max(1.0) as usize;
                let pad = match args.get(2) {
                    Some(e) => self.evaluate_ast(e, context, row, col, deps, scope)?,
                    None => ResultData::Error("#N/A".to_string()),
                };
                if func_name == "WRAPROWS" {
                    // Row-major flat storage with num_cols == wrap is
                    // exactly the padded input sequence itself.
                    let mut result = flat;
                    let rem = result.len() % wrap;
                    if rem != 0 {
                        result.extend(std::iter::repeat_n(pad, wrap - rem));
                    }
                    Ok(ResultData::List(result))
                } else {
                    let num_result_cols = flat.len().div_ceil(wrap).max(1);
                    let total = wrap * num_result_cols;
                    let mut result = Vec::with_capacity(total);
                    for i in 0..total {
                        let col = i / wrap;
                        let r = i % wrap;
                        let target = r * num_result_cols + col;
                        while result.len() <= target {
                            result.push(pad.clone());
                        }
                        if i < flat.len() {
                            result[target] = flat[i].clone();
                        }
                    }
                    Ok(ResultData::List(result))
                }
            }
            "UNIQUE" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let (flat, _cols) = self.array_shape(arg, context, row, col, deps, scope)?;
                let exactly_once = match args.get(2) {
                    Some(e) => self.to_bool(&self.evaluate_ast(e, context, row, col, deps, scope)?),
                    None => false,
                };
                let mut seen: Vec<(String, ResultData, usize)> = Vec::new();
                for v in &flat {
                    // UNIQUE compares values without the cross-type coercion
                    // used by worksheet comparison operators: text "3" and
                    // numeric 3 are distinct values.
                    let key = match v {
                        ResultData::None => "blank:".to_string(),
                        ResultData::Boolean(b) => format!("bool:{b}"),
                        ResultData::Integer(i) => format!("num:{}", *i as f64),
                        ResultData::Float(f) => format!("num:{f}"),
                        ResultData::String(s) => format!("str:{s}"),
                        ResultData::Error(e) => format!("err:{e}"),
                        ResultData::List(_) | ResultData::Dict(_) => format!("other:{v}"),
                    };
                    match seen.iter_mut().find(|(k, ..)| k == &key) {
                        Some(entry) => entry.2 += 1,
                        None => seen.push((key, v.clone(), 1)),
                    }
                }
                let result: Vec<ResultData> = seen
                    .into_iter()
                    .filter(|(_, _, count)| !exactly_once || *count == 1)
                    .map(|(_, v, _)| v)
                    .collect();
                Ok(ResultData::List(result))
            }
            "SORT" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let (flat, cols) = self.array_shape(arg, context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let sort_index = match args.get(1) {
                    Some(e) => self
                        .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope)?)
                        .unwrap_or(1.0) as usize,
                    None => 1,
                };
                let sort_order = match args.get(2) {
                    Some(e) => self
                        .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope)?)
                        .unwrap_or(1.0),
                    None => 1.0,
                };
                let col_idx = sort_index.saturating_sub(1).min(cols.saturating_sub(1));
                let mut row_indices: Vec<usize> = (0..rows).collect();
                row_indices.sort_by(|&a, &b| {
                    Self::sort_compare_blanks_last(
                        &flat[a * cols + col_idx],
                        &flat[b * cols + col_idx],
                        sort_order,
                    )
                });
                let mut result = Vec::with_capacity(flat.len());
                for r in row_indices {
                    for c in 0..cols {
                        result.push(flat[r * cols + c].clone());
                    }
                }
                Ok(ResultData::List(result))
            }
            "SORTBY" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (flat, cols) = self.array_shape(&args[0], context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let by = self.eval_as_array(&args[1], context, row, col, deps, scope)?;
                let order = match args.get(2) {
                    Some(e) => self
                        .to_f64(&self.evaluate_ast(e, context, row, col, deps, scope)?)
                        .unwrap_or(1.0),
                    None => 1.0,
                };
                let mut row_indices: Vec<usize> = (0..rows).collect();
                row_indices.sort_by(|&a, &b| {
                    let va = by.get(a).cloned().unwrap_or(ResultData::None);
                    let vb = by.get(b).cloned().unwrap_or(ResultData::None);
                    Self::sort_compare_blanks_last(&va, &vb, order)
                });
                let mut result = Vec::with_capacity(flat.len());
                for r in row_indices {
                    for c in 0..cols {
                        result.push(flat[r * cols + c].clone());
                    }
                }
                Ok(ResultData::List(result))
            }
            "FILTER" => {
                if args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let (flat, cols) = self.array_shape(&args[0], context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let include = self.eval_as_array(&args[1], context, row, col, deps, scope)?;
                let mut result = Vec::new();
                for r in 0..rows {
                    let keep = include.get(r).map(|v| self.to_bool(v)).unwrap_or(false);
                    if keep {
                        for c in 0..cols {
                            result.push(flat[r * cols + c].clone());
                        }
                    }
                }
                if result.is_empty() {
                    match args.get(2) {
                        Some(e) => Ok(self.evaluate_ast(e, context, row, col, deps, scope)?),
                        None => Ok(ResultData::Error("#CALC!".to_string())),
                    }
                } else {
                    Ok(ResultData::List(result))
                }
            }
            "TRIMRANGE" => {
                let Some(arg) = args.first() else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                let (flat, cols) = self.array_shape(arg, context, row, col, deps, scope)?;
                let rows = flat.len().checked_div(cols).unwrap_or(0);
                let is_blank = |v: &ResultData| {
                    matches!(v, ResultData::None)
                        || matches!(v, ResultData::String(s) if s.is_empty())
                };
                let row_blank = |r: usize| (0..cols).all(|c| is_blank(&flat[r * cols + c]));
                let col_blank = |c: usize| (0..rows).all(|r| is_blank(&flat[r * cols + c]));
                let mut r_start = 0;
                while r_start < rows && row_blank(r_start) {
                    r_start += 1;
                }
                let mut r_end = rows;
                while r_end > r_start && row_blank(r_end - 1) {
                    r_end -= 1;
                }
                let mut c_start = 0;
                while c_start < cols && col_blank(c_start) {
                    c_start += 1;
                }
                let mut c_end = cols;
                while c_end > c_start && col_blank(c_end - 1) {
                    c_end -= 1;
                }
                let mut result = Vec::new();
                for r in r_start..r_end {
                    for c in c_start..c_end {
                        result.push(flat[r * cols + c].clone());
                    }
                }
                Ok(ResultData::List(result))
            }
            _ => unreachable!(),
        }
    }

    /// The raw text typed into a cell -- `"10"`, `"=SUM(A1:A2)"` -- or `None`
    /// if the cell is outside the sheet's allocated grid.
    ///
    /// This is the input, not the result; see [`Sheet::get_result_data`] for
    /// the computed value and `Sheet::get_display_string` for what a user
    /// should see.
    pub fn get_src(&self, cell: &CellRef) -> Option<&String> {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.src.get(cell.row)
        } else {
            None
        }
    }

    /// [`Sheet::get_src`] with an out-of-range cell flattened to an owned
    /// empty string.
    pub fn get_src_str(&self, cell: &CellRef) -> String {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.src.get(cell.row).cloned().unwrap_or("".to_string())
        } else {
            "".to_string()
        }
    }

    /// [`Sheet::get_src`] as a borrowed `&str`, for callers that only read.
    pub fn get_src_str_ref(&self, cell: &CellRef) -> Option<&str> {
        let col = self.columns.get(cell.col)?;
        col.src.get(cell.row).map(|s| s.as_str())
    }

    /// The word surrounding `char_offset` in a cell's source text, as a
    /// half-open range of character (not byte) indices -- what an editor needs
    /// for word-wise selection. See [`get_word_boundaries_from_str`].
    pub fn get_word_boundaries(&self, cell: &CellRef, char_offset: usize) -> (usize, usize) {
        let text = self.get_src_str(cell);
        get_word_boundaries_from_str(&text, char_offset)
    }
}
