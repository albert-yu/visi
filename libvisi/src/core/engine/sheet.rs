use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use web_time::Instant;

use super::cell::{CellRef, Dependency, EngineError, EvalError, TextCellRef, generate_unique_id};
use super::column::{ColumnPosition, DataColumn};
use super::result_data::ResultData;
use crate::core::finance;
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

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Direction {
    None,
    Up,
    Down,
    Left,
    Right,
}

/// Represents a sheet in a spreadsheet.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    #[serde(default = "generate_unique_id")]
    pub id: u64,
    pub name: String,
    pub columns: Vec<DataColumn>,
    #[serde(default)]
    pub tables: Vec<crate::core::table::ExcelTable>,
    #[serde(skip, default)]
    pub dependencies: HashMap<Dependency, HashSet<CellRef>>,
    #[serde(skip, default)]
    pub dependencies_rev: HashMap<CellRef, HashSet<Dependency>>,
    #[serde(skip)]
    pub uncommitted_actions: Vec<crate::core::SheetAction>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetInit {
    #[serde(default)]
    pub id: Option<u64>,
    pub name: Option<String>,
    pub rows: usize,
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
    /// Counts as 0 (GCD/LCM).
    Zero,
    /// Dropped entirely, shifting later elements (SERIESSUM).
    Skip,
    /// #VALUE!, like text (LINEST/TREND/GROWTH/LOGEST/MMULT).
    Reject,
}

impl Sheet {
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
        }
    }

    pub fn setup_after_deserialization(&mut self) {
        for col in &mut self.columns {
            let size = col.src.len();
            col.data.resize(size);
            col.compiled_src = vec![crate::core::CompiledFormula::default(); size].into();
        }
        self.mark_all_dirty();
    }

    pub fn get_all_tables_for_compilation(&self, context: Option<&Context>) -> Vec<Sheet> {
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

            let (result, new_deps, compiled_to_cache) = {
                let src = self.get_src_str_ref(&cell_ref).unwrap_or("");
                if !src.starts_with('=') {
                    let res = if src.is_empty() {
                        ResultData::None
                    } else if src.starts_with('"') && src.ends_with('"') && src.len() >= 2 {
                        ResultData::String(src[1..src.len() - 1].to_string())
                    } else if let Ok(i) = src.parse::<i64>() {
                        ResultData::Integer(i)
                    } else if let Ok(f) = src.parse::<f64>() {
                        ResultData::Float(f)
                    } else if src.eq_ignore_ascii_case("true") {
                        ResultData::Boolean(true)
                    } else if src.eq_ignore_ascii_case("false") {
                        ResultData::Boolean(false)
                    } else {
                        ResultData::String(src.to_string())
                    };
                    (res, vec![], None)
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
                    (final_res, deps, Some(compiled))
                }
            };

            // Write compiled cache
            if let Some(col) = self.columns.get_mut(cell_ref.col)
                && cell_ref.row < col.compiled_src.len()
            {
                col.compiled_src[cell_ref.row] = compiled_to_cache.unwrap_or_default();
            }

            // Update dependencies
            // 1. Remove old reverse dependencies
            if let Some(old_deps) = self.dependencies_rev.remove(&cell_ref) {
                for provider in old_deps {
                    if let Some(dependents) = self.dependencies.get_mut(&provider) {
                        dependents.remove(&cell_ref);
                    }
                }
            }

            // 2. Add new dependencies (only if not empty to save map allocations)
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

            // Update data
            if let Some(col) = self.columns.get_mut(cell_ref.col)
                && cell_ref.row < col.data.len()
            {
                col.data.set(cell_ref.row, result.clone());
                updated_cells.insert(cell_ref);
            }
            if let Some(comp_sheet) = tables_for_compilation
                .iter_mut()
                .find(|s| s.name == self.name)
                && let Some(col) = comp_sheet.columns.get_mut(cell_ref.col)
                && cell_ref.row < col.data.len()
            {
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

    /// Mark cells as dirty if they depend on the given dependency
    pub fn invalidate_dependency(&mut self, dep: &Dependency) {
        if let Some(dependents) = self.dependencies.get(dep) {
            for dependent in dependents {
                // Mark as dirty so commit will pick it up
                if let Some(col) = self.columns.get_mut(dependent.col)
                    && !col.dirty_indices.contains(&dependent.row)
                {
                    col.dirty_indices.push(dependent.row);
                }
            }
        }
    }

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

                let actual_end_row = if *end_row == usize::MAX {
                    if is_self {
                        self.row_count().saturating_sub(1)
                    } else if let Some(ctx) = context {
                        ctx.sheets
                            .get(sheet.as_ref().unwrap())
                            .map(|t| t.row_count().saturating_sub(1))
                            .unwrap_or(0)
                    } else {
                        0
                    }
                } else {
                    *end_row
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
                // referential formula bounded by commit()'s max_ops, see
                // the fix just above), that quadratic-in-width blowup was
                // the difference between finishing in under a second and
                // taking tens of seconds to minutes -- found via the same
                // libvisi/fuzz formula_eval run (#26).
                let mut seen_col_deps: HashSet<usize> = HashSet::new();

                let mut results = Vec::new();
                for r in *start_row..=actual_end_row {
                    for c in *start_col..=*end_col {
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
                            // itself one level deeper each pass --
                            // unbounded growth that only stops at a stack
                            // overflow in recursive Clone/Drop, found via
                            // libvisi/fuzz's formula_eval target (#26).
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
                let r_val = self.evaluate_ast(right, context, row, col, deps, scope)?;

                match op {
                    Op::Eq | Op::Ne | Op::Lt | Op::Gt | Op::Le | Op::Ge => {
                        if let ResultData::Error(_) = &l_val {
                            return Ok(l_val);
                        }
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
        // Coerce ResultData::None against the type of the opposing operand
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
    /// Found via the differential fuzzer: `SORT({-215.8,,-100,-240.97,-88},1,-1)`
    /// put the blank first (coerced to 0, the largest value once reversed)
    /// instead of last, so `INDEX(...,1)` returned 0 instead of -88.
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
        let a_low = a.to_lowercase();
        let b_low = b.to_lowercase();
        let char_weight = |ch: char| -> u32 {
            match ch {
                '-' => 1,
                '(' => 2,
                ')' => 3,
                _ => (ch as u32) + 10,
            }
        };
        for (ca, cb) in a_low.chars().zip(b_low.chars()) {
            if ca != cb {
                let wa = char_weight(ca);
                let wb = char_weight(cb);
                return wa.cmp(&wb);
            }
        }
        a_low.len().cmp(&b_low.len())
    }

    pub fn clean_float(val: f64) -> f64 {
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

    pub fn to_f64(&self, val: &ResultData) -> Option<f64> {
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
                    if let Some((date, _)) = crate::core::date::parse_date(s_trim) {
                        return Some(crate::core::date::date_to_excel_serial(date));
                    }
                    None
                } else if let Some((date, _)) = crate::core::date::parse_date(s_trim) {
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
    /// `blanks` selects between the three behaviours real Excel actually
    /// exhibits here, each established by probing it directly:
    ///  - `BlankPolicy::Zero` (GCD/LCM): a blank counts as 0 and the call
    ///    still succeeds. `GCD` over `{4, 6, <blank>, 8}` is 2 and `LCM`
    ///    over it is 0 (i.e. the blank really did participate as a zero),
    ///    while the same range with `TRUE` in place of the blank is
    ///    #VALUE!.
    ///  - `BlankPolicy::Skip` (SERIESSUM): a blank is dropped outright,
    ///    which *shifts* every later coefficient down a power.
    ///    `SERIESSUM(0.5, 0, 2, {4, 6, <blank>, 8})` is 6.0 -- exactly the
    ///    3-coefficient answer -- not the 5.625 a zero in that slot gives.
    ///  - `BlankPolicy::Reject` (LINEST/TREND/GROWTH/LOGEST/MMULT): text,
    ///    booleans *and* blanks are all #VALUE!. LINEST returns #VALUE!
    ///    for each of those three separately and only computes when every
    ///    cell is a real number.
    ///
    /// Note this deliberately does not go through `to_f64`, which is the
    /// lenient coercion used for scalar arguments -- that maps a blank to
    /// 0, a boolean to 1/0, and a numeric-looking string to its value,
    /// none of which these functions accept.
    fn flatten_strict_inner(
        &self,
        arg: &ResultData,
        blanks: BlankPolicy,
        out: &mut Vec<f64>,
    ) -> Result<(), String> {
        match arg {
            ResultData::List(items) => {
                for item in items {
                    self.flatten_strict_inner(item, blanks, out)?;
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
            _ => Err("#VALUE!".to_string()),
        }
    }

    fn flatten_strict_numbers(&self, arg: &ResultData) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        self.flatten_strict_inner(arg, BlankPolicy::Zero, &mut out)?;
        Ok(out)
    }

    /// flatten_strict_numbers with blanks dropped rather than zero-filled.
    fn flatten_skipping_blanks(&self, arg: Option<&ResultData>) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        if let Some(a) = arg {
            self.flatten_strict_inner(a, BlankPolicy::Skip, &mut out)?;
        }
        Ok(out)
    }

    /// flatten_strict_numbers with the stricter "a blank is also #VALUE!"
    /// rule the regression-array and matrix functions use.
    fn flatten_numbers_only(&self, arg: &ResultData) -> Result<Vec<f64>, String> {
        let mut out = Vec::new();
        self.flatten_strict_inner(arg, BlankPolicy::Reject, &mut out)?;
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
            ResultData::String(s) => {
                if s.is_empty() {
                    0
                } else {
                    1
                }
            }
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
            let remainder = rest.trim().to_string();
            val.to_string() != remainder
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
            let remainder = rest.trim().to_string();
            val.to_string() == remainder
        } else {
            val.to_string() == crit_str
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
        if actual_end_row < start_row || end_col < start_col {
            return Some(Vec::new());
        }
        let mut grid = Vec::with_capacity(actual_end_row - start_row + 1);
        for r in start_row..=actual_end_row {
            let mut row = Vec::with_capacity(end_col - start_col + 1);
            for c in start_col..=end_col {
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
    /// `INDEX(EXPAND(A1:B2,3,3,0),3,3)`) previously fell back to treating
    /// the whole result as a single row, corrupting the flat-index math.
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
    /// `parse_range_ref`): `libvisi` cannot depend on the `visi` crate
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
                    let key = v.to_string();
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

    #[allow(clippy::too_many_arguments)]
    fn evaluate_function(
        &self,
        name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::parser::Expr;
        let mut upper_name = name.to_uppercase();
        if upper_name.starts_with("_XLFN.") {
            upper_name = upper_name["_XLFN.".len()..].to_string();
        }
        // Real Excel's OOXML writer additionally nests some dynamic-array
        // worksheet functions (UNIQUE, SORT, FILTER, ...) under a second
        // `_xlws.` prefix inside `_xlfn.` -- e.g. `_xlfn._xlws.SORT`, not
        // just `_xlfn.SORT`. Without stripping it too, the un-stripped
        // name never matches any dispatch arm here -- confirmed as a real
        // mismatch by checking real Excel's own OOXML export for these
        // functions directly.
        if upper_name.starts_with("_XLWS.") {
            upper_name = upper_name["_XLWS.".len()..].to_string();
        }

        if upper_name == "LET" {
            return self.evaluate_let(args, context, row, col, deps, scope);
        }

        if upper_name == "PLOT" {
            let mut x_vals = Vec::new();
            let mut y_vals = Vec::new();

            let mut color = [0.2, 0.6, 1.0, 1.0];
            let mut radius = 0.005; // default radius
            let mut is_line = false;
            let mut title = None;
            let mut xlabel = None;
            let mut ylabel = None;

            let mut positional_count = 0;

            for arg in args {
                use crate::core::parser::Expr;
                use crate::core::parser::Op;
                if let Expr::BinaryOp {
                    op: Op::Eq,
                    left,
                    right,
                } = arg
                    && let Expr::Identifier(name) = &**left
                {
                    let val = self.evaluate_ast(right, context, row, col, deps, scope)?;
                    match name.to_lowercase().as_str() {
                        "color" => {
                            if let ResultData::List(list) = val {
                                for (i, v) in list.iter().take(4).enumerate() {
                                    color[i] = self.to_f64(v).unwrap_or(color[i] as f64) as f32;
                                }
                            }
                        }
                        "radius" => {
                            radius = self.to_f64(&val).unwrap_or(radius as f64) as f32;
                        }
                        "type" => {
                            if val.to_string() == "line" {
                                is_line = true;
                            }
                        }
                        "title" => {
                            title = Some(val.to_string());
                        }
                        "xlabel" => {
                            xlabel = Some(val.to_string());
                        }
                        "ylabel" => {
                            ylabel = Some(val.to_string());
                        }
                        _ => {}
                    }
                    continue;
                }

                let val = self.evaluate_ast(arg, context, row, col, deps, scope)?;
                match positional_count {
                    0 => {
                        if let ResultData::List(list) = val {
                            x_vals = list
                                .iter()
                                .map(|v| self.to_f64(v).unwrap_or(0.0) as f32)
                                .collect();
                        }
                    }
                    1 => {
                        if let ResultData::List(list) = val {
                            y_vals = list
                                .iter()
                                .map(|v| self.to_f64(v).unwrap_or(0.0) as f32)
                                .collect();
                        }
                    }
                    2 => match val {
                        ResultData::String(s) => {
                            if s == "line" {
                                is_line = true;
                            }
                        }
                        ResultData::List(list) => {
                            for (i, v) in list.iter().take(4).enumerate() {
                                color[i] = self.to_f64(v).unwrap_or(color[i] as f64) as f32;
                            }
                        }
                        ResultData::Float(f) => {
                            radius = f as f32;
                        }
                        ResultData::Integer(i) => {
                            radius = i as f32;
                        }
                        _ => {}
                    },
                    3 => match val {
                        ResultData::String(s) => {
                            if s == "line" {
                                is_line = true;
                            }
                        }
                        ResultData::Float(f) => {
                            radius = f as f32;
                        }
                        ResultData::Integer(i) => {
                            radius = i as f32;
                        }
                        _ => {}
                    },
                    4 if val.to_string() == "line" => {
                        is_line = true;
                    }
                    _ => {}
                }
                positional_count += 1;
            }

            let mut points = Vec::new();
            for i in 0..x_vals.len().min(y_vals.len()) {
                points.push((x_vals[i], y_vals[i]));
            }

            Ok(ResultData::Plot {
                points,
                color,
                radius,
                is_line,
                title,
                xlabel,
                ylabel,
            })
        } else {
            if upper_name == "IF" {
                if args.len() < 3 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "IF requires 3 arguments".to_string(),
                    )));
                }
                let cond_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
                if let ResultData::Error(_) = cond_val {
                    return Ok(cond_val);
                }
                let condition = match self.to_bool_opt(&cond_val) {
                    Some(b) => b,
                    None => return Ok(ResultData::Error("#VALUE!".to_string())),
                };
                if condition {
                    return self.evaluate_ast(&args[1], context, row, col, deps, scope);
                } else {
                    return self.evaluate_ast(&args[2], context, row, col, deps, scope);
                }
            }

            if upper_name == "IFERROR" {
                if args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "IFERROR requires 2 arguments".to_string(),
                    )));
                }
                let first_res = self.evaluate_ast(&args[0], context, row, col, deps, scope);
                match first_res {
                    Ok(ResultData::Error(_)) | Err(_) => {
                        return self.evaluate_ast(&args[1], context, row, col, deps, scope);
                    }
                    Ok(val) => return Ok(val),
                }
            }

            if upper_name == "IFNA" {
                if args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "IFNA requires 2 arguments".to_string(),
                    )));
                }
                let first_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
                if let ResultData::Error(ref e) = first_val
                    && e == "#N/A"
                {
                    return self.evaluate_ast(&args[1], context, row, col, deps, scope);
                }
                return Ok(first_val);
            }

            if upper_name == "IFS" {
                // Lazily evaluated: only the arms up to and including the
                // first TRUE condition are ever computed, so an error
                // sitting in a later (unselected) value never propagates.
                // Confirmed against real Excel: `IFS(TRUE, 42, TRUE, 1/0)`
                // is 42, while `IFS(FALSE, 42, TRUE, 1/0)` is #DIV/0!.
                let mut i = 0;
                while i + 1 < args.len() {
                    let cond = self.evaluate_ast(&args[i], context, row, col, deps, scope)?;
                    if let ResultData::Error(_) = cond {
                        return Ok(cond);
                    }
                    if self.to_bool(&cond) {
                        return self.evaluate_ast(&args[i + 1], context, row, col, deps, scope);
                    }
                    i += 2;
                }
                return Ok(ResultData::Error("#N/A".to_string()));
            }

            if upper_name == "SWITCH" {
                // Lazily evaluated for the same reason as IFS: an error in
                // a value arm that isn't selected must not propagate
                // (`SWITCH(2, 1, 1/0, 2, 99, -1)` is 99 in real Excel).
                if args.len() < 3 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let target = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
                if let ResultData::Error(_) = target {
                    return Ok(target);
                }
                let mut i = 1;
                while i + 1 < args.len() {
                    let case = self.evaluate_ast(&args[i], context, row, col, deps, scope)?;
                    if let ResultData::Error(_) = case {
                        return Ok(case);
                    }
                    if target.to_string() == case.to_string() {
                        return self.evaluate_ast(&args[i + 1], context, row, col, deps, scope);
                    }
                    i += 2;
                }
                // A trailing odd argument is the default.
                if i < args.len() {
                    return self.evaluate_ast(&args[i], context, row, col, deps, scope);
                }
                return Ok(ResultData::Error("#N/A".to_string()));
            }

            if upper_name == "CHOOSE" {
                if args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "CHOOSE requires at least 2 arguments".to_string(),
                    )));
                }
                let idx_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
                if let ResultData::Error(_) = idx_val {
                    return Ok(idx_val);
                }
                let idx = match self.to_f64(&idx_val) {
                    Some(f) => f.round() as isize,
                    None => return Ok(ResultData::Error("#VALUE!".to_string())),
                };
                let choices = &args[1..];
                if idx >= 1 && (idx as usize) <= choices.len() {
                    return self.evaluate_ast(
                        &choices[(idx - 1) as usize],
                        context,
                        row,
                        col,
                        deps,
                        scope,
                    );
                } else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
            }

            if upper_name == "LAMBDA" {
                // A bare, uninvoked LAMBDA (not nested as another
                // function's argument, e.g. `=LAMBDA(x, x*2)` alone in a
                // cell) has nothing to apply it to -- the parser doesn't
                // support the `LAMBDA(...)(args)` immediate-invocation
                // syntax (that would need the grammar to allow calling an
                // arbitrary sub-expression, not just a bare identifier),
                // so this mirrors Excel's #CALC! for an unusable lambda.
                return Ok(ResultData::Error("#CALC!".to_string()));
            }

            if matches!(
                upper_name.as_str(),
                "MAP" | "BYROW" | "BYCOL" | "REDUCE" | "SCAN" | "MAKEARRAY"
            ) {
                return self.evaluate_lambda_function(
                    upper_name.as_str(),
                    args,
                    context,
                    row,
                    col,
                    deps,
                    scope,
                );
            }

            if upper_name == "ISOMITTED" {
                // Best-effort: every lambda invocation path here
                // (MAP/BYROW/BYCOL/REDUCE/SCAN/MAKEARRAY) always supplies
                // exactly as many argument values as the lambda declares
                // parameters, so a declared parameter is never actually
                // left unbound -- this can only ever observe "not found
                // in scope at all", which is the honest limitation to
                // report rather than silently guessing.
                let is_omitted = match args.first() {
                    Some(Expr::Identifier(name)) => scope.get(name).is_none(),
                    _ => false,
                };
                return Ok(ResultData::Boolean(is_omitted));
            }

            if matches!(
                upper_name.as_str(),
                "ROW"
                    | "ROWS"
                    | "COLUMN"
                    | "COLUMNS"
                    | "AREAS"
                    | "ISREF"
                    | "FORMULATEXT"
                    | "ISFORMULA"
                    | "INDIRECT"
                    | "OFFSET"
                    | "SHEET"
                    | "SHEETS"
                    | "CELL"
                    | "INFO"
            ) {
                return self.evaluate_range_info_function(
                    upper_name.as_str(),
                    args,
                    context,
                    row,
                    col,
                    deps,
                    scope,
                );
            }

            if matches!(
                upper_name.as_str(),
                "TRANSPOSE"
                    | "HSTACK"
                    | "VSTACK"
                    | "CHOOSEROWS"
                    | "CHOOSECOLS"
                    | "DROP"
                    | "EXPAND"
                    | "TAKE"
                    | "TOCOL"
                    | "TOROW"
                    | "WRAPROWS"
                    | "WRAPCOLS"
                    | "UNIQUE"
                    | "SORT"
                    | "SORTBY"
                    | "FILTER"
                    | "TRIMRANGE"
            ) {
                return self.evaluate_array_reshape_function(
                    upper_name.as_str(),
                    args,
                    context,
                    row,
                    col,
                    deps,
                    scope,
                );
            }

            if upper_name == "GETPIVOTDATA" {
                return self.evaluate_getpivotdata(args, context, row, col, deps, scope);
            }

            if upper_name == "ISERROR" {
                if args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                let res = self.evaluate_ast(&args[0], context, row, col, deps, scope);
                return match res {
                    Ok(ResultData::Error(_)) | Err(_) => Ok(ResultData::Boolean(true)),
                    _ => Ok(ResultData::Boolean(false)),
                };
            }

            if upper_name == "ISNA" {
                if args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                let res = self.evaluate_ast(&args[0], context, row, col, deps, scope);
                return match res {
                    Ok(ResultData::Error(e)) => Ok(ResultData::Boolean(e.contains("#N/A"))),
                    _ => Ok(ResultData::Boolean(false)),
                };
            }

            let mut evaluated_args = Vec::new();
            let mut arg_is_direct = Vec::new();
            for arg in args {
                let is_direct_arg = match arg {
                    Expr::CellRef { .. } | Expr::RangeRef { .. } | Expr::StructuredRef { .. } => {
                        false
                    }
                    Expr::FunctionCall { name, .. } => {
                        let n = name.to_uppercase();
                        n != "IF" && n != "IFERROR" && n != "CHOOSE"
                    }
                    _ => true,
                };
                arg_is_direct.push(is_direct_arg);
                let eval_res = match self.evaluate_ast(arg, context, row, col, deps, scope) {
                    Ok(r) => r,
                    Err(EngineError::EvalError(EvalError::UnknownFunction(err_str)))
                        if err_str.starts_with('#') =>
                    {
                        ResultData::Error(err_str)
                    }
                    Err(e) => return Err(e),
                };
                evaluated_args.push(eval_res);
            }

            let uses_ordered_arg_error_check = matches!(
                upper_name.as_str(),
                "SUM" | "AVERAGE" | "MIN" | "MAX" | "PRODUCT"
            );
            // The type-introspection functions must see an error value
            // rather than have it propagate past them: real Excel answers
            // TYPE(1/0) = 16, ISNONTEXT(1/0) = TRUE, and
            // ISTEXT/ISNUMBER/ISLOGICAL/ISBLANK(1/0) = FALSE. (Math
            // functions like ISODD do still propagate -- ISODD(1/0) is
            // #DIV/0! -- so they stay out of this list.)
            let inspects_errors = matches!(
                upper_name.as_str(),
                "IFERROR"
                    | "ISERROR"
                    | "ISNA"
                    | "ISERR"
                    | "ERROR.TYPE"
                    | "TYPE"
                    | "ISTEXT"
                    | "ISNONTEXT"
                    | "ISNUMBER"
                    | "ISLOGICAL"
                    | "ISBLANK"
            );
            if !inspects_errors
                // COUNTA counts an error argument as one more non-blank
                // value, and COUNT skips it, rather than either
                // propagating it (both match real Excel).
                && upper_name != "COUNTA"
                && upper_name != "COUNT"
                // COUNTBLANK just asks which cells are empty; an error in
                // the range is a non-blank cell, not a reason to fail.
                && upper_name != "COUNTBLANK"
                // AGGREGATE decides for itself whether to propagate or
                // ignore an error in its data, based on its `options`
                // argument, so it must see the raw arguments.
                && upper_name != "AGGREGATE"
                // The paired statistical functions check their two ranges'
                // shapes before anything else -- a size mismatch is #N/A
                // even when a range also holds an error value -- so they
                // re-raise errors themselves (see paired_args).
                && !matches!(
                    upper_name.as_str(),
                    "CORREL"
                        | "PEARSON"
                        | "COVAR"
                        | "COVARIANCE.P"
                        | "COVARIANCE.S"
                        | "SLOPE"
                        | "INTERCEPT"
                        | "RSQ"
                        | "STEYX"
                        | "FORECAST"
                        | "FORECAST.LINEAR"
                        | "SUMX2MY2"
                        | "SUMX2PY2"
                        | "SUMXMY2"
                        | "CHISQ.TEST"
                        | "CHITEST"
                )
                && !uses_ordered_arg_error_check
                && let Some(err) = Self::find_error_in_args(&evaluated_args)
            {
                return Ok(err);
            }

            let res_to_rd = |res: Result<f64, String>| -> Result<ResultData, EngineError> {
                match res {
                    Ok(v) => Ok(ResultData::Float(v)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            };

            // A NaN can only come from a math function evaluated outside
            // its domain (ASIN/ACOS of |x|>1, SQRT/LN/LOG10 of a negative,
            // ...), and an infinity only from one that overflowed
            // (POWER(42, 600), EXP(1000)). Excel has neither -- it reports
            // #NUM! for both -- so rather than bolting a domain/overflow
            // guard onto each of those call sites individually, normalize
            // here at the single point every function result flows
            // through.
            let dispatched = match upper_name.as_str() {
                // --- STATISTICAL FUNCTIONS ---
                "AVEDEV" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::avedev(&nums))
                }
                "AVERAGEA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    if nums.is_empty() {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else {
                        Ok(ResultData::Float(
                            nums.iter().sum::<f64>() / nums.len() as f64,
                        ))
                    }
                }
                "AVERAGEIF" => {
                    if evaluated_args.len() < 2 {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let range_list = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Error("#DIV/0!".to_string())),
                    };
                    let criteria = &evaluated_args[1];
                    let avg_range = if evaluated_args.len() >= 3 {
                        match &evaluated_args[2] {
                            ResultData::List(l) => l,
                            _ => range_list,
                        }
                    } else {
                        range_list
                    };
                    let mut sum = 0.0;
                    let mut count = 0;
                    for (i, val) in range_list.iter().enumerate() {
                        if self.match_criteria(val, criteria)
                            && let Some(target_val) = avg_range.get(i)
                            && let Some(f) = Self::aggregate_range_number(target_val)
                        {
                            sum += f;
                            count += 1;
                        }
                    }
                    if count == 0 {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else {
                        Ok(ResultData::Float(sum / count as f64))
                    }
                }
                "AVERAGEIFS" => {
                    if evaluated_args.len() < 3 || (evaluated_args.len() - 1) % 2 != 0 {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let avg_range = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Error("#DIV/0!".to_string())),
                    };
                    let mut criteria_pairs = Vec::new();
                    let mut i = 1;
                    while i < evaluated_args.len() {
                        let crit_range = match &evaluated_args[i] {
                            ResultData::List(l) => l,
                            _ => return Ok(ResultData::Error("#DIV/0!".to_string())),
                        };
                        let crit_val = &evaluated_args[i + 1];
                        criteria_pairs.push((crit_range, crit_val));
                        i += 2;
                    }
                    let mut sum = 0.0;
                    let mut count = 0;
                    for (idx, target_val) in avg_range.iter().enumerate() {
                        let mut all_match = true;
                        for (crit_range, crit_val) in &criteria_pairs {
                            if idx >= crit_range.len()
                                || !self.match_criteria(&crit_range[idx], crit_val)
                            {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match && let Some(f) = Self::aggregate_range_number(target_val) {
                            sum += f;
                            count += 1;
                        }
                    }
                    if count == 0 {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else {
                        Ok(ResultData::Float(sum / count as f64))
                    }
                }
                "BETA.DIST" | "BETADIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "BETA.DIST")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "BETA.DIST")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "BETA.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    let a = evaluated_args
                        .get(4)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(0.0);
                    let b = evaluated_args
                        .get(5)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(1.0);
                    res_to_rd(crate::core::stats::beta_dist(
                        x, alpha, beta, cumulative, a, b,
                    ))
                }
                "BETA.INV" | "BETAINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "BETA.INV")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "BETA.INV")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "BETA.INV")?;
                    let a = evaluated_args
                        .get(3)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(0.0);
                    let b = evaluated_args
                        .get(4)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(1.0);
                    res_to_rd(crate::core::stats::beta_inv(p, alpha, beta, a, b))
                }
                "BINOM.DIST" | "BINOMDIST" => {
                    let k = self.to_f64_arg(evaluated_args.first(), "BINOM.DIST")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "BINOM.DIST")?;
                    let p = self.to_f64_arg(evaluated_args.get(2), "BINOM.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(false);
                    res_to_rd(crate::core::stats::binom_dist(k, n, p, cumulative))
                }
                "BINOM.DIST.RANGE" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "BINOM.DIST.RANGE")?;
                    let p = self.to_f64_arg(evaluated_args.get(1), "BINOM.DIST.RANGE")?;
                    let k1 = self.to_f64_arg(evaluated_args.get(2), "BINOM.DIST.RANGE")?;
                    let k2 = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::stats::binom_dist_range(n, p, k1, k2))
                }
                "BINOM.INV" | "CRITBINOM" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "BINOM.INV")?;
                    let p = self.to_f64_arg(evaluated_args.get(1), "BINOM.INV")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(2), "BINOM.INV")?;
                    res_to_rd(crate::core::stats::binom_inv(n, p, alpha))
                }
                "CHISQ.DIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "CHISQ.DIST")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.DIST")?;
                    let cumulative = evaluated_args
                        .get(2)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::chisq_dist(x, df, cumulative))
                }
                "CHISQ.DIST.RT" | "CHIDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "CHISQ.DIST.RT")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.DIST.RT")?;
                    res_to_rd(crate::core::stats::chisq_dist_rt(x, df))
                }
                "CHISQ.INV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "CHISQ.INV")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.INV")?;
                    res_to_rd(crate::core::stats::chisq_inv(p, df))
                }
                "CHISQ.INV.RT" | "CHIINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "CHISQ.INV.RT")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.INV.RT")?;
                    res_to_rd(crate::core::stats::chisq_inv_rt(p, df))
                }
                "CHISQ.TEST" | "CHITEST" => {
                    // Like the paired statistical functions, CHITEST
                    // compares its two ranges' *raw* cell counts first --
                    // a mismatch is #N/A even when a range also holds an
                    // error value. It does not, however, pairwise-exclude
                    // the way CORREL and friends do (Excel keeps the
                    // original dimensions when working out the degrees of
                    // freedom), so the values themselves still come from
                    // the lenient flatten.
                    // Same shape as the paired sums: a range holding no
                    // numeric value at all is #DIV/0!, while a range that
                    // merely loses every *pair* to exclusion still computes
                    // (the statistic is 0, so the answer is 1).
                    if self.paired_sum_has_no_numbers(evaluated_args.first())
                        || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                    {
                        return Ok(ResultData::Error("#DIV/0!".to_string()));
                    }
                    let mut first_err = None;
                    let a_raw = self.positional_numbers(evaluated_args.first(), &mut first_err);
                    let e_raw = self.positional_numbers(evaluated_args.get(1), &mut first_err);
                    if a_raw.len() != e_raw.len() {
                        return Ok(ResultData::Error("#N/A".to_string()));
                    }
                    // A single category leaves zero degrees of freedom, so
                    // there is no chi-square distribution to evaluate
                    // against and Excel reports #N/A. Judged on the *raw*
                    // range size: applying it after pairwise filtering
                    // would turn a two-cell pair that merely holds one text
                    // cell into #N/A, where Excel still reports the
                    // underlying #DIV/0!.
                    if a_raw.len() < 2 {
                        return Ok(ResultData::Error("#N/A".to_string()));
                    }
                    if let Some(e) = first_err {
                        return Ok(ResultData::Error(e));
                    }
                    // Values are taken pairwise so a non-numeric cell in
                    // one range can't leave the two sides different lengths
                    // and turn a computable call into a spurious #N/A --
                    // Excel still returns a value there (CHITEST over a
                    // 2-cell pair whose expected range holds one text cell
                    // computes rather than failing).
                    let (actual, expected) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    // Degrees of freedom come from the raw range size, not
                    // from how many pairs survived the filtering above.
                    res_to_rd(crate::core::stats::chisq_test(
                        &actual,
                        &expected,
                        a_raw.len(),
                    ))
                }
                "CONFIDENCE.NORM" | "CONFIDENCE" => {
                    let alpha = self.to_f64_arg(evaluated_args.first(), "CONFIDENCE.NORM")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(1), "CONFIDENCE.NORM")?;
                    let size = self.to_f64_arg(evaluated_args.get(2), "CONFIDENCE.NORM")?;
                    res_to_rd(crate::core::stats::confidence_norm(alpha, std_dev, size))
                }
                "CONFIDENCE.T" => {
                    let alpha = self.to_f64_arg(evaluated_args.first(), "CONFIDENCE.T")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(1), "CONFIDENCE.T")?;
                    let size = self.to_f64_arg(evaluated_args.get(2), "CONFIDENCE.T")?;
                    res_to_rd(crate::core::stats::confidence_t(alpha, std_dev, size))
                }
                "CORREL" | "PEARSON" => {
                    let (xs, ys) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::correl(&xs, &ys))
                }
                "COUNTBLANK" => {
                    let mut count = 0;
                    fn count_blank_rec(arg: &ResultData) -> usize {
                        match arg {
                            ResultData::None => 1,
                            ResultData::String(s) if s.is_empty() => 1,
                            ResultData::List(list) => list.iter().map(count_blank_rec).sum(),
                            _ => 0,
                        }
                    }
                    for arg in &evaluated_args {
                        count += count_blank_rec(arg);
                    }
                    Ok(ResultData::Float(count as f64))
                }
                "COVARIANCE.P" | "COVAR" => {
                    let (xs, ys) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::covariance_p(&xs, &ys))
                }
                "COVARIANCE.S" => {
                    let (xs, ys) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::covariance_s(&xs, &ys))
                }
                "DEVSQ" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::devsq(&nums))
                }
                "EXPON.DIST" | "EXPONDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "EXPON.DIST")?;
                    let lambda = self.to_f64_arg(evaluated_args.get(1), "EXPON.DIST")?;
                    let cumulative = evaluated_args
                        .get(2)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::expon_dist(x, lambda, cumulative))
                }
                "F.DIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "F.DIST")?;
                    let df1 = self.to_f64_arg(evaluated_args.get(1), "F.DIST")?;
                    let df2 = self.to_f64_arg(evaluated_args.get(2), "F.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::f_dist(x, df1, df2, cumulative))
                }
                "F.DIST.RT" | "FDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "F.DIST.RT")?;
                    let df1 = self.to_f64_arg(evaluated_args.get(1), "F.DIST.RT")?;
                    let df2 = self.to_f64_arg(evaluated_args.get(2), "F.DIST.RT")?;
                    res_to_rd(crate::core::stats::f_dist_rt(x, df1, df2))
                }
                "F.INV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "F.INV")?;
                    let df1 = self.to_f64_arg(evaluated_args.get(1), "F.INV")?;
                    let df2 = self.to_f64_arg(evaluated_args.get(2), "F.INV")?;
                    res_to_rd(crate::core::stats::f_inv(p, df1, df2))
                }
                "F.INV.RT" | "FINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "F.INV.RT")?;
                    let df1 = self.to_f64_arg(evaluated_args.get(1), "F.INV.RT")?;
                    let df2 = self.to_f64_arg(evaluated_args.get(2), "F.INV.RT")?;
                    res_to_rd(crate::core::stats::f_inv_rt(p, df1, df2))
                }
                "F.TEST" | "FTEST" => {
                    let array1: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let array2: Vec<f64> = evaluated_args
                        .get(1)
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    res_to_rd(crate::core::stats::f_test(&array1, &array2))
                }
                "FISHER" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "FISHER")?;
                    res_to_rd(crate::core::stats::fisher(x))
                }
                "FISHERINV" => {
                    let y = self.to_f64_arg(evaluated_args.first(), "FISHERINV")?;
                    res_to_rd(crate::core::stats::fisherinv(y))
                }
                "FORECAST" | "FORECAST.LINEAR" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "FORECAST")?;
                    let (ys, xs) =
                        match self.paired_args(evaluated_args.get(1), evaluated_args.get(2)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::forecast_linear(x, &ys, &xs))
                }
                "FORECAST.ETS" | "FORECAST.ETS.CONFINT" => {
                    // FORECAST.ETS(target, values, timeline,
                    //              [seasonality], [data_completion], [aggregation])
                    // FORECAST.ETS.CONFINT(target, values, timeline,
                    //              [confidence], [seasonality], [data_completion], [aggregation])
                    let is_confint = upper_name == "FORECAST.ETS.CONFINT";
                    let target = self.to_f64_arg(evaluated_args.first(), "FORECAST.ETS")?;
                    let (values, timeline) =
                        match self.paired_args(evaluated_args.get(1), evaluated_args.get(2)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    let (confidence, seasonality_idx) = if is_confint {
                        (self.opt_f64_arg(&evaluated_args, 3, 0.95)?, 4)
                    } else {
                        (0.95, 3)
                    };
                    let seasonality = self.opt_f64_arg(&evaluated_args, seasonality_idx, 1.0)?;
                    let completion =
                        self.opt_f64_arg(&evaluated_args, seasonality_idx + 1, 1.0)? != 0.0;

                    let series =
                        match crate::core::ets::build_series(&values, &timeline, completion) {
                            Ok(s) => s,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    let h = match crate::core::ets::horizon(
                        series.start,
                        series.step,
                        series.values.len(),
                        target,
                    ) {
                        Ok(h) => h,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    let model = match crate::core::ets::prepare(
                        &values,
                        &timeline,
                        seasonality,
                        completion,
                    ) {
                        Ok(m) => m,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    if is_confint {
                        res_to_rd(model.confint(h, confidence))
                    } else {
                        Ok(ResultData::Float(model.forecast(h)))
                    }
                }
                "FORECAST.ETS.SEASONALITY" => {
                    // FORECAST.ETS.SEASONALITY(values, timeline,
                    //                          [data_completion], [aggregation])
                    // -- note there is no leading target-date argument.
                    let (values, timeline) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    let completion = self.opt_f64_arg(&evaluated_args, 2, 1.0)? != 0.0;
                    match crate::core::ets::build_series(&values, &timeline, completion) {
                        Ok(series) => Ok(ResultData::Float(crate::core::ets::detect_period(
                            &series.values,
                        ) as f64)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "FORECAST.ETS.STAT" => {
                    // FORECAST.ETS.STAT(values, timeline, statistic_type,
                    //                   [seasonality], [data_completion], [aggregation])
                    let (values, timeline) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    let which = self.to_f64_arg(evaluated_args.get(2), "FORECAST.ETS.STAT")?;
                    let seasonality = self.opt_f64_arg(&evaluated_args, 3, 1.0)?;
                    let completion = self.opt_f64_arg(&evaluated_args, 4, 1.0)? != 0.0;
                    match crate::core::ets::prepare(&values, &timeline, seasonality, completion) {
                        Ok(model) => res_to_rd(model.stat(which.round() as usize)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "FREQUENCY" => {
                    let data: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let bins: Vec<f64> = evaluated_args
                        .get(1)
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    match crate::core::stats::frequency(&data, &bins) {
                        Ok(counts) => Ok(ResultData::List(
                            counts.into_iter().map(ResultData::Float).collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "GAMMA" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "GAMMA")?;
                    let val = crate::core::stats::gamma(x);
                    if val.is_nan() {
                        Ok(ResultData::Error("#NUM!".to_string()))
                    } else {
                        Ok(ResultData::Float(val))
                    }
                }
                "GAMMA.DIST" | "GAMMADIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "GAMMA.DIST")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "GAMMA.DIST")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "GAMMA.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::gamma_dist(x, alpha, beta, cumulative))
                }
                "GAMMA.INV" | "GAMMAINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "GAMMA.INV")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "GAMMA.INV")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "GAMMA.INV")?;
                    res_to_rd(crate::core::stats::gamma_inv(p, alpha, beta))
                }
                "GAMMALN" | "GAMMALN.PRECISE" => {
                    // ln(Gamma(x)) is only defined for x > 0 in Excel --
                    // GAMMALN(-5), GAMMALN(0) and GAMMALN of a large
                    // negative are all #NUM!. The underlying lgamma here
                    // uses the reflection formula and happily returns a
                    // value for negative non-integers, so the domain has to
                    // be enforced at the boundary.
                    let x = self.to_f64_arg(evaluated_args.first(), "GAMMALN")?;
                    if x <= 0.0 {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    let val = crate::core::stats::lgamma(x);
                    if val.is_nan() {
                        Ok(ResultData::Error("#NUM!".to_string()))
                    } else {
                        Ok(ResultData::Float(val))
                    }
                }
                "GAUSS" => {
                    let z = self.to_f64_arg(evaluated_args.first(), "GAUSS")?;
                    res_to_rd(crate::core::stats::gauss(z))
                }
                "GEOMEAN" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::geomean(&nums))
                }
                "GROWTH" | "LOGEST" => {
                    // LINEST/TREND/GROWTH/LOGEST are the *array* form of
                    // the regression family and, unlike scalar FORECAST
                    // (which drops a non-numeric pair and carries on),
                    // real Excel rejects any non-numeric cell outright
                    // with #VALUE! -- confirmed by probing all five
                    // against the same text-containing range.
                    let ys = match self.flatten_numbers_only_arg(evaluated_args.first()) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    let xs = match evaluated_args.get(1) {
                        Some(arg) => match self.flatten_numbers_only(arg) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        },
                        None => (1..=ys.len()).map(|i| i as f64).collect(),
                    };
                    let ln_ys: Vec<f64> = ys.iter().map(|y| y.ln()).collect();
                    let m = match crate::core::stats::slope(&ln_ys, &xs) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    let b = match crate::core::stats::intercept(&ln_ys, &xs) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    if upper_name == "LOGEST" {
                        Ok(ResultData::List(vec![
                            ResultData::Float(m.exp()),
                            ResultData::Float(b.exp()),
                        ]))
                    } else {
                        let new_x = evaluated_args
                            .get(2)
                            .and_then(|v| self.to_f64(v))
                            .unwrap_or_else(|| xs.first().copied().unwrap_or(1.0));
                        Ok(ResultData::Float((b + m * new_x).exp()))
                    }
                }
                "HARMEAN" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::harmean(&nums))
                }
                "HYPGEOM.DIST" | "HYPGEOMDIST" => {
                    let sample_s = self.to_f64_arg(evaluated_args.first(), "HYPGEOM.DIST")?;
                    let sample_size = self.to_f64_arg(evaluated_args.get(1), "HYPGEOM.DIST")?;
                    let pop_s = self.to_f64_arg(evaluated_args.get(2), "HYPGEOM.DIST")?;
                    let pop_size = self.to_f64_arg(evaluated_args.get(3), "HYPGEOM.DIST")?;
                    // Legacy HYPGEOMDIST takes no cumulative flag at all --
                    // it's always the point probability mass, never the
                    // cumulative sum (unlike HYPGEOM.DIST, whose 5th
                    // argument is required and selects between the two).
                    let cumulative = if upper_name == "HYPGEOMDIST" {
                        false
                    } else {
                        evaluated_args
                            .get(4)
                            .map(|v| self.to_bool(v))
                            .unwrap_or(true)
                    };
                    res_to_rd(crate::core::stats::hypgeom_dist(
                        sample_s,
                        sample_size,
                        pop_s,
                        pop_size,
                        cumulative,
                    ))
                }
                "INTERCEPT" => {
                    let (ys, xs) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::intercept(&ys, &xs))
                }
                "KURT" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::kurt(&nums))
                }
                "LARGE" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "LARGE")?.round() as usize;
                    res_to_rd(crate::core::stats::large(&nums, k))
                }
                "LINEST" | "TREND" => {
                    // LINEST/TREND/GROWTH/LOGEST are the *array* form of
                    // the regression family and, unlike scalar FORECAST
                    // (which drops a non-numeric pair and carries on),
                    // real Excel rejects any non-numeric cell outright
                    // with #VALUE! -- confirmed by probing all five
                    // against the same text-containing range.
                    let ys = match self.flatten_numbers_only_arg(evaluated_args.first()) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    let xs = match evaluated_args.get(1) {
                        Some(arg) => match self.flatten_numbers_only(arg) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        },
                        None => (1..=ys.len()).map(|i| i as f64).collect(),
                    };
                    let m = match crate::core::stats::slope(&ys, &xs) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    let b = match crate::core::stats::intercept(&ys, &xs) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    if upper_name == "LINEST" {
                        Ok(ResultData::List(vec![
                            ResultData::Float(m),
                            ResultData::Float(b),
                        ]))
                    } else {
                        let new_x = evaluated_args
                            .get(2)
                            .and_then(|v| self.to_f64(v))
                            .unwrap_or_else(|| xs.first().copied().unwrap_or(1.0));
                        Ok(ResultData::Float(m * new_x + b))
                    }
                }
                "LOGNORM.DIST" | "LOGNORMDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "LOGNORM.DIST")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "LOGNORM.DIST")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "LOGNORM.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::lognorm_dist(
                        x, mean, std_dev, cumulative,
                    ))
                }
                "LOGNORM.INV" | "LOGINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "LOGNORM.INV")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "LOGNORM.INV")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "LOGNORM.INV")?;
                    res_to_rd(crate::core::stats::lognorm_inv(p, mean, std_dev))
                }
                "MAXA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    if nums.is_empty() {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(
                            nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max),
                        ))
                    }
                }
                "MAXIFS" => {
                    if evaluated_args.len() < 3 || (evaluated_args.len() - 1) % 2 != 0 {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let max_range = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let mut criteria_pairs = Vec::new();
                    let mut i = 1;
                    while i < evaluated_args.len() {
                        let crit_range = match &evaluated_args[i] {
                            ResultData::List(l) => l,
                            _ => return Ok(ResultData::Float(0.0)),
                        };
                        let crit_val = &evaluated_args[i + 1];
                        criteria_pairs.push((crit_range, crit_val));
                        i += 2;
                    }
                    let mut max_val = f64::NEG_INFINITY;
                    let mut found = false;
                    for (idx, target_val) in max_range.iter().enumerate() {
                        let mut all_match = true;
                        for (crit_range, crit_val) in &criteria_pairs {
                            if idx >= crit_range.len()
                                || !self.match_criteria(&crit_range[idx], crit_val)
                            {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match && let Some(f) = Self::aggregate_range_number(target_val) {
                            max_val = max_val.max(f);
                            found = true;
                        }
                    }
                    if !found {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(max_val))
                    }
                }
                "MEDIAN" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::median(&nums))
                }
                "MINA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    if nums.is_empty() {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(
                            nums.iter().cloned().fold(f64::INFINITY, f64::min),
                        ))
                    }
                }
                "MINIFS" => {
                    if evaluated_args.len() < 3 || (evaluated_args.len() - 1) % 2 != 0 {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let min_range = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let mut criteria_pairs = Vec::new();
                    let mut i = 1;
                    while i < evaluated_args.len() {
                        let crit_range = match &evaluated_args[i] {
                            ResultData::List(l) => l,
                            _ => return Ok(ResultData::Float(0.0)),
                        };
                        let crit_val = &evaluated_args[i + 1];
                        criteria_pairs.push((crit_range, crit_val));
                        i += 2;
                    }
                    let mut min_val = f64::INFINITY;
                    let mut found = false;
                    for (idx, target_val) in min_range.iter().enumerate() {
                        let mut all_match = true;
                        for (crit_range, crit_val) in &criteria_pairs {
                            if idx >= crit_range.len()
                                || !self.match_criteria(&crit_range[idx], crit_val)
                            {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match && let Some(f) = Self::aggregate_range_number(target_val) {
                            min_val = min_val.min(f);
                            found = true;
                        }
                    }
                    if !found {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(min_val))
                    }
                }
                "MODE.MULT" => {
                    // The MODE family rejects a lone blank operand where
                    // its neighbours tolerate it -- MODE(x, <blank>) is
                    // #VALUE! while MEDIAN(x, <blank>) is x. Applies to
                    // all three spellings. See is_empty_scalar_operand.
                    if evaluated_args.iter().any(Self::is_empty_scalar_operand) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    match crate::core::stats::mode_mult(&nums) {
                        Ok(modes) => Ok(ResultData::List(
                            modes.into_iter().map(ResultData::Float).collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "MODE.SNGL" | "MODE" => {
                    // The MODE family rejects a lone blank operand where
                    // its neighbours tolerate it -- MODE(x, <blank>) is
                    // #VALUE! while MEDIAN(x, <blank>) is x. Applies to
                    // all three spellings. See is_empty_scalar_operand.
                    if evaluated_args.iter().any(Self::is_empty_scalar_operand) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::mode_sngl(&nums))
                }
                "NEGBINOM.DIST" | "NEGBINOMDIST" => {
                    let k = self.to_f64_arg(evaluated_args.first(), "NEGBINOM.DIST")?;
                    let r = self.to_f64_arg(evaluated_args.get(1), "NEGBINOM.DIST")?;
                    let p = self.to_f64_arg(evaluated_args.get(2), "NEGBINOM.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(false);
                    res_to_rd(crate::core::stats::negbinom_dist(k, r, p, cumulative))
                }
                "NORM.DIST" | "NORMDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "NORM.DIST")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "NORM.DIST")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "NORM.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::norm_dist(x, mean, std_dev, cumulative))
                }
                "NORM.INV" | "NORMINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "NORM.INV")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "NORM.INV")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "NORM.INV")?;
                    res_to_rd(crate::core::stats::norm_inv(p, mean, std_dev))
                }
                "NORM.S.DIST" | "NORMSDIST" => {
                    let z = self.to_f64_arg(evaluated_args.first(), "NORM.S.DIST")?;
                    let cumulative = evaluated_args
                        .get(1)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::norm_s_dist(z, cumulative))
                }
                "NORM.S.INV" | "NORMSINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "NORM.S.INV")?;
                    res_to_rd(crate::core::stats::norm_s_inv(p))
                }
                "PERCENTILE.EXC" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "PERCENTILE.EXC")?;
                    res_to_rd(crate::core::stats::percentile_exc(&nums, k))
                }
                "PERCENTILE.INC" | "PERCENTILE" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "PERCENTILE.INC")?;
                    res_to_rd(crate::core::stats::percentile_inc(&nums, k))
                }
                "PERCENTRANK.EXC" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let x = self.to_f64_arg(evaluated_args.get(1), "PERCENTRANK.EXC")?;
                    let sig = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(3.0) as usize;
                    res_to_rd(crate::core::stats::percentrank_exc(&nums, x, sig))
                }
                "PERCENTRANK.INC" | "PERCENTRANK" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let x = self.to_f64_arg(evaluated_args.get(1), "PERCENTRANK.INC")?;
                    let sig = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(3.0) as usize;
                    res_to_rd(crate::core::stats::percentrank_inc(&nums, x, sig))
                }
                "PERMUT" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "PERMUT")?;
                    let k = self.to_f64_arg(evaluated_args.get(1), "PERMUT")?;
                    res_to_rd(crate::core::stats::permut(n, k))
                }
                "PERMUTATIONA" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "PERMUTATIONA")?;
                    let k = self.to_f64_arg(evaluated_args.get(1), "PERMUTATIONA")?;
                    res_to_rd(crate::core::stats::permutationa(n, k))
                }
                "PHI" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "PHI")?;
                    res_to_rd(crate::core::stats::phi(x))
                }
                "POISSON.DIST" | "POISSON" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "POISSON.DIST")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "POISSON.DIST")?;
                    let cumulative = evaluated_args
                        .get(2)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::poisson_dist(x, mean, cumulative))
                }
                "PROB" => {
                    let (x_range, prob_range) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    let lower = self.to_f64_arg(evaluated_args.get(2), "PROB")?;
                    let upper = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::stats::prob(
                        &x_range,
                        &prob_range,
                        lower,
                        upper,
                    ))
                }
                "QUARTILE.EXC" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let q = self
                        .to_f64_arg(evaluated_args.get(1), "QUARTILE.EXC")?
                        .round() as usize;
                    res_to_rd(crate::core::stats::quartile_exc(&nums, q))
                }
                "QUARTILE.INC" | "QUARTILE" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let q = self
                        .to_f64_arg(evaluated_args.get(1), "QUARTILE.INC")?
                        .round() as usize;
                    res_to_rd(crate::core::stats::quartile_inc(&nums, q))
                }
                "RANK.AVG" => {
                    let number = self.to_f64_arg(evaluated_args.first(), "RANK.AVG")?;
                    let ref_data: Vec<f64> = evaluated_args
                        .get(1)
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let order = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(0.0) as usize;
                    res_to_rd(crate::core::stats::rank_avg(number, &ref_data, order))
                }
                "RANK.EQ" | "RANK" => {
                    let number = self.to_f64_arg(evaluated_args.first(), "RANK.EQ")?;
                    let ref_data: Vec<f64> = evaluated_args
                        .get(1)
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let order = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(0.0) as usize;
                    res_to_rd(crate::core::stats::rank_eq(number, &ref_data, order))
                }
                "RSQ" => {
                    let (ys, xs) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::rsq(&ys, &xs))
                }
                "SKEW" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::skew(&nums))
                }
                "SKEW.P" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::skew_p(&nums))
                }
                "SLOPE" => {
                    let (ys, xs) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::slope(&ys, &xs))
                }
                "SMALL" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "SMALL")?.round() as usize;
                    res_to_rd(crate::core::stats::small(&nums, k))
                }
                "STANDARDIZE" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "STANDARDIZE")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "STANDARDIZE")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "STANDARDIZE")?;
                    res_to_rd(crate::core::stats::standardize(x, mean, std_dev))
                }
                "STDEV.P" | "STDEVP" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::stdev_p(&nums))
                }
                "STDEV.S" | "STDEV" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::stdev_s(&nums))
                }
                "STDEVA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::stdev_s(&nums))
                }
                "STDEVPA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::stdev_p(&nums))
                }
                "STEYX" => {
                    let (ys, xs) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::steyx(&ys, &xs))
                }
                "T.DIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "T.DIST")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST")?;
                    let cumulative = evaluated_args
                        .get(2)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::t_dist(x, df, cumulative))
                }
                "T.DIST.2T" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "T.DIST.2T")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST.2T")?;
                    res_to_rd(crate::core::stats::t_dist_2t(x, df))
                }
                "TDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "TDIST")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "TDIST")?;
                    let tails = self.to_f64_arg(evaluated_args.get(2), "TDIST")?;
                    if tails == 1.0 {
                        res_to_rd(crate::core::stats::t_dist_rt(x, df))
                    } else {
                        res_to_rd(crate::core::stats::t_dist_2t(x, df))
                    }
                }
                "T.DIST.RT" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "T.DIST.RT")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST.RT")?;
                    res_to_rd(crate::core::stats::t_dist_rt(x, df))
                }
                "T.INV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "T.INV")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.INV")?;
                    res_to_rd(crate::core::stats::t_inv(p, df))
                }
                "T.INV.2T" | "TINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "T.INV.2T")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.INV.2T")?;
                    res_to_rd(crate::core::stats::t_inv_2t(p, df))
                }
                "T.TEST" | "TTEST" => {
                    let tails = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(2.0) as usize;
                    let test_type = evaluated_args
                        .get(3)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(1.0) as usize;
                    // Only test_type 1 is the *paired* test, where the two
                    // arrays must be the same size (#N/A otherwise) and a
                    // non-numeric cell drops its whole (x, y) pair. Types 2
                    // and 3 are two-*sample* tests that compare two
                    // independent groups, so they accept different sizes
                    // and each array drops its own non-numerics
                    // independently. Both confirmed against real Excel:
                    // `TTEST(4-cell-with-text, 4-cell, 1, 2)` equals
                    // `TTEST(full-4-cell, 3-cell-survivor, 1, 2)`, while
                    // the same call with type 1 instead equals the
                    // 3-vs-3 pairwise-survivor form.
                    let (array1, array2) = if test_type == 1 {
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        }
                    } else {
                        (
                            evaluated_args
                                .first()
                                .map(|arg| self.flatten_stat_numbers(arg, false))
                                .unwrap_or_default(),
                            evaluated_args
                                .get(1)
                                .map(|arg| self.flatten_stat_numbers(arg, false))
                                .unwrap_or_default(),
                        )
                    };
                    res_to_rd(crate::core::stats::t_test(
                        &array1, &array2, tails, test_type,
                    ))
                }
                "TRIMMEAN" => {
                    let nums: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let percent = self.to_f64_arg(evaluated_args.get(1), "TRIMMEAN")?;
                    res_to_rd(crate::core::stats::trimmean(&nums, percent))
                }
                "VAR.P" | "VARP" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::var_p(&nums))
                }
                "VAR.S" | "VAR" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::var_s(&nums))
                }
                "VARA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::var_s(&nums))
                }
                "VARPA" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers_a(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::stats::var_p(&nums))
                }
                "WEIBULL.DIST" | "WEIBULL" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "WEIBULL.DIST")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "WEIBULL.DIST")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "WEIBULL.DIST")?;
                    let cumulative = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    res_to_rd(crate::core::stats::weibull_dist(x, alpha, beta, cumulative))
                }
                "Z.TEST" | "ZTEST" => {
                    let array: Vec<f64> = evaluated_args
                        .first()
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    let x = self.to_f64_arg(evaluated_args.get(1), "Z.TEST")?;
                    let sigma = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::stats::z_test(&array, x, sigma))
                }

                // --- MATH AND TRIGONOMETRY FUNCTIONS ---
                "ACOSH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ACOSH")?;
                    res_to_rd(crate::core::math_trig::acosh(x))
                }
                "ACOT" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ACOT")?;
                    res_to_rd(crate::core::math_trig::acot(x))
                }
                "ACOTH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ACOTH")?;
                    res_to_rd(crate::core::math_trig::acoth(x))
                }
                "AGGREGATE" => {
                    // AGGREGATE(function_num, options, ref1, ...) -- unlike
                    // SUBTOTAL(function_num, ref1, ...), its *second*
                    // argument is the options flag, not data. Sharing
                    // SUBTOTAL's handler (which skips only the first
                    // argument) folded that options value straight into
                    // the aggregated numbers, so e.g. AGGREGATE(4, 6, ...)
                    // computed MAX over the data *plus a literal 6*.
                    let fn_num = self
                        .to_f64_arg(evaluated_args.first(), "AGGREGATE")?
                        .round() as usize;
                    let options = evaluated_args
                        .get(1)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(0.0)
                        .round() as usize;
                    // Function numbers 14-19 (LARGE/SMALL/PERCENTILE.INC/
                    // QUARTILE.INC/PERCENTILE.EXC/QUARTILE.EXC) take a
                    // trailing k argument after the array.
                    let takes_k = (14..=19).contains(&fn_num);
                    let data_end = if takes_k {
                        evaluated_args.len().saturating_sub(1)
                    } else {
                        evaluated_args.len()
                    };
                    let k = if takes_k {
                        evaluated_args
                            .last()
                            .and_then(|v| self.to_f64(v))
                            .unwrap_or(1.0)
                    } else {
                        1.0
                    };
                    // Options 2/3/6/7 mean "ignore error values"; every
                    // option this engine can express other than that still
                    // propagates an error in the data, matching Excel.
                    let ignores_errors = matches!(options, 2 | 3 | 6 | 7);
                    let data_args = &evaluated_args[2.min(evaluated_args.len())..data_end];
                    if !ignores_errors && let Some(err) = Self::find_error_in_args(data_args) {
                        return Ok(err);
                    }
                    let nums: Vec<f64> = data_args
                        .iter()
                        .flat_map(|arg| self.flatten_stat_numbers(arg, false))
                        .collect();
                    match fn_num {
                        1 => res_to_rd(if nums.is_empty() {
                            Err("#DIV/0!".to_string())
                        } else {
                            Ok(nums.iter().sum::<f64>() / nums.len() as f64)
                        }),
                        2 | 3 => Ok(ResultData::Float(nums.len() as f64)),
                        // MAX/MIN over nothing is 0, not an infinity --
                        // which the dispatch-level NaN/infinity guard would
                        // otherwise turn into #NUM!.
                        4 => Ok(ResultData::Float(if nums.is_empty() {
                            0.0
                        } else {
                            nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
                        })),
                        5 => Ok(ResultData::Float(if nums.is_empty() {
                            0.0
                        } else {
                            nums.iter().cloned().fold(f64::INFINITY, f64::min)
                        })),
                        6 => Ok(ResultData::Float(nums.iter().product())),
                        7 => res_to_rd(crate::core::stats::stdev_s(&nums)),
                        8 => res_to_rd(crate::core::stats::stdev_p(&nums)),
                        9 => Ok(ResultData::Float(nums.iter().sum())),
                        10 => res_to_rd(crate::core::stats::var_s(&nums)),
                        11 => res_to_rd(crate::core::stats::var_p(&nums)),
                        12 => res_to_rd(crate::core::stats::median(&nums)),
                        13 => res_to_rd(crate::core::stats::mode_sngl(&nums)),
                        14 => res_to_rd(crate::core::stats::large(&nums, k.round() as usize)),
                        15 => res_to_rd(crate::core::stats::small(&nums, k.round() as usize)),
                        16 => res_to_rd(crate::core::stats::percentile_inc(&nums, k)),
                        17 => {
                            res_to_rd(crate::core::stats::quartile_inc(&nums, k.round() as usize))
                        }
                        18 => res_to_rd(crate::core::stats::percentile_exc(&nums, k)),
                        19 => {
                            res_to_rd(crate::core::stats::quartile_exc(&nums, k.round() as usize))
                        }
                        _ => Ok(ResultData::Error("#VALUE!".to_string())),
                    }
                }
                "SUBTOTAL" => {
                    let fn_num =
                        self.to_f64_arg(evaluated_args.first(), "SUBTOTAL")?.round() as usize;
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .skip(1)
                        .flat_map(|arg| self.flatten_stat_numbers(arg, false))
                        .collect();
                    match fn_num % 100 {
                        1 => res_to_rd(if nums.is_empty() {
                            Err("#DIV/0!".to_string())
                        } else {
                            Ok(nums.iter().sum::<f64>() / nums.len() as f64)
                        }),
                        2 | 3 => Ok(ResultData::Float(nums.len() as f64)),
                        // MAX/MIN over nothing is 0, matching plain
                        // MAX/MIN (and not an infinity, which the
                        // dispatch-level guard would turn into #NUM!).
                        4 => Ok(ResultData::Float(if nums.is_empty() {
                            0.0
                        } else {
                            nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
                        })),
                        5 => Ok(ResultData::Float(if nums.is_empty() {
                            0.0
                        } else {
                            nums.iter().cloned().fold(f64::INFINITY, f64::min)
                        })),
                        6 => Ok(ResultData::Float(nums.iter().product())),
                        7 => res_to_rd(crate::core::stats::stdev_s(&nums)),
                        8 => res_to_rd(crate::core::stats::stdev_p(&nums)),
                        9 => Ok(ResultData::Float(nums.iter().sum())),
                        10 => res_to_rd(crate::core::stats::var_s(&nums)),
                        11 => res_to_rd(crate::core::stats::var_p(&nums)),
                        12 => res_to_rd(crate::core::stats::median(&nums)),
                        _ => Ok(ResultData::Float(nums.iter().sum())),
                    }
                }
                "ARABIC" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::math_trig::arabic(&text))
                }
                "ASINH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ASINH")?;
                    res_to_rd(crate::core::math_trig::asinh(x))
                }
                "ATAN2" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ATAN2")?;
                    let y = self.to_f64_arg(evaluated_args.get(1), "ATAN2")?;
                    res_to_rd(crate::core::math_trig::atan2(x, y))
                }
                "ATANH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ATANH")?;
                    res_to_rd(crate::core::math_trig::atanh(x))
                }
                "BASE" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "BASE")?;
                    let radix = self.to_f64_arg(evaluated_args.get(1), "BASE")?;
                    let min_len = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    match crate::core::math_trig::base(num, radix, min_len) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "CEILING.MATH" | "CEILING.PRECISE" | "ISO.CEILING" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "CEILING.MATH")?;
                    let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    let mode = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::math_trig::ceiling_math(x, sig, mode))
                }
                "COMBIN" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "COMBIN")?;
                    let k = self.to_f64_arg(evaluated_args.get(1), "COMBIN")?;
                    res_to_rd(crate::core::math_trig::combin(n, k))
                }
                "COMBINA" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "COMBINA")?;
                    let k = self.to_f64_arg(evaluated_args.get(1), "COMBINA")?;
                    res_to_rd(crate::core::math_trig::combina(n, k))
                }
                "COSH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "COSH")?;
                    res_to_rd(crate::core::math_trig::cosh(x))
                }
                "COT" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "COT")?;
                    res_to_rd(crate::core::math_trig::cot(x))
                }
                "COTH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "COTH")?;
                    res_to_rd(crate::core::math_trig::coth(x))
                }
                "CSC" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "CSC")?;
                    res_to_rd(crate::core::math_trig::csc(x))
                }
                "CSCH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "CSCH")?;
                    res_to_rd(crate::core::math_trig::csch(x))
                }
                "DECIMAL" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let radix = self.to_f64_arg(evaluated_args.get(1), "DECIMAL")?;
                    res_to_rd(crate::core::math_trig::decimal(&text, radix))
                }
                "DEGREES" => {
                    let rad = self.to_f64_arg(evaluated_args.first(), "DEGREES")?;
                    res_to_rd(crate::core::math_trig::degrees(rad))
                }
                "EVEN" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "EVEN")?;
                    res_to_rd(crate::core::math_trig::even(x))
                }
                "FACT" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "FACT")?;
                    res_to_rd(crate::core::math_trig::fact(n))
                }
                "FACTDOUBLE" => {
                    // FACTDOUBLE(TRUE) is #VALUE! even though
                    // FACTDOUBLE(1) is 1. See first_arg_is_boolean.
                    if Self::first_arg_is_boolean(&evaluated_args) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let n = self.to_f64_arg(evaluated_args.first(), "FACTDOUBLE")?;
                    res_to_rd(crate::core::math_trig::factdouble(n))
                }
                "FLOOR.MATH" | "FLOOR.PRECISE" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "FLOOR.MATH")?;
                    let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    let mode = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::math_trig::floor_math(x, sig, mode))
                }
                "GCD" | "LCM" => {
                    let mut nums = Vec::new();
                    for arg in &evaluated_args {
                        match self.flatten_strict_numbers(arg) {
                            Ok(v) => nums.extend(v),
                            Err(e) => return Ok(ResultData::Error(e)),
                        }
                    }
                    if upper_name == "GCD" {
                        res_to_rd(crate::core::math_trig::gcd(&nums))
                    } else {
                        res_to_rd(crate::core::math_trig::lcm(&nums))
                    }
                }
                "LOG" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "LOG")?;
                    let base = self.opt_f64_arg(&evaluated_args, 1, 10.0)?;
                    // Base 1 is #DIV/0!, not #NUM!: log(n)/log(1) divides
                    // by zero. Everything else out of domain stays #NUM!
                    // (both confirmed against real Excel).
                    if base == 1.0 {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else if num <= 0.0 || base <= 0.0 {
                        Ok(ResultData::Error("#NUM!".to_string()))
                    } else {
                        Ok(ResultData::Float(num.log(base)))
                    }
                }
                "MDETERM" => {
                    let matrix = match (args.first(), evaluated_args.first()) {
                        (Some(e), Some(v)) => self.matrix_from_arg(e, v),
                        _ => Vec::new(),
                    };
                    res_to_rd(crate::core::math_trig::mdeterm(&matrix))
                }
                "MINVERSE" => {
                    let matrix = match (args.first(), evaluated_args.first()) {
                        (Some(e), Some(v)) => self.matrix_from_arg(e, v),
                        _ => Vec::new(),
                    };
                    match crate::core::math_trig::minverse(&matrix) {
                        Ok(inv) => Ok(ResultData::List(
                            inv.into_iter()
                                .map(|row| {
                                    ResultData::List(
                                        row.into_iter().map(ResultData::Float).collect(),
                                    )
                                })
                                .collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "MROUND" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "MROUND")?;
                    let mult = self.to_f64_arg(evaluated_args.get(1), "MROUND")?;
                    res_to_rd(crate::core::math_trig::mround(x, mult))
                }
                "MULTINOMIAL" => {
                    // Like GCD/LCM, MULTINOMIAL rejects a non-numeric cell
                    // outright (#VALUE!) instead of skipping it the way
                    // SUM does -- a blank inside a range still counts as 0.
                    // ... and a blank operand is only a *missing* operand
                    // when there is nothing else: MULTINOMIAL(<blank>) and
                    // MULTINOMIAL(<blank>, <blank>) are #VALUE! while
                    // MULTINOMIAL(3, <blank>) is 1, the blank counting as
                    // 0. That is narrower than SUMPRODUCT, where any lone
                    // blank operand is #VALUE! even beside a number.
                    if !evaluated_args.is_empty()
                        && evaluated_args.iter().all(Self::is_empty_scalar_operand)
                    {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let mut nums = Vec::new();
                    for arg in &evaluated_args {
                        match self.flatten_strict_numbers(arg) {
                            Ok(v) => nums.extend(v),
                            Err(e) => return Ok(ResultData::Error(e)),
                        }
                    }
                    res_to_rd(crate::core::math_trig::multinomial(&nums))
                }
                "MUNIT" => {
                    let dim = self.to_f64_arg(evaluated_args.first(), "MUNIT")?;
                    match crate::core::math_trig::munit(dim) {
                        Ok(mat) => Ok(ResultData::List(
                            mat.into_iter()
                                .map(|row| {
                                    ResultData::List(
                                        row.into_iter().map(ResultData::Float).collect(),
                                    )
                                })
                                .collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "ODD" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ODD")?;
                    res_to_rd(crate::core::math_trig::odd(x))
                }
                "PERCENTOF" => {
                    // PERCENTOF(subset, all) is SUM(subset)/SUM(all), and
                    // it inherits SUM's leniency rather than erroring on a
                    // non-numeric argument: real Excel gives 0 for
                    // PERCENTOF(<text>, 10) (the numerator sums to 0) and
                    // #DIV/0! for PERCENTOF(10, <text>) or PERCENTOF(10, 0)
                    // (the denominator does). Routing both arguments
                    // through to_f64_arg instead made any text #VALUE!.
                    // Text *inside a referenced range* sums as 0 (so
                    // PERCENTOF(<text cell>, 10) is 0 and
                    // PERCENTOF(10, <text cell>) is #DIV/0!), but a
                    // directly-supplied non-numeric value -- a literal, or
                    // the result of a nested call like LOWER(...) -- is
                    // #VALUE!. That's the same direct-vs-reference split
                    // the SUM/AVERAGE helpers already make.
                    let mut sums = [0.0f64; 2];
                    for (i, slot) in sums.iter_mut().enumerate() {
                        let Some(v) = evaluated_args.get(i) else {
                            continue;
                        };
                        if arg_is_direct.get(i).copied().unwrap_or(false) {
                            match v {
                                ResultData::None => {}
                                other => match self.to_f64(other) {
                                    Some(f) => *slot = f,
                                    None => {
                                        return Ok(ResultData::Error("#VALUE!".to_string()));
                                    }
                                },
                            }
                        } else {
                            *slot = self.flatten_stat_numbers(v, false).iter().sum();
                        }
                    }
                    let [data_val, target_val] = sums;
                    if target_val == 0.0 {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else {
                        res_to_rd(crate::core::math_trig::percentof(data_val, target_val))
                    }
                }
                "PI" => Ok(ResultData::Float(std::f64::consts::PI)),
                "POWER" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "POWER")?;
                    let p = self.to_f64_arg(evaluated_args.get(1), "POWER")?;
                    res_to_rd(crate::core::math_trig::power(num, p))
                }
                "QUOTIENT" => {
                    // QUOTIENT rejects *booleans* but still coerces
                    // numeric text: QUOTIENT(12, TRUE) is #VALUE! while
                    // QUOTIENT("12", 5) is 2 and QUOTIENT(12, "ab") is
                    // #VALUE!. (MOD differs again -- MOD(TRUE, 2) is 1.)
                    // So this is to_f64's coercion with booleans excluded,
                    // not a numbers-only rule -- rejecting numeric strings
                    // too made QUOTIENT over a CONCATENATE/RIGHT result
                    // #VALUE! where Excel computes.
                    let coerce = |v: Option<&ResultData>| -> Option<f64> {
                        match v {
                            Some(ResultData::Boolean(_)) => None,
                            Some(other) => self.to_f64(other),
                            None => None,
                        }
                    };
                    let (Some(num), Some(den)) = (
                        coerce(evaluated_args.first()),
                        coerce(evaluated_args.get(1)),
                    ) else {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    };
                    res_to_rd(crate::core::math_trig::quotient(num, den))
                }
                "RADIANS" => {
                    let deg = self.to_f64_arg(evaluated_args.first(), "RADIANS")?;
                    res_to_rd(crate::core::math_trig::radians(deg))
                }
                "RANDARRAY" => {
                    let rows = evaluated_args.first().and_then(|v| self.to_f64(v));
                    let cols = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    let min = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    let max = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    let whole = evaluated_args.get(4).map(|v| self.to_bool(v));
                    match crate::core::math_trig::randarray(rows, cols, min, max, whole) {
                        Ok(grid) => Ok(ResultData::List(
                            grid.into_iter()
                                .map(|row| {
                                    ResultData::List(
                                        row.into_iter().map(ResultData::Float).collect(),
                                    )
                                })
                                .collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "ROMAN" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "ROMAN")?;
                    let form = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::math_trig::roman(num, form) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "SEC" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "SEC")?;
                    res_to_rd(crate::core::math_trig::sec(x))
                }
                "SECH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "SECH")?;
                    res_to_rd(crate::core::math_trig::sech(x))
                }
                "SEQUENCE" => {
                    let rows = self.to_f64_arg(evaluated_args.first(), "SEQUENCE")?;
                    let cols = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    let start = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    let step = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    match crate::core::math_trig::sequence(rows, cols, start, step) {
                        Ok(grid) => Ok(ResultData::List(
                            grid.into_iter()
                                .map(|row| {
                                    ResultData::List(
                                        row.into_iter().map(ResultData::Float).collect(),
                                    )
                                })
                                .collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "SERIESSUM" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "SERIESSUM")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "SERIESSUM")?;
                    let m = self.to_f64_arg(evaluated_args.get(2), "SERIESSUM")?;
                    let coeffs = match self.flatten_skipping_blanks(evaluated_args.get(3)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                    res_to_rd(crate::core::math_trig::seriessum(x, n, m, &coeffs))
                }
                "SIGN" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "SIGN")?;
                    res_to_rd(crate::core::math_trig::sign(x))
                }
                "SINH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "SINH")?;
                    res_to_rd(crate::core::math_trig::sinh(x))
                }
                "SQRTPI" => {
                    if Self::first_arg_is_boolean(&evaluated_args) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let x = self.to_f64_arg(evaluated_args.first(), "SQRTPI")?;
                    res_to_rd(crate::core::math_trig::sqrtpi(x))
                }
                "SUMPRODUCT" => {
                    // SUMPRODUCT treats non-numeric entries as zeros rather
                    // than skipping or rejecting them, which matters twice
                    // over: the term contributes 0, and -- because the
                    // entry still occupies its slot -- the arrays stay the
                    // same length so the remaining terms keep lining up.
                    // Dropping them instead made SUMPRODUCT(2, "abc")
                    // #VALUE! (length 1 against length 0) where real Excel
                    // answers 0, and SUMPRODUCT({1,2}, {3,"x"}) is 3.
                    let mut arrays: Vec<Vec<f64>> = Vec::new();
                    let mut first_err = None;
                    for arg in &evaluated_args {
                        // A single blank cell is a missing operand, not an
                        // empty array: SUMPRODUCT over one blank cell is
                        // #VALUE! where over two it is 0.
                        if Self::is_empty_scalar_operand(arg) {
                            return Ok(ResultData::Error("#VALUE!".to_string()));
                        }
                        let mut slots = Vec::new();
                        self.flatten_positional(arg, &mut slots, &mut first_err);
                        arrays.push(slots.into_iter().map(|v| v.unwrap_or(0.0)).collect());
                    }
                    if let Some(e) = first_err {
                        return Ok(ResultData::Error(e));
                    }
                    res_to_rd(crate::core::math_trig::sumproduct(&arrays))
                }
                "SUMSQ" => {
                    let nums: Vec<f64> =
                        match self.flatten_args_stat_numbers(&evaluated_args, &arg_is_direct) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::math_trig::sumsq(&nums))
                }
                "SUMX2MY2" => {
                    if self.paired_sum_has_no_numbers(evaluated_args.first())
                        || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                    {
                        return Ok(ResultData::Error("#DIV/0!".to_string()));
                    }
                    let (xs, ys) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::math_trig::sumx2my2(&xs, &ys))
                }
                "SUMX2PY2" => {
                    if self.paired_sum_has_no_numbers(evaluated_args.first())
                        || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                    {
                        return Ok(ResultData::Error("#DIV/0!".to_string()));
                    }
                    let (xs, ys) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::math_trig::sumx2py2(&xs, &ys))
                }
                "SUMXMY2" => {
                    if self.paired_sum_has_no_numbers(evaluated_args.first())
                        || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                    {
                        return Ok(ResultData::Error("#DIV/0!".to_string()));
                    }
                    let (xs, ys) =
                        match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                            Ok(v) => v,
                            Err(e) => return Ok(ResultData::Error(e)),
                        };
                    res_to_rd(crate::core::math_trig::sumxmy2(&xs, &ys))
                }
                "TANH" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "TANH")?;
                    res_to_rd(crate::core::math_trig::tanh(x))
                }
                "TRUNC" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "TRUNC")?;
                    let digits = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::math_trig::trunc(x, digits))
                }

                // --- TEXT FUNCTIONS ---
                "ARRAYTOTEXT" => {
                    // Every element's own text (numbers via
                    // format_excel_number, TRUE/FALSE, raw strings, ...)
                    // via ResultData's Display -- not flatten_stat_numbers,
                    // which silently drops non-numeric cells and so only
                    // ever produced a text/bool-free (and often empty)
                    // result for a mixed range.
                    fn flatten_text(val: &ResultData, out: &mut Vec<String>) {
                        match val {
                            ResultData::List(items) => {
                                for item in items {
                                    flatten_text(item, out);
                                }
                            }
                            other => out.push(other.to_string()),
                        }
                    }
                    let mut items = Vec::new();
                    if let Some(arg) = evaluated_args.first() {
                        flatten_text(arg, &mut items);
                    }
                    // A *single* empty cell has no text to render at all
                    // and is #VALUE!. A multi-cell range of blanks is not:
                    // ARRAYTOTEXT over two empty cells is "," in real
                    // Excel, i.e. the separators still show.
                    if items.len() == 1 && items[0].is_empty() {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let fmt = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::text::arraytotext(&items, fmt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "ASC" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::asc(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "JIS" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::jis(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "BAHTTEXT" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "BAHTTEXT")?;
                    match crate::core::text::bahttext(num) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "CHAR" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "CHAR")?;
                    match crate::core::text::char_fn(num) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "CLEAN" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::clean(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "CODE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::text::code(&text))
                }
                "DBCS" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::dbcs(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "DETECTLANGUAGE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::detectlanguage(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "DOLLAR" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "DOLLAR")?;
                    let dec = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::text::dollar(num, dec) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "EXACT" => {
                    let t1 = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let t2 = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::exact(&t1, &t2) {
                        Ok(b) => Ok(ResultData::Boolean(b)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "FIND" | "FINDB" => {
                    let find_text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let within_text = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let start_num = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::text::find(&find_text, &within_text, start_num))
                }
                "FIXED" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "FIXED")?;
                    let dec = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    let no_commas = evaluated_args.get(2).map(|v| self.to_bool(v));
                    match crate::core::text::fixed(num, dec, no_commas) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "NUMBERVALUE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let dec = evaluated_args.get(1).map(|v| v.to_string());
                    let grp = evaluated_args.get(2).map(|v| v.to_string());
                    res_to_rd(crate::core::text::numbervalue(
                        &text,
                        dec.as_deref(),
                        grp.as_deref(),
                    ))
                }
                "PHONETIC" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::phonetic(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REGEXEXTRACT" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let pat = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::regexextract(&text, &pat) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REGEXREPLACE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let pat = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let rep = evaluated_args
                        .get(2)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::regexreplace(&text, &pat, &rep) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REGEXTEST" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let pat = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::regextest(&text, &pat) {
                        Ok(b) => Ok(ResultData::Boolean(b)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REPLACE" | "REPLACEB" => {
                    let old_text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let start_num = self.to_f64_arg(evaluated_args.get(1), "REPLACE")?;
                    let num_chars = self.to_f64_arg(evaluated_args.get(2), "REPLACE")?;
                    let new_text = evaluated_args
                        .get(3)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::replace_fn(&old_text, start_num, num_chars, &new_text)
                    {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REPT" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let cnt = self.to_f64_arg(evaluated_args.get(1), "REPT")?;
                    match crate::core::text::rept(&text, cnt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "SEARCH" | "SEARCHB" => {
                    let find_text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let within_text = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let start_num = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::text::search(
                        &find_text,
                        &within_text,
                        start_num,
                    ))
                }
                "SUBSTITUTE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let old_text = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let new_text = evaluated_args
                        .get(2)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let instance = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    match crate::core::text::substitute(&text, &old_text, &new_text, instance) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "T" => {
                    let is_str = matches!(evaluated_args.first(), Some(ResultData::String(_)));
                    let val = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    Ok(ResultData::String(crate::core::text::t_fn(&val, is_str)))
                }
                "TEXT" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "TEXT")?;
                    let fmt = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::text_fn(num, &fmt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTAFTER" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let delim = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let instance = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    match crate::core::text::textafter(&text, &delim, instance) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTBEFORE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let delim = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let instance = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    match crate::core::text::textbefore(&text, &delim, instance) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTJOIN" => {
                    let delim = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let ignore = evaluated_args
                        .get(1)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    let texts: Vec<String> = evaluated_args
                        .iter()
                        .skip(2)
                        .map(|v| v.to_string())
                        .collect();
                    match crate::core::text::textjoin(&delim, ignore, &texts) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTSPLIT" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let delim = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::textsplit(&text, &delim) {
                        Ok(parts) => Ok(ResultData::List(
                            parts.into_iter().map(ResultData::String).collect(),
                        )),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TRANSLATE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let from = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let to = evaluated_args
                        .get(2)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::text::translate(&text, &from, &to) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "UNICHAR" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "UNICHAR")?;
                    match crate::core::text::unichar(num) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "UNICODE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::text::unicode(&text))
                }
                "VALUE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::text::value(&text))
                }
                "VALUETOTEXT" => {
                    let val = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let fmt = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::text::valuetotext(&val, fmt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "LEFTB" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let count = evaluated_args
                        .get(1)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(1.0)
                        .floor() as usize;
                    let res: String = text.chars().take(count).collect();
                    Ok(ResultData::String(res))
                }
                "RIGHTB" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let count = evaluated_args
                        .get(1)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(1.0)
                        .floor() as usize;
                    let chars: Vec<char> = text.chars().collect();
                    let skip = chars.len().saturating_sub(count);
                    let res: String = chars.into_iter().skip(skip).collect();
                    Ok(ResultData::String(res))
                }
                "LENB" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    Ok(ResultData::Float(text.len() as f64))
                }
                "MIDB" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let start = self.to_f64_arg(evaluated_args.get(1), "MIDB")?.floor() as usize;
                    let count = self.to_f64_arg(evaluated_args.get(2), "MIDB")?.floor() as usize;
                    if start < 1 {
                        Ok(ResultData::Error("#VALUE!".to_string()))
                    } else {
                        let chars: Vec<char> = text.chars().collect();
                        let start_idx = (start - 1).min(chars.len());
                        let res: String = chars.into_iter().skip(start_idx).take(count).collect();
                        Ok(ResultData::String(res))
                    }
                }

                // --- DATE AND TIME FUNCTIONS ---
                "DATE" => {
                    let y = self.to_f64_arg(evaluated_args.first(), "DATE")?;
                    let m = self.to_f64_arg(evaluated_args.get(1), "DATE")?;
                    let d = self.to_f64_arg(evaluated_args.get(2), "DATE")?;
                    res_to_rd(crate::core::date_fn::date_fn(y, m, d))
                }
                "DATEDIF" => {
                    let start = self.to_f64_arg(evaluated_args.first(), "DATEDIF")?;
                    let end = self.to_f64_arg(evaluated_args.get(1), "DATEDIF")?;
                    let unit = evaluated_args
                        .get(2)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::date_fn::datedif(start, end, &unit))
                }
                "DATEVALUE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::date_fn::datevalue(&text))
                }
                "DAY" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "DAY")?;
                    res_to_rd(crate::core::date_fn::day_fn(s))
                }
                "DAYS" => {
                    let e = self.to_f64_arg(evaluated_args.first(), "DAYS")?;
                    let s = self.to_f64_arg(evaluated_args.get(1), "DAYS")?;
                    res_to_rd(crate::core::date_fn::days(e, s))
                }
                "DAYS360" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "DAYS360")?;
                    let e = self.to_f64_arg(evaluated_args.get(1), "DAYS360")?;
                    let method = evaluated_args.get(2).map(|v| self.to_bool(v));
                    res_to_rd(crate::core::date_fn::days360(s, e, method))
                }
                "EDATE" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "EDATE")?;
                    let m = self.to_f64_arg(evaluated_args.get(1), "EDATE")?;
                    res_to_rd(crate::core::date_fn::edate(s, m))
                }
                "EOMONTH" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "EOMONTH")?;
                    let m = self.to_f64_arg(evaluated_args.get(1), "EOMONTH")?;
                    res_to_rd(crate::core::date_fn::eomonth(s, m))
                }
                "HOUR" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "HOUR")?;
                    res_to_rd(crate::core::date_fn::hour_fn(s))
                }
                "ISOWEEKNUM" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "ISOWEEKNUM")?;
                    res_to_rd(crate::core::date_fn::isoweeknum(s))
                }
                "MINUTE" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "MINUTE")?;
                    res_to_rd(crate::core::date_fn::minute_fn(s))
                }
                "MONTH" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "MONTH")?;
                    res_to_rd(crate::core::date_fn::month_fn(s))
                }
                "NETWORKDAYS" | "NETWORKDAYS.INTL" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "NETWORKDAYS")?;
                    let e = self.to_f64_arg(evaluated_args.get(1), "NETWORKDAYS")?;
                    let holidays: Vec<f64> = evaluated_args
                        .get(2)
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    res_to_rd(crate::core::date_fn::networkdays(s, e, &holidays))
                }
                "SECOND" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "SECOND")?;
                    res_to_rd(crate::core::date_fn::second_fn(s))
                }
                "TIME" => {
                    let h = self.to_f64_arg(evaluated_args.first(), "TIME")?;
                    let m = self.to_f64_arg(evaluated_args.get(1), "TIME")?;
                    let s = self.to_f64_arg(evaluated_args.get(2), "TIME")?;
                    res_to_rd(crate::core::date_fn::time_fn(h, m, s))
                }
                "TIMEVALUE" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::date_fn::timevalue(&text))
                }
                "WEEKDAY" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "WEEKDAY")?;
                    let r_type = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::date_fn::weekday(s, r_type))
                }
                "WEEKNUM" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "WEEKNUM")?;
                    let r_type = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::date_fn::weeknum(s, r_type))
                }
                "WORKDAY" | "WORKDAY.INTL" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "WORKDAY")?;
                    let days = self.to_f64_arg(evaluated_args.get(1), "WORKDAY")?;
                    let holidays: Vec<f64> = evaluated_args
                        .get(2)
                        .map(|arg| self.flatten_stat_numbers(arg, false))
                        .unwrap_or_default();
                    res_to_rd(crate::core::date_fn::workday(s, days, &holidays))
                }
                "YEAR" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "YEAR")?;
                    res_to_rd(crate::core::date_fn::year_fn(s))
                }
                "YEARFRAC" => {
                    let s = self.to_f64_arg(evaluated_args.first(), "YEARFRAC")?;
                    let e = self.to_f64_arg(evaluated_args.get(1), "YEARFRAC")?;
                    let basis = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::date_fn::yearfrac(s, e, basis))
                }

                // --- ENGINEERING FUNCTIONS ---
                "BESSELI" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "BESSELI")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "BESSELI")?;
                    res_to_rd(crate::core::engineering::besseli(x, n))
                }
                "BESSELJ" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "BESSELJ")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "BESSELJ")?;
                    res_to_rd(crate::core::engineering::besselj(x, n))
                }
                "BESSELK" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "BESSELK")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "BESSELK")?;
                    res_to_rd(crate::core::engineering::besselk(x, n))
                }
                "BESSELY" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "BESSELY")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "BESSELY")?;
                    res_to_rd(crate::core::engineering::bessely(x, n))
                }
                "BIN2DEC" => {
                    if Self::first_arg_is_boolean(&evaluated_args) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::bin2dec(&t))
                }
                "BIN2HEX" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::bin2hex(&t, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "BIN2OCT" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::bin2oct(&t, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "BITAND" => {
                    let n1 = self.to_f64_arg(evaluated_args.first(), "BITAND")?;
                    let n2 = self.to_f64_arg(evaluated_args.get(1), "BITAND")?;
                    res_to_rd(crate::core::engineering::bitand(n1, n2))
                }
                "BITLSHIFT" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "BITLSHIFT")?;
                    let s = self.to_f64_arg(evaluated_args.get(1), "BITLSHIFT")?;
                    res_to_rd(crate::core::engineering::bitlshift(n, s))
                }
                "BITOR" => {
                    let n1 = self.to_f64_arg(evaluated_args.first(), "BITOR")?;
                    let n2 = self.to_f64_arg(evaluated_args.get(1), "BITOR")?;
                    res_to_rd(crate::core::engineering::bitor(n1, n2))
                }
                "BITRSHIFT" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "BITRSHIFT")?;
                    let s = self.to_f64_arg(evaluated_args.get(1), "BITRSHIFT")?;
                    res_to_rd(crate::core::engineering::bitrshift(n, s))
                }
                "BITXOR" => {
                    let n1 = self.to_f64_arg(evaluated_args.first(), "BITXOR")?;
                    let n2 = self.to_f64_arg(evaluated_args.get(1), "BITXOR")?;
                    res_to_rd(crate::core::engineering::bitxor(n1, n2))
                }
                "COMPLEX" => {
                    let r = self.to_f64_arg(evaluated_args.first(), "COMPLEX")?;
                    let i = self.to_f64_arg(evaluated_args.get(1), "COMPLEX")?;
                    let s = evaluated_args.get(2).map(|v| v.to_string());
                    match crate::core::engineering::complex_fn(r, i, s.as_deref()) {
                        Ok(res) => Ok(ResultData::String(res)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "CONVERT" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "CONVERT")?;
                    let u1 = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let u2 = evaluated_args
                        .get(2)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::convert(val, &u1, &u2))
                }
                "DEC2BIN" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "DEC2BIN")?;
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::dec2bin(n, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "DEC2HEX" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "DEC2HEX")?;
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::dec2hex(n, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "DEC2OCT" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "DEC2OCT")?;
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::dec2oct(n, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "DELTA" => {
                    let n1 = self.to_f64_arg(evaluated_args.first(), "DELTA")?;
                    let n2 = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::engineering::delta(n1, n2))
                }
                "ERF" | "ERFC" | "ERF.PRECISE" | "ERFC.PRECISE" => {
                    // Unlike SQRT/ABS/INT/MOD (which all accept a boolean
                    // as 1/0), the error functions reject booleans: real
                    // Excel answers #VALUE! for ERF(TRUE) and ERF(FALSE).
                    // Numeric *text* is coerced though, from a literal or
                    // from a text cell, and surrounding whitespace is
                    // tolerated -- ERF("1") and ERF(" 1 ") both give
                    // 0.8427007929497149. Non-numeric text is #VALUE!.
                    // A blank argument coerces to 0 (ERF(<blank>) is 0 and
                    // ERFC(<blank>) is 1). Same rule as QUOTIENT.
                    let x = match evaluated_args.first() {
                        None | Some(ResultData::None) => 0.0,
                        Some(v) => {
                            // A one-cell range arrives as a one-element List.
                            let scalar = match v {
                                ResultData::List(items) if items.len() == 1 => &items[0],
                                other => other,
                            };
                            if matches!(scalar, ResultData::Boolean(_)) {
                                // See first_arg_is_boolean.
                                return Ok(ResultData::Error("#VALUE!".to_string()));
                            }
                            match self.to_f64(scalar) {
                                Some(f) => f,
                                None => return Ok(ResultData::Error("#VALUE!".to_string())),
                            }
                        }
                    };
                    let v = if upper_name.starts_with("ERFC") {
                        crate::core::stats::erfc(x)
                    } else {
                        crate::core::stats::erf(x)
                    };
                    res_to_rd(Ok(v))
                }
                "GESTEP" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "GESTEP")?;
                    let step = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::engineering::gestep(n, step))
                }
                "HEX2BIN" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::hex2bin(&t, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "HEX2DEC" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::hex2dec(&t))
                }
                "HEX2OCT" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::hex2oct(&t, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMABS" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::imabs(&t))
                }
                "IMAGINARY" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::imaginary(&t))
                }
                "IMARGUMENT" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::imargument(&t))
                }
                "IMCONJUGATE" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::engineering::imconjugate(&t) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMDIV" => {
                    let t1 = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let t2 = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::engineering::imdiv(&t1, &t2) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMPRODUCT" => {
                    let strs: Vec<String> = evaluated_args.iter().map(|v| v.to_string()).collect();
                    let refs: Vec<&str> = strs.iter().map(|s| s.as_str()).collect();
                    match crate::core::engineering::improduct(&refs) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMREAL" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::imreal(&t))
                }
                "IMSUB" => {
                    let t1 = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let t2 = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::engineering::imsub(&t1, &t2) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMSUM" => {
                    let strs: Vec<String> = evaluated_args.iter().map(|v| v.to_string()).collect();
                    let refs: Vec<&str> = strs.iter().map(|s| s.as_str()).collect();
                    match crate::core::engineering::imsum(&refs) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "OCT2BIN" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::oct2bin(&t, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "OCT2DEC" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    res_to_rd(crate::core::engineering::oct2dec(&t))
                }
                "OCT2HEX" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::engineering::oct2hex(&t, p) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMCOS" | "IMCOSH" | "IMCOT" | "IMCSC" | "IMCSCH" | "IMEXP" | "IMLN"
                | "IMLOG10" | "IMLOG2" | "IMSEC" | "IMSECH" | "IMSIN" | "IMSINH" | "IMSQRT"
                | "IMTAN" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let result = match upper_name.as_str() {
                        "IMCOS" => crate::core::engineering::imcos(&t),
                        "IMCOSH" => crate::core::engineering::imcosh(&t),
                        "IMCOT" => crate::core::engineering::imcot(&t),
                        "IMCSC" => crate::core::engineering::imcsc(&t),
                        "IMCSCH" => crate::core::engineering::imcsch(&t),
                        "IMEXP" => crate::core::engineering::imexp(&t),
                        "IMLN" => crate::core::engineering::imln(&t),
                        "IMLOG10" => crate::core::engineering::imlog10(&t),
                        "IMLOG2" => crate::core::engineering::imlog2(&t),
                        "IMSEC" => crate::core::engineering::imsec(&t),
                        "IMSECH" => crate::core::engineering::imsech(&t),
                        "IMSIN" => crate::core::engineering::imsin(&t),
                        "IMSINH" => crate::core::engineering::imsinh(&t),
                        "IMSQRT" => crate::core::engineering::imsqrt(&t),
                        "IMTAN" => crate::core::engineering::imtan(&t),
                        _ => unreachable!(),
                    };
                    match result {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "IMPOWER" => {
                    let t = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let n = self.to_f64_arg(evaluated_args.get(1), "IMPOWER")?;
                    match crate::core::engineering::impower(&t, n) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }

                // --- INFORMATION & LOGICAL & DATABASE & LOOKUP & WEB & CUBE FUNCTIONS ---
                "ERROR.TYPE" => {
                    let t = match evaluated_args.first() {
                        Some(ResultData::Error(e)) => e.clone(),
                        _ => String::new(),
                    };
                    res_to_rd(crate::core::extended_fn::error_type(&t))
                }
                "ISERR" => {
                    let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                    Ok(ResultData::Boolean(crate::core::extended_fn::iserr(&val)))
                }
                "ISEVEN" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "ISEVEN")?;
                    Ok(ResultData::Boolean(crate::core::extended_fn::iseven(n)))
                }
                "ISLOGICAL" => {
                    let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                    Ok(ResultData::Boolean(crate::core::extended_fn::islogical(
                        &val,
                    )))
                }
                "ISNONTEXT" => {
                    let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                    Ok(ResultData::Boolean(crate::core::extended_fn::isnontext(
                        &val,
                    )))
                }
                "ISODD" => {
                    let n = self.to_f64_arg(evaluated_args.first(), "ISODD")?;
                    Ok(ResultData::Boolean(crate::core::extended_fn::isodd(n)))
                }
                "N" => {
                    let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                    Ok(ResultData::Float(crate::core::extended_fn::n_fn(&val)))
                }
                "NA" => Ok(crate::core::extended_fn::na_fn()),
                "TYPE" => {
                    let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                    Ok(ResultData::Float(crate::core::extended_fn::type_fn(&val)))
                }
                "XOR" => {
                    let bools: Vec<bool> = evaluated_args.iter().map(|v| self.to_bool(v)).collect();
                    Ok(ResultData::Boolean(crate::core::extended_fn::xor_fn(
                        &bools,
                    )))
                }
                "ADDRESS" => {
                    let r = self.to_f64_arg(evaluated_args.first(), "ADDRESS")?;
                    let c = self.to_f64_arg(evaluated_args.get(1), "ADDRESS")?;
                    let abs_n = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    let a1 = evaluated_args.get(3).map(|v| self.to_bool(v));
                    let s_name = evaluated_args.get(4).map(|v| v.to_string());
                    match crate::core::extended_fn::address_fn(r, c, abs_n, a1, s_name.as_deref()) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "HLOOKUP" => {
                    if evaluated_args.len() < 3 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "HLOOKUP requires at least 3 arguments".to_string(),
                        )));
                    }
                    let lookup_val = &evaluated_args[0];
                    let row_idx = self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as usize;
                    let range_lookup = if evaluated_args.len() >= 4 {
                        self.to_bool(&evaluated_args[3])
                    } else {
                        true
                    };

                    // The mirror image of VLOOKUP just below: the range's
                    // flat, row-major `List` is reshaped using the
                    // *unevaluated* range's column span, the first *row*
                    // (not column) is searched, and the match is read back
                    // out of the target row. Previously this went through
                    // `extract_matrix`, which coerces every cell through
                    // `to_f64` and silently drops non-numeric ones -- so a
                    // text header row (the common HLOOKUP case) never
                    // matched.
                    if let ResultData::List(list) = &evaluated_args[1] {
                        let num_cols = match &args[1] {
                            Expr::RangeRef {
                                start_col, end_col, ..
                            } => end_col - start_col + 1,
                            _ => list.len(),
                        };

                        let num_rows = list.len().checked_div(num_cols).unwrap_or(0);
                        if num_rows == 0 || row_idx == 0 || row_idx > num_rows {
                            return Ok(ResultData::Error("#N/A".to_string()));
                        }

                        let first_row = &list[..num_cols];
                        let mut found_col_idx: Option<usize> = None;
                        if !range_lookup {
                            for (c, item) in first_row.iter().enumerate() {
                                if Self::exact_lookup_matches(lookup_val, item) {
                                    found_col_idx = Some(c);
                                    break;
                                }
                            }
                        } else {
                            let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                            for (c, item) in first_row.iter().enumerate() {
                                let val_f = self.to_f64(item).unwrap_or(0.0);
                                if val_f <= lookup_f {
                                    found_col_idx = Some(c);
                                } else {
                                    break;
                                }
                            }
                        }

                        match found_col_idx {
                            Some(c) => Ok(list[(row_idx - 1) * num_cols + c].clone()),
                            None => Ok(ResultData::Error("#N/A".to_string())),
                        }
                    } else {
                        Ok(ResultData::Error("#N/A".to_string()))
                    }
                }
                "ENCODEURL" => {
                    let text = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::extended_fn::encodeurl(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                // Both need a live external data source this engine has no
                // access to (Microsoft's undocumented stock-data cloud
                // service for STOCKHISTORY; a registered Windows COM
                // IRtdServer for RTD) -- #N/A matches what real Excel shows
                // once that connection is unavailable, rather than the
                // misleading echo-the-last-argument placeholder these used
                // to fall through to.
                "STOCKHISTORY" | "RTD" => Ok(ResultData::Error("#N/A".to_string())),
                "DAVERAGE" | "DCOUNT" | "DCOUNTA" | "DGET" | "DMAX" | "DMIN" | "DPRODUCT"
                | "DSTDEV" | "DSTDEVP" | "DSUM" | "DVAR" | "DVARP" => self
                    .evaluate_database_function(
                        upper_name.as_str(),
                        args,
                        &evaluated_args,
                        context,
                    ),
                "HYPERLINK" => {
                    // No clickable-hyperlink concept in this engine --
                    // returns the display value a formula-based consumer
                    // would see: friendly_name if given, else the raw
                    // link_location text.
                    match evaluated_args.get(1) {
                        Some(friendly) => Ok(friendly.clone()),
                        None => Ok(evaluated_args
                            .first()
                            .cloned()
                            .unwrap_or(ResultData::Error("#VALUE!".to_string()))),
                    }
                }
                // No OLAP cube connection concept exists in this engine --
                // #N/A matches what real Excel shows once a cube function's
                // underlying connection is unavailable, the same reasoning
                // already applied to RTD/STOCKHISTORY above, rather than
                // the misleading echo-the-last-argument placeholder these
                // (and GROUPBY/PIVOTBY/IMAGE/WEBSERVICE below) used to fall
                // through to -- a plausible-looking wrong value is worse
                // than a visible error, since it can silently corrupt a
                // downstream calculation with no signal anything is wrong.
                "CUBEKPIMEMBER" | "CUBEMEMBER" | "CUBEMEMBERPROPERTY" | "CUBERANKEDMEMBER"
                | "CUBESET" | "CUBESETCOUNT" | "CUBEVALUE" => {
                    Ok(ResultData::Error("#N/A".to_string()))
                }
                // WEBSERVICE needs actual network access to an arbitrary
                // URL; #VALUE! matches Microsoft's own documented error
                // for a request that can't be completed.
                "WEBSERVICE" => Ok(ResultData::Error("#VALUE!".to_string())),
                // IMAGE needs to fetch/decode real image data, which this
                // engine has no concept of; #VALUE! matches real Excel's
                // error for a source it can't resolve to a usable image.
                "IMAGE" => Ok(ResultData::Error("#VALUE!".to_string())),
                // GROUPBY/PIVOTBY are genuine, deterministic array
                // functions (not connection-dependent like the above) --
                // unlike those, faking an error here would be its own
                // regression from the previous echo-last-arg placeholder,
                // which at least degraded gracefully for the single-
                // aggregate-function common case. Properly implementing
                // Excel's full row/column-field grouping and dynamic-array
                // spill semantics is real, separately-scoped work (this
                // engine already has the pivot-table grouping machinery in
                // pivot.rs that a real implementation would build on) --
                // left as a stub for now, but returning #N/A like the
                // genuinely-unimplementable functions above would be
                // actively misleading about *why* it's unimplemented.
                "GROUPBY" | "PIVOTBY" => {
                    Ok(evaluated_args.last().cloned().unwrap_or(ResultData::None))
                }
                "FILTERXML" => {
                    let xml = evaluated_args
                        .first()
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let xpath = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    match crate::core::xml::filterxml(&xml, &xpath) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }

                "SUM" => {
                    if let Some(err) = self.check_arg_errors(&evaluated_args, &arg_is_direct) {
                        return Ok(err);
                    }
                    let mut sum = 0.0;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        sum += self.sum_helper(arg, arg_is_direct[i]);
                    }
                    Ok(ResultData::Float(sum))
                }
                "AVERAGE" => {
                    if let Some(err) = self.check_arg_errors(&evaluated_args, &arg_is_direct) {
                        return Ok(err);
                    }
                    let mut sum = 0.0;
                    let mut count = 0;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        let (s, c) = self.average_helper(arg, arg_is_direct[i]);
                        sum += s;
                        count += c;
                    }
                    if count == 0 {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else {
                        Ok(ResultData::Float(sum / count as f64))
                    }
                }
                "COUNT" => {
                    // A boolean counts when it is typed directly as an
                    // argument, but not when it merely sits inside a
                    // referenced range -- Excel's documented split, and the
                    // same is_direct distinction the SUM/AVERAGE helpers
                    // already make.
                    let mut count = 0;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        let direct = arg_is_direct.get(i).copied().unwrap_or(false);
                        if direct && matches!(arg, ResultData::Boolean(_)) {
                            count += 1;
                        } else if direct
                            && matches!(arg, ResultData::String(_))
                            && self.to_f64(arg).is_some()
                        {
                            // Numeric text typed directly counts too --
                            // COUNT("12", 3, 4, 5) is 4. Text that will not
                            // coerce is simply not counted; unlike the rest
                            // of the family COUNT never reports #VALUE!.
                            count += 1;
                        } else {
                            count += self.count_helper(arg);
                        }
                    }
                    Ok(ResultData::Float(count as f64))
                }
                "MIN" => {
                    if let Some(err) = self.check_arg_errors(&evaluated_args, &arg_is_direct) {
                        return Ok(err);
                    }
                    let mut min_val = f64::INFINITY;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        min_val = min_val.min(self.min_helper(arg, arg_is_direct[i]));
                    }
                    if min_val.is_infinite() {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(min_val))
                    }
                }
                "MAX" => {
                    if let Some(err) = self.check_arg_errors(&evaluated_args, &arg_is_direct) {
                        return Ok(err);
                    }
                    let mut max_val = f64::NEG_INFINITY;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        max_val = max_val.max(self.max_helper(arg, arg_is_direct[i]));
                    }
                    if max_val.is_infinite() {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(max_val))
                    }
                }

                "STR" => {
                    if evaluated_args.is_empty() {
                        Ok(ResultData::String(String::new()))
                    } else {
                        Ok(ResultData::String(evaluated_args[0].to_string()))
                    }
                }
                "SQRT" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "SQRT")?;
                    if val < 0.0 {
                        Ok(ResultData::Error("#NUM!".to_string()))
                    } else {
                        Ok(ResultData::Float(val.sqrt()))
                    }
                }
                "RAND" => {
                    use rand::Rng;
                    let mut rng = rand::thread_rng();
                    Ok(ResultData::Float(rng.r#gen::<f64>()))
                }
                "RANDBETWEEN" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "RANDBETWEEN requires 2 arguments".to_string(),
                        )));
                    }
                    let bottom = self
                        .to_f64_arg(evaluated_args.first(), "RANDBETWEEN")?
                        .round() as i64;
                    let top = self
                        .to_f64_arg(evaluated_args.get(1), "RANDBETWEEN")?
                        .round() as i64;
                    use rand::Rng;
                    let mut rng = rand::thread_rng();
                    let val = if bottom <= top {
                        rng.gen_range(bottom..=top)
                    } else {
                        rng.gen_range(top..=bottom)
                    };
                    Ok(ResultData::Integer(val))
                }
                "SIN" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "SIN")?;
                    Ok(ResultData::Float(val.sin()))
                }
                "COS" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "COS")?;
                    Ok(ResultData::Float(val.cos()))
                }
                "TAN" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "TAN")?;
                    Ok(ResultData::Float(val.tan()))
                }
                "ACOS" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ACOS")?;
                    Ok(ResultData::Float(val.acos()))
                }
                "ASIN" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ASIN")?;
                    Ok(ResultData::Float(val.asin()))
                }
                "ATAN" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ATAN")?;
                    Ok(ResultData::Float(val.atan()))
                }
                "FLOOR" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "FLOOR")?;
                    let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::math_trig::floor_math(val, sig, None))
                }
                "CEILING" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "CEILING")?;
                    let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::math_trig::ceiling_math(val, sig, None))
                }
                "LOG10" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "LOG10")?;
                    Ok(ResultData::Float(val.log10()))
                }
                "LN" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "LN")?;
                    Ok(ResultData::Float(val.ln()))
                }
                "EXP" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "EXP")?;
                    Ok(ResultData::Float(val.exp()))
                }
                "GET" => {
                    if evaluated_args.len() == 2 {
                        let row = self.to_f64(&evaluated_args[0]).unwrap_or(0.0) as usize;
                        let col = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as usize;
                        let cell_ref = CellRef::new(row, col);
                        deps.push(Dependency::Local(cell_ref));
                        Ok(self.get_result_data(&cell_ref))
                    } else if evaluated_args.len() == 3 {
                        let sheet_name = evaluated_args[0].to_string();
                        let row = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as usize;
                        let col = self.to_f64(&evaluated_args[2]).unwrap_or(0.0) as usize;
                        let cell_ref = CellRef::new(row, col);

                        if sheet_name == self.name {
                            deps.push(Dependency::Local(cell_ref));
                            Ok(self.get_result_data(&cell_ref))
                        } else {
                            deps.push(Dependency::Remote {
                                sheet: sheet_name.clone(),
                                cell: cell_ref,
                            });
                            if let Some(ctx) = context {
                                if let Some(t) = ctx.sheets.get(&sheet_name) {
                                    Ok(t.get_result_data(&cell_ref))
                                } else {
                                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                                        format!("Sheet not found: {}", sheet_name),
                                    )))
                                }
                            } else {
                                Err(EngineError::EvalError(EvalError::UnknownFunction(
                                    "No context".to_string(),
                                )))
                            }
                        }
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "get() takes 2 or 3 arguments".to_string(),
                        )))
                    }
                }
                "GET_COL" => {
                    if evaluated_args.len() == 2 {
                        let sheet_name = evaluated_args[0].to_string();
                        let col_name = evaluated_args[1].to_string();
                        let is_self = sheet_name == self.name;

                        if let Some(ctx) = context {
                            if let Some(sheet) = ctx.sheets.get(&sheet_name) {
                                if let Some(col_idx) =
                                    sheet.columns.iter().position(|c| c.name == col_name)
                                {
                                    if is_self {
                                        deps.push(Dependency::LocalColumn(col_idx));
                                    } else {
                                        deps.push(Dependency::RemoteColumn {
                                            sheet: sheet_name.clone(),
                                            col: col_idx,
                                        });
                                    }
                                    let mut results = Vec::new();
                                    for row in 0..sheet.row_count() {
                                        let cell_ref = CellRef::new(row, col_idx);
                                        results.push(sheet.get_result_data(&cell_ref));
                                    }
                                    Ok(ResultData::List(results))
                                } else {
                                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                                        format!("Column not found: {}", col_name),
                                    )))
                                }
                            } else {
                                Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                    "Sheet not found: {}",
                                    sheet_name
                                ))))
                            }
                        } else if is_self {
                            if let Some(col_idx) =
                                self.columns.iter().position(|c| c.name == col_name)
                            {
                                deps.push(Dependency::LocalColumn(col_idx));
                                let mut results = Vec::new();
                                for row in 0..self.row_count() {
                                    let cell_ref = CellRef::new(row, col_idx);
                                    results.push(self.get_result_data(&cell_ref));
                                }
                                Ok(ResultData::List(results))
                            } else {
                                Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                    "Column not found: {}",
                                    col_name
                                ))))
                            }
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(
                                "No context to resolve sheet reference".to_string(),
                            )))
                        }
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "get_col() takes 2 arguments".to_string(),
                        )))
                    }
                }
                "ABS" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ABS")?;
                    Ok(ResultData::Float(val.abs()))
                }
                "INT" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "INT")?;
                    Ok(ResultData::Float(val.floor()))
                }
                "ROUND" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ROUND")?;
                    let digits = if evaluated_args.len() >= 2 {
                        self.to_f64_arg(evaluated_args.get(1), "ROUND")? as i32
                    } else {
                        0
                    };
                    let factor = 10.0f64.powi(digits);
                    let mut scaled = val * factor;
                    if scaled.abs() >= 1e-12 {
                        scaled = (scaled * 1e12).round() / 1e12;
                    }
                    Ok(ResultData::Float(scaled.round() / factor))
                }
                "ROUNDUP" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ROUNDUP")?;
                    let digits = if evaluated_args.len() >= 2 {
                        self.to_f64_arg(evaluated_args.get(1), "ROUNDUP")? as i32
                    } else {
                        0
                    };
                    let factor = 10.0f64.powi(digits);
                    let mut scaled = val * factor;
                    if scaled.abs() >= 1e-12 {
                        scaled = (scaled * 1e12).round() / 1e12;
                    }
                    let rounded = if val >= 0.0 {
                        scaled.ceil() / factor
                    } else {
                        scaled.floor() / factor
                    };
                    Ok(ResultData::Float(rounded))
                }
                "ROUNDDOWN" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "ROUNDDOWN")?;
                    let digits = if evaluated_args.len() >= 2 {
                        self.to_f64_arg(evaluated_args.get(1), "ROUNDDOWN")? as i32
                    } else {
                        0
                    };
                    let factor = 10.0f64.powi(digits);
                    let mut scaled = val * factor;
                    if scaled.abs() >= 1e-12 {
                        scaled = (scaled * 1e12).round() / 1e12;
                    }
                    let rounded = if val >= 0.0 {
                        scaled.floor() / factor
                    } else {
                        scaled.ceil() / factor
                    };
                    Ok(ResultData::Float(rounded))
                }
                "SLICE" => {
                    if evaluated_args.len() == 3 {
                        if let ResultData::List(list) = &evaluated_args[0] {
                            let start = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as isize;
                            let mut end = self.to_f64(&evaluated_args[2]).unwrap_or(-1.0) as isize;

                            let len = list.len() as isize;
                            let start_idx = if start < 0 {
                                (len + start).max(0)
                            } else {
                                start.min(len)
                            } as usize;

                            if end == -1 {
                                end = len;
                            }
                            let end_idx = if end < 0 {
                                (len + end).max(0)
                            } else {
                                end.min(len)
                            } as usize;

                            let sliced = if start_idx < end_idx && start_idx < list.len() {
                                list[start_idx..end_idx.min(list.len())].to_vec()
                            } else {
                                Vec::new()
                            };
                            Ok(ResultData::List(sliced))
                        } else {
                            Ok(ResultData::None)
                        }
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "SLICE requires 3 arguments".to_string(),
                        )))
                    }
                }
                "INDEX" => {
                    if evaluated_args.len() == 2 {
                        if let ResultData::List(raw) = &evaluated_args[0] {
                            let (list, _) = Self::flatten_row_major(raw.clone());
                            let idx = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as isize;
                            let len = list.len() as isize;
                            // 1-based like every other INDEX form -- found
                            // via the differential fuzzer (LAMBDA/MAP/
                            // BYROW testing was the first thing to ever
                            // exercise this 2-arg path; the standalone
                            // INDEX fuzz generator only ever used the
                            // 3-arg row/col form) that this returned the
                            // element one past the requested position.
                            let real_idx = if idx < 0 { len + idx } else { idx - 1 };
                            if real_idx >= 0 && real_idx < len {
                                Ok(list[real_idx as usize].clone())
                            } else {
                                Ok(ResultData::None)
                            }
                        } else {
                            Ok(ResultData::None)
                        }
                    } else if evaluated_args.len() == 3 {
                        let row_num = self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as isize;
                        let col_num = self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as isize;

                        if let ResultData::List(raw) = &evaluated_args[0] {
                            let (list, nested_cols) = Self::flatten_row_major(raw.clone());
                            let num_cols = match nested_cols {
                                Some(c) => c as isize,
                                None => match &args[0] {
                                    Expr::RangeRef {
                                        start_col, end_col, ..
                                    } => (end_col - start_col + 1) as isize,
                                    Expr::FunctionCall { name, args: fargs } => self
                                        .function_call_cols(
                                            name, fargs, context, row, col, deps, scope,
                                        )
                                        .unwrap_or(1)
                                        as isize,
                                    _ => 1,
                                },
                            };
                            let r_idx = row_num - 1;
                            let c_idx = col_num - 1;
                            let flat_idx = r_idx * num_cols + c_idx;
                            if flat_idx >= 0 && flat_idx < list.len() as isize {
                                Ok(list[flat_idx as usize].clone())
                            } else {
                                Ok(ResultData::None)
                            }
                        } else {
                            Ok(ResultData::None)
                        }
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "INDEX requires 2 or 3 arguments".to_string(),
                        )))
                    }
                }
                "GET_COL_IDX" => {
                    if evaluated_args.len() == 2 {
                        let sheet_name = evaluated_args[0].to_string();
                        let col_idx = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as usize;
                        let is_self = sheet_name == self.name;

                        if let Some(ctx) = context {
                            if let Some(sheet) = ctx.sheets.get(&sheet_name) {
                                if is_self {
                                    deps.push(Dependency::LocalColumn(col_idx));
                                } else {
                                    deps.push(Dependency::RemoteColumn {
                                        sheet: sheet_name.clone(),
                                        col: col_idx,
                                    });
                                }
                                let mut results = Vec::new();
                                for row in 0..sheet.row_count() {
                                    let cell_ref = CellRef::new(row, col_idx);
                                    results.push(sheet.get_result_data(&cell_ref));
                                }
                                Ok(ResultData::List(results))
                            } else {
                                Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                    "Sheet not found: {}",
                                    sheet_name
                                ))))
                            }
                        } else if is_self {
                            deps.push(Dependency::LocalColumn(col_idx));
                            let mut results = Vec::new();
                            for row in 0..self.row_count() {
                                let cell_ref = CellRef::new(row, col_idx);
                                results.push(self.get_result_data(&cell_ref));
                            }
                            Ok(ResultData::List(results))
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(
                                "No context to resolve sheet reference".to_string(),
                            )))
                        }
                    } else if evaluated_args.len() == 1 {
                        let col_idx = self.to_f64(&evaluated_args[0]).unwrap_or(0.0) as usize;
                        deps.push(Dependency::LocalColumn(col_idx));
                        let mut results = Vec::new();
                        for row in 0..self.row_count() {
                            let cell_ref = CellRef::new(row, col_idx);
                            results.push(self.get_result_data(&cell_ref));
                        }
                        Ok(ResultData::List(results))
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "GET_COL_IDX requires 1 or 2 arguments".to_string(),
                        )))
                    }
                }
                "COUNTA" => {
                    let mut count = 0;
                    for arg in evaluated_args {
                        count += self.counta_helper(&arg);
                    }
                    Ok(ResultData::Float(count as f64))
                }
                "CONCAT" | "CONCATENATE" => {
                    let mut out = String::new();
                    for arg in evaluated_args {
                        self.concat_helper(&arg, &mut out);
                    }
                    Ok(ResultData::String(out))
                }

                "AND" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(true));
                    }
                    let mut res = true;
                    let mut first_err = None;
                    for arg in &evaluated_args {
                        match arg {
                            ResultData::Error(e) => {
                                if first_err.is_none() {
                                    first_err = Some(ResultData::Error(e.clone()));
                                }
                            }
                            ResultData::List(list) => {
                                for item in list {
                                    if let ResultData::Error(e) = item {
                                        if first_err.is_none() {
                                            first_err = Some(ResultData::Error(e.clone()));
                                        }
                                    } else if !self.to_bool(item) {
                                        res = false;
                                        if first_err.is_none() {
                                            return Ok(ResultData::Boolean(false));
                                        }
                                        break;
                                    }
                                }
                            }
                            other => {
                                if !self.to_bool(other) {
                                    res = false;
                                    if first_err.is_none() {
                                        return Ok(ResultData::Boolean(false));
                                    }
                                }
                            }
                        }
                    }
                    if let Some(err) = first_err {
                        Ok(err)
                    } else {
                        Ok(ResultData::Boolean(res))
                    }
                }
                "OR" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(false));
                    }
                    let mut res = false;
                    let mut first_err = None;
                    for arg in &evaluated_args {
                        match arg {
                            ResultData::Error(e) => {
                                if first_err.is_none() {
                                    first_err = Some(ResultData::Error(e.clone()));
                                }
                            }
                            ResultData::List(list) => {
                                for item in list {
                                    if let ResultData::Error(e) = item {
                                        if first_err.is_none() {
                                            first_err = Some(ResultData::Error(e.clone()));
                                        }
                                    } else if self.to_bool(item) {
                                        res = true;
                                        if first_err.is_none() {
                                            return Ok(ResultData::Boolean(true));
                                        }
                                        break;
                                    }
                                }
                            }
                            other => {
                                if self.to_bool(other) {
                                    res = true;
                                    if first_err.is_none() {
                                        return Ok(ResultData::Boolean(true));
                                    }
                                }
                            }
                        }
                    }
                    if let Some(err) = first_err {
                        Ok(err)
                    } else {
                        Ok(ResultData::Boolean(res))
                    }
                }
                "TRUE" => Ok(ResultData::Boolean(true)),
                "FALSE" => Ok(ResultData::Boolean(false)),
                "NOT" => {
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                        return Ok(err);
                    }
                    let val = evaluated_args.first().ok_or_else(|| {
                        EngineError::EvalError(EvalError::UnknownFunction(
                            "NOT requires 1 argument".to_string(),
                        ))
                    })?;
                    Ok(ResultData::Boolean(!self.to_bool(val)))
                }
                "LEFT" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::String(String::new()));
                    }
                    let s = evaluated_args[0].to_string();
                    let num_chars = if evaluated_args.len() >= 2 {
                        self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as usize
                    } else {
                        1
                    };
                    let prefix: String = s.chars().take(num_chars).collect();
                    Ok(ResultData::String(prefix))
                }
                "RIGHT" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::String(String::new()));
                    }
                    let s = evaluated_args[0].to_string();
                    let num_chars = if evaluated_args.len() >= 2 {
                        self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as usize
                    } else {
                        1
                    };
                    let total_chars = s.chars().count();
                    let skip_chars = total_chars.saturating_sub(num_chars);
                    let suffix: String = s.chars().skip(skip_chars).collect();
                    Ok(ResultData::String(suffix))
                }
                "MID" => {
                    if evaluated_args.len() < 3 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "MID requires 3 arguments".to_string(),
                        )));
                    }
                    let s = evaluated_args[0].to_string();
                    let start_num = self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as usize;
                    let num_chars = self.to_f64(&evaluated_args[2]).unwrap_or(0.0) as usize;

                    let start_idx = start_num.saturating_sub(1);
                    let mid_str: String = s.chars().skip(start_idx).take(num_chars).collect();
                    Ok(ResultData::String(mid_str))
                }
                "LEN" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Float(0.0));
                    }
                    let s = evaluated_args[0].to_string();
                    Ok(ResultData::Float(s.chars().count() as f64))
                }
                "TRIM" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::String(String::new()));
                    }
                    let s = evaluated_args[0].to_string();
                    let trimmed = s.trim();
                    let mut result = String::new();
                    let mut last_was_space = false;
                    for c in trimmed.chars() {
                        if c.is_whitespace() {
                            if !last_was_space {
                                result.push(' ');
                                last_was_space = true;
                            }
                        } else {
                            result.push(c);
                            last_was_space = false;
                        }
                    }
                    Ok(ResultData::String(result))
                }
                "UPPER" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::None);
                    }
                    match &evaluated_args[0] {
                        ResultData::None => Ok(ResultData::String(String::new())),
                        ResultData::Error(e) => Ok(ResultData::Error(e.clone())),
                        v => Ok(ResultData::String(v.to_string().to_uppercase())),
                    }
                }
                "LOWER" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::None);
                    }
                    match &evaluated_args[0] {
                        ResultData::None => Ok(ResultData::String(String::new())),
                        ResultData::Error(e) => Ok(ResultData::Error(e.clone())),
                        v => Ok(ResultData::String(v.to_string().to_lowercase())),
                    }
                }
                "PROPER" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::None);
                    }
                    match &evaluated_args[0] {
                        ResultData::None => Ok(ResultData::String(String::new())),
                        ResultData::Error(e) => Ok(ResultData::Error(e.clone())),
                        v => Ok(ResultData::String(self.proper(&v.to_string()))),
                    }
                }
                "ISNUMBER" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(false));
                    }
                    match &evaluated_args[0] {
                        ResultData::Float(_) | ResultData::Integer(_) => {
                            Ok(ResultData::Boolean(true))
                        }
                        _ => Ok(ResultData::Boolean(false)),
                    }
                }
                "ISTEXT" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(false));
                    }
                    match &evaluated_args[0] {
                        ResultData::String(_) => Ok(ResultData::Boolean(true)),
                        _ => Ok(ResultData::Boolean(false)),
                    }
                }
                "ISBLANK" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(true));
                    }
                    match &evaluated_args[0] {
                        ResultData::None => Ok(ResultData::Boolean(true)),
                        ResultData::String(s) => Ok(ResultData::Boolean(s.is_empty())),
                        _ => Ok(ResultData::Boolean(false)),
                    }
                }
                "ISERROR" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(false));
                    }
                    match &evaluated_args[0] {
                        ResultData::Error(_) => Ok(ResultData::Boolean(true)),
                        _ => Ok(ResultData::Boolean(false)),
                    }
                }
                "PRODUCT" => {
                    if let Some(err) = self.check_arg_errors(&evaluated_args, &arg_is_direct) {
                        return Ok(err);
                    }
                    let mut prod = 1.0;
                    let mut has_nums = false;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        let is_dir = arg_is_direct.get(i).copied().unwrap_or(false);
                        let (p, h) = self.product_helper(arg, is_dir);
                        if h {
                            prod *= p;
                            has_nums = true;
                        }
                    }
                    if has_nums {
                        // Excel snaps a formula's result to 15 significant
                        // digits, and that is observable beyond display:
                        // PRODUCT(-35, -0.617, -40, -34) is
                        // 29369.199999999997 in raw f64, and
                        // ROUNDDOWN(.., 2) of it gives 29369.19, but Excel
                        // answers 29369.2 because the snap happens first.
                        //
                        // Crucially it is applied *once*, to the finished
                        // product. Doing it per factor compounds: over
                        // seven factors PRODUCT drifted ~14 ULP and
                        // rendered 189124133819.665 where Excel gives
                        // 189124133819.664.
                        Ok(ResultData::Float(Self::clean_float(prod)))
                    } else {
                        Ok(ResultData::Float(0.0))
                    }
                }
                "MOD" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "MOD requires 2 arguments".to_string(),
                        )));
                    }
                    let n = self.to_f64_arg(evaluated_args.first(), "MOD")?;
                    let d = self.to_f64_arg(evaluated_args.get(1), "MOD")?;
                    if d == 0.0 {
                        return Ok(ResultData::Error("#DIV/0!".to_string()));
                    }
                    // Excel gives up once the quotient gets large enough
                    // that `n - d * INT(n / d)` stops being meaningful, and
                    // reports #NUM! rather than a number built out of noise
                    // -- MOD(28^31, 3) is #NUM! there, while visi used to
                    // answer 0 from a value 28^31 cannot represent anyway.
                    //
                    // The cutoff is on the quotient, not on either operand
                    // (MOD(1E15, 1E7) is fine, MOD(1E13, 3) is not), and is
                    // identical for different divisors. Bisected against
                    // real Excel to between 1.024 and 1.026 times 2^40; the
                    // exact constant isn't a round number and isn't worth
                    // more probes, so this uses 2^40. That is very slightly
                    // conservative -- inside that 0.2%-wide band visi
                    // reports #NUM! a little before Excel does -- but it is
                    // right everywhere else, which is where the quotients
                    // that actually turn up land.
                    const MOD_QUOTIENT_LIMIT: f64 = 1_099_511_627_776.0; // 2^40
                    let quotient = n / d;
                    if !quotient.is_finite() || quotient.abs() > MOD_QUOTIENT_LIMIT {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    let val = n - d * quotient.floor();
                    Ok(ResultData::Float(val))
                }
                "TODAY" => {
                    let ((y, m, d), _) = self.get_ymd_hms();
                    Ok(ResultData::String(format!("{:04}-{:02}-{:02}", y, m, d)))
                }
                "NOW" => {
                    let ((y, m, d), (hr, min, sec)) = self.get_ymd_hms();
                    Ok(ResultData::String(format!(
                        "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                        y, m, d, hr, min, sec
                    )))
                }
                "MATCH" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "MATCH requires at least 2 arguments".to_string(),
                        )));
                    }
                    let lookup_val = &evaluated_args[0];
                    let match_type = if evaluated_args.len() >= 3 {
                        self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as isize
                    } else {
                        1
                    };

                    if let ResultData::List(list) = &evaluated_args[1] {
                        let mut match_idx: Option<usize> = None;
                        if match_type == 0 {
                            for (idx, item) in list.iter().enumerate() {
                                if Self::exact_lookup_matches(lookup_val, item) {
                                    match_idx = Some(idx);
                                    break;
                                }
                            }
                        } else if match_type == 1 {
                            let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                            for (idx, item) in list.iter().enumerate() {
                                let item_f = self.to_f64(item).unwrap_or(0.0);
                                if item_f <= lookup_f {
                                    match_idx = Some(idx);
                                } else {
                                    break;
                                }
                            }
                        } else if match_type == -1 {
                            let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                            for (idx, item) in list.iter().enumerate() {
                                let item_f = self.to_f64(item).unwrap_or(0.0);
                                if item_f >= lookup_f {
                                    match_idx = Some(idx);
                                } else {
                                    break;
                                }
                            }
                        }

                        if let Some(idx) = match_idx {
                            Ok(ResultData::Integer((idx + 1) as i64))
                        } else {
                            Ok(ResultData::Error("#N/A".to_string()))
                        }
                    } else {
                        Ok(ResultData::Error("#N/A".to_string()))
                    }
                }
                "LOOKUP" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "LOOKUP requires at least 2 arguments".to_string(),
                        )));
                    }
                    let lookup_val = &evaluated_args[0];
                    let lookup_vec: Vec<ResultData> = match &evaluated_args[1] {
                        ResultData::List(l) => l.clone(),
                        other => vec![other.clone()],
                    };
                    let result_vec: Vec<ResultData> = match evaluated_args.get(2) {
                        Some(ResultData::List(l)) => l.clone(),
                        Some(other) => vec![other.clone()],
                        None => lookup_vec.clone(),
                    };
                    let lookup_f = self.to_f64(lookup_val);
                    let mut match_idx: Option<usize> = None;
                    for (idx, item) in lookup_vec.iter().enumerate() {
                        let is_match = match lookup_f {
                            Some(lf) => self.to_f64(item).map(|f| f <= lf).unwrap_or(false),
                            None => item.to_string() <= lookup_val.to_string(),
                        };
                        if is_match {
                            match_idx = Some(idx);
                        } else {
                            break;
                        }
                    }
                    match match_idx {
                        Some(idx) => Ok(result_vec
                            .get(idx)
                            .cloned()
                            .unwrap_or(ResultData::Error("#N/A".to_string()))),
                        None => Ok(ResultData::Error("#N/A".to_string())),
                    }
                }
                "XMATCH" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "XMATCH requires at least 2 arguments".to_string(),
                        )));
                    }
                    let lookup_val = &evaluated_args[0];
                    let arr: Vec<ResultData> = match &evaluated_args[1] {
                        ResultData::List(l) => l.clone(),
                        other => vec![other.clone()],
                    };
                    // search_mode (a 4th argument) isn't supported beyond
                    // the default forward linear search, nor is wildcard
                    // match_mode (2).
                    let match_mode = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(0.0) as isize;
                    let found = match match_mode {
                        -1 | 1 => {
                            let lf = self.to_f64(lookup_val).unwrap_or(0.0);
                            let mut best: Option<(usize, f64)> = None;
                            for (i, item) in arr.iter().enumerate() {
                                if let Some(f) = self.to_f64(item) {
                                    let candidate =
                                        if match_mode == -1 { f <= lf } else { f >= lf };
                                    let better = match best {
                                        Some((_, bf)) => {
                                            if match_mode == -1 {
                                                f > bf
                                            } else {
                                                f < bf
                                            }
                                        }
                                        None => true,
                                    };
                                    if candidate && better {
                                        best = Some((i, f));
                                    }
                                }
                            }
                            best.map(|(i, _)| i)
                        }
                        // XMATCH deliberately does NOT use
                        // exact_lookup_matches: unlike MATCH/VLOOKUP,
                        // real Excel's XMATCH *does* match a blank lookup
                        // value against a blank cell (XMATCH over a blank
                        // A1 in A1:A4 returns 1 where MATCH returns #N/A).
                        _ => arr
                            .iter()
                            .position(|item| item.to_string() == lookup_val.to_string()),
                    };
                    match found {
                        Some(i) => Ok(ResultData::Float((i + 1) as f64)),
                        None => Ok(ResultData::Error("#N/A".to_string())),
                    }
                }
                "VLOOKUP" => {
                    if evaluated_args.len() < 3 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "VLOOKUP requires at least 3 arguments".to_string(),
                        )));
                    }
                    let lookup_val = &evaluated_args[0];
                    let col_idx = self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as usize - 1;
                    let range_lookup = if evaluated_args.len() >= 4 {
                        self.to_bool(&evaluated_args[3])
                    } else {
                        true
                    };

                    if let ResultData::List(list) = &evaluated_args[1] {
                        let num_cols = match &args[1] {
                            Expr::RangeRef {
                                start_col, end_col, ..
                            } => end_col - start_col + 1,
                            _ => 1,
                        };

                        let num_rows = list.len() / num_cols;
                        if num_rows == 0 || num_cols == 0 {
                            return Ok(ResultData::Error("#N/A".to_string()));
                        }

                        let mut found_row_idx: Option<usize> = None;
                        if !range_lookup {
                            for r in 0..num_rows {
                                let first_col_val = &list[r * num_cols];
                                if Self::exact_lookup_matches(lookup_val, first_col_val) {
                                    found_row_idx = Some(r);
                                    break;
                                }
                            }
                        } else {
                            let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                            for r in 0..num_rows {
                                let first_col_val = &list[r * num_cols];
                                let val_f = self.to_f64(first_col_val).unwrap_or(0.0);
                                if val_f <= lookup_f {
                                    found_row_idx = Some(r);
                                } else {
                                    break;
                                }
                            }
                        }

                        if let Some(r) = found_row_idx {
                            if col_idx < num_cols {
                                Ok(list[r * num_cols + col_idx].clone())
                            } else {
                                Ok(ResultData::Error("#REF!".to_string()))
                            }
                        } else {
                            Ok(ResultData::Error("#N/A".to_string()))
                        }
                    } else {
                        Ok(ResultData::Error("#N/A".to_string()))
                    }
                }
                "XLOOKUP" => {
                    if evaluated_args.len() < 3 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "XLOOKUP requires at least 3 arguments".to_string(),
                        )));
                    }
                    let lookup_val = &evaluated_args[0];
                    let if_not_found = if evaluated_args.len() >= 4 {
                        evaluated_args[3].clone()
                    } else {
                        ResultData::Error("#N/A".to_string())
                    };
                    let search_mode = if evaluated_args.len() >= 6 {
                        self.to_f64(&evaluated_args[5]).unwrap_or(1.0) as isize
                    } else {
                        1
                    };

                    if let (ResultData::List(lookup_list), ResultData::List(return_list)) =
                        (&evaluated_args[1], &evaluated_args[2])
                    {
                        let mut found_idx: Option<usize> = None;
                        let len = lookup_list.len();

                        let iter_indices: Vec<usize> = if search_mode == -1 {
                            (0..len).rev().collect()
                        } else {
                            (0..len).collect()
                        };

                        for idx in iter_indices {
                            // Like XMATCH (and unlike VLOOKUP/MATCH),
                            // XLOOKUP matches a blank lookup value against
                            // a blank cell rather than reporting #N/A.
                            if lookup_list[idx].to_string() == lookup_val.to_string() {
                                found_idx = Some(idx);
                                break;
                            }
                        }

                        if let Some(idx) = found_idx {
                            if idx < return_list.len() {
                                Ok(return_list[idx].clone())
                            } else {
                                Ok(ResultData::Error("#REF!".to_string()))
                            }
                        } else {
                            Ok(if_not_found)
                        }
                    } else {
                        Ok(if_not_found)
                    }
                }
                "SUMIF" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "SUMIF requires at least 2 arguments".to_string(),
                        )));
                    }
                    let range_list = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let criteria = &evaluated_args[1];
                    let sum_list = if evaluated_args.len() >= 3 {
                        match &evaluated_args[2] {
                            ResultData::List(l) => l,
                            _ => range_list,
                        }
                    } else {
                        range_list
                    };

                    let mut sum = 0.0;
                    for idx in 0..range_list.len() {
                        if idx < sum_list.len() && self.match_criteria(&range_list[idx], criteria) {
                            sum += Self::aggregate_range_number(&sum_list[idx]).unwrap_or(0.0);
                        }
                    }
                    Ok(ResultData::Float(sum))
                }
                "SUMIFS" => {
                    if evaluated_args.len() < 3 || evaluated_args.len() % 2 == 0 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "SUMIFS requires sum_range and at least one criteria_range/criteria pair".to_string(),
                        )));
                    }
                    let sum_list = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };

                    let mut criteria_pairs = Vec::new();
                    let mut i = 1;
                    while i < evaluated_args.len() {
                        let crit_range = match &evaluated_args[i] {
                            ResultData::List(l) => l,
                            _ => return Ok(ResultData::Float(0.0)),
                        };
                        let crit_val = &evaluated_args[i + 1];
                        criteria_pairs.push((crit_range, crit_val));
                        i += 2;
                    }

                    let mut sum = 0.0;
                    for idx in 0..sum_list.len() {
                        let mut all_match = true;
                        for (crit_range, crit_val) in &criteria_pairs {
                            if idx >= crit_range.len()
                                || !self.match_criteria(&crit_range[idx], crit_val)
                            {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match {
                            sum += Self::aggregate_range_number(&sum_list[idx]).unwrap_or(0.0);
                        }
                    }
                    Ok(ResultData::Float(sum))
                }
                "COUNTIF" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "COUNTIF requires 2 arguments".to_string(),
                        )));
                    }
                    let range_list = match &evaluated_args[0] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let criteria = &evaluated_args[1];
                    let mut count = 0;
                    for val in range_list {
                        if self.match_criteria(val, criteria) {
                            count += 1;
                        }
                    }
                    Ok(ResultData::Float(count as f64))
                }
                "COUNTIFS" => {
                    if evaluated_args.len() < 2 || evaluated_args.len() % 2 != 0 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "COUNTIFS requires at least one criteria_range/criteria pair"
                                .to_string(),
                        )));
                    }
                    let mut criteria_pairs = Vec::new();
                    let mut i = 0;
                    while i < evaluated_args.len() {
                        let crit_range = match &evaluated_args[i] {
                            ResultData::List(l) => l,
                            _ => return Ok(ResultData::Float(0.0)),
                        };
                        let crit_val = &evaluated_args[i + 1];
                        criteria_pairs.push((crit_range, crit_val));
                        i += 2;
                    }

                    let mut count = 0;
                    let first_len = criteria_pairs[0].0.len();
                    for idx in 0..first_len {
                        let mut all_match = true;
                        for (crit_range, crit_val) in &criteria_pairs {
                            if idx >= crit_range.len()
                                || !self.match_criteria(&crit_range[idx], crit_val)
                            {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match {
                            count += 1;
                        }
                    }
                    Ok(ResultData::Float(count as f64))
                }
                "MMULT" => {
                    if evaluated_args.len() < 2 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "MMULT requires 2 arguments".to_string(),
                        )));
                    }

                    if let (ResultData::List(list1), ResultData::List(list2)) =
                        (&evaluated_args[0], &evaluated_args[1])
                    {
                        let (rows1, cols1) = match &args[0] {
                            Expr::RangeRef {
                                sheet,
                                start_row,
                                end_row,
                                start_col,
                                end_col,
                                ..
                            } => {
                                let is_self = match sheet {
                                    Some(name) => name == &self.name,
                                    None => true,
                                };
                                let actual_end = if *end_row == usize::MAX {
                                    if is_self {
                                        self.row_count().saturating_sub(1)
                                    } else if let Some(ctx) = context {
                                        ctx.sheets
                                            .get(sheet.as_ref().unwrap())
                                            .map(|t| t.row_count().saturating_sub(1))
                                            .unwrap_or(0)
                                    } else {
                                        0
                                    }
                                } else {
                                    *end_row
                                };
                                (actual_end - start_row + 1, end_col - start_col + 1)
                            }
                            _ => (1, list1.len()),
                        };

                        let (rows2, cols2) = match &args[1] {
                            Expr::RangeRef {
                                sheet,
                                start_row,
                                end_row,
                                start_col,
                                end_col,
                                ..
                            } => {
                                let is_self = match sheet {
                                    Some(name) => name == &self.name,
                                    None => true,
                                };
                                let actual_end = if *end_row == usize::MAX {
                                    if is_self {
                                        self.row_count().saturating_sub(1)
                                    } else if let Some(ctx) = context {
                                        ctx.sheets
                                            .get(sheet.as_ref().unwrap())
                                            .map(|t| t.row_count().saturating_sub(1))
                                            .unwrap_or(0)
                                    } else {
                                        0
                                    }
                                } else {
                                    *end_row
                                };
                                (actual_end - start_row + 1, end_col - start_col + 1)
                            }
                            _ => {
                                if list2.len() == cols1 {
                                    (cols1, 1)
                                } else {
                                    (1, list2.len())
                                }
                            }
                        };

                        if cols1 != rows2 {
                            return Ok(ResultData::Error("#VALUE!".to_string()));
                        }

                        // A non-numeric cell anywhere in either operand
                        // makes the whole call #VALUE! in real Excel, not
                        // a silent 0 -- MMULT doesn't ignore text the way
                        // SUM/AVERAGE-style aggregates do.
                        fn as_plain_number(v: &ResultData) -> Option<f64> {
                            match v {
                                ResultData::Float(f) => Some(*f),
                                ResultData::Integer(i) => Some(*i as f64),
                                _ => None,
                            }
                        }
                        let mut result_list = Vec::with_capacity(rows1 * cols2);
                        for r in 0..rows1 {
                            for c in 0..cols2 {
                                let mut val = 0.0;
                                for k in 0..cols1 {
                                    // Only a real number is acceptable --
                                    // MMULT rejects text, booleans and
                                    // blanks alike (all confirmed #VALUE!
                                    // against real Excel), so this can't
                                    // use to_f64's lenient coercion.
                                    let (Some(v1), Some(v2)) = (
                                        as_plain_number(&list1[r * cols1 + k]),
                                        as_plain_number(&list2[k * cols2 + c]),
                                    ) else {
                                        return Ok(ResultData::Error("#VALUE!".to_string()));
                                    };
                                    val += v1 * v2;
                                }
                                result_list.push(ResultData::Float(val));
                            }
                        }
                        Ok(ResultData::List(result_list))
                    } else {
                        Ok(ResultData::Error("#VALUE!".to_string()))
                    }
                }
                "PV" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "PV")?;
                    let nper = self.to_f64_arg(evaluated_args.get(1), "PV")?;
                    let pmt = self.to_f64_arg(evaluated_args.get(2), "PV")?;
                    let fv = self.opt_f64(&evaluated_args, 3, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::pv(
                        rate, nper, pmt, fv, pmt_type,
                    )))
                }
                "FV" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "FV")?;
                    let nper = self.to_f64_arg(evaluated_args.get(1), "FV")?;
                    let pmt = self.to_f64_arg(evaluated_args.get(2), "FV")?;
                    let pv = self.opt_f64(&evaluated_args, 3, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::fv(
                        rate, nper, pmt, pv, pmt_type,
                    )))
                }
                "PMT" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "PMT")?;
                    let nper = self.to_f64_arg(evaluated_args.get(1), "PMT")?;
                    let pv = self.to_f64_arg(evaluated_args.get(2), "PMT")?;
                    let fv = self.opt_f64(&evaluated_args, 3, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::pmt(
                        rate, nper, pv, fv, pmt_type,
                    )))
                }
                "NPER" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "NPER")?;
                    let pmt = self.to_f64_arg(evaluated_args.get(1), "NPER")?;
                    let pv = self.to_f64_arg(evaluated_args.get(2), "NPER")?;
                    let fv = self.opt_f64(&evaluated_args, 3, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 4, 0.0);
                    match finance::nper(rate, pmt, pv, fv, pmt_type) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "RATE" => {
                    let nper = self.to_f64_arg(evaluated_args.first(), "RATE")?;
                    let pmt = self.to_f64_arg(evaluated_args.get(1), "RATE")?;
                    let pv = self.to_f64_arg(evaluated_args.get(2), "RATE")?;
                    let fv = self.opt_f64(&evaluated_args, 3, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 4, 0.0);
                    let guess = self.opt_f64(&evaluated_args, 5, 0.1);
                    match finance::rate(nper, pmt, pv, fv, pmt_type, guess) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "IPMT" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "IPMT")?;
                    let per = self.to_f64_arg(evaluated_args.get(1), "IPMT")?;
                    let nper = self.to_f64_arg(evaluated_args.get(2), "IPMT")?;
                    let pv = self.to_f64_arg(evaluated_args.get(3), "IPMT")?;
                    let fv = self.opt_f64(&evaluated_args, 4, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 5, 0.0);
                    Ok(ResultData::Float(finance::ipmt(
                        rate, per, nper, pv, fv, pmt_type,
                    )))
                }
                "PPMT" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "PPMT")?;
                    let per = self.to_f64_arg(evaluated_args.get(1), "PPMT")?;
                    let nper = self.to_f64_arg(evaluated_args.get(2), "PPMT")?;
                    let pv = self.to_f64_arg(evaluated_args.get(3), "PPMT")?;
                    let fv = self.opt_f64(&evaluated_args, 4, 0.0);
                    let pmt_type = self.opt_f64(&evaluated_args, 5, 0.0);
                    Ok(ResultData::Float(finance::ppmt(
                        rate, per, nper, pv, fv, pmt_type,
                    )))
                }
                "CUMIPMT" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "CUMIPMT")?;
                    let nper = self.to_f64_arg(evaluated_args.get(1), "CUMIPMT")?;
                    let pv = self.to_f64_arg(evaluated_args.get(2), "CUMIPMT")?;
                    let start = self.to_f64_arg(evaluated_args.get(3), "CUMIPMT")?;
                    let end = self.to_f64_arg(evaluated_args.get(4), "CUMIPMT")?;
                    let pmt_type = self.to_f64_arg(evaluated_args.get(5), "CUMIPMT")?;
                    Ok(ResultData::Float(finance::cumipmt(
                        rate, nper, pv, start, end, pmt_type,
                    )))
                }
                "CUMPRINC" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "CUMPRINC")?;
                    let nper = self.to_f64_arg(evaluated_args.get(1), "CUMPRINC")?;
                    let pv = self.to_f64_arg(evaluated_args.get(2), "CUMPRINC")?;
                    let start = self.to_f64_arg(evaluated_args.get(3), "CUMPRINC")?;
                    let end = self.to_f64_arg(evaluated_args.get(4), "CUMPRINC")?;
                    let pmt_type = self.to_f64_arg(evaluated_args.get(5), "CUMPRINC")?;
                    Ok(ResultData::Float(finance::cumprinc(
                        rate, nper, pv, start, end, pmt_type,
                    )))
                }
                "NPV" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "NPV")?;
                    let mut values = Vec::new();
                    for (i, arg) in evaluated_args.iter().enumerate().skip(1) {
                        values.extend(self.flatten_finance_numbers(arg, arg_is_direct[i]));
                    }
                    Ok(ResultData::Float(finance::npv(rate, &values)))
                }
                "IRR" => {
                    let values = evaluated_args
                        .first()
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[0]))
                        .unwrap_or_default();
                    let guess = self.opt_f64(&evaluated_args, 1, 0.1);
                    match finance::irr(&values, guess) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "MIRR" => {
                    let values = evaluated_args
                        .first()
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[0]))
                        .unwrap_or_default();
                    let finance_rate = self.to_f64_arg(evaluated_args.get(1), "MIRR")?;
                    let reinvest_rate = self.to_f64_arg(evaluated_args.get(2), "MIRR")?;
                    match finance::mirr(&values, finance_rate, reinvest_rate) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "XNPV" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "XNPV")?;
                    let values = evaluated_args
                        .get(1)
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[1]))
                        .unwrap_or_default();
                    let dates = evaluated_args
                        .get(2)
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[2]))
                        .unwrap_or_default();
                    if values.is_empty() || values.len() != dates.len() {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    Ok(ResultData::Float(finance::xnpv(rate, &values, &dates)))
                }
                "XIRR" => {
                    let values = evaluated_args
                        .first()
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[0]))
                        .unwrap_or_default();
                    let dates = evaluated_args
                        .get(1)
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[1]))
                        .unwrap_or_default();
                    let guess = self.opt_f64(&evaluated_args, 2, 0.1);
                    if values.is_empty() || values.len() != dates.len() {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    match finance::xirr(&values, &dates, guess) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "SLN" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "SLN")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(1), "SLN")?;
                    let life = self.to_f64_arg(evaluated_args.get(2), "SLN")?;
                    Ok(ResultData::Float(finance::sln(cost, salvage, life)))
                }
                "SYD" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "SYD")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(1), "SYD")?;
                    let life = self.to_f64_arg(evaluated_args.get(2), "SYD")?;
                    let per = self.to_f64_arg(evaluated_args.get(3), "SYD")?;
                    Ok(ResultData::Float(finance::syd(cost, salvage, life, per)))
                }
                "DB" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "DB")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(1), "DB")?;
                    let life = self.to_f64_arg(evaluated_args.get(2), "DB")?;
                    let period = self.to_f64_arg(evaluated_args.get(3), "DB")?;
                    let month = self.opt_f64(&evaluated_args, 4, 12.0);
                    Ok(ResultData::Float(finance::db(
                        cost, salvage, life, period, month,
                    )))
                }
                "DDB" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "DDB")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(1), "DDB")?;
                    let life = self.to_f64_arg(evaluated_args.get(2), "DDB")?;
                    let period = self.to_f64_arg(evaluated_args.get(3), "DDB")?;
                    let factor = self.opt_f64(&evaluated_args, 4, 2.0);
                    Ok(ResultData::Float(finance::ddb(
                        cost, salvage, life, period, factor,
                    )))
                }
                "VDB" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "VDB")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(1), "VDB")?;
                    let life = self.to_f64_arg(evaluated_args.get(2), "VDB")?;
                    let start = self.to_f64_arg(evaluated_args.get(3), "VDB")?;
                    let end = self.to_f64_arg(evaluated_args.get(4), "VDB")?;
                    let factor = self.opt_f64(&evaluated_args, 5, 2.0);
                    let no_switch = evaluated_args
                        .get(6)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(false);
                    match finance::vdb(cost, salvage, life, start, end, factor, no_switch) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "EFFECT" => {
                    let nominal_rate = self.to_f64_arg(evaluated_args.first(), "EFFECT")?;
                    let npery = self.to_f64_arg(evaluated_args.get(1), "EFFECT")?;
                    Ok(ResultData::Float(finance::effect(nominal_rate, npery)))
                }
                "NOMINAL" => {
                    let effect_rate = self.to_f64_arg(evaluated_args.first(), "NOMINAL")?;
                    let npery = self.to_f64_arg(evaluated_args.get(1), "NOMINAL")?;
                    Ok(ResultData::Float(finance::nominal(effect_rate, npery)))
                }
                "DOLLARDE" => {
                    let fractional_dollar = self.to_f64_arg(evaluated_args.first(), "DOLLARDE")?;
                    let fraction = self.to_f64_arg(evaluated_args.get(1), "DOLLARDE")?;
                    match finance::dollarde(fractional_dollar, fraction) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "DOLLARFR" => {
                    let decimal_dollar = self.to_f64_arg(evaluated_args.first(), "DOLLARFR")?;
                    let fraction = self.to_f64_arg(evaluated_args.get(1), "DOLLARFR")?;
                    match finance::dollarfr(decimal_dollar, fraction) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "FVSCHEDULE" => {
                    let principal = self.to_f64_arg(evaluated_args.first(), "FVSCHEDULE")?;
                    let schedule = evaluated_args
                        .get(1)
                        .map(|v| self.flatten_finance_numbers(v, arg_is_direct[1]))
                        .unwrap_or_default();
                    Ok(ResultData::Float(finance::fvschedule(principal, &schedule)))
                }
                "RRI" => {
                    let nper = self.to_f64_arg(evaluated_args.first(), "RRI")?;
                    let pv = self.to_f64_arg(evaluated_args.get(1), "RRI")?;
                    let fv = self.to_f64_arg(evaluated_args.get(2), "RRI")?;
                    match finance::rri(nper, pv, fv) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "PDURATION" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "PDURATION")?;
                    let pv = self.to_f64_arg(evaluated_args.get(1), "PDURATION")?;
                    let fv = self.to_f64_arg(evaluated_args.get(2), "PDURATION")?;
                    match finance::pduration(rate, pv, fv) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "ISPMT" => {
                    let rate = self.to_f64_arg(evaluated_args.first(), "ISPMT")?;
                    let per = self.to_f64_arg(evaluated_args.get(1), "ISPMT")?;
                    let nper = self.to_f64_arg(evaluated_args.get(2), "ISPMT")?;
                    let pv = self.to_f64_arg(evaluated_args.get(3), "ISPMT")?;
                    Ok(ResultData::Float(finance::ispmt(rate, per, nper, pv)))
                }
                "COUPDAYBS" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "COUPDAYBS")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPDAYBS")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPDAYBS")?;
                    let basis = self.opt_f64(&evaluated_args, 3, 0.0);
                    Ok(ResultData::Float(finance::coupdaybs(
                        settlement, maturity, frequency, basis,
                    )))
                }
                "COUPDAYS" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "COUPDAYS")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPDAYS")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPDAYS")?;
                    let basis = self.opt_f64(&evaluated_args, 3, 0.0);
                    Ok(ResultData::Float(finance::coupdays(
                        settlement, maturity, frequency, basis,
                    )))
                }
                "COUPDAYSNC" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "COUPDAYSNC")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPDAYSNC")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPDAYSNC")?;
                    let basis = self.opt_f64(&evaluated_args, 3, 0.0);
                    Ok(ResultData::Float(finance::coupdaysnc(
                        settlement, maturity, frequency, basis,
                    )))
                }
                "COUPNCD" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "COUPNCD")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPNCD")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPNCD")?;
                    Ok(ResultData::Float(finance::coupncd(
                        settlement, maturity, frequency,
                    )))
                }
                "COUPNUM" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "COUPNUM")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPNUM")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPNUM")?;
                    Ok(ResultData::Float(finance::coupnum(
                        settlement, maturity, frequency,
                    )))
                }
                "COUPPCD" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "COUPPCD")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPPCD")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPPCD")?;
                    Ok(ResultData::Float(finance::couppcd(
                        settlement, maturity, frequency,
                    )))
                }
                "PRICE" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "PRICE")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "PRICE")?;
                    let rate = self.to_f64_arg(evaluated_args.get(2), "PRICE")?;
                    let yld = self.to_f64_arg(evaluated_args.get(3), "PRICE")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(4), "PRICE")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(5), "PRICE")?;
                    let basis = self.opt_f64(&evaluated_args, 6, 0.0);
                    Ok(ResultData::Float(finance::price(
                        settlement, maturity, rate, yld, redemption, frequency, basis,
                    )))
                }
                "YIELD" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "YIELD")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "YIELD")?;
                    let rate = self.to_f64_arg(evaluated_args.get(2), "YIELD")?;
                    let pr = self.to_f64_arg(evaluated_args.get(3), "YIELD")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(4), "YIELD")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(5), "YIELD")?;
                    let basis = self.opt_f64(&evaluated_args, 6, 0.0);
                    match finance::yield_(
                        settlement, maturity, rate, pr, redemption, frequency, basis,
                    ) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "DURATION" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "DURATION")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "DURATION")?;
                    let coupon = self.to_f64_arg(evaluated_args.get(2), "DURATION")?;
                    let yld = self.to_f64_arg(evaluated_args.get(3), "DURATION")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(4), "DURATION")?;
                    let basis = self.opt_f64(&evaluated_args, 5, 0.0);
                    Ok(ResultData::Float(finance::duration(
                        settlement, maturity, coupon, yld, frequency, basis,
                    )))
                }
                "MDURATION" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "MDURATION")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "MDURATION")?;
                    let coupon = self.to_f64_arg(evaluated_args.get(2), "MDURATION")?;
                    let yld = self.to_f64_arg(evaluated_args.get(3), "MDURATION")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(4), "MDURATION")?;
                    let basis = self.opt_f64(&evaluated_args, 5, 0.0);
                    Ok(ResultData::Float(finance::mduration(
                        settlement, maturity, coupon, yld, frequency, basis,
                    )))
                }
                "DISC" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "DISC")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "DISC")?;
                    let pr = self.to_f64_arg(evaluated_args.get(2), "DISC")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(3), "DISC")?;
                    let basis = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::disc(
                        settlement, maturity, pr, redemption, basis,
                    )))
                }
                "PRICEDISC" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "PRICEDISC")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "PRICEDISC")?;
                    let discount = self.to_f64_arg(evaluated_args.get(2), "PRICEDISC")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(3), "PRICEDISC")?;
                    let basis = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::pricedisc(
                        settlement, maturity, discount, redemption, basis,
                    )))
                }
                "YIELDDISC" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "YIELDDISC")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "YIELDDISC")?;
                    let pr = self.to_f64_arg(evaluated_args.get(2), "YIELDDISC")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(3), "YIELDDISC")?;
                    let basis = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::yielddisc(
                        settlement, maturity, pr, redemption, basis,
                    )))
                }
                "PRICEMAT" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "PRICEMAT")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "PRICEMAT")?;
                    let issue = self.to_f64_arg(evaluated_args.get(2), "PRICEMAT")?;
                    let rate = self.to_f64_arg(evaluated_args.get(3), "PRICEMAT")?;
                    let yld = self.to_f64_arg(evaluated_args.get(4), "PRICEMAT")?;
                    let basis = self.opt_f64(&evaluated_args, 5, 0.0);
                    Ok(ResultData::Float(finance::pricemat(
                        settlement, maturity, issue, rate, yld, basis,
                    )))
                }
                "YIELDMAT" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "YIELDMAT")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "YIELDMAT")?;
                    let issue = self.to_f64_arg(evaluated_args.get(2), "YIELDMAT")?;
                    let rate = self.to_f64_arg(evaluated_args.get(3), "YIELDMAT")?;
                    let pr = self.to_f64_arg(evaluated_args.get(4), "YIELDMAT")?;
                    let basis = self.opt_f64(&evaluated_args, 5, 0.0);
                    Ok(ResultData::Float(finance::yieldmat(
                        settlement, maturity, issue, rate, pr, basis,
                    )))
                }
                "RECEIVED" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "RECEIVED")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "RECEIVED")?;
                    let investment = self.to_f64_arg(evaluated_args.get(2), "RECEIVED")?;
                    let discount = self.to_f64_arg(evaluated_args.get(3), "RECEIVED")?;
                    let basis = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::received(
                        settlement, maturity, investment, discount, basis,
                    )))
                }
                "INTRATE" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "INTRATE")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "INTRATE")?;
                    let investment = self.to_f64_arg(evaluated_args.get(2), "INTRATE")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(3), "INTRATE")?;
                    let basis = self.opt_f64(&evaluated_args, 4, 0.0);
                    Ok(ResultData::Float(finance::intrate(
                        settlement, maturity, investment, redemption, basis,
                    )))
                }
                "TBILLPRICE" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "TBILLPRICE")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "TBILLPRICE")?;
                    let discount = self.to_f64_arg(evaluated_args.get(2), "TBILLPRICE")?;
                    Ok(ResultData::Float(finance::tbillprice(
                        settlement, maturity, discount,
                    )))
                }
                "TBILLYIELD" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "TBILLYIELD")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "TBILLYIELD")?;
                    let pr = self.to_f64_arg(evaluated_args.get(2), "TBILLYIELD")?;
                    Ok(ResultData::Float(finance::tbillyield(
                        settlement, maturity, pr,
                    )))
                }
                "TBILLEQ" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "TBILLEQ")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "TBILLEQ")?;
                    let discount = self.to_f64_arg(evaluated_args.get(2), "TBILLEQ")?;
                    match finance::tbilleq(settlement, maturity, discount) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "ACCRINTM" => {
                    let issue = self.to_f64_arg(evaluated_args.first(), "ACCRINTM")?;
                    let settlement = self.to_f64_arg(evaluated_args.get(1), "ACCRINTM")?;
                    let rate = self.to_f64_arg(evaluated_args.get(2), "ACCRINTM")?;
                    let par = self.to_f64_arg(evaluated_args.get(3), "ACCRINTM")?;
                    let basis = self.opt_f64(&evaluated_args, 4, 0.0);
                    res_to_rd(finance::accrintm(issue, settlement, rate, par, basis))
                }
                "ACCRINT" => {
                    let issue = self.to_f64_arg(evaluated_args.first(), "ACCRINT")?;
                    let first_interest = self.to_f64_arg(evaluated_args.get(1), "ACCRINT")?;
                    let settlement = self.to_f64_arg(evaluated_args.get(2), "ACCRINT")?;
                    let rate = self.to_f64_arg(evaluated_args.get(3), "ACCRINT")?;
                    let par = self.to_f64_arg(evaluated_args.get(4), "ACCRINT")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(5), "ACCRINT")?;
                    let basis = self.opt_f64(&evaluated_args, 6, 0.0);
                    let calc_method = evaluated_args
                        .get(7)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true);
                    Ok(ResultData::Float(finance::accrint(
                        issue,
                        first_interest,
                        settlement,
                        rate,
                        par,
                        frequency,
                        basis,
                        calc_method,
                    )))
                }
                "AMORLINC" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "AMORLINC")?;
                    let date_purchased = self.to_f64_arg(evaluated_args.get(1), "AMORLINC")?;
                    let first_period = self.to_f64_arg(evaluated_args.get(2), "AMORLINC")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(3), "AMORLINC")?;
                    let period = self.to_f64_arg(evaluated_args.get(4), "AMORLINC")?;
                    let rate = self.to_f64_arg(evaluated_args.get(5), "AMORLINC")?;
                    let basis = self.opt_f64(&evaluated_args, 6, 0.0);
                    res_to_rd(finance::amorlinc(
                        cost,
                        date_purchased,
                        first_period,
                        salvage,
                        period,
                        rate,
                        basis,
                    ))
                }
                "AMORDEGRC" => {
                    let cost = self.to_f64_arg(evaluated_args.first(), "AMORDEGRC")?;
                    let date_purchased = self.to_f64_arg(evaluated_args.get(1), "AMORDEGRC")?;
                    let first_period = self.to_f64_arg(evaluated_args.get(2), "AMORDEGRC")?;
                    let salvage = self.to_f64_arg(evaluated_args.get(3), "AMORDEGRC")?;
                    let period = self.to_f64_arg(evaluated_args.get(4), "AMORDEGRC")?;
                    let rate = self.to_f64_arg(evaluated_args.get(5), "AMORDEGRC")?;
                    let basis = self.opt_f64(&evaluated_args, 6, 0.0);
                    res_to_rd(finance::amordegrc(
                        cost,
                        date_purchased,
                        first_period,
                        salvage,
                        period,
                        rate,
                        basis,
                    ))
                }
                "ODDFPRICE" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "ODDFPRICE")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDFPRICE")?;
                    let issue = self.to_f64_arg(evaluated_args.get(2), "ODDFPRICE")?;
                    let first_coupon = self.to_f64_arg(evaluated_args.get(3), "ODDFPRICE")?;
                    let rate = self.to_f64_arg(evaluated_args.get(4), "ODDFPRICE")?;
                    let yld = self.to_f64_arg(evaluated_args.get(5), "ODDFPRICE")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(6), "ODDFPRICE")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(7), "ODDFPRICE")?;
                    let basis = self.opt_f64(&evaluated_args, 8, 0.0);
                    if settlement <= issue {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    Ok(ResultData::Float(finance::oddfprice(
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        yld,
                        redemption,
                        frequency,
                        basis,
                    )))
                }
                "ODDFYIELD" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "ODDFYIELD")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDFYIELD")?;
                    let issue = self.to_f64_arg(evaluated_args.get(2), "ODDFYIELD")?;
                    let first_coupon = self.to_f64_arg(evaluated_args.get(3), "ODDFYIELD")?;
                    let rate = self.to_f64_arg(evaluated_args.get(4), "ODDFYIELD")?;
                    let pr = self.to_f64_arg(evaluated_args.get(5), "ODDFYIELD")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(6), "ODDFYIELD")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(7), "ODDFYIELD")?;
                    let basis = self.opt_f64(&evaluated_args, 8, 0.0);
                    if settlement <= issue {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    match finance::oddfyield(
                        settlement,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        pr,
                        redemption,
                        frequency,
                        basis,
                    ) {
                        Some(v) => Ok(ResultData::Float(v)),
                        None => Ok(ResultData::Error("#NUM!".to_string())),
                    }
                }
                "ODDLPRICE" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "ODDLPRICE")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDLPRICE")?;
                    let last_interest = self.to_f64_arg(evaluated_args.get(2), "ODDLPRICE")?;
                    let rate = self.to_f64_arg(evaluated_args.get(3), "ODDLPRICE")?;
                    let yld = self.to_f64_arg(evaluated_args.get(4), "ODDLPRICE")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(5), "ODDLPRICE")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(6), "ODDLPRICE")?;
                    let basis = self.opt_f64(&evaluated_args, 7, 0.0);
                    if settlement <= last_interest {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    Ok(ResultData::Float(finance::oddlprice(
                        settlement,
                        maturity,
                        last_interest,
                        rate,
                        yld,
                        redemption,
                        frequency,
                        basis,
                    )))
                }
                "ODDLYIELD" => {
                    let settlement = self.to_f64_arg(evaluated_args.first(), "ODDLYIELD")?;
                    let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDLYIELD")?;
                    let last_interest = self.to_f64_arg(evaluated_args.get(2), "ODDLYIELD")?;
                    let rate = self.to_f64_arg(evaluated_args.get(3), "ODDLYIELD")?;
                    let pr = self.to_f64_arg(evaluated_args.get(4), "ODDLYIELD")?;
                    let redemption = self.to_f64_arg(evaluated_args.get(5), "ODDLYIELD")?;
                    let frequency = self.to_f64_arg(evaluated_args.get(6), "ODDLYIELD")?;
                    let basis = self.opt_f64(&evaluated_args, 7, 0.0);
                    if settlement <= last_interest {
                        return Ok(ResultData::Error("#NUM!".to_string()));
                    }
                    Ok(ResultData::Float(finance::oddlyield(
                        settlement,
                        maturity,
                        last_interest,
                        rate,
                        pr,
                        redemption,
                        frequency,
                        basis,
                    )))
                }
                "EUROCONVERT" => {
                    let number = self.to_f64_arg(evaluated_args.first(), "EUROCONVERT")?;
                    let source = evaluated_args
                        .get(1)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let target = evaluated_args
                        .get(2)
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let full_precision = evaluated_args
                        .get(3)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(false);
                    let triangulation_precision =
                        evaluated_args.get(4).and_then(|v| self.to_f64(v));
                    res_to_rd(finance::euroconvert(
                        number,
                        &source,
                        &target,
                        full_precision,
                        triangulation_precision,
                    ))
                }
                _ => Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                    "Unknown function: {}",
                    name
                )))),
            };
            match dispatched {
                Ok(ResultData::Float(f)) if !f.is_finite() => {
                    Ok(ResultData::Error("#NUM!".to_string()))
                }
                other => other,
            }
        }
    }

    pub fn get_src(&self, cell: &CellRef) -> Option<&String> {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.src.get(cell.row)
        } else {
            None
        }
    }

    pub fn get_src_str(&self, cell: &CellRef) -> String {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.src.get(cell.row).cloned().unwrap_or("".to_string())
        } else {
            "".to_string()
        }
    }

    pub fn get_src_str_ref(&self, cell: &CellRef) -> Option<&str> {
        let col = self.columns.get(cell.col)?;
        col.src.get(cell.row).map(|s| s.as_str())
    }

    pub fn get_word_boundaries(&self, cell: &CellRef, char_offset: usize) -> (usize, usize) {
        let text = self.get_src_str(cell);
        get_word_boundaries_from_str(&text, char_offset)
    }
}

pub fn get_word_boundaries_from_str(text: &str, char_offset: usize) -> (usize, usize) {
    if text.is_empty() {
        return (0, 0);
    }

    let chars: Vec<char> = text.chars().collect();
    let len = chars.len();
    let offset = char_offset.min(len);

    let is_word_char = |c: char| c.is_alphanumeric() || c == '_';

    let (on_idx, on_c) = if offset < len {
        if chars[offset].is_whitespace() && offset > 0 && is_word_char(chars[offset - 1]) {
            (offset - 1, chars[offset - 1])
        } else {
            (offset, chars[offset])
        }
    } else if offset > 0 {
        (offset - 1, chars[offset - 1])
    } else {
        return (0, 0);
    };

    if on_c.is_whitespace() {
        let mut start = on_idx;
        while start > 0 && chars[start - 1].is_whitespace() {
            start -= 1;
        }
        let mut end = on_idx + 1;
        while end < len && chars[end].is_whitespace() {
            end += 1;
        }
        return (start, end);
    }

    if is_word_char(on_c) {
        let mut start = on_idx;
        while start > 0 && is_word_char(chars[start - 1]) {
            start -= 1;
        }
        let mut end = on_idx + 1;
        while end < len && is_word_char(chars[end]) {
            end += 1;
        }
        return (start, end);
    }

    let mut start = on_idx;
    while start > 0 && !is_word_char(chars[start - 1]) && !chars[start - 1].is_whitespace() {
        start -= 1;
    }
    let mut end = on_idx + 1;
    while end < len && !is_word_char(chars[end]) && !chars[end].is_whitespace() {
        end += 1;
    }
    (start, end)
}
impl Sheet {
    pub fn get_result_data(&self, cell: &CellRef) -> ResultData {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.data.get(cell.row).unwrap_or(ResultData::None)
        } else {
            ResultData::None
        }
    }

    /// Updates the src text of a particular cell but does
    /// not automatically evaluate. Call [`Sheet::commit`] to evaluate
    /// updated cells.
    /// Directly sets the src of a cell and marks it dirty.
    pub fn set_cell_src(&mut self, row: usize, col: usize, src: String) {
        let table_clone = self.clone();
        if let Some(column) = self.columns.get_mut(col)
            && row < column.src.len()
        {
            column.src[row] = src.clone();
            let compiled = crate::core::parser::compile_formula(&src, &[table_clone]);
            column.compiled_src[row] = compiled;
            column.mark_dirty(row);

            self.uncommitted_actions
                .push(crate::core::SheetAction::SetCellSrc {
                    sheet_name: self.name.clone(),
                    col,
                    row,
                    src,
                });
        }
    }

    pub fn insert(&mut self, pos: TextCellRef, input: &str) {
        let TextCellRef {
            row,
            col,
            char_offset,
        } = pos;
        let table_clone = self.clone();
        let existing_col = self.columns.get_mut(col);
        match existing_col {
            Some(existing_column) => {
                existing_column.insert(ColumnPosition { row, char_offset }, input);
                let src = existing_column.src[row].clone();
                let compiled = crate::core::parser::compile_formula(&src, &[table_clone]);
                existing_column.compiled_src[row] = compiled;
                existing_column.mark_dirty(row);
                self.uncommitted_actions
                    .push(crate::core::SheetAction::SetCellSrc {
                        sheet_name: self.name.clone(),
                        col,
                        row,
                        src,
                    });
            }
            None => {
                println!("Warning: column {} does not exist", col)
            }
        }
    }

    /// Delete one before (like backspace)
    pub fn delete_one_before(&mut self, pos: TextCellRef) {
        let char_offset = pos.char_offset;
        let start = if char_offset > 0 {
            TextCellRef {
                row: pos.row,
                col: pos.col,
                char_offset: char_offset - 1,
            }
        } else {
            pos.clone()
        };
        let end = pos;
        self.delete(start, end);
    }

    pub fn delete(&mut self, start: TextCellRef, end: TextCellRef) {
        // Validate positions are in correct order
        if start.col > end.col || (start.col == end.col && start.row > end.row) {
            return;
        }
        let table_clone = self.clone();
        // Handle deletion within a single column
        if start.col == end.col {
            let start_index = start.row;
            let end_index = end.row;

            if let Some(column) = self.columns.get_mut(start.col) {
                // Handle single row deletion
                if start.row == end.row && start_index < column.src.len() {
                    let src = &mut column.src[start_index];
                    let end_offset = std::cmp::min(end.char_offset, src.len());
                    if start.char_offset < end_offset {
                        src.replace_range(start.char_offset..end_offset, "");
                        column.dirty_indices.push(start_index);
                        let updated_src = src.clone();
                        let compiled =
                            crate::core::parser::compile_formula(&updated_src, &[table_clone]);
                        column.compiled_src[start_index] = compiled;
                    }
                }
                // Handle multi-row deletion
                else if start_index < column.src.len() {
                    // Delete complete rows between start and end
                    if end_index >= start_index {
                        column.src.drain(start_index..=end_index);
                        column.compiled_src.drain(start_index..=end_index);
                        column.data.drain(start_index..=end_index);
                    }
                }
            }
        } else {
            // Handle multi-column deletion
            for col in start.col..=end.col {
                if let Some(column) = self.columns.get_mut(col) {
                    let start_index = if col == start.col { col } else { 0 };

                    let end_index = if col == end.col {
                        col
                    } else {
                        column.src.len() - 1
                    };

                    if start_index < column.src.len() {
                        // Delete rows in this column
                        if end_index >= start_index {
                            column.src.drain(start_index..=end_index);
                            column.compiled_src.drain(start_index..=end_index);
                            column.data.drain(start_index..=end_index);
                        }
                    }
                }
            }
        }
    }

    pub fn extend(&mut self, direction: Direction) {
        if self.columns.is_empty() {
            return;
        }
        let row_count = self.columns[0].src.len();
        const MAX_COLS: usize = 26;
        match direction {
            Direction::Up => {
                for column in &mut self.columns {
                    column.src.insert(0, String::new());
                    column
                        .compiled_src
                        .insert(0, crate::core::CompiledFormula::default());
                    column.data.insert(0, ResultData::None);
                }
                self.uncommitted_actions
                    .push(crate::core::SheetAction::InsertRow {
                        sheet_name: self.name.clone(),
                        index: 0,
                    });
            }
            Direction::Down => {
                for column in &mut self.columns {
                    column.src.push(String::new());
                    column
                        .compiled_src
                        .push(crate::core::CompiledFormula::default());
                    column.data.push(ResultData::None);
                }
                self.uncommitted_actions
                    .push(crate::core::SheetAction::InsertRow {
                        sheet_name: self.name.clone(),
                        index: row_count,
                    });
            }
            Direction::Left => {
                if self.columns.len() < MAX_COLS {
                    self.columns.insert(0, DataColumn::new(row_count));
                    self.uncommitted_actions
                        .push(crate::core::SheetAction::InsertCol {
                            sheet_name: self.name.clone(),
                            index: 0,
                        });
                }
            }
            Direction::Right => {
                if self.columns.len() < MAX_COLS {
                    self.columns.push(DataColumn::new(row_count));
                    self.uncommitted_actions
                        .push(crate::core::SheetAction::InsertCol {
                            sheet_name: self.name.clone(),
                            index: self.columns.len() - 1,
                        });
                }
            }
            Direction::None => {}
        }
    }

    /// Insert a new empty row at the specified index
    /// If index is >= row_count, appends at the end
    pub fn insert_row(&mut self, index: usize) {
        let row_count = self.row_count();
        if index >= row_count {
            // Append at the end
            for column in &mut self.columns {
                column.src.push(String::new());
                column
                    .compiled_src
                    .push(crate::core::CompiledFormula::default());
                column.data.push(ResultData::None);
            }
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertRow {
                    sheet_name: self.name.clone(),
                    index: row_count,
                });
        } else {
            // Insert at the specified index
            for column in &mut self.columns {
                column.src.insert(index, String::new());
                column
                    .compiled_src
                    .insert(index, crate::core::CompiledFormula::default());
                column.data.insert(index, ResultData::None);
            }
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertRow {
                    sheet_name: self.name.clone(),
                    index,
                });
        }
    }

    pub fn delete_row(&mut self, index: usize) {
        let row_count = self.row_count();
        if index < row_count {
            for column in &mut self.columns {
                column.src.remove(index);
                column.compiled_src.remove(index);
                column.data.remove(index);
                // Adjust dirty indices
                column.dirty_indices.retain(|&i| i != index);
                for i in 0..column.dirty_indices.len() {
                    if column.dirty_indices[i] > index {
                        column.dirty_indices[i] -= 1;
                    }
                }
            }
            self.uncommitted_actions
                .push(crate::core::SheetAction::DeleteRow {
                    sheet_name: self.name.clone(),
                    index,
                });
            self.mark_all_dirty();
        }
    }

    pub fn delete_col(&mut self, index: usize) {
        if index < self.columns.len() {
            self.columns.remove(index);
            self.uncommitted_actions
                .push(crate::core::SheetAction::DeleteCol {
                    sheet_name: self.name.clone(),
                    index,
                });
            self.mark_all_dirty();
        }
    }

    /// Insert a new empty column at the specified index
    /// If index is >= columns.len(), appends at the end
    pub fn insert_col(&mut self, index: usize) {
        let row_count = self.row_count();
        let new_col = DataColumn::new(row_count);
        let col_count = self.columns.len();
        if index >= col_count {
            self.columns.push(new_col);
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertCol {
                    sheet_name: self.name.clone(),
                    index: col_count,
                });
        } else {
            self.columns.insert(index, new_col);
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertCol {
                    sheet_name: self.name.clone(),
                    index,
                });
        }
        self.mark_all_dirty();
    }

    pub fn row_count(&self) -> usize {
        self.columns.first().map(|c| c.src.len()).unwrap_or(0)
    }

    pub fn col_count(&self) -> usize {
        self.columns.len()
    }
}
