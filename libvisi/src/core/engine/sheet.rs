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
}

impl<'a> Context<'a> {
    /// Create a new empty context
    pub fn new() -> Self {
        Self {
            sheets: HashMap::new(),
        }
    }

    /// Add a sheet to the context for lookup during evaluation
    pub fn add_table(&mut self, name: String, sheet: &'a Sheet) {
        self.sheets.insert(name, sheet);
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
                    let (res, deps) =
                        match self.eval_with_row(&eval_src, context, Some(cell_ref.row)) {
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
            if let Some(col) = self.columns.get_mut(cell_ref.col) {
                if cell_ref.row < col.compiled_src.len() {
                    col.compiled_src[cell_ref.row] = compiled_to_cache.unwrap_or_default();
                }
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
            if let Some(col) = self.columns.get_mut(cell_ref.col) {
                if cell_ref.row < col.data.len() {
                    col.data.set(cell_ref.row, result.clone());
                    updated_cells.insert(cell_ref);
                }
            }
            if let Some(comp_sheet) = tables_for_compilation
                .iter_mut()
                .find(|s| s.name == self.name)
            {
                if let Some(col) = comp_sheet.columns.get_mut(cell_ref.col) {
                    if cell_ref.row < col.data.len() {
                        col.data.set(cell_ref.row, result);
                    }
                }
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
                if let Some(col) = self.columns.get_mut(dependent.col) {
                    if !col.dirty_indices.contains(&dependent.row) {
                        col.dirty_indices.push(dependent.row);
                    }
                }
            }
        }
    }

    pub fn eval_with_row(
        &self,
        input: &str,
        context: Option<&Context>,
        row: Option<usize>,
    ) -> Result<(ResultData, Vec<Dependency>), EngineError> {
        if input.is_empty() {
            return Ok((ResultData::None, vec![]));
        }
        if input.starts_with('=') {
            self.eval_excel(&input[1..], context, row)
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
        self.eval_with_row(input, context, None)
    }

    fn eval_excel(
        &self,
        code: &str,
        context: Option<&Context>,
        row: Option<usize>,
    ) -> Result<(ResultData, Vec<Dependency>), EngineError> {
        let ast = crate::core::parser::parse_excel_formula(code)
            .map_err(|e| EngineError::EvalError(EvalError::UnknownFunction(e)))?;

        let mut deps = Vec::new();
        let result = match self.evaluate_ast(&ast, context, row, &mut deps) {
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
        deps: &mut Vec<Dependency>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::SheetSection;
        use crate::core::parser::Expr;
        use crate::core::parser::Op;

        match ast {
            Expr::Number(n) => Ok(ResultData::Float(*n)),
            Expr::String(s) => Ok(ResultData::String(s.clone())),
            Expr::Boolean(b) => Ok(ResultData::Boolean(*b)),
            Expr::Identifier(_) => Ok(ResultData::None),
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
                if found.is_none() {
                    if let Some(ctx) = context {
                        for s in ctx.sheets.values() {
                            if let Some(t) = s.find_table(&ref_name) {
                                found = Some((s, t));
                                break;
                            }
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

                let mut results = Vec::new();
                for r in *start_row..=actual_end_row {
                    for c in *start_col..=*end_col {
                        let cell_ref = CellRef::new(r, c);
                        if is_self {
                            if is_col_range {
                                let col_dep = Dependency::LocalColumn(c);
                                if !deps.contains(&col_dep) {
                                    deps.push(col_dep);
                                }
                            } else {
                                deps.push(Dependency::Local(cell_ref));
                            }
                            results.push(self.get_result_data(&cell_ref));
                        } else {
                            let name = sheet.as_ref().unwrap().clone();
                            if is_col_range {
                                let col_dep = Dependency::RemoteColumn {
                                    sheet: name.clone(),
                                    col: c,
                                };
                                if !deps.contains(&col_dep) {
                                    deps.push(col_dep);
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
                    results.push(self.evaluate_ast(item, context, row, deps)?);
                }
                Ok(ResultData::List(results))
            }
            Expr::UnaryOp { op, expr } => {
                let val = self.evaluate_ast(expr, context, row, deps)?;
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
                let l_val = self.evaluate_ast(left, context, row, deps)?;
                let r_val = self.evaluate_ast(right, context, row, deps)?;

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
                self.evaluate_function(name, args, context, row, deps)
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
                return "".cmp(&b.to_lowercase().as_str());
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
                ResultData::String(_) => {
                    if is_direct.get(i).copied().unwrap_or(false) {
                        if self.to_f64(arg).is_none() {
                            return Some(ResultData::Error("#VALUE!".to_string()));
                        }
                    }
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

    fn flatten_stat_numbers_a(&self, arg: &ResultData) -> Vec<f64> {
        match arg {
            ResultData::Float(f) => vec![*f],
            ResultData::Integer(i) => vec![*i as f64],
            ResultData::Boolean(b) => vec![if *b { 1.0 } else { 0.0 }],
            ResultData::String(_) => vec![0.0],
            ResultData::List(list) => list
                .iter()
                .flat_map(|v| self.flatten_stat_numbers_a(v))
                .collect(),
            ResultData::None => vec![],
            _ => vec![0.0],
        }
    }

    fn extract_matrix(&self, arg: &ResultData) -> Vec<Vec<f64>> {
        match arg {
            ResultData::List(list) => {
                let mut rows = Vec::new();
                for item in list {
                    match item {
                        ResultData::List(sub_list) => {
                            let row: Vec<f64> = sub_list.iter().flat_map(|v| self.to_f64(v)).collect();
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
                        prod = Self::clean_float(prod * p);
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

    fn match_criteria(&self, val: &ResultData, criteria: &ResultData) -> bool {
        let crit_str = criteria.to_string();
        if crit_str.starts_with('>') {
            let val_f = self.to_f64(val).unwrap_or(0.0);
            if crit_str.starts_with(">=") {
                let crit_f = crit_str[2..].trim().parse::<f64>().unwrap_or(0.0);
                val_f >= crit_f
            } else {
                let crit_f = crit_str[1..].trim().parse::<f64>().unwrap_or(0.0);
                val_f > crit_f
            }
        } else if crit_str.starts_with('<') {
            let val_f = self.to_f64(val).unwrap_or(0.0);
            if crit_str.starts_with("<=") {
                let crit_f = crit_str[2..].trim().parse::<f64>().unwrap_or(0.0);
                val_f <= crit_f
            } else if crit_str.starts_with("<>") {
                let remainder = crit_str[2..].trim().to_string();
                val.to_string() != remainder
            } else {
                let crit_f = crit_str[1..].trim().parse::<f64>().unwrap_or(0.0);
                val_f < crit_f
            }
        } else if crit_str.starts_with('=') {
            let remainder = crit_str[1..].trim().to_string();
            val.to_string() == remainder
        } else {
            val.to_string() == crit_str
        }
    }

    fn proper(&self, s: &str) -> String {
        let mut c_chars = Vec::new();
        let mut capitalize_next = true;
        for c in s.chars() {
            if c.is_alphanumeric() {
                if capitalize_next {
                    c_chars.extend(c.to_uppercase());
                    capitalize_next = false;
                } else {
                    c_chars.extend(c.to_lowercase());
                }
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

    fn evaluate_function(
        &self,
        name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        deps: &mut Vec<Dependency>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::parser::Expr;
        let mut upper_name = name.to_uppercase();
        if upper_name.starts_with("_XLFN.") {
            upper_name = upper_name["_XLFN.".len()..].to_string();
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
                {
                    if let Expr::Identifier(name) = &**left {
                        let val = self.evaluate_ast(right, context, row, deps)?;
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
                }

                let val = self.evaluate_ast(arg, context, row, deps)?;
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
                    4 => {
                        if val.to_string() == "line" {
                            is_line = true;
                        }
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
                let cond_val = self.evaluate_ast(&args[0], context, row, deps)?;
                if let ResultData::Error(_) = cond_val {
                    return Ok(cond_val);
                }
                let condition = match self.to_bool_opt(&cond_val) {
                    Some(b) => b,
                    None => return Ok(ResultData::Error("#VALUE!".to_string())),
                };
                if condition {
                    return self.evaluate_ast(&args[1], context, row, deps);
                } else {
                    return self.evaluate_ast(&args[2], context, row, deps);
                }
            }

            if upper_name == "IFERROR" {
                if args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "IFERROR requires 2 arguments".to_string(),
                    )));
                }
                let first_res = self.evaluate_ast(&args[0], context, row, deps);
                match first_res {
                    Ok(ResultData::Error(_)) | Err(_) => {
                        return self.evaluate_ast(&args[1], context, row, deps);
                    }
                    Ok(val) => return Ok(val),
                }
            }

            if upper_name == "ISERROR" {
                if args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                let res = self.evaluate_ast(&args[0], context, row, deps);
                return match res {
                    Ok(ResultData::Error(_)) | Err(_) => Ok(ResultData::Boolean(true)),
                    _ => Ok(ResultData::Boolean(false)),
                };
            }

            if upper_name == "ISNA" {
                if args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                let res = self.evaluate_ast(&args[0], context, row, deps);
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
                let eval_res = match self.evaluate_ast(arg, context, row, deps) {
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
            if upper_name != "IFERROR"
                && upper_name != "ISERROR"
                && upper_name != "ISNA"
                && !uses_ordered_arg_error_check
            {
                if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                    return Ok(err);
                }
            }

            let res_to_rd = |res: Result<f64, String>| -> Result<ResultData, EngineError> {
                match res {
                    Ok(v) => Ok(ResultData::Float(v)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            };

            match upper_name.as_str() {
                // --- STATISTICAL FUNCTIONS ---
                "AVEDEV" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::avedev(&nums))
                }
                "AVERAGEA" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .flat_map(|arg| self.flatten_stat_numbers_a(arg))
                        .collect();
                    if nums.is_empty() {
                        Ok(ResultData::Error("#DIV/0!".to_string()))
                    } else {
                        Ok(ResultData::Float(nums.iter().sum::<f64>() / nums.len() as f64))
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
                        if self.match_criteria(val, criteria) {
                            if let Some(target_val) = avg_range.get(i) {
                                if let Some(f) = self.to_f64(target_val) {
                                    sum += f;
                                    count += 1;
                                }
                            }
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
                            if idx >= crit_range.len() || !self.match_criteria(&crit_range[idx], crit_val) {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match {
                            if let Some(f) = self.to_f64(target_val) {
                                sum += f;
                                count += 1;
                            }
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
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(true);
                    let a = evaluated_args.get(4).and_then(|v| self.to_f64(v)).unwrap_or(0.0);
                    let b = evaluated_args.get(5).and_then(|v| self.to_f64(v)).unwrap_or(1.0);
                    res_to_rd(crate::core::stats::beta_dist(x, alpha, beta, cumulative, a, b))
                }
                "BETA.INV" | "BETAINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "BETA.INV")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "BETA.INV")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "BETA.INV")?;
                    let a = evaluated_args.get(3).and_then(|v| self.to_f64(v)).unwrap_or(0.0);
                    let b = evaluated_args.get(4).and_then(|v| self.to_f64(v)).unwrap_or(1.0);
                    res_to_rd(crate::core::stats::beta_inv(p, alpha, beta, a, b))
                }
                "BINOM.DIST" | "BINOMDIST" => {
                    let k = self.to_f64_arg(evaluated_args.first(), "BINOM.DIST")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "BINOM.DIST")?;
                    let p = self.to_f64_arg(evaluated_args.get(2), "BINOM.DIST")?;
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(false);
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
                    let cumulative = evaluated_args.get(2).map(|v| self.to_bool(v)).unwrap_or(true);
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
                    let actual: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let expected: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::chisq_test(&actual, &expected))
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
                    let xs: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
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
                    let xs: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::covariance_p(&xs, &ys))
                }
                "COVARIANCE.S" => {
                    let xs: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::covariance_s(&xs, &ys))
                }
                "DEVSQ" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::devsq(&nums))
                }
                "EXPON.DIST" | "EXPONDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "EXPON.DIST")?;
                    let lambda = self.to_f64_arg(evaluated_args.get(1), "EXPON.DIST")?;
                    let cumulative = evaluated_args.get(2).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::expon_dist(x, lambda, cumulative))
                }
                "F.DIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "F.DIST")?;
                    let df1 = self.to_f64_arg(evaluated_args.get(1), "F.DIST")?;
                    let df2 = self.to_f64_arg(evaluated_args.get(2), "F.DIST")?;
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(true);
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
                    let array1: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let array2: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
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
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(2).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::forecast_linear(x, &ys, &xs))
                }
                "FORECAST.ETS" | "FORECAST.ETS.CONFINT" | "FORECAST.ETS.SEASONALITY" | "FORECAST.ETS.STAT" => {
                    let target_date = evaluated_args.first().and_then(|v| self.to_f64(v)).unwrap_or(0.0);
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(2).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    match upper_name.as_str() {
                        "FORECAST.ETS.SEASONALITY" => Ok(ResultData::Float(1.0)),
                        "FORECAST.ETS.STAT" => Ok(ResultData::Float(0.5)),
                        "FORECAST.ETS.CONFINT" => Ok(ResultData::Float(0.0)),
                        _ => res_to_rd(crate::core::stats::forecast_linear(target_date, &ys, &xs)),
                    }
                }
                "FREQUENCY" => {
                    let data: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let bins: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    match crate::core::stats::frequency(&data, &bins) {
                        Ok(counts) => Ok(ResultData::List(counts.into_iter().map(ResultData::Float).collect())),
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
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::gamma_dist(x, alpha, beta, cumulative))
                }
                "GAMMA.INV" | "GAMMAINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "GAMMA.INV")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "GAMMA.INV")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "GAMMA.INV")?;
                    res_to_rd(crate::core::stats::gamma_inv(p, alpha, beta))
                }
                "GAMMALN" | "GAMMALN.PRECISE" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "GAMMALN")?;
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
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::geomean(&nums))
                }
                "GROWTH" | "LOGEST" => {
                    let ys: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_else(|| (1..=ys.len()).map(|i| i as f64).collect());
                    let ln_ys: Vec<f64> = ys.iter().map(|y| y.ln()).collect();
                    let m = crate::core::stats::slope(&ln_ys, &xs).unwrap_or(0.0);
                    let b = crate::core::stats::intercept(&ln_ys, &xs).unwrap_or(0.0);
                    if upper_name == "LOGEST" {
                        Ok(ResultData::List(vec![ResultData::Float(m.exp()), ResultData::Float(b.exp())]))
                    } else {
                        let new_x = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or_else(|| xs.first().copied().unwrap_or(1.0));
                        Ok(ResultData::Float((b + m * new_x).exp()))
                    }
                }
                "HARMEAN" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::harmean(&nums))
                }
                "HYPGEOM.DIST" | "HYPGEOMDIST" => {
                    let sample_s = self.to_f64_arg(evaluated_args.first(), "HYPGEOM.DIST")?;
                    let sample_size = self.to_f64_arg(evaluated_args.get(1), "HYPGEOM.DIST")?;
                    let pop_s = self.to_f64_arg(evaluated_args.get(2), "HYPGEOM.DIST")?;
                    let pop_size = self.to_f64_arg(evaluated_args.get(3), "HYPGEOM.DIST")?;
                    let cumulative = evaluated_args.get(4).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::hypgeom_dist(sample_s, sample_size, pop_s, pop_size, cumulative))
                }
                "INTERCEPT" => {
                    let ys: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::intercept(&ys, &xs))
                }
                "KURT" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::kurt(&nums))
                }
                "LARGE" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "LARGE")?.round() as usize;
                    res_to_rd(crate::core::stats::large(&nums, k))
                }
                "LINEST" | "TREND" => {
                    let ys: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_else(|| (1..=ys.len()).map(|i| i as f64).collect());
                    let m = crate::core::stats::slope(&ys, &xs).unwrap_or(0.0);
                    let b = crate::core::stats::intercept(&ys, &xs).unwrap_or(0.0);
                    if upper_name == "LINEST" {
                        Ok(ResultData::List(vec![ResultData::Float(m), ResultData::Float(b)]))
                    } else {
                        let new_x = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or_else(|| xs.first().copied().unwrap_or(1.0));
                        Ok(ResultData::Float(m * new_x + b))
                    }
                }
                "LOGNORM.DIST" | "LOGNORMDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "LOGNORM.DIST")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "LOGNORM.DIST")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "LOGNORM.DIST")?;
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::lognorm_dist(x, mean, std_dev, cumulative))
                }
                "LOGNORM.INV" | "LOGINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "LOGNORM.INV")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "LOGNORM.INV")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "LOGNORM.INV")?;
                    res_to_rd(crate::core::stats::lognorm_inv(p, mean, std_dev))
                }
                "MAXA" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers_a(arg)).collect();
                    if nums.is_empty() {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max)))
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
                            if idx >= crit_range.len() || !self.match_criteria(&crit_range[idx], crit_val) {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match {
                            if let Some(f) = self.to_f64(target_val) {
                                max_val = max_val.max(f);
                                found = true;
                            }
                        }
                    }
                    if !found {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(max_val))
                    }
                }
                "MEDIAN" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::median(&nums))
                }
                "MINA" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers_a(arg)).collect();
                    if nums.is_empty() {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(nums.iter().cloned().fold(f64::INFINITY, f64::min)))
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
                            if idx >= crit_range.len() || !self.match_criteria(&crit_range[idx], crit_val) {
                                all_match = false;
                                break;
                            }
                        }
                        if all_match {
                            if let Some(f) = self.to_f64(target_val) {
                                min_val = min_val.min(f);
                                found = true;
                            }
                        }
                    }
                    if !found {
                        Ok(ResultData::Float(0.0))
                    } else {
                        Ok(ResultData::Float(min_val))
                    }
                }
                "MODE.MULT" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    match crate::core::stats::mode_mult(&nums) {
                        Ok(modes) => Ok(ResultData::List(modes.into_iter().map(ResultData::Float).collect())),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "MODE.SNGL" | "MODE" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::mode_sngl(&nums))
                }
                "NEGBINOM.DIST" | "NEGBINOMDIST" => {
                    let k = self.to_f64_arg(evaluated_args.first(), "NEGBINOM.DIST")?;
                    let r = self.to_f64_arg(evaluated_args.get(1), "NEGBINOM.DIST")?;
                    let p = self.to_f64_arg(evaluated_args.get(2), "NEGBINOM.DIST")?;
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(false);
                    res_to_rd(crate::core::stats::negbinom_dist(k, r, p, cumulative))
                }
                "NORM.DIST" | "NORMDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "NORM.DIST")?;
                    let mean = self.to_f64_arg(evaluated_args.get(1), "NORM.DIST")?;
                    let std_dev = self.to_f64_arg(evaluated_args.get(2), "NORM.DIST")?;
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(true);
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
                    let cumulative = evaluated_args.get(1).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::norm_s_dist(z, cumulative))
                }
                "NORM.S.INV" | "NORMSINV" => {
                    let p = self.to_f64_arg(evaluated_args.first(), "NORM.S.INV")?;
                    res_to_rd(crate::core::stats::norm_s_inv(p))
                }
                "PERCENTILE.EXC" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "PERCENTILE.EXC")?;
                    res_to_rd(crate::core::stats::percentile_exc(&nums, k))
                }
                "PERCENTILE.INC" | "PERCENTILE" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let k = self.to_f64_arg(evaluated_args.get(1), "PERCENTILE.INC")?;
                    res_to_rd(crate::core::stats::percentile_inc(&nums, k))
                }
                "PERCENTRANK.EXC" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let x = self.to_f64_arg(evaluated_args.get(1), "PERCENTRANK.EXC")?;
                    let sig = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or(3.0) as usize;
                    res_to_rd(crate::core::stats::percentrank_exc(&nums, x, sig))
                }
                "PERCENTRANK.INC" | "PERCENTRANK" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let x = self.to_f64_arg(evaluated_args.get(1), "PERCENTRANK.INC")?;
                    let sig = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or(3.0) as usize;
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
                    let cumulative = evaluated_args.get(2).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::poisson_dist(x, mean, cumulative))
                }
                "PROB" => {
                    let x_range: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let prob_range: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let lower = self.to_f64_arg(evaluated_args.get(2), "PROB")?;
                    let upper = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::stats::prob(&x_range, &prob_range, lower, upper))
                }
                "QUARTILE.EXC" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let q = self.to_f64_arg(evaluated_args.get(1), "QUARTILE.EXC")?.round() as usize;
                    res_to_rd(crate::core::stats::quartile_exc(&nums, q))
                }
                "QUARTILE.INC" | "QUARTILE" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let q = self.to_f64_arg(evaluated_args.get(1), "QUARTILE.INC")?.round() as usize;
                    res_to_rd(crate::core::stats::quartile_inc(&nums, q))
                }
                "RANK.AVG" => {
                    let number = self.to_f64_arg(evaluated_args.first(), "RANK.AVG")?;
                    let ref_data: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let order = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or(0.0) as usize;
                    res_to_rd(crate::core::stats::rank_avg(number, &ref_data, order))
                }
                "RANK.EQ" | "RANK" => {
                    let number = self.to_f64_arg(evaluated_args.first(), "RANK.EQ")?;
                    let ref_data: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let order = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or(0.0) as usize;
                    res_to_rd(crate::core::stats::rank_eq(number, &ref_data, order))
                }
                "RSQ" => {
                    let ys: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::rsq(&ys, &xs))
                }
                "SKEW" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::skew(&nums))
                }
                "SKEW.P" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::skew_p(&nums))
                }
                "SLOPE" => {
                    let ys: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::slope(&ys, &xs))
                }
                "SMALL" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
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
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::stdev_p(&nums))
                }
                "STDEV.S" | "STDEV" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::stdev_s(&nums))
                }
                "STDEVA" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers_a(arg)).collect();
                    res_to_rd(crate::core::stats::stdev_s(&nums))
                }
                "STDEVPA" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers_a(arg)).collect();
                    res_to_rd(crate::core::stats::stdev_p(&nums))
                }
                "STDEYX" => {
                    let ys: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let xs: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::stats::steyx(&ys, &xs))
                }
                "T.DIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "T.DIST")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST")?;
                    let cumulative = evaluated_args.get(2).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::t_dist(x, df, cumulative))
                }
                "T.DIST.2T" | "TDIST" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "T.DIST.2T")?;
                    let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST.2T")?;
                    res_to_rd(crate::core::stats::t_dist_2t(x, df))
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
                    let array1: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let array2: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let tails = evaluated_args.get(2).and_then(|v| self.to_f64(v)).unwrap_or(2.0) as usize;
                    let test_type = evaluated_args.get(3).and_then(|v| self.to_f64(v)).unwrap_or(1.0) as usize;
                    res_to_rd(crate::core::stats::t_test(&array1, &array2, tails, test_type))
                }
                "TRIMMEAN" => {
                    let nums: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let percent = self.to_f64_arg(evaluated_args.get(1), "TRIMMEAN")?;
                    res_to_rd(crate::core::stats::trimmean(&nums, percent))
                }
                "VAR.P" | "VARP" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::var_p(&nums))
                }
                "VAR.S" | "VAR" => {
                    let nums: Vec<f64> = evaluated_args
                        .iter()
                        .enumerate()
                        .flat_map(|(i, arg)| self.flatten_stat_numbers(arg, arg_is_direct[i]))
                        .collect();
                    res_to_rd(crate::core::stats::var_s(&nums))
                }
                "VARA" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers_a(arg)).collect();
                    res_to_rd(crate::core::stats::var_s(&nums))
                }
                "VARPA" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers_a(arg)).collect();
                    res_to_rd(crate::core::stats::var_p(&nums))
                }
                "WEIBULL.DIST" | "WEIBULL" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "WEIBULL.DIST")?;
                    let alpha = self.to_f64_arg(evaluated_args.get(1), "WEIBULL.DIST")?;
                    let beta = self.to_f64_arg(evaluated_args.get(2), "WEIBULL.DIST")?;
                    let cumulative = evaluated_args.get(3).map(|v| self.to_bool(v)).unwrap_or(true);
                    res_to_rd(crate::core::stats::weibull_dist(x, alpha, beta, cumulative))
                }
                "Z.TEST" | "ZTEST" => {
                    let array: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
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
                "AGGREGATE" | "SUBTOTAL" => {
                    let fn_num = self.to_f64_arg(evaluated_args.first(), "AGGREGATE")?.round() as usize;
                    let nums: Vec<f64> = evaluated_args.iter().skip(1).flat_map(|arg| self.flatten_stat_numbers(arg, false)).collect();
                    match fn_num % 100 {
                        1 => res_to_rd(if nums.is_empty() { Err("#DIV/0!".to_string()) } else { Ok(nums.iter().sum::<f64>() / nums.len() as f64) }),
                        2 | 3 => Ok(ResultData::Float(nums.len() as f64)),
                        4 => Ok(ResultData::Float(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))),
                        5 => Ok(ResultData::Float(nums.iter().cloned().fold(f64::INFINITY, f64::min))),
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
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
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
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
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
                    let n = self.to_f64_arg(evaluated_args.first(), "FACTDOUBLE")?;
                    res_to_rd(crate::core::math_trig::factdouble(n))
                }
                "FLOOR.MATH" | "FLOOR.PRECISE" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "FLOOR.MATH")?;
                    let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    let mode = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::math_trig::floor_math(x, sig, mode))
                }
                "GCD" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers(arg, false)).collect();
                    res_to_rd(crate::core::math_trig::gcd(&nums))
                }
                "LCM" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers(arg, false)).collect();
                    res_to_rd(crate::core::math_trig::lcm(&nums))
                }
                "LET" => {
                    if let Some(last) = evaluated_args.last() {
                        Ok(last.clone())
                    } else {
                        Ok(ResultData::None)
                    }
                }
                "LOG" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "LOG")?;
                    let base = evaluated_args.get(1).and_then(|v| self.to_f64(v)).unwrap_or(10.0);
                    if num <= 0.0 || base <= 0.0 || base == 1.0 {
                        Ok(ResultData::Error("#NUM!".to_string()))
                    } else {
                        Ok(ResultData::Float(num.log(base)))
                    }
                }
                "MDETERM" => {
                    let matrix = evaluated_args.first().map(|arg| self.extract_matrix(arg)).unwrap_or_default();
                    res_to_rd(crate::core::math_trig::mdeterm(&matrix))
                }
                "MINVERSE" => {
                    let matrix = evaluated_args.first().map(|arg| self.extract_matrix(arg)).unwrap_or_default();
                    match crate::core::math_trig::minverse(&matrix) {
                        Ok(inv) => Ok(ResultData::List(inv.into_iter().map(|row| ResultData::List(row.into_iter().map(ResultData::Float).collect())).collect())),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "MROUND" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "MROUND")?;
                    let mult = self.to_f64_arg(evaluated_args.get(1), "MROUND")?;
                    res_to_rd(crate::core::math_trig::mround(x, mult))
                }
                "MULTINOMIAL" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers(arg, false)).collect();
                    res_to_rd(crate::core::math_trig::multinomial(&nums))
                }
                "MUNIT" => {
                    let dim = self.to_f64_arg(evaluated_args.first(), "MUNIT")?;
                    match crate::core::math_trig::munit(dim) {
                        Ok(mat) => Ok(ResultData::List(mat.into_iter().map(|row| ResultData::List(row.into_iter().map(ResultData::Float).collect())).collect())),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "ODD" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "ODD")?;
                    res_to_rd(crate::core::math_trig::odd(x))
                }
                "PERCENTOF" => {
                    let data_val = self.to_f64_arg(evaluated_args.first(), "PERCENTOF")?;
                    let target_val = self.to_f64_arg(evaluated_args.get(1), "PERCENTOF")?;
                    res_to_rd(crate::core::math_trig::percentof(data_val, target_val))
                }
                "PI" => Ok(ResultData::Float(std::f64::consts::PI)),
                "POWER" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "POWER")?;
                    let p = self.to_f64_arg(evaluated_args.get(1), "POWER")?;
                    res_to_rd(crate::core::math_trig::power(num, p))
                }
                "QUOTIENT" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "QUOTIENT")?;
                    let den = self.to_f64_arg(evaluated_args.get(1), "QUOTIENT")?;
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
                        Ok(grid) => Ok(ResultData::List(grid.into_iter().map(|row| ResultData::List(row.into_iter().map(ResultData::Float).collect())).collect())),
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
                        Ok(grid) => Ok(ResultData::List(grid.into_iter().map(|row| ResultData::List(row.into_iter().map(ResultData::Float).collect())).collect())),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "SERIESSUM" => {
                    let x = self.to_f64_arg(evaluated_args.first(), "SERIESSUM")?;
                    let n = self.to_f64_arg(evaluated_args.get(1), "SERIESSUM")?;
                    let m = self.to_f64_arg(evaluated_args.get(2), "SERIESSUM")?;
                    let coeffs: Vec<f64> = evaluated_args.get(3).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
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
                    let x = self.to_f64_arg(evaluated_args.first(), "SQRTPI")?;
                    res_to_rd(crate::core::math_trig::sqrtpi(x))
                }
                "SUMPRODUCT" => {
                    let arrays: Vec<Vec<f64>> = evaluated_args.iter().map(|arg| self.flatten_stat_numbers(arg, false)).collect();
                    res_to_rd(crate::core::math_trig::sumproduct(&arrays))
                }
                "SUMSQ" => {
                    let nums: Vec<f64> = evaluated_args.iter().flat_map(|arg| self.flatten_stat_numbers(arg, false)).collect();
                    res_to_rd(crate::core::math_trig::sumsq(&nums))
                }
                "SUMX2MY2" => {
                    let xs: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::math_trig::sumx2my2(&xs, &ys))
                }
                "SUMX2PY2" => {
                    let xs: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    res_to_rd(crate::core::math_trig::sumx2py2(&xs, &ys))
                }
                "SUMXMY2" => {
                    let xs: Vec<f64> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
                    let ys: Vec<f64> = evaluated_args.get(1).map(|arg| self.flatten_stat_numbers(arg, false)).unwrap_or_default();
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
                    let items: Vec<String> = evaluated_args.first().map(|arg| self.flatten_stat_numbers(arg, false).iter().map(|v| v.to_string()).collect()).unwrap_or_default();
                    let fmt = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::text::arraytotext(&items, fmt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "ASC" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::asc(&text) {
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
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::clean(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "CODE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    res_to_rd(crate::core::text::code(&text))
                }
                "DBCS" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::dbcs(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "DETECTLANGUAGE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
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
                    let t1 = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let t2 = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::exact(&t1, &t2) {
                        Ok(b) => Ok(ResultData::Boolean(b)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "FIND" | "FINDB" => {
                    let find_text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let within_text = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
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
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let dec = evaluated_args.get(1).map(|v| v.to_string());
                    let grp = evaluated_args.get(2).map(|v| v.to_string());
                    res_to_rd(crate::core::text::numbervalue(&text, dec.as_deref(), grp.as_deref()))
                }
                "PHONETIC" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::phonetic(&text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REGEXEXTRACT" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let pat = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::regexextract(&text, &pat) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REGEXREPLACE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let pat = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    let rep = evaluated_args.get(2).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::regexreplace(&text, &pat, &rep) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REGEXTEST" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let pat = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::regextest(&text, &pat) {
                        Ok(b) => Ok(ResultData::Boolean(b)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REPLACE" | "REPLACEB" => {
                    let old_text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let start_num = self.to_f64_arg(evaluated_args.get(1), "REPLACE")?;
                    let num_chars = self.to_f64_arg(evaluated_args.get(2), "REPLACE")?;
                    let new_text = evaluated_args.get(3).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::replace_fn(&old_text, start_num, num_chars, &new_text) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "REPT" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let cnt = self.to_f64_arg(evaluated_args.get(1), "REPT")?;
                    match crate::core::text::rept(&text, cnt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "SEARCH" | "SEARCHB" => {
                    let find_text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let within_text = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    let start_num = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    res_to_rd(crate::core::text::search(&find_text, &within_text, start_num))
                }
                "SUBSTITUTE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let old_text = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    let new_text = evaluated_args.get(2).map(|v| v.to_string()).unwrap_or_default();
                    let instance = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                    match crate::core::text::substitute(&text, &old_text, &new_text, instance) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "T" => {
                    let is_str = matches!(evaluated_args.first(), Some(ResultData::String(_)));
                    let val = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    Ok(ResultData::String(crate::core::text::t_fn(&val, is_str)))
                }
                "TEXT" => {
                    let num = self.to_f64_arg(evaluated_args.first(), "TEXT")?;
                    let fmt = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::text_fn(num, &fmt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTAFTER" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let delim = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    let instance = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    match crate::core::text::textafter(&text, &delim, instance) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTBEFORE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let delim = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    let instance = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                    match crate::core::text::textbefore(&text, &delim, instance) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTJOIN" => {
                    let delim = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let ignore = evaluated_args.get(1).map(|v| self.to_bool(v)).unwrap_or(true);
                    let texts: Vec<String> = evaluated_args.iter().skip(2).map(|v| v.to_string()).collect();
                    match crate::core::text::textjoin(&delim, ignore, &texts) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TEXTSPLIT" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let delim = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    match crate::core::text::textsplit(&text, &delim) {
                        Ok(parts) => Ok(ResultData::List(parts.into_iter().map(ResultData::String).collect())),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "TRANSLATE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let from = evaluated_args.get(1).map(|v| v.to_string()).unwrap_or_default();
                    let to = evaluated_args.get(2).map(|v| v.to_string()).unwrap_or_default();
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
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    res_to_rd(crate::core::text::unicode(&text))
                }
                "VALUE" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    res_to_rd(crate::core::text::value(&text))
                }
                "VALUETOTEXT" => {
                    let val = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let fmt = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                    match crate::core::text::valuetotext(&val, fmt) {
                        Ok(s) => Ok(ResultData::String(s)),
                        Err(e) => Ok(ResultData::Error(e)),
                    }
                }
                "LEFTB" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let count = evaluated_args.get(1).and_then(|v| self.to_f64(v)).unwrap_or(1.0).floor() as usize;
                    let res: String = text.chars().take(count).collect();
                    Ok(ResultData::String(res))
                }
                "RIGHTB" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    let count = evaluated_args.get(1).and_then(|v| self.to_f64(v)).unwrap_or(1.0).floor() as usize;
                    let chars: Vec<char> = text.chars().collect();
                    let skip = chars.len().saturating_sub(count);
                    let res: String = chars.into_iter().skip(skip).collect();
                    Ok(ResultData::String(res))
                }
                "LENB" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
                    Ok(ResultData::Float(text.len() as f64))
                }
                "MIDB" => {
                    let text = evaluated_args.first().map(|v| v.to_string()).unwrap_or_default();
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
                    let mut count = 0;
                    for arg in evaluated_args {
                        count += self.count_helper(&arg);
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
                    Ok(ResultData::Float(val.floor()))
                }
                "CEILING" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "CEILING")?;
                    Ok(ResultData::Float(val.ceil()))
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
                        if let ResultData::List(list) = &evaluated_args[0] {
                            let idx = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as isize;
                            let len = list.len() as isize;
                            let real_idx = if idx < 0 { len + idx } else { idx };
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

                        if let ResultData::List(list) = &evaluated_args[0] {
                            let num_cols = match &args[0] {
                                Expr::RangeRef {
                                    start_col, end_col, ..
                                } => (end_col - start_col + 1) as isize,
                                _ => 1,
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
                        Ok(ResultData::Float(prod))
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
                    let n = self.to_f64(&evaluated_args[0]).unwrap_or(0.0);
                    let d = self.to_f64(&evaluated_args[1]).unwrap_or(1.0);
                    if d == 0.0 {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "MOD divisor cannot be zero".to_string(),
                        )));
                    }
                    let val = n - d * (n / d).floor();
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
                                if item.to_string() == lookup_val.to_string() {
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
                            } => (end_col - start_col + 1) as usize,
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
                                if first_col_val.to_string() == lookup_val.to_string() {
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
                        if idx < sum_list.len() {
                            if self.match_criteria(&range_list[idx], criteria) {
                                sum += self.to_f64(&sum_list[idx]).unwrap_or(0.0);
                            }
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
                            sum += self.to_f64(&sum_list[idx]).unwrap_or(0.0);
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
                        if self.match_criteria(&val, criteria) {
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

                        let mut result_list = Vec::with_capacity(rows1 * cols2);
                        for r in 0..rows1 {
                            for c in 0..cols2 {
                                let mut val = 0.0;
                                for k in 0..cols1 {
                                    let v1 = self.to_f64(&list1[r * cols1 + k]).unwrap_or(0.0);
                                    let v2 = self.to_f64(&list2[k * cols2 + c]).unwrap_or(0.0);
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
                _ => Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                    "Unknown function: {}",
                    name
                )))),
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
        if let Some(column) = self.columns.get_mut(col) {
            if row < column.src.len() {
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
