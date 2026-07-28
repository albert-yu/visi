use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use web_time::Instant;

use super::cell::{CellRef, Dependency, EngineError, EvalError, TextCellRef, generate_unique_id};
use super::column::{ColumnPosition, DataColumn};
use super::result_data::ResultData;
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
        let result = self.evaluate_ast(&ast, context, row, &mut deps)?;
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
                let sheet_name = match sheet {
                    Some(name) => name.clone(),
                    None => self.name.clone(),
                };
                let is_self = sheet_name == self.name;

                let target_table = if is_self {
                    self
                } else if let Some(ctx) = context {
                    if let Some(t) = ctx.sheets.get(&sheet_name) {
                        t
                    } else {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                            "Sheet not found: {}",
                            sheet_name
                        ))));
                    }
                } else {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                        "No context to resolve sheet reference: {}",
                        sheet_name
                    ))));
                };

                let col_idx = if let Some(col_name) = column {
                    if let Some(pos) = target_table
                        .columns
                        .iter()
                        .position(|c| c.name == *col_name)
                    {
                        pos
                    } else {
                        return Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                            "Column not found: {}",
                            col_name
                        ))));
                    }
                } else {
                    0
                };

                match section {
                    SheetSection::Headers => {
                        if let Some(col) = target_table.columns.get(col_idx) {
                            Ok(ResultData::String(col.name.clone()))
                        } else {
                            Ok(ResultData::None)
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
                            let cell_ref = CellRef::new(r, col_idx);
                            if is_self {
                                deps.push(Dependency::Local(cell_ref));
                            } else {
                                deps.push(Dependency::Remote {
                                    sheet: sheet_name,
                                    cell: cell_ref,
                                });
                            }
                            Ok(target_table.get_result_data(&cell_ref))
                        } else {
                            if is_self {
                                deps.push(Dependency::LocalColumn(col_idx));
                            } else {
                                deps.push(Dependency::RemoteColumn {
                                    sheet: sheet_name,
                                    col: col_idx,
                                });
                            }
                            let mut results = Vec::new();
                            for r in 0..target_table.row_count() {
                                let cell_ref = CellRef::new(r, col_idx);
                                results.push(target_table.get_result_data(&cell_ref));
                            }
                            Ok(ResultData::List(results))
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
                                let res =
                                    if lf < 0.0 && rf.fract() == 0.0 && rf.abs() <= i32::MAX as f64
                                    {
                                        lf.powi(rf as i32)
                                    } else {
                                        lf.powf(rf)
                                    };
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
                return 0.0.partial_cmp(&(*b as f64)).unwrap_or(std::cmp::Ordering::Equal);
            }
            (ResultData::None, ResultData::Float(b)) => {
                return 0.0.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal);
            }
            (ResultData::Integer(a), ResultData::None) => {
                return (*a as f64).partial_cmp(&0.0).unwrap_or(std::cmp::Ordering::Equal);
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
            (ResultData::String(a), ResultData::String(b)) => {
                a.to_lowercase().cmp(&b.to_lowercase())
            }
            _ => std::cmp::Ordering::Equal,
        }
    }

    pub fn to_f64(&self, val: &ResultData) -> Option<f64> {
        match val {
            ResultData::None => Some(0.0),
            ResultData::Float(f) => Some(*f),
            ResultData::Integer(i) => Some(*i as f64),
            ResultData::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
            ResultData::String(s) => {
                let s_trim = s.trim();
                if s_trim.eq_ignore_ascii_case("true") {
                    Some(1.0)
                } else if s_trim.eq_ignore_ascii_case("false") {
                    Some(0.0)
                } else {
                    s_trim.parse::<f64>().ok()
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

    fn check_direct_string_error(
        &self,
        args: &[ResultData],
        is_direct: &[bool],
    ) -> Option<ResultData> {
        for (i, arg) in args.iter().enumerate() {
            if is_direct.get(i).copied().unwrap_or(false) {
                if let ResultData::String(_) = arg {
                    if self.to_f64(arg).is_none() {
                        return Some(ResultData::Error("#VALUE!".to_string()));
                    }
                }
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
                        prod *= p;
                        has_nums = true;
                    }
                }
                (prod, has_nums)
            }
            _ => (1.0, false),
        }
    }

    fn to_bool(&self, val: &ResultData) -> bool {
        match val {
            ResultData::Boolean(b) => *b,
            ResultData::Integer(i) => *i != 0,
            ResultData::Float(f) => *f != 0.0,
            ResultData::String(s) => !s.is_empty() && s.to_lowercase() != "false",
            _ => false,
        }
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
                let condition = self.to_bool(&cond_val);
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
                    Expr::Number(_) | Expr::String(_) | Expr::Boolean(_) => true,
                    Expr::FunctionCall { name, .. } => {
                        let n = name.to_uppercase();
                        n != "IF" && n != "IFERROR" && n != "CHOOSE"
                    }
                    _ => false,
                };
                arg_is_direct.push(is_direct_arg);
                let eval_res = match self.evaluate_ast(arg, context, row, deps) {
                    Ok(r) => r,
                    Err(EngineError::EvalError(EvalError::UnknownFunction(err_str))) if err_str.starts_with('#') => {
                        ResultData::Error(err_str)
                    }
                    Err(e) => return Err(e),
                };
                evaluated_args.push(eval_res);
            }

            if upper_name != "IFERROR" && upper_name != "ISERROR" && upper_name != "ISNA" {
                if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                    return Ok(err);
                }
            }

            match upper_name.as_str() {
                "SUM" => {
                    if let Some(err) =
                        self.check_direct_string_error(&evaluated_args, &arg_is_direct)
                    {
                        return Ok(err);
                    }
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                        return Ok(err);
                    }
                    let mut sum = 0.0;
                    for (i, arg) in evaluated_args.iter().enumerate() {
                        sum += self.sum_helper(arg, arg_is_direct[i]);
                    }
                    Ok(ResultData::Float(sum))
                }
                "AVERAGE" => {
                    if let Some(err) =
                        self.check_direct_string_error(&evaluated_args, &arg_is_direct)
                    {
                        return Ok(err);
                    }
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
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
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                        return Ok(err);
                    }
                    let mut count = 0;
                    for arg in evaluated_args {
                        count += self.count_helper(&arg);
                    }
                    Ok(ResultData::Float(count as f64))
                }
                "MIN" => {
                    if let Some(err) =
                        self.check_direct_string_error(&evaluated_args, &arg_is_direct)
                    {
                        return Ok(err);
                    }
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
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
                    if let Some(err) =
                        self.check_direct_string_error(&evaluated_args, &arg_is_direct)
                    {
                        return Ok(err);
                    }
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
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
                "CEIL" => {
                    let val = self.to_f64_arg(evaluated_args.first(), "CEIL")?;
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
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                        return Ok(err);
                    }
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(false));
                    }
                    let mut res = true;
                    for arg in &evaluated_args {
                        match arg {
                            ResultData::List(list) => {
                                for item in list {
                                    if !self.to_bool(item) {
                                        res = false;
                                        break;
                                    }
                                }
                            }
                            other => {
                                if !self.to_bool(other) {
                                    res = false;
                                }
                            }
                        }
                        if !res {
                            break;
                        }
                    }
                    Ok(ResultData::Boolean(res))
                }
                "OR" => {
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
                        return Ok(err);
                    }
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::Boolean(false));
                    }
                    let mut res = false;
                    for arg in &evaluated_args {
                        match arg {
                            ResultData::List(list) => {
                                for item in list {
                                    if self.to_bool(item) {
                                        res = true;
                                        break;
                                    }
                                }
                            }
                            other => {
                                if self.to_bool(other) {
                                    res = true;
                                }
                            }
                        }
                        if res {
                            break;
                        }
                    }
                    Ok(ResultData::Boolean(res))
                }
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
                        return Ok(ResultData::String(String::new()));
                    }
                    Ok(ResultData::String(
                        evaluated_args[0].to_string().to_uppercase(),
                    ))
                }
                "LOWER" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::String(String::new()));
                    }
                    Ok(ResultData::String(
                        evaluated_args[0].to_string().to_lowercase(),
                    ))
                }
                "PROPER" => {
                    if evaluated_args.is_empty() {
                        return Ok(ResultData::String(String::new()));
                    }
                    let s = evaluated_args[0].to_string();
                    Ok(ResultData::String(self.proper(&s)))
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
                    if let Some(err) =
                        self.check_direct_string_error(&evaluated_args, &arg_is_direct)
                    {
                        return Ok(err);
                    }
                    if let Some(err) = Self::find_error_in_args(&evaluated_args) {
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
