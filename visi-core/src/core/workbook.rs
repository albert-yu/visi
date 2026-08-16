//! Workbook-level orchestration: the sheets, charts, pivot tables and VBA
//! project of one `.xlsx` file, plus the CRUD and recalculation operations
//! that span more than a single [`Sheet`].
//!
//! This layer is where cross-sheet behavior lives. [`Sheet::commit`] only
//! propagates *local* dependencies; [`WorkbookManager::evaluate`] is what
//! carries values across sheets. Likewise nothing recomputes a pivot table
//! implicitly -- [`WorkbookManager::refresh_pivot_table`] is the only thing
//! that writes a computed grid into cells. Embedders should drive the
//! workbook through this type rather than manipulating [`Sheet`] directly,
//! or cross-sheet formulas and pivot output will silently go stale.
//!
//! Loading and saving are byte-oriented ([`WorkbookManager::load_bytes`] /
//! [`WorkbookManager::save_bytes`]) so this stays usable off the filesystem,
//! including on wasm. The `visi` CLI layers path- and stdio-based helpers on
//! top via its own `WorkbookFile` trait.

use crate::core::formula::CompiledFormula;
use crate::core::grid_edit::{Axis, GridEdit};
use crate::core::parser::col_idx_to_letters;
use crate::core::xlsx::{export_xlsx_data, import_xlsx_data};
use crate::core::{
    ExcelTable, PivotAggregation, PivotArea, PivotField, PivotFilterField, PivotGrid, PivotSource,
    PivotTable, PivotValueField, VbaModule, VbaModuleKind, VbaProject,
    chart::{Chart, ChartType},
    compute_pivot,
    engine::{Context, DataColumn, ResultData, Sheet, generate_unique_id},
    validate_vba_module_name,
};
use crate::{Error, ObjectKind};

/// Keeps a table's column names lined up with its sheet columns after a
/// column insert or delete cut through it.
///
/// `ExcelTable::columns` has one entry per sheet column in
/// `start_col..=end_col`, so a column added or removed inside that span has to
/// add or remove a name at the matching offset -- otherwise every name past
/// the edit describes the wrong column. Called with the table's *pre-edit*
/// extent still in place, which is what `edit.at` is compared against.
fn resize_table_columns(
    table: &mut ExcelTable,
    new_start_col: usize,
    new_end_col: usize,
    edit: &GridEdit,
) {
    if edit.insert {
        // Inserting at or before the table's first column moves the table
        // rather than widening it, so only a strictly-interior insert adds a
        // column -- the same asymmetry `shift_span` encodes for references.
        if edit.at > table.start_col && edit.at <= table.end_col {
            let offset = (edit.at - table.start_col).min(table.columns.len());
            for _ in 0..edit.count {
                table.columns.insert(offset, String::new());
            }
        }
    } else {
        let first = edit.at.max(table.start_col);
        let last = (edit.at + edit.count).min(table.end_col + 1);
        if first < last {
            let lo = (first - table.start_col).min(table.columns.len());
            let hi = (last - table.start_col).min(table.columns.len());
            table.columns.drain(lo..hi);
        }
    }
    // A table loaded from elsewhere may already disagree with its own extent;
    // the new width is the authority either way.
    table
        .columns
        .resize(new_end_col - new_start_col + 1, String::new());
}

/// A single sheet's line in a [`WorkbookSummary`].
pub struct SheetSummary {
    /// The sheet's name.
    pub name: String,
    /// Allocated rows.
    pub row_count: usize,
    /// Allocated columns.
    pub col_count: usize,
    /// How many of its cells hold a formula rather than a literal.
    pub formula_count: usize,
}

/// An overview of a workbook's shape, for reporting rather than editing.
pub struct WorkbookSummary {
    /// The file the workbook was loaded from, as the caller named it.
    pub file_name: String,
    /// How many sheets it has.
    pub sheet_count: usize,
    /// How many charts it has.
    pub chart_count: usize,
    /// One entry per sheet, in workbook order.
    pub sheets: Vec<SheetSummary>,
}

/// A whole workbook: its sheets, charts, pivot tables and VBA project, and the
/// operations that span more than one of them.
///
/// The entry point to this crate, and the layer an embedder should drive.
/// Two behaviors are only correct at this level:
///
/// - **Cross-sheet formulas.** [`Sheet::commit`] propagates local dependencies
///   only; [`WorkbookManager::evaluate`] is what carries values between
///   sheets.
/// - **Pivot tables.** Nothing recomputes one implicitly.
///   [`WorkbookManager::refresh_pivot_table`] is the only thing that writes a
///   computed grid into cells.
///
/// Editing the [`Sheet`]s directly is allowed -- the fields are public -- but
/// skips both, so cross-sheet formulas and pivot output go stale silently.
pub struct WorkbookManager {
    /// The worksheets, in workbook order. Cell coordinates within them are
    /// 0-based.
    pub sheets: Vec<Sheet>,
    /// The charts. Workbook-level rather than sheet-scoped; which sheet a
    /// chart is drawn on comes from its `data_range`.
    pub charts: Vec<Chart>,
    /// The pivot table definitions. Workbook-level, since a pivot's source and
    /// destination may be on different sheets.
    pub pivot_tables: Vec<PivotTable>,
    /// The VBA project, if the workbook has macros.
    pub vba_project: Option<VbaProject>,
}

/// Quotes a materialized pivot label that would otherwise be re-parsed as a
/// number, boolean, or formula by `Sheet::commit`'s literal-cell parsing
/// (mirrors `xlsx::text_cell_src`'s treatment of imported text cells).
fn pivot_label_literal(text: &str) -> String {
    if text.is_empty() {
        String::new()
    } else if text.starts_with('=')
        || text.parse::<f64>().is_ok()
        || text.eq_ignore_ascii_case("true")
        || text.eq_ignore_ascii_case("false")
    {
        format!("\"{}\"", text)
    } else {
        text.to_string()
    }
}

/// Renders one aggregated pivot value as literal cell text; errors (e.g.
/// `AVERAGE` over zero numeric records) are written as their Excel error
/// string rather than `ResultData`'s human-readable `"Error: ..."` form.
fn pivot_value_literal(v: &ResultData) -> String {
    match v {
        ResultData::Error(e) => e.clone(),
        other => other.to_string(),
    }
}

fn remove_pivot_field(fields: &mut Vec<PivotField>, column: &str) -> bool {
    let before = fields.len();
    fields.retain(|f| !f.column.eq_ignore_ascii_case(column));
    before != fields.len()
}

impl WorkbookManager {
    /// Load Excel workbook from bytes buffer
    pub fn load_bytes(buffer: &[u8]) -> crate::Result<Self> {
        let (imported_tables, charts, pivot_tables, vba_project) =
            import_xlsx_data(buffer, &[], |_, _, _| {})?;

        let sheets = imported_tables.into_iter().map(|it| it.sheet).collect();
        Ok(Self {
            sheets,
            charts,
            pivot_tables,
            vba_project,
        })
    }

    /// Serialize the workbook to `.xlsx` bytes.
    ///
    /// The byte-level counterpart to [`Self::load_bytes`]. Callers that want
    /// to read or write an actual file supply their own IO -- the `visi` CLI
    /// does so through its `WorkbookFile` trait.
    pub fn save_bytes(&self) -> crate::Result<Vec<u8>> {
        export_xlsx_data(
            &self.sheets,
            &self.charts,
            &self.pivot_tables,
            self.vba_project.as_ref(),
        )
    }

    /// A new workbook containing a single empty sheet named `Sheet1`.
    pub fn new_empty() -> crate::Result<Self> {
        let mut wb = Self {
            sheets: Vec::new(),
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            vba_project: None,
        };
        wb.add_sheet("Sheet1")?;
        Ok(wb)
    }

    /// Recalculate all formulas in all sheets using visi-core engine
    pub fn evaluate(&mut self) -> crate::Result<()> {
        if self.sheets.is_empty() {
            return Ok(());
        }

        // `self.sheets` is the one place true workbook order exists --
        // `Context.sheets` is an unordered `HashMap` -- so `SHEET()` needs
        // this collected once up front rather than derived from a context.
        let sheet_order: Vec<String> = self.sheets.iter().map(|s| s.name.clone()).collect();

        // Multi-pass evaluation to resolve cross-sheet formula dependencies.
        // Every cell is re-marked dirty at the start of *each* pass, not
        // just once before the loop -- `Sheet::commit` drains and clears a
        // sheet's dirty queue as it processes it, so without re-marking,
        // passes 2 and 3 had nothing left dirty and were silent no-ops.
        // That meant a cross-sheet chain more than one hop deep (sheet A's
        // formula depends on sheet B's formula depending on sheet A) kept
        // whatever stale value pass 1 happened to compute before B had a
        // chance to update -- found via the fuzzer's new cross-sheet
        // generator block (#26).
        for _pass in 0..3 {
            for sheet in &mut self.sheets {
                sheet.mark_all_dirty();
            }
            for i in 0..self.sheets.len() {
                let (left, right) = self.sheets.split_at_mut(i);
                let (target_sheet, right_tail) = right.split_first_mut().unwrap();

                let mut context = Context::new();
                for s in left.iter() {
                    context.add_table(s.name.clone(), s);
                }
                for s in right_tail.iter() {
                    context.add_table(s.name.clone(), s);
                }
                context.pivot_tables = &self.pivot_tables;
                context.sheet_order = sheet_order.clone();

                let _ = target_sheet.commit(Some(&context));
            }
        }

        Ok(())
    }

    /// Evaluates one Excel function against this workbook, outside any cell.
    ///
    /// What `Application.WorksheetFunction.X` calls. Every sheet is in the
    /// context, so an argument naming a range on any of them resolves; the
    /// call itself is hosted on the first sheet, which only matters for the
    /// handful of functions that read the calling cell's position -- and a
    /// macro's call has no calling cell to read.
    pub(crate) fn call_worksheet_function(
        &self,
        name: &str,
        args: &[crate::core::parser::Expr],
    ) -> Result<ResultData, crate::core::EngineError> {
        let Some(host) = self.sheets.first() else {
            return Err(crate::core::EngineError::EvalError(
                crate::core::EvalError::UnknownFunction("no worksheets".to_string()),
            ));
        };
        let mut context = Context::new();
        for s in &self.sheets {
            context.add_table(s.name.clone(), s);
        }
        context.pivot_tables = &self.pivot_tables;
        context.sheet_order = self.sheets.iter().map(|s| s.name.clone()).collect();
        host.call_worksheet_function(name, args, Some(&context))
    }

    /// Find index of sheet by name, or return default index 0 if name is None.
    pub fn find_sheet_index(&self, name_opt: Option<&str>) -> crate::Result<usize> {
        if self.sheets.is_empty() {
            return Err(Error::EmptyWorkbook);
        }

        match name_opt {
            Some(name) => {
                if let Some(idx) = self
                    .sheets
                    .iter()
                    .position(|s| s.name.eq_ignore_ascii_case(name))
                {
                    Ok(idx)
                } else {
                    let available: Vec<String> =
                        self.sheets.iter().map(|s| s.name.clone()).collect();
                    Err(Error::not_found_among(
                        ObjectKind::Sheet,
                        name.to_string(),
                        available,
                    ))
                }
            }
            None => Ok(0),
        }
    }

    /// Get structural summary of workbook
    pub fn get_summary(&self, file_name: &str) -> WorkbookSummary {
        let sheet_summaries = self
            .sheets
            .iter()
            .map(|sheet| {
                let row_count = sheet.row_count();
                let col_count = sheet.col_count();
                let mut formula_count = 0;

                for col in &sheet.columns {
                    for src in &col.src {
                        if src.starts_with('=') {
                            formula_count += 1;
                        }
                    }
                }

                SheetSummary {
                    name: sheet.name.clone(),
                    row_count,
                    col_count,
                    formula_count,
                }
            })
            .collect();

        WorkbookSummary {
            file_name: file_name.to_string(),
            sheet_count: self.sheets.len(),
            chart_count: self.charts.len(),
            sheets: sheet_summaries,
        }
    }

    /// Ensure sheet bounds can accommodate specified target_row and target_col
    pub fn ensure_capacity(&mut self, sheet_idx: usize, target_row: usize, target_col: usize) {
        if sheet_idx >= self.sheets.len() {
            return;
        }
        self.sheets[sheet_idx].ensure_capacity(target_row, target_col);
    }

    /// Merges `style` into one cell's existing style.
    ///
    /// Row and column are 0-based, like every other coordinate on this type.
    /// A1 notation is a CLI/parser-boundary concern: callers holding a
    /// string like `"Sheet2!B3"` parse it themselves and resolve the sheet
    /// prefix against `sheet_name` before calling in.
    pub fn set_cell_style(
        &mut self,
        sheet_name: Option<&str>,
        row: usize,
        col: usize,
        style: crate::core::CellStyle,
    ) -> crate::Result<()> {
        let sheet_idx = self.find_sheet_index(sheet_name)?;
        self.sheets[sheet_idx].update_cell_style(row, col, |s| s.merge(&style));
        Ok(())
    }

    /// Merges `style` into every cell of an inclusive 0-based range.
    pub fn set_range_style(
        &mut self,
        sheet_name: Option<&str>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        style: crate::core::CellStyle,
    ) -> crate::Result<()> {
        if end_row < start_row || end_col < start_col {
            return Err(Error::InvalidRange(
                "range end must not precede its start".to_string(),
            ));
        }
        let sheet_idx = self.find_sheet_index(sheet_name)?;
        for r in start_row..=end_row {
            for c in start_col..=end_col {
                self.sheets[sheet_idx].update_cell_style(r, c, |s| s.merge(&style));
            }
        }
        Ok(())
    }

    /// The style applied to one 0-based cell, if it has one.
    pub fn get_cell_style(
        &self,
        sheet_name: Option<&str>,
        row: usize,
        col: usize,
    ) -> crate::Result<Option<crate::core::CellStyle>> {
        let sheet_idx = self.find_sheet_index(sheet_name)?;
        Ok(self.sheets[sheet_idx].get_cell_style(row, col).cloned())
    }

    /// Sets an Excel Table's visual style, looking the table up by name
    /// across every sheet.
    ///
    /// # Errors
    ///
    /// [`Error::NotFound`] if no table in the workbook has that name.
    pub fn set_table_style(&mut self, table_name: &str, style_name: &str) -> crate::Result<()> {
        for sheet in &mut self.sheets {
            for table in &mut sheet.tables {
                if table.name.eq_ignore_ascii_case(table_name) {
                    table.set_style_name(Some(style_name.to_string()));
                    return Ok(());
                }
            }
        }
        Err(Error::not_found(ObjectKind::Table, table_name.to_string()))
    }

    /// An Excel Table's visual style, or `None` if it has none set.
    ///
    /// # Errors
    ///
    /// [`Error::NotFound`] if no table in the workbook has that name.
    pub fn get_table_style(&self, table_name: &str) -> crate::Result<Option<String>> {
        for sheet in &self.sheets {
            for table in &sheet.tables {
                if table.name.eq_ignore_ascii_case(table_name) {
                    return Ok(table.style_name.clone());
                }
            }
        }
        Err(Error::not_found(ObjectKind::Table, table_name.to_string()))
    }

    /// Update cell source / value at (row, col)
    pub fn set_cell(&mut self, sheet_idx: usize, row: usize, col: usize, value: String) {
        self.ensure_capacity(sheet_idx, row, col);
        let sheet = &mut self.sheets[sheet_idx];
        sheet.set_cell_src(row, col, value);
    }

    /// Insert row at 0-based index.
    ///
    /// Formulas throughout the workbook are rewritten to follow the cells
    /// that moved, as in Excel, and Excel Table and pivot ranges move with
    /// the cells they cover. See `core::grid_edit` for the rules.
    pub fn insert_row(&mut self, sheet_idx: usize, row_idx: usize) -> crate::Result<()> {
        let sheet = &self.sheets[sheet_idx];
        // `Sheet::insert_row` appends when the index is past the end, so the
        // edit the rewrite is told about has to say the same thing.
        let at = row_idx.min(sheet.row_count());
        let edit = GridEdit::insert_row(sheet.id, at);
        self.apply_grid_edit(edit, &[], |wb| wb.sheets[sheet_idx].insert_row(at));
        self.evaluate()
    }

    /// Delete row at 0-based index.
    ///
    /// References to the deleted row become `#REF!` and references below it
    /// move up, as in Excel.
    pub fn delete_row(&mut self, sheet_idx: usize, row_idx: usize) -> crate::Result<()> {
        let sheet = &self.sheets[sheet_idx];
        if row_idx >= sheet.row_count() {
            return Err(Error::OutOfBounds {
                what: "row",
                index: row_idx,
                len: sheet.row_count(),
            });
        }
        let edit = GridEdit::delete_row(sheet.id, row_idx);
        self.apply_grid_edit(edit, &[], |wb| wb.sheets[sheet_idx].delete_row(row_idx));
        self.evaluate()
    }

    /// Insert column at 0-based index.
    pub fn insert_col(&mut self, sheet_idx: usize, col_idx: usize) -> crate::Result<()> {
        let sheet = &self.sheets[sheet_idx];
        let at = col_idx.min(sheet.col_count());
        let edit = GridEdit::insert_col(sheet.id, at);
        self.apply_grid_edit(edit, &[], |wb| wb.sheets[sheet_idx].insert_col(at));
        self.evaluate()
    }

    /// Delete column at 0-based index.
    pub fn delete_col(&mut self, sheet_idx: usize, col_idx: usize) -> crate::Result<()> {
        let sheet = &self.sheets[sheet_idx];
        if col_idx >= sheet.col_count() {
            return Err(Error::OutOfBounds {
                what: "column",
                index: col_idx,
                len: sheet.col_count(),
            });
        }
        // A whole-column reference is held by column id, not position, so the
        // only way to tell whether the edit broke one is to note the id
        // before the column is gone.
        let deleted_col_ids = vec![sheet.columns()[col_idx].id];
        let edit = GridEdit::delete_col(sheet.id, col_idx);
        self.apply_grid_edit(edit, &deleted_col_ids, |wb| {
            wb.sheets[sheet_idx].delete_col(col_idx)
        });
        self.evaluate()
    }

    /// Runs a structural edit, keeping everything that holds a coordinate
    /// pointing at what it pointed at before.
    ///
    /// Three phases, and the order is the whole point:
    ///
    /// 1. **Before the edit**, compile every formula in the workbook and
    ///    shift its references. Compiling needs the grid the formula text was
    ///    written against.
    /// 2. Apply the edit itself, via `apply`.
    /// 3. **After the edit**, serialize the shifted formulas back to text and
    ///    write each one at wherever its own cell moved to.
    ///
    /// Phase 3 cannot be folded into phase 1. A whole-column reference
    /// renders as the column's *current* letter, so serializing `=SUM(B:B)`
    /// before a column is inserted to its left would write `B:B` into a cell
    /// where `B` now names a different column -- and `src` is what the next
    /// recompile reads, so the wrong text wins.
    fn apply_grid_edit(
        &mut self,
        edit: GridEdit,
        deleted_col_ids: &[u64],
        apply: impl FnOnce(&mut Self),
    ) {
        // Phase 1: compile against the pre-edit grid and shift. Formulas the
        // edit does not touch are skipped outright rather than rewritten to
        // an equivalent spelling.
        let mut shifted: Vec<(usize, usize, usize, CompiledFormula)> = Vec::new();
        for (sheet_idx, sheet) in self.sheets.iter().enumerate() {
            for (col_idx, column) in sheet.columns().iter().enumerate() {
                for row_idx in 0..column.len() {
                    let Some(src) = column.src(row_idx).filter(|s| s.starts_with('=')) else {
                        continue;
                    };
                    let compiled = crate::core::parser::compile_formula(src, &self.sheets);
                    if let Some(next) =
                        crate::core::grid_edit::shift_formula(&compiled, &edit, deleted_col_ids)
                    {
                        shifted.push((sheet_idx, col_idx, row_idx, next));
                    }
                }
            }
        }

        // Phase 2.
        apply(self);
        self.shift_table_and_pivot_ranges(&edit);

        // Phase 3.
        for (sheet_idx, col_idx, row_idx, compiled) in shifted {
            let Some((row, col)) = self.moved_cell(&edit, sheet_idx, row_idx, col_idx) else {
                // The cell holding the formula was itself deleted.
                continue;
            };
            let text = crate::core::parser::serialize_formula(&compiled, &self.sheets);
            self.sheets[sheet_idx].set_cell_src(row, col, text);
        }
    }

    /// Where the cell at `(row, col)` on `sheet_idx` ends up after `edit`, or
    /// `None` if the edit deleted it.
    fn moved_cell(
        &self,
        edit: &GridEdit,
        sheet_idx: usize,
        row: usize,
        col: usize,
    ) -> Option<(usize, usize)> {
        if self.sheets[sheet_idx].id != edit.sheet_id {
            return Some((row, col));
        }
        let moved = |index: usize| {
            crate::core::grid_edit::shift_point(index, edit.at, edit.count, edit.insert)
        };
        match edit.axis {
            Axis::Row => Some((moved(row)?, col)),
            Axis::Col => Some((row, moved(col)?)),
        }
    }

    /// Moves the Excel Table and pivot rectangles the edit passed through.
    ///
    /// A table or a pivot source whose every row (or every column) was
    /// deleted has nothing left to describe, so it is dropped -- the same
    /// thing Excel does when you delete the last row of a one-row table.
    fn shift_table_and_pivot_ranges(&mut self, edit: &GridEdit) {
        use crate::core::grid_edit::{shift_point, shift_rect};

        for sheet in &mut self.sheets {
            if sheet.id != edit.sheet_id {
                continue;
            }
            sheet.tables.retain_mut(|table| {
                match shift_rect(
                    edit,
                    table.start_row,
                    table.start_col,
                    table.end_row,
                    table.end_col,
                ) {
                    Some((r0, c0, r1, c1)) => {
                        // Column names are per sheet-column, so dropping a
                        // column has to drop its name with it or every name
                        // past it shifts onto the wrong column.
                        if edit.axis == Axis::Col {
                            resize_table_columns(table, c0, c1, edit);
                        }
                        table.start_row = r0;
                        table.start_col = c0;
                        table.end_row = r1;
                        table.end_col = c1;
                        true
                    }
                    None => false,
                }
            });
        }

        for pivot in &mut self.pivot_tables {
            if let PivotSource::Range {
                sheet_id,
                start_row,
                start_col,
                end_row,
                end_col,
            } = &mut pivot.source
                && *sheet_id == edit.sheet_id
                && let Some((r0, c0, r1, c1)) =
                    shift_rect(edit, *start_row, *start_col, *end_row, *end_col)
            {
                *start_row = r0;
                *start_col = c0;
                *end_row = r1;
                *end_col = c1;
            }

            if pivot.dest_sheet_id == edit.sheet_id {
                // The destination is a corner, not a span. A deleted corner
                // clamps to the edit rather than vanishing: the grid is
                // rewritten wholesale on the next refresh anyway, so what
                // matters is that it names a live cell.
                match edit.axis {
                    Axis::Row => {
                        pivot.dest_row =
                            shift_point(pivot.dest_row, edit.at, edit.count, edit.insert)
                                .unwrap_or(edit.at);
                    }
                    Axis::Col => {
                        pivot.dest_col =
                            shift_point(pivot.dest_col, edit.at, edit.count, edit.insert)
                                .unwrap_or(edit.at);
                    }
                }
                // The last rendered extent is only used to clear stale cells,
                // so a stale one over-clears rather than under-clears; drop it
                // and let the next refresh re-record it.
                pivot.last_output_end_row = None;
                pivot.last_output_end_col = None;
            }
        }
    }

    /// Add new sheet with specified name
    pub fn add_sheet(&mut self, name: &str) -> crate::Result<()> {
        if self
            .sheets
            .iter()
            .any(|s| s.name.eq_ignore_ascii_case(name))
        {
            return Err(Error::AlreadyExists {
                kind: ObjectKind::Sheet,
                name: name.to_string(),
            });
        }

        let mut columns = Vec::new();
        for col_idx in 0..5 {
            let mut col = DataColumn::new(10);
            col.id = generate_unique_id();
            col.name = col_idx_to_letters(col_idx);
            columns.push(col);
        }

        let new_sheet = Sheet {
            id: generate_unique_id(),
            name: name.to_string(),
            columns,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
        };

        self.sheets.push(new_sheet);
        Ok(())
    }

    /// Delete sheet by name
    pub fn delete_sheet(&mut self, name: &str) -> crate::Result<()> {
        let idx = self.find_sheet_index(Some(name))?;
        if self.sheets.len() <= 1 {
            return Err(Error::LastSheetInWorkbook);
        }
        self.sheets.remove(idx);
        Ok(())
    }

    /// Rename sheet
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> crate::Result<()> {
        let idx = self.find_sheet_index(Some(old_name))?;
        if self
            .sheets
            .iter()
            .enumerate()
            .any(|(i, s)| i != idx && s.name.eq_ignore_ascii_case(new_name))
        {
            return Err(Error::NameTaken {
                kind: ObjectKind::Sheet,
                name: new_name.to_string(),
            });
        }
        self.sheets[idx].name = new_name.to_string();
        Ok(())
    }

    /// Add chart to workbook
    #[allow(clippy::too_many_arguments)]
    pub fn add_chart(
        &mut self,
        sheet_name: &str,
        chart_type: ChartType,
        range: String,
        title: Option<String>,
        anchor: Option<(usize, usize)>,
    ) -> crate::Result<u64> {
        let _ = self.find_sheet_index(Some(sheet_name))?;
        let id = generate_unique_id();
        let name = format!("Chart {}", self.charts.len() + 1);
        let (anchor_row, anchor_col) = anchor.unwrap_or((0, 0));

        let chart = Chart {
            id,
            name,
            chart_type,
            data_range: range,
            title,
            xlabel: None,
            ylabel: None,
            show_legend: true,
            anchor_row,
            anchor_col,
        };

        self.charts.push(chart);
        Ok(id)
    }

    /// Edit an existing chart's properties. Every parameter is optional;
    /// `None` leaves that field unchanged. `title`/`xlabel`/`ylabel` are
    /// tri-state (`Option<Option<String>>`): outer `None` leaves the field
    /// unchanged, `Some(None)` clears it, `Some(Some(text))` sets it.
    #[allow(clippy::too_many_arguments)]
    pub fn edit_chart(
        &mut self,
        id: u64,
        name: Option<String>,
        chart_type: Option<ChartType>,
        data_range: Option<String>,
        title: Option<Option<String>>,
        xlabel: Option<Option<String>>,
        ylabel: Option<Option<String>>,
        show_legend: Option<bool>,
        anchor: Option<(usize, usize)>,
    ) -> crate::Result<()> {
        let chart = self
            .charts
            .iter_mut()
            .find(|c| c.id == id)
            .ok_or_else(|| Error::not_found(ObjectKind::Chart, id.to_string()))?;
        if let Some(name) = name {
            chart.name = name;
        }
        if let Some(chart_type) = chart_type {
            chart.chart_type = chart_type;
        }
        if let Some(data_range) = data_range {
            chart.data_range = data_range;
        }
        if let Some(title) = title {
            chart.title = title;
        }
        if let Some(xlabel) = xlabel {
            chart.xlabel = xlabel;
        }
        if let Some(ylabel) = ylabel {
            chart.ylabel = ylabel;
        }
        if let Some(show_legend) = show_legend {
            chart.show_legend = show_legend;
        }
        if let Some((anchor_row, anchor_col)) = anchor {
            chart.anchor_row = anchor_row;
            chart.anchor_col = anchor_col;
        }
        Ok(())
    }

    /// Whether the workbook carries a VBA project.
    pub fn has_vba_project(&self) -> bool {
        self.vba_project.is_some()
    }

    /// Lists every module in the workbook's VBA project, if it has one.
    pub fn list_vba_modules(&self) -> Vec<&VbaModule> {
        self.vba_project
            .as_ref()
            .map(|p| p.modules.iter().collect())
            .unwrap_or_default()
    }

    /// Creates an empty, entirely synthetic VBA project (see
    /// `VbaProject::new_empty`) if this workbook doesn't already have one.
    /// Idempotent.
    pub fn ensure_vba_project(&mut self) -> crate::Result<()> {
        if self.vba_project.is_some() {
            return Ok(());
        }
        self.vba_project = Some(VbaProject::new_empty());
        Ok(())
    }

    /// Adds a new module to the workbook's VBA project (creating the
    /// project from the bundled template first, if needed). `bound_sheet_id`
    /// is required for `VbaModuleKind::Document` (except when `name` is
    /// `"ThisWorkbook"`, which -- like real Excel's own always-present
    /// ThisWorkbook module -- isn't tied to a specific sheet; any
    /// `bound_sheet_id` passed alongside it is ignored rather than stored)
    /// -- note this does NOT rename the sheet, or vice versa; Excel allows a
    /// document module's own name and its sheet's display name to diverge,
    /// and this codebase deliberately doesn't cascade one into the other.
    pub fn add_vba_module(
        &mut self,
        name: String,
        kind: VbaModuleKind,
        source: String,
        bound_sheet_id: Option<u64>,
    ) -> crate::Result<()> {
        validate_vba_module_name(&name).map_err(|reason| Error::InvalidName {
            kind: ObjectKind::VbaModule,
            name: name.clone(),
            reason,
        })?;
        let is_this_workbook = kind == VbaModuleKind::Document && name == "ThisWorkbook";
        if kind == VbaModuleKind::Document && !is_this_workbook {
            let sheet_id = bound_sheet_id
                .ok_or_else(|| Error::Vba("document modules require a bound sheet".to_string()))?;
            if !self.sheets.iter().any(|s| s.id == sheet_id) {
                return Err(Error::not_found(ObjectKind::Sheet, sheet_id.to_string()));
            }
        }
        self.ensure_vba_project()?;
        let project = self.vba_project.as_mut().unwrap();
        if project.module_name_taken(&name) {
            return Err(Error::AlreadyExists {
                kind: ObjectKind::VbaModule,
                name: name.to_string(),
            });
        }
        if kind == VbaModuleKind::Document
            && bound_sheet_id.is_some()
            && project
                .modules
                .iter()
                .any(|m| m.kind == VbaModuleKind::Document && m.bound_sheet_id == bound_sheet_id)
        {
            return Err(Error::DocumentModuleExists);
        }
        // Donate p-code prefix bytes and a module cookie from any existing
        // module in this same project -- proven (via a scratchpad
        // proof-of-concept against real Excel) that the prefix's content
        // doesn't need to correspond to the module it precedes, only its
        // shape matters. If this project has no modules yet (the common
        // case for one freshly created by `ensure_vba_project`), fall back
        // to the synthetic seed values instead.
        let prefix_bytes = project
            .modules
            .first()
            .map(|m| m.prefix_bytes.clone())
            .unwrap_or_else(|| project.seed_prefix_bytes.clone());
        let module_cookie = project
            .modules
            .first()
            .map(|m| m.module_cookie)
            .unwrap_or(project.seed_module_cookie);
        let stored_bound_sheet_id = if kind == VbaModuleKind::Document && !is_this_workbook {
            bound_sheet_id
        } else {
            None
        };
        project.modules.push(VbaModule {
            name,
            kind,
            source,
            bound_sheet_id: stored_bound_sheet_id,
            prefix_bytes,
            module_cookie,
            // A brand-new module has no already-compressed form to reuse.
            cached_compressed_source: None,
        });
        Ok(())
    }

    /// Removes a VBA module by name, matched case-insensitively.
    ///
    /// # Errors
    ///
    /// [`Error::Vba`] if the workbook has no VBA project, or
    /// [`Error::NotFound`] if it has no module by that name.
    pub fn remove_vba_module(&mut self, name: &str) -> crate::Result<()> {
        let project = self
            .vba_project
            .as_mut()
            .ok_or_else(|| Error::Vba("workbook has no VBA project".to_string()))?;
        let before = project.modules.len();
        project
            .modules
            .retain(|m| !m.name.eq_ignore_ascii_case(name));
        if project.modules.len() == before {
            return Err(Error::not_found(ObjectKind::VbaModule, name.to_string()));
        }
        Ok(())
    }

    /// Renames a VBA module.
    ///
    /// Renames only the module; VBA source that calls into it is not
    /// rewritten, so a module referenced by name elsewhere will no longer
    /// resolve.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidName`] if `new_name` is not a valid VBA identifier,
    /// [`Error::AlreadyExists`] if another module already has it,
    /// [`Error::Vba`] if the workbook has no VBA project, or
    /// [`Error::NotFound`] if it has no module called `old_name`.
    pub fn rename_vba_module(&mut self, old_name: &str, new_name: &str) -> crate::Result<()> {
        validate_vba_module_name(new_name).map_err(|reason| Error::InvalidName {
            kind: ObjectKind::VbaModule,
            name: new_name.to_string(),
            reason,
        })?;
        let project = self
            .vba_project
            .as_mut()
            .ok_or_else(|| Error::Vba("workbook has no VBA project".to_string()))?;
        if !old_name.eq_ignore_ascii_case(new_name) && project.module_name_taken(new_name) {
            return Err(Error::AlreadyExists {
                kind: ObjectKind::VbaModule,
                name: new_name.to_string(),
            });
        }
        let module = project
            .find_module_mut(old_name)
            .ok_or_else(|| Error::not_found(ObjectKind::VbaModule, old_name))?;
        module.name = new_name.to_string();
        Ok(())
    }

    /// Replaces a VBA module's source text.
    ///
    /// The caller supplies the whole module body, including its
    /// `Attribute VB_Name = "..."` line, matching how real Excel-authored
    /// module streams are shaped.
    ///
    /// # Errors
    ///
    /// [`Error::Vba`] if the workbook has no VBA project, or
    /// [`Error::NotFound`] if it has no module by that name.
    pub fn set_vba_module_source(&mut self, name: &str, source: String) -> crate::Result<()> {
        let project = self
            .vba_project
            .as_mut()
            .ok_or_else(|| Error::Vba("workbook has no VBA project".to_string()))?;
        let module = project
            .find_module_mut(name)
            .ok_or_else(|| Error::not_found(ObjectKind::VbaModule, name))?;
        module.source = source;
        // Invalidate the cached compressed form -- see
        // VbaModule::cached_compressed_source -- since it no longer matches
        // the new `source`.
        module.cached_compressed_source = None;
        Ok(())
    }

    /// Delete chart by u64 ID
    pub fn delete_chart(&mut self, id: u64) -> crate::Result<()> {
        if let Some(pos) = self.charts.iter().position(|c| c.id == id) {
            self.charts.remove(pos);
            Ok(())
        } else {
            Err(Error::not_found(ObjectKind::Chart, id.to_string()))
        }
    }

    /// Find the sheet that owns the table with the given name, and the
    /// table itself. Table names are unique across the whole workbook.
    pub fn find_table(&self, name: &str) -> Option<(&Sheet, &ExcelTable)> {
        self.sheets
            .iter()
            .find_map(|s| s.find_table(name).map(|t| (s, t)))
    }

    /// List every table in the workbook, alongside the name of the sheet it
    /// lives on.
    pub fn list_tables(&self) -> Vec<(&str, &ExcelTable)> {
        self.sheets
            .iter()
            .flat_map(|s| s.tables.iter().map(move |t| (s.name.as_str(), t)))
            .collect()
    }

    fn find_table_sheet_index(&self, name: &str) -> crate::Result<usize> {
        self.sheets
            .iter()
            .position(|s| s.find_table(name).is_some())
            .ok_or_else(|| Error::not_found(ObjectKind::Table, name))
    }

    fn table_name_taken(&self, name: &str) -> bool {
        self.sheets
            .iter()
            .any(|s| s.tables.iter().any(|t| t.name.eq_ignore_ascii_case(name)))
    }

    /// Define a new Excel Table over an existing cell range on a sheet.
    /// Table names are unique across the entire workbook (not just the
    /// sheet), matching how Excel itself scopes structured-reference names.
    #[allow(clippy::too_many_arguments)]
    pub fn add_table(
        &mut self,
        sheet_name: Option<&str>,
        name: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        has_header_row: bool,
        has_totals_row: bool,
    ) -> crate::Result<u64> {
        if self.table_name_taken(name) {
            return Err(Error::AlreadyExists {
                kind: ObjectKind::Table,
                name: name.to_string(),
            });
        }
        let idx = self.find_sheet_index(sheet_name)?;
        self.sheets[idx]
            .add_table(
                name.to_string(),
                start_row,
                start_col,
                end_row,
                end_col,
                has_header_row,
                has_totals_row,
            )
            .map_err(Error::InvalidArgument)
    }

    /// Delete a table by name (leaves the underlying cell contents alone).
    pub fn delete_table(&mut self, name: &str) -> crate::Result<()> {
        let idx = self.find_table_sheet_index(name)?;
        self.sheets[idx]
            .delete_table_by_name(name)
            .map_err(Error::InvalidArgument)
    }

    /// Rename a table.
    pub fn rename_table(&mut self, old_name: &str, new_name: &str) -> crate::Result<()> {
        if !old_name.eq_ignore_ascii_case(new_name) && self.table_name_taken(new_name) {
            return Err(Error::NameTaken {
                kind: ObjectKind::Table,
                name: new_name.to_string(),
            });
        }
        let idx = self.find_table_sheet_index(old_name)?;
        self.sheets[idx]
            .rename_table(old_name, new_name)
            .map_err(Error::InvalidArgument)?;
        // Excel updates every structured reference to a renamed table, so a
        // formula like `=SUM(Sales[Amount])` keeps working after "Sales" is
        // renamed; match that instead of silently breaking those formulas.
        self.rewrite_table_references(old_name, Some(new_name), None);
        self.evaluate()
    }

    /// Rewrites every formula in the workbook that structurally references
    /// `table_name` (optionally renaming the table and/or one column),
    /// mirroring how Excel keeps structured references in sync when a Table
    /// or one of its column headers is renamed.
    fn rewrite_table_references(
        &mut self,
        table_name: &str,
        new_table_name: Option<&str>,
        col_rename: Option<(&str, &str)>,
    ) {
        for sheet in &mut self.sheets {
            for col_idx in 0..sheet.columns.len() {
                let row_count = sheet.columns[col_idx].src.len();
                for row_idx in 0..row_count {
                    let src = sheet.columns[col_idx].src[row_idx].clone();
                    if let Some(new_src) = crate::core::parser::rewrite_structured_table_reference(
                        &src,
                        table_name,
                        new_table_name,
                        col_rename,
                    ) {
                        sheet.set_cell_src(row_idx, col_idx, new_src);
                    }
                }
            }
        }
    }

    /// Resize a table by moving its bottom-right corner.
    pub fn resize_table(
        &mut self,
        name: &str,
        new_end_row: usize,
        new_end_col: usize,
    ) -> crate::Result<()> {
        let idx = self.find_table_sheet_index(name)?;
        self.sheets[idx]
            .resize_table(name, new_end_row, new_end_col)
            .map_err(Error::InvalidArgument)
    }

    /// Rename one column (0-based, relative to the table) of a table.
    pub fn rename_table_column(
        &mut self,
        table_name: &str,
        col_index: usize,
        new_name: &str,
    ) -> crate::Result<()> {
        let idx = self.find_table_sheet_index(table_name)?;
        let old_col_name = self.sheets[idx]
            .find_table(table_name)
            .and_then(|t| t.columns.get(col_index).cloned())
            .ok_or_else(|| {
                Error::InvalidArgument(format!(
                    "column index {col_index} out of bounds for table '{table_name}'"
                ))
            })?;
        self.sheets[idx]
            .rename_table_column(table_name, col_index, new_name)
            .map_err(Error::InvalidArgument)?;
        // As with rename_table, keep dependent formulas working across the
        // rename instead of leaving them referencing the old column name.
        self.rewrite_table_references(table_name, None, Some((&old_col_name, new_name)));
        self.evaluate()
    }

    /// Find a pivot table by name (case-insensitive).
    pub fn find_pivot_table(&self, name: &str) -> Option<&PivotTable> {
        self.pivot_tables
            .iter()
            .find(|p| p.name.eq_ignore_ascii_case(name))
    }

    fn find_pivot_table_index(&self, name: &str) -> crate::Result<usize> {
        self.pivot_tables
            .iter()
            .position(|p| p.name.eq_ignore_ascii_case(name))
            .ok_or_else(|| Error::not_found(ObjectKind::PivotTable, name))
    }

    /// List every pivot table in the workbook.
    pub fn list_pivot_tables(&self) -> &[PivotTable] {
        &self.pivot_tables
    }

    fn pivot_table_name_taken(&self, name: &str) -> bool {
        self.pivot_tables
            .iter()
            .any(|p| p.name.eq_ignore_ascii_case(name))
    }

    /// Defines a new pivot table sourced from an existing Excel Table, with
    /// no fields assigned yet -- mirroring Excel inserting an empty
    /// PivotTable shell that fills in as fields are added to it.
    #[allow(clippy::too_many_arguments)]
    pub fn add_pivot_table_from_table(
        &mut self,
        name: &str,
        source_table_name: &str,
        dest_sheet_name: Option<&str>,
        dest_row: usize,
        dest_col: usize,
        grand_totals_row: bool,
        grand_totals_col: bool,
    ) -> crate::Result<u64> {
        if self.pivot_table_name_taken(name) {
            return Err(Error::AlreadyExists {
                kind: ObjectKind::PivotTable,
                name: name.to_string(),
            });
        }
        self.find_table(source_table_name)
            .ok_or_else(|| Error::not_found(ObjectKind::Table, source_table_name))?;
        let dest_idx = self.find_sheet_index(dest_sheet_name)?;
        let id = generate_unique_id();
        self.pivot_tables.push(PivotTable {
            id,
            name: name.to_string(),
            source: PivotSource::Table {
                name: source_table_name.to_string(),
            },
            dest_sheet_id: self.sheets[dest_idx].id,
            dest_row,
            dest_col,
            row_fields: Vec::new(),
            col_fields: Vec::new(),
            value_fields: Vec::new(),
            filter_fields: Vec::new(),
            grand_totals_row,
            grand_totals_col,
            last_output_end_row: None,
            last_output_end_col: None,
        });
        self.refresh_pivot_table(name)?;
        Ok(id)
    }

    /// Defines a new pivot table sourced from a plain cell range (its first
    /// row is treated as column headers), with no fields assigned yet.
    #[allow(clippy::too_many_arguments)]
    pub fn add_pivot_table_from_range(
        &mut self,
        name: &str,
        source_sheet_name: Option<&str>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        dest_sheet_name: Option<&str>,
        dest_row: usize,
        dest_col: usize,
        grand_totals_row: bool,
        grand_totals_col: bool,
    ) -> crate::Result<u64> {
        if self.pivot_table_name_taken(name) {
            return Err(Error::AlreadyExists {
                kind: ObjectKind::PivotTable,
                name: name.to_string(),
            });
        }
        let src_idx = self.find_sheet_index(source_sheet_name)?;
        let dest_idx = self.find_sheet_index(dest_sheet_name)?;
        let id = generate_unique_id();
        self.pivot_tables.push(PivotTable {
            id,
            name: name.to_string(),
            source: PivotSource::Range {
                sheet_id: self.sheets[src_idx].id,
                start_row,
                start_col,
                end_row,
                end_col,
            },
            dest_sheet_id: self.sheets[dest_idx].id,
            dest_row,
            dest_col,
            row_fields: Vec::new(),
            col_fields: Vec::new(),
            value_fields: Vec::new(),
            filter_fields: Vec::new(),
            grand_totals_row,
            grand_totals_col,
            last_output_end_row: None,
            last_output_end_col: None,
        });
        self.refresh_pivot_table(name)?;
        Ok(id)
    }

    /// Deletes a pivot table definition and clears its last rendered output
    /// range (leaves the source data untouched).
    pub fn delete_pivot_table(&mut self, name: &str) -> crate::Result<()> {
        let idx = self.find_pivot_table_index(name)?;
        let pivot = self.pivot_tables.remove(idx);
        if let (Some(end_row), Some(end_col)) =
            (pivot.last_output_end_row, pivot.last_output_end_col)
            && let Some(sheet_idx) = self.sheets.iter().position(|s| s.id == pivot.dest_sheet_id)
        {
            self.clear_range(sheet_idx, pivot.dest_row, pivot.dest_col, end_row, end_col);
        }
        Ok(())
    }

    /// Renames a pivot table (names are unique workbook-wide, like tables).
    pub fn rename_pivot_table(&mut self, old_name: &str, new_name: &str) -> crate::Result<()> {
        if !old_name.eq_ignore_ascii_case(new_name) && self.pivot_table_name_taken(new_name) {
            return Err(Error::NameTaken {
                kind: ObjectKind::PivotTable,
                name: new_name.to_string(),
            });
        }
        let idx = self.find_pivot_table_index(old_name)?;
        self.pivot_tables[idx].name = new_name.to_string();
        Ok(())
    }

    /// Adds a field to one of a pivot table's four areas (Row/Column/
    /// Value/Filter) and immediately refreshes its output, mirroring
    /// Excel's live-updating field list.
    ///
    /// A field can only occupy one area at a time, exactly like dragging a
    /// field to a new area in Excel's field list moves it rather than
    /// duplicating it (confirmed via the win32com driver: setting
    /// `PivotField.Orientation` a second time relocates the field). Value
    /// fields are the one exception -- Excel allows the same source column
    /// to appear as multiple value fields simultaneously (e.g. both "Sum of
    /// Amount" and "Min of Amount"), so adding to `PivotArea::Value` does
    /// not evict the column from Row/Column/Filter, and vice versa a
    /// Row/Column/Filter add does not evict existing value fields.
    pub fn add_pivot_field(
        &mut self,
        pivot_name: &str,
        area: PivotArea,
        column: &str,
        aggregation: Option<PivotAggregation>,
    ) -> crate::Result<()> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        if !matches!(area, PivotArea::Value) {
            let pivot = &mut self.pivot_tables[idx];
            remove_pivot_field(&mut pivot.row_fields, column);
            remove_pivot_field(&mut pivot.col_fields, column);
            pivot
                .filter_fields
                .retain(|f| !f.column.eq_ignore_ascii_case(column));
        }
        match area {
            PivotArea::Row => self.pivot_tables[idx]
                .row_fields
                .push(PivotField::new(column)),
            PivotArea::Column => self.pivot_tables[idx]
                .col_fields
                .push(PivotField::new(column)),
            PivotArea::Value => {
                let agg = aggregation.unwrap_or(PivotAggregation::Sum);
                self.pivot_tables[idx]
                    .value_fields
                    .push(PivotValueField::new(column, agg));
            }
            PivotArea::Filter => self.pivot_tables[idx]
                .filter_fields
                .push(PivotFilterField::new(column)),
        }
        self.refresh_pivot_table(pivot_name)
    }

    /// Removes a field from one of a pivot table's four areas and
    /// refreshes its output.
    pub fn remove_pivot_field(
        &mut self,
        pivot_name: &str,
        area: PivotArea,
        column: &str,
    ) -> crate::Result<()> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        let removed = match area {
            PivotArea::Row => remove_pivot_field(&mut self.pivot_tables[idx].row_fields, column),
            PivotArea::Column => remove_pivot_field(&mut self.pivot_tables[idx].col_fields, column),
            PivotArea::Value => {
                let before = self.pivot_tables[idx].value_fields.len();
                self.pivot_tables[idx]
                    .value_fields
                    .retain(|f| !f.column.eq_ignore_ascii_case(column));
                before != self.pivot_tables[idx].value_fields.len()
            }
            PivotArea::Filter => {
                let before = self.pivot_tables[idx].filter_fields.len();
                self.pivot_tables[idx]
                    .filter_fields
                    .retain(|f| !f.column.eq_ignore_ascii_case(column));
                before != self.pivot_tables[idx].filter_fields.len()
            }
        };
        if !removed {
            return Err(Error::not_found(
                ObjectKind::PivotField,
                format!("{column}' in pivot table '{pivot_name}"),
            ));
        }
        self.refresh_pivot_table(pivot_name)
    }

    /// Restricts (or clears, with `values: None`) a filter field's allowed
    /// values and refreshes the pivot table's output.
    pub fn set_pivot_filter(
        &mut self,
        pivot_name: &str,
        column: &str,
        values: Option<Vec<String>>,
    ) -> crate::Result<()> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        let field = self.pivot_tables[idx]
            .filter_fields
            .iter_mut()
            .find(|f| f.column.eq_ignore_ascii_case(column))
            .ok_or_else(|| {
                Error::not_found(
                    ObjectKind::PivotField,
                    format!("{column}' on pivot table '{pivot_name}"),
                )
            })?;
        field.selected_values = values;
        self.refresh_pivot_table(pivot_name)
    }

    /// Recomputes a pivot table's aggregation and re-materializes it as
    /// plain values onto its destination sheet. Like Excel, a pivot table
    /// only updates on an explicit refresh, never automatically as its
    /// source data changes.
    pub fn refresh_pivot_table(&mut self, pivot_name: &str) -> crate::Result<()> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        let pivot = self.pivot_tables[idx].clone();
        let dest_idx = self
            .sheets
            .iter()
            .position(|s| s.id == pivot.dest_sheet_id)
            .ok_or_else(|| {
                Error::InvalidArgument(
                    "pivot table's destination sheet no longer exists".to_string(),
                )
            })?;

        let grid: Option<PivotGrid> = if pivot.value_fields.is_empty() {
            None
        } else {
            let sheet_refs: Vec<&Sheet> = self.sheets.iter().collect();
            Some(compute_pivot(&sheet_refs, &pivot).map_err(Error::InvalidArgument)?)
        };

        // Clear the previously rendered area first, since a refresh can
        // shrink the grid (fewer groups, a narrower filter, etc).
        if let (Some(old_end_row), Some(old_end_col)) =
            (pivot.last_output_end_row, pivot.last_output_end_col)
        {
            self.clear_range(
                dest_idx,
                pivot.dest_row,
                pivot.dest_col,
                old_end_row,
                old_end_col,
            );
        }

        let new_bounds = grid.as_ref().map(|grid| {
            let height = grid.height();
            let width = grid.width.max(1);
            self.ensure_capacity(
                dest_idx,
                pivot.dest_row + height.saturating_sub(1),
                pivot.dest_col + width.saturating_sub(1),
            );

            let mut r = pivot.dest_row;
            for (name, state) in &grid.filter_rows {
                self.set_cell(dest_idx, r, pivot.dest_col, pivot_label_literal(name));
                self.set_cell(dest_idx, r, pivot.dest_col + 1, pivot_label_literal(state));
                r += 1;
            }
            if !grid.filter_rows.is_empty() {
                r += 1; // blank spacer row before the grid, matching Excel
            }
            for header in &grid.header_rows {
                for (c, text) in header.iter().enumerate() {
                    self.set_cell(dest_idx, r, pivot.dest_col + c, pivot_label_literal(text));
                }
                r += 1;
            }
            for body in &grid.body_rows {
                for (c, label) in body.row_labels.iter().enumerate() {
                    self.set_cell(dest_idx, r, pivot.dest_col + c, pivot_label_literal(label));
                }
                for (c, val) in body.values.iter().enumerate() {
                    self.set_cell(
                        dest_idx,
                        r,
                        pivot.dest_col + body.row_labels.len() + c,
                        pivot_value_literal(val),
                    );
                }
                r += 1;
            }
            (
                pivot.dest_row + height.saturating_sub(1),
                pivot.dest_col + width.saturating_sub(1),
            )
        });

        self.pivot_tables[idx].last_output_end_row = new_bounds.map(|(r, _)| r);
        self.pivot_tables[idx].last_output_end_col = new_bounds.map(|(_, c)| c);
        self.evaluate()
    }

    /// Blanks every cell in the given rectangular range (inclusive),
    /// clipped to the sheet's current bounds. Used to wipe a pivot table's
    /// previous output before re-rendering a possibly smaller grid.
    fn clear_range(
        &mut self,
        sheet_idx: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) {
        if sheet_idx >= self.sheets.len() {
            return;
        }
        let (row_count, col_count) = {
            let s = &self.sheets[sheet_idx];
            (s.row_count(), s.col_count())
        };
        if row_count == 0 || col_count == 0 {
            return;
        }
        for r in start_row..=end_row.min(row_count - 1) {
            for c in start_col..=end_col.min(col_count - 1) {
                self.sheets[sheet_idx].set_cell_src(r, c, String::new());
            }
        }
    }
}
