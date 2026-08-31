//! The VBA host object model: Phase 2 of `docs/vba-macro-support.md`.
//!
//! This is the layer that lets a macro touch a workbook. Everything in it was
//! measured against Excel for Mac 16.112 with `fuzz/vba_host_probe.py` rather
//! than taken from documentation, for the same reason Phase 1's `Variant`
//! rules were: a careful reading of the docs already got one of those
//! backwards, and a host model that is subtly wrong produces wrong *numbers*
//! in a saved file, which is worse than an error.
//!
//! Five design decisions carry most of the weight.
//!
//! **An object is a handle, not a pointer and not a bare value.** [`ObjRef`]
//! holds ids, so it is `Copy` and never borrows the workbook. That is what
//! lets [`Host`] hold `&mut WorkbookManager` for the whole run without the
//! borrow checker fighting every statement. A `Range` carries `sheet_id`, not
//! a sheet index or name, for exactly the reason compiled formulas do: a
//! macro can rename or reorder sheets mid-run.
//!
//! **`Is` compares an identity token, not the coordinates.** Measured, and
//! the opposite of the obvious design: `ws.Range("A1") Is ws.Range("A1")` is
//! **False** in Excel, because each call constructs a fresh object, while
//! `Set r = ws.Range("A1")` followed by `Set q = r` makes `q Is r` True. So
//! every constructed `Range` gets a token that copying preserves and
//! reconstructing does not. Worksheets are the other way round --
//! `ws Is wb.Worksheets(1)` is True -- because Excel hands out a cached
//! object per sheet, so a `Worksheet`'s identity is its `sheet_id`.
//!
//! **A `Range`'s coordinates live in the host, not in the object, because a
//! `Range` tracks structural edits.** `Set r = ws.Range("A5")` then
//! `ws.Rows(1).Insert` leaves `r` reading `$A$6` -- it follows the data, and
//! `r.Value` is still what was in `A5`. Measured, along with the rest of
//! `fuzz/vba_range_tracking_probe.py`. So `ObjRef::Range` is a handle into
//! [`Host::ranges`] and the edit rewrites the table, which is also what makes
//! a copy taken *before* the edit track: `q Is r` and both read `$A$6`.
//!
//! The geometry turned out to be exactly `core::grid_edit`'s, case for case
//! -- inserting at a range's first row moves it while inserting one row lower
//! grows it, deleting part of it shrinks it -- so both go through
//! [`crate::core::grid_edit::shift_span`] rather than through two
//! hand-written rules that could drift. A range whose every cell is deleted
//! becomes [`RangeState::Dead`]; see that type for what Excel does with one.
//!
//! **A read recalculates if a write is outstanding.** Excel in automatic mode
//! recalculates after every assignment, and it is observable: writing `A1`
//! and then reading a `D1` that holds `=A1*2` gives the new value. Doing that
//! literally here would mean running [`WorkbookManager::evaluate`] -- three
//! passes over every sheet -- once per assignment, which a loop writing a
//! thousand cells cannot afford. Instead a write sets a `stale` flag and the
//! next read that could observe it pays for one recalculation, so a run of
//! consecutive writes costs one rather than one each. The observable
//! behaviour is the same; only the timing differs, and `evaluate`'s own fixed
//! three-pass limit is the deeper staleness risk (see its doc comment).
//!
//! **Everything outside the allow-list raises 438 naming the construct.** The
//! refusal is the feature, not a gap in it. Widening the list is a decision.

use std::collections::HashMap;

use crate::core::engine::{CellRef, ResultData, Sheet};
use crate::core::grid_edit::{Axis, GridEdit, shift_span};
use crate::core::{CellStyle, ExcelTable, PivotTable};

use super::color;
use crate::core::parser::{Expr as FExpr, col_idx_to_letters};
use crate::core::workbook::WorkbookManager;

use super::value::{VResult, VarArray, Variant, VbaError};

/// Rows in an Excel worksheet. `ws.Cells` is the whole grid, not the part
/// `visi` happens to have allocated, which is why `ws.Cells.Count` overflows
/// a `Long` in Excel -- and, now, here.
pub const MAX_ROWS: u32 = 1_048_576;
/// Columns in an Excel worksheet (`A` through `XFD`).
pub const MAX_COLS: u32 = 16_384;

/// How many cells a macro may cause to be *allocated*.
///
/// Excel's grid is sparse; `visi`'s [`Sheet`] is a dense `Vec` per column, so
/// `ws.Range("XFD1048576").Value = 1` would ask for 17 billion cells. Excel
/// would shrug; this would exhaust memory and take the process down, which is
/// not an outcome a guard may have. Error 7 ("Out of memory") is what VBA
/// itself reports when an allocation fails, so a macro that trips this sees a
/// number it could plausibly have seen from Excel.
const MAX_ALLOCATED_CELLS: u64 = 4_000_000;

/// A reference to a host object, or `Nothing`.
///
/// Deliberately a plain value: no lifetimes, no borrow of the workbook, so a
/// [`Variant`] holding one stays `Clone` and can outlive any single statement.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ObjRef {
    /// An unset object reference. `TypeName` says `"Nothing"`.
    Nothing,
    /// `Application`.
    Application,
    /// `Application.WorksheetFunction`.
    ///
    /// A separate object from [`ObjRef::Application`] because the *same*
    /// function reached through the two behaves differently on failure:
    /// `WorksheetFunction.VLookup` raises error 1004, while
    /// `Application.VLookup` returns an error `Variant` that `IsError`
    /// detects. Both measured. One implementation, two call paths.
    WorksheetFunction,
    /// `ThisWorkbook` / `ActiveWorkbook`. There is only ever one.
    Workbook,
    /// The `Worksheets` / `Sheets` collection. `TypeName` says `"Sheets"`.
    Worksheets,
    /// One worksheet, by its stable id. Identity *is* the id: Excel hands out
    /// a cached object per sheet, so `ws Is wb.Worksheets(1)` is True.
    Worksheet(u64),
    /// The `ListObjects` collection of one worksheet, by sheet id.
    ListObjects(u64),
    /// One Excel Table, by its workbook-unique id.
    ///
    /// By id rather than by name because the id is what survives a rename --
    /// and `ListObject.Name = "X"` is a supported write, so a macro can hold
    /// a table across one.
    ListObject(u64),
    /// A table's `ListColumns` collection, by table id.
    ListColumns(u64),
    /// One table column, by table id and 0-based position within the table.
    ListColumn(u64, u32),
    /// A table's `ListRows` collection, by table id.
    ListRows(u64),
    /// One table data row, by table id and 0-based position within the data
    /// body.
    ListRow(u64, u32),
    /// The `PivotTables` collection of one worksheet, by sheet id.
    PivotTables(u64),
    /// One pivot table, by its workbook-unique id.
    PivotTable(u64),
    /// A pivot table's `PivotFields` collection, by pivot id.
    PivotFields(u64),
    /// One pivot field, by pivot id and source-column position -- Excel's
    /// `PivotFields` has one entry per *source column*, whatever area (if
    /// any) it currently occupies.
    PivotField(u64, u32),
    /// `Range.Interior`, by the handle of the range it belongs to.
    ///
    /// Excel hands out a distinct object (`TypeName` is `"Interior"`,
    /// measured) but it has no identity of its own worth modelling: it is a
    /// view onto the same cells, so it rides on the range's handle and
    /// tracks structural edits for free.
    Interior(u64),
    /// `Range.Font`, by the handle of the range it belongs to. `TypeName` is
    /// `"Font"`, measured.
    Font(u64),
    /// A rectangular range of cells, by handle into [`Host::ranges`].
    ///
    /// A handle rather than the coordinates because a `Range` **tracks
    /// structural edits**: inserting a row above one moves it, and every copy
    /// of it moves too, so the location has to live in one place that the
    /// edit can rewrite. The handle doubles as the identity token for `Is`.
    Range(u64),
    /// A user-defined class instance, by handle into the interpreter's class instance table.
    UserClass(u64),
}

/// Where a `Range` currently points, or that it no longer points anywhere.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RangeState {
    /// The range covers these cells.
    Live(RangeRef),
    /// Every cell the range covered was deleted.
    ///
    /// Measured, and none of it is guessable: the object is **not**
    /// `Nothing` (`r Is Nothing` is False) and still reports
    /// `TypeName(r)` as `"Range"`, but *every* member access raises
    /// `Method '<name>' of object 'Range' failed`.
    ///
    /// Excel's `Err.Number` for this is not reproducible -- the same case
    /// came back as `-1667945984` on one run and `-1667949824` on the next,
    /// and two different members gave the same number on one run and
    /// different numbers on another. Pinning it would be pinning noise, so
    /// visi raises 1004, the number Excel on Windows documents for this
    /// message and the one this module already uses for the rest of the
    /// "object-defined error" family. Recorded in `docs/excel-discrepancies.md`.
    Dead,
}

/// A `Range`'s rectangle: which sheet, and which cells.
///
/// Identity is *not* in here -- that is the [`ObjRef::Range`] handle, so that
/// two ranges over the same cells stay different objects and a range that
/// moves stays the same one.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeRef {
    /// The sheet the range lives on, by stable id rather than index or name.
    pub sheet_id: u64,
    /// 0-based top row.
    pub row: u32,
    /// 0-based left column.
    pub col: u32,
    /// Rows spanned; never zero.
    pub height: u32,
    /// Columns spanned; never zero.
    pub width: u32,
}

impl RangeRef {
    /// Whether this range is exactly one cell, which decides whether `.Value`
    /// reads a scalar or an array.
    pub fn is_single(&self) -> bool {
        self.height == 1 && self.width == 1
    }

    /// Cells covered. `u64` because a whole sheet has more than a `u32` holds
    /// -- and more than `Range.Count`'s `Long` holds, which is why
    /// `ws.Cells.Count` is error 6 in Excel.
    pub fn count(&self) -> u64 {
        self.height as u64 * self.width as u64
    }
}

impl ObjRef {
    /// What `TypeName()` reports. All measured.
    pub fn type_name(&self) -> &'static str {
        match self {
            ObjRef::Nothing => "Nothing",
            ObjRef::Application => "Application",
            ObjRef::WorksheetFunction => "WorksheetFunction",
            ObjRef::Workbook => "Workbook",
            // Not "Worksheets": the collection reports itself as `Sheets`
            // whether it was reached through `.Worksheets` or `.Sheets`.
            ObjRef::Worksheets => "Sheets",
            ObjRef::Worksheet(_) => "Worksheet",
            ObjRef::Range(_) => "Range",
            ObjRef::Interior(_) => "Interior",
            ObjRef::Font(_) => "Font",
            ObjRef::ListObjects(_) => "ListObjects",
            ObjRef::ListObject(_) => "ListObject",
            ObjRef::ListColumns(_) => "ListColumns",
            ObjRef::ListColumn(..) => "ListColumn",
            ObjRef::ListRows(_) => "ListRows",
            ObjRef::ListRow(..) => "ListRow",
            ObjRef::PivotTables(_) => "PivotTables",
            ObjRef::PivotTable(_) => "PivotTable",
            ObjRef::PivotFields(_) => "PivotFields",
            ObjRef::PivotField(..) => "PivotField",
            ObjRef::UserClass(_) => "Object",
        }
    }

    /// `Is`: reference identity, not value equality.
    pub fn same_object(&self, other: &ObjRef) -> bool {
        match (self, other) {
            (ObjRef::Nothing, ObjRef::Nothing)
            | (ObjRef::Application, ObjRef::Application)
            | (ObjRef::WorksheetFunction, ObjRef::WorksheetFunction)
            | (ObjRef::Workbook, ObjRef::Workbook)
            | (ObjRef::Worksheets, ObjRef::Worksheets) => true,
            (ObjRef::Worksheet(a), ObjRef::Worksheet(b)) => a == b,
            // The handle *is* the identity, so a range that moved is still
            // the same object and two ranges over the same cells are not.
            (ObjRef::Range(a), ObjRef::Range(b)) => a == b,
            (ObjRef::UserClass(a), ObjRef::UserClass(b)) => a == b,
            (ObjRef::Interior(a), ObjRef::Interior(b)) => a == b,
            (ObjRef::Font(a), ObjRef::Font(b)) => a == b,
            // A table is identified by its id, so it stays the same object
            // across a rename -- the same reasoning as a worksheet.
            (ObjRef::ListObjects(a), ObjRef::ListObjects(b))
            | (ObjRef::ListObject(a), ObjRef::ListObject(b))
            | (ObjRef::ListColumns(a), ObjRef::ListColumns(b))
            | (ObjRef::ListRows(a), ObjRef::ListRows(b)) => a == b,
            (ObjRef::ListColumn(a, i), ObjRef::ListColumn(b, j))
            | (ObjRef::ListRow(a, i), ObjRef::ListRow(b, j))
            | (ObjRef::PivotField(a, i), ObjRef::PivotField(b, j)) => a == b && i == j,
            (ObjRef::PivotTables(a), ObjRef::PivotTables(b))
            | (ObjRef::PivotTable(a), ObjRef::PivotTable(b))
            | (ObjRef::PivotFields(a), ObjRef::PivotFields(b)) => a == b,
            _ => false,
        }
    }
}

/// Whether a bare name belongs to the host object model.
///
/// Consulted when there is *no* workbook attached, so that
/// `Range("A1")` in a host-free run reports "this needs a workbook" rather
/// than "Sub or Function not defined" -- which would be true but useless, and
/// would let a typo and a missing workbook look identical.
pub fn is_host_name(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "thisworkbook"
            | "activeworkbook"
            | "application"
            | "worksheets"
            | "sheets"
            | "range"
            | "cells"
    )
}

/// Error 438, naming the construct that is out of scope.
fn unsupported(what: &str) -> VbaError {
    VbaError::new(
        438,
        format!("Object doesn't support this property or method: {what}"),
    )
}

/// What `PivotField.CurrentPage` reads as when the field is unfiltered --
/// and, measured, also when *several* items are selected. It only ever
/// reflects a single selection.
const ALL_PAGES: &str = "(All)";

/// Error 1004 -- what Excel reports for a bad address, an out-of-sheet
/// `Offset`, and a `WorksheetFunction` call that fails. All measured.
fn app_defined(message: impl Into<String>) -> VbaError {
    VbaError::new(1004, message.into())
}

/// Reaching a member through a `Range` whose cells were all deleted.
///
/// The message is Excel's, verbatim. The *number* is not: see
/// [`RangeState::Dead`] for why Excel's is not reproducible and why 1004 is
/// what visi raises instead.
fn dead_range(member: &str) -> VbaError {
    VbaError::new(1004, format!("Method '{member}' of object 'Range' failed"))
}

/// The names `Sheet::evaluate_function` implements that are *not* Excel
/// functions.
///
/// `WorksheetFunction` must not expose them: a macro calling
/// `WorksheetFunction.Slice(...)` would work here and fail in Excel, which is
/// the one direction of divergence a differential harness cannot catch (it
/// generates what Excel accepts).
const ENGINE_ONLY_FUNCTIONS: &[&str] = &["GET", "GET_COL", "GET_COL_IDX", "SLICE", "STR"];

/// The workbook a macro is running against.
///
/// Holds the workbook mutably for the whole run, which is why every object is
/// a plain value: nothing else may hold a `&Sheet` across a statement.
pub struct Host<'w> {
    wb: &'w mut WorkbookManager,
    /// A write happened and no recalculation has run since. The next read
    /// that could observe it pays for one.
    stale: bool,
    /// Whether this run changed the workbook at all, which is what decides
    /// whether the caller has something worth saving.
    mutated: bool,
    /// Every `Range` handed out this run, by handle.
    ///
    /// The location lives here rather than in the [`ObjRef`] so a structural
    /// edit can move it, which is what makes `Set r = ws.Range("A5")` read
    /// `$A$6` after a row is inserted above it -- and makes a copy of `r`
    /// taken before the edit read `$A$6` too. Measured; see the module doc.
    ///
    /// Grows for the lifetime of one macro run and is dropped with the
    /// [`Host`]. A loop constructing a range per iteration therefore
    /// accumulates entries, which is bounded by the run rather than by the
    /// workbook.
    ranges: HashMap<u64, RangeState>,
    /// Next [`ObjRef::Range`] handle. A counter, not a hash of the
    /// coordinates -- two ranges over the same cells must not be the same
    /// object.
    next_token: u64,
    /// The sheet an unqualified `Range(...)` / `Cells(...)` resolves against.
    active_sheet: u64,
    /// Whether application events are enabled. Defaults to true.
    pub enable_events: bool,
    /// Pending cell range mutations for event dispatch.
    pub pending_cell_changes: Vec<RangeRef>,
    /// Pending sheet recalculations for calculate event dispatch.
    pub pending_calculate_sheets: Vec<u64>,
}

impl<'w> Host<'w> {
    /// Binds a workbook for the duration of a run.
    ///
    /// The first sheet is the active one, since nothing in the supported
    /// surface can change the selection.
    pub fn new(wb: &'w mut WorkbookManager) -> VResult<Self> {
        let active_sheet = wb
            .sheets
            .first()
            .map(|s| s.id)
            .ok_or_else(|| app_defined("The workbook has no worksheets"))?;
        Ok(Self {
            wb,
            stale: false,
            mutated: false,
            ranges: HashMap::new(),
            next_token: 1,
            active_sheet,
            enable_events: true,
            pending_cell_changes: Vec::new(),
            pending_calculate_sheets: Vec::new(),
        })
    }

    /// Whether the run changed anything in the workbook.
    pub fn mutated(&self) -> bool {
        self.mutated
    }

    /// Settles any outstanding recalculation, so a workbook about to be saved
    /// holds the values a reader would have seen.
    pub fn finish(&mut self) {
        self.recalculate();
    }

    fn recalculate(&mut self) {
        if self.stale {
            let _ = self.wb.evaluate();
            self.stale = false;
            if self.enable_events {
                for s in &self.wb.sheets {
                    self.pending_calculate_sheets.push(s.id);
                }
            }
        }
    }

    fn sheet_index(&self, id: u64) -> VResult<usize> {
        self.wb
            .sheets
            .iter()
            .position(|s| s.id == id)
            // A sheet can only vanish mid-run if something deleted it, which
            // is not in scope -- but "subscript out of range" is the right
            // report if it ever does.
            .ok_or_else(VbaError::subscript)
    }

    fn sheet(&self, id: u64) -> VResult<&Sheet> {
        Ok(&self.wb.sheets[self.sheet_index(id)?])
    }

    /// Creates a new range handle.
    pub fn new_range(
        &mut self,
        sheet_id: u64,
        row: u32,
        col: u32,
        height: u32,
        width: u32,
    ) -> ObjRef {
        self.next_token += 1;
        self.ranges.insert(
            self.next_token,
            RangeState::Live(RangeRef {
                sheet_id,
                row,
                col,
                height,
                width,
            }),
        );
        ObjRef::Range(self.next_token)
    }

    /// Where the range behind a handle currently points.
    ///
    /// `name` is the member being reached through it, purely so a dead range
    /// reports the same `Method '<name>' of object 'Range' failed` Excel
    /// does. An unknown handle cannot happen -- handles are only minted by
    /// [`Host::new_range`] and never removed -- but is reported rather than
    /// panicking, since a `Variant` holding one crosses the interpreter.
    fn range(&self, token: u64, name: &str) -> VResult<RangeRef> {
        match self.ranges.get(&token).copied() {
            Some(RangeState::Live(r)) => Ok(r),
            Some(RangeState::Dead) => Err(dead_range(name)),
            None => Err(dead_range(name)),
        }
    }

    // -- entry points the interpreter calls -------------------------------

    /// A bare identifier that names a host object, or `None` if it does not.
    pub fn global(&mut self, name: &str) -> Option<ObjRef> {
        Some(match name.to_ascii_lowercase().as_str() {
            // One workbook is open, so `ActiveWorkbook` is `ThisWorkbook`.
            "thisworkbook" | "activeworkbook" => ObjRef::Workbook,
            "application" => ObjRef::Application,
            "worksheets" | "sheets" => ObjRef::Worksheets,
            _ => return None,
        })
    }

    /// A call to a bare name that belongs to the host -- `Range("A1")`,
    /// `Cells(2, 3)`, `Worksheets(1)` -- resolved against the active sheet.
    ///
    /// Returns `None` for a name the host does not own, so the interpreter
    /// can go on to report "Sub or Function not defined" itself.
    pub fn global_call(&mut self, name: &str, args: &[Variant]) -> Option<VResult<Variant>> {
        let lower = name.to_ascii_lowercase();
        let obj = match lower.as_str() {
            "range" | "cells" => ObjRef::Worksheet(self.active_sheet),
            "worksheets" | "sheets" => ObjRef::Worksheets,
            _ => return None,
        };
        let member = match lower.as_str() {
            "worksheets" | "sheets" => return Some(self.call_object(&ObjRef::Worksheets, args)),
            other => other,
        };
        Some(self.get_member(&obj, member, args))
    }

    /// Reads a property, or calls a method, on an object.
    pub fn get_member(&mut self, obj: &ObjRef, name: &str, args: &[Variant]) -> VResult<Variant> {
        match obj {
            ObjRef::Nothing => Err(VbaError::new(
                91,
                format!("Object variable or With block variable not set: .{name}"),
            )),
            ObjRef::Application => self.application_member(name, args),
            ObjRef::WorksheetFunction => self.worksheet_function(name, args, true),
            ObjRef::Workbook => self.workbook_member(name, args),
            ObjRef::Worksheets => self.worksheets_member(name, args),
            ObjRef::Worksheet(id) => self.worksheet_member(*id, name, args),
            ObjRef::Range(token) => self.range_member(*token, name, args),
            ObjRef::Interior(token) => self.interior_member(*token, name),
            ObjRef::Font(token) => self.font_member(*token, name),
            ObjRef::ListObjects(sheet_id) => self.list_objects_member(*sheet_id, name, args),
            ObjRef::ListObject(id) => self.list_object_member(*id, name, args),
            ObjRef::ListColumns(id) => self.list_columns_member(*id, name, args),
            ObjRef::ListColumn(id, idx) => self.list_column_member(*id, *idx, name),
            ObjRef::ListRows(id) => self.list_rows_member(*id, name, args),
            ObjRef::ListRow(id, idx) => self.list_row_member(*id, *idx, name),
            ObjRef::PivotTables(sheet_id) => self.pivot_tables_member(*sheet_id, name, args),
            ObjRef::PivotTable(id) => self.pivot_table_member(*id, name, args),
            ObjRef::PivotFields(id) => self.pivot_fields_member(*id, name, args),
            ObjRef::PivotField(id, idx) => self.pivot_field_member(*id, *idx, name),
            ObjRef::UserClass(_) => Err(unsupported("user class member access on host")),
        }
    }

    /// Writes a property on an object.
    pub fn set_member(
        &mut self,
        obj: &ObjRef,
        name: &str,
        args: &[Variant],
        value: &Variant,
    ) -> VResult<()> {
        match (obj, name.to_ascii_lowercase().as_str()) {
            (ObjRef::Application, "enableevents") => {
                self.enable_events = value.to_bool()?;
                Ok(())
            }
            // `ListObject.Name` is not a field write: names are unique
            // workbook-wide and a rename cascades into formula *text*
            // everywhere. Routing through `WorkbookManager` is what keeps
            // `Sales[Amount]` pointing at the renamed table -- measured, the
            // formula really does change. A name already in use is error
            // 1004, also measured, and notably *not* the 9 that a failed
            // lookup gives.
            (ObjRef::ListObject(id), "name") => {
                let (_, t) = self.table(*id)?;
                let new_name = value.to_vba_string()?;
                self.wb
                    .rename_table(&t.name, &new_name)
                    .map_err(|e| app_defined(e.to_string()))?;
                self.mutated = true;
                self.stale = false;
                Ok(())
            }
            (ObjRef::ListObject(id), "showtotals") => {
                let (sheet_idx, t) = self.table(*id)?;
                let on = value.to_bool()?;
                if on == t.has_totals_row {
                    return Ok(());
                }
                if on {
                    // The totals row is a new row at the bottom of the
                    // table's own columns, so it shifts what is under it.
                    self.wb
                        .insert_cells_shift_down(
                            sheet_idx,
                            t.end_row + 1,
                            t.start_col,
                            t.end_col,
                            1,
                        )
                        .map_err(|e| app_defined(e.to_string()))?;
                } else {
                    self.wb
                        .delete_cells_shift_up(sheet_idx, t.end_row, t.start_col, t.end_col, 1)
                        .map_err(|e| app_defined(e.to_string()))?;
                }
                let (sheet_idx, _) = self.table(*id)?;
                if let Some(table) = self.wb.sheets[sheet_idx]
                    .tables
                    .iter_mut()
                    .find(|t| t.id == *id)
                {
                    table.has_totals_row = on;
                    table.end_row = if on {
                        t.end_row + 1
                    } else {
                        t.end_row.saturating_sub(1)
                    };
                }
                self.mutated = true;
                self.stale = false;
                Ok(())
            }
            // Same cascade, for a column: renaming one rewrites every
            // structured reference that names it.
            (ObjRef::ListColumn(id, idx), "name") => {
                let (_, t) = self.table(*id)?;
                let new_name = value.to_vba_string()?;
                self.wb
                    .rename_table_column(&t.name, *idx as usize, &new_name)
                    .map_err(|e| app_defined(e.to_string()))?;
                self.mutated = true;
                self.stale = false;
                Ok(())
            }
            // Measured: assigning re-renders the grid **immediately**, with
            // no `RefreshTable` -- reading a pivot cell straight afterwards
            // shows the filtered value. That is a deliberate exception to
            // this crate's rule that nothing recomputes a pivot implicitly;
            // the rule describes visi's CRUD, and here Excel's behaviour
            // wins because a macro can observe the difference.
            (ObjRef::PivotField(id, idx), "currentpage") => {
                let (_, p) = self.pivot(*id)?;
                let columns = self.pivot_source_columns(&p)?;
                let column = columns
                    .get(*idx as usize)
                    .cloned()
                    .ok_or_else(|| app_defined("Unable to set the CurrentPage property"))?;
                if !p
                    .filter_fields
                    .iter()
                    .any(|f| f.column.eq_ignore_ascii_case(&column))
                {
                    return Err(app_defined("Unable to set the CurrentPage property"));
                }
                let wanted = value.to_vba_string()?;
                let selection = if wanted == ALL_PAGES {
                    None
                } else {
                    // Measured: a value that is not one of the field's items
                    // is 1004 rather than an empty selection, which would
                    // silently blank the grid.
                    if !self.pivot_field_has_item(&p, &column, &wanted)? {
                        return Err(app_defined(
                            "Unable to set the CurrentPage property of the PivotField class",
                        ));
                    }
                    Some(vec![wanted])
                };
                self.wb
                    .set_pivot_filter(&p.name, &column, selection)
                    .map_err(|e| app_defined(e.to_string()))?;
                // `CurrentPage` *is* the single-select page mode, which is
                // what makes the page-field cell show the item's own name
                // rather than `(Multiple Items)`.
                if let Some(pivot) = self.wb.pivot_tables.iter_mut().find(|t| t.id == *id)
                    && let Some(f) = pivot
                        .filter_fields
                        .iter_mut()
                        .find(|f| f.column.eq_ignore_ascii_case(&column))
                {
                    f.multiple_selection = false;
                }
                self.wb
                    .refresh_pivot_table(&p.name)
                    .map_err(|e| app_defined(e.to_string()))?;
                self.mutated = true;
                self.stale = false;
                Ok(())
            }
            (ObjRef::Interior(_) | ObjRef::Font(_), _) => {
                if !args.is_empty() {
                    return Err(unsupported(&format!(
                        "{}.{name} with arguments",
                        obj.type_name()
                    )));
                }
                self.style_set(obj, name, value)
            }
            (ObjRef::Range(token), "numberformat") => {
                let r = self.range(*token, name)?;
                let format = value.to_vba_string()?;
                // Measured: setting `General` on a date cell leaves the
                // serial alone -- the value stays a number and only the
                // rendering changes, which is exactly `core::date`'s model.
                let stored =
                    (!format.eq_ignore_ascii_case(color::GENERAL_FORMAT)).then_some(format);
                self.style_write(r, move |s| s.num_format = stored.clone())
            }
            (ObjRef::Range(token), "value" | "value2" | "formula") => {
                if !args.is_empty() {
                    return Err(unsupported(&format!("Range.{name} with arguments")));
                }
                let r = self.range(*token, name)?;
                self.write_range(r, value)
            }
            (ObjRef::Worksheet(id), "name") => {
                let new_name = value.to_vba_string()?;
                let idx = self.sheet_index(*id)?;
                let old = self.wb.sheets[idx].name.clone();
                if old == new_name {
                    return Ok(());
                }
                self.wb
                    .rename_sheet(&old, &new_name)
                    .map_err(|e| app_defined(e.to_string()))?;
                self.mutated = true;
                // A rename rewrites cross-sheet formula text, so anything
                // read afterwards must see the rebuilt references.
                self.stale = true;
                Ok(())
            }
            (ObjRef::Nothing, _) => Err(VbaError::new(
                91,
                format!("Object variable or With block variable not set: .{name}"),
            )),
            _ => Err(unsupported(&format!("{}.{name} =", obj.type_name()))),
        }
    }

    /// Calls an object as if it were its own default member: `Worksheets(1)`,
    /// `ws.Cells(2, 3)`.
    pub fn call_object(&mut self, obj: &ObjRef, args: &[Variant]) -> VResult<Variant> {
        match obj {
            ObjRef::Worksheets => {
                let key = args.first().ok_or_else(VbaError::subscript)?;
                Ok(Variant::Object(self.worksheet_by_key(key)?))
            }
            // `ws.Cells` is a Range over the whole grid, so `ws.Cells(2, 3)`
            // is that range indexed -- which is exactly Excel's own model and
            // gives `$C$2` without a second code path.
            ObjRef::Range(token) => {
                let r = self.range(*token, "Item")?;
                self.range_index(r, args)
            }
            ObjRef::ListObjects(sheet_id) => {
                let key = args.first().ok_or_else(VbaError::subscript)?;
                Ok(Variant::Object(self.table_by_key(*sheet_id, key)?))
            }
            ObjRef::ListColumns(id) => {
                let key = args.first().ok_or_else(VbaError::subscript)?;
                Ok(Variant::Object(self.list_column_by_key(*id, key)?))
            }
            ObjRef::ListRows(id) => {
                let key = args.first().ok_or_else(VbaError::subscript)?;
                Ok(Variant::Object(self.list_row_by_index(*id, key)?))
            }
            ObjRef::PivotTables(sheet_id) => {
                let key = args
                    .first()
                    .ok_or_else(|| app_defined("PivotTables needs a key"))?;
                Ok(Variant::Object(self.pivot_by_key(*sheet_id, key)?))
            }
            ObjRef::PivotFields(id) => {
                let key = args
                    .first()
                    .ok_or_else(|| app_defined("PivotFields needs a key"))?;
                Ok(Variant::Object(self.pivot_field_by_key(*id, key)?))
            }
            _ => Err(unsupported(&format!("calling a {}", obj.type_name()))),
        }
    }

    /// The value an object stands for when it is used without `Set`.
    ///
    /// `x = ws.Range("A1")` reads the cell, and `MsgBox ws.Range("A1")` would
    /// too. Only `Range` has a default member in this scope.
    pub fn default_value(&mut self, obj: &ObjRef) -> VResult<Variant> {
        match obj {
            ObjRef::Range(token) => {
                let r = self.range(*token, "Value")?;
                self.read_range(r, false)
            }
            ObjRef::Nothing => Err(VbaError::new(
                91,
                "Object variable or With block variable not set",
            )),
            other => Err(unsupported(&format!(
                "using a {} as a value",
                other.type_name()
            ))),
        }
    }

    /// Assigning to an object without `Set`, which writes its default member.
    pub fn assign_default(&mut self, obj: &ObjRef, value: &Variant) -> VResult<()> {
        match obj {
            ObjRef::Range(token) => {
                let r = self.range(*token, "Value")?;
                self.write_range(r, value)
            }
            other => Err(unsupported(&format!(
                "assigning to a {}",
                other.type_name()
            ))),
        }
    }

    /// The elements `For Each` walks.
    ///
    /// Measured: a `Range` iterates one cell at a time in **row-major** order
    /// (`A1 B1 A2 B2` over `A1:B2`), and `Worksheets` iterates in workbook
    /// order.
    pub fn iterate(&mut self, obj: &ObjRef) -> VResult<Vec<Variant>> {
        match obj {
            ObjRef::Range(token) => {
                let r = self.range(*token, "Item")?;
                if r.count() > MAX_ALLOCATED_CELLS {
                    return Err(VbaError::new(
                        7,
                        "Out of memory: For Each over a range this large",
                    ));
                }
                let mut out = Vec::with_capacity(r.count() as usize);
                for dr in 0..r.height {
                    for dc in 0..r.width {
                        out.push(Variant::Object(self.new_range(
                            r.sheet_id,
                            r.row + dr,
                            r.col + dc,
                            1,
                            1,
                        )));
                    }
                }
                Ok(out)
            }
            ObjRef::Worksheets => Ok(self
                .wb
                .sheets
                .iter()
                .map(|s| Variant::Object(ObjRef::Worksheet(s.id)))
                .collect()),
            other => Err(unsupported(&format!(
                "For Each over a {}",
                other.type_name()
            ))),
        }
    }

    // -- per-object members -----------------------------------------------

    fn workbook_member(&mut self, name: &str, args: &[Variant]) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "worksheets" | "sheets" => {
                if args.is_empty() {
                    Ok(Variant::Object(ObjRef::Worksheets))
                } else {
                    self.call_object(&ObjRef::Worksheets, args)
                }
            }
            // The workbook has no filename until something saves it, and this
            // layer has none to offer -- the CLI decides the output path, and
            // `visi-core` never sees it. Reporting the sheet-less placeholder
            // Excel uses for an unsaved book is the honest answer.
            "name" => Ok(Variant::Str("Book1".to_string())),
            // `.Save` is in scope as a no-op with a real meaning: the CLI
            // writes the file when the run finishes, so a macro asking for a
            // save gets one. What it must not do is fail.
            "save" => Ok(Variant::Empty),
            _ => Err(unsupported(&format!("Workbook.{name}"))),
        }
    }

    fn worksheets_member(&mut self, name: &str, args: &[Variant]) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Variant::Long(self.wb.sheets.len() as i32)),
            "item" => self.call_object(&ObjRef::Worksheets, args),
            _ => Err(unsupported(&format!("Sheets.{name}"))),
        }
    }

    /// `Worksheets(x)`, where `x` is a 1-based index or a name.
    ///
    /// Measured: both a missing name and an out-of-range index are error 9,
    /// and the name match is case-insensitive.
    fn worksheet_by_key(&mut self, key: &Variant) -> VResult<ObjRef> {
        if let Variant::Str(name) = key {
            return self
                .wb
                .sheets
                .iter()
                .find(|s| s.name.eq_ignore_ascii_case(name))
                .map(|s| ObjRef::Worksheet(s.id))
                .ok_or_else(VbaError::subscript);
        }
        let idx = key.to_f64()?;
        if idx < 1.0 || idx > self.wb.sheets.len() as f64 {
            return Err(VbaError::subscript());
        }
        Ok(ObjRef::Worksheet(self.wb.sheets[idx as usize - 1].id))
    }

    fn worksheet_member(&mut self, id: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "name" => Ok(Variant::Str(self.sheet(id)?.name.clone())),
            "range" => {
                let obj = self.resolve_range_args(id, args)?;
                Ok(Variant::Object(obj))
            }
            "cells" => {
                let whole = self.new_range(id, 0, 0, MAX_ROWS, MAX_COLS);
                if args.is_empty() {
                    return Ok(Variant::Object(whole));
                }
                self.call_object(&whole, args)
            }
            // `ws.Rows(3)` is a whole-row Range -- measured:
            // `TypeName(ws.Rows(3))` is `"Range"` and its `.Address` is
            // `$3:$3`, which `format_address` already renders from a
            // full-width rectangle.
            "listobjects" => {
                if args.is_empty() {
                    return Ok(Variant::Object(ObjRef::ListObjects(id)));
                }
                self.call_object(&ObjRef::ListObjects(id), args)
            }
            "pivottables" => {
                if args.is_empty() {
                    return Ok(Variant::Object(ObjRef::PivotTables(id)));
                }
                self.call_object(&ObjRef::PivotTables(id), args)
            }
            "rows" => {
                let (at, count) = band_args(args, MAX_ROWS, parse_row)?;
                Ok(Variant::Object(self.new_range(id, at, 0, count, MAX_COLS)))
            }
            "columns" => {
                let (at, count) = band_args(args, MAX_COLS, parse_col)?;
                Ok(Variant::Object(self.new_range(id, 0, at, MAX_ROWS, count)))
            }
            _ => Err(unsupported(&format!("Worksheet.{name}"))),
        }
    }

    /// `Range("A1")`, `Range("A1:B2")` and `Range(cell1, cell2)`.
    fn resolve_range_args(&mut self, sheet_id: u64, args: &[Variant]) -> VResult<ObjRef> {
        let first = args.first().ok_or_else(VbaError::invalid_call)?;
        if args.len() >= 2 {
            let a = self.corner(first)?;
            let b = self.corner(&args[1])?;
            let (row, col) = (a.0.min(b.0), a.1.min(b.1));
            let (row2, col2) = (a.0.max(b.0), a.1.max(b.1));
            return Ok(self.new_range(sheet_id, row, col, row2 - row + 1, col2 - col + 1));
        }
        match first {
            Variant::Object(ObjRef::Range(token)) => {
                let r = self.range(*token, "Range")?;
                Ok(self.new_range(sheet_id, r.row, r.col, r.height, r.width))
            }
            other => {
                let text = other.to_vba_string()?;
                let (row, col, height, width) = parse_address(&text).ok_or_else(|| {
                    app_defined(format!("Method 'Range' of object failed: {text}"))
                })?;
                Ok(self.new_range(sheet_id, row, col, height, width))
            }
        }
    }

    /// One corner of a two-argument `Range(cell1, cell2)`, as `(row, col)`.
    ///
    /// Takes no sheet: an address is a sheet-free coordinate, and a `Range`
    /// passed as a corner contributes only its top-left. `Range(a, b)` on one
    /// worksheet with a corner from *another* therefore lands on the first,
    /// which is what Excel does too.
    fn corner(&self, v: &Variant) -> VResult<(u32, u32)> {
        match v {
            Variant::Object(ObjRef::Range(token)) => {
                let r = self.range(*token, "Range")?;
                Ok((r.row, r.col))
            }
            other => {
                let text = other.to_vba_string()?;
                let (row, col, _, _) = parse_address(&text).ok_or_else(|| {
                    app_defined(format!("Method 'Range' of object failed: {text}"))
                })?;
                Ok((row, col))
            }
        }
    }

    /// Takes the handle rather than the rectangle so that a dead range names
    /// the member the macro actually reached for, exactly as Excel's
    /// `Method '<name>' of object 'Range' failed` does.
    fn range_member(&mut self, token: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        let r = self.range(token, name)?;
        match name.to_ascii_lowercase().as_str() {
            "value" => self.read_range(r, false),
            "value2" => self.read_range(r, true),
            "formula" => self.read_formula(r),
            "text" => {
                self.recalculate();
                let sheet = self.sheet(r.sheet_id)?;
                let cell = CellRef::new(r.row as usize, r.col as usize);
                // Measured: `.Text` on a cell holding `=1/0` is `#DIV/0!`,
                // the string Excel puts on screen -- not `ResultData`'s
                // human-readable `Error: #DIV/0!` rendering, which is for a
                // terminal rather than for a cell.
                Ok(Variant::Str(match sheet.get_result_data(&cell) {
                    ResultData::Error(e) => e,
                    _ => sheet.get_display_string(&cell),
                }))
            }
            // `.Row` and `.Column` report the *top-left* cell of a
            // multi-cell range: measured, `ws.Range("B2:C4").Row` is 2.
            "row" => Ok(Variant::Long(r.row as i32 + 1)),
            "column" => Ok(Variant::Long(r.col as i32 + 1)),
            "count" => {
                let n = r.count();
                if n > i32::MAX as u64 {
                    // Measured: `ws.Cells.Count` really is error 6 in Excel,
                    // because the property is typed `Long` and the whole grid
                    // does not fit in one.
                    return Err(VbaError::overflow());
                }
                Ok(Variant::Long(n as i32))
            }
            "address" => {
                let row_abs = match args.first() {
                    Some(v) => v.to_bool()?,
                    None => true,
                };
                let col_abs = match args.get(1) {
                    Some(v) => v.to_bool()?,
                    None => true,
                };
                Ok(Variant::Str(format_address(&r, row_abs, col_abs)))
            }
            "offset" => {
                let dr = int_arg(args, 0)?;
                let dc = int_arg(args, 1)?;
                let row = i64::from(r.row) + dr;
                let col = i64::from(r.col) + dc;
                if row < 0
                    || col < 0
                    || row + i64::from(r.height) > i64::from(MAX_ROWS)
                    || col + i64::from(r.width) > i64::from(MAX_COLS)
                {
                    return Err(app_defined("Offset moves the range off the worksheet"));
                }
                Ok(Variant::Object(self.new_range(
                    r.sheet_id, row as u32, col as u32, r.height, r.width,
                )))
            }
            "resize" => {
                // An omitted dimension keeps the current one, which is what
                // makes `.Resize(, 3)` mean "same rows, three columns".
                let height = match args.first() {
                    Some(v) if !v.is_empty() => positive_dim(v)?,
                    _ => r.height,
                };
                let width = match args.get(1) {
                    Some(v) if !v.is_empty() => positive_dim(v)?,
                    _ => r.width,
                };
                if u64::from(r.row) + u64::from(height) > u64::from(MAX_ROWS)
                    || u64::from(r.col) + u64::from(width) > u64::from(MAX_COLS)
                {
                    return Err(app_defined("Resize extends the range off the worksheet"));
                }
                Ok(Variant::Object(
                    self.new_range(r.sheet_id, r.row, r.col, height, width),
                ))
            }
            "cells" => {
                if args.is_empty() {
                    return Ok(Variant::Object(
                        self.new_range(r.sheet_id, r.row, r.col, r.height, r.width),
                    ));
                }
                self.range_index(r, args)
            }
            // Measured: `ws.Range("B5").EntireRow.Address` is `$5:$5` and
            // `.EntireColumn` is `$B:$B`, so each widens one axis to the
            // whole grid and leaves the other alone.
            // Measured: the table containing the range's top-left, or
            // `Nothing` for a cell outside every table.
            "listobject" => Ok(Variant::Object(
                self.table_at(r.sheet_id, r.row, r.col)
                    .map(ObjRef::ListObject)
                    .unwrap_or(ObjRef::Nothing),
            )),
            // A view onto the same cells, riding on this range's handle, so
            // it tracks a structural edit exactly as the range does.
            "interior" => Ok(Variant::Object(ObjRef::Interior(token))),
            "font" => Ok(Variant::Object(ObjRef::Font(token))),
            // Measured: `General` on a cell carrying no format, and `Null`
            // over a range whose cells disagree.
            "numberformat" => self.style_fold(
                r,
                Variant::Null,
                |s| {
                    s.and_then(|s| s.num_format.clone())
                        .unwrap_or_else(|| color::GENERAL_FORMAT.to_string())
                },
                Variant::Str,
            ),
            "entirerow" => Ok(Variant::Object(
                self.new_range(r.sheet_id, r.row, 0, r.height, MAX_COLS),
            )),
            "entirecolumn" => Ok(Variant::Object(
                self.new_range(r.sheet_id, 0, r.col, MAX_ROWS, r.width),
            )),
            "insert" | "delete" => {
                self.structural_edit(r, name.eq_ignore_ascii_case("insert"), name)?;
                // Excel's `Insert`/`Delete` do return a value, but a macro
                // calls them as statements and this is not measured, so
                // nothing is claimed about it.
                Ok(Variant::Empty)
            }
            // `name`, not the lowercased match binding: the message is
            // user-facing and `Range.interior` reads like a typo where
            // `Range.Interior` reads like the refusal it is.
            _ => Err(unsupported(&format!("Range.{name}"))),
        }
    }

    // -- Excel Tables (`ListObjects`) --------------------------------------

    /// The sheet index and a snapshot of the table with this id.
    ///
    /// Returns a copy for the same reason every host object is a value: the
    /// caller needs the workbook mutably a moment later. Tables are small
    /// (an extent plus its column names), and this is not a hot path.
    ///
    /// A table that no longer exists reports error 9, the same number every
    /// other "no such table" case does -- measured for the lookup cases.
    fn table(&self, id: u64) -> VResult<(usize, ExcelTable)> {
        self.wb
            .sheets
            .iter()
            .enumerate()
            .find_map(|(i, sheet)| {
                sheet
                    .tables
                    .iter()
                    .find(|t| t.id == id)
                    .map(|t| (i, t.clone()))
            })
            .ok_or_else(VbaError::subscript)
    }

    /// The id of the table covering a cell, or `None` if it is in no table.
    fn table_at(&self, sheet_id: u64, row: u32, col: u32) -> Option<u64> {
        let (row, col) = (row as usize, col as usize);
        self.wb
            .sheets
            .iter()
            .find(|s| s.id == sheet_id)?
            .tables
            .iter()
            .find(|t| {
                (t.start_row..=t.end_row).contains(&row) && (t.start_col..=t.end_col).contains(&col)
            })
            .map(|t| t.id)
    }

    /// `ListObjects(x)`, where `x` is a 1-based index or a name.
    ///
    /// Measured: a missing name and an out-of-range index are both error 9,
    /// and the name match is case-insensitive.
    fn table_by_key(&mut self, sheet_id: u64, key: &Variant) -> VResult<ObjRef> {
        let sheet = self.sheet(sheet_id)?;
        let table = match key {
            Variant::Str(name) => sheet
                .tables
                .iter()
                .find(|t| t.name.eq_ignore_ascii_case(name)),
            other => {
                let n = other.to_f64()?;
                let idx = crate::core::vba::value::bankers_round(n) as i64;
                if idx < 1 {
                    return Err(VbaError::subscript());
                }
                sheet.tables.get(idx as usize - 1)
            }
        };
        table
            .map(|t| ObjRef::ListObject(t.id))
            .ok_or_else(VbaError::subscript)
    }

    fn list_objects_member(
        &mut self,
        sheet_id: u64,
        name: &str,
        args: &[Variant],
    ) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Variant::Long(self.sheet(sheet_id)?.tables.len() as i32)),
            "item" => self.call_object(&ObjRef::ListObjects(sheet_id), args),
            _ => Err(unsupported(&format!("ListObjects.{name}"))),
        }
    }

    /// A `Range` over a table sub-rectangle, or `Nothing` when the part does
    /// not exist -- which is measured behaviour for all three of
    /// `HeaderRowRange` on a headerless table, `TotalsRowRange` without a
    /// totals row, and `DataBodyRange` on a table with zero data rows.
    fn table_part(&mut self, t: &ExcelTable, rows: Option<(usize, usize)>) -> Variant {
        match rows {
            Some((first, last)) if first <= last => Variant::Object(self.new_range(
                t.sheet_id,
                first as u32,
                t.start_col as u32,
                (last - first + 1) as u32,
                t.col_count() as u32,
            )),
            _ => Variant::Object(ObjRef::Nothing),
        }
    }

    fn list_object_member(&mut self, id: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        let (_, t) = self.table(id)?;
        match name.to_ascii_lowercase().as_str() {
            "name" => Ok(Variant::Str(t.name.clone())),
            "range" => Ok(self.table_part(&t, Some((t.start_row, t.end_row)))),
            "headerrowrange" => {
                let rows = t.header_row().map(|r| (r, r));
                Ok(self.table_part(&t, rows))
            }
            // Measured: `Nothing` when the table has no data rows, which is
            // *not* the same as an empty range -- see `ExcelTable::has_insert_row`.
            "databodyrange" => {
                let rows = (t.data_row_count() > 0).then(|| (t.data_start_row(), t.data_end_row()));
                Ok(self.table_part(&t, rows))
            }
            "totalsrowrange" => {
                let rows = t.totals_row().map(|r| (r, r));
                Ok(self.table_part(&t, rows))
            }
            "showtotals" => Ok(Variant::Boolean(t.has_totals_row)),
            "showheaders" => Ok(Variant::Boolean(t.has_header_row)),
            "listcolumns" => {
                if args.is_empty() {
                    return Ok(Variant::Object(ObjRef::ListColumns(id)));
                }
                self.call_object(&ObjRef::ListColumns(id), args)
            }
            "listrows" => {
                if args.is_empty() {
                    return Ok(Variant::Object(ObjRef::ListRows(id)));
                }
                self.call_object(&ObjRef::ListRows(id), args)
            }
            _ => Err(unsupported(&format!("ListObject.{name}"))),
        }
    }

    fn list_columns_member(&mut self, id: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Variant::Long(self.table(id)?.1.col_count() as i32)),
            "item" => self.call_object(&ObjRef::ListColumns(id), args),
            _ => Err(unsupported(&format!("ListColumns.{name}"))),
        }
    }

    fn list_column_by_key(&mut self, id: u64, key: &Variant) -> VResult<ObjRef> {
        let (_, t) = self.table(id)?;
        let idx = match key {
            Variant::Str(name) => t.local_column_index(name).ok_or_else(VbaError::subscript)?,
            other => {
                let n = crate::core::vba::value::bankers_round(other.to_f64()?) as i64;
                if n < 1 || n as usize > t.col_count() {
                    return Err(VbaError::subscript());
                }
                n as usize - 1
            }
        };
        Ok(ObjRef::ListColumn(id, idx as u32))
    }

    fn list_column_member(&mut self, id: u64, idx: u32, name: &str) -> VResult<Variant> {
        let (_, t) = self.table(id)?;
        let idx = idx as usize;
        if idx >= t.col_count() {
            return Err(VbaError::subscript());
        }
        let col = (t.start_col + idx) as u32;
        match name.to_ascii_lowercase().as_str() {
            "name" => Ok(Variant::Str(
                t.columns.get(idx).cloned().unwrap_or_default(),
            )),
            "index" => Ok(Variant::Long(idx as i32 + 1)),
            // Measured: `.Range` on a column includes the header row.
            "range" => Ok(Variant::Object(self.new_range(
                t.sheet_id,
                t.start_row as u32,
                col,
                t.row_count() as u32,
                1,
            ))),
            "databodyrange" => {
                if t.data_row_count() == 0 {
                    return Ok(Variant::Object(ObjRef::Nothing));
                }
                Ok(Variant::Object(self.new_range(
                    t.sheet_id,
                    t.data_start_row() as u32,
                    col,
                    t.data_row_count() as u32,
                    1,
                )))
            }
            _ => Err(unsupported(&format!("ListColumn.{name}"))),
        }
    }

    fn list_rows_member(&mut self, id: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Variant::Long(self.table(id)?.1.data_row_count() as i32)),
            "item" => self.call_object(&ObjRef::ListRows(id), args),
            "add" => self.list_rows_add(id, args),
            _ => Err(unsupported(&format!("ListRows.{name}"))),
        }
    }

    fn list_row_by_index(&mut self, id: u64, key: &Variant) -> VResult<ObjRef> {
        let (_, t) = self.table(id)?;
        let n = crate::core::vba::value::bankers_round(key.to_f64()?) as i64;
        if n < 1 || n as usize > t.data_row_count() {
            return Err(VbaError::subscript());
        }
        Ok(ObjRef::ListRow(id, n as u32 - 1))
    }

    fn list_row_member(&mut self, id: u64, idx: u32, name: &str) -> VResult<Variant> {
        let (sheet_idx, t) = self.table(id)?;
        if idx as usize >= t.data_row_count() {
            return Err(VbaError::subscript());
        }
        let row = t.data_start_row() + idx as usize;
        match name.to_ascii_lowercase().as_str() {
            "index" => Ok(Variant::Long(idx as i32 + 1)),
            "range" => Ok(Variant::Object(self.new_range(
                t.sheet_id,
                row as u32,
                t.start_col as u32,
                1,
                t.col_count() as u32,
            ))),
            "delete" => {
                // Deleting a table row shifts only the table's own columns
                // up, exactly as adding one shifts them down.
                self.wb
                    .delete_cells_shift_up(sheet_idx, row, t.start_col, t.end_col, 1)
                    .map_err(|e| app_defined(e.to_string()))?;
                self.resize_table_rows(id, t.end_row, -1)?;
                self.mutated = true;
                self.stale = false;
                Ok(Variant::Empty)
            }
            _ => Err(unsupported(&format!("ListRow.{name}"))),
        }
    }

    /// `ListRows.Add([Position])`, which appends by default.
    ///
    /// Measured: the new row is blank, the table grows by one row, and the
    /// returned object is a `ListRow` pointing at it. A table sitting on its
    /// insert-row placeholder already has the row reserved, so that case only
    /// clears the flag -- which is why `.Add` on an emptied table leaves the
    /// extent at `A1:C2` rather than growing it to `A1:C3`.
    fn list_rows_add(&mut self, id: u64, args: &[Variant]) -> VResult<Variant> {
        let (sheet_idx, t) = self.table(id)?;
        let position = match args.first().filter(|v| !v.is_empty()) {
            Some(v) => {
                let n = crate::core::vba::value::bankers_round(v.to_f64()?) as i64;
                if n < 1 || n as usize > t.data_row_count() + 1 {
                    return Err(VbaError::subscript());
                }
                n as usize - 1
            }
            None => t.data_row_count(),
        };

        if t.has_insert_row {
            // The placeholder row is already there and already blank.
            if let Some(table) = self.wb.sheets[sheet_idx]
                .tables
                .iter_mut()
                .find(|t| t.id == id)
            {
                table.has_insert_row = false;
            }
        } else {
            let row = t.data_start_row() + position;
            self.wb
                .insert_cells_shift_down(sheet_idx, row, t.start_col, t.end_col, 1)
                .map_err(|e| app_defined(e.to_string()))?;
            self.resize_table_rows(id, t.end_row, 1)?;
        }
        self.mutated = true;
        self.stale = false;
        Ok(Variant::Object(ObjRef::ListRow(id, position as u32)))
    }

    /// Sets a table's bottom edge to `old_end_row + delta`, or marks it as
    /// sitting on its insert row when that would leave it with no data.
    ///
    /// Assigns absolutely rather than adding, because
    /// `WorkbookManager::apply_grid_edit` may already have grown the extent:
    /// an insert *inside* the table moves its bottom edge, an insert just
    /// past it does not, and the caller should not have to know which.
    fn resize_table_rows(&mut self, id: u64, old_end_row: usize, delta: isize) -> VResult<()> {
        let (sheet_idx, _) = self.table(id)?;
        let Some(table) = self.wb.sheets[sheet_idx]
            .tables
            .iter_mut()
            .find(|t| t.id == id)
        else {
            return Err(VbaError::subscript());
        };
        table.end_row = old_end_row.saturating_add_signed(delta);
        // Removing the last data row leaves the extent alone and raises the
        // insert-row flag, which is what Excel writes to the file.
        if delta < 0 && table.data_row_count() == 0 && !table.has_insert_row {
            table.end_row = old_end_row;
            table.has_insert_row = true;
        }
        Ok(())
    }

    // -- Pivot tables ------------------------------------------------------
    //
    // Every failure in this surface is **1004**, not the 9 the `ListObjects`
    // collection uses for the same shape of mistake. Measured, and the two
    // are easy to conflate.

    /// Whether a pivot field actually has an item with this value, matched
    /// the way the pivot engine merges them (case-insensitively).
    fn pivot_field_has_item(&self, p: &PivotTable, column: &str, wanted: &str) -> VResult<bool> {
        let sheets: Vec<&Sheet> = self.wb.sheets.iter().collect();
        let (src, names, cols, rows) =
            crate::core::pivot::resolve_source(&sheets, &p.source).map_err(app_defined)?;
        let idx = crate::core::pivot::column_index(&names, column).map_err(app_defined)?;
        Ok(rows.iter().any(|&r| {
            crate::core::pivot::group_key(&src.get_result_data(&CellRef::new(r, cols[idx])))
                .eq_ignore_ascii_case(wanted)
        }))
    }

    /// A snapshot of the pivot with this id, and its index.
    fn pivot(&self, id: u64) -> VResult<(usize, PivotTable)> {
        self.wb
            .pivot_tables
            .iter()
            .position(|p| p.id == id)
            .map(|i| (i, self.wb.pivot_tables[i].clone()))
            .ok_or_else(|| app_defined("The pivot table no longer exists"))
    }

    /// The source column names of a pivot, which is what `PivotFields`
    /// enumerates -- one entry per source column, whatever area it occupies.
    fn pivot_source_columns(&self, pivot: &PivotTable) -> VResult<Vec<String>> {
        let sheets: Vec<&Sheet> = self.wb.sheets.iter().collect();
        crate::core::pivot::resolve_source(&sheets, &pivot.source)
            .map(|(_, names, _, _)| names)
            .map_err(app_defined)
    }

    fn pivot_by_key(&mut self, sheet_id: u64, key: &Variant) -> VResult<ObjRef> {
        // Excel scopes `PivotTables` to a worksheet by where the pivot is
        // *drawn*, which is `dest_sheet_id` here -- a pivot's source may be
        // on another sheet entirely.
        let on_sheet: Vec<&PivotTable> = self
            .wb
            .pivot_tables
            .iter()
            .filter(|p| p.dest_sheet_id == sheet_id)
            .collect();
        let found = match key {
            Variant::Str(name) => on_sheet.iter().find(|p| p.name.eq_ignore_ascii_case(name)),
            other => {
                let n = crate::core::vba::value::bankers_round(other.to_f64()?) as i64;
                if n < 1 {
                    return Err(app_defined("PivotTables index must be 1 or more"));
                }
                on_sheet.get(n as usize - 1)
            }
        };
        found
            .map(|p| ObjRef::PivotTable(p.id))
            .ok_or_else(|| app_defined("Unable to get the PivotTables property"))
    }

    fn pivot_tables_member(
        &mut self,
        sheet_id: u64,
        name: &str,
        args: &[Variant],
    ) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "count" => Ok(Variant::Long(
                self.wb
                    .pivot_tables
                    .iter()
                    .filter(|p| p.dest_sheet_id == sheet_id)
                    .count() as i32,
            )),
            "item" => self.call_object(&ObjRef::PivotTables(sheet_id), args),
            _ => Err(unsupported(&format!("PivotTables.{name}"))),
        }
    }

    fn pivot_table_member(&mut self, id: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        let (_, p) = self.pivot(id)?;
        match name.to_ascii_lowercase().as_str() {
            "name" => Ok(Variant::Str(p.name.clone())),
            // Measured: returns the Boolean True.
            "refreshtable" => {
                self.wb
                    .refresh_pivot_table(&p.name)
                    .map_err(|e| app_defined(e.to_string()))?;
                self.mutated = true;
                self.stale = false;
                Ok(Variant::Boolean(true))
            }
            "pivotfields" => {
                if args.is_empty() {
                    return Ok(Variant::Object(ObjRef::PivotFields(id)));
                }
                self.call_object(&ObjRef::PivotFields(id), args)
            }
            // Measured: `TableRange1` is the grid alone and `TableRange2`
            // includes the page-field rows above it -- `$F$3:$G$7` and
            // `$F$1:$G$7` for the same pivot at `F1` with one filter.
            "tablerange1" | "tablerange2" => {
                let (end_row, end_col) = self.pivot_extent(&p)?;
                let with_pages = name.eq_ignore_ascii_case("tablerange2");
                let offset = if with_pages || p.filter_fields.is_empty() {
                    0
                } else {
                    // One row per filter field plus the blank one under them,
                    // the same reservation `pivot_xlsx` makes on export.
                    p.filter_fields.len() + 1
                };
                let top = p.dest_row + offset;
                if top > end_row {
                    return Err(app_defined("The pivot table has no output yet"));
                }
                Ok(Variant::Object(self.new_range(
                    p.dest_sheet_id,
                    top as u32,
                    p.dest_col as u32,
                    (end_row - top + 1) as u32,
                    (end_col - p.dest_col + 1) as u32,
                )))
            }
            _ => Err(unsupported(&format!("PivotTable.{name}"))),
        }
    }

    /// The pivot's rendered bottom-right corner.
    ///
    /// Prefers the extent the last refresh recorded, and computes one
    /// otherwise so that reading a range is a read -- refreshing here would
    /// make a property access mutate the workbook.
    fn pivot_extent(&self, p: &PivotTable) -> VResult<(usize, usize)> {
        if let (Some(r), Some(c)) = (p.last_output_end_row, p.last_output_end_col) {
            return Ok((r, c));
        }
        let sheets: Vec<&Sheet> = self.wb.sheets.iter().collect();
        let grid = crate::core::pivot::compute_pivot(&sheets, p).map_err(app_defined)?;
        let rows = grid.filter_rows.len()
            + usize::from(!grid.filter_rows.is_empty())
            + grid.header_rows.len()
            + grid.body_rows.len();
        Ok((
            p.dest_row + rows.saturating_sub(1),
            p.dest_col + grid.width.saturating_sub(1),
        ))
    }

    fn pivot_fields_member(&mut self, id: u64, name: &str, args: &[Variant]) -> VResult<Variant> {
        match name.to_ascii_lowercase().as_str() {
            "count" => {
                let (_, p) = self.pivot(id)?;
                Ok(Variant::Long(self.pivot_source_columns(&p)?.len() as i32))
            }
            "item" => self.call_object(&ObjRef::PivotFields(id), args),
            _ => Err(unsupported(&format!("PivotFields.{name}"))),
        }
    }

    fn pivot_field_by_key(&mut self, id: u64, key: &Variant) -> VResult<ObjRef> {
        let (_, p) = self.pivot(id)?;
        let columns = self.pivot_source_columns(&p)?;
        let idx = match key {
            Variant::Str(name) => columns
                .iter()
                .position(|c| c.eq_ignore_ascii_case(name))
                .ok_or_else(|| app_defined("Unable to get the PivotFields property"))?,
            other => {
                let n = crate::core::vba::value::bankers_round(other.to_f64()?) as i64;
                if n < 1 || n as usize > columns.len() {
                    return Err(app_defined("Unable to get the PivotFields property"));
                }
                n as usize - 1
            }
        };
        Ok(ObjRef::PivotField(id, idx as u32))
    }

    fn pivot_field_member(&mut self, id: u64, idx: u32, name: &str) -> VResult<Variant> {
        let (_, p) = self.pivot(id)?;
        let columns = self.pivot_source_columns(&p)?;
        let column = columns
            .get(idx as usize)
            .cloned()
            .ok_or_else(|| app_defined("Unable to get the PivotFields property"))?;
        let matches = |c: &String| c.eq_ignore_ascii_case(&column);
        match name.to_ascii_lowercase().as_str() {
            "name" => Ok(Variant::Str(column)),
            // xlRowField 1, xlColumnField 2, xlPageField 3, xlHidden 0.
            // Measured: a *data* field reports 0, not xlDataField -- the
            // orientation belongs to the source field, and aggregating one
            // leaves it unoriented.
            "orientation" => Ok(Variant::Long(
                if p.row_fields.iter().any(|f| matches(&f.column)) {
                    1
                } else if p.col_fields.iter().any(|f| matches(&f.column)) {
                    2
                } else if p.filter_fields.iter().any(|f| matches(&f.column)) {
                    3
                } else {
                    0
                },
            )),
            // Measured: `(All)` when nothing is filtered *and* when several
            // items are selected -- `CurrentPage` only ever reflects a single
            // selection. On a field that is not a page field it raises.
            "currentpage" => {
                let field = p
                    .filter_fields
                    .iter()
                    .find(|f| matches(&f.column))
                    .ok_or_else(|| app_defined("Unable to get the CurrentPage property"))?;
                Ok(Variant::Str(match &field.selected_values {
                    Some(values) if values.len() == 1 => values[0].clone(),
                    _ => ALL_PAGES.to_string(),
                }))
            }
            _ => Err(unsupported(&format!("PivotField.{name}"))),
        }
    }

    /// A style attribute read over a range, or `Null` where the cells
    /// disagree.
    ///
    /// Measured, and the asymmetry is Excel's: `Font.Bold`, `Font.Size`,
    /// `Font.Name`, `NumberFormat` and `Interior.ColorIndex` all report
    /// `Null` over a range whose cells differ, but `Interior.Color` reports
    /// **0**. Callers pass `mixed` for that one case rather than it being
    /// inferred, because 0 is also a legitimate uniform value (black).
    fn style_fold<T: PartialEq>(
        &mut self,
        r: RangeRef,
        mixed: Variant,
        read: impl Fn(Option<&CellStyle>) -> T,
        wrap: impl Fn(T) -> Variant,
    ) -> VResult<Variant> {
        let sheet = self.sheet(r.sheet_id)?;
        let mut seen: Option<T> = None;
        for row in r.row..r.row.saturating_add(r.height).min(MAX_ROWS) {
            for col in r.col..r.col.saturating_add(r.width).min(MAX_COLS) {
                let value = read(sheet.get_cell_style(row as usize, col as usize));
                match &seen {
                    None => seen = Some(value),
                    Some(first) if *first == value => {}
                    Some(_) => return Ok(mixed),
                }
            }
        }
        Ok(seen.map(wrap).unwrap_or(Variant::Empty))
    }

    /// Applies a style change to every cell of a range.
    fn style_write(&mut self, r: RangeRef, edit: impl Fn(&mut CellStyle)) -> VResult<()> {
        if r.count() > MAX_ALLOCATED_CELLS {
            return Err(VbaError::new(7, "Out of memory: range too large to style"));
        }
        let idx = self.sheet_index(r.sheet_id)?;
        let sheet = &mut self.wb.sheets[idx];
        for row in r.row..r.row.saturating_add(r.height).min(MAX_ROWS) {
            for col in r.col..r.col.saturating_add(r.width).min(MAX_COLS) {
                sheet.update_cell_style(row as usize, col as usize, &edit);
            }
        }
        self.mutated = true;
        Ok(())
    }

    fn interior_member(&mut self, token: u64, name: &str) -> VResult<Variant> {
        let r = self.range(token, name)?;
        match name.to_ascii_lowercase().as_str() {
            // Measured: an unfilled cell reports white (16777215), and a
            // range whose fills differ reports 0 rather than Null -- the one
            // property that breaks the Null rule.
            "color" => self.style_fold(r, Variant::Double(0.0), color::interior_color, |c| {
                Variant::Double(f64::from(c))
            }),
            // Measured: an *unfilled* cell is `xlNone` (-4142), but a fill
            // that is not one of the 56 slots reports the **nearest** slot
            // rather than `xlNone` -- `RGB(250, 10, 10)` reports 3.
            "colorindex" => self.style_fold(
                r,
                Variant::Null,
                |s| {
                    s.and_then(|s| s.bg_color.as_deref())
                        .and_then(color::nearest_color_index)
                        .unwrap_or(color::COLOR_INDEX_NONE)
                },
                Variant::Long,
            ),
            _ => Err(unsupported(&format!("Interior.{name}"))),
        }
    }

    fn font_member(&mut self, token: u64, name: &str) -> VResult<Variant> {
        let r = self.range(token, name)?;
        match name.to_ascii_lowercase().as_str() {
            "bold" => self.style_fold(
                r,
                Variant::Null,
                |s| s.and_then(|s| s.bold).unwrap_or(false),
                Variant::Boolean,
            ),
            "italic" => self.style_fold(
                r,
                Variant::Null,
                |s| s.and_then(|s| s.italic).unwrap_or(false),
                Variant::Boolean,
            ),
            // Measured: `Font.Size` is a `Double`, not a `Long`, and a
            // half-point size round-trips -- which is why `CellStyle`'s is an
            // `f64`.
            "size" => self.style_fold(
                r,
                Variant::Null,
                |s| {
                    s.and_then(|s| s.font_size)
                        .unwrap_or(color::DEFAULT_FONT_SIZE)
                        .to_bits()
                },
                |bits| Variant::Double(f64::from_bits(bits)),
            ),
            "name" => self.style_fold(
                r,
                Variant::Null,
                |s| {
                    s.and_then(|s| s.font_family.clone())
                        .unwrap_or_else(|| color::DEFAULT_FONT_NAME.to_string())
                },
                Variant::Str,
            ),
            "color" => self.style_fold(r, Variant::Null, color::font_color, |c| {
                Variant::Double(f64::from(c))
            }),
            // Measured: a cell with no explicit font colour reports slot 1
            // (black), not `xlNone` -- the opposite of `Interior.ColorIndex`.
            "colorindex" => self.style_fold(
                r,
                Variant::Null,
                |s| {
                    s.and_then(|s| s.font_color.as_deref())
                        .and_then(color::nearest_color_index)
                        .unwrap_or(color::FONT_COLOR_INDEX_AUTOMATIC)
                },
                Variant::Long,
            ),
            _ => Err(unsupported(&format!("Font.{name}"))),
        }
    }

    /// Writing `Interior.X` / `Font.X`, which is where the BGR conversion
    /// actually happens.
    fn style_set(&mut self, obj: &ObjRef, name: &str, value: &Variant) -> VResult<()> {
        let (token, on_font) = match obj {
            ObjRef::Interior(t) => (*t, false),
            ObjRef::Font(t) => (*t, true),
            _ => unreachable!("style_set is only reached for Interior and Font"),
        };
        let r = self.range(token, name)?;
        let lower = name.to_ascii_lowercase();

        // `Color` is a BGR `Long`; `CellStyle` stores `"#RRGGBB"`.
        if lower == "color" {
            let hex = color::bgr_to_hex(long_arg(value)?);
            return self.style_write(r, move |s| {
                if on_font {
                    s.font_color = Some(hex.clone());
                } else {
                    s.bg_color = Some(hex.clone());
                }
            });
        }
        if lower == "colorindex" {
            let index = long_arg(value)?;
            // Measured: setting `xlNone` clears the fill, after which
            // `Interior.Color` reads white again.
            let hex = if index == color::COLOR_INDEX_NONE {
                None
            } else {
                let slot = usize::try_from(index)
                    .ok()
                    .filter(|i| (1..=color::COLOR_INDEX_PALETTE.len()).contains(i))
                    .ok_or_else(|| app_defined("ColorIndex is out of range"))?;
                Some(color::COLOR_INDEX_PALETTE[slot - 1].to_string())
            };
            return self.style_write(r, move |s| {
                if on_font {
                    s.font_color = hex.clone();
                } else {
                    s.bg_color = hex.clone();
                }
            });
        }
        if !on_font {
            return Err(unsupported(&format!("Interior.{name}")));
        }
        match lower.as_str() {
            "bold" => {
                let on = value.to_bool()?;
                self.style_write(r, move |s| s.bold = Some(on))
            }
            "italic" => {
                let on = value.to_bool()?;
                self.style_write(r, move |s| s.italic = Some(on))
            }
            "underline" => {
                let on = value.to_bool()?;
                self.style_write(r, move |s| s.underline = Some(on))
            }
            "size" => {
                let size = value.to_f64()?;
                if !size.is_finite() || size <= 0.0 {
                    return Err(app_defined("Font.Size must be a positive number"));
                }
                self.style_write(r, move |s| s.font_size = Some(size))
            }
            "name" => {
                let family = value.to_vba_string()?;
                self.style_write(r, move |s| s.font_family = Some(family.clone()))
            }
            _ => Err(unsupported(&format!("Font.{name}"))),
        }
    }

    /// `Rows(n).Insert` / `.Delete` and the column equivalents.
    ///
    /// Only a whole-row or whole-column band is accepted. Excel *does* accept
    /// a partial range and picks the shift direction from its shape -- and
    /// measured, `Range("A2:A3").Insert` shifts **right**, not down, which is
    /// the opposite of what the obvious reading suggests. Guessing at that
    /// silently moves a macro's data sideways, so it is refused instead.
    fn structural_edit(&mut self, r: RangeRef, insert: bool, member: &str) -> VResult<()> {
        let (axis, at, count) = if r.width == MAX_COLS {
            (Axis::Row, r.row, r.height)
        } else if r.height == MAX_ROWS {
            (Axis::Col, r.col, r.width)
        } else {
            return Err(unsupported(&format!(
                "Range.{member} on a range that is not a whole row or column"
            )));
        };

        let sheet_idx = self.sheet_index(r.sheet_id)?;
        let (rows, cols) = {
            let sheet = &self.wb.sheets[sheet_idx];
            (sheet.row_count(), sheet.col_count())
        };

        // How many the *grid* has to change by. Deleting past the end of the
        // allocated grid is a no-op there but still moves anything below it,
        // so the clamp applies to the workbook ops and not to the shift.
        let grid_count = if insert {
            // A dense column vector per sheet column means inserting a
            // million rows really would allocate them, where Excel's sparse
            // grid would shrug. Same guard, and same error 7, as a write to
            // a far-off cell.
            let added = u64::from(count)
                * match axis {
                    Axis::Row => cols as u64,
                    Axis::Col => rows as u64,
                };
            if added > MAX_ALLOCATED_CELLS {
                return Err(VbaError::new(
                    7,
                    format!("Out of memory: Range.{member} would allocate {added} cells"),
                ));
            }
            count
        } else {
            let allocated = match axis {
                Axis::Row => rows as u32,
                Axis::Col => cols as u32,
            };
            count.min(allocated.saturating_sub(at))
        };

        for _ in 0..grid_count {
            // Each call rewrites the formulas in the whole workbook and
            // re-evaluates, so the result is the same as one wider edit; see
            // `WorkbookManager::apply_grid_edit`.
            let done = match (axis, insert) {
                (Axis::Row, true) => self.wb.insert_row(sheet_idx, at as usize),
                (Axis::Row, false) => self.wb.delete_row(sheet_idx, at as usize),
                (Axis::Col, true) => self.wb.insert_col(sheet_idx, at as usize),
                (Axis::Col, false) => self.wb.delete_col(sheet_idx, at as usize),
            };
            if done.is_err() {
                // Only `OutOfBounds`, which the clamp above already means we
                // do not expect; nothing is left to remove either way.
                break;
            }
        }

        self.shift_ranges(&GridEdit {
            sheet_id: r.sheet_id,
            axis,
            at: at as usize,
            count: count as usize,
            insert,
            band: None,
        });

        self.mutated = true;
        // `insert_row` and friends re-evaluate the workbook themselves, so
        // nothing is outstanding once they return.
        self.stale = false;
        Ok(())
    }

    /// Moves every live `Range` this run has handed out.
    ///
    /// This is what makes a `Range` track the edit the way Excel's does, and
    /// it deliberately reuses [`shift_span`] rather than reimplementing the
    /// geometry: `fuzz/vba_range_tracking_probe.py` found Excel's rules for a
    /// `Range` object to be the same ones, case for case, that it applies to
    /// a formula's range reference.
    fn shift_ranges(&mut self, edit: &GridEdit) {
        for state in self.ranges.values_mut() {
            let RangeState::Live(r) = state else {
                continue;
            };
            if r.sheet_id != edit.sheet_id {
                continue;
            }
            let (start, len) = match edit.axis {
                Axis::Row => (r.row, r.height),
                Axis::Col => (r.col, r.width),
            };
            // A whole-row or whole-column band already spans the axis it is
            // unbounded on, so an edit along that axis cannot move it -- and
            // `start + len - 1` would be the grid maximum, which shifting
            // would push past the end.
            let unbounded = match edit.axis {
                Axis::Row => r.height == MAX_ROWS,
                Axis::Col => r.width == MAX_COLS,
            };
            if unbounded {
                continue;
            }
            let end = start as usize + len as usize - 1;
            match shift_span(start as usize, end, edit.at, edit.count, edit.insert) {
                Some((new_start, new_end)) => {
                    let new_len = (new_end - new_start + 1) as u32;
                    match edit.axis {
                        Axis::Row => {
                            r.row = new_start as u32;
                            r.height = new_len;
                        }
                        Axis::Col => {
                            r.col = new_start as u32;
                            r.width = new_len;
                        }
                    }
                }
                // Every cell it covered is gone. Measured: the object stays a
                // `Range` and is not `Nothing`, but every member now raises.
                None => *state = RangeState::Dead,
            }
        }
    }

    /// `range(r, c)`, 1-based and relative to the range's own top-left.
    ///
    /// Excel lets this run outside the range -- `ws.Cells(2, 3)` works
    /// because `ws.Cells` starts at `A1` -- so only the sheet bounds apply.
    fn range_index(&mut self, r: RangeRef, args: &[Variant]) -> VResult<Variant> {
        let row = int_arg(args, 0)?;
        let col = if args.len() > 1 { int_arg(args, 1)? } else { 1 };
        if row < 1 || col < 1 {
            return Err(VbaError::subscript());
        }
        let row = i64::from(r.row) + row - 1;
        let col = i64::from(r.col) + col - 1;
        if row >= i64::from(MAX_ROWS) || col >= i64::from(MAX_COLS) {
            return Err(VbaError::subscript());
        }
        Ok(Variant::Object(
            self.new_range(r.sheet_id, row as u32, col as u32, 1, 1),
        ))
    }

    // -- reading and writing cells ----------------------------------------

    /// `.Value` / `.Value2`.
    ///
    /// A single cell reads as a scalar; anything larger reads as a 2-D array
    /// indexed `(row, column)` from 1, which is what makes
    /// `v = ws.Range("A1:B2").Value` legal and `CStr(v)` error 13.
    fn read_range(&mut self, r: RangeRef, value2: bool) -> VResult<Variant> {
        self.recalculate();
        if r.is_single() {
            let sheet = self.sheet(r.sheet_id)?;
            return Ok(cell_value(sheet, r.row, r.col, value2));
        }
        if r.count() > MAX_ALLOCATED_CELLS {
            return Err(VbaError::new(7, "Out of memory: range too large to read"));
        }
        let sheet = self.sheet(r.sheet_id)?;
        let mut values = Vec::with_capacity(r.count() as usize);
        for dr in 0..r.height {
            for dc in 0..r.width {
                values.push(cell_value(sheet, r.row + dr, r.col + dc, value2));
            }
        }
        Ok(Variant::Array(std::rc::Rc::new(VarArray {
            rows: r.height as usize,
            cols: r.width as usize,
            values,
        })))
    }

    /// `.Formula`, which is the cell's source text -- `"=A1*2"` for a
    /// formula, `"hi"` for text, `""` for an empty cell. All measured.
    fn read_formula(&mut self, r: RangeRef) -> VResult<Variant> {
        self.recalculate();
        let sheet = self.sheet(r.sheet_id)?;
        if r.is_single() {
            return Ok(Variant::Str(cell_formula(sheet, r.row, r.col)));
        }
        if r.count() > MAX_ALLOCATED_CELLS {
            return Err(VbaError::new(7, "Out of memory: range too large to read"));
        }
        let mut values = Vec::with_capacity(r.count() as usize);
        for dr in 0..r.height {
            for dc in 0..r.width {
                values.push(Variant::Str(cell_formula(sheet, r.row + dr, r.col + dc)));
            }
        }
        Ok(Variant::Array(std::rc::Rc::new(VarArray {
            rows: r.height as usize,
            cols: r.width as usize,
            values,
        })))
    }

    /// Writes one value across every cell of a range.
    ///
    /// Measured: assigning to a multi-cell range fills all of it, and
    /// assigning an *array* to a single cell writes only its first element.
    fn write_range(&mut self, r: RangeRef, value: &Variant) -> VResult<()> {
        let value = match value {
            Variant::Object(o) => {
                let o = *o;
                &self.default_value(&o)?
            }
            other => other,
        };
        let scalar = match value {
            Variant::Array(a) => a.get(1, 1)?,
            other => other.clone(),
        };
        let src = cell_src(&scalar)?;
        let date_format = match &scalar {
            // Measured: `.Value = #6/22/2026#` leaves the cell formatted
            // `m/d/yy`, so writing a Date writes the notation too -- which is
            // exactly how `core/date.rs` models a date in the first place.
            Variant::Date(serial) => Some(if serial.fract() == 0.0 {
                "m/d/yy"
            } else {
                "m/d/yy h:mm"
            }),
            _ => None,
        };

        let last_row = u64::from(r.row) + u64::from(r.height) - 1;
        let last_col = u64::from(r.col) + u64::from(r.width) - 1;
        if (last_row + 1) * (last_col + 1) > MAX_ALLOCATED_CELLS {
            return Err(VbaError::new(
                7,
                format!(
                    "Out of memory: writing {} would allocate more than {} cells",
                    format_address(&r, true, true),
                    MAX_ALLOCATED_CELLS
                ),
            ));
        }

        let idx = self.sheet_index(r.sheet_id)?;
        let sheet = &mut self.wb.sheets[idx];
        sheet.ensure_capacity(last_row as usize, last_col as usize);
        for dr in 0..r.height {
            for dc in 0..r.width {
                let (row, col) = ((r.row + dr) as usize, (r.col + dc) as usize);
                sheet.set_cell_src(row, col, src.clone());
                if let Some(code) = date_format {
                    sheet.update_cell_style(row, col, |s| {
                        s.num_format = Some(code.to_string());
                    });
                }
            }
        }
        self.mutated = true;
        self.stale = true;
        if self.enable_events {
            self.pending_cell_changes.push(r);
        }
        Ok(())
    }

    // -- the WorksheetFunction bridge --------------------------------------

    /// `Application.<name>`, which is either an object or a non-raising
    /// worksheet function.
    fn application_member(&mut self, name: &str, args: &[Variant]) -> VResult<Variant> {
        if name.eq_ignore_ascii_case("worksheetfunction") {
            return Ok(Variant::Object(ObjRef::WorksheetFunction));
        }
        if name.eq_ignore_ascii_case("enableevents") {
            return Ok(Variant::Boolean(self.enable_events));
        }
        // `Application.Sum(...)` works and returns 6, exactly as
        // `WorksheetFunction.Sum` does -- the two only part company on
        // failure. Measured.
        self.worksheet_function(name, args, false)
    }

    /// Calls an Excel worksheet function through the engine's own
    /// implementation.
    ///
    /// `raises` is the whole difference between the two call paths:
    /// `WorksheetFunction.VLookup` failing is a trappable error 1004, while
    /// `Application.VLookup` failing returns an error `Variant` that
    /// `IsError` detects. Both measured.
    fn worksheet_function(
        &mut self,
        name: &str,
        args: &[Variant],
        raises: bool,
    ) -> VResult<Variant> {
        let upper = name.to_uppercase();
        if ENGINE_ONLY_FUNCTIONS.contains(&upper.as_str()) {
            return Err(unsupported(&format!("WorksheetFunction.{name}")));
        }
        self.recalculate();

        let mut exprs = Vec::with_capacity(args.len());
        for a in args {
            // An error value handed in directly short-circuits, matching what
            // the engine does with an error inside a range.
            if let Some(n) = a.error_number() {
                return self.function_failure(raises, name, error_text(n));
            }
            exprs.push(self.arg_expr(a)?);
        }

        match self.wb.call_worksheet_function(&upper, &exprs) {
            Ok(ResultData::Error(e)) => self.function_failure(raises, name, &e),
            Ok(value) => Ok(result_to_variant(&value)),
            // The engine reports an unknown name as an evaluation failure.
            // Excel makes that a *compile* error for `WorksheetFunction`,
            // which nothing here can reproduce, so 438 -- the late-bound
            // equivalent, and the number every other out-of-scope construct
            // reports -- names it rather than inventing a 1004.
            Err(_) => Err(unsupported(&format!("WorksheetFunction.{name}"))),
        }
    }

    fn function_failure(&self, raises: bool, name: &str, error: &str) -> VResult<Variant> {
        if raises {
            Err(app_defined(format!(
                "Unable to get the {name} property of the WorksheetFunction class: {error}"
            )))
        } else {
            Ok(Variant::ErrValue(error_code(error)))
        }
    }

    /// One VBA argument as the formula AST the engine evaluates.
    ///
    /// A `Range` becomes a real range reference so the function sees cells --
    /// which is what makes `Sum` skip the text and booleans inside one, while
    /// `Sum("1", 2)` coerces the string. Both measured, and both fall out of
    /// reusing the formula path rather than being coded twice.
    fn arg_expr(&mut self, v: &Variant) -> VResult<FExpr> {
        Ok(match v {
            Variant::Object(ObjRef::Range(token)) => {
                let r = self.range(*token, "Value")?;
                let sheet = self.sheet(r.sheet_id)?.name.clone();
                FExpr::RangeRef {
                    sheet: Some(sheet),
                    start_row: r.row as usize,
                    start_col: r.col as usize,
                    end_row: (r.row + r.height - 1) as usize,
                    end_col: (r.col + r.width - 1) as usize,
                    start_row_abs: true,
                    start_col_abs: true,
                    end_row_abs: true,
                    end_col_abs: true,
                }
            }
            Variant::Object(o) => {
                return Err(unsupported(&format!(
                    "a {} as a worksheet function argument",
                    o.type_name()
                )));
            }
            Variant::Array(a) => FExpr::List(
                a.values
                    .iter()
                    .map(scalar_expr)
                    .collect::<VResult<Vec<_>>>()?,
            ),
            other => scalar_expr(other)?,
        })
    }
}

/// A scalar VBA value as a formula literal.
fn scalar_expr(v: &Variant) -> VResult<FExpr> {
    Ok(match v {
        Variant::Boolean(b) => FExpr::Boolean(*b),
        Variant::Str(s) => FExpr::String(s.clone()),
        // Excel has no blank literal in a formula, and a blank argument is
        // treated as zero by every function that accepts one.
        Variant::Empty | Variant::Null => FExpr::Number(0.0),
        other => FExpr::Number(other.to_f64()?),
    })
}

/// One argument as a whole number, for `Offset`/`Resize`/`Cells`.
fn int_arg(args: &[Variant], i: usize) -> VResult<i64> {
    let v = args.get(i).cloned().unwrap_or(Variant::Empty);
    let f = v.to_f64()?;
    if !f.is_finite() || f.abs() > i64::MAX as f64 {
        return Err(VbaError::overflow());
    }
    Ok(crate::core::vba::value::bankers_round(f) as i64)
}

/// A colour or palette index written to a style property.
///
/// Excel's colour properties are typed `Long`, so a fractional value rounds
/// rather than erroring; `int_arg`'s banker's rounding is the conversion the
/// rest of the interpreter uses.
fn long_arg(v: &Variant) -> VResult<i32> {
    let n = int_arg(std::slice::from_ref(v), 0)?;
    i32::try_from(n).map_err(|_| VbaError::overflow())
}

/// A `Resize` dimension, which Excel rejects at zero or below.
fn positive_dim(v: &Variant) -> VResult<u32> {
    let f = crate::core::vba::value::bankers_round(v.to_f64()?);
    if f < 1.0 || f > f64::from(MAX_ROWS) {
        return Err(app_defined("Resize dimension must be at least 1"));
    }
    Ok(f as u32)
}

/// `$A$1` / `$A$1:$B$2`, and the whole-row form Excel uses for a range that
/// spans every column (`ws.Cells.Address` is `$1:$1048576`, measured).
/// The `(start, count)` a `Rows(...)` / `Columns(...)` argument selects,
/// 0-based.
///
/// Accepts the three spellings Excel does: nothing at all (the whole axis),
/// a 1-based number, and a string that is either one index or an `a:b` band
/// -- `ws.Columns("B")` and `ws.Rows("5:7")` both being ordinary VBA.
/// `parse_one` is what turns one side of that string into a 0-based index,
/// which is where rows and columns differ (`"5"` versus `"B"`).
fn band_args(
    args: &[Variant],
    axis_len: u32,
    parse_one: fn(&str) -> Option<u32>,
) -> VResult<(u32, u32)> {
    let Some(arg) = args.first().filter(|v| !v.is_empty()) else {
        return Ok((0, axis_len));
    };
    if let Variant::Str(text) = arg {
        let text = text.trim();
        let (first, last) = match text.split_once(':') {
            Some((a, b)) => (
                parse_one(a.trim()).ok_or_else(VbaError::subscript)?,
                parse_one(b.trim()).ok_or_else(VbaError::subscript)?,
            ),
            None => {
                let one = parse_one(text).ok_or_else(VbaError::subscript)?;
                (one, one)
            }
        };
        let (lo, hi) = (first.min(last), first.max(last));
        if hi >= axis_len {
            return Err(VbaError::subscript());
        }
        return Ok((lo, hi - lo + 1));
    }
    let n = int_arg(args, 0)?;
    if n < 1 || n > i64::from(axis_len) {
        return Err(VbaError::subscript());
    }
    Ok((n as u32 - 1, 1))
}

fn format_address(r: &RangeRef, row_abs: bool, col_abs: bool) -> String {
    let rd = if row_abs { "$" } else { "" };
    let cd = if col_abs { "$" } else { "" };
    let (r1, r2) = (r.row + 1, r.row + r.height);
    let (c1, c2) = (r.col, r.col + r.width - 1);
    if r.width == MAX_COLS {
        return format!("{rd}{r1}:{rd}{r2}");
    }
    if r.height == MAX_ROWS {
        return format!(
            "{cd}{}:{cd}{}",
            col_idx_to_letters(c1 as usize),
            col_idx_to_letters(c2 as usize)
        );
    }
    let start = format!("{cd}{}{rd}{r1}", col_idx_to_letters(c1 as usize));
    if r.is_single() {
        return start;
    }
    format!("{start}:{cd}{}{rd}{r2}", col_idx_to_letters(c2 as usize))
}

/// `"A1"`, `"A1:B2"`, `"A:B"`, `"3:5"` -> `(row, col, height, width)`, all
/// 0-based. `None` for anything else, which the caller turns into 1004.
fn parse_address(text: &str) -> Option<(u32, u32, u32, u32)> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    let (a, b) = match text.split_once(':') {
        Some((a, b)) => (a, Some(b)),
        None => (text, None),
    };
    // Whole columns (`A:B`) and whole rows (`3:5`) are the two forms where a
    // part carries only half a coordinate; both are all-or-nothing.
    if let Some(b) = b
        && let (Some(c1), Some(c2)) = (parse_col(a), parse_col(b))
    {
        let (lo, hi) = (c1.min(c2), c1.max(c2));
        return Some((0, lo, MAX_ROWS, hi - lo + 1));
    }
    if let Some(b) = b
        && let (Some(r1), Some(r2)) = (parse_row(a), parse_row(b))
    {
        let (lo, hi) = (r1.min(r2), r1.max(r2));
        return Some((lo, 0, hi - lo + 1, MAX_COLS));
    }
    let (r1, c1) = parse_cell(a)?;
    let Some(b) = b else {
        return Some((r1, c1, 1, 1));
    };
    let (r2, c2) = parse_cell(b)?;
    let (rlo, rhi) = (r1.min(r2), r1.max(r2));
    let (clo, chi) = (c1.min(c2), c1.max(c2));
    Some((rlo, clo, rhi - rlo + 1, chi - clo + 1))
}

fn parse_cell(s: &str) -> Option<(u32, u32)> {
    let s = s.trim().trim_start_matches('$');
    let split = s.find(|c: char| c.is_ascii_digit())?;
    let (letters, rest) = s.split_at(split);
    let col = parse_col(letters)?;
    let row = parse_row(rest)?;
    Some((row, col))
}

fn parse_col(s: &str) -> Option<u32> {
    let s = s.trim().trim_start_matches('$');
    if s.is_empty() || s.len() > 3 || !s.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let mut col: u32 = 0;
    for c in s.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    if col == 0 || col > MAX_COLS {
        return None;
    }
    Some(col - 1)
}

fn parse_row(s: &str) -> Option<u32> {
    let s = s.trim().trim_start_matches('$');
    if s.is_empty() || !s.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let row: u32 = s.parse().ok()?;
    if row == 0 || row > MAX_ROWS {
        return None;
    }
    Some(row - 1)
}

/// One cell as `.Value` or `.Value2` reports it.
///
/// Two measured rules do all the work here. A numeric cell always reads back
/// as a `Double`, never an `Integer` -- `TypeName(ws.Range("A1").Value)` is
/// `Double` for a cell holding `1`. And a cell whose style carries a date
/// number format reads back as a `Date` through `.Value` and a `Double`
/// through `.Value2`, which is the whole reason [`Variant::Date`] exists:
/// the engine stores only the serial (see `core/date.rs`), and the notation
/// lives on the cell.
fn cell_value(sheet: &Sheet, row: u32, col: u32, value2: bool) -> Variant {
    let cell = CellRef::new(row as usize, col as usize);
    let is_date = !value2
        && sheet
            .get_cell_style(row as usize, col as usize)
            .and_then(|s| s.num_format.as_deref())
            .is_some_and(crate::core::date::is_date_code);
    match sheet.get_result_data(&cell) {
        ResultData::None => Variant::Empty,
        ResultData::Boolean(b) => Variant::Boolean(b),
        ResultData::Integer(i) if is_date && i >= 0 => Variant::Date(i as f64),
        ResultData::Integer(i) => Variant::Double(i as f64),
        ResultData::Float(f) if is_date && f >= 0.0 => Variant::Date(f),
        ResultData::Float(f) => Variant::Double(f),
        ResultData::String(s) => Variant::Str(s),
        ResultData::Error(e) => Variant::ErrValue(error_code(&e)),
        // `List` and `Dict` are engine-internal shapes with no VBA analogue
        // (nothing Excel-compatible puts one in a cell). Rendering rather
        // than erroring keeps a macro that stumbles onto one readable.
        other => Variant::Str(other.to_string()),
    }
}

/// One cell as `.Formula` reports it: its source text.
///
/// Measured: a formula cell gives `"=A1*2"`, a text cell gives `"hi"` and an
/// empty cell gives `""`. The one adjustment is the quoting -- `visi` stores
/// a text cell whose content would otherwise re-parse as a number, a boolean
/// or a date in a quoted literal (`xlsx::text_cell_src`), and that quoting is
/// storage, not something a user ever typed. Excel's `.Formula` shows what
/// was typed, so it comes back off.
fn cell_formula(sheet: &Sheet, row: u32, col: u32) -> String {
    let cell = CellRef::new(row as usize, col as usize);
    let src = sheet.get_src_str(&cell);
    if src.starts_with('=') {
        return src;
    }
    if let Some(inner) = src.strip_prefix('"').and_then(|s| s.strip_suffix('"')) {
        return inner.to_string();
    }
    src
}

/// A `Variant` as the cell source text that reproduces it.
///
/// A string is written *verbatim*, which is not laziness: measured,
/// `.Value = "=G1*3"` really does make the cell a formula, and
/// `.Value = "6/22/2026"` really does make it a date. `Sheet::commit` already
/// parses a literal exactly the way Excel parses typed-in text, so handing it
/// the raw string reproduces Excel's behaviour instead of re-deriving it.
fn cell_src(v: &Variant) -> VResult<String> {
    Ok(match v {
        Variant::Empty | Variant::Null => String::new(),
        Variant::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        Variant::Str(s) => s.clone(),
        Variant::ErrValue(n) => error_text(*n).to_string(),
        Variant::Object(_) | Variant::Array(_) => return Err(VbaError::type_mismatch()),
        // Rust's shortest-round-trip float formatting, not VBA's `CStr`:
        // this text is re-parsed by `Sheet::commit`, so it has to survive the
        // round trip exactly rather than look like VBA.
        other => {
            let f = other.to_f64()?;
            // Excel has no infinities and no NaNs: assigning one stores
            // `#NUM!`. Measured -- `wsh.Cells(2, 5).Value = (-2.5 ^ va)` with
            // `va = 1000` leaves an error cell, not a number. (Between
            // *constants* the same expression is error 6 before the
            // assignment is ever reached, which is `value::ArithMode`'s doing
            // rather than this.)
            if !f.is_finite() {
                "#NUM!".to_string()
            } else if f == f.trunc() && f.abs() < 1e15 {
                format!("{}", f as i64)
            } else {
                format!("{f}")
            }
        }
    })
}

/// An engine result as a `Variant`.
///
/// No date handling: a worksheet function's result carries no cell and so no
/// number format, and `WorksheetFunction.Sum` over date cells is measured to
/// return a plain `Double` serial.
fn result_to_variant(v: &ResultData) -> Variant {
    match v {
        ResultData::None => Variant::Empty,
        ResultData::Boolean(b) => Variant::Boolean(*b),
        ResultData::Integer(i) => Variant::Double(*i as f64),
        ResultData::Float(f) => Variant::Double(*f),
        ResultData::String(s) => Variant::Str(s.clone()),
        ResultData::Error(e) => Variant::ErrValue(error_code(e)),
        other => Variant::Str(other.to_string()),
    }
}

/// Excel's error strings and their `CVErr` numbers, which are what `CLng` on
/// an error `Variant` gives back (measured: 2042 for a failed lookup).
const ERROR_CODES: &[(&str, i32)] = &[
    ("#NULL!", 2000),
    ("#DIV/0!", 2007),
    ("#VALUE!", 2015),
    ("#REF!", 2023),
    ("#NAME?", 2029),
    ("#NUM!", 2036),
    ("#N/A", 2042),
];

fn error_code(text: &str) -> i32 {
    ERROR_CODES
        .iter()
        .find(|(s, _)| text.starts_with(s))
        .map(|(_, n)| *n)
        // `#SPILL!`, `#CALC!` and friends postdate the `xlErr*` constants and
        // have no `CVErr` number at all; `#VALUE!`'s is the closest thing to
        // "something is wrong with this value" the enumeration offers.
        .unwrap_or(2015)
}

fn error_text(code: i32) -> &'static str {
    ERROR_CODES
        .iter()
        .find(|(_, n)| *n == code)
        .map(|(s, _)| *s)
        .unwrap_or("#VALUE!")
}

#[cfg(test)]
mod tests {
    use crate::core::engine::{Sheet, SheetInit};
    use crate::core::workbook::WorkbookManager;

    /// The same grid `fuzz/vba_host_probe.py` builds, so an expectation here
    /// can be read straight off a probe run against real Excel.
    ///
    /// `A1:A3` = 1/2/3, `B1:B3` = 10/20/30, `C1` a date-formatted serial,
    /// `C2` a fractional one, `D1` a formula, `D2` text, `E1` a boolean,
    /// `E2` empty, `F1` an error.
    fn fixture() -> WorkbookManager {
        let mut wb = WorkbookManager {
            sheets: vec![
                Sheet::new(SheetInit {
                    name: Some("Sheet1".to_string()),
                    rows: 6,
                    cols: 8,
                    ..Default::default()
                }),
                Sheet::new(SheetInit {
                    name: Some("Sheet2".to_string()),
                    rows: 4,
                    cols: 4,
                    ..Default::default()
                }),
            ],
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            vba_project: None,
            locale: crate::core::locale::Locale::default(),
        };
        let s = &mut wb.sheets[0];
        for (row, (a, b)) in [(1, 10), (2, 20), (3, 30)].into_iter().enumerate() {
            s.set_cell_src(row, 0, a.to_string());
            s.set_cell_src(row, 1, b.to_string());
        }
        s.set_cell_src(0, 2, "46195".to_string());
        s.set_cell_src(1, 2, "46195.5".to_string());
        for row in 0..2 {
            s.update_cell_style(row, 2, |st| st.num_format = Some("m/d/yy".to_string()));
        }
        s.set_cell_src(0, 3, "=A1*2".to_string());
        s.set_cell_src(1, 3, "\"hi\"".to_string());
        s.set_cell_src(0, 4, "TRUE".to_string());
        s.set_cell_src(0, 5, "=1/0".to_string());
        wb.evaluate().unwrap();
        wb
    }

    /// Runs one expression as a macro over the fixture and reports
    /// `TypeName|CStr`, or `ERR|number` -- the exact shape
    /// `fuzz/vba_host_probe.py` prints from Excel.
    fn probe(setup_and_expr: &str) -> String {
        let (setup, expr) = match setup_and_expr.rsplit_once("::") {
            Some((s, e)) => (s.replace("\\n", "\n    "), e.to_string()),
            None => (String::new(), setup_and_expr.to_string()),
        };
        let source = format!(
            "Attribute VB_Name = \"H\"\n\
             Function Gen()\n    \
             Dim ws As Worksheet, wb As Workbook\n    \
             Set wb = ThisWorkbook\n    \
             Set ws = wb.Worksheets(\"Sheet1\")\n    \
             Dim v, s, c\n    \
             {setup}\n    \
             Gen = {expr}\nEnd Function\n"
        );
        let mut wb = fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "H".to_string(),
            crate::core::VbaModuleKind::Standard,
            source,
            None,
        )
        .unwrap();
        match wb.run_macro(Some("H"), "Gen", &[]) {
            Ok(out) => format!("{}|{}", out.type_name, out.value.unwrap_or_default()),
            Err(crate::Error::VbaRuntime { number, .. }) => format!("ERR|{number}"),
            Err(e) => format!("FAIL|{e}"),
        }
    }

    // Every expectation below is what `fuzz/vba_host_probe.py` returned from
    // Excel for Mac 16.112 for the same expression, minus the `OK|` prefix
    // the harness adds. Do not "correct" one from memory -- re-measure.

    #[test]
    fn objects_report_the_type_names_excel_reports() {
        assert_eq!(probe("TypeName(wb)"), "String|Workbook");
        assert_eq!(probe("TypeName(ws)"), "String|Worksheet");
        assert_eq!(probe("TypeName(ws.Range(\"A1\"))"), "String|Range");
        assert_eq!(probe("TypeName(ws.Range(\"A1:B2\"))"), "String|Range");
        assert_eq!(probe("TypeName(ws.Cells)"), "String|Range");
        assert_eq!(probe("TypeName(wb.Worksheets)"), "String|Sheets");
        assert_eq!(probe("TypeName(wb.Sheets)"), "String|Sheets");
        assert_eq!(probe("TypeName(wb.Worksheets(1))"), "String|Worksheet");
    }

    #[test]
    fn addresses_and_shape() {
        assert_eq!(probe("ws.Range(\"A1\").Address"), "String|$A$1");
        assert_eq!(probe("ws.Range(\"A1:B2\").Address"), "String|$A$1:$B$2");
        assert_eq!(probe("ws.Range(\"A1\").Address(False, False)"), "String|A1");
        assert_eq!(probe("ws.Cells(2, 3).Address"), "String|$C$2");
        assert_eq!(
            probe("ws.Range(\"A1\", \"B2\").Address"),
            "String|$A$1:$B$2"
        );
        assert_eq!(
            probe("ws.Range(ws.Cells(1, 1), ws.Cells(2, 2)).Address"),
            "String|$A$1:$B$2"
        );
        assert_eq!(
            probe("ws.Range(\"B2\").Offset(1, 1).Address"),
            "String|$C$3"
        );
        assert_eq!(
            probe("ws.Range(\"B2\").Offset(-1, 0).Address"),
            "String|$B$1"
        );
        assert_eq!(
            probe("ws.Range(\"A1:B2\").Resize(3, 1).Address"),
            "String|$A$1:$A$3"
        );
        // Measured: `ws.Cells.Address` renders as whole rows, not
        // `$A$1:$XFD$1048576`.
        assert_eq!(probe("ws.Cells.Address"), "String|$1:$1048576");
        assert_eq!(probe("CStr(ws.Range(\"A1:B2\").Count)"), "String|4");
        assert_eq!(probe("TypeName(ws.Range(\"A1:B2\").Count)"), "String|Long");
        assert_eq!(
            probe("ws.Range(\"A1\").Row & \",\" & ws.Range(\"B2\").Column"),
            "String|1,2"
        );
        // `.Row`/`.Column` report the top-left of a multi-cell range.
        assert_eq!(
            probe("ws.Range(\"B2:C4\").Row & \",\" & ws.Range(\"B2:C4\").Column"),
            "String|2,2"
        );
    }

    #[test]
    fn a_whole_sheet_count_overflows_a_long_exactly_as_in_excel() {
        // Not an implementation limit: `Range.Count` is typed `Long` and the
        // grid does not fit in one, so Excel itself reports error 6.
        assert_eq!(probe("CStr(ws.Cells.Count)"), "ERR|6");
    }

    #[test]
    fn range_errors_are_the_numbers_excel_reports() {
        assert_eq!(probe("ws.Range(\"nope!!\").Address"), "ERR|1004");
        assert_eq!(probe("ws.Range(\"\").Address"), "ERR|1004");
        assert_eq!(probe("ws.Range(\"A1\").Offset(-1, 0).Address"), "ERR|1004");
        assert_eq!(probe("ws.Range(\"A1\").Resize(0, 1).Address"), "ERR|1004");
        assert_eq!(probe("wb.Worksheets(\"nope\").Name"), "ERR|9");
        assert_eq!(probe("wb.Worksheets(5).Name"), "ERR|9");
    }

    #[test]
    fn for_each_over_a_range_is_row_major() {
        // Excel walks left-to-right, then down.
        assert_eq!(
            probe(
                "For Each c In ws.Range(\"A1:B2\")\\n s = s & c.Address(False, False) & \" \"\\nNext :: s"
            ),
            "String|A1 B1 A2 B2 "
        );
        assert_eq!(
            probe("For Each c In wb.Worksheets\\n s = s & c.Name & \" \"\\nNext :: s"),
            "String|Sheet1 Sheet2 "
        );
    }

    #[test]
    fn a_numeric_cell_always_reads_back_as_a_double() {
        // Even a cell holding `1`: there is no `Integer` on this path.
        assert_eq!(probe("TypeName(ws.Range(\"A1\").Value)"), "String|Double");
        assert_eq!(probe("CStr(ws.Range(\"A1\").Value)"), "String|1");
        assert_eq!(probe("TypeName(ws.Range(\"D1\").Value)"), "String|Double");
        assert_eq!(probe("CStr(ws.Range(\"D1\").Value)"), "String|2");
        assert_eq!(probe("TypeName(ws.Range(\"D2\").Value)"), "String|String");
        assert_eq!(probe("TypeName(ws.Range(\"E1\").Value)"), "String|Boolean");
        assert_eq!(probe("TypeName(ws.Range(\"E2\").Value)"), "String|Empty");
        assert_eq!(probe("CStr(ws.Range(\"E2\").Value)"), "String|");
    }

    #[test]
    fn a_date_formatted_cell_reads_as_date_through_value_and_double_through_value2() {
        // The whole reason `Variant::Date` exists. The engine stores only the
        // serial; the notation is the cell's number format.
        assert_eq!(probe("TypeName(ws.Range(\"C1\").Value)"), "String|Date");
        assert_eq!(probe("CStr(ws.Range(\"C1\").Value)"), "String|6/22/26");
        assert_eq!(probe("TypeName(ws.Range(\"C1\").Value2)"), "String|Double");
        assert_eq!(probe("CStr(ws.Range(\"C1\").Value2)"), "String|46195");
        // A fractional serial is still a Date, and CStr shows the time.
        assert_eq!(probe("TypeName(ws.Range(\"C2\").Value)"), "String|Date");
        assert_eq!(
            probe("CStr(ws.Range(\"C2\").Value)"),
            "String|6/22/26 12:00:00 PM"
        );
        assert_eq!(probe("CStr(ws.Range(\"C2\").Value2)"), "String|46195.5");
    }

    #[test]
    fn formula_and_text_read_the_source_and_the_rendering() {
        assert_eq!(probe("ws.Range(\"D1\").Formula"), "String|=A1*2");
        assert_eq!(probe("ws.Range(\"D2\").Formula"), "String|hi");
        assert_eq!(probe("ws.Range(\"E2\").Formula"), "String|");
        assert_eq!(probe("ws.Range(\"C1\").Text"), "String|6/22/26");
        assert_eq!(probe("ws.Range(\"A1\").Text"), "String|1");
        assert_eq!(probe("ws.Range(\"F1\").Text"), "String|#DIV/0!");
    }

    #[test]
    fn an_error_in_a_cell_reads_back_as_an_error_variant() {
        assert_eq!(probe("TypeName(ws.Range(\"F1\").Value)"), "String|Error");
        assert_eq!(probe("CStr(CLng(ws.Range(\"F1\").Value))"), "String|2007");
        assert_eq!(
            probe("CStr(IsError(ws.Range(\"F1\").Value))"),
            "String|True"
        );
        // ... but arithmetic on one is a type mismatch, not a propagation.
        assert_eq!(probe("v = ws.Range(\"F1\").Value :: CStr(v + 1)"), "ERR|13");
    }

    #[test]
    fn a_multi_cell_value_read_is_a_two_dimensional_array() {
        assert_eq!(
            probe("v = ws.Range(\"A1:A3\").Value :: TypeName(v)"),
            "String|Variant()"
        );
        assert_eq!(
            probe("v = ws.Range(\"A1:A3\").Value :: CStr(v(2, 1))"),
            "String|2"
        );
        assert_eq!(
            probe(
                "v = ws.Range(\"A1:A3\").Value :: CStr(UBound(v, 1)) & \",\" & CStr(UBound(v, 2))"
            ),
            "String|3,1"
        );
        // Reading a Range without `.Value` takes the same default member.
        assert_eq!(
            probe("v = ws.Range(\"A1:A3\") :: TypeName(v)"),
            "String|Variant()"
        );
        // And an array is not a scalar.
        assert_eq!(probe("v = ws.Range(\"A1:B2\").Value :: CStr(v)"), "ERR|13");
        assert_eq!(
            probe("v = ws.Range(\"A1\").Value :: TypeName(v)"),
            "String|Double"
        );
    }

    #[test]
    fn is_compares_object_identity_not_coordinates() {
        // The measurement that decided the design: two `Range()` calls over
        // the same cell are different objects, while a worksheet is cached.
        assert_eq!(
            probe("CStr(ws.Range(\"A1\") Is ws.Range(\"A1\"))"),
            "String|False"
        );
        assert_eq!(probe("CStr(ws Is wb.Worksheets(1))"), "String|True");
        assert_eq!(probe("CStr(ws Is wb.Worksheets(2))"), "String|False");
        assert_eq!(probe("CStr(ws Is Nothing)"), "String|False");
        // Copying preserves identity, which is what makes `Is` on a Range
        // usable at all.
        assert_eq!(
            probe("Dim r As Range :: Set r = ws.Range(\"A1\") :: CStr(r Is r)"),
            "String|True"
        );
        assert_eq!(
            probe(
                "Dim r As Range, q As Range :: Set r = ws.Range(\"A1\") :: Set q = r :: CStr(q Is r)"
            ),
            "String|True"
        );
        assert_eq!(
            probe("Dim r As Range :: Set r = ws.Range(\"A1\") :: CStr(r Is ws.Range(\"A1\"))"),
            "String|False"
        );
    }

    #[test]
    fn an_unset_object_variable_is_nothing() {
        assert_eq!(probe("Dim r As Range :: TypeName(r)"), "String|Nothing");
        assert_eq!(probe("Dim r As Range :: CStr(r Is Nothing)"), "String|True");
    }

    #[test]
    fn with_binds_the_object_not_its_value() {
        assert_eq!(
            probe("With ws.Range(\"A2\")\\n s = CStr(.Value) & \"/\" & .Address\\nEnd With :: s"),
            "String|2/$A$2"
        );
    }

    #[test]
    fn worksheet_function_bridges_onto_the_formula_library() {
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(ws.Range(\"A1:A3\")))"),
            "String|6"
        );
        assert_eq!(
            probe("TypeName(Application.WorksheetFunction.Sum(ws.Range(\"A1:A3\")))"),
            "String|Double"
        );
        // A direct string argument coerces; text and booleans *inside a
        // range* do not. Both measured, and both fall out of reusing the
        // formula path rather than being written twice.
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(\"1\", 2))"),
            "String|3"
        );
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(ws.Range(\"D2\")))"),
            "String|0"
        );
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(ws.Range(\"E1\")))"),
            "String|0"
        );
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(ws.Range(\"C1\")))"),
            "String|46195"
        );
        assert_eq!(
            probe("Application.WorksheetFunction.Text(ws.Range(\"C1\").Value, \"yyyy-mm-dd\")"),
            "String|2026-06-22"
        );
        // A mixed range-and-scalar argument list, and an array argument.
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(ws.Range(\"A1\"), 5))"),
            "String|6"
        );
        assert_eq!(
            probe("v = ws.Range(\"A1:A3\").Value :: CStr(Application.WorksheetFunction.Sum(v))"),
            "String|6"
        );
        // `Application.X` is the same implementation.
        assert_eq!(
            probe("CStr(Application.Sum(ws.Range(\"A1:A3\")))"),
            "String|6"
        );
    }

    #[test]
    fn worksheet_function_raises_where_application_returns_an_error() {
        // Two call paths over one implementation, and the difference is
        // entirely in what failure looks like.
        assert_eq!(
            probe(
                "CStr(Application.WorksheetFunction.VLookup(\"zzz\", ws.Range(\"A1:B3\"), 2, False))"
            ),
            "ERR|1004"
        );
        assert_eq!(
            probe("TypeName(Application.VLookup(\"zzz\", ws.Range(\"A1:B3\"), 2, False))"),
            "String|Error"
        );
        assert_eq!(
            probe("CStr(IsError(Application.VLookup(\"zzz\", ws.Range(\"A1:B3\"), 2, False)))"),
            "String|True"
        );
        assert_eq!(
            probe(
                "v = Application.VLookup(\"zzz\", ws.Range(\"A1:B3\"), 2, False) :: CStr(CLng(v))"
            ),
            "String|2042"
        );
        // An error inside the data raises rather than propagating.
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Sum(ws.Range(\"F1\")))"),
            "ERR|1004"
        );
    }

    #[test]
    fn worksheet_function_does_not_expose_the_engines_own_functions() {
        // `SLICE` and friends are `evaluate_function` names that Excel has
        // never heard of. A macro using one would work here and fail in
        // Excel -- the one direction a differential harness cannot catch,
        // since it only generates what Excel accepts.
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Slice(ws.Range(\"A1:A3\"), 1))"),
            "ERR|438"
        );
        assert_eq!(
            probe("CStr(Application.WorksheetFunction.Get(1))"),
            "ERR|438"
        );
    }

    #[test]
    fn unqualified_range_and_cells_resolve_against_the_active_sheet() {
        assert_eq!(
            probe("Range(\"A1\").Address & \"/\" & Cells(2, 2).Address"),
            "String|$A$1/$B$2"
        );
        assert_eq!(probe("CStr(Range(\"A2\").Value)"), "String|2");
    }

    #[test]
    fn worksheets_are_reachable_by_name_index_and_count() {
        assert_eq!(probe("CStr(wb.Worksheets.Count)"), "String|2");
        assert_eq!(probe("wb.Worksheets(1).Name"), "String|Sheet1");
        // Case-insensitive, as every VBA name lookup is.
        assert_eq!(
            probe("CStr(wb.Worksheets(\"SHEET1\").Name)"),
            "String|Sheet1"
        );
        assert_eq!(
            probe("wb.Worksheets(\"Sheet1\").Range(\"B3\").Address"),
            "String|$B$3"
        );
    }

    // -- writes ----------------------------------------------------------

    #[test]
    fn a_write_is_visible_to_the_next_read() {
        assert_eq!(
            probe("ws.Range(\"G1\").Value = 5 :: CStr(ws.Range(\"G1\").Value)"),
            "String|5"
        );
        // The default member: no `.Value` on the left.
        assert_eq!(
            probe("ws.Range(\"G2\") = 7 :: CStr(ws.Range(\"G2\").Value)"),
            "String|7"
        );
        // Assigning to a multi-cell range fills all of it.
        assert_eq!(
            probe("ws.Range(\"G3:H4\").Value = 3 :: CStr(ws.Range(\"H4\").Value)"),
            "String|3"
        );
        // An array assigned to one cell writes its first element.
        assert_eq!(
            probe(
                "ws.Range(\"G5\").Value = ws.Range(\"A1:A3\").Value :: TypeName(ws.Range(\"G5\").Value) & \"|\" & CStr(ws.Range(\"G5\").Value)"
            ),
            "String|Double|1"
        );
    }

    #[test]
    fn a_write_recalculates_before_the_next_read_that_could_see_it() {
        // `D1` holds `=A1*2`. This is the behaviour the lazy-recalculation
        // design exists to preserve: Excel in automatic mode would have
        // recalculated at the assignment, and the difference is invisible.
        assert_eq!(
            probe("ws.Range(\"A1\").Value = 5 :: CStr(ws.Range(\"D1\").Value)"),
            "String|10"
        );
        // A formula written by the macro evaluates too.
        assert_eq!(
            probe(
                "ws.Range(\"G1\").Value = 5\\nws.Range(\"G2\").Formula = \"=G1*2\" :: CStr(ws.Range(\"G2\").Value)"
            ),
            "String|10"
        );
    }

    #[test]
    fn writing_a_string_parses_it_the_way_typing_it_would() {
        // Measured, and the reason the writer hands the raw text to
        // `Sheet::commit` rather than deciding for itself: a leading `=`
        // makes a formula, and date-looking text makes a date.
        assert_eq!(
            probe(
                "ws.Range(\"G4\").Value = \"=1+2\" :: ws.Range(\"G4\").Formula & \"|\" & CStr(ws.Range(\"G4\").Value)"
            ),
            "String|=1+2|3"
        );
        assert_eq!(
            probe(
                "ws.Range(\"G6\").Value = \"6/22/2026\" :: TypeName(ws.Range(\"G6\").Value) & \"|\" & CStr(ws.Range(\"G6\").Value2)"
            ),
            "String|Date|46195"
        );
        assert_eq!(
            probe(
                "ws.Range(\"G7\").Value = \"hello\" :: TypeName(ws.Range(\"G7\").Value) & \"|\" & ws.Range(\"G7\").Formula"
            ),
            "String|String|hello"
        );
    }

    #[test]
    fn writing_a_date_writes_the_notation_too() {
        // A date is a serial plus a number format, so writing one has to set
        // both -- otherwise `.Text` would show 46195.
        assert_eq!(
            probe(
                "ws.Range(\"G8\").Value = #6/22/2026# :: TypeName(ws.Range(\"G8\").Value) & \"|\" & ws.Range(\"G8\").Text"
            ),
            "String|Date|6/22/26"
        );
    }

    #[test]
    fn writing_whitespace_padded_numeric_text_stores_a_number() {
        // Entering `"  3  "` into a cell gives the *number* 3, in Excel and
        // through `Range.Value`. Found by `fuzz/fuzz_vba.py`'s cell
        // comparison on two independent cases, where the return value agreed
        // and only the cell differed.
        assert_eq!(
            probe(
                "ws.Range(\"H5\").Value = \"  3  \" :: TypeName(ws.Range(\"H5\").Value) & \"|\" & CStr(ws.Range(\"H5\").Value)"
            ),
            "String|Double|3"
        );
        // Text that merely has spaces around it is still text.
        assert_eq!(
            probe("ws.Range(\"H6\").Value = \"  x  \" :: TypeName(ws.Range(\"H6\").Value)"),
            "String|String"
        );
    }

    #[test]
    fn overflowing_pow_raises_before_a_cell_write() {
        // Measured after fuzz/fuzz_vba.py found that runtime `^` overflows
        // rather than producing an infinity that can be assigned to a cell.
        assert_eq!(
            probe(
                "v = 1000\nws.Range(\"G9\").Value = (-2.5 ^ v) :: TypeName(ws.Range(\"G9\").Value)"
            ),
            "ERR|6"
        );
        // An error value assigned as text is the error too, which is what
        // Excel does with the same assignment.
        assert_eq!(
            probe(
                "ws.Range(\"H1\").Value = \"#N/A\" :: TypeName(ws.Range(\"H1\").Value) & \"|\" & ws.Range(\"H1\").Formula"
            ),
            "String|Error|#N/A"
        );
        assert_eq!(
            probe(
                "ws.Range(\"H2\").Value = CVErr(2036) :: TypeName(ws.Range(\"H2\").Value) & \"|\" & ws.Range(\"H2\").Formula"
            ),
            "String|Error|#NUM!"
        );
    }

    #[test]
    fn a_sheet_can_be_renamed_and_the_rename_is_visible() {
        assert_eq!(
            probe("wb.Worksheets(2).Name = \"Renamed\" :: wb.Worksheets(2).Name"),
            "String|Renamed"
        );
    }

    #[test]
    fn a_run_that_only_reads_does_not_report_a_mutation() {
        let mut wb = fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "M".to_string(),
            crate::core::VbaModuleKind::Standard,
            "Attribute VB_Name = \"M\"\n\
             Function Reads()\n    Reads = ThisWorkbook.Worksheets(1).Range(\"A1\").Value\nEnd Function\n\
             Sub Writes()\n    ThisWorkbook.Worksheets(1).Range(\"Z1\").Value = 1\nEnd Sub\n"
                .to_string(),
            None,
        )
        .unwrap();
        assert!(!wb.run_macro(None, "Reads", &[]).unwrap().mutated);
        assert!(wb.run_macro(None, "Writes", &[]).unwrap().mutated);
    }

    #[test]
    fn a_macro_that_raises_after_writing_still_leaves_the_write() {
        // Excel does not roll back. Reporting the error while discarding the
        // writes would be a quieter kind of wrong.
        let mut wb = fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "M".to_string(),
            crate::core::VbaModuleKind::Standard,
            "Attribute VB_Name = \"M\"\n\
             Sub Half()\n    \
             ThisWorkbook.Worksheets(1).Range(\"Z1\").Value = 42\n    \
             Err.Raise 5\nEnd Sub\n"
                .to_string(),
            None,
        )
        .unwrap();
        assert!(wb.run_macro(None, "Half", &[]).is_err());
        assert_eq!(
            wb.sheets[0].get_display_string(&crate::core::CellRef::new(0, 25)),
            "42"
        );
    }

    #[test]
    fn a_missing_module_or_procedure_is_reported_rather_than_guessed() {
        let mut wb = fixture();
        assert!(wb.run_macro(None, "Nope", &[]).is_err());
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "M".to_string(),
            crate::core::VbaModuleKind::Standard,
            "Attribute VB_Name = \"M\"\nSub Only()\nEnd Sub\n".to_string(),
            None,
        )
        .unwrap();
        assert!(wb.run_macro(Some("Missing"), "Only", &[]).is_err());
        assert!(wb.run_macro(None, "NotDeclared", &[]).is_err());
        assert!(wb.run_macro(Some("M"), "Only", &[]).is_ok());
    }

    #[test]
    fn a_write_far_outside_the_sheet_is_refused_rather_than_allocated() {
        // Excel's grid is sparse and `visi`'s is dense, so the honest answer
        // to `XFD1048576` is error 7 rather than seventeen billion cells.
        assert_eq!(probe("ws.Range(\"XFD1048576\").Value = 1 :: \"\""), "ERR|7");
        assert_eq!(probe("CStr(ws.Range(\"A1000000\").Value)"), "String|");
    }

    #[test]
    fn out_of_scope_members_still_name_themselves() {
        // The refusal is the feature. Each of these is a real Excel member
        // that is still out of scope; `Interior.Color`, `Font.Bold` and
        // `NumberFormat` have left this list because they are now
        // implemented, which is the only reason a member may leave it.
        for case in [
            "wb.PivotTables(1).Name",
            "ws.Range(\"A1:B2\").Sort",
            "ws.Range(\"A1:B2\").Find(\"x\")",
            "ws.Range(\"A1\").Interior.Pattern",
            "ws.Range(\"A1\").Font.Strikethrough",
            "ws.Range(\"A1\").Borders",
        ] {
            assert_eq!(probe(case), "ERR|438", "{case}");
        }
    }

    // ---------------------------------------------------------------
    // Range tracking under a structural edit.
    //
    // These run against `tracking_fixture`, which is the *same* grid
    // `fuzz/vba_range_tracking_probe.py` builds, so every expectation below
    // is literally the string Excel for Mac 16.112 returned for the same
    // expression (minus the harness's `OK|` prefix). Re-run the probe rather
    // than adjusting one from memory.
    // ---------------------------------------------------------------

    /// 1..10 down column A, 101..110 down B, 201..210 down C, plus an empty
    /// `Sheet2` -- distinct per row so a tracked range can be asked what it
    /// now reads, not only where it now points.
    fn tracking_fixture() -> WorkbookManager {
        let mut wb = WorkbookManager {
            sheets: vec![
                Sheet::new(SheetInit {
                    name: Some("Sheet1".to_string()),
                    rows: 10,
                    cols: 3,
                    ..Default::default()
                }),
                Sheet::new(SheetInit {
                    name: Some("Sheet2".to_string()),
                    rows: 4,
                    cols: 4,
                    ..Default::default()
                }),
            ],
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            vba_project: None,
            locale: crate::core::locale::Locale::default(),
        };
        let s = &mut wb.sheets[0];
        for row in 0..10 {
            s.set_cell_src(row, 0, (row + 1).to_string());
            s.set_cell_src(row, 1, (101 + row).to_string());
            s.set_cell_src(row, 2, (201 + row).to_string());
        }
        wb.evaluate().unwrap();
        wb
    }

    fn tracking_probe(setup_and_expr: &str) -> String {
        let (setup, expr) = match setup_and_expr.rsplit_once("::") {
            Some((s, e)) => (s.replace("\\n", "\n    "), e.to_string()),
            None => (String::new(), setup_and_expr.to_string()),
        };
        let source = format!(
            "Attribute VB_Name = \"H\"\n\
             Function Gen()\n    \
             Dim ws As Worksheet, wb As Workbook\n    \
             Set wb = ThisWorkbook\n    \
             Set ws = wb.Worksheets(\"Sheet1\")\n    \
             Dim r As Range, q As Range, s\n    \
             {setup}\n    \
             Gen = {expr}\nEnd Function\n"
        );
        let mut wb = tracking_fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "H".to_string(),
            crate::core::VbaModuleKind::Standard,
            source,
            None,
        )
        .unwrap();
        match wb.run_macro(Some("H"), "Gen", &[]) {
            Ok(out) => format!("{}|{}", out.type_name, out.value.unwrap_or_default()),
            Err(crate::Error::VbaRuntime { number, .. }) => format!("ERR|{number}"),
            Err(e) => format!("FAIL|{e}"),
        }
    }

    #[test]
    fn a_held_range_follows_a_row_inserted_above_it() {
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(1).Insert :: r.Address"#),
            "String|$A$6"
        );
        // It follows the *data*, not just the coordinates: A5 held 5, and
        // after the insert that 5 is at A6 and `r` still reads it.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(1).Insert :: CStr(r.Value)"#),
            "String|5"
        );
        // An insert below it changes nothing.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(9).Insert :: r.Address"#),
            "String|$A$5"
        );
    }

    #[test]
    fn inserting_at_a_held_spans_first_row_moves_it_and_inserting_inside_grows_it() {
        // The same asymmetry `core::grid_edit` encodes for a formula's range
        // reference, and measured separately here because there was no
        // reason a priori for Excel to treat the two the same way.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:A7") :: ws.Rows(5).Insert :: r.Address"#),
            "String|$A$6:$A$8"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:A7") :: ws.Rows(6).Insert :: r.Address"#),
            "String|$A$5:$A$8"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:A7") :: ws.Rows(8).Insert :: r.Address"#),
            "String|$A$5:$A$7"
        );
    }

    #[test]
    fn deleting_rows_moves_and_shrinks_a_held_range() {
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(1).Delete :: r.Address"#),
            "String|$A$4"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:A7") :: ws.Rows(6).Delete :: r.Address"#),
            "String|$A$5:$A$6"
        );
        // Deleting the span's own first row shrinks it from the bottom, not
        // the top -- the start stays where it is because everything below
        // moved up into it.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:A7") :: ws.Rows(5).Delete :: r.Address"#),
            "String|$A$5:$A$6"
        );
    }

    #[test]
    fn a_range_whose_cells_are_all_deleted_is_dead_but_not_nothing() {
        // Measured: the object is *not* `Nothing` and still calls itself a
        // `Range`, but every member access raises.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(5).Delete :: CStr(r Is Nothing)"#),
            "String|False"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(5).Delete :: TypeName(r)"#),
            "String|Range"
        );
        // The *number* is a deliberate divergence: Excel for Mac's is not
        // reproducible (the same case gave -1667945984 and then -1667949824),
        // so visi raises 1004 -- the number Excel on Windows documents for
        // this message, and the one the rest of this module already uses.
        // See `docs/excel-discrepancies.md`.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(5).Delete :: r.Address"#),
            "ERR|1004"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5") :: ws.Rows(5).Delete :: CStr(r.Value)"#),
            "ERR|1004"
        );
        // A multi-row span loses every row it had.
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:A7") :: ws.Rows("5:7").Delete :: r.Address"#),
            "ERR|1004"
        );
    }

    #[test]
    fn a_held_range_tracks_column_edits_on_the_column_axis() {
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("C5") :: ws.Columns(1).Insert :: r.Address"#),
            "String|$D$5"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("A5:C5") :: ws.Columns(2).Insert :: r.Address"#),
            "String|$A$5:$D$5"
        );
        assert_eq!(
            tracking_probe(r#"Set r = ws.Range("C5") :: ws.Columns(3).Delete :: r.Address"#),
            "ERR|1004"
        );
    }

    #[test]
    fn an_edit_on_another_sheet_leaves_a_held_range_alone() {
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("A5") :: wb.Worksheets("Sheet2").Rows(1).Insert :: r.Address"#
            ),
            "String|$A$5"
        );
    }

    #[test]
    fn every_copy_of_a_range_tracks_the_edit_and_stays_the_same_object() {
        // This is the case that forced interning: `q` was copied *before* the
        // edit, and Excel moves it too, which a by-value Range cannot do.
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("A5") :: Set q = r :: ws.Rows(1).Insert :: CStr(q Is r) & "/" & q.Address"#
            ),
            "String|True/$A$6"
        );
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("A5") :: Set q = r :: ws.Rows(1).Insert :: r.Address & "/" & q.Address"#
            ),
            "String|$A$6/$A$6"
        );
    }

    #[test]
    fn rows_columns_and_entire_row_are_whole_band_ranges() {
        assert_eq!(tracking_probe("ws.Rows(3).Address"), "String|$3:$3");
        assert_eq!(tracking_probe("ws.Columns(3).Address"), "String|$C:$C");
        assert_eq!(tracking_probe("TypeName(ws.Rows(3))"), "String|Range");
        assert_eq!(
            tracking_probe(r#"ws.Range("B5").EntireRow.Address"#),
            "String|$5:$5"
        );
        assert_eq!(
            tracking_probe(r#"ws.Range("B5").EntireColumn.Address"#),
            "String|$B:$B"
        );
    }

    #[test]
    fn an_edit_spelled_through_entire_row_does_the_same_thing() {
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("A5") :: ws.Range("A1").EntireRow.Insert :: r.Address"#
            ),
            "String|$A$6"
        );
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("C5") :: ws.Range("A1").EntireColumn.Insert :: r.Address"#
            ),
            "String|$D$5"
        );
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("A5") :: ws.Range("A5").EntireRow.Delete :: r.Address"#
            ),
            "ERR|1004"
        );
    }

    #[test]
    fn a_partial_range_insert_is_refused_rather_than_guessed_at() {
        // Excel accepts this and picks the shift direction from the range's
        // shape -- measured, `Range("A2:A3").Insert` shifts *right*. Picking
        // one here would silently move a macro's data the wrong way, so 438
        // names the construct instead, as everywhere else in this module.
        assert_eq!(
            tracking_probe(r#"ws.Range("A2:A3").Insert :: "unreachable""#),
            "ERR|438"
        );
    }

    #[test]
    fn a_structural_edit_from_a_macro_shifts_the_formulas_too() {
        // The engine-level rewrite and the range tracking are separate
        // mechanisms; this is the one case that exercises both at once.
        assert_eq!(
            tracking_probe(
                r#"ws.Range("E1").Formula = "=A5" :: ws.Rows(1).Insert :: ws.Range("E2").Formula"#
            ),
            "String|=A6"
        );
    }

    // ---------------------------------------------------------------
    // Styles: `.Interior`, `.Font` and `.NumberFormat`.
    //
    // Every expectation is what `fuzz/vba_style_probe.py` returned from Excel
    // for Mac 16.112 for the same expression. The colour cases are also
    // pinned from the *other* side by that probe's `--paint` channel, which
    // has Excel save the workbook and reads the real ARGB back with openpyxl
    // -- the VBA channel alone cannot catch a consistent-but-wrong
    // convention, since both engines would round-trip the same wrong Long.
    // ---------------------------------------------------------------

    #[test]
    fn rgb_is_a_builtin_returning_the_long_excel_returns() {
        assert_eq!(probe("CStr(RGB(255, 0, 0))"), "String|255");
        assert_eq!(probe("CStr(RGB(0, 255, 0))"), "String|65280");
        assert_eq!(probe("CStr(RGB(0, 0, 255))"), "String|16711680");
        assert_eq!(probe("CStr(RGB(1, 2, 3))"), "String|197121");
        assert_eq!(probe("TypeName(RGB(1, 2, 3))"), "String|Long");
        // Measured: clamps rather than carrying into the next byte, which
        // would have made `RGB(300, 0, 0)` green.
        assert_eq!(probe("CStr(RGB(300, 0, 0))"), "String|255");
        // Measured: a negative component is error 5, not a clamp to zero.
        assert_eq!(probe("CStr(RGB(-1, 0, 0))"), "ERR|5");
    }

    #[test]
    fn interior_color_is_bgr_so_ff0000_is_blue() {
        // `&HFF0000` must land in the *blue* channel of the stored style,
        // not the red one.
        assert_eq!(
            probe(
                r#"ws.Range("G2").Interior.Color = &HFF0000 :: CStr(ws.Range("G2").Interior.Color)"#
            ),
            "String|16711680"
        );
        assert_eq!(
            probe(
                r#"ws.Range("G1").Interior.Color = RGB(255, 0, 0) :: CStr(ws.Range("G1").Interior.Color)"#
            ),
            "String|255"
        );
        // And the same read through the stored representation, which is where
        // a byte swap would show up rather than cancelling out.
        let mut wb = fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "H".to_string(),
            crate::core::VbaModuleKind::Standard,
            "Attribute VB_Name = \"H\"\nSub Go()\n    \
             ThisWorkbook.Worksheets(\"Sheet1\").Range(\"G1\").Interior.Color = &HFF0000\n    \
             ThisWorkbook.Worksheets(\"Sheet1\").Range(\"G2\").Interior.Color = RGB(255, 0, 0)\n\
             End Sub\n"
                .to_string(),
            None,
        )
        .unwrap();
        wb.run_macro(Some("H"), "Go", &[]).unwrap();
        assert_eq!(
            wb.sheets[0]
                .get_cell_style(0, 6)
                .unwrap()
                .bg_color
                .as_deref(),
            Some("#0000FF"),
            "&HFF0000 must store as blue"
        );
        assert_eq!(
            wb.sheets[0]
                .get_cell_style(1, 6)
                .unwrap()
                .bg_color
                .as_deref(),
            Some("#FF0000"),
            "RGB(255,0,0) must store as red"
        );
    }

    #[test]
    fn an_unstyled_cell_reports_excels_style_defaults() {
        assert_eq!(
            probe(r#"TypeName(ws.Range("A1").Interior)"#),
            "String|Interior"
        );
        assert_eq!(probe(r#"TypeName(ws.Range("A1").Font)"#), "String|Font");
        // White, not zero -- an unfilled cell is not "no colour" to Excel.
        assert_eq!(
            probe(r#"CStr(ws.Range("A1").Interior.Color)"#),
            "String|16777215"
        );
        assert_eq!(
            probe(r#"CStr(ws.Range("A1").Interior.ColorIndex)"#),
            "String|-4142"
        );
        assert_eq!(probe(r#"CStr(ws.Range("A1").Font.Bold)"#), "String|False");
        assert_eq!(probe(r#"CStr(ws.Range("A1").Font.Italic)"#), "String|False");
        assert_eq!(probe(r#"CStr(ws.Range("A1").Font.Size)"#), "String|11");
        // A `Double`, not a `Long` -- which is why `CellStyle::font_size` is
        // an `f64`.
        assert_eq!(
            probe(r#"TypeName(ws.Range("A1").Font.Size)"#),
            "String|Double"
        );
        assert_eq!(probe(r#"ws.Range("A1").Font.Name"#), "String|Calibri");
        assert_eq!(probe(r#"CStr(ws.Range("A1").Font.Color)"#), "String|0");
        // Slot 1 (black), *not* xlNone -- the opposite of Interior's default.
        assert_eq!(probe(r#"CStr(ws.Range("A1").Font.ColorIndex)"#), "String|1");
        assert_eq!(probe(r#"ws.Range("A1").NumberFormat"#), "String|General");
    }

    #[test]
    fn font_attributes_round_trip() {
        assert_eq!(
            probe(r#"ws.Range("H1").Font.Bold = True :: CStr(ws.Range("H1").Font.Bold)"#),
            "String|True"
        );
        assert_eq!(
            probe(r#"ws.Range("H2").Font.Italic = True :: CStr(ws.Range("H2").Font.Italic)"#),
            "String|True"
        );
        assert_eq!(
            probe(r#"ws.Range("H3").Font.Size = 14 :: CStr(ws.Range("H3").Font.Size)"#),
            "String|14"
        );
        // A half-point size survives, which an integer `font_size` could not
        // have stored.
        assert_eq!(
            probe(r#"ws.Range("M1").Font.Size = 10.5 :: CStr(ws.Range("M1").Font.Size)"#),
            "String|10.5"
        );
        assert_eq!(
            probe(r#"ws.Range("H4").Font.Name = "Courier New" :: ws.Range("H4").Font.Name"#),
            "String|Courier New"
        );
        assert_eq!(
            probe(
                r#"ws.Range("H5").Font.Color = RGB(0, 0, 255) :: CStr(ws.Range("H5").Font.Color)"#
            ),
            "String|16711680"
        );
    }

    #[test]
    fn color_index_is_the_palette_excel_reported() {
        // Setting a colour reports the slot it occupies...
        assert_eq!(
            probe(
                r#"ws.Range("G4").Interior.Color = RGB(255, 0, 0) :: CStr(ws.Range("G4").Interior.ColorIndex)"#
            ),
            "String|3"
        );
        // ...and setting the slot gives back the colour.
        assert_eq!(
            probe(
                r#"ws.Range("G5").Interior.ColorIndex = 3 :: CStr(ws.Range("G5").Interior.Color)"#
            ),
            "String|255"
        );
        // An off-palette colour reports the *nearest* slot, not xlNone. The
        // first implementation here guessed xlNone and was wrong; these three
        // are what Excel actually returned.
        assert_eq!(
            probe(
                r#"ws.Range("G7").Interior.Color = RGB(250, 10, 10) :: CStr(ws.Range("G7").Interior.ColorIndex)"#
            ),
            "String|3"
        );
        assert_eq!(
            probe(
                r#"ws.Range("G8").Interior.Color = RGB(10, 200, 10) :: CStr(ws.Range("G8").Interior.ColorIndex)"#
            ),
            "String|4"
        );
        assert_eq!(
            probe(
                r#"ws.Range("G9").Interior.Color = RGB(1, 2, 3) :: CStr(ws.Range("G9").Interior.ColorIndex)"#
            ),
            "String|1"
        );
        // xlNone clears the fill, after which Color reads white again.
        assert_eq!(
            probe(
                "ws.Range(\"G6\").Interior.Color = RGB(255, 0, 0)\\n\
                 ws.Range(\"G6\").Interior.ColorIndex = -4142 :: CStr(ws.Range(\"G6\").Interior.Color)"
            ),
            "String|16777215"
        );
    }

    #[test]
    fn a_mixed_range_reads_null_except_interior_color_which_reads_zero() {
        // Excel's own asymmetry, and the reason `style_fold` takes the
        // mixed-value rather than inferring it: 0 is also a legitimate
        // uniform answer, so it cannot double as a sentinel.
        assert_eq!(
            probe(
                r#"ws.Range("K1").Font.Bold = True :: CStr(IsNull(ws.Range("K1:K2").Font.Bold))"#
            ),
            "String|True"
        );
        assert_eq!(
            probe(r#"ws.Range("Q1").Font.Size = 14 :: TypeName(ws.Range("Q1:Q2").Font.Size)"#),
            "String|Null"
        );
        assert_eq!(
            probe(
                r#"ws.Range("R1").Font.Name = "Courier New" :: TypeName(ws.Range("R1:R2").Font.Name)"#
            ),
            "String|Null"
        );
        assert_eq!(
            probe(
                r#"ws.Range("S1").NumberFormat = "m/d/yy" :: TypeName(ws.Range("S1:S2").NumberFormat)"#
            ),
            "String|Null"
        );
        // The exception.
        assert_eq!(
            probe(
                r#"ws.Range("L1").Interior.Color = RGB(255, 0, 0) :: CStr(ws.Range("L1:L2").Interior.Color)"#
            ),
            "String|0"
        );
        // A range whose cells *agree* reports the value, not the sentinel.
        assert_eq!(
            probe(
                "ws.Range(\"N1\").Interior.Color = RGB(255, 0, 0)\\n\
                 ws.Range(\"N2\").Interior.Color = RGB(255, 0, 0) :: CStr(ws.Range(\"N1:N2\").Interior.Color)"
            ),
            "String|255"
        );
    }

    #[test]
    fn a_style_written_over_a_range_reaches_every_cell() {
        assert_eq!(
            probe(
                r#"ws.Range("J1:J3").Interior.Color = RGB(0, 255, 0) :: CStr(ws.Range("J3").Interior.Color)"#
            ),
            "String|65280"
        );
        assert_eq!(
            probe(r#"ws.Range("J1:J3").Font.Bold = True :: CStr(ws.Range("J2").Font.Bold)"#),
            "String|True"
        );
    }

    #[test]
    fn number_format_changes_the_rendering_and_leaves_the_serial_alone() {
        // Writing `NumberFormat` from a macro can turn a date cell into
        // a plain number *visually* while the serial stays put, which is
        // correct and is exactly `core::date`'s model -- there is no date
        // value type to lose.
        assert_eq!(
            probe(
                "ws.Range(\"I1\").Value = 46195\\n\
                 ws.Range(\"I1\").NumberFormat = \"m/d/yy\" :: ws.Range(\"I1\").Text"
            ),
            "String|6/22/26"
        );
        assert_eq!(
            probe(
                "ws.Range(\"C1\").NumberFormat = \"General\" :: \
                 CStr(ws.Range(\"C1\").Value2) & \"|\" & TypeName(ws.Range(\"C1\").Value)"
            ),
            "String|46195|Double"
        );
        assert_eq!(
            probe(
                "ws.Range(\"I2\").Value = 46195\\n\
                 ws.Range(\"I2\").NumberFormat = \"m/d/yy\"\\n\
                 ws.Range(\"I2\").NumberFormat = \"General\" :: \
                 ws.Range(\"I2\").Text & \"|\" & CStr(ws.Range(\"I2\").Value2)"
            ),
            "String|46195|46195"
        );
    }

    #[test]
    fn a_styled_range_tracks_a_structural_edit_like_any_other() {
        // `Interior` and `Font` ride on the range's handle rather than
        // carrying coordinates, so this falls out rather than needing its own
        // mechanism -- but it would be silently wrong if they did carry them.
        assert_eq!(
            tracking_probe(
                r#"Set r = ws.Range("A5") :: r.Interior.Color = RGB(255, 0, 0)\nws.Rows(1).Insert :: r.Address & "/" & CStr(r.Interior.Color)"#
            ),
            "String|$A$6/255"
        );
    }

    // ---------------------------------------------------------------
    // Excel Tables (`ListObjects`).
    //
    // `table_fixture` is the grid `fuzz/vba_table_probe.py` builds, so every
    // expectation is what Excel for Mac 16.112 returned for the same
    // expression. `Sales` is A1:C4 with a header row and no totals row, and
    // `E1` holds `=SUM(Sales[Amount])` so a rename's cascade is observable.
    // ---------------------------------------------------------------

    fn table_fixture() -> WorkbookManager {
        let mut wb = WorkbookManager {
            sheets: vec![Sheet::new(SheetInit {
                name: Some("Sheet1".to_string()),
                rows: 10,
                cols: 6,
                ..Default::default()
            })],
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            vba_project: None,
            locale: crate::core::locale::Locale::default(),
        };
        let rows: [[&str; 3]; 4] = [
            ["Region", "Product", "Amount"],
            ["East", "Widget", "10"],
            ["West", "Gadget", "20"],
            ["North", "Widget", "30"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                wb.sheets[0].set_cell_src(r, c, v.to_string());
            }
        }
        // A second table, so the duplicate-name case has a real collision.
        wb.sheets[0].set_cell_src(7, 0, "Key".to_string());
        wb.sheets[0].set_cell_src(7, 1, "Val".to_string());
        wb.sheets[0].set_cell_src(8, 0, "k".to_string());
        wb.sheets[0].set_cell_src(8, 1, "1".to_string());
        wb.sheets[0].set_cell_src(0, 4, "=SUM(Sales[Amount])".to_string());
        wb.evaluate().unwrap();
        wb.add_table(None, "Sales", 0, 0, 3, 2, true, false)
            .unwrap();
        wb.add_table(None, "Other", 7, 0, 8, 1, true, false)
            .unwrap();
        wb.evaluate().unwrap();
        wb
    }

    fn table_probe(setup_and_expr: &str) -> String {
        let (setup, expr) = match setup_and_expr.rsplit_once("::") {
            Some((s, e)) => (s.replace("\\n", "\n    "), e.to_string()),
            None => (String::new(), setup_and_expr.to_string()),
        };
        let source = format!(
            "Attribute VB_Name = \"H\"\n\
             Function Gen()\n    \
             Dim ws As Worksheet, wb As Workbook\n    \
             Set wb = ThisWorkbook\n    \
             Set ws = wb.Worksheets(\"Sheet1\")\n    \
             Dim lo, v, s\n    \
             {setup}\n    \
             Gen = {expr}\nEnd Function\n"
        );
        let mut wb = table_fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "H".to_string(),
            crate::core::VbaModuleKind::Standard,
            source,
            None,
        )
        .unwrap();
        match wb.run_macro(Some("H"), "Gen", &[]) {
            Ok(out) => format!("{}|{}", out.type_name, out.value.unwrap_or_default()),
            Err(crate::Error::VbaRuntime { number, .. }) => format!("ERR|{number}"),
            Err(e) => format!("FAIL|{e}"),
        }
    }

    #[test]
    fn table_objects_report_the_type_names_excel_reports() {
        assert_eq!(
            table_probe("TypeName(ws.ListObjects)"),
            "String|ListObjects"
        );
        assert_eq!(
            table_probe("TypeName(ws.ListObjects(1))"),
            "String|ListObject"
        );
        assert_eq!(
            table_probe(r#"TypeName(ws.ListObjects("Sales"))"#),
            "String|ListObject"
        );
        assert_eq!(
            table_probe("TypeName(ws.ListObjects(1).ListColumns)"),
            "String|ListColumns"
        );
        assert_eq!(
            table_probe("TypeName(ws.ListObjects(1).ListRows)"),
            "String|ListRows"
        );
        assert_eq!(table_probe("CStr(ws.ListObjects.Count)"), "String|2");
        assert_eq!(table_probe("ws.ListObjects(1).Name"), "String|Sales");
    }

    #[test]
    fn a_tables_ranges_are_the_ones_excel_reports() {
        assert_eq!(
            table_probe("ws.ListObjects(1).Range.Address"),
            "String|$A$1:$C$4"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).HeaderRowRange.Address"),
            "String|$A$1:$C$1"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).DataBodyRange.Address"),
            "String|$A$2:$C$4"
        );
        // No totals row: `Nothing`, not an empty range and not an error.
        assert_eq!(
            table_probe("TypeName(ws.ListObjects(1).TotalsRowRange)"),
            "String|Nothing"
        );
        assert_eq!(
            table_probe("CStr(ws.ListObjects(1).ShowTotals)"),
            "String|False"
        );
        assert_eq!(
            table_probe("CStr(ws.ListObjects(1).ShowHeaders)"),
            "String|True"
        );
    }

    #[test]
    fn table_columns_and_rows_address_what_excel_addresses() {
        assert_eq!(
            table_probe("CStr(ws.ListObjects(1).ListColumns.Count)"),
            "String|3"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).ListColumns(1).Name"),
            "String|Region"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).ListColumns(3).Name"),
            "String|Amount"
        );
        // A column's `.Range` includes the header row; its `.DataBodyRange`
        // does not.
        assert_eq!(
            table_probe("ws.ListObjects(1).ListColumns(3).Range.Address"),
            "String|$C$1:$C$4"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).ListColumns(3).DataBodyRange.Address"),
            "String|$C$2:$C$4"
        );
        assert_eq!(
            table_probe(r#"ws.ListObjects(1).ListColumns("Amount").Range.Address"#),
            "String|$C$1:$C$4"
        );
        assert_eq!(
            table_probe("CStr(ws.ListObjects(1).ListColumns(3).Index)"),
            "String|3"
        );
        // `ListRows` counts data rows only.
        assert_eq!(
            table_probe("CStr(ws.ListObjects(1).ListRows.Count)"),
            "String|3"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).ListRows(1).Range.Address"),
            "String|$A$2:$C$2"
        );
    }

    #[test]
    fn a_range_knows_which_table_it_is_in() {
        assert_eq!(
            table_probe(r#"TypeName(ws.Range("A2").ListObject)"#),
            "String|ListObject"
        );
        assert_eq!(
            table_probe(r#"ws.Range("A2").ListObject.Name"#),
            "String|Sales"
        );
        // Outside every table: `Nothing`.
        assert_eq!(
            table_probe(r#"TypeName(ws.Range("G1").ListObject)"#),
            "String|Nothing"
        );
    }

    #[test]
    fn a_missing_table_is_error_9_not_1004() {
        // Measured, and worth pinning: every *lookup* failure is 9, while a
        // duplicate name on assignment is 1004. The two are easy to conflate.
        assert_eq!(table_probe(r#"ws.ListObjects("nope").Name"#), "ERR|9");
        assert_eq!(table_probe("ws.ListObjects(5).Name"), "ERR|9");
        assert_eq!(
            table_probe("ws.ListObjects(1).ListColumns(9).Name"),
            "ERR|9"
        );
        assert_eq!(
            table_probe("ws.ListObjects(1).ListRows(9).Range.Address"),
            "ERR|9"
        );
    }

    #[test]
    fn renaming_a_table_cascades_into_formula_text() {
        // The reason `.Name` routes through `WorkbookManager` rather than
        // writing the field: measured, Excel rewrites the formula too.
        assert_eq!(
            table_probe(
                r#"ws.ListObjects(1).Name = "Revenue" :: ws.ListObjects(1).Name & "|" & ws.Range("E1").Formula"#
            ),
            "String|Revenue|=SUM(Revenue[Amount])"
        );
        // A name already taken workbook-wide is 1004.
        assert_eq!(
            table_probe(r#"ws.ListObjects(1).Name = "Other" :: "unreachable""#),
            "ERR|1004"
        );
        // Renaming a column cascades the same way.
        assert_eq!(
            table_probe(
                r#"ws.ListObjects(1).ListColumns(3).Name = "Total" :: ws.ListObjects(1).ListColumns(3).Name & "|" & ws.Range("E1").Formula"#
            ),
            "String|Total|=SUM(Sales[Total])"
        );
    }

    #[test]
    fn list_rows_add_grows_the_table_and_returns_the_new_row() {
        assert_eq!(
            table_probe(
                "Set lo = ws.ListObjects(1)\\nlo.ListRows.Add :: lo.Range.Address & \"|\" & CStr(lo.ListRows.Count)"
            ),
            "String|$A$1:$C$5|4"
        );
        assert_eq!(
            table_probe("Set lo = ws.ListObjects(1)\\nSet v = lo.ListRows.Add :: TypeName(v)"),
            "String|ListRow"
        );
        assert_eq!(
            table_probe("Set v = ws.ListObjects(1).ListRows.Add :: v.Range.Address"),
            "String|$A$5:$C$5"
        );
        // Adding at a position pushes the existing rows down, and the new row
        // is blank.
        assert_eq!(
            table_probe(
                r#"Set v = ws.ListObjects(1).ListRows.Add(1) :: v.Range.Address & "/" & CStr(ws.Range("A2").Value)"#
            ),
            "String|$A$2:$C$2/"
        );
    }

    #[test]
    fn adding_a_table_row_shifts_only_the_tables_own_columns() {
        // The measurement that forced the band operation to exist: content
        // below the table in its columns moves, content beside it does not.
        assert_eq!(
            table_probe(
                "ws.Range(\"A8\").Value = \"BELOW\"\\n\
                 ws.Range(\"E2\").Value = \"BESIDE\"\\n\
                 ws.ListObjects(1).ListRows.Add :: \
                 CStr(ws.Range(\"A9\").Value) & \"/\" & CStr(ws.Range(\"E2\").Value)"
            ),
            "String|BELOW/BESIDE"
        );
    }

    #[test]
    fn a_table_with_no_data_rows_has_no_data_body_range() {
        // Measured: deleting the only data row leaves the extent alone --
        // the table keeps its insert-row placeholder -- while `DataBodyRange`
        // becomes `Nothing` and `ListRows.Count` becomes 0.
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: TypeName(ws.ListObjects("Other").DataBodyRange)"#
            ),
            "String|Nothing"
        );
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: CStr(ws.ListObjects("Other").DataBodyRange Is Nothing)"#
            ),
            "String|True"
        );
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: CStr(ws.ListObjects("Other").ListRows.Count)"#
            ),
            "String|0"
        );
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: ws.ListObjects("Other").Range.Address"#
            ),
            "String|$A$8:$B$9"
        );
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: ws.ListObjects("Other").HeaderRowRange.Address"#
            ),
            "String|$A$8:$B$8"
        );
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: CStr(ws.ListObjects("Other").ListColumns.Count)"#
            ),
            "String|2"
        );
        assert_eq!(
            table_probe(
                r#"ws.ListObjects("Other").ListRows(1).Delete :: TypeName(ws.ListObjects("Other").ListColumns(1).DataBodyRange)"#
            ),
            "String|Nothing"
        );
        // Adding a row back reuses the reserved row rather than growing the
        // extent -- measured, `$A$8:$B$9` again, not `$A$8:$B$10`.
        assert_eq!(
            table_probe(
                "ws.ListObjects(\"Other\").ListRows(1).Delete\\n\
                 Set lo = ws.ListObjects(\"Other\")\\n\
                 lo.ListRows.Add :: lo.Range.Address & \"|\" & lo.DataBodyRange.Address"
            ),
            "String|$A$8:$B$9|$A$9:$B$9"
        );
    }

    #[test]
    fn show_totals_adds_a_row_to_the_table() {
        assert_eq!(
            table_probe(
                "Set lo = ws.ListObjects(1)\\nlo.ShowTotals = True :: \
                 lo.Range.Address & \"|\" & lo.TotalsRowRange.Address"
            ),
            "String|$A$1:$C$5|$A$5:$C$5"
        );
    }

    #[test]
    fn writing_through_a_table_range_reaches_the_cells() {
        assert_eq!(
            table_probe(
                r#"ws.ListObjects(1).DataBodyRange.Cells(1, 3).Value = 999 :: CStr(ws.Range("C2").Value)"#
            ),
            "String|999"
        );
    }

    // ---------------------------------------------------------------
    // Pivot tables.
    //
    // `pivot_fixture` is the grid `fuzz/vba_pivot_probe.py` builds, and the
    // pivot has the same shape: Region down the rows, Sum of Amount as the
    // value, Product as a page field, drawn at F1. Every expectation is what
    // Excel for Mac 16.112 returned for the same expression.
    //
    // Note the error number: **1004** throughout, where the `ListObjects`
    // collection uses 9 for the same shape of mistake. Both measured.
    // ---------------------------------------------------------------

    fn pivot_fixture() -> WorkbookManager {
        use crate::core::pivot::{PivotAggregation, PivotArea};
        let mut wb = WorkbookManager {
            sheets: vec![Sheet::new(SheetInit {
                name: Some("Sheet1".to_string()),
                rows: 8,
                cols: 8,
                ..Default::default()
            })],
            charts: Vec::new(),
            pivot_tables: Vec::new(),
            vba_project: None,
            locale: crate::core::locale::Locale::default(),
        };
        let rows: [[&str; 3]; 7] = [
            ["Region", "Product", "Amount"],
            ["East", "Widget", "10"],
            ["East", "Gadget", "5"],
            ["West", "Widget", "30"],
            ["West", "Gadget", "40"],
            ["North", "Doohickey", "7"],
            ["North", "Widget", "3"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                wb.sheets[0].set_cell_src(r, c, v.to_string());
            }
        }
        wb.evaluate().unwrap();
        wb.add_pivot_table_from_range("P1", None, 0, 0, 6, 2, None, 0, 5, true, true)
            .unwrap();
        wb.add_pivot_field("P1", PivotArea::Row, "Region", None)
            .unwrap();
        wb.add_pivot_field(
            "P1",
            PivotArea::Value,
            "Amount",
            Some(PivotAggregation::Sum),
        )
        .unwrap();
        wb.add_pivot_field("P1", PivotArea::Filter, "Product", None)
            .unwrap();
        wb.refresh_pivot_table("P1").unwrap();
        wb
    }

    fn pivot_probe(setup_and_expr: &str) -> String {
        let (setup, expr) = match setup_and_expr.rsplit_once("::") {
            Some((s, e)) => (s.replace("\\n", "\n    "), e.to_string()),
            None => (String::new(), setup_and_expr.to_string()),
        };
        let source = format!(
            "Attribute VB_Name = \"H\"\n\
             Function Gen()\n    \
             Dim ws As Worksheet, wb As Workbook, pt As PivotTable\n    \
             Set wb = ThisWorkbook\n    \
             Set ws = wb.Worksheets(\"Sheet1\")\n    \
             Set pt = ws.PivotTables(\"P1\")\n    \
             Dim v, s\n    \
             {setup}\n    \
             Gen = {expr}\nEnd Function\n"
        );
        let mut wb = pivot_fixture();
        wb.ensure_vba_project().unwrap();
        wb.add_vba_module(
            "H".to_string(),
            crate::core::VbaModuleKind::Standard,
            source,
            None,
        )
        .unwrap();
        match wb.run_macro(Some("H"), "Gen", &[]) {
            Ok(out) => format!("{}|{}", out.type_name, out.value.unwrap_or_default()),
            Err(crate::Error::VbaRuntime { number, .. }) => format!("ERR|{number}"),
            Err(e) => format!("FAIL|{e}"),
        }
    }

    #[test]
    fn pivot_objects_report_the_type_names_excel_reports() {
        assert_eq!(
            pivot_probe("TypeName(ws.PivotTables)"),
            "String|PivotTables"
        );
        assert_eq!(
            pivot_probe("TypeName(ws.PivotTables(1))"),
            "String|PivotTable"
        );
        assert_eq!(
            pivot_probe(r#"TypeName(ws.PivotTables("P1"))"#),
            "String|PivotTable"
        );
        assert_eq!(
            pivot_probe(r#"TypeName(pt.PivotFields("Product"))"#),
            "String|PivotField"
        );
        assert_eq!(pivot_probe("CStr(ws.PivotTables.Count)"), "String|1");
        assert_eq!(pivot_probe("pt.Name"), "String|P1");
        // One entry per *source column*, whatever area it occupies.
        assert_eq!(pivot_probe("CStr(pt.PivotFields.Count)"), "String|3");
    }

    #[test]
    fn a_missing_pivot_is_1004_not_the_9_a_missing_table_gives() {
        assert_eq!(pivot_probe(r#"ws.PivotTables("nope").Name"#), "ERR|1004");
        assert_eq!(pivot_probe("ws.PivotTables(5).Name"), "ERR|1004");
        assert_eq!(
            pivot_probe(r#"pt.PivotFields("nope").Orientation"#),
            "ERR|1004"
        );
    }

    #[test]
    fn orientation_is_the_area_and_a_data_field_reports_hidden() {
        // xlRowField 1, xlPageField 3, xlHidden 0. The surprise is `Amount`:
        // it is the value field, yet reports 0 rather than xlDataField --
        // aggregating a column leaves the *source* field unoriented.
        assert_eq!(
            pivot_probe(r#"CStr(pt.PivotFields("Region").Orientation)"#),
            "String|1"
        );
        assert_eq!(
            pivot_probe(r#"CStr(pt.PivotFields("Product").Orientation)"#),
            "String|3"
        );
        assert_eq!(
            pivot_probe(r#"CStr(pt.PivotFields("Amount").Orientation)"#),
            "String|0"
        );
    }

    #[test]
    fn current_page_reads_all_until_exactly_one_item_is_selected() {
        assert_eq!(
            pivot_probe(r#"pt.PivotFields("Product").CurrentPage"#),
            "String|(All)"
        );
        assert_eq!(
            pivot_probe(
                r#"pt.PivotFields("Product").CurrentPage = "Widget" :: pt.PivotFields("Product").CurrentPage"#
            ),
            "String|Widget"
        );
        // Reading it on a field that is not a page field raises.
        assert_eq!(
            pivot_probe(r#"pt.PivotFields("Region").CurrentPage"#),
            "ERR|1004"
        );
        // A value the field does not have raises rather than blanking the
        // grid with an empty selection.
        assert_eq!(
            pivot_probe(r#"pt.PivotFields("Product").CurrentPage = "Nonesuch" :: "unreachable""#),
            "ERR|1004"
        );
    }

    #[test]
    fn setting_current_page_rerenders_without_an_explicit_refresh() {
        // The deliberate exception to "nothing recomputes a pivot
        // implicitly": measured, the grid is already filtered on the very
        // next read, with no `RefreshTable` in between. G5 is the second
        // data row -- North, which is 3 for Widget alone and 10 overall.
        assert_eq!(
            pivot_probe(
                r#"pt.PivotFields("Product").CurrentPage = "Widget" :: CStr(ws.Range("G5").Value)"#
            ),
            "String|3"
        );
        // An explicit refresh afterwards changes nothing.
        assert_eq!(
            pivot_probe(
                "pt.PivotFields(\"Product\").CurrentPage = \"Widget\"\\n\
                 pt.RefreshTable :: CStr(ws.Range(\"G5\").Value)"
            ),
            "String|3"
        );
        // The page-field cell shows the selection.
        assert_eq!(
            pivot_probe(
                r#"pt.PivotFields("Product").CurrentPage = "Widget" :: CStr(ws.Range("G1").Value)"#
            ),
            "String|Widget"
        );
        // And `(All)` puts it back.
        assert_eq!(
            pivot_probe(
                "pt.PivotFields(\"Product\").CurrentPage = \"Widget\"\\n\
                 pt.PivotFields(\"Product\").CurrentPage = \"(All)\" :: \
                 pt.PivotFields(\"Product\").CurrentPage & \"/\" & CStr(ws.Range(\"G5\").Value)"
            ),
            "String|(All)/10"
        );
    }

    #[test]
    fn refresh_table_returns_true() {
        assert_eq!(pivot_probe("TypeName(pt.RefreshTable)"), "String|Boolean");
        assert_eq!(pivot_probe("CStr(pt.RefreshTable)"), "String|True");
    }

    #[test]
    fn table_range_1_is_the_grid_and_2_includes_the_page_rows() {
        // Measured for a pivot at F1 with one filter field: the filter row
        // and the blank under it sit above the grid.
        assert_eq!(pivot_probe("pt.TableRange1.Address"), "String|$F$3:$G$7");
        assert_eq!(pivot_probe("pt.TableRange2.Address"), "String|$F$1:$G$7");
    }
}
