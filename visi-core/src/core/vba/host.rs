//! The VBA host object model: Phase 2 of `docs/vba-macro-support.md`.
//!
//! This is the layer that lets a macro touch a workbook. Everything in it was
//! measured against Excel for Mac 16.112 with `fuzz/vba_host_probe.py` rather
//! than taken from documentation, for the same reason Phase 1's `Variant`
//! rules were: a careful reading of the docs already got one of those
//! backwards, and a host model that is subtly wrong produces wrong *numbers*
//! in a saved file, which is worse than an error.
//!
//! Four design decisions carry most of the weight.
//!
//! **An object is a value, not a pointer.** [`ObjRef`] holds ids and
//! coordinates, so it is `Clone` and never borrows the workbook. That is what
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

use crate::core::engine::{CellRef, ResultData, Sheet};
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
#[derive(Debug, Clone, PartialEq)]
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
    /// A rectangular range of cells.
    Range(RangeRef),
}

/// A `Range`: which sheet, which rectangle, and which object it is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeRef {
    /// Object identity, for `Is`. Fresh per constructed `Range`, preserved by
    /// copying -- see this module's doc comment for the measurements that
    /// forced this to exist rather than comparing the coordinates.
    pub token: u64,
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
            (ObjRef::Range(a), ObjRef::Range(b)) => a.token == b.token,
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

/// Error 1004 -- what Excel reports for a bad address, an out-of-sheet
/// `Offset`, and a `WorksheetFunction` call that fails. All measured.
fn app_defined(message: impl Into<String>) -> VbaError {
    VbaError::new(1004, message.into())
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
    /// Next `RangeRef::token`. A counter, not a hash of the coordinates --
    /// two ranges over the same cells must not be the same object.
    next_token: u64,
    /// The sheet an unqualified `Range(...)` / `Cells(...)` resolves against.
    active_sheet: u64,
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
            next_token: 1,
            active_sheet,
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

    fn new_range(&mut self, sheet_id: u64, row: u32, col: u32, height: u32, width: u32) -> ObjRef {
        self.next_token += 1;
        ObjRef::Range(RangeRef {
            token: self.next_token,
            sheet_id,
            row,
            col,
            height,
            width,
        })
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
            ObjRef::Range(r) => self.range_member(*r, name, args),
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
            (ObjRef::Range(r), "value" | "value2" | "formula") => {
                if !args.is_empty() {
                    return Err(unsupported(&format!("Range.{name} with arguments")));
                }
                self.write_range(*r, value)
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
            ObjRef::Range(r) => self.range_index(*r, args),
            _ => Err(unsupported(&format!("calling a {}", obj.type_name()))),
        }
    }

    /// The value an object stands for when it is used without `Set`.
    ///
    /// `x = ws.Range("A1")` reads the cell, and `MsgBox ws.Range("A1")` would
    /// too. Only `Range` has a default member in this scope.
    pub fn default_value(&mut self, obj: &ObjRef) -> VResult<Variant> {
        match obj {
            ObjRef::Range(r) => self.read_range(*r, false),
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
            ObjRef::Range(r) => self.write_range(*r, value),
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
            ObjRef::Range(r) => {
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
            _ => Err(unsupported(&format!("Worksheet.{name}"))),
        }
    }

    /// `Range("A1")`, `Range("A1:B2")` and `Range(cell1, cell2)`.
    fn resolve_range_args(&mut self, sheet_id: u64, args: &[Variant]) -> VResult<ObjRef> {
        let first = args.first().ok_or_else(VbaError::invalid_call)?;
        if args.len() >= 2 {
            let a = Self::corner(first)?;
            let b = Self::corner(&args[1])?;
            let (row, col) = (a.0.min(b.0), a.1.min(b.1));
            let (row2, col2) = (a.0.max(b.0), a.1.max(b.1));
            return Ok(self.new_range(sheet_id, row, col, row2 - row + 1, col2 - col + 1));
        }
        match first {
            Variant::Object(ObjRef::Range(r)) => {
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
    fn corner(v: &Variant) -> VResult<(u32, u32)> {
        match v {
            Variant::Object(ObjRef::Range(r)) => Ok((r.row, r.col)),
            other => {
                let text = other.to_vba_string()?;
                let (row, col, _, _) = parse_address(&text).ok_or_else(|| {
                    app_defined(format!("Method 'Range' of object failed: {text}"))
                })?;
                Ok((row, col))
            }
        }
    }

    fn range_member(&mut self, r: RangeRef, name: &str, args: &[Variant]) -> VResult<Variant> {
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
            // `name`, not the lowercased match binding: the message is
            // user-facing and `Range.interior` reads like a typo where
            // `Range.Interior` reads like the refusal it is.
            _ => Err(unsupported(&format!("Range.{name}"))),
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
                let o = o.clone();
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
        Ok(())
    }

    // -- the WorksheetFunction bridge --------------------------------------

    /// `Application.<name>`, which is either an object or a non-raising
    /// worksheet function.
    fn application_member(&mut self, name: &str, args: &[Variant]) -> VResult<Variant> {
        if name.eq_ignore_ascii_case("worksheetfunction") {
            return Ok(Variant::Object(ObjRef::WorksheetFunction));
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
            Variant::Object(ObjRef::Range(r)) => {
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
    if let (Some(inner), ResultData::String(value)) = (
        src.strip_prefix('"').and_then(|s| s.strip_suffix('"')),
        sheet.get_result_data(&cell),
    ) && inner == value
    {
        return value;
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
        // The one ordering question issue #57 asked to confirm rather than
        // assume. Excel walks left-to-right, then down.
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
    fn writing_an_infinity_stores_the_num_error_excel_stores() {
        // `^` is the one VBA operator that produces an infinity rather than
        // raising error 6, so a macro really can reach here. Excel has no
        // infinities: measured, assigning one leaves an error cell.
        //
        // Found by `fuzz/fuzz_vba.py` -- and only by its *cell* comparison,
        // since the case returned the same value on both sides while leaving
        // `-inf` in a cell here and `#NUM!` in Excel.
        assert_eq!(
            probe(
                "v = 1000\nws.Range(\"G9\").Value = (-2.5 ^ v) :: TypeName(ws.Range(\"G9\").Value)"
            ),
            "String|Error"
        );
        assert_eq!(
            probe("v = 1000\nws.Range(\"G9\").Value = (-2.5 ^ v) :: ws.Range(\"G9\").Text"),
            "String|#NUM!"
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
        // that this phase deliberately does not implement.
        for case in [
            "ws.Range(\"A1\").Interior.Color",
            "ws.ListObjects(1).Name",
            "wb.PivotTables(1).Name",
            "ws.Range(\"A1\").Font.Bold",
            "ws.Range(\"A1:B2\").Sort",
            "ws.Range(\"A1\").NumberFormat",
        ] {
            assert_eq!(probe(case), "ERR|438", "{case}");
        }
    }
}
