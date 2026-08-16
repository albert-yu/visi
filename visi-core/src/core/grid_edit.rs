//! What a row or column insert/delete does to everything that holds a
//! coordinate.
//!
//! Inserting a row above a formula does not just move the formula down: the
//! references *inside* it have to move too, or `=A3` keeps pointing at row 3
//! after the value it meant slid to row 4. Excel rewrites the formula text;
//! so does this module, over [`CompiledFormula`], whose references are
//! already numeric `(row, col)` pairs rather than text that would have to be
//! re-parsed.
//!
//! The same geometry governs the other things that hold coordinates -- Excel
//! Table rectangles and pivot source/destination ranges -- so the primitives
//! ([`shift_point`], [`shift_span`]) are shared rather than reimplemented per
//! caller.
//!
//! # The rules, and why they are these rules
//!
//! Every rule below was **measured against Microsoft Excel**, not recalled --
//! `fuzz/grid_edit_probe.py` drives Excel through the same edits and prints
//! its formula text next to visi's. Re-run it rather than "correcting" one of
//! these from memory; the first two are the ones most often remembered
//! backwards.
//!
//! - **A `$` does not pin a reference against a structural edit.** `$A$3`
//!   shifts on a row insert exactly as `A3` does. Absolute addressing governs
//!   fill and copy, not this. So a reference's `RefType` is deliberately
//!   never consulted below.
//! - **Inserting at a span's first row moves the span; inserting inside it
//!   grows it.** `SUM(A2:A4)` with a row inserted at 2 becomes `SUM(A3:A5)`,
//!   but with a row inserted at 3 becomes `SUM(A2:A5)`. Both fall out of
//!   shifting the two endpoints independently.
//! - **Deleting inside a span shrinks it; deleting all of it is `#REF!`.**
//!   `SUM(A2:A4)` less row 3 is `SUM(A2:A3)`; `SUM(A2:A2)` less row 2 is
//!   `SUM(#REF!)`. A single-cell reference is just a span of one, so it needs
//!   no separate rule.
//! - **`#REF!` replaces the reference, not the formula.** `=A3+1` less row 3
//!   is `=#REF!+1`, which still evaluates -- to `#REF!`, by ordinary error
//!   propagation.

use crate::core::formula::{CompiledFormula, FormulaPart};

/// Which axis a structural edit runs along.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Axis {
    /// Rows were inserted or deleted, moving cells vertically.
    Row,
    /// Columns were inserted or deleted, moving cells horizontally.
    Col,
}

/// A row or column insert/delete on one sheet.
///
/// Carries the sheet it happened on because the rewrite runs workbook-wide:
/// a formula on `Sheet2` referring to `Sheet1!A3` has to move when a row is
/// inserted on `Sheet1`, and must *not* move when one is inserted on
/// `Sheet2`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct GridEdit {
    /// The sheet whose grid changed, by stable id.
    pub sheet_id: u64,
    /// Whether rows or columns moved.
    pub axis: Axis,
    /// First index inserted or deleted, 0-based.
    pub at: usize,
    /// How many were inserted or deleted; never zero.
    pub count: usize,
    /// `true` for an insert, `false` for a delete.
    pub insert: bool,
    /// Restricts the edit to an inclusive column range, for Excel's
    /// *Insert / Delete cells, shift down / up* over a band -- which is what
    /// `ListRows.Add` actually is, not a row insert.
    ///
    /// `None` is a whole-row edit. Only meaningful with [`Axis::Row`]; the
    /// mirrored column-axis form (a shift-right over a *row* band) is not
    /// implemented, and [`GridEdit::band_rows`] is the only constructor that
    /// sets this.
    ///
    /// Measured: **a reference moves if and only if its columns lie entirely
    /// inside the band**, and then the ordinary rules apply unchanged. A
    /// range straddling the band's edge does not move at all -- it cannot
    /// both shift and not shift, and Excel resolves that by leaving it
    /// alone. See `fuzz/band_insert_probe.py`.
    pub band: Option<(usize, usize)>,
}

impl GridEdit {
    /// A single row inserted before `at`.
    pub fn insert_row(sheet_id: u64, at: usize) -> Self {
        Self {
            sheet_id,
            axis: Axis::Row,
            at,
            count: 1,
            insert: true,
            band: None,
        }
    }

    /// A single row deleted at `at`.
    pub fn delete_row(sheet_id: u64, at: usize) -> Self {
        Self {
            sheet_id,
            axis: Axis::Row,
            at,
            count: 1,
            insert: false,
            band: None,
        }
    }

    /// A single column inserted before `at`.
    pub fn insert_col(sheet_id: u64, at: usize) -> Self {
        Self {
            sheet_id,
            axis: Axis::Col,
            at,
            count: 1,
            insert: true,
            band: None,
        }
    }

    /// A single column deleted at `at`.
    pub fn delete_col(sheet_id: u64, at: usize) -> Self {
        Self {
            sheet_id,
            axis: Axis::Col,
            at,
            count: 1,
            insert: false,
            band: None,
        }
    }

    /// Rows inserted or deleted within an inclusive column band, which is
    /// Excel's *Insert cells, shift down* rather than a row insert.
    pub fn band_rows(
        sheet_id: u64,
        at: usize,
        count: usize,
        first_col: usize,
        last_col: usize,
        insert: bool,
    ) -> Self {
        Self {
            sheet_id,
            axis: Axis::Row,
            at,
            count,
            insert,
            band: Some((first_col, last_col)),
        }
    }

    /// Whether an inclusive column span lies entirely inside the band, which
    /// is the whole test for whether a reference moves. A whole-row edit has
    /// no band and so covers everything.
    ///
    /// Also the test for whether an Excel Table or pivot rectangle moves --
    /// same rule, since a table straddling the band's edge is in exactly the
    /// position a straddling range reference is.
    pub(crate) fn covers_columns(&self, first_col: usize, last_col: usize) -> bool {
        match self.band {
            None => true,
            Some((lo, hi)) => lo <= first_col && last_col <= hi,
        }
    }

    /// Where a single index on this edit's axis ends up, or `None` if it was
    /// deleted.
    fn point(&self, index: usize) -> Option<usize> {
        shift_point(index, self.at, self.count, self.insert)
    }

    /// Where an inclusive span on this edit's axis ends up, or `None` if all
    /// of it was deleted.
    fn span(&self, start: usize, end: usize) -> Option<(usize, usize)> {
        shift_span(start, end, self.at, self.count, self.insert)
    }
}

/// Where a single index ends up after `count` rows or columns are inserted
/// before `at`, or deleted starting at `at`.
///
/// `None` means the index itself was deleted -- the caller turns that into
/// `#REF!` or drops the object, depending on what holds the coordinate.
pub(crate) fn shift_point(index: usize, at: usize, count: usize, insert: bool) -> Option<usize> {
    if insert {
        Some(if index >= at { index + count } else { index })
    } else if index < at {
        Some(index)
    } else if index < at + count {
        None
    } else {
        Some(index - count)
    }
}

/// Where an inclusive `start..=end` span ends up after the same edit.
///
/// `None` means every index in the span was deleted. A partly-deleted span
/// survives as the part that is left, which is how `SUM(A2:A4)` becomes
/// `SUM(A2:A3)` rather than `#REF!` when row 3 goes.
///
/// `end` must be a real index. Callers holding a compiled formula have to
/// screen out the unbounded-row sentinel (`usize::MAX`, as in `A:C`) first,
/// which is what the `debug_assert` is here to catch.
pub(crate) fn shift_span(
    start: usize,
    end: usize,
    at: usize,
    count: usize,
    insert: bool,
) -> Option<(usize, usize)> {
    debug_assert!(end < usize::MAX, "unbounded span reached shift_span");
    if insert {
        // Each endpoint moves on its own, which is what makes an insert at
        // the span's first index move the span and an insert inside it grow
        // the span, with no special case for either.
        let new_start = if start >= at { start + count } else { start };
        let new_end = if end >= at { end + count } else { end };
        return Some((new_start, new_end));
    }

    // Work in half-open `[start, end + 1)` so the two ends use the same rule:
    // subtract however many deleted indices lie below each bound. An
    // inclusive `end` cannot, since the index it names may be one of the
    // deleted ones.
    let removed_below = |bound: usize| count.min(bound.saturating_sub(at));
    let new_start = start - removed_below(start);
    let new_end_exclusive = (end + 1) - removed_below(end + 1);
    if new_end_exclusive <= new_start {
        None
    } else {
        Some((new_start, new_end_exclusive - 1))
    }
}

/// Where a rectangle ends up after the edit, or `None` if the edit deleted
/// every row or every column of it.
///
/// Both bounds are inclusive. Used for the coordinate-holding objects that
/// are not formulas: Excel Table extents and pivot source ranges.
pub(crate) fn shift_rect(
    edit: &GridEdit,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
) -> Option<(usize, usize, usize, usize)> {
    match edit.axis {
        Axis::Row => {
            let (r0, r1) = edit.span(start_row, end_row)?;
            Some((r0, start_col, r1, end_col))
        }
        Axis::Col => {
            let (c0, c1) = edit.span(start_col, end_col)?;
            Some((start_row, c0, end_row, c1))
        }
    }
}

/// Rewrites a compiled formula's references for the edit, returning `None`
/// if nothing in it was affected.
///
/// `deleted_col_ids` are the ids of columns the edit is about to remove,
/// which is what a whole-column reference (`=SUM(B:B)`, held by column id
/// rather than by position) has to be checked against -- once the column is
/// gone there is nothing left to compare with.
pub(crate) fn shift_formula(
    formula: &CompiledFormula,
    edit: &GridEdit,
    deleted_col_ids: &[u64],
) -> Option<CompiledFormula> {
    let mut changed = false;
    let parts = formula
        .parts
        .iter()
        .map(|part| {
            let (next, part_changed) = shift_part(part, edit, deleted_col_ids);
            changed |= part_changed;
            next
        })
        .collect();
    changed.then_some(CompiledFormula { parts })
}

/// The text a reference collapses to once what it pointed at is gone.
///
/// Only the reference is replaced, so the rest of the formula still
/// evaluates and the error propagates through it the way Excel's does.
const REF_ERROR: &str = "#REF!";

fn shift_part(part: &FormulaPart, edit: &GridEdit, deleted_col_ids: &[u64]) -> (FormulaPart, bool) {
    let unchanged = || (part.clone(), false);
    let broken = || (FormulaPart::Text(REF_ERROR.to_string()), true);

    match part {
        FormulaPart::Text(_) | FormulaPart::StructuredReference { .. } => unchanged(),

        // A structured reference names its table and column, and a table
        // tracks its own extent, so it needs no coordinate fix-up here --
        // shifting the `ExcelTable` rectangle is what keeps it correct.
        FormulaPart::ColumnReference { sheet_id, col_id } => {
            if *sheet_id != edit.sheet_id {
                unchanged()
            } else if deleted_col_ids.contains(col_id) {
                broken()
            } else {
                // Held by id, so an insert or a delete elsewhere on the sheet
                // moves the column without changing what the reference means.
                unchanged()
            }
        }

        FormulaPart::SheetReference {
            sheet_id,
            row,
            col,
            row_ref_type,
            col_ref_type,
        } => {
            if *sheet_id != edit.sheet_id || !edit.covers_columns(*col, *col) {
                return unchanged();
            }
            let (new_row, new_col) = match edit.axis {
                Axis::Row => match edit.point(*row) {
                    Some(r) => (r, *col),
                    None => return broken(),
                },
                Axis::Col => match edit.point(*col) {
                    Some(c) => (*row, c),
                    None => return broken(),
                },
            };
            if (new_row, new_col) == (*row, *col) {
                return unchanged();
            }
            (
                FormulaPart::SheetReference {
                    sheet_id: *sheet_id,
                    row: new_row,
                    col: new_col,
                    row_ref_type: *row_ref_type,
                    col_ref_type: *col_ref_type,
                },
                true,
            )
        }

        FormulaPart::RangeReference {
            sheet_id,
            start_row,
            start_col,
            end_row,
            end_col,
            start_row_ref_type,
            start_col_ref_type,
            end_row_ref_type,
            end_col_ref_type,
        } => {
            if *sheet_id != edit.sheet_id || !edit.covers_columns(*start_col, *end_col) {
                return unchanged();
            }
            // `A:C` compiles to a range whose `end_row` is the unbounded
            // sentinel (see `parser::serialize_formula`, which renders it back
            // without any row part). It already covers every row, so a row
            // edit leaves it alone -- and must, since `end_row + 1` would
            // overflow. A *column* edit still moves its column bounds.
            if *end_row == usize::MAX && edit.axis == Axis::Row {
                return unchanged();
            }
            let rect = match shift_rect(edit, *start_row, *start_col, *end_row, *end_col) {
                Some(rect) => rect,
                None => return broken(),
            };
            if rect == (*start_row, *start_col, *end_row, *end_col) {
                return unchanged();
            }
            let (r0, c0, r1, c1) = rect;
            (
                FormulaPart::RangeReference {
                    sheet_id: *sheet_id,
                    start_row: r0,
                    start_col: c0,
                    end_row: r1,
                    end_col: c1,
                    start_row_ref_type: *start_row_ref_type,
                    start_col_ref_type: *start_col_ref_type,
                    end_row_ref_type: *end_row_ref_type,
                    end_col_ref_type: *end_col_ref_type,
                },
                true,
            )
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::RefType;

    // Every case below names the Excel behaviour it encodes. `at` is 0-based,
    // so "insert at 1" is Excel's "insert before row 2".

    #[test]
    fn an_insert_moves_what_is_at_or_below_it() {
        assert_eq!(shift_point(0, 1, 1, true), Some(0));
        assert_eq!(shift_point(1, 1, 1, true), Some(2));
        assert_eq!(shift_point(5, 1, 3, true), Some(8));
    }

    #[test]
    fn a_delete_removes_its_own_indices_and_pulls_the_rest_up() {
        assert_eq!(shift_point(0, 1, 1, false), Some(0));
        assert_eq!(shift_point(1, 1, 1, false), None);
        assert_eq!(shift_point(2, 1, 1, false), Some(1));
        assert_eq!(shift_point(4, 1, 3, false), Some(1));
        assert_eq!(shift_point(3, 1, 3, false), None);
    }

    #[test]
    fn inserting_at_a_spans_start_moves_it_and_inserting_inside_grows_it() {
        // SUM(A2:A4) -> SUM(A3:A5): insert at the first row moves the span.
        assert_eq!(shift_span(1, 3, 1, 1, true), Some((2, 4)));
        // SUM(A2:A4) -> SUM(A2:A5): insert inside grows it.
        assert_eq!(shift_span(1, 3, 2, 1, true), Some((1, 4)));
        // Insert below the span leaves it alone.
        assert_eq!(shift_span(1, 3, 4, 1, true), Some((1, 3)));
    }

    #[test]
    fn deleting_inside_a_span_shrinks_it() {
        // SUM(A2:A4) -> SUM(A2:A3).
        assert_eq!(shift_span(1, 3, 2, 1, false), Some((1, 2)));
        // Deleting the span's first row keeps the start where it is.
        assert_eq!(shift_span(1, 3, 1, 1, false), Some((1, 2)));
        // Deleting above the span slides the whole thing up.
        assert_eq!(shift_span(1, 3, 0, 1, false), Some((0, 2)));
        // Deleting below it changes nothing.
        assert_eq!(shift_span(1, 3, 4, 1, false), Some((1, 3)));
    }

    #[test]
    fn a_span_deleted_in_full_is_gone() {
        // A single-cell reference is a span of one.
        assert_eq!(shift_span(2, 2, 2, 1, false), None);
        // A multi-row span entirely inside the deleted run.
        assert_eq!(shift_span(2, 4, 1, 5, false), None);
        // Overlapping only part of it survives as that part.
        assert_eq!(shift_span(2, 4, 3, 5, false), Some((2, 2)));
    }

    #[test]
    fn a_dollar_sign_does_not_pin_a_reference_against_a_structural_edit() {
        // $A$3 and A3 shift identically; the ref types ride along untouched.
        let formula = CompiledFormula {
            parts: vec![
                FormulaPart::Text("=".to_string()),
                FormulaPart::SheetReference {
                    sheet_id: 1,
                    row: 2,
                    col: 0,
                    row_ref_type: RefType::Absolute,
                    col_ref_type: RefType::Absolute,
                },
            ],
        };
        let shifted = shift_formula(&formula, &GridEdit::insert_row(1, 0), &[]).unwrap();
        assert_eq!(
            shifted.parts[1],
            FormulaPart::SheetReference {
                sheet_id: 1,
                row: 3,
                col: 0,
                row_ref_type: RefType::Absolute,
                col_ref_type: RefType::Absolute,
            }
        );
    }

    #[test]
    fn an_edit_on_another_sheet_leaves_a_reference_alone() {
        let formula = CompiledFormula {
            parts: vec![FormulaPart::SheetReference {
                sheet_id: 1,
                row: 2,
                col: 0,
                row_ref_type: RefType::Relative,
                col_ref_type: RefType::Relative,
            }],
        };
        assert!(shift_formula(&formula, &GridEdit::insert_row(2, 0), &[]).is_none());
    }

    #[test]
    fn only_the_deleted_reference_becomes_ref_error() {
        // `=A3+1` less row 3 is `=#REF!+1`, not a wholly destroyed formula.
        let formula = CompiledFormula {
            parts: vec![
                FormulaPart::Text("=".to_string()),
                FormulaPart::SheetReference {
                    sheet_id: 1,
                    row: 2,
                    col: 0,
                    row_ref_type: RefType::Relative,
                    col_ref_type: RefType::Relative,
                },
                FormulaPart::Text("+1".to_string()),
            ],
        };
        let shifted = shift_formula(&formula, &GridEdit::delete_row(1, 2), &[]).unwrap();
        assert_eq!(
            shifted.parts,
            vec![
                FormulaPart::Text("=".to_string()),
                FormulaPart::Text("#REF!".to_string()),
                FormulaPart::Text("+1".to_string()),
            ]
        );
    }

    #[test]
    fn a_whole_column_reference_survives_a_move_and_breaks_on_its_own_deletion() {
        let formula = CompiledFormula {
            parts: vec![FormulaPart::ColumnReference {
                sheet_id: 1,
                col_id: 7,
            }],
        };
        // Held by id, so inserting a column beside it changes nothing.
        assert!(shift_formula(&formula, &GridEdit::insert_col(1, 0), &[]).is_none());
        // Deleting some other column likewise.
        assert!(shift_formula(&formula, &GridEdit::delete_col(1, 0), &[9]).is_none());
        // Deleting the column it names is the one case that breaks it.
        let shifted = shift_formula(&formula, &GridEdit::delete_col(1, 0), &[7]).unwrap();
        assert_eq!(shifted.parts, vec![FormulaPart::Text("#REF!".to_string())]);
    }

    #[test]
    fn an_unbounded_row_range_survives_a_row_edit_and_still_tracks_columns() {
        // `A:C` compiles to a range with `end_row: usize::MAX`. A row edit
        // must leave it alone -- it already covers every row, and the
        // arithmetic would overflow. A column edit still has to move it.
        let unbounded = |start_col, end_col| CompiledFormula {
            parts: vec![FormulaPart::RangeReference {
                sheet_id: 1,
                start_row: 0,
                start_col,
                end_row: usize::MAX,
                end_col,
                start_row_ref_type: RefType::Absolute,
                start_col_ref_type: RefType::Relative,
                end_row_ref_type: RefType::Absolute,
                end_col_ref_type: RefType::Relative,
            }],
        };
        assert!(shift_formula(&unbounded(0, 2), &GridEdit::delete_row(1, 0), &[]).is_none());
        assert!(shift_formula(&unbounded(0, 2), &GridEdit::insert_row(1, 0), &[]).is_none());

        let shifted = shift_formula(&unbounded(0, 2), &GridEdit::insert_col(1, 0), &[]).unwrap();
        assert_eq!(shifted.parts, unbounded(1, 3).parts);
    }

    #[test]
    fn a_row_edit_leaves_columns_alone_and_a_column_edit_leaves_rows_alone() {
        let rect = (2, 2, 4, 4);
        let (r0, c0, r1, c1) =
            shift_rect(&GridEdit::insert_row(1, 0), rect.0, rect.1, rect.2, rect.3).unwrap();
        assert_eq!((r0, c0, r1, c1), (3, 2, 5, 4));
        let (r0, c0, r1, c1) =
            shift_rect(&GridEdit::insert_col(1, 0), rect.0, rect.1, rect.2, rect.3).unwrap();
        assert_eq!((r0, c0, r1, c1), (2, 3, 4, 5));
    }

    #[test]
    fn a_band_edit_moves_only_references_wholly_inside_the_band() {
        // `ListRows.Add` is an insert over the table's columns, not a row
        // insert, so a formula beside the table must not move. Every case
        // here is from `fuzz/band_insert_probe.py` with the band A:C and the
        // insert at row 2 (0-based row 1, cols 0..=2).
        let edit = GridEdit::band_rows(1, 1, 1, 0, 2, true);
        let cell = |row, col| CompiledFormula {
            parts: vec![FormulaPart::SheetReference {
                sheet_id: 1,
                row,
                col,
                row_ref_type: RefType::Relative,
                col_ref_type: RefType::Relative,
            }],
        };
        // `=A5` -> `=A6`: inside the band, below the insert.
        assert_eq!(
            shift_formula(&cell(4, 0), &edit, &[]).unwrap().parts,
            cell(5, 0).parts
        );
        // `=A2` -> `=A3`: the insert point itself moves.
        assert_eq!(
            shift_formula(&cell(1, 0), &edit, &[]).unwrap().parts,
            cell(2, 0).parts
        );
        // `=A1`: above the insert, untouched.
        assert!(shift_formula(&cell(0, 0), &edit, &[]).is_none());
        // `=E5`: outside the band, untouched even though it is below.
        assert!(shift_formula(&cell(4, 4), &edit, &[]).is_none());
    }

    #[test]
    fn a_range_straddling_the_bands_edge_does_not_move_at_all() {
        // The case with no obvious answer, and the reason `covers_columns`
        // tests the *whole* span: `=SUM(A5:E5)` cannot both shift (its A part
        // is inside the band) and not shift (its E part is not), and Excel
        // resolves that by leaving it alone. Measured.
        let edit = GridEdit::band_rows(1, 1, 1, 0, 2, true);
        let range = |start_col, end_col| CompiledFormula {
            parts: vec![FormulaPart::RangeReference {
                sheet_id: 1,
                start_row: 4,
                start_col,
                end_row: 5,
                end_col,
                start_row_ref_type: RefType::Relative,
                start_col_ref_type: RefType::Relative,
                end_row_ref_type: RefType::Relative,
                end_col_ref_type: RefType::Relative,
            }],
        };
        // A5:E6 straddles the edge -- unchanged.
        assert!(shift_formula(&range(0, 4), &edit, &[]).is_none());
        // A5:C6 is wholly inside -- moves.
        let moved = shift_formula(&range(0, 2), &edit, &[]).unwrap();
        let FormulaPart::RangeReference {
            start_row, end_row, ..
        } = moved.parts[0]
        else {
            panic!("expected a range");
        };
        assert_eq!((start_row, end_row), (5, 6));
        // E5:F6 is wholly outside -- unchanged.
        assert!(shift_formula(&range(4, 5), &edit, &[]).is_none());
    }

    #[test]
    fn a_band_edit_grows_a_range_that_spans_its_insert_point() {
        // Inside the band the ordinary rules apply unchanged, which is the
        // point of reusing `shift_span`: `=SUM(A1:A6)` becomes `=SUM(A1:A7)`.
        let edit = GridEdit::band_rows(1, 1, 1, 0, 2, true);
        let formula = CompiledFormula {
            parts: vec![FormulaPart::RangeReference {
                sheet_id: 1,
                start_row: 0,
                start_col: 0,
                end_row: 5,
                end_col: 0,
                start_row_ref_type: RefType::Relative,
                start_col_ref_type: RefType::Relative,
                end_row_ref_type: RefType::Relative,
                end_col_ref_type: RefType::Relative,
            }],
        };
        let grown = shift_formula(&formula, &edit, &[]).unwrap();
        let FormulaPart::RangeReference {
            start_row, end_row, ..
        } = grown.parts[0]
        else {
            panic!("expected a range");
        };
        assert_eq!((start_row, end_row), (0, 6));
    }

    #[test]
    fn a_whole_column_reference_ignores_a_band_edit() {
        // Measured: `=SUM(A:A)` is unchanged by an insert inside A:C, since
        // it already spans every row.
        let edit = GridEdit::band_rows(1, 1, 1, 0, 2, true);
        let whole = CompiledFormula {
            parts: vec![FormulaPart::ColumnReference {
                sheet_id: 1,
                col_id: 7,
            }],
        };
        assert!(shift_formula(&whole, &edit, &[]).is_none());
    }
}
