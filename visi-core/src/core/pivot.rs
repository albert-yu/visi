//! Pivot table definitions and the pure function that computes one.
//!
//! A [`PivotTable`] is a *definition*: where the records come from, which
//! fields go in the row, column, value and filter areas, and where the result
//! should land. [`compute_pivot`] turns that definition plus the workbook's
//! sheets into a [`PivotGrid`], a display-ready set of header and body rows.
//!
//! Computing a grid never touches a sheet. Writing one into cells is
//! `WorkbookManager::refresh_pivot_table`'s job, and -- as in Excel -- it only
//! happens when something asks for it: **nothing recomputes a pivot table
//! implicitly**, not `Sheet::commit` and not `WorkbookManager::evaluate`, so
//! editing the source data leaves the rendered grid stale until a refresh.
//! Every CRUD operation on a pivot definition refreshes explicitly afterward.
//!
//! Unlike an [`ExcelTable`](crate::core::table::ExcelTable), which is scoped
//! to one sheet, a pivot table is workbook-level: its source and destination
//! ranges may live on different sheets, so `WorkbookManager` owns the list.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

use crate::core::engine::{CellRef, ResultData, Sheet};

/// Where a `PivotTable` reads its source records from: either an existing
/// `ExcelTable` (looked up by name at compute time, so renames/resizes of
/// the table are picked up automatically on refresh) or a plain cell range
/// whose first row is treated as column headers.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum PivotSource {
    /// An `ExcelTable`, resolved by name on every refresh.
    Table {
        /// The table's name, matched case-insensitively workbook-wide.
        name: String,
    },
    /// A raw rectangular range, whose first row supplies the field names.
    Range {
        /// Sheet the range lives on.
        sheet_id: u64,
        /// First row of the range, 0-based, and the header row.
        start_row: usize,
        /// First column of the range, 0-based.
        start_col: usize,
        /// Last row of the range, 0-based and inclusive.
        end_row: usize,
        /// Last column of the range, 0-based and inclusive.
        end_col: usize,
    },
}

/// Matches the "Summarize value field by" choices Excel exposes for a data
/// field; the five most commonly used ones plus the numeric-only count.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum PivotAggregation {
    /// Total of the numeric values.
    Sum,
    /// How many non-blank values there are, text included.
    Count,
    /// How many values are numbers.
    CountNumbers,
    /// Mean of the numeric values.
    Average,
    /// Largest numeric value.
    Max,
    /// Smallest numeric value.
    Min,
}

impl PivotAggregation {
    /// The caption Excel uses for this aggregation in a value field's default
    /// label ("Sum of Amount").
    ///
    /// [`PivotAggregation::CountNumbers`] shares `Count`'s caption, which is
    /// why two such fields on one column collide and get disambiguated by
    /// [`value_field_labels`].
    pub fn label(&self) -> &'static str {
        match self {
            PivotAggregation::Sum => "Sum",
            // Excel's default value-field caption for "Count Numbers" is
            // "Count of <field>" -- identical to plain "Count" -- not
            // "Count Numbers of <field>"; there's no separate caption text
            // for it in Excel's own UI (confirmed via fuzz/fuzz_pivot.py
            // against real Excel).
            PivotAggregation::Count | PivotAggregation::CountNumbers => "Count",
            PivotAggregation::Average => "Average",
            PivotAggregation::Max => "Max",
            PivotAggregation::Min => "Min",
        }
    }

    /// Parses a user-supplied aggregation name, ignoring case, spaces,
    /// underscores and hyphens, and accepting the common short forms (`avg`,
    /// `countnums`, `maximum`). `None` if it names nothing.
    pub fn parse(s: &str) -> Option<Self> {
        match s.to_ascii_lowercase().replace(['_', '-', ' '], "").as_str() {
            "sum" => Some(Self::Sum),
            "count" => Some(Self::Count),
            "countnumbers" | "countnums" => Some(Self::CountNumbers),
            "average" | "avg" => Some(Self::Average),
            "max" | "maximum" => Some(Self::Max),
            "min" | "minimum" => Some(Self::Min),
            _ => None,
        }
    }
}

/// One field placed in the Row or Column area.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PivotField {
    /// Name of the source column to group by, matched against the header row.
    pub column: String,
    /// Whether a subtotal line is emitted for this field when it isn't the
    /// innermost field in its area (Excel's per-field "Subtotals" toggle).
    pub subtotal: bool,
}

impl PivotField {
    /// A field on `column` with subtotals enabled, Excel's default.
    pub fn new(column: impl Into<String>) -> Self {
        Self {
            column: column.into(),
            subtotal: true,
        }
    }
}

/// One field placed in the Values area.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PivotValueField {
    /// Name of the source column to aggregate, matched against the header row.
    pub column: String,
    /// How the column's values are summarized.
    pub aggregation: PivotAggregation,
    /// Overrides the default "Sum of Amount" caption. A custom name is used
    /// verbatim and takes no part in [`value_field_labels`]' disambiguation.
    pub custom_name: Option<String>,
}

impl PivotValueField {
    /// A value field on `column` with the default caption.
    pub fn new(column: impl Into<String>, aggregation: PivotAggregation) -> Self {
        Self {
            column: column.into(),
            aggregation,
            custom_name: None,
        }
    }

    /// This field's caption considered on its own, ignoring any collision
    /// with the pivot's other value fields. Use [`value_field_labels`] to
    /// caption a whole list the way Excel would.
    pub fn label(&self) -> String {
        self.custom_name
            .clone()
            .unwrap_or_else(|| format!("{} of {}", self.aggregation.label(), self.column))
    }
}

/// Default display labels for a pivot's whole value-field list, matching
/// Excel's own (surprisingly convoluted) disambiguation for repeated
/// source columns -- derived empirically against real Excel via
/// fuzz/fuzz_pivot.py plus direct probing (see the probe script referenced
/// in the PR that added this comment), since none of it is documented.
///
/// Two independent mechanisms are in play, both scoped per source column:
///
/// 1. **The "Sum" clone.** The *first* value field for a column that uses
///    the `Sum` aggregation causes Excel to silently clone that column
///    into a new pseudo-field ("Amount" -> "Amount2") for every value
///    field *after* it in the list (not before) -- regardless of their own
///    aggregation. A *second* `Sum` on the same column clones again
///    ("Amount2" -> "Amount3"), but non-`Sum` aggregations never trigger a
///    further clone; they just ride whatever clone slot is already active.
///    E.g. `[Sum, Max, Count]` on "Amount" -> `["Sum of Amount", "Max of
///    Amount2", "Count of Amount2"]` (both non-Sum fields share slot 2);
///    `[Sum, Sum, Count]` -> `["Sum of Amount", "Sum of Amount2", "Count
///    of Amount3"]` (the second Sum clones again). A column with *no* Sum
///    value field anywhere is never cloned at all.
/// 2. **Literal caption collision.** Independent of the above, if two
///    value fields end up wanting the exact same caption text, Excel still
///    has to disambiguate. If neither is in a Sum-cloned slot, it appends
///    a plain digit straight onto the column name (`"Count of Amount"`,
///    `"Count of Amount2"`, `"Count of Amount3"`, ...) -- this is also how
///    `CountNumbers` colliding with `Count` gets suffixed, since both
///    share the caption label "Count" (see `PivotAggregation::label`). If
///    the collision instead happens *inside* an already Sum-cloned slot
///    (two non-Sum fields sharing one clone with the same aggregation),
///    Excel instead appends an underscored counter to the *whole* already-
///    suffixed caption (`"Max of Amount2"`, `"Max of Amount2_2"`) rather
///    than incrementing the clone number again.
///
/// An explicit `custom_name` bypasses both mechanisms entirely -- it's
/// used as-is and doesn't consume a collision slot or trigger a clone.
pub fn value_field_labels(value_fields: &[PivotValueField]) -> Vec<String> {
    let mut clone_suffix: HashMap<&str, usize> = HashMap::new();
    let mut next_clone: HashMap<&str, usize> = HashMap::new();
    let mut label_counts: HashMap<String, usize> = HashMap::new();

    value_fields
        .iter()
        .map(|vf| {
            if let Some(name) = &vf.custom_name {
                return name.clone();
            }
            let agg_label = vf.aggregation.label();
            let in_clone_slot = clone_suffix.contains_key(vf.column.as_str());
            let base_column = match clone_suffix.get(vf.column.as_str()) {
                Some(n) => format!("{}{}", vf.column, n),
                None => vf.column.clone(),
            };
            let base_label = format!("{} of {}", agg_label, base_column);
            let count = label_counts.entry(base_label.clone()).or_insert(0);
            *count += 1;
            let label = if *count == 1 {
                base_label
            } else if in_clone_slot {
                format!("{}_{}", base_label, count)
            } else {
                format!("{} of {}{}", agg_label, vf.column, count)
            };
            if vf.aggregation == PivotAggregation::Sum {
                let assigned = *next_clone.entry(vf.column.as_str()).or_insert(2);
                next_clone.insert(vf.column.as_str(), assigned + 1);
                clone_suffix.insert(vf.column.as_str(), assigned);
            }
            label
        })
        .collect()
}

/// One field placed in the Filter (Page) area.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PivotFilterField {
    /// Name of the source column to filter on, matched against the header row.
    pub column: String,
    /// `None` means every value is allowed (no filtering applied yet).
    ///
    /// **Reconstructed on xlsx import**, resolved through the cache's
    /// `<sharedItems>` to plain value strings rather than kept as indices --
    /// which is what makes it safe. The indices are trusted only against the
    /// cache definition in the same file, which is self-consistent by
    /// construction, and a value that no longer exists in changed source data
    /// simply matches nothing.
    ///
    /// Two things do not survive, both because the format cannot hold them:
    ///
    /// - A selection covering *every* value marks nothing hidden, so it is
    ///   indistinguishable from no filter and reads back as `None`. The
    ///   grid is the same either way.
    /// - A filter on a column that is *also* a row or column field is lost
    ///   entirely: a pivot field carries one `axis`, so there is nowhere to
    ///   record it. Excel cannot express that config at all -- a field has
    ///   exactly one orientation there.
    ///
    /// Matching is case-insensitive, because the items themselves are merged
    /// that way; a selection naming `east` picks the merged `East` item.
    pub selected_values: Option<Vec<String>>,
    /// Whether the field is in Excel's *multi-select* page mode
    /// (`multipleItemSelectionAllowed` in the file) rather than its classic
    /// single-select one.
    ///
    /// The two differ in what the page-field cell says, which is observable:
    /// with one item chosen, multi-select shows `(Multiple Items)` while
    /// single-select shows the **item's own name**. Both measured -- the
    /// first through `PivotItems(x).Visible = False`, the second through
    /// `PivotField.CurrentPage = "Widget"`, which is what puts a field into
    /// single-select mode in the first place.
    ///
    /// Defaults to `true`, matching `set_pivot_filter` and the CLI, which
    /// select a set of values rather than one page.
    #[serde(default = "default_true")]
    pub multiple_selection: bool,
}

fn default_true() -> bool {
    true
}

impl PivotFilterField {
    /// A filter field on `column` with nothing filtered out yet.
    pub fn new(column: impl Into<String>) -> Self {
        Self {
            column: column.into(),
            selected_values: None,
            multiple_selection: true,
        }
    }
}

/// The area of a pivot table a field can be assigned to, used by the
/// add/remove-field CRUD operations.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum PivotArea {
    /// Groups down the left edge; adds to `PivotTable::row_fields`.
    Row,
    /// Groups across the top; adds to `PivotTable::col_fields`.
    Column,
    /// Aggregated data; adds to `PivotTable::value_fields`.
    Value,
    /// Restricts which source records take part; adds to
    /// `PivotTable::filter_fields`.
    Filter,
}

/// A pivot table definition: a summary of `source`, grouped by `row_fields`
/// nested within `col_fields`, restricted by `filter_fields`, and
/// aggregated per `value_fields`. This is a workbook-level object (like
/// `Chart`) rather than sheet-scoped like `ExcelTable`, since its source and
/// destination ranges may live on different sheets.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PivotTable {
    /// Workbook-unique identifier, stable across renames.
    pub id: u64,
    /// Display name, unique workbook-wide.
    pub name: String,
    /// Where the records come from.
    pub source: PivotSource,
    /// Sheet the grid is written to, which need not be the source's sheet.
    pub dest_sheet_id: u64,
    /// Top-left row of the output grid, 0-based.
    pub dest_row: usize,
    /// Top-left column of the output grid, 0-based.
    pub dest_col: usize,
    /// Fields grouped down the left edge, outermost first.
    pub row_fields: Vec<PivotField>,
    /// Fields grouped across the top, outermost first.
    pub col_fields: Vec<PivotField>,
    /// Fields aggregated into the body. At least one is required for
    /// [`compute_pivot`] to succeed.
    pub value_fields: Vec<PivotValueField>,
    /// Fields restricting which source records take part.
    pub filter_fields: Vec<PivotFilterField>,
    /// Whether a grand-total row is appended below the body.
    pub grand_totals_row: bool,
    /// Whether a grand-total column is appended to the right of the body.
    pub grand_totals_col: bool,
    /// Bottom-right corner of the last rendered output grid, so a refresh
    /// that produces a smaller grid can clear the now-stale cells.
    #[serde(default)]
    pub last_output_end_row: Option<usize>,
    /// Column half of that corner; see [`PivotTable::last_output_end_row`].
    #[serde(default)]
    pub last_output_end_col: Option<usize>,
}

/// Width, in columns, reserved for row-field labels: one column per row
/// field when there are any. With no row fields at all, Excel only
/// reserves a single placeholder column when there's *exactly one* value
/// field *and* at least one column field for it to sit to the left of --
/// that lone cell holds the value field's own label (e.g. "Max of
/// Amount"), the same way the header's "Row Labels | Sum of X" corner
/// would if there were row fields. With no column fields either (the fully
/// "flat" single-aggregate pivot) or with more than one value field (whose
/// labels already show up elsewhere in the header), there's nothing
/// unambiguous to put in a corner, so Excel reserves no column there at
/// all. All three shapes verified against real Excel via
/// fuzz/fuzz_pivot.py. Shared between `compute_pivot` (which must actually
/// size `PivotBodyRow::row_labels` this way) and `pivot_xlsx.rs` (which
/// needs the same number for `firstDataCol`).
pub(crate) fn row_label_width(pivot: &PivotTable) -> usize {
    if !pivot.row_fields.is_empty() {
        return pivot.row_fields.len();
    }
    if pivot.value_fields.len() == 1 && !pivot.col_fields.is_empty() {
        1
    } else {
        0
    }
}

/// A fully computed pivot result, ready to be materialized into a sheet:
/// `filter_rows` (if any) come first, then a blank spacer row, then
/// `header_rows`, then one entry of `body_rows` per output row -- mirroring
/// Excel's own report-filter placement (verified against real Excel: it
/// always reserves one row per filter field plus a blank spacer above the
/// row/column header grid, and captions each with a "(All)"/"(Multiple
/// Items)" state -- never a specific value's name, since that's specific to
/// the classic single-select page-field mode Excel no longer defaults to).
#[derive(Debug, Clone)]
pub struct PivotGrid {
    /// One `(field name, state)` pair per filter field, in the order they
    /// were added.
    ///
    /// The state is `"(All)"` when every value is allowed, the **item's own
    /// name** when exactly one is selected, and `"(Multiple Items)"`
    /// otherwise -- which is what Excel puts in the page-field cell, and what
    /// `PivotField.CurrentPage` reports alongside it.
    pub filter_rows: Vec<(String, String)>,
    /// The column-header block above the body: one row per column field,
    /// plus a value-field row when there is more than one value field.
    pub header_rows: Vec<Vec<String>>,
    /// The body, one entry per output row, subtotal and grand-total rows
    /// included.
    pub body_rows: Vec<PivotBodyRow>,
    /// Total width in columns (row-label columns + data columns), used by
    /// the caller to know how large a range to clear/allocate. Always >= 2,
    /// so `filter_rows`' two columns (name, state) always fit within it.
    pub width: usize,
    /// The flattened row/column axis groups underlying `body_rows`/the data
    /// columns, exposed (independent of display formatting) so an xlsx
    /// exporter can reconstruct a native `pivotTableDefinition`'s
    /// `rowItems`/`colItems` without re-deriving the grouping itself.
    pub row_axis: Vec<PivotAxisItem>,
    /// Column half of that axis pair; see [`PivotGrid::row_axis`].
    pub col_axis: Vec<PivotAxisItem>,
}

/// One row of a computed pivot's body: its row-field labels and its
/// aggregated values.
#[derive(Debug, Clone)]
pub struct PivotBodyRow {
    /// One entry per row field (or a single "Grand Total" entry when there
    /// are no row fields); blank entries mean "same as the row above".
    pub row_labels: Vec<String>,
    /// Whether this row is the grand total rather than a data or subtotal row.
    pub is_grand_total: bool,
    /// One entry per data column, aligned with the last `header_rows` row.
    pub values: Vec<ResultData>,
}

/// One flattened group along a row or column axis: a label per axis field
/// (`None` past its own depth), plus whether it's a subtotal or grand-total
/// pseudo-group rather than a real leaf group.
#[derive(Debug, Clone)]
pub struct PivotAxisItem {
    /// One entry per field in this axis, `None` past this group's own depth.
    pub labels: Vec<Option<String>>,
    /// Whether this is a subtotal pseudo-group rather than a leaf group.
    pub is_subtotal: bool,
    /// Whether this is the axis's grand-total pseudo-group.
    pub is_grand_total: bool,
}

impl PivotGrid {
    /// Row offset from the pivot's `dest_row` anchor to where the row/col
    /// header + data grid actually begins: 0 with no filter fields, else
    /// one row per filter field plus a blank spacer row.
    pub fn grid_row_offset(&self) -> usize {
        if self.filter_rows.is_empty() {
            0
        } else {
            self.filter_rows.len() + 1
        }
    }

    /// Total height in rows, filter rows and spacer included -- what the
    /// caller needs to allocate or clear at the pivot's `dest_row` anchor.
    pub fn height(&self) -> usize {
        self.grid_row_offset() + self.header_rows.len() + self.body_rows.len()
    }
}

/// A flattened, labeled group of source records along one axis (row or
/// column), produced by recursively grouping by each field in that axis in
/// turn. `record_indices` is the union of every record folded into this
/// group -- for a leaf group that's just its own bucket, for a subtotal or
/// grand-total pseudo-group it's every record under it.
struct FlatGroup {
    /// One label per field in this axis; `None` past the group's own depth
    /// (e.g. a subtotal group has no label for deeper fields).
    labels: Vec<Option<String>>,
    record_indices: Vec<usize>,
    is_subtotal: bool,
    is_grand_total: bool,
}

struct GroupNode {
    label: String,
    record_indices: Vec<usize>,
    children: Vec<GroupNode>,
}

pub(crate) fn group_key(result: &ResultData) -> String {
    match result {
        ResultData::None => "(blank)".to_string(),
        ResultData::String(s) if s.is_empty() => "(blank)".to_string(),
        other => other.to_string(),
    }
}

/// Whether every non-blank value of `records[..][field_idx]` is a genuine
/// number (`Integer`/`Float`), as opposed to text that merely looks
/// numeric (e.g. a zero-padded code like `"08"`, or digits kept as text on
/// purpose). Determines sort order for that field's pivot groups --
/// Excel sorts a real numeric field numerically but a text field
/// alphabetically even when its values happen to look like numbers
/// (verified against real Excel via fuzz/fuzz_pivot.py's `NumStr` column,
/// whose whole purpose is generating quoted numeric-looking text to probe
/// exactly this) -- with one refinement found on Windows: a value that
/// looks like a *negative* number sorts by its digits with the leading
/// `-` stripped, not by the `-` character itself. See
/// `sort_group_entries` and `text_sort_key`. Grouping already collapsed
/// values to strings by this point (`group_key`), which can no longer
/// tell a real `22` from a text `"22"` -- this has to be decided from the
/// original `ResultData`s.
pub(crate) fn field_is_numeric(records: &[Vec<ResultData>], field_idx: usize) -> bool {
    !records.is_empty()
        && records.iter().all(|r| {
            matches!(
                r.get(field_idx),
                Some(ResultData::Integer(_)) | Some(ResultData::Float(_)) | Some(ResultData::None)
            )
        })
}

/// The key `sort_group_entries`'s text-field branch compares siblings by:
/// the value itself, lowercased, *unless* it looks like a negative number
/// (`"-7"`, `"-25"`), in which case the leading `-` is stripped first.
/// Measured on Windows real Excel across three independent sibling sets
/// (fuzz/fuzz_pivot.py's `NumStr` column):
///   `{-7, .0152, 13, 34, 4}`        -> `.0152, 13, 34, 4, -7`
///   `{-46, .097, 01, 02, 1, 10, 35}` -> `.097, 01, 02, 1, 10, 35, -46`
///   `{-25, .0599, .0839, 01, 02, 08, 1, 12, 37}`
///                                   -> `.0599, .0839, 01, 02, 08, 1, 12, -25, 37`
/// A "sorts last" rule (visi's first attempt at this) fits the first two
/// but not the third, where "-25" lands *before* "37" -- comparing "25"
/// (the stripped digits) against the other keys fits all three: "25"
/// falls between "12" and "37" alphabetically, exactly where Excel put
/// "-25". Not tested (no evidence either way): two negative-looking
/// siblings compared against each other -- both get stripped, so they
/// fall back to comparing their digit strings.
fn text_sort_key(s: &str) -> String {
    let trimmed = s.trim();
    let key = match trimmed.strip_prefix('-') {
        Some(rest) if rest.starts_with(|c: char| c.is_ascii_digit()) => rest,
        _ => trimmed,
    };
    key.to_lowercase()
}

fn sort_group_entries(pairs: &mut [(String, Vec<usize>)], numeric: bool) {
    // A blank/empty group always sorts last, regardless of the field's
    // otherwise-numeric-or-text order (verified against real Excel via
    // fuzz/fuzz_pivot.py).
    pairs.sort_by(|a, b| match (a.0 == "(blank)", b.0 == "(blank)") {
        (true, true) => std::cmp::Ordering::Equal,
        (true, false) => std::cmp::Ordering::Greater,
        (false, true) => std::cmp::Ordering::Less,
        (false, false) if numeric => {
            let fa: f64 = a.0.trim().parse().unwrap_or(0.0);
            let fb: f64 = b.0.trim().parse().unwrap_or(0.0);
            fa.partial_cmp(&fb).unwrap_or(std::cmp::Ordering::Equal)
        }
        (false, false) => text_sort_key(&a.0).cmp(&text_sort_key(&b.0)),
    });
}

fn build_group_tree(
    indices: &[usize],
    keys: &[Vec<String>],
    depth: usize,
    num_fields: usize,
    numeric_by_depth: &[bool],
) -> Vec<GroupNode> {
    // Case-insensitive merge (verified against real Excel via
    // fuzz/fuzz_pivot.py, whose generator deliberately mixes casings like
    // "East"/"east" to probe this): Excel's PivotTable field grouping
    // treats text values that differ only in case as the same group,
    // captioned with whichever casing appeared first in the source data --
    // which fewer distinct `groups` entries than `keys` naturally
    // preserves here, since only the first-seen spelling of a key ever
    // becomes `entry.0`.
    let mut groups: Vec<(String, Vec<usize>)> = Vec::new();
    for &idx in indices {
        let key = &keys[idx][depth];
        if let Some(entry) = groups.iter_mut().find(|(k, _)| k.eq_ignore_ascii_case(key)) {
            entry.1.push(idx);
        } else {
            groups.push((key.clone(), vec![idx]));
        }
    }
    sort_group_entries(&mut groups, numeric_by_depth[depth]);
    groups
        .into_iter()
        .map(|(label, idxs)| {
            let children = if depth + 1 < num_fields {
                build_group_tree(&idxs, keys, depth + 1, num_fields, numeric_by_depth)
            } else {
                Vec::new()
            };
            GroupNode {
                label,
                record_indices: idxs,
                children,
            }
        })
        .collect()
}

/// Recursively flattens a group tree into a list of `FlatGroup`s: every leaf
/// group, plus (when enabled for that field) a subtotal pseudo-group after
/// each non-innermost group's children.
fn flatten_groups(
    nodes: &[GroupNode],
    fields: &[PivotField],
    depth: usize,
    num_fields: usize,
    prefix: &[Option<String>],
    out: &mut Vec<FlatGroup>,
) {
    for node in nodes {
        // `labels` holds exactly this node's own depth (depth+1 entries) so
        // that a child's `push` lands at the right position; it's only
        // padded out to `num_fields` at the point a `FlatGroup` is actually
        // emitted (leaf or subtotal), never before recursing further.
        let mut labels = prefix.to_vec();
        labels.push(Some(node.label.clone()));

        if node.children.is_empty() {
            let mut leaf_labels = labels.clone();
            leaf_labels.resize(num_fields, None);
            out.push(FlatGroup {
                labels: leaf_labels,
                record_indices: node.record_indices.clone(),
                is_subtotal: false,
                is_grand_total: false,
            });
        } else {
            flatten_groups(&node.children, fields, depth + 1, num_fields, &labels, out);
            let is_innermost = depth + 1 >= num_fields;
            if fields[depth].subtotal && !is_innermost {
                let mut subtotal_labels = labels.clone();
                subtotal_labels.resize(num_fields, None);
                out.push(FlatGroup {
                    labels: subtotal_labels,
                    record_indices: node.record_indices.clone(),
                    is_subtotal: true,
                    is_grand_total: false,
                });
            }
        }
    }
}

/// Builds the flattened axis groups for `fields` over `record_indices`,
/// optionally appending a grand-total pseudo-group. Returns a single
/// implicit "all records" group when `fields` is empty.
fn build_axis(
    record_indices: &[usize],
    keys: &[Vec<String>],
    fields: &[PivotField],
    grand_total: bool,
    numeric_by_depth: &[bool],
) -> Vec<FlatGroup> {
    if fields.is_empty() {
        return vec![FlatGroup {
            labels: Vec::new(),
            record_indices: record_indices.to_vec(),
            is_subtotal: false,
            is_grand_total: false,
        }];
    }
    let tree = build_group_tree(record_indices, keys, 0, fields.len(), numeric_by_depth);
    let mut flat = Vec::new();
    flatten_groups(&tree, fields, 0, fields.len(), &[], &mut flat);
    // Excel shows the grand total whenever the toggle is on, even when
    // there's only one real group and the grand total would be a literal
    // duplicate of it -- confirmed against real Excel via
    // fuzz/fuzz_pivot.py: a column axis with a single field, filtered down
    // to exactly one distinct value (so there's no possible subtotal
    // either), still got its own redundant "Grand Total" column. Only
    // skip it when there's no data to total at all.
    if grand_total && !flat.is_empty() {
        flat.push(FlatGroup {
            labels: vec![None; fields.len()],
            record_indices: record_indices.to_vec(),
            is_subtotal: false,
            is_grand_total: true,
        });
    }
    flat
}

fn aggregate(sheet: &Sheet, values: &[ResultData], agg: PivotAggregation) -> ResultData {
    // A row/column intersection with zero underlying records (a sparse
    // cell in the cross-tab -- e.g. a row group and column group that
    // simply never co-occur in the source data) renders as a genuinely
    // blank cell in Excel, not a computed zero or #DIV/0! error, for every
    // aggregation kind (verified against real Excel via fuzz/fuzz_pivot.py:
    // even Count and Sum, which have an obvious "zero" answer, still show
    // blank there). This is distinct from records existing but this
    // column's values all being blank for them, which the per-aggregation
    // branches below already handle on their own terms (e.g. Max/Min over
    // an all-blank column already fall back to `ResultData::None`).
    if values.is_empty() {
        return ResultData::None;
    }
    match agg {
        PivotAggregation::Count => ResultData::Integer(
            values
                .iter()
                .filter(|v| !matches!(v, ResultData::None))
                .count() as i64,
        ),
        PivotAggregation::CountNumbers => ResultData::Integer(
            values
                .iter()
                .filter(|v| matches!(v, ResultData::Integer(_) | ResultData::Float(_)))
                .count() as i64,
        ),
        _ => {
            let nums: Vec<f64> = values
                .iter()
                .filter_map(|v| match v {
                    ResultData::Integer(_) | ResultData::Float(_) => sheet.to_f64(v),
                    _ => None,
                })
                .collect();
            match agg {
                PivotAggregation::Sum => {
                    if nums.is_empty() {
                        ResultData::Integer(0)
                    } else {
                        ResultData::Float(Sheet::clean_float(nums.iter().sum()))
                    }
                }
                PivotAggregation::Average => {
                    if nums.is_empty() {
                        ResultData::Error("#DIV/0!".to_string())
                    } else {
                        let avg = nums.iter().sum::<f64>() / nums.len() as f64;
                        ResultData::Float(Sheet::clean_float(avg))
                    }
                }
                PivotAggregation::Max => nums
                    .into_iter()
                    .fold(None, |acc: Option<f64>, x| {
                        Some(acc.map_or(x, |a| a.max(x)))
                    })
                    .map(ResultData::Float)
                    .unwrap_or(ResultData::None),
                PivotAggregation::Min => nums
                    .into_iter()
                    .fold(None, |acc: Option<f64>, x| {
                        Some(acc.map_or(x, |a| a.min(x)))
                    })
                    .map(ResultData::Float)
                    .unwrap_or(ResultData::None),
                PivotAggregation::Count | PivotAggregation::CountNumbers => unreachable!(),
            }
        }
    }
}

/// Resolves a `PivotSource` against the workbook's sheets, returning the
/// owning sheet, the source's column names (in source-column order), the
/// matching absolute sheet-column indices, and the absolute sheet-row
/// indices holding data (i.e. excluding any header/totals row).
/// (owning sheet, source column names, absolute sheet-column indices, absolute data-row indices).
pub(crate) type ResolvedSource<'a> = (&'a Sheet, Vec<String>, Vec<usize>, Vec<usize>);

pub(crate) fn resolve_source<'a>(
    sheets: &'a [&'a Sheet],
    source: &PivotSource,
) -> Result<ResolvedSource<'a>, String> {
    match source {
        PivotSource::Table { name } => {
            let (sheet, table) = sheets
                .iter()
                .find_map(|s| s.find_table(name).map(|t| (*s, t)))
                .ok_or_else(|| format!("Table '{}' not found", name))?;
            let cols: Vec<usize> = (table.start_col..=table.end_col).collect();
            let rows: Vec<usize> = (table.data_start_row()..=table.data_end_row()).collect();
            Ok((sheet, table.columns.clone(), cols, rows))
        }
        PivotSource::Range {
            sheet_id,
            start_row,
            start_col,
            end_row,
            end_col,
        } => {
            let sheet = *sheets
                .iter()
                .find(|s| s.id == *sheet_id)
                .ok_or_else(|| "Pivot source sheet no longer exists".to_string())?;
            if *end_row < *start_row || *end_col < *start_col {
                return Err("Pivot source range end must not precede its start".to_string());
            }
            let cols: Vec<usize> = (*start_col..=*end_col).collect();
            let names: Vec<String> = cols
                .iter()
                .map(|&c| {
                    let v = sheet.get_result_data(&CellRef::new(*start_row, c));
                    let s = v.to_string();
                    if s.is_empty() {
                        crate::core::parser::col_idx_to_letters(c)
                    } else {
                        s
                    }
                })
                .collect();
            let rows: Vec<usize> = if *end_row > *start_row {
                (*start_row + 1..=*end_row).collect()
            } else {
                Vec::new()
            };
            Ok((sheet, names, cols, rows))
        }
    }
}

pub(crate) fn column_index(names: &[String], target: &str) -> Result<usize, String> {
    names
        .iter()
        .position(|c| c.eq_ignore_ascii_case(target))
        .ok_or_else(|| {
            format!(
                "Source column '{}' not found (columns: {})",
                target,
                names.join(", ")
            )
        })
}

/// Computes a pivot table's result grid from the current state of `sheets`.
/// Pure and read-only: callers materialize the returned `PivotGrid` into
/// sheet cells themselves.
/// Computes `pivot` against `sheets`, returning a display-ready grid.
///
/// Pure: it reads source records, applies the filter fields, groups by the
/// row and column fields, aggregates the value fields, and returns the
/// result. Nothing is written -- materializing the grid into cells is
/// `WorkbookManager::refresh_pivot_table`'s job.
///
/// `sheets` must include both the source's sheet and, for a
/// [`PivotSource::Table`] source, whichever sheet carries that table.
///
/// # Errors
///
/// Returns a message if the source cannot be resolved, if a named field is
/// not among the source's columns, or if the pivot has no value fields.
pub fn compute_pivot(sheets: &[&Sheet], pivot: &PivotTable) -> Result<PivotGrid, String> {
    let (sheet, col_names, sheet_cols, data_rows) = resolve_source(sheets, &pivot.source)?;

    for f in pivot.row_fields.iter().chain(pivot.col_fields.iter()) {
        column_index(&col_names, &f.column)?;
    }
    for vf in &pivot.value_fields {
        column_index(&col_names, &vf.column)?;
    }
    for ff in &pivot.filter_fields {
        column_index(&col_names, &ff.column)?;
    }
    if pivot.value_fields.is_empty() {
        return Err("Pivot table has no value fields".to_string());
    }

    // Read every source record unfiltered first -- the filter-row captions
    // below need every distinct value that actually exists in the source,
    // not just the ones that survive filtering, to tell "(All)" apart from
    // "(Multiple Items)".
    let mut all_rows: Vec<Vec<ResultData>> = Vec::with_capacity(data_rows.len());
    for &r in &data_rows {
        let mut row_vals = Vec::with_capacity(sheet_cols.len());
        for &c in &sheet_cols {
            row_vals.push(sheet.get_result_data(&CellRef::new(r, c)));
        }
        all_rows.push(row_vals);
    }

    // A filter field's selectable items are Excel pivot-cache items, which
    // (like row/col group labels) merge case-different text into one item
    // -- so both the "(All)"/"(Multiple Items)" state and the actual
    // row-inclusion test below must compare case-insensitively, not by
    // exact string equality. Verified against real Excel via
    // fuzz/fuzz_pivot.py (iteration 8, seed 599783): a source column with
    // both "East" and "east" rows, filtered to a selection containing
    // "east", must include every row of either casing -- Excel's pivot
    // cache only ever offers one merged "East"/"east" checkbox, not two.
    let mut filter_rows: Vec<(String, String)> = Vec::new();
    for ff in &pivot.filter_fields {
        let idx = column_index(&col_names, &ff.column)?;
        let distinct: std::collections::HashSet<String> = all_rows
            .iter()
            .map(|row| group_key(&row[idx]).to_ascii_lowercase())
            .collect();
        let state = match &ff.selected_values {
            None => "(All)".to_string(),
            Some(selected) => {
                let selected_set: std::collections::HashSet<String> =
                    selected.iter().map(|v| v.to_ascii_lowercase()).collect();
                let is_all = selected_set.len() == distinct.len()
                    && distinct.iter().all(|v| selected_set.contains(v));
                if is_all {
                    "(All)".to_string()
                } else if !ff.multiple_selection && selected_set.len() == 1 {
                    // Single-select mode names the item; multi-select says
                    // `(Multiple Items)` even for one. Both measured -- see
                    // `PivotFilterField::multiple_selection`. The item's own
                    // casing is used, since the cache merges case variants
                    // onto whichever it saw first.
                    let wanted = &selected_set;
                    all_rows
                        .iter()
                        .map(|row| group_key(&row[idx]))
                        .find(|v| wanted.contains(&v.to_ascii_lowercase()))
                        .unwrap_or_else(|| "(Multiple Items)".to_string())
                } else {
                    "(Multiple Items)".to_string()
                }
            }
        };
        filter_rows.push((ff.column.clone(), state));
    }

    // Apply filter fields to build the working record set.
    let mut records: Vec<Vec<ResultData>> = Vec::new();
    'row: for row_vals in &all_rows {
        for ff in &pivot.filter_fields {
            if let Some(selected) = &ff.selected_values {
                let idx = column_index(&col_names, &ff.column)?;
                let key = group_key(&row_vals[idx]);
                if !selected.iter().any(|v| v.eq_ignore_ascii_case(&key)) {
                    continue 'row;
                }
            }
        }
        records.push(row_vals.clone());
    }

    let record_indices: Vec<usize> = (0..records.len()).collect();

    let row_field_idxs: Vec<usize> = pivot
        .row_fields
        .iter()
        .map(|f| column_index(&col_names, &f.column))
        .collect::<Result<_, _>>()?;
    let col_field_idxs: Vec<usize> = pivot
        .col_fields
        .iter()
        .map(|f| column_index(&col_names, &f.column))
        .collect::<Result<_, _>>()?;
    // The casing a case-insensitively-merged group displays under must be
    // decided once per field, from that field's first occurrence anywhere
    // in the source data -- not independently within whichever nested
    // branch of the *other* axis it happens to first appear under.
    // `build_group_tree`'s merge only sees one branch's records at a time,
    // so canonicalizing case up front here (before grouping) is what makes
    // every branch agree on the same casing for the same value (verified
    // against real Excel via fuzz/fuzz_pivot.py: its pivot cache assigns
    // one canonical spelling per distinct value field-wide).
    let mut case_canon: HashMap<usize, HashMap<String, String>> = HashMap::new();
    let mut canonical_key = |field_idx: usize, raw: String| -> String {
        let map = case_canon.entry(field_idx).or_default();
        map.entry(raw.to_ascii_lowercase()).or_insert(raw).clone()
    };
    // Seed the canonical casing from *every* source row, not just the ones
    // that survive `pivot.filter_fields` -- Excel's pivot cache assigns a
    // value's canonical casing once, field-wide, from the raw source data,
    // and a filter only hides cached items afterward rather than rebuilding
    // the cache from the filtered subset. Skipping this seeding step used
    // to let a filter change which occurrence of a case-variant value
    // counted as "first" (whichever one happened to survive the filter),
    // even though Excel's own choice never depends on the filter at all.
    for row_vals in &all_rows {
        for &i in row_field_idxs.iter().chain(col_field_idxs.iter()) {
            canonical_key(i, group_key(&row_vals[i]));
        }
    }
    let row_keys: Vec<Vec<String>> = if pivot.row_fields.is_empty() {
        Vec::new()
    } else {
        records
            .iter()
            .map(|rec| {
                row_field_idxs
                    .iter()
                    .map(|&i| canonical_key(i, group_key(&rec[i])))
                    .collect()
            })
            .collect()
    };
    let col_keys: Vec<Vec<String>> = if pivot.col_fields.is_empty() {
        Vec::new()
    } else {
        records
            .iter()
            .map(|rec| {
                col_field_idxs
                    .iter()
                    .map(|&i| canonical_key(i, group_key(&rec[i])))
                    .collect()
            })
            .collect()
    };
    let row_numeric: Vec<bool> = row_field_idxs
        .iter()
        .map(|&i| field_is_numeric(&records, i))
        .collect();
    let col_numeric: Vec<bool> = col_field_idxs
        .iter()
        .map(|&i| field_is_numeric(&records, i))
        .collect();

    let row_groups = build_axis(
        &record_indices,
        &row_keys,
        &pivot.row_fields,
        pivot.grand_totals_row,
        &row_numeric,
    );
    let col_groups = build_axis(
        &record_indices,
        &col_keys,
        &pivot.col_fields,
        pivot.grand_totals_col,
        &col_numeric,
    );

    let value_multiplier = if pivot.value_fields.len() > 1 {
        pivot.value_fields.len()
    } else {
        1
    };
    let value_idxs: Vec<usize> = pivot
        .value_fields
        .iter()
        .map(|vf| column_index(&col_names, &vf.column))
        .collect::<Result<_, _>>()?;
    let value_labels = value_field_labels(&pivot.value_fields);

    // --- Header rows ---
    // Matches Excel's default "compact form" display, verified against real
    // Excel via fuzz/fuzz_pivot.py (see fuzz/README.md's pivot section):
    // the outermost row field's caption becomes the literal text "Row
    // Labels" (deeper row fields keep their real name), and -- whenever
    // there's at least one column field -- an extra header row captioned
    // "Column Labels" is inserted above the column-field-value rows. Excel
    // can't be made to use its alternate "tabular form" (the per-field
    // LayoutForm VBA property that would show real field names instead is
    // confirmed to have no effect on Mac Excel, and the table-wide
    // RowAxisLayout/ColumnAxisLayout methods that do work hang Mac Excel
    // outright when driven via VBA/AppleScript), so matching this on visi's
    // side is the only tractable way to reach parity.
    let n_col_header_rows = pivot.col_fields.len().max(1);
    // The extra value-label row (needed to tell a column group's own value
    // apart from which value field a sub-column holds) only makes sense
    // when there's a column-group-values row for it to sit below in the
    // first place. With no column fields at all, there's no such row --
    // Excel just lists every value field as a plain adjacent column in the
    // single header row instead, exactly like a flat table's header
    // (verified against real Excel via fuzz/fuzz_pivot.py: 2 value fields
    // with no column fields produced one header row with both labels side
    // by side, not two stacked rows).
    let n_header_rows = if value_multiplier > 1 && !pivot.col_fields.is_empty() {
        n_col_header_rows + 1
    } else {
        n_col_header_rows
    };
    let row_label_width = row_label_width(pivot);

    let mut header_rows: Vec<Vec<String>> = Vec::new();
    for r in 0..n_header_rows {
        let mut row: Vec<String> = Vec::new();
        for i in 0..row_label_width {
            // Row-label captions ("Row Labels" plus any deeper row fields'
            // real names) sit on the *last* header row -- the one right
            // above the data -- not the first: with multiple value fields
            // that's the extra value-label row, not the column-field-value
            // row above it (confirmed against real Excel: with 2 value
            // fields, "Row Labels" lands on the value-label row while the
            // column-value row directly above it leaves that same spot
            // blank).
            if r == n_header_rows - 1 {
                row.push(if i == 0 && !pivot.row_fields.is_empty() {
                    "Row Labels".to_string()
                } else {
                    pivot
                        .row_fields
                        .get(i)
                        .map(|f| f.column.clone())
                        .unwrap_or_default()
                });
            } else {
                row.push(String::new());
            }
        }
        // Excel merges a repeated label across the columns it spans -- a
        // value field fanning a single column group out into several
        // adjacent sub-columns is one way that happens, a shallower column
        // field repeating over several deeper-field sub-columns under the
        // *same* ancestor chain is another -- showing the label once at the
        // leftmost column and blank for the rest. The two cases need
        // different adjacency tests: within one group, every `vf` beyond
        // the first is *always* a repeat (they all render that group's same
        // `labels[r]`, `vf` doesn't affect it). Across groups, `labels[r]`
        // matching alone isn't enough -- two unrelated groups can
        // coincidentally share a leaf value at depth `r` (e.g. two
        // different outer-field branches both happening to have a "west"
        // child) without being siblings under the same parent, so merging
        // them would silently drop one's real value. Only merge when every
        // depth from 0 up to and including `r` matches the immediately
        // preceding group, which is exactly the condition for them being
        // adjacent leaves of the same parent in `col_groups`'s tree order.
        let mut prev_group: Option<&FlatGroup> = None;
        for group in &col_groups {
            // A subtotal group's labels hold exactly one real value, at
            // whichever depth it was inserted -- e.g. `[Some("-3"), None]`
            // for an outer-field subtotal over a 2-level axis. That's the
            // one row its caption becomes "<value> Total" (or, with 2+
            // value fields, "<value> <value field label>" per sub-column,
            // mirroring the grand-total column's "Total <value label>"
            // treatment below -- confirmed against real Excel via
            // fuzz/fuzz_pivot.py: with 2 value fields it repeats the value
            // field's own name under a subtotal group instead of the
            // literal word "Total", and doesn't emit a separate
            // value-label row beneath it the way non-subtotal groups do);
            // every other column-field row either inherits an ancestor's
            // label (already handled below) or stays blank.
            let subtotal_depth = group
                .is_subtotal
                .then(|| group.labels.iter().rposition(|l| l.is_some()))
                .flatten();
            for vf in 0..value_multiplier {
                let label = if r < pivot.col_fields.len() {
                    if group.is_grand_total {
                        // The grand-total column's caption always lands on
                        // the *outermost* column-field row (r == 0), not
                        // the deepest one -- confirmed against real Excel
                        // with a 2-level column axis, where "Grand Total"
                        // showed up on the shallow row while the deep row
                        // beneath it stayed blank (the two coincide, and so
                        // looked identical, in every single-column-field
                        // case tested before that).
                        if r == 0 {
                            if value_multiplier > 1 {
                                format!("Total {}", value_labels[vf])
                            } else {
                                "Grand Total".to_string()
                            }
                        } else {
                            String::new()
                        }
                    } else if subtotal_depth == Some(r) {
                        let value = group.labels[r].clone().unwrap();
                        if value_multiplier > 1 {
                            format!("{} {}", value, value_labels[vf])
                        } else {
                            format!("{} Total", value)
                        }
                    } else {
                        let is_repeat = if vf > 0 {
                            true
                        } else {
                            prev_group.is_some_and(|pg| {
                                (0..=r).all(|d| pg.labels.get(d) == group.labels.get(d))
                            })
                        };
                        if is_repeat {
                            String::new()
                        } else {
                            group
                                .labels
                                .get(r)
                                .and_then(|l| l.clone())
                                .unwrap_or_default()
                        }
                    }
                } else if group.is_grand_total || group.is_subtotal {
                    // Already captioned "Total <value label>" (grand total)
                    // or "<value> <value label>" (subtotal) on the
                    // column-field row above -- no separate value-label row
                    // for these groups.
                    String::new()
                } else {
                    value_labels.get(vf).cloned().unwrap_or_default()
                };
                row.push(label);
            }
            prev_group = Some(group);
        }
        header_rows.push(row);
    }
    // If there's exactly one column group with no column fields, put the
    // single value field's label directly in the header row (mirrors the
    // classic single-value-field pivot layout: "Row Labels | Sum of X").
    if pivot.col_fields.is_empty()
        && value_multiplier == 1
        && let Some(last) = header_rows.last_mut()
        && let Some(cell) = last.last_mut()
        && let Some(label) = value_labels.first()
    {
        *cell = label.clone();
    }
    // Whenever there's at least one column field, Excel prepends a header
    // row captioned "Column Labels" above the column-field-value rows.
    // Its row-label area is blank, except: when there's exactly one value
    // field *and* at least one row field, that field's label goes in the
    // very first cell (mirrors the single-value-field layout's "Row Labels
    // | Sum of X" convention, just one row up since the row-label area's
    // own first cell is taken by the "Row Labels" caption instead). With no
    // row fields, that label has nowhere to go here -- the row-label area
    // has no field caption to displace -- so it surfaces on the sole body
    // row's corner instead (see the "Total" fallback below).
    if !pivot.col_fields.is_empty() {
        let mut row = vec![String::new(); row_label_width];
        if value_multiplier == 1
            && !pivot.row_fields.is_empty()
            && let Some(label) = value_labels.first()
        {
            row[0] = label.clone();
        }
        row.push("Column Labels".to_string());
        row.resize(
            row_label_width + col_groups.len() * value_multiplier,
            String::new(),
        );
        header_rows.insert(0, row);
    }

    // --- Body rows ---
    let mut body_rows: Vec<PivotBodyRow> = Vec::new();
    let mut prev_labels: Vec<Option<String>> = vec![None; row_label_width];
    for rg in &row_groups {
        let mut display_labels = vec![String::new(); row_label_width];
        if rg.is_grand_total {
            display_labels[0] = "Grand Total".to_string();
            for l in prev_labels.iter_mut() {
                *l = None;
            }
        } else {
            let mut changed = false;
            for d in 0..row_label_width {
                let cur = if pivot.row_fields.is_empty() {
                    None
                } else {
                    rg.labels.get(d).cloned().flatten()
                };
                let is_subtotal_marker =
                    rg.is_subtotal && rg.labels.get(d).map(|l| l.is_some()).unwrap_or(false);
                let show = changed || cur != prev_labels[d] || is_subtotal_marker;
                if show {
                    if let Some(ref v) = cur {
                        display_labels[d] = if is_subtotal_marker {
                            format!("{} Total", v)
                        } else {
                            v.clone()
                        };
                    }
                    changed = true;
                }
                prev_labels[d] = cur;
            }
            // When there are no row fields *and* no column fields either,
            // `row_label_width` is 0 (see `row_label_width`'s doc comment)
            // -- there's no label cell here at all, just the value itself.
            if pivot.row_fields.is_empty() && row_label_width > 0 {
                // With no row fields there's exactly one body row (the
                // aggregate over everything), and no "Row Labels"-captioned
                // header row above it to hold a single value field's label
                // the way the col_fields-empty layout does in the header
                // (see the header construction above) -- so it surfaces
                // here instead, on the one row that exists. Falls back to
                // "Total" when there's more than one value field, same as
                // the header's equivalent case.
                display_labels[0] = if !pivot.col_fields.is_empty() && value_multiplier == 1 {
                    value_labels.first().cloned().unwrap_or_default()
                } else {
                    "Total".to_string()
                };
            }
        }

        let row_record_set: std::collections::HashSet<usize> =
            rg.record_indices.iter().copied().collect();
        let mut values: Vec<ResultData> = Vec::new();
        for cg in &col_groups {
            for (vf_pos, &vidx) in value_idxs.iter().enumerate() {
                if vf_pos > 0 && value_multiplier == 1 {
                    break;
                }
                let col_vals: Vec<ResultData> = cg
                    .record_indices
                    .iter()
                    .filter(|i| row_record_set.contains(i))
                    .map(|&i| records[i][vidx].clone())
                    .collect();
                values.push(aggregate(
                    sheet,
                    &col_vals,
                    pivot.value_fields[vf_pos].aggregation,
                ));
            }
        }

        body_rows.push(PivotBodyRow {
            row_labels: display_labels,
            is_grand_total: rg.is_grand_total,
            values,
        });
    }

    let width = row_label_width + col_groups.len() * value_multiplier;
    let to_axis_items = |groups: &[FlatGroup]| -> Vec<PivotAxisItem> {
        groups
            .iter()
            .map(|g| PivotAxisItem {
                labels: g.labels.clone(),
                is_subtotal: g.is_subtotal,
                is_grand_total: g.is_grand_total,
            })
            .collect()
    };
    Ok(PivotGrid {
        filter_rows,
        header_rows,
        body_rows,
        width,
        row_axis: to_axis_items(&row_groups),
        col_axis: to_axis_items(&col_groups),
    })
}

/// Finds the unique row/col-axis group matching `criteria` -- `(field
/// depth, item text)` pairs restricted to one axis -- for `GETPIVOTDATA`.
/// Empty `criteria` means "the axis's grand total". A non-empty `criteria`
/// that doesn't specify every field on the axis matches the subtotal group
/// at that depth (mirrors Excel: naming only the outer field(s) of a nested
/// row/col axis returns that branch's subtotal, not an arbitrary leaf under
/// it); naming every field down to the innermost one matches the leaf.
/// Ambiguous or absent matches are both reported as `#REF!`, matching real
/// Excel's error for a `GETPIVOTDATA` criteria pair that doesn't resolve.
fn match_pivot_axis(
    axis: &[PivotAxisItem],
    criteria: &[(usize, &str)],
    field_count: usize,
) -> Result<usize, String> {
    if criteria.is_empty() {
        return axis
            .iter()
            .position(|g| g.is_grand_total)
            .or(if field_count == 0 && axis.len() == 1 {
                Some(0)
            } else {
                None
            })
            .ok_or_else(|| "#REF!".to_string());
    }
    let max_depth = criteria.iter().map(|(d, _)| *d).max().unwrap_or(0);
    let want_leaf = max_depth + 1 == field_count;
    let matches: Vec<usize> = axis
        .iter()
        .enumerate()
        .filter(|(_, group)| {
            if group.is_grand_total {
                return false;
            }
            if want_leaf {
                if group.is_subtotal {
                    return false;
                }
            } else {
                let own_depth = group.labels.iter().rposition(|l| l.is_some());
                if !(group.is_subtotal && own_depth == Some(max_depth)) {
                    return false;
                }
            }
            criteria.iter().all(|(depth, item)| {
                group
                    .labels
                    .get(*depth)
                    .and_then(|l| l.as_deref())
                    .map(|l| l.eq_ignore_ascii_case(item))
                    .unwrap_or(false)
            })
        })
        .map(|(i, _)| i)
        .collect();
    match matches.len() {
        1 => Ok(matches[0]),
        _ => Err("#REF!".to_string()),
    }
}

/// Implements `GETPIVOTDATA`: extracts a single summarized value out of a
/// pivot table's computed grid by data-field name plus `(row/col field,
/// item)` criteria pairs, the same way real Excel's formula does when
/// pointed at a rendered pivot. Recomputes the grid fresh from `sheets`
/// rather than caching it, consistent with formulas re-evaluating from
/// current sheet state on every recalculation pass.
pub fn getpivotdata(
    sheets: &[&Sheet],
    pivot: &PivotTable,
    data_field: &str,
    criteria: &[(String, String)],
) -> Result<ResultData, String> {
    let grid = compute_pivot(sheets, pivot)?;

    let value_labels = value_field_labels(&pivot.value_fields);
    let value_multiplier = if pivot.value_fields.len() > 1 {
        pivot.value_fields.len()
    } else {
        1
    };
    let value_field_idx = pivot
        .value_fields
        .iter()
        .position(|vf| vf.column.eq_ignore_ascii_case(data_field))
        .or_else(|| {
            value_labels
                .iter()
                .position(|l| l.eq_ignore_ascii_case(data_field))
        })
        .ok_or_else(|| "#VALUE!".to_string())?;

    let mut row_criteria: Vec<(usize, &str)> = Vec::new();
    let mut col_criteria: Vec<(usize, &str)> = Vec::new();
    for (field, item) in criteria {
        if let Some(depth) = pivot
            .row_fields
            .iter()
            .position(|f| f.column.eq_ignore_ascii_case(field))
        {
            row_criteria.push((depth, item.as_str()));
        } else if let Some(depth) = pivot
            .col_fields
            .iter()
            .position(|f| f.column.eq_ignore_ascii_case(field))
        {
            col_criteria.push((depth, item.as_str()));
        } else {
            return Err("#REF!".to_string());
        }
    }

    let row_idx = match_pivot_axis(&grid.row_axis, &row_criteria, pivot.row_fields.len())?;
    let col_idx = match_pivot_axis(&grid.col_axis, &col_criteria, pivot.col_fields.len())?;

    let pos = col_idx * value_multiplier + value_field_idx;
    grid.body_rows
        .get(row_idx)
        .and_then(|r| r.values.get(pos))
        .cloned()
        .ok_or_else(|| "#REF!".to_string())
}

/// Returns the distinct values of `values`, sorted the same way pivot
/// groups are (ascending numeric if every value parses as a number,
/// otherwise case-insensitive ascending text) -- used by the xlsx exporter
/// to build a pivot field's flat `<items>` enumeration.
pub(crate) fn sorted_distinct_strings(values: &[String], numeric: bool) -> Vec<String> {
    let mut pairs: Vec<(String, Vec<usize>)> = distinct_strings(values)
        .into_iter()
        .map(|s| (s, Vec::new()))
        .collect();
    sort_group_entries(&mut pairs, numeric);
    pairs.into_iter().map(|(s, _)| s).collect()
}

/// The distinct values in **first-seen** order, which is the order a pivot
/// cache stores them in.
///
/// Measured: Excel's `<sharedItems>` are in source order while a pivot
/// field's `<items>` are sorted for display and reference sharedItems by
/// index, so the two orders are both needed and are different. See
/// `fuzz/pivot_filter_probe.py`.
///
/// Case-insensitive dedup (first-seen casing kept), matching
/// `build_group_tree`'s merge -- this feeds the exported pivot cache, so it
/// must agree with how `compute_pivot` actually groups these same values or a
/// reimported or refreshed pivot's item list falls out of sync with its own
/// displayed grouping.
pub(crate) fn distinct_strings(values: &[String]) -> Vec<String> {
    let mut seen: Vec<String> = Vec::new();
    for v in values {
        if !seen.iter().any(|s| s.eq_ignore_ascii_case(v)) {
            seen.push(v.clone());
        }
    }
    seen
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::engine::SheetInit;

    fn source_sheet() -> Sheet {
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 9,
            cols: 4,
            ..Default::default()
        });
        let header = ["Region", "Product", "Rep", "Amount"];
        for (c, h) in header.iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows: [[&str; 4]; 8] = [
            ["East", "Widget", "Alice", "10"],
            ["East", "Widget", "Bob", "20"],
            ["East", "Gadget", "Alice", "5"],
            ["West", "Widget", "Carol", "30"],
            ["West", "Gadget", "Carol", "40"],
            ["West", "Gadget", "Dave", "50"],
            ["East", "Gadget", "Bob", "15"],
            ["West", "Widget", "Dave", "25"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 8, 3, true, false)
            .unwrap();
        sheet
    }

    fn base_pivot() -> PivotTable {
        PivotTable {
            id: 1,
            name: "Pivot1".to_string(),
            source: PivotSource::Table {
                name: "Sales".to_string(),
            },
            dest_sheet_id: 0,
            dest_row: 0,
            dest_col: 0,
            row_fields: vec![PivotField::new("Region")],
            col_fields: vec![],
            value_fields: vec![PivotValueField::new("Amount", PivotAggregation::Sum)],
            filter_fields: vec![],
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        }
    }

    fn value_at(row: &PivotBodyRow, col: usize) -> f64 {
        match &row.values[col] {
            ResultData::Float(f) => *f,
            ResultData::Integer(i) => *i as f64,
            other => panic!("expected numeric, got {:?}", other),
        }
    }

    #[test]
    fn test_single_row_field_sum_with_grand_total() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // East: 10+20+5+15=50, West: 30+40+50+25=145, Grand Total: 195
        assert_eq!(grid.body_rows.len(), 3);
        assert_eq!(grid.body_rows[0].row_labels[0], "East");
        assert_eq!(value_at(&grid.body_rows[0], 0), 50.0);
        assert_eq!(grid.body_rows[1].row_labels[0], "West");
        assert_eq!(value_at(&grid.body_rows[1], 0), 145.0);
        assert!(grid.body_rows[2].is_grand_total);
        assert_eq!(grid.body_rows[2].row_labels[0], "Grand Total");
        assert_eq!(value_at(&grid.body_rows[2], 0), 195.0);
    }

    #[test]
    fn test_getpivotdata_matches_a_row_group() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let result = getpivotdata(
            &[&sheet],
            &pivot,
            "Amount",
            &[("Region".to_string(), "East".to_string())],
        )
        .unwrap();
        assert!(matches!(result, ResultData::Float(f) if f == 50.0));
    }

    #[test]
    fn test_getpivotdata_empty_criteria_matches_grand_total() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let result = getpivotdata(&[&sheet], &pivot, "Amount", &[]).unwrap();
        assert!(matches!(result, ResultData::Float(f) if f == 195.0));
    }

    #[test]
    fn test_getpivotdata_partial_criteria_matches_subtotal() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        // East: Widget=10+20=30, Gadget=5+15=20 -> Region subtotal 50
        let result = getpivotdata(
            &[&sheet],
            &pivot,
            "Amount",
            &[("Region".to_string(), "East".to_string())],
        )
        .unwrap();
        assert!(matches!(result, ResultData::Float(f) if f == 50.0));
    }

    #[test]
    fn test_getpivotdata_full_path_matches_leaf() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        let result = getpivotdata(
            &[&sheet],
            &pivot,
            "Amount",
            &[
                ("Region".to_string(), "East".to_string()),
                ("Product".to_string(), "Widget".to_string()),
            ],
        )
        .unwrap();
        assert!(matches!(result, ResultData::Float(f) if f == 30.0));
    }

    #[test]
    fn test_getpivotdata_unknown_field_is_ref_error() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let err = getpivotdata(
            &[&sheet],
            &pivot,
            "Amount",
            &[("NotAField".to_string(), "East".to_string())],
        )
        .unwrap_err();
        assert_eq!(err, "#REF!");
    }

    #[test]
    fn test_getpivotdata_unknown_item_is_ref_error() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let err = getpivotdata(
            &[&sheet],
            &pivot,
            "Amount",
            &[("Region".to_string(), "North".to_string())],
        )
        .unwrap_err();
        assert_eq!(err, "#REF!");
    }

    #[test]
    fn test_getpivotdata_unknown_data_field_is_value_error() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let err = getpivotdata(
            &[&sheet],
            &pivot,
            "NotAField",
            &[("Region".to_string(), "East".to_string())],
        )
        .unwrap_err();
        assert_eq!(err, "#VALUE!");
    }

    #[test]
    fn test_row_and_col_fields_with_subtotals() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        pivot.col_fields = vec![PivotField::new("Rep")];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // Region subtotal rows should appear (2 regions x (2 products + 1 subtotal)) + grand total
        let subtotal_rows: Vec<&PivotBodyRow> = grid
            .body_rows
            .iter()
            .filter(|r| r.row_labels[0].ends_with("Total") && !r.is_grand_total)
            .collect();
        assert_eq!(subtotal_rows.len(), 2); // one per region
        assert!(grid.body_rows.last().unwrap().is_grand_total);
    }

    #[test]
    fn test_nested_row_field_second_level_labels_are_not_lost() {
        // The second (innermost) row field's own labels must survive being
        // nested under the first field's groups, not be truncated away when
        // the group tree is flattened.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        let leaf_rows: Vec<&PivotBodyRow> = grid
            .body_rows
            .iter()
            .filter(|r| !r.row_labels[0].ends_with("Total") && !r.is_grand_total)
            .collect();
        // East has Widget+Gadget, West has Widget+Gadget: 4 leaf rows.
        assert_eq!(leaf_rows.len(), 4);
        // Every leaf row must show a real (non-blank) Product label, not "".
        for row in &leaf_rows {
            assert!(
                !row.row_labels[1].is_empty(),
                "expected a Product label on leaf row {:?}, got blank",
                row.row_labels
            );
        }
        let products: Vec<&str> = leaf_rows.iter().map(|r| r.row_labels[1].as_str()).collect();
        assert!(products.contains(&"Widget"));
        assert!(products.contains(&"Gadget"));
    }

    #[test]
    fn test_count_aggregation() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.value_fields = vec![PivotValueField::new("Rep", PivotAggregation::Count)];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.body_rows.len(), 2);
        // East has 4 records, West has 4 records
        for row in &grid.body_rows {
            assert_eq!(value_at(row, 0), 4.0);
        }
    }

    #[test]
    fn test_filter_field_restricts_records() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.filter_fields = vec![PivotFilterField {
            column: "Product".to_string(),
            selected_values: Some(vec!["Widget".to_string()]),
            multiple_selection: true,
        }];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        // East widgets: 10+20=30, West widgets: 30+25=55
        assert_eq!(grid.body_rows.len(), 2);
        assert_eq!(value_at(&grid.body_rows[0], 0), 30.0);
        assert_eq!(value_at(&grid.body_rows[1], 0), 55.0);
    }

    #[test]
    fn test_filter_field_selection_matches_case_insensitively() {
        // A filter field's selectable items are Excel pivot-cache items,
        // which merge case-different text into a single item exactly like
        // row/col group labels do (see
        // test_case_variant_values_merge_using_globally_first_seen_casing)
        // -- so selecting "east" must match *every* row spelled "East" or
        // "east", not just rows with that exact casing.
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 4,
            cols: 2,
            ..Default::default()
        });
        for (c, h) in ["Mixed", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows: [[&str; 2]; 3] = [["East", "10"], ["east", "20"], ["West", "30"]];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 1, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.source = PivotSource::Table {
            name: "Sales".to_string(),
        };
        pivot.row_fields = vec![];
        pivot.value_fields = vec![PivotValueField::new("Amount", PivotAggregation::Sum)];
        pivot.filter_fields = vec![PivotFilterField {
            column: "Mixed".to_string(),
            selected_values: Some(vec!["east".to_string()]),
            multiple_selection: true,
        }];
        pivot.grand_totals_row = false;
        pivot.grand_totals_col = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        // Both "East" (10) and "east" (20) rows must be included: 30, not 20.
        assert_eq!(value_at(&grid.body_rows[0], 0), 30.0);
    }

    #[test]
    fn test_no_filter_fields_means_no_reserved_rows() {
        let sheet = source_sheet();
        let pivot = base_pivot();
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert!(grid.filter_rows.is_empty());
        assert_eq!(grid.grid_row_offset(), 0);
        assert_eq!(grid.height(), grid.header_rows.len() + grid.body_rows.len());
    }

    #[test]
    fn test_filter_field_state_label_all_vs_multiple_items() {
        // Product has exactly two distinct values in `source_sheet`: Widget, Gadget.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.filter_fields = vec![PivotFilterField {
            column: "Product".to_string(),
            selected_values: None,
            multiple_selection: true,
        }];

        // No selection at all -> "(All)".
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(
            grid.filter_rows,
            vec![("Product".to_string(), "(All)".to_string())]
        );
        assert_eq!(grid.grid_row_offset(), 2); // 1 filter row + 1 blank spacer

        // Explicitly selecting every existing distinct value is equivalent to "(All)".
        pivot.filter_fields[0].selected_values =
            Some(vec!["Widget".to_string(), "Gadget".to_string()]);
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.filter_rows[0].1, "(All)");

        // A strict subset -> "(Multiple Items)". Verified against real
        // Excel: even a single selected value out of several shows this,
        // never the value's own name -- that's specific to the classic
        // single-select page-field mode Excel no longer defaults to.
        pivot.filter_fields[0].selected_values = Some(vec!["Widget".to_string()]);
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.filter_rows[0].1, "(Multiple Items)");

        // ...and that single-select mode is exactly where the item's own
        // name does show, which is what `PivotField.CurrentPage = "Widget"`
        // produces. Measured: the page-field cell reads `Widget`.
        pivot.filter_fields[0].multiple_selection = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.filter_rows[0].1, "Widget");
    }

    #[test]
    fn test_col_axis_subtotal_group_gets_total_caption_and_grand_total_stays_outermost() {
        // With a 2-level column axis (both fields' subtotals enabled by
        // default), the header logic gives a column-axis subtotal group its
        // own "<value> Total" caption. The grand-total column's caption is
        // placed on the *outermost* row.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Rep")];
        pivot.col_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // header_rows[0] is the prepended "Column Labels" row; [1] is the
        // outermost column field (Region), [2] is the deepest (Product).
        let region_row = &grid.header_rows[1];
        assert!(region_row.contains(&"East Total".to_string()));
        assert!(region_row.contains(&"West Total".to_string()));
        assert!(region_row.contains(&"Grand Total".to_string()));
        let product_row = &grid.header_rows[2];
        assert_eq!(product_row.last().unwrap(), "");
    }

    #[test]
    fn test_col_axis_subtotal_caption_uses_value_field_label_with_multiple_value_fields() {
        // With 2+ value fields, a col-field subtotal group repeats the value
        // field's own name directly on the subtotal's caption row ("<n> Min of Amount",
        // "<n> Sum of Amount") and emits no separate label row underneath
        // for those sub-columns.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![];
        pivot.col_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        pivot.value_fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Min),
            PivotValueField::new("Amount", PivotAggregation::Sum),
        ];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // header_rows[0] is "Column Labels", [1] is Region (outer, with the
        // subtotal), [2] is Product (deepest), [3] is the value-label row.
        let region_row = &grid.header_rows[1];
        // Min and Sum are different aggregations, so their default
        // captions are distinct on their own and Excel leaves the reused
        // "Amount" source column unsuffixed (see
        // test_value_field_labels_leaves_distinct_aggregations_on_same_column_unsuffixed).
        assert!(region_row.contains(&"East Min of Amount".to_string()));
        assert!(region_row.contains(&"East Sum of Amount".to_string()));
        assert!(region_row.contains(&"West Min of Amount".to_string()));
        assert!(region_row.contains(&"West Sum of Amount".to_string()));
        assert!(
            !region_row
                .iter()
                .any(|c| c == "East Total" || c == "West Total")
        );

        // The value-label row must stay blank under the subtotal's
        // sub-columns (no redundant second label row for them), while still
        // showing the value labels under the non-subtotal leaf columns.
        let value_label_row = grid.header_rows.last().unwrap();
        assert!(value_label_row.contains(&"Min of Amount".to_string()));
        assert!(value_label_row.contains(&"Sum of Amount".to_string()));
        let east_subtotal_idx = region_row
            .iter()
            .position(|c| c == "East Min of Amount")
            .unwrap();
        assert_eq!(value_label_row[east_subtotal_idx], "");
        assert_eq!(value_label_row[east_subtotal_idx + 1], "");
    }

    #[test]
    fn test_col_axis_repeated_leaf_value_under_different_parents_is_not_falsely_merged() {
        // Repeated leaf values under different parent groups are preserved
        // rather than merged across unrelated outer-field branches.
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 3,
            cols: 3,
            ..Default::default()
        });
        for (c, h) in ["Group", "Sub", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        // GroupA's only Sub child and GroupB's only Sub child are both "X",
        // with nothing else between them once flattened.
        let rows: [[&str; 3]; 2] = [["GroupA", "X", "1"], ["GroupB", "X", "2"]];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 2, 2, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.row_fields = vec![];
        pivot.col_fields = vec![PivotField::new("Group"), PivotField::new("Sub")];
        pivot.grand_totals_col = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // Deepest (Sub) row: "X" must appear for *both* groups, not just
        // the first (with the second silently blanked as a false "repeat").
        let sub_row = &grid.header_rows[2];
        let x_count = sub_row.iter().filter(|c| *c == "X").count();
        assert_eq!(
            x_count, 2,
            "expected \"X\" under both GroupA and GroupB, got {sub_row:?}"
        );
    }

    #[test]
    fn test_multiple_value_fields_become_column_labels() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.value_fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Count),
        ];
        pivot.grand_totals_row = false;
        pivot.grand_totals_col = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.header_rows.last().unwrap()[1], "Sum of Amount");
        // The first value field on "Amount" uses Sum, which clones the
        // column for every value field after it (see `value_field_labels`'s
        // doc comment) -- so the second value field's default label
        // disambiguates as "Amount2", matching real Excel.
        assert_eq!(grid.header_rows.last().unwrap()[2], "Count of Amount2");
        assert_eq!(grid.body_rows[0].values.len(), 2);
        assert_eq!(value_at(&grid.body_rows[0], 0), 50.0); // Sum for East
        assert_eq!(value_at(&grid.body_rows[0], 1), 4.0); // Count for East
    }

    #[test]
    fn test_row_labels_caption_replaces_outermost_row_field_name() {
        // Matches Excel's default "compact form" display (verified against
        // real Excel via fuzz/fuzz_pivot.py): the outermost row field's own
        // name never appears in the header at all -- it's always the
        // literal text "Row Labels".
        let sheet = source_sheet();
        let pivot = base_pivot(); // row_fields=[Region], col_fields=[]
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.header_rows.last().unwrap()[0], "Row Labels");
    }

    #[test]
    fn test_column_labels_row_prepended_and_deeper_row_field_keeps_its_name() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        pivot.col_fields = vec![PivotField::new("Rep")];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // Whenever there's at least one column field, Excel inserts an
        // extra header row above the column-value rows, captioned
        // "Column Labels".
        assert!(grid.header_rows[0].iter().any(|c| c == "Column Labels"));
        // Row-label captions land on the last header row: the outermost
        // row field ("Region") becomes "Row Labels", but a *deeper* row
        // field ("Product") keeps its own real name.
        let last = grid.header_rows.last().unwrap();
        assert_eq!(last[0], "Row Labels");
        assert_eq!(last[1], "Product");
    }

    #[test]
    fn test_grand_total_column_shows_total_prefixed_value_label_with_multiple_value_fields() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.col_fields = vec![PivotField::new("Product")];
        pivot.value_fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Min),
        ];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // The grand-total column's caption lands on the column-field row
        // (not repeated per value field as plain "Grand Total"), combining
        // "Total " with each value field's own label. Sum is first on
        // "Amount", so it clones the column for the following value field
        // (see `value_field_labels`'s doc comment), giving Min the
        // disambiguated "Amount2".
        let col_values_row = &grid.header_rows[1];
        assert!(col_values_row.contains(&"Total Sum of Amount".to_string()));
        assert!(col_values_row.contains(&"Total Min of Amount2".to_string()));
        // The value-label row directly below leaves the grand-total's
        // columns blank, since the caption already appeared above it.
        assert_eq!(grid.header_rows.last().unwrap().last().unwrap(), "");
    }

    #[test]
    fn test_grand_total_still_shows_with_only_one_leaf_group() {
        // Excel shows the grand total whenever the toggle is on, regardless of
        // how many groups it's summarizing (even with only one leaf group).
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.filter_fields = vec![PivotFilterField {
            column: "Region".to_string(),
            selected_values: Some(vec!["East".to_string()]),
            multiple_selection: true,
        }];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert!(grid.body_rows.iter().any(|r| r.is_grand_total));
    }

    #[test]
    fn test_case_variant_values_merge_using_globally_first_seen_casing() {
        // Case-insensitive grouping merges values consistently across all
        // branches using the field's first occurrence anywhere in the source data.
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 5,
            cols: 3,
            ..Default::default()
        });
        for (c, h) in ["Group", "Mixed", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        // "EAST" (uppercase) appears first in sheet order under Group=G1;
        // "east" (lowercase) appears later, nested under a *different*
        // Group=G2 branch.
        let rows: [[&str; 3]; 3] = [
            ["G1", "EAST", "10"],
            ["G1", "West", "20"],
            ["G2", "east", "30"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 2, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Group"), PivotField::new("Mixed")];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        let mixed_labels: Vec<&str> = grid
            .body_rows
            .iter()
            .map(|r| r.row_labels[1].as_str())
            .filter(|l| !l.is_empty())
            .collect();
        assert!(
            mixed_labels.contains(&"EAST") && !mixed_labels.contains(&"east"),
            "expected every occurrence to use the globally first-seen casing \"EAST\", got {mixed_labels:?}"
        );
    }

    #[test]
    fn test_case_canonicalization_uses_first_seen_casing_from_unfiltered_source_not_just_surviving_rows()
     {
        // Canonical casing for a case-insensitively merged group is determined
        // from the full source data field-wide, not just the filtered record set.
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 4,
            cols: 3,
            ..Default::default()
        });
        for (c, h) in ["Cat", "Mixed", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        // The true first occurrence of the "west"/"WEST" value is "WEST"
        // (row 1), but it's filtered out below (Cat="Alpha" excluded);
        // "west" (row 3, Cat="Beta", which survives the filter) must still
        // canonicalize to "WEST", not to itself.
        let rows: [[&str; 3]; 3] = [
            ["Alpha", "WEST", "10"],
            ["Beta", "East", "20"],
            ["Beta", "west", "30"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 2, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Mixed")];
        pivot.filter_fields = vec![PivotFilterField {
            column: "Cat".to_string(),
            selected_values: Some(vec!["Beta".to_string()]),
            multiple_selection: true,
        }];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        let labels: Vec<&str> = grid
            .body_rows
            .iter()
            .map(|r| r.row_labels[0].as_str())
            .collect();
        assert!(
            labels.contains(&"WEST") && !labels.contains(&"west"),
            "expected the filtered-out row's casing \"WEST\" to still win, got {labels:?}"
        );
    }

    #[test]
    fn test_blank_group_sorts_last_even_among_numeric_siblings() {
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 4,
            cols: 2,
            ..Default::default()
        });
        for (c, h) in ["Code", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        // 30 < ... numerically, but the blank row's Code cell is left
        // empty entirely -- deliberately out of numeric order so a sort
        // that just treated "(blank)" as any other value would put it
        // first (its group_key text "(blank)" sorts alphabetically before
        // digits) rather than last.
        sheet.set_cell_src(1, 0, "30".to_string());
        sheet.set_cell_src(1, 1, "1".to_string());
        sheet.set_cell_src(3, 0, "10".to_string());
        sheet.set_cell_src(3, 1, "3".to_string());
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 1, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Code")];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        let codes: Vec<&str> = grid
            .body_rows
            .iter()
            .map(|r| r.row_labels[0].as_str())
            .collect();
        assert_eq!(codes, vec!["10", "30", "(blank)"]);
    }

    #[test]
    fn test_negative_looking_text_sorts_last_among_text_siblings() {
        // Real Windows Excel sorts negative-looking text by its digits
        // with the '-' stripped ("7"), which happens to land it last
        // among these particular siblings -- see `text_sort_key` and the
        // next test for a case where stripped-sign placement is *not*
        // last.
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 6,
            cols: 2,
            ..Default::default()
        });
        for (c, h) in ["Code", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows: [(&str, &str); 5] = [
            ("\"-7\"", "1"),
            ("\".0152\"", "2"),
            ("\"13\"", "3"),
            ("\"34\"", "4"),
            ("\"4\"", "5"),
        ];
        for (r, (code, amount)) in rows.iter().enumerate() {
            sheet.set_cell_src(r + 1, 0, code.to_string());
            sheet.set_cell_src(r + 1, 1, amount.to_string());
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 5, 1, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Code")];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        let codes: Vec<&str> = grid
            .body_rows
            .iter()
            .map(|r| r.row_labels[0].as_str())
            .collect();
        assert_eq!(codes, vec![".0152", "13", "34", "4", "-7"]);
    }

    #[test]
    fn test_negative_looking_text_sorts_by_stripped_digits_not_last() {
        // Harvested from fuzz/fuzz_pivot.py's win32com (Windows) run, seed
        // 118859: among siblings "12" and "37", real Excel placed "-25"
        // *between* them, not after both -- comparing "-25" by its
        // stripped digit string "25" (which alphabetically falls between
        // "12" and "37") is what predicts this; a simpler "negative always
        // sorts last" rule (as in the previous test) would wrongly put
        // "-25" after "37" here.
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 4,
            cols: 2,
            ..Default::default()
        });
        for (c, h) in ["Code", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows: [(&str, &str); 3] = [("\"12\"", "1"), ("\"37\"", "2"), ("\"-25\"", "3")];
        for (r, (code, amount)) in rows.iter().enumerate() {
            sheet.set_cell_src(r + 1, 0, code.to_string());
            sheet.set_cell_src(r + 1, 1, amount.to_string());
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 1, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Code")];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        let codes: Vec<&str> = grid
            .body_rows
            .iter()
            .map(|r| r.row_labels[0].as_str())
            .collect();
        assert_eq!(codes, vec!["12", "-25", "37"]);
    }

    #[test]
    fn test_empty_row_col_intersection_renders_blank_not_zero_or_error() {
        // A row/column combination with zero underlying records (a sparse
        // cell in the cross-tab) renders as a genuinely blank cell in
        // Excel for every aggregation kind, not a computed zero or error
        // (verified against real Excel via fuzz/fuzz_pivot.py).
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Data".to_string()),
            rows: 3,
            cols: 3,
            ..Default::default()
        });
        for (c, h) in ["Region", "Product", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        // East only ever pairs with Widget; West only ever pairs with
        // Gadget -- so (East, Gadget) and (West, Widget) are both
        // genuinely empty intersections.
        let rows: [[&str; 3]; 2] = [["East", "Widget", "10"], ["West", "Gadget", "20"]];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 2, 2, true, false)
            .unwrap();

        let mut pivot = base_pivot();
        pivot.col_fields = vec![PivotField::new("Product")];
        pivot.value_fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Average),
        ];
        pivot.grand_totals_row = false;
        pivot.grand_totals_col = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // Row "East" only has Widget data, so both of its Gadget-column
        // cells (Sum and Average) must be blank.
        let east_row = grid
            .body_rows
            .iter()
            .find(|r| r.row_labels[0] == "East")
            .unwrap();
        for v in &east_row.values[..2] {
            assert!(
                matches!(v, ResultData::None),
                "expected blank for an empty intersection, got {v:?}"
            );
        }
    }

    #[test]
    fn test_value_field_labels_distinct_aggregations_without_sum_stay_unsuffixed() {
        // Reusing a source column across multiple value fields with
        // *different*, non-Sum aggregations produces distinct default
        // captions on its own ("Max of Amount", "Count of Amount"), so
        // real Excel leaves them alone -- no "Amount2" suffix.
        let fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Count),
            PivotValueField::new("Amount", PivotAggregation::Max),
        ];
        assert_eq!(
            value_field_labels(&fields),
            vec!["Count of Amount".to_string(), "Max of Amount".to_string()]
        );
    }

    #[test]
    fn test_value_field_labels_sum_clones_column_for_later_fields() {
        // Unlike other aggregations, the *first* value field on a column that
        // uses `Sum` clones that column ("Amount" -> "Amount2") for every value
        // field *after* it in the list, regardless of their own aggregation.
        // "Rate" here has no Sum field at all, so it's unaffected and stays plain.
        let fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Rate", PivotAggregation::Average),
            PivotValueField::new("Amount", PivotAggregation::Min),
            PivotValueField::new("Amount", PivotAggregation::Max),
        ];
        assert_eq!(
            value_field_labels(&fields),
            vec![
                "Sum of Amount".to_string(),
                "Average of Rate".to_string(),
                "Min of Amount2".to_string(),
                "Max of Amount2".to_string(),
            ]
        );
    }

    #[test]
    fn test_value_field_labels_second_sum_clones_again() {
        // A second `Sum` value field on the same column clones *again*
        // ("Amount2" -> "Amount3"), rather than reusing the first clone --
        // verified by direct real-Excel probing (see the test above).
        let fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Count),
        ];
        assert_eq!(
            value_field_labels(&fields),
            vec![
                "Sum of Amount".to_string(),
                "Sum of Amount2".to_string(),
                "Count of Amount3".to_string(),
            ]
        );
    }

    #[test]
    fn test_value_field_labels_disambiguates_identical_aggregation_and_column() {
        // Two value fields on the same column with the *same* aggregation
        // do produce an identical default caption ("Sum of Amount" twice),
        // so this is the one shape where real Excel's plain digit-suffix
        // disambiguation kicks in even without any preceding clone.
        let fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Sum),
        ];
        assert_eq!(
            value_field_labels(&fields),
            vec![
                "Sum of Amount".to_string(),
                "Sum of Amount2".to_string(),
                "Sum of Amount3".to_string(),
            ]
        );
    }

    #[test]
    fn test_value_field_labels_collision_within_sum_clone_uses_underscore_suffix() {
        // When a caption collision happens *inside* an already Sum-cloned
        // slot (two non-Sum fields on the same clone sharing an
        // aggregation), real Excel disambiguates by appending an
        // underscored counter to the whole already-suffixed caption
        // instead of incrementing the clone number again -- verified by
        // direct real-Excel probing.
        let fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Max),
            PivotValueField::new("Amount", PivotAggregation::Max),
        ];
        assert_eq!(
            value_field_labels(&fields),
            vec![
                "Sum of Amount".to_string(),
                "Max of Amount2".to_string(),
                "Max of Amount2_2".to_string(),
            ]
        );
    }

    #[test]
    fn test_value_field_labels_count_numbers_shares_plain_count_caption() {
        // Excel's default caption for the "Count Numbers" summary function
        // is "Count of <field>" -- identical to plain "Count" -- not
        // "Count Numbers of <field>". Since both aggregations generate
        // the same caption text, using both on the same column is exactly
        // the collide-and-suffix case above.
        let fields = vec![
            PivotValueField::new("Rate", PivotAggregation::CountNumbers),
            PivotValueField::new("Rate", PivotAggregation::Count),
        ];
        assert_eq!(
            value_field_labels(&fields),
            vec!["Count of Rate".to_string(), "Count of Rate2".to_string()]
        );
    }

    #[test]
    fn test_value_field_labels_leaves_custom_name_untouched() {
        let mut fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Min),
        ];
        fields[1].custom_name = Some("Lowest Amount".to_string());
        assert_eq!(
            value_field_labels(&fields),
            vec!["Sum of Amount".to_string(), "Lowest Amount".to_string()]
        );
    }

    #[test]
    fn test_flat_pivot_with_no_row_or_col_fields_has_no_reserved_label_column() {
        // With neither row nor column fields (a single aggregate value, no
        // grouping at all), Excel doesn't reserve a separate row-label
        // column the way it does whenever *either* axis has fields -- the
        // value field's own header sits directly above the value, one column
        // wide total.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        assert_eq!(grid.width, 1);
        assert_eq!(
            grid.header_rows.last().unwrap(),
            &vec!["Sum of Amount".to_string()]
        );
        assert_eq!(grid.body_rows.len(), 1);
        assert!(grid.body_rows[0].row_labels.is_empty());
        assert_eq!(value_at(&grid.body_rows[0], 0), 195.0);
    }

    #[test]
    fn test_no_row_fields_with_multiple_value_fields_has_no_reserved_label_column_either() {
        // Unlike the single-value-field case (which reserves one corner
        // column for that field's own label, e.g. "Max of Amount"), with
        // *multiple* value fields and no row fields there's no single
        // unambiguous label to put in a corner -- each value field's label
        // already shows up in its own column further along the header -- so
        // Excel reserves no column for it at all, regardless of whether
        // column fields are present.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![];
        pivot.col_fields = vec![PivotField::new("Product")];
        pivot.value_fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Count),
        ];
        pivot.grand_totals_col = false;
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        // width = 0 reserved + 2 column groups (Gadget, Widget) * 2 value
        // fields.
        assert_eq!(grid.width, 4);
        assert_eq!(grid.body_rows.len(), 1);
        assert!(grid.body_rows[0].row_labels.is_empty());
    }

    #[test]
    fn test_multiple_value_fields_with_no_column_fields_share_one_header_row() {
        // With no column fields at all there's no column-group-values row
        // in the first place, so Excel lists each value field as a
        // plain adjacent column in the single header row, like an ordinary
        // flat table.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.value_fields = vec![
            PivotValueField::new("Amount", PivotAggregation::Sum),
            PivotValueField::new("Amount", PivotAggregation::Count),
        ];
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();

        assert_eq!(grid.header_rows.len(), 1);
        // Sum is first on "Amount", so it clones the column for the
        // following value field (see `value_field_labels`'s doc comment).
        assert_eq!(
            grid.header_rows[0],
            vec![
                "Row Labels".to_string(),
                "Sum of Amount".to_string(),
                "Count of Amount2".to_string(),
            ]
        );
    }

    #[test]
    fn test_missing_column_errors() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Nope")];
        let err = compute_pivot(&[&sheet], &pivot).unwrap_err();
        assert!(err.contains("not found"));
    }

    #[test]
    fn test_range_source_matches_table_source() {
        // A pivot sourced from a raw range covering exactly a table's
        // declared bounds must produce the same grid as one sourced from
        // the table itself.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.source = PivotSource::Range {
            sheet_id: sheet.id,
            start_row: 0,
            start_col: 0,
            end_row: 8,
            end_col: 3,
        };
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        assert_eq!(grid.body_rows.len(), 3);
        assert_eq!(value_at(&grid.body_rows[0], 0), 50.0);
        assert_eq!(value_at(&grid.body_rows[1], 0), 145.0);
        assert_eq!(value_at(&grid.body_rows[2], 0), 195.0);
    }

    #[test]
    fn test_zero_data_rows_produces_empty_grid_without_panicking() {
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Empty".to_string()),
            rows: 1,
            cols: 2,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "Region".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet.commit(None).unwrap();
        sheet
            .add_table("Empty".to_string(), 0, 0, 0, 1, true, false)
            .unwrap();

        let pivot = PivotTable {
            id: 1,
            name: "EmptyPivot".to_string(),
            source: PivotSource::Table {
                name: "Empty".to_string(),
            },
            dest_sheet_id: sheet.id,
            dest_row: 0,
            dest_col: 0,
            row_fields: vec![PivotField::new("Region")],
            col_fields: vec![],
            value_fields: vec![PivotValueField::new("Amount", PivotAggregation::Sum)],
            filter_fields: vec![],
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        };
        let grid = compute_pivot(&[&sheet], &pivot).unwrap();
        // No records at all -> no groups, and (per `build_axis`) a grand
        // total is only appended when there's more than one group, so none
        // is emitted here either.
        assert!(grid.body_rows.is_empty());
        assert!(grid.row_axis.is_empty());
    }

    // ---- Randomized invariant fuzzing --------------------------------
    //
    // Builds many random source sheets + pivot configurations and checks
    // internal self-consistency (never panics; every output cell, whether
    // leaf/subtotal/grand-total, equals an independently-derived aggregate
    // over the same filtered records; xlsx export/import round-trips
    // field assignments faithfully). This is a self-consistency fuzzer,
    // not a check against real Excel -- that's `fuzz/fuzz_pivot.py`'s job
    // -- but it's cheap to run in `cargo test` and catches crashes and logic
    // issues in the group-tree flattening/subtotal/grand-total code.
    use rand::rngs::StdRng;
    use rand::{Rng, SeedableRng};

    const FUZZ_COLS: [&str; 6] = ["Cat", "Mixed", "NumStr", "Amount", "Rate", "Flag"];
    const FUZZ_CATEGORIES: [&str; 5] = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon"];
    const FUZZ_CASE_VARIANTS: [&str; 5] = ["East", "east", "WEST", "west", "North"];

    /// Builds a random source sheet with columns chosen to exercise
    /// grouping edge cases: a low-cardinality category column with
    /// occasional blanks, a case-variant category column (case-insensitive
    /// grouping parity), a quoted numeric-looking-string column (the
    /// numeric-vs-text sort ambiguity `sort_group_entries` has to resolve),
    /// two numeric columns (ints and floats, including negative/zero), and
    /// a boolean column (ignored by Sum/Average/Max/Min).
    fn fuzz_source_sheet(rng: &mut StdRng, num_rows: usize) -> (Sheet, Vec<String>) {
        let mut sheet = Sheet::new(SheetInit {
            name: Some("FuzzData".to_string()),
            rows: num_rows + 1,
            cols: FUZZ_COLS.len(),
            ..Default::default()
        });
        for (c, h) in FUZZ_COLS.iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        for r in 0..num_rows {
            let cat = if rng.gen_bool(0.1) {
                String::new()
            } else {
                FUZZ_CATEGORIES[rng.gen_range(0..FUZZ_CATEGORIES.len())].to_string()
            };
            sheet.set_cell_src(r + 1, 0, cat);

            let mixed = FUZZ_CASE_VARIANTS[rng.gen_range(0..FUZZ_CASE_VARIANTS.len())].to_string();
            sheet.set_cell_src(r + 1, 1, mixed);

            let numstr = match rng.gen_range(0u8..4u8) {
                0 => String::new(),
                1 => format!("\"0{}\"", rng.gen_range(0u32..10u32)),
                2 => format!("\".0{}\"", rng.gen_range(0u32..1000u32)),
                _ => format!("\"{}\"", rng.gen_range(-50i64..50i64)),
            };
            sheet.set_cell_src(r + 1, 2, numstr);

            sheet.set_cell_src(r + 1, 3, rng.gen_range(-100i64..=100i64).to_string());

            let rate =
                (rng.gen_range(-500i64..=500i64) as f64) / (rng.gen_range(1i64..=100i64) as f64);
            sheet.set_cell_src(r + 1, 4, format!("{:.4}", rate));

            sheet.set_cell_src(r + 1, 5, rng.gen_bool(0.5).to_string());
        }
        sheet.commit(None).unwrap();
        (sheet, FUZZ_COLS.iter().map(|s| s.to_string()).collect())
    }

    fn random_aggregation(rng: &mut StdRng) -> PivotAggregation {
        match rng.gen_range(0u8..6u8) {
            0 => PivotAggregation::Sum,
            1 => PivotAggregation::Count,
            2 => PivotAggregation::CountNumbers,
            3 => PivotAggregation::Average,
            4 => PivotAggregation::Max,
            _ => PivotAggregation::Min,
        }
    }

    /// Builds a random, always-valid `PivotTable` config over `sheet`:
    /// 0-2 row fields and 0-2 col fields (drawn without replacement from
    /// the categorical columns), 1-2 value fields (from the numeric
    /// columns), an optional filter field with a random subset of its
    /// actual distinct values selected (including the all-excluded case),
    /// and random per-field subtotal / grand-total toggles.
    fn fuzz_pivot_config(
        rng: &mut StdRng,
        sheet: &Sheet,
        col_names: &[String],
        num_rows: usize,
        use_table: bool,
    ) -> PivotTable {
        let mut pool: Vec<usize> = vec![0, 1, 2]; // Cat, Mixed, NumStr
        let numeric: [usize; 2] = [3, 4]; // Amount, Rate

        let n_row = rng.gen_range(0..=pool.len().min(2));
        let row_cols: Vec<usize> = (0..n_row)
            .map(|_| pool.remove(rng.gen_range(0..pool.len())))
            .collect();
        let n_col = rng.gen_range(0..=pool.len().min(2));
        let col_cols: Vec<usize> = (0..n_col)
            .map(|_| pool.remove(rng.gen_range(0..pool.len())))
            .collect();

        let row_fields: Vec<PivotField> = row_cols
            .iter()
            .map(|&i| PivotField {
                column: col_names[i].clone(),
                subtotal: rng.gen_bool(0.7),
            })
            .collect();
        let col_fields: Vec<PivotField> = col_cols
            .iter()
            .map(|&i| PivotField {
                column: col_names[i].clone(),
                subtotal: rng.gen_bool(0.7),
            })
            .collect();

        let n_value = rng.gen_range(1..=2);
        let value_fields: Vec<PivotValueField> = (0..n_value)
            .map(|_| {
                let col = numeric[rng.gen_range(0..numeric.len())];
                PivotValueField::new(col_names[col].clone(), random_aggregation(rng))
            })
            .collect();

        let mut filter_fields = Vec::new();
        if rng.gen_bool(0.5) {
            let candidates = [0usize, 1, 2, 5];
            let fcol = candidates[rng.gen_range(0..candidates.len())];
            let mut distinct: Vec<String> = (1..=num_rows)
                .map(|r| group_key(&sheet.get_result_data(&CellRef::new(r, fcol))))
                .collect();
            distinct.sort();
            distinct.dedup();
            let selected = if distinct.is_empty() || rng.gen_bool(0.2) {
                None
            } else {
                // May legitimately come out empty -> filters out every record.
                Some(distinct.into_iter().filter(|_| rng.gen_bool(0.5)).collect())
            };
            filter_fields.push(PivotFilterField {
                column: col_names[fcol].clone(),
                selected_values: selected,
                multiple_selection: true,
            });
        }

        let source = if use_table {
            PivotSource::Table {
                name: "FuzzTable".to_string(),
            }
        } else {
            PivotSource::Range {
                sheet_id: sheet.id,
                start_row: 0,
                start_col: 0,
                end_row: num_rows,
                end_col: col_names.len() - 1,
            }
        };

        PivotTable {
            id: 1,
            name: "FuzzPivot".to_string(),
            source,
            dest_sheet_id: sheet.id,
            dest_row: num_rows + 20,
            dest_col: 0,
            row_fields,
            col_fields,
            value_fields,
            filter_fields,
            grand_totals_row: rng.gen_bool(0.7),
            grand_totals_col: rng.gen_bool(0.7),
            last_output_end_row: None,
            last_output_end_col: None,
        }
    }

    fn results_close(a: &ResultData, b: &ResultData) -> bool {
        match (a, b) {
            (ResultData::Integer(x), ResultData::Integer(y)) => x == y,
            (ResultData::Float(x), ResultData::Float(y)) => (x - y).abs() < 1e-6,
            (ResultData::Integer(x), ResultData::Float(y))
            | (ResultData::Float(y), ResultData::Integer(x)) => (*x as f64 - y).abs() < 1e-6,
            (ResultData::None, ResultData::None) => true,
            (ResultData::Error(x), ResultData::Error(y)) => x == y,
            (ResultData::String(x), ResultData::String(y)) => x == y,
            (ResultData::Boolean(x), ResultData::Boolean(y)) => x == y,
            _ => false,
        }
    }

    /// A row/col axis label vector (`Some` per own depth, `None` past it --
    /// see `FlatGroup`) is a *partial key*: `None` positions are wildcards.
    /// This is exactly what a subtotal or grand-total group represents, so
    /// the same matcher works uniformly for leaf, subtotal, and grand-total
    /// groups.
    fn matches_partial(key: &[String], labels: &[Option<String>]) -> bool {
        // Case-insensitive, matching `build_group_tree`'s merge: an axis
        // label is whichever casing was first seen for that group, so a
        // record whose own key differs only in case must still match it.
        key.iter()
            .zip(labels)
            .all(|(k, want)| want.as_ref().is_none_or(|w| w.eq_ignore_ascii_case(k)))
    }

    /// Cross-checks every cell of `grid` against an aggregate computed by a
    /// structurally independent path: instead of `compute_pivot`'s
    /// recursive group-tree + flatten, this filters the same record set by
    /// simple partial-key matching against each axis item's labels. Catches
    /// bugs in the tree-based grouping/flattening/subtotal-insertion logic
    /// specifically, since the aggregation math itself (`aggregate`) is
    /// shared and already covered by the fixed-data tests above.
    fn verify_grid_matches_records(sheet: &Sheet, pivot: &PivotTable, grid: &PivotGrid) {
        let (_, col_names, sheet_cols, data_rows) =
            resolve_source(&[sheet], &pivot.source).unwrap();
        let row_idxs: Vec<usize> = pivot
            .row_fields
            .iter()
            .map(|f| column_index(&col_names, &f.column).unwrap())
            .collect();
        let col_idxs: Vec<usize> = pivot
            .col_fields
            .iter()
            .map(|f| column_index(&col_names, &f.column).unwrap())
            .collect();

        let mut records: Vec<(Vec<String>, Vec<String>, Vec<ResultData>)> = Vec::new();
        'row: for &r in &data_rows {
            let row_vals: Vec<ResultData> = sheet_cols
                .iter()
                .map(|&c| sheet.get_result_data(&CellRef::new(r, c)))
                .collect();
            for ff in &pivot.filter_fields {
                if let Some(selected) = &ff.selected_values {
                    let idx = column_index(&col_names, &ff.column).unwrap();
                    let key = group_key(&row_vals[idx]);
                    // Case-insensitive, matching `compute_pivot`'s own filter
                    // step (a filter field's items are merged case-different
                    // text, same as row/col group labels).
                    if !selected.iter().any(|v| v.eq_ignore_ascii_case(&key)) {
                        continue 'row;
                    }
                }
            }
            let row_key: Vec<String> = row_idxs.iter().map(|&i| group_key(&row_vals[i])).collect();
            let col_key: Vec<String> = col_idxs.iter().map(|&i| group_key(&row_vals[i])).collect();
            records.push((row_key, col_key, row_vals));
        }

        let value_idxs: Vec<usize> = pivot
            .value_fields
            .iter()
            .map(|vf| column_index(&col_names, &vf.column).unwrap())
            .collect();
        let value_multiplier = if pivot.value_fields.len() > 1 {
            pivot.value_fields.len()
        } else {
            1
        };
        let width = row_label_width(pivot);

        assert_eq!(grid.body_rows.len(), grid.row_axis.len());
        assert_eq!(grid.width, width + grid.col_axis.len() * value_multiplier);
        for hrow in &grid.header_rows {
            assert_eq!(hrow.len(), grid.width);
        }

        for (i, (body_row, row_axis)) in grid.body_rows.iter().zip(grid.row_axis.iter()).enumerate()
        {
            assert_eq!(
                body_row.is_grand_total, row_axis.is_grand_total,
                "row {i} grand-total flag mismatch"
            );
            assert_eq!(body_row.row_labels.len(), width, "row {i} label width");
            assert_eq!(
                body_row.values.len(),
                grid.col_axis.len() * value_multiplier,
                "row {i} value count"
            );

            for (j, col_axis) in grid.col_axis.iter().enumerate() {
                let matching: Vec<&Vec<ResultData>> = records
                    .iter()
                    .filter(|(rk, ck, _)| {
                        matches_partial(rk, &row_axis.labels)
                            && matches_partial(ck, &col_axis.labels)
                    })
                    .map(|(_, _, row)| row)
                    .collect();

                for (vf_pos, &vidx) in value_idxs.iter().enumerate() {
                    if vf_pos > 0 && value_multiplier == 1 {
                        break;
                    }
                    let col_vals: Vec<ResultData> =
                        matching.iter().map(|row| row[vidx].clone()).collect();
                    let expected =
                        aggregate(sheet, &col_vals, pivot.value_fields[vf_pos].aggregation);
                    let actual = &body_row.values[j * value_multiplier + vf_pos];
                    assert!(
                        results_close(&expected, actual),
                        "row {i} col {j} value-field {vf_pos}: expected {expected:?}, got {actual:?} \
                         (row_labels={:?}, col_labels={:?})",
                        row_axis.labels,
                        col_axis.labels,
                    );
                }
            }
        }
    }

    /// A grand-total pseudo-group is appended whenever the toggle is on,
    /// *except* when the axis has no fields at all (`build_axis`'s
    /// no-fields early return never adds one -- there's no separate
    /// grouping to total distinctly from the single implicit group).
    /// Otherwise Excel shows it regardless of how many real groups exist,
    /// even just one (confirmed against real Excel via fuzz/fuzz_pivot.py).
    fn verify_grand_total_placement(
        axis: &[PivotAxisItem],
        grand_total_requested: bool,
        axis_has_fields: bool,
        label: &str,
    ) {
        let grand_count = axis.iter().filter(|a| a.is_grand_total).count();
        let has_any_real_group = axis.iter().any(|a| !a.is_grand_total);
        assert!(grand_count <= 1, "{label}: more than one grand-total group");
        if grand_total_requested && axis_has_fields && has_any_real_group {
            assert_eq!(
                grand_count, 1,
                "{label}: expected a grand total to be appended"
            );
        } else {
            assert_eq!(grand_count, 0, "{label}: did not expect a grand total");
        }
    }

    #[test]
    fn test_fuzz_pivot_random_invariants() {
        for seed in 0u64..300 {
            let mut rng: StdRng = SeedableRng::seed_from_u64(seed);
            let use_table = seed % 2 == 0;
            // Zero-data-row (header-only) Excel Table coverage.
            let num_rows = rng.gen_range(0..=40usize);
            let (mut sheet, col_names) = fuzz_source_sheet(&mut rng, num_rows);
            if use_table {
                sheet
                    .add_table(
                        "FuzzTable".to_string(),
                        0,
                        0,
                        num_rows,
                        col_names.len() - 1,
                        true,
                        false,
                    )
                    .unwrap();
            }
            let pivot = fuzz_pivot_config(&mut rng, &sheet, &col_names, num_rows, use_table);

            let grid = compute_pivot(&[&sheet], &pivot)
                .unwrap_or_else(|e| panic!("seed {seed}: compute_pivot failed: {e}"));

            verify_grid_matches_records(&sheet, &pivot, &grid);
            verify_grand_total_placement(
                &grid.row_axis,
                pivot.grand_totals_row,
                !pivot.row_fields.is_empty(),
                "row axis",
            );
            verify_grand_total_placement(
                &grid.col_axis,
                pivot.grand_totals_col,
                !pivot.col_fields.is_empty(),
                "col axis",
            );

            // Round-trip through xlsx export/import: field/aggregation
            // assignments, grand-total flags, and subtotal toggles must
            // survive; filter selections are documented (pivot_xlsx.rs) as
            // resetting to "all" rather than surviving.
            let xlsx = crate::core::xlsx::export_xlsx_data(
                std::slice::from_ref(&sheet),
                &[],
                std::slice::from_ref(&pivot),
                None,
            )
            .unwrap_or_else(|e| panic!("seed {seed}: export failed: {e}"));
            let (imported_sheets, _, imported_pivots, _) =
                crate::core::xlsx::import_xlsx_data(&xlsx, &[], |_, _, _| {})
                    .unwrap_or_else(|e| panic!("seed {seed}: import failed: {e}"));
            assert_eq!(
                imported_pivots.len(),
                1,
                "seed {seed}: pivot lost on round-trip"
            );
            let reimported = &imported_pivots[0];

            assert_eq!(
                reimported
                    .row_fields
                    .iter()
                    .map(|f| &f.column)
                    .collect::<Vec<_>>(),
                pivot
                    .row_fields
                    .iter()
                    .map(|f| &f.column)
                    .collect::<Vec<_>>(),
                "seed {seed}: row field columns changed on round-trip"
            );
            assert_eq!(
                reimported
                    .col_fields
                    .iter()
                    .map(|f| &f.column)
                    .collect::<Vec<_>>(),
                pivot
                    .col_fields
                    .iter()
                    .map(|f| &f.column)
                    .collect::<Vec<_>>(),
                "seed {seed}: col field columns changed on round-trip"
            );
            assert_eq!(
                reimported
                    .value_fields
                    .iter()
                    .map(|f| (&f.column, f.aggregation))
                    .collect::<Vec<_>>(),
                pivot
                    .value_fields
                    .iter()
                    .map(|f| (&f.column, f.aggregation))
                    .collect::<Vec<_>>(),
                "seed {seed}: value fields changed on round-trip"
            );
            assert_eq!(reimported.grand_totals_row, pivot.grand_totals_row);
            assert_eq!(reimported.grand_totals_col, pivot.grand_totals_col);
            assert_eq!(
                reimported
                    .row_fields
                    .iter()
                    .map(|f| f.subtotal)
                    .collect::<Vec<_>>(),
                pivot
                    .row_fields
                    .iter()
                    .map(|f| f.subtotal)
                    .collect::<Vec<_>>(),
                "seed {seed}: row field subtotal toggle should round-trip"
            );
            assert_eq!(
                reimported
                    .col_fields
                    .iter()
                    .map(|f| f.subtotal)
                    .collect::<Vec<_>>(),
                pivot
                    .col_fields
                    .iter()
                    .map(|f| f.subtotal)
                    .collect::<Vec<_>>(),
                "seed {seed}: col field subtotal toggle should round-trip"
            );
            // If nothing lossy was actually in play, the reimported grid
            // must be structurally identical -- this is where a genuine
            // round-trip bug (e.g. losing a value field's aggregation)
            // would show up as a shape mismatch rather than a field-list
            // diff the assertions above already caught. Subtotal toggles
            // now round-trip exactly, so only filter selections remain
            // lossy.
            // Filter selections round-trip now too, so a lossless round trip
            // is the ordinary case rather than the exception.
            let nothing_lossy = true;
            let any_filter_is_also_an_axis_field = pivot.filter_fields.iter().any(|ff| {
                pivot
                    .row_fields
                    .iter()
                    .chain(pivot.col_fields.iter())
                    .any(|f| f.column.eq_ignore_ascii_case(&ff.column))
            });
            let reimported_sheets: Vec<Sheet> =
                imported_sheets.into_iter().map(|s| s.sheet).collect();
            let reimported_sheet_refs: Vec<&Sheet> = reimported_sheets.iter().collect();
            let reimported_grid = compute_pivot(&reimported_sheet_refs, reimported)
                .unwrap_or_else(|e| panic!("seed {seed}: reimported compute_pivot failed: {e}"));
            // A field Excel could not represent at all -- see
            // `axis_bound` below -- is excluded from the shape check for the
            // same reason it is excluded from the selection check.
            if nothing_lossy && !any_filter_is_also_an_axis_field {
                assert_eq!(
                    reimported_grid.body_rows.len(),
                    grid.body_rows.len(),
                    "seed {seed}: grid shape changed on lossless round-trip"
                );
            }

            // Filter selections round-trip now: they are written as indices
            // into the cache's `<sharedItems>` and resolved back to plain
            // values on import. What must match is the *set* of selected
            // values, since the file stores them in the cache's first-seen
            // order rather than the caller's.
            //
            // The one legitimate difference: a selection covering every
            // value marks nothing hidden, so it is indistinguishable from no
            // filter once written and comes back as `None`. That is only
            // acceptable if it really was a no-op, which the grid proves.
            // Compared case-insensitively, because the engine merges
            // case-variant values into one item (keyed by the first casing
            // seen in the source). So a selection naming both `WEST` and
            // `west` picks a single item and legitimately reads back as
            // whichever casing the cache stored -- a canonicalization, not a
            // loss.
            let sorted = |f: &PivotFilterField| {
                f.selected_values.as_ref().map(|v| {
                    let mut v: Vec<String> = v.iter().map(|s| s.to_lowercase()).collect();
                    v.sort();
                    v.dedup();
                    v
                })
            };
            // A filter column that is *also* a row or column field has no
            // representation in the file: a pivot field carries one `axis`,
            // so the row/column orientation wins and there is nowhere left to
            // record the selection. Excel cannot express that config either
            // -- a field has exactly one orientation there -- so this is a
            // shape visi's model admits and the format does not, rather than
            // a round-trip bug.
            let axis_bound = |column: &str| {
                pivot
                    .row_fields
                    .iter()
                    .chain(pivot.col_fields.iter())
                    .any(|f| f.column.eq_ignore_ascii_case(column))
            };
            for (before, after) in pivot
                .filter_fields
                .iter()
                .zip(reimported.filter_fields.iter())
            {
                if axis_bound(&before.column) {
                    continue;
                }
                if before.selected_values.is_some() && after.selected_values.is_none() {
                    assert_eq!(
                        reimported_grid.body_rows.len(),
                        grid.body_rows.len(),
                        "seed {seed}: filter on '{}' was dropped and it mattered",
                        before.column
                    );
                } else {
                    assert_eq!(
                        sorted(before),
                        sorted(after),
                        "seed {seed}: filter selection should round-trip for '{}'",
                        before.column
                    );
                }
            }
        }
    }
}
