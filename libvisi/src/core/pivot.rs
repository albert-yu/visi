use serde::{Deserialize, Serialize};

use crate::core::engine::{CellRef, ResultData, Sheet};

/// Where a `PivotTable` reads its source records from: either an existing
/// `ExcelTable` (looked up by name at compute time, so renames/resizes of
/// the table are picked up automatically on refresh) or a plain cell range
/// whose first row is treated as column headers.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum PivotSource {
    Table {
        name: String,
    },
    Range {
        sheet_id: u64,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
}

/// Matches the "Summarize value field by" choices Excel exposes for a data
/// field; the five most commonly used ones plus the numeric-only count.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum PivotAggregation {
    Sum,
    Count,
    CountNumbers,
    Average,
    Max,
    Min,
}

impl PivotAggregation {
    pub fn label(&self) -> &'static str {
        match self {
            PivotAggregation::Sum => "Sum",
            PivotAggregation::Count => "Count",
            PivotAggregation::CountNumbers => "Count Numbers",
            PivotAggregation::Average => "Average",
            PivotAggregation::Max => "Max",
            PivotAggregation::Min => "Min",
        }
    }

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
    pub column: String,
    /// Whether a subtotal line is emitted for this field when it isn't the
    /// innermost field in its area (Excel's per-field "Subtotals" toggle).
    pub subtotal: bool,
}

impl PivotField {
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
    pub column: String,
    pub aggregation: PivotAggregation,
    pub custom_name: Option<String>,
}

impl PivotValueField {
    pub fn new(column: impl Into<String>, aggregation: PivotAggregation) -> Self {
        Self {
            column: column.into(),
            aggregation,
            custom_name: None,
        }
    }

    pub fn label(&self) -> String {
        self.custom_name
            .clone()
            .unwrap_or_else(|| format!("{} of {}", self.aggregation.label(), self.column))
    }
}

/// One field placed in the Filter (Page) area.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PivotFilterField {
    pub column: String,
    /// `None` means every value is allowed (no filtering applied yet).
    pub selected_values: Option<Vec<String>>,
}

impl PivotFilterField {
    pub fn new(column: impl Into<String>) -> Self {
        Self {
            column: column.into(),
            selected_values: None,
        }
    }
}

/// The area of a pivot table a field can be assigned to, used by the
/// add/remove-field CRUD operations.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum PivotArea {
    Row,
    Column,
    Value,
    Filter,
}

/// A pivot table definition: a summary of `source`, grouped by `row_fields`
/// nested within `col_fields`, restricted by `filter_fields`, and
/// aggregated per `value_fields`. This is a workbook-level object (like
/// `Chart`) rather than sheet-scoped like `ExcelTable`, since its source and
/// destination ranges may live on different sheets.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PivotTable {
    pub id: u64,
    pub name: String,
    pub source: PivotSource,
    pub dest_sheet_id: u64,
    pub dest_row: usize,
    pub dest_col: usize,
    pub row_fields: Vec<PivotField>,
    pub col_fields: Vec<PivotField>,
    pub value_fields: Vec<PivotValueField>,
    pub filter_fields: Vec<PivotFilterField>,
    pub grand_totals_row: bool,
    pub grand_totals_col: bool,
    /// Bottom-right corner of the last rendered output grid, so a refresh
    /// that produces a smaller grid can clear the now-stale cells.
    #[serde(default)]
    pub last_output_end_row: Option<usize>,
    #[serde(default)]
    pub last_output_end_col: Option<usize>,
}

/// A fully computed pivot result, ready to be materialized into a sheet:
/// `header_rows` come first, then one entry of `body_rows` per output row.
#[derive(Debug, Clone)]
pub struct PivotGrid {
    pub header_rows: Vec<Vec<String>>,
    pub body_rows: Vec<PivotBodyRow>,
    /// Total width in columns (row-label columns + data columns), used by
    /// the caller to know how large a range to clear/allocate.
    pub width: usize,
    /// The flattened row/column axis groups underlying `body_rows`/the data
    /// columns, exposed (independent of display formatting) so an xlsx
    /// exporter can reconstruct a native `pivotTableDefinition`'s
    /// `rowItems`/`colItems` without re-deriving the grouping itself.
    pub row_axis: Vec<PivotAxisItem>,
    pub col_axis: Vec<PivotAxisItem>,
}

#[derive(Debug, Clone)]
pub struct PivotBodyRow {
    /// One entry per row field (or a single "Grand Total" entry when there
    /// are no row fields); blank entries mean "same as the row above".
    pub row_labels: Vec<String>,
    pub is_grand_total: bool,
    /// One entry per data column, aligned with the last `header_rows` row.
    pub values: Vec<ResultData>,
}

/// One flattened group along a row or column axis: a label per axis field
/// (`None` past its own depth), plus whether it's a subtotal or grand-total
/// pseudo-group rather than a real leaf group.
#[derive(Debug, Clone)]
pub struct PivotAxisItem {
    pub labels: Vec<Option<String>>,
    pub is_subtotal: bool,
    pub is_grand_total: bool,
}

impl PivotGrid {
    pub fn height(&self) -> usize {
        self.header_rows.len() + self.body_rows.len()
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

fn sort_group_entries(pairs: &mut [(String, Vec<usize>)]) {
    let all_numeric = pairs
        .iter()
        .all(|(k, _)| k != "(blank)" && k.trim().parse::<f64>().is_ok());
    if all_numeric {
        pairs.sort_by(|a, b| {
            let fa: f64 = a.0.trim().parse().unwrap_or(0.0);
            let fb: f64 = b.0.trim().parse().unwrap_or(0.0);
            fa.partial_cmp(&fb).unwrap_or(std::cmp::Ordering::Equal)
        });
    } else {
        pairs.sort_by_key(|(k, _)| k.to_lowercase());
    }
}

fn build_group_tree(
    indices: &[usize],
    keys: &[Vec<String>],
    depth: usize,
    num_fields: usize,
) -> Vec<GroupNode> {
    let mut groups: Vec<(String, Vec<usize>)> = Vec::new();
    for &idx in indices {
        let key = &keys[idx][depth];
        if let Some(entry) = groups.iter_mut().find(|(k, _)| k == key) {
            entry.1.push(idx);
        } else {
            groups.push((key.clone(), vec![idx]));
        }
    }
    sort_group_entries(&mut groups);
    groups
        .into_iter()
        .map(|(label, idxs)| {
            let children = if depth + 1 < num_fields {
                build_group_tree(&idxs, keys, depth + 1, num_fields)
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
) -> Vec<FlatGroup> {
    if fields.is_empty() {
        return vec![FlatGroup {
            labels: Vec::new(),
            record_indices: record_indices.to_vec(),
            is_subtotal: false,
            is_grand_total: false,
        }];
    }
    let tree = build_group_tree(record_indices, keys, 0, fields.len());
    let mut flat = Vec::new();
    flatten_groups(&tree, fields, 0, fields.len(), &[], &mut flat);
    if grand_total && flat.len() > 1 {
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
    sheets: &'a [Sheet],
    source: &PivotSource,
) -> Result<ResolvedSource<'a>, String> {
    match source {
        PivotSource::Table { name } => {
            let (sheet, table) = sheets
                .iter()
                .find_map(|s| s.find_table(name).map(|t| (s, t)))
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
            let sheet = sheets
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
pub fn compute_pivot(sheets: &[Sheet], pivot: &PivotTable) -> Result<PivotGrid, String> {
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

    // Read every source record (one row per data row, one ResultData per
    // source column), applying filter fields as we go.
    let mut records: Vec<Vec<ResultData>> = Vec::new();
    'row: for &r in &data_rows {
        let mut row_vals = Vec::with_capacity(sheet_cols.len());
        for &c in &sheet_cols {
            row_vals.push(sheet.get_result_data(&CellRef::new(r, c)));
        }
        for ff in &pivot.filter_fields {
            if let Some(selected) = &ff.selected_values {
                let idx = column_index(&col_names, &ff.column)?;
                let key = group_key(&row_vals[idx]);
                if !selected.iter().any(|v| v == &key) {
                    continue 'row;
                }
            }
        }
        records.push(row_vals);
    }

    let record_indices: Vec<usize> = (0..records.len()).collect();

    let row_keys: Vec<Vec<String>> = if pivot.row_fields.is_empty() {
        Vec::new()
    } else {
        let idxs: Vec<usize> = pivot
            .row_fields
            .iter()
            .map(|f| column_index(&col_names, &f.column))
            .collect::<Result<_, _>>()?;
        records
            .iter()
            .map(|rec| idxs.iter().map(|&i| group_key(&rec[i])).collect())
            .collect()
    };
    let col_keys: Vec<Vec<String>> = if pivot.col_fields.is_empty() {
        Vec::new()
    } else {
        let idxs: Vec<usize> = pivot
            .col_fields
            .iter()
            .map(|f| column_index(&col_names, &f.column))
            .collect::<Result<_, _>>()?;
        records
            .iter()
            .map(|rec| idxs.iter().map(|&i| group_key(&rec[i])).collect())
            .collect()
    };

    let row_groups = build_axis(
        &record_indices,
        &row_keys,
        &pivot.row_fields,
        pivot.grand_totals_row,
    );
    let col_groups = build_axis(
        &record_indices,
        &col_keys,
        &pivot.col_fields,
        pivot.grand_totals_col,
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

    // --- Header rows ---
    let n_col_header_rows = pivot.col_fields.len().max(1);
    let n_header_rows = if value_multiplier > 1 {
        n_col_header_rows + 1
    } else {
        n_col_header_rows
    };
    let row_label_width = pivot.row_fields.len().max(1);

    let mut header_rows: Vec<Vec<String>> = Vec::new();
    for r in 0..n_header_rows {
        let mut row: Vec<String> = Vec::new();
        for i in 0..row_label_width {
            if r == 0 {
                row.push(
                    pivot
                        .row_fields
                        .get(i)
                        .map(|f| f.column.clone())
                        .unwrap_or_else(|| "Total".to_string()),
                );
            } else {
                row.push(String::new());
            }
        }
        for group in &col_groups {
            for vf in 0..value_multiplier {
                let label = if r < pivot.col_fields.len() {
                    if group.is_grand_total {
                        if r == 0 {
                            "Grand Total".to_string()
                        } else {
                            String::new()
                        }
                    } else {
                        group
                            .labels
                            .get(r)
                            .and_then(|l| l.clone())
                            .unwrap_or_else(|| {
                                if group.is_subtotal && r == pivot.col_fields.len() - 1 {
                                    // shouldn't happen (subtotal never at innermost depth)
                                    String::new()
                                } else {
                                    String::new()
                                }
                            })
                    }
                } else if value_multiplier > 1 {
                    pivot.value_fields[vf].label()
                } else {
                    pivot
                        .value_fields
                        .first()
                        .map(|v| v.label())
                        .unwrap_or_default()
                };
                row.push(label);
            }
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
        && let Some(vf) = pivot.value_fields.first()
    {
        *cell = vf.label();
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
            if pivot.row_fields.is_empty() {
                display_labels[0] = "Total".to_string();
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
        header_rows,
        body_rows,
        width,
        row_axis: to_axis_items(&row_groups),
        col_axis: to_axis_items(&col_groups),
    })
}

/// Returns the distinct values of `values`, sorted the same way pivot
/// groups are (ascending numeric if every value parses as a number,
/// otherwise case-insensitive ascending text) -- used by the xlsx exporter
/// to build a pivot field's flat `<items>` enumeration.
pub(crate) fn sorted_distinct_strings(values: &[String]) -> Vec<String> {
    let mut seen: Vec<String> = Vec::new();
    for v in values {
        if !seen.contains(v) {
            seen.push(v.clone());
        }
    }
    let mut pairs: Vec<(String, Vec<usize>)> = seen.into_iter().map(|s| (s, Vec::new())).collect();
    sort_group_entries(&mut pairs);
    pairs.into_iter().map(|(s, _)| s).collect()
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
        let grid = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap();

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
    fn test_row_and_col_fields_with_subtotals() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        pivot.col_fields = vec![PivotField::new("Rep")];
        let grid = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap();

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
        // Regression test: the second (innermost) row field's own labels
        // must survive being nested under the first field's groups, not be
        // truncated away when the group tree is flattened.
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Region"), PivotField::new("Product")];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap();

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
        let grid = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap();
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
        }];
        pivot.grand_totals_row = false;
        let grid = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap();
        // East widgets: 10+20=30, West widgets: 30+25=55
        assert_eq!(grid.body_rows.len(), 2);
        assert_eq!(value_at(&grid.body_rows[0], 0), 30.0);
        assert_eq!(value_at(&grid.body_rows[1], 0), 55.0);
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
        let grid = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap();
        assert_eq!(grid.header_rows.last().unwrap()[1], "Sum of Amount");
        assert_eq!(grid.header_rows.last().unwrap()[2], "Count of Amount");
        assert_eq!(grid.body_rows[0].values.len(), 2);
        assert_eq!(value_at(&grid.body_rows[0], 0), 50.0); // Sum for East
        assert_eq!(value_at(&grid.body_rows[0], 1), 4.0); // Count for East
    }

    #[test]
    fn test_missing_column_errors() {
        let sheet = source_sheet();
        let mut pivot = base_pivot();
        pivot.row_fields = vec![PivotField::new("Nope")];
        let err = compute_pivot(std::slice::from_ref(&sheet), &pivot).unwrap_err();
        assert!(err.contains("not found"));
    }
}
