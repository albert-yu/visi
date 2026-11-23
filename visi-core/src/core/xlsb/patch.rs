use std::collections::{HashMap, HashSet};
use std::io::{Seek, Write};

use crate::core::engine::{CellType, ResultData, Sheet};
use crate::core::opc::{
    OpcPackage, PartDiff, REL_SHARED_STRINGS, REL_STYLES, REL_WORKSHEET, Relationship, next_rid,
    rel_type_matches,
};
use crate::core::style::CellStyle;

use super::package::workbook_sheets;
use super::shared_strings::SharedStrings;

const CT_XLSB_WORKSHEET: &str = "application/vnd.ms-excel.worksheet";
const CT_XLSB_SHARED_STRINGS: &str = "application/vnd.ms-excel.sharedStrings";
const CT_XLSB_STYLES: &str = "application/vnd.ms-excel.styles";

#[derive(Debug, Clone)]
struct XlsbCellSnapshot {
    src: String,
    res: ResultData,
    cell_type: CellType,
    style: Option<CellStyle>,
}

impl PartialEq for XlsbCellSnapshot {
    fn eq(&self, other: &Self) -> bool {
        self.src == other.src
            && self.cell_type == other.cell_type
            && self.style == other.style
            && result_data_eq(&self.res, &other.res)
    }
}

fn result_data_eq(a: &ResultData, b: &ResultData) -> bool {
    match (a, b) {
        (ResultData::None, ResultData::None) => true,
        (ResultData::Integer(x), ResultData::Integer(y)) => x == y,
        (ResultData::Float(x), ResultData::Float(y)) => x.to_bits() == y.to_bits(),
        (ResultData::String(x), ResultData::String(y)) => x == y,
        (ResultData::Boolean(x), ResultData::Boolean(y)) => x == y,
        (ResultData::Error(x), ResultData::Error(y)) => x == y,
        (ResultData::List(x), ResultData::List(y)) => {
            x.len() == y.len() && x.iter().zip(y.iter()).all(|(l, r)| result_data_eq(l, r))
        }
        _ => false,
    }
}

fn snapshot_cells(sheet: &Sheet) -> HashMap<(usize, usize), XlsbCellSnapshot> {
    let mut map = HashMap::new();
    for (col_idx, column) in sheet.columns.iter().enumerate() {
        for row in 0..column.len() {
            let src = column.src(row).unwrap_or_default();
            let res = column.value(row).unwrap_or(ResultData::None);
            let cell_type = column.cell_type(row).unwrap_or(CellType::Empty);
            let style = column.styles.get(row).and_then(|s| s.clone());
            let occupied = !src.is_empty() || !matches!(res, ResultData::None) || style.is_some();
            if occupied {
                map.insert(
                    (row, col_idx),
                    XlsbCellSnapshot {
                        src: src.to_string(),
                        res,
                        cell_type,
                        style,
                    },
                );
            }
        }
    }
    map
}

fn snapshot_col_widths(sheet: &Sheet) -> HashMap<usize, u64> {
    sheet
        .columns
        .iter()
        .enumerate()
        .filter_map(|(i, col)| col.width.map(|w| (i, w.to_bits())))
        .collect()
}

fn snapshot_row_heights(sheet: &Sheet) -> HashMap<usize, u64> {
    sheet
        .row_heights
        .iter()
        .enumerate()
        .filter_map(|(i, h)| h.map(|v| (i, v.to_bits())))
        .collect()
}

#[derive(Debug, Clone)]
struct XlsbSheetSnapshot {
    sheet_id: u64,
    name: String,
    part_path: String,
    rid: String,
    cells: HashMap<(usize, usize), XlsbCellSnapshot>,
    column_widths: HashMap<usize, u64>,
    row_heights: HashMap<usize, u64>,
}

impl XlsbSheetSnapshot {
    fn from_sheet(sheet: &Sheet, part_path: String, rid: String) -> Self {
        Self {
            sheet_id: sheet.id,
            name: sheet.name.clone(),
            part_path,
            rid,
            cells: snapshot_cells(sheet),
            column_widths: snapshot_col_widths(sheet),
            row_heights: snapshot_row_heights(sheet),
        }
    }
}

fn allocate_free_numbered_part(pkg: &OpcPackage, prefix: &str, suffix: &str) -> String {
    let mut n = 1usize;
    loop {
        let candidate = format!("{prefix}{n}{suffix}");
        if !pkg.has_part(&candidate) {
            return candidate;
        }
        n += 1;
    }
}

#[derive(Debug, Clone)]
pub(crate) struct XlsbSourceState {
    package: OpcPackage,
    workbook_part: String,
    shared_strings_part: String,
    shared_strings: SharedStrings,
    sheet_snapshots: Vec<XlsbSheetSnapshot>,
}

impl XlsbSourceState {
    pub(crate) fn capture(package: OpcPackage, sheets: &[Sheet]) -> Result<Self, String> {
        let workbook_part = "xl/workbook.bin".to_string();
        let wb_bytes = package
            .get_bytes(&workbook_part)
            .ok_or_else(|| "XLSB package is missing xl/workbook.bin".to_string())?;
        let wb_sheets = workbook_sheets(wb_bytes)?;
        let wb_rels = package.read_relationships(&workbook_part);

        let shared_strings_part = wb_rels
            .iter()
            .find(|r| !r.is_external() && rel_type_matches(&r.rel_type, "sharedStrings"))
            .map(|r| r.resolved_target(&workbook_part))
            .unwrap_or_else(|| "xl/sharedStrings.bin".to_string());

        let shared_strings = match package.get_bytes(&shared_strings_part) {
            Some(bytes) => SharedStrings::parse(bytes)?,
            None => SharedStrings::default(),
        };

        let mut sheet_snapshots = Vec::with_capacity(sheets.len());
        for (sheet, (_, rid)) in sheets.iter().zip(wb_sheets.iter()) {
            let part_path = wb_rels
                .iter()
                .find(|r| r.id == *rid)
                .map(|r| r.resolved_target(&workbook_part))
                .ok_or_else(|| format!("missing XLSB worksheet relationship '{rid}'"))?;
            sheet_snapshots.push(XlsbSheetSnapshot::from_sheet(sheet, part_path, rid.clone()));
        }

        if sheet_snapshots.len() != sheets.len() {
            return Err("XLSB worksheet count mismatch during baseline capture".to_string());
        }

        Ok(Self {
            package,
            workbook_part,
            shared_strings_part,
            shared_strings,
            sheet_snapshots,
        })
    }

    pub(crate) fn can_patch_surgically(&self, sheets: &[Sheet]) -> bool {
        !sheets.is_empty()
    }

    pub(crate) fn has_pending_changes(&self, sheets: &[Sheet]) -> bool {
        if sheets.len() != self.sheet_snapshots.len() {
            return true;
        }
        for (sheet, snap) in sheets.iter().zip(self.sheet_snapshots.iter()) {
            if sheet.id != snap.sheet_id
                || sheet.name != snap.name
                || snapshot_cells(sheet) != snap.cells
                || snapshot_col_widths(sheet) != snap.column_widths
                || snapshot_row_heights(sheet) != snap.row_heights
            {
                return true;
            }
        }
        false
    }

    pub(crate) fn write_to<W: Write + Seek>(&self, writer: W) -> Result<W, String> {
        self.package.write_to(writer)
    }

    pub(crate) fn apply_patch(&mut self, sheets: &[Sheet]) -> Result<PartDiff, String> {
        if !self.has_pending_changes(sheets) {
            return Ok(PartDiff::default());
        }
        let (pkg, shared_strings, new_snapshots, diff) = self.build_patched_package(sheets)?;
        self.package = pkg;
        self.shared_strings = shared_strings;
        self.sheet_snapshots = new_snapshots;
        Ok(diff)
    }

    pub(crate) fn patch_to_bytes(&self, sheets: &[Sheet]) -> Result<Vec<u8>, String> {
        let (pkg, _, _, _) = self.build_patched_package(sheets)?;
        pkg.to_bytes()
    }

    fn build_patched_package(
        &self,
        sheets: &[Sheet],
    ) -> Result<(OpcPackage, SharedStrings, Vec<XlsbSheetSnapshot>, PartDiff), String> {
        if sheets.is_empty() {
            return Err("cannot write an XLSB workbook without worksheets".to_string());
        }

        let mut pkg = self.package.clone();
        pkg.start_tracking();

        let mut shared_strings = self.shared_strings.clone();
        shared_strings.dirty = false;

        let snap_by_id: HashMap<u64, &XlsbSheetSnapshot> = self
            .sheet_snapshots
            .iter()
            .map(|s| (s.sheet_id, s))
            .collect();
        let active_ids: HashSet<u64> = sheets.iter().map(|s| s.id).collect();
        let topology_changed = sheets.len() != self.sheet_snapshots.len()
            || sheets
                .iter()
                .zip(self.sheet_snapshots.iter())
                .any(|(s, snap)| s.id != snap.sheet_id);

        let mut wb_rels = pkg.read_relationships(&self.workbook_part);
        let mut wb_rels_changed = false;
        let mut ct = pkg.read_content_types();
        let mut ct_changed = false;

        let mut added_sheet_info: HashMap<u64, (String, String)> = HashMap::new();
        if topology_changed {
            for snap in &self.sheet_snapshots {
                if !active_ids.contains(&snap.sheet_id) {
                    pkg.remove_part(&snap.part_path);
                    ct.remove_override(&snap.part_path);
                    ct_changed = true;
                    let before = wb_rels.len();
                    wb_rels.retain(|r| r.id != snap.rid);
                    if wb_rels.len() != before {
                        wb_rels_changed = true;
                    }
                }
            }

            for sheet in sheets {
                if !snap_by_id.contains_key(&sheet.id) {
                    let new_part = allocate_free_numbered_part(&pkg, "xl/worksheets/sheet", ".bin");
                    let rid = next_rid(&wb_rels);
                    let target = new_part.trim_start_matches("xl/").to_string();
                    wb_rels.push(Relationship {
                        id: rid.clone(),
                        rel_type: REL_WORKSHEET.to_string(),
                        target,
                        target_mode: None,
                    });
                    wb_rels_changed = true;
                    ct.set_override(&new_part, CT_XLSB_WORKSHEET);
                    ct_changed = true;
                    added_sheet_info.insert(sheet.id, (rid, new_part));
                }
            }
        }

        let mut new_snapshots = Vec::with_capacity(sheets.len());
        let mut names_and_rids = Vec::with_capacity(sheets.len());
        let mut any_renamed = false;

        for sheet in sheets {
            if let Some(snap) = snap_by_id.get(&sheet.id) {
                if sheet.name != snap.name {
                    any_renamed = true;
                }
                let cur_cells = snapshot_cells(sheet);
                let cur_widths = snapshot_col_widths(sheet);
                let cur_heights = snapshot_row_heights(sheet);
                let sheet_dirty = cur_cells != snap.cells
                    || cur_widths != snap.column_widths
                    || cur_heights != snap.row_heights;
                if sheet_dirty {
                    let ws_bytes = super::worksheet::build(sheet, &mut shared_strings)?;
                    pkg.set_part(&snap.part_path, ws_bytes);
                }
                names_and_rids.push((sheet.name.clone(), snap.rid.clone()));
                new_snapshots.push(XlsbSheetSnapshot {
                    sheet_id: sheet.id,
                    name: sheet.name.clone(),
                    part_path: snap.part_path.clone(),
                    rid: snap.rid.clone(),
                    cells: cur_cells,
                    column_widths: cur_widths,
                    row_heights: cur_heights,
                });
            } else if let Some((rid, new_part)) = added_sheet_info.get(&sheet.id) {
                let ws_bytes = super::worksheet::build(sheet, &mut shared_strings)?;
                pkg.set_part(new_part, ws_bytes);
                names_and_rids.push((sheet.name.clone(), rid.clone()));
                new_snapshots.push(XlsbSheetSnapshot::from_sheet(
                    sheet,
                    new_part.clone(),
                    rid.clone(),
                ));
            }
        }

        if topology_changed || any_renamed {
            let wb_bytes = super::workbook::build_with_rids(&names_and_rids);
            pkg.set_part(&self.workbook_part, wb_bytes);
        }

        if shared_strings.dirty {
            let had_sst = pkg.has_part(&self.shared_strings_part);
            pkg.set_part(&self.shared_strings_part, shared_strings.encode());
            if !had_sst {
                ct.set_override(&self.shared_strings_part, CT_XLSB_SHARED_STRINGS);
                ct_changed = true;
                if !wb_rels
                    .iter()
                    .any(|r| rel_type_matches(&r.rel_type, "sharedStrings"))
                {
                    let rid = next_rid(&wb_rels);
                    let target = self
                        .shared_strings_part
                        .trim_start_matches("xl/")
                        .to_string();
                    wb_rels.push(Relationship {
                        id: rid,
                        rel_type: REL_SHARED_STRINGS.to_string(),
                        target,
                        target_mode: None,
                    });
                    wb_rels_changed = true;
                }
            }
            shared_strings.dirty = false;
        }

        if !pkg.has_part("xl/styles.bin") {
            pkg.set_part("xl/styles.bin", super::styles::build());
            ct.set_override("xl/styles.bin", CT_XLSB_STYLES);
            ct_changed = true;
            if !wb_rels
                .iter()
                .any(|r| rel_type_matches(&r.rel_type, "styles"))
            {
                let rid = next_rid(&wb_rels);
                wb_rels.push(Relationship {
                    id: rid,
                    rel_type: REL_STYLES.to_string(),
                    target: "styles.bin".to_string(),
                    target_mode: None,
                });
                wb_rels_changed = true;
            }
        }

        if wb_rels_changed {
            pkg.write_relationships(&self.workbook_part, &wb_rels);
        }
        if ct_changed {
            pkg.write_content_types(&ct);
        }

        let diff = pkg.finish_tracking();
        Ok((pkg, shared_strings, new_snapshots, diff))
    }
}
