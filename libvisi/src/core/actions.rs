use serde::{Deserialize, Serialize};

use crate::render::Position;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum SheetAction {
    SetCellSrc {
        sheet_name: String,
        col: usize,
        row: usize,
        src: String,
    },
    UpdateColName {
        sheet_name: String,
        col: usize,
        name: String,
    },
    UpdateTableName {
        old_name: String,
        new_name: String,
    },
    InsertRow {
        sheet_name: String,
        index: usize,
    },
    DeleteRow {
        sheet_name: String,
        index: usize,
    },
    InsertCol {
        sheet_name: String,
        index: usize,
    },
    DeleteCol {
        sheet_name: String,
        index: usize,
    },
    MoveTable {
        sheet_name: String,
        pos: Position,
    },
    AddTable {
        sheet: crate::core::Sheet,
    },
    DeleteTable {
        sheet_name: String,
    },
    SetTableZIndex {
        sheet_name: String,
        z_index: i32,
    },
    AddChart {
        chart: crate::core::chart::Chart,
    },
    DeleteChart {
        chart_id: u64,
    },
    UpdateChartLayout {
        chart_id: u64,
        from_anchor: crate::core::chart::ChartAnchor,
        to_anchor: crate::core::chart::ChartAnchor,
    },
    UpdateChartProperties {
        chart_id: u64,
        chart_type: crate::core::chart::ChartType,
        data_range: String,
        title: Option<String>,
        xlabel: Option<String>,
        ylabel: Option<String>,
        show_legend: bool,
    },
}
