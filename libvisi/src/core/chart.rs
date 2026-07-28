use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub enum ChartType {
    Column,
    Bar,
    Line,
    Pie,
    Scatter,
    Area,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct ChartAnchor {
    pub col: isize,
    pub row: isize,
    pub col_offset_pixels: f32,
    pub row_offset_pixels: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct Chart {
    pub id: u64,
    pub name: String,
    pub chart_type: ChartType,
    pub data_range: String, // e.g. "table_1!A1:B10"
    pub from_anchor: ChartAnchor,
    pub to_anchor: ChartAnchor,
    pub title: Option<String>,
    pub xlabel: Option<String>,
    pub ylabel: Option<String>,
    pub show_legend: bool,
    pub z_index: i32,
}
