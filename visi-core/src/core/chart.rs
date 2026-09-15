use serde::{Deserialize, Serialize};

/// [LLM-generated] The chart shapes visi can read and write.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub enum ChartType {
    /// [LLM-generated] Vertical bars.
    Column,
    /// [LLM-generated] Horizontal bars.
    Bar,
    /// [LLM-generated] Points joined by lines.
    Line,
    /// [LLM-generated] A pie, from a single series.
    Pie,
    /// [LLM-generated] X/Y points, taking the first column as the x values.
    Scatter,
    /// [LLM-generated] A line chart with the area beneath it filled.
    Area,
}

/// [LLM-generated] A chart over a range of cells.
///
/// Workbook-level rather than sheet-scoped: `WorkbookManager::charts` owns
/// these, and which worksheet a chart is drawn on comes from `data_range`'s
/// sheet prefix.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct Chart {
    /// [LLM-generated] Workbook-unique identifier, stable across renames.
    pub id: u64,
    /// [LLM-generated] Display name.
    pub name: String,
    /// [LLM-generated] What shape to draw.
    pub chart_type: ChartType,
    /// [LLM-generated] The source cells, in A1 notation with a sheet prefix
    /// (`"table_1!A1:B10"`). The prefix also decides which worksheet the
    /// chart is anchored to.
    pub data_range: String,
    /// [LLM-generated] Chart title, if it has one.
    pub title: Option<String>,
    /// [LLM-generated] X-axis label, if it has one.
    pub xlabel: Option<String>,
    /// [LLM-generated] Y-axis label, if it has one.
    pub ylabel: Option<String>,
    /// [LLM-generated] Whether to draw a legend.
    pub show_legend: bool,
    /// [LLM-generated] 0-based row/col of the cell the chart's top-left corner is anchored
    /// to on its worksheet (which worksheet that is comes from
    /// `data_range`'s sheet/table prefix, same as today).
    #[serde(default)]
    pub anchor_row: usize,
    /// [LLM-generated] Column half of that anchor; see [`Chart::anchor_row`].
    #[serde(default)]
    pub anchor_col: usize,
}
