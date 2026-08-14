//! Chart definitions.

use serde::{Deserialize, Serialize};

/// The chart shapes visi can read and write.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub enum ChartType {
    /// Vertical bars.
    Column,
    /// Horizontal bars.
    Bar,
    /// Points joined by lines.
    Line,
    /// A pie, from a single series.
    Pie,
    /// X/Y points, taking the first column as the x values.
    Scatter,
    /// A line chart with the area beneath it filled.
    Area,
}

/// A chart over a range of cells.
///
/// Workbook-level rather than sheet-scoped: `WorkbookManager::charts` owns
/// these, and which worksheet a chart is drawn on comes from `data_range`'s
/// sheet prefix.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct Chart {
    /// Workbook-unique identifier, stable across renames.
    pub id: u64,
    /// Display name.
    pub name: String,
    /// What shape to draw.
    pub chart_type: ChartType,
    /// The source cells, in A1 notation with a sheet prefix
    /// (`"table_1!A1:B10"`). The prefix also decides which worksheet the
    /// chart is anchored to.
    pub data_range: String,
    /// Chart title, if it has one.
    pub title: Option<String>,
    /// X-axis label, if it has one.
    pub xlabel: Option<String>,
    /// Y-axis label, if it has one.
    pub ylabel: Option<String>,
    /// Whether to draw a legend.
    pub show_legend: bool,
    /// 0-based row/col of the cell the chart's top-left corner is anchored
    /// to on its worksheet (which worksheet that is comes from
    /// `data_range`'s sheet/table prefix, same as today).
    #[serde(default)]
    pub anchor_row: usize,
    /// Column half of that anchor; see [`Chart::anchor_row`].
    #[serde(default)]
    pub anchor_col: usize,
}
