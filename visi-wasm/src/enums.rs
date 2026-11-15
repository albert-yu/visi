use visi_engine::core::chart::ChartType;
use visi_engine::core::{CellType, PivotAggregation, PivotArea, VbaModuleKind};
use visi_engine::{Error as CoreError, WorkbookFormat};

pub fn parse_workbook_format(s: &str) -> Result<WorkbookFormat, CoreError> {
    match s.trim().to_ascii_lowercase().as_str() {
        "xlsx" | "xlsm" => Ok(WorkbookFormat::Xlsx),
        "xlsb" => Ok(WorkbookFormat::Xlsb),
        other => Err(CoreError::InvalidArgument(format!(
            "unsupported workbook format '{other}'; expected xlsx, xlsm, or xlsb"
        ))),
    }
}

pub fn parse_chart_type(s: &str) -> Result<ChartType, CoreError> {
    match s.to_ascii_lowercase().as_str() {
        "column" => Ok(ChartType::Column),
        "bar" => Ok(ChartType::Bar),
        "line" => Ok(ChartType::Line),
        "pie" => Ok(ChartType::Pie),
        "scatter" => Ok(ChartType::Scatter),
        "area" => Ok(ChartType::Area),
        other => Err(CoreError::InvalidArgument(format!(
            "unknown chart type {other:?}; expected one of: column, bar, line, pie, scatter, area"
        ))),
    }
}

pub fn parse_cell_type(s: &str) -> Result<CellType, CoreError> {
    match s.to_ascii_lowercase().as_str() {
        "empty" => Ok(CellType::Empty),
        "int" => Ok(CellType::Int),
        "float" => Ok(CellType::Float),
        "string" => Ok(CellType::String),
        "bool" => Ok(CellType::Bool),
        "date-time" => Ok(CellType::DateTime),
        "date-time-iso" => Ok(CellType::DateTimeIso),
        "duration-iso" => Ok(CellType::DurationIso),
        "error" => Ok(CellType::Error),
        other => Err(CoreError::InvalidArgument(format!(
            "unknown cell type {other:?}; expected one of: empty, int, float, string, bool, date-time, date-time-iso, duration-iso, error"
        ))),
    }
}

pub fn parse_pivot_area(s: &str) -> Result<PivotArea, CoreError> {
    match s.to_ascii_lowercase().as_str() {
        "row" => Ok(PivotArea::Row),
        "column" => Ok(PivotArea::Column),
        "value" => Ok(PivotArea::Value),
        "filter" => Ok(PivotArea::Filter),
        other => Err(CoreError::InvalidArgument(format!(
            "unknown pivot area {other:?}; expected one of: row, column, value, filter"
        ))),
    }
}

pub fn parse_pivot_agg(s: &str) -> Result<PivotAggregation, CoreError> {
    match s.to_ascii_lowercase().as_str() {
        "sum" => Ok(PivotAggregation::Sum),
        "count" => Ok(PivotAggregation::Count),
        "count-numbers" => Ok(PivotAggregation::CountNumbers),
        "average" => Ok(PivotAggregation::Average),
        "max" => Ok(PivotAggregation::Max),
        "min" => Ok(PivotAggregation::Min),
        other => Err(CoreError::InvalidArgument(format!(
            "unknown aggregation {other:?}; expected one of: sum, count, count-numbers, average, max, min"
        ))),
    }
}

pub fn parse_vba_module_kind(s: &str) -> Result<VbaModuleKind, CoreError> {
    match s.to_ascii_lowercase().as_str() {
        "standard" => Ok(VbaModuleKind::Standard),
        "class" => Ok(VbaModuleKind::Class),
        "document" => Ok(VbaModuleKind::Document),
        other => Err(CoreError::InvalidArgument(format!(
            "unknown VBA module kind {other:?}; expected one of: standard, class, document"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_every_standard_spelling() {
        for s in ["xlsx", "xlsm", "xlsb"] {
            assert!(parse_workbook_format(s).is_ok(), "workbook format {s:?}");
        }
        for s in ["column", "bar", "line", "pie", "scatter", "area"] {
            assert!(parse_chart_type(s).is_ok(), "chart type {s:?}");
        }
        for s in [
            "empty",
            "int",
            "float",
            "string",
            "bool",
            "date-time",
            "date-time-iso",
            "duration-iso",
            "error",
        ] {
            assert!(parse_cell_type(s).is_ok(), "cell type {s:?}");
        }
        for s in ["row", "column", "value", "filter"] {
            assert!(parse_pivot_area(s).is_ok(), "pivot area {s:?}");
        }
        for s in ["sum", "count", "count-numbers", "average", "max", "min"] {
            assert!(parse_pivot_agg(s).is_ok(), "aggregation {s:?}");
        }
        for s in ["standard", "class", "document"] {
            assert!(parse_vba_module_kind(s).is_ok(), "module kind {s:?}");
        }
    }

    #[test]
    fn is_case_insensitive() {
        assert_eq!(parse_workbook_format("XLSB").unwrap(), WorkbookFormat::Xlsb);
        assert_eq!(parse_chart_type("COLUMN").unwrap(), ChartType::Column);
        assert_eq!(parse_cell_type("STRING").unwrap(), CellType::String);
        assert_eq!(parse_pivot_area("Row").unwrap(), PivotArea::Row);
        assert_eq!(
            parse_pivot_agg("Count-Numbers").unwrap(),
            PivotAggregation::CountNumbers
        );
    }

    #[test]
    fn rejects_unknown_values() {
        assert!(parse_workbook_format("csv").is_err());
        assert!(parse_chart_type("doughnut").is_err());
        assert!(parse_cell_type("unknown").is_err());
        assert!(parse_pivot_area("page").is_err());
        assert!(parse_pivot_agg("counta").is_err());
        assert!(parse_pivot_agg("count_numbers").is_err());
        assert!(parse_vba_module_kind("bas").is_err());
        assert!(parse_vba_module_kind("cls").is_err());
    }
}
