use crate::errors::invalid_argument;
use pyo3::PyResult;
use visi_engine::core::chart::ChartType;
use visi_engine::core::{CellType, PivotAggregation, PivotArea, VbaModuleKind};

pub fn parse_chart_type(s: &str) -> PyResult<ChartType> {
    match s.to_ascii_lowercase().as_str() {
        "column" => Ok(ChartType::Column),
        "bar" => Ok(ChartType::Bar),
        "line" => Ok(ChartType::Line),
        "pie" => Ok(ChartType::Pie),
        "scatter" => Ok(ChartType::Scatter),
        "area" => Ok(ChartType::Area),
        other => Err(invalid_argument(format!(
            "unknown chart type {other:?}; expected one of: column, bar, line, pie, scatter, area"
        ))),
    }
}

pub fn parse_cell_type(s: &str) -> PyResult<CellType> {
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
        other => Err(invalid_argument(format!(
            "unknown cell type {other:?}; expected one of: empty, int, float, string, bool, date-time, date-time-iso, duration-iso, error"
        ))),
    }
}

pub fn parse_pivot_area(s: &str) -> PyResult<PivotArea> {
    match s.to_ascii_lowercase().as_str() {
        "row" => Ok(PivotArea::Row),
        "column" => Ok(PivotArea::Column),
        "value" => Ok(PivotArea::Value),
        "filter" => Ok(PivotArea::Filter),
        other => Err(invalid_argument(format!(
            "unknown pivot area {other:?}; expected one of: row, column, value, filter"
        ))),
    }
}

pub fn parse_pivot_agg(s: &str) -> PyResult<PivotAggregation> {
    match s.to_ascii_lowercase().as_str() {
        "sum" => Ok(PivotAggregation::Sum),
        "count" => Ok(PivotAggregation::Count),
        "count-numbers" => Ok(PivotAggregation::CountNumbers),
        "average" => Ok(PivotAggregation::Average),
        "max" => Ok(PivotAggregation::Max),
        "min" => Ok(PivotAggregation::Min),
        other => Err(invalid_argument(format!(
            "unknown aggregation {other:?}; expected one of: sum, count, count-numbers, average, max, min"
        ))),
    }
}

pub fn parse_vba_module_kind(s: &str) -> PyResult<VbaModuleKind> {
    match s.to_ascii_lowercase().as_str() {
        "standard" => Ok(VbaModuleKind::Standard),
        "class" => Ok(VbaModuleKind::Class),
        "document" => Ok(VbaModuleKind::Document),
        other => Err(invalid_argument(format!(
            "unknown VBA module kind {other:?}; expected one of: standard, class, document"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_every_spelling_the_fuzz_harness_emits() {
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
        assert!(matches!(parse_chart_type("COLUMN"), Ok(ChartType::Column)));
        assert!(matches!(parse_cell_type("STRING"), Ok(CellType::String)));
        assert!(matches!(parse_pivot_area("Row"), Ok(PivotArea::Row)));
        assert!(matches!(
            parse_pivot_agg("Count-Numbers"),
            Ok(PivotAggregation::CountNumbers)
        ));
    }

    #[test]
    fn rejects_unknown_values() {
        assert!(parse_chart_type("doughnut").is_err());
        assert!(parse_cell_type("unknown").is_err());
        assert!(parse_pivot_area("page").is_err());
        assert!(parse_pivot_agg("counta").is_err());
        assert!(parse_pivot_agg("count_numbers").is_err());
        assert!(parse_vba_module_kind("bas").is_err());
        assert!(parse_vba_module_kind("cls").is_err());
    }

    #[test]
    fn every_vba_module_kind_has_a_spelling() {
        for kind in [
            VbaModuleKind::Standard,
            VbaModuleKind::Class,
            VbaModuleKind::Document,
        ] {
            let name = match kind {
                VbaModuleKind::Standard => "standard",
                VbaModuleKind::Class => "class",
                VbaModuleKind::Document => "document",
            };
            assert_eq!(parse_vba_module_kind(name).unwrap(), kind);
        }
    }

    #[test]
    fn every_cell_type_has_a_spelling() {
        for ct in [
            CellType::Empty,
            CellType::Int,
            CellType::Float,
            CellType::String,
            CellType::Bool,
            CellType::DateTime,
            CellType::DateTimeIso,
            CellType::DurationIso,
            CellType::Error,
        ] {
            let name = match ct {
                CellType::Empty => "empty",
                CellType::Int => "int",
                CellType::Float => "float",
                CellType::String => "string",
                CellType::Bool => "bool",
                CellType::DateTime => "date-time",
                CellType::DateTimeIso => "date-time-iso",
                CellType::DurationIso => "duration-iso",
                CellType::Error => "error",
            };
            assert_eq!(parse_cell_type(name).unwrap(), ct);
        }
    }

    #[test]
    fn every_chart_type_has_a_spelling() {
        for ct in [
            ChartType::Column,
            ChartType::Bar,
            ChartType::Line,
            ChartType::Pie,
            ChartType::Scatter,
            ChartType::Area,
        ] {
            let name = match ct {
                ChartType::Column => "column",
                ChartType::Bar => "bar",
                ChartType::Line => "line",
                ChartType::Pie => "pie",
                ChartType::Scatter => "scatter",
                ChartType::Area => "area",
            };
            assert_eq!(parse_chart_type(name).unwrap(), ct);
        }
    }
}
