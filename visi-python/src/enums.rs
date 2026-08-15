//! `&str` -> enum parsers.
//!
//! The spellings are clap's, so the fuzz harness's existing config strings
//! (`fuzz_chart.py`'s `CHART_TYPES`, `fuzz_pivot.py`'s `AGGREGATIONS`) work
//! unchanged whether they are passed to the CLI or to these bindings. That
//! 1:1 correspondence is what `fuzz/test_backend_parity.py` relies on, so
//! adding a spelling here that the CLI does not accept would quietly break
//! the equivalence it checks.
//!
//! The CLI's own mapping lives in the `visi` crate (`ChartTypeArg`,
//! `PivotAreaArg`, `PivotAggArg`), which this crate deliberately does not
//! depend on -- hence the duplication. The tests below pin the spellings.

use crate::errors::invalid_argument;
use pyo3::PyResult;
use visi_engine::core::chart::ChartType;
use visi_engine::core::{PivotAggregation, PivotArea, VbaModuleKind};

/// Parses a chart type name (`"column"`, `"bar"`, `"line"`, ...).
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

/// Parses a pivot area name (`"row"`, `"column"`, `"value"`, `"filter"`).
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

/// Parses a pivot aggregation name (`"sum"`, `"count-numbers"`, ...).
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

/// Parses a VBA module kind (`"standard"`, `"class"`, `"document"`).
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

    // These lists are the fuzz harness's own (fuzz_chart.py:75,
    // fuzz_pivot.py:93). If a spelling here stops matching, the bindings and
    // the CLI diverge for inputs the fuzzer generates every run.
    #[test]
    fn accepts_every_spelling_the_fuzz_harness_emits() {
        for s in ["column", "bar", "line", "pie", "scatter", "area"] {
            assert!(parse_chart_type(s).is_ok(), "chart type {s:?}");
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
        assert!(matches!(parse_pivot_area("Row"), Ok(PivotArea::Row)));
        assert!(matches!(
            parse_pivot_agg("Count-Numbers"),
            Ok(PivotAggregation::CountNumbers)
        ));
    }

    #[test]
    fn rejects_unknown_values() {
        assert!(parse_chart_type("doughnut").is_err());
        assert!(parse_pivot_area("page").is_err());
        // Excel's own name for it, but not the spelling the CLI takes.
        assert!(parse_pivot_agg("counta").is_err());
        // Underscores are not an accepted alias -- clap renders kebab-case.
        assert!(parse_pivot_agg("count_numbers").is_err());
        // The file extensions, not the kinds -- the CLI takes neither.
        assert!(parse_vba_module_kind("bas").is_err());
        assert!(parse_vba_module_kind("cls").is_err());
    }

    /// Same exhaustiveness guard as `every_chart_type_has_a_spelling`.
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

    // A guard on exhaustiveness: if `ChartType` gains a variant, this match
    // stops compiling, which is the signal to add a spelling above.
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
