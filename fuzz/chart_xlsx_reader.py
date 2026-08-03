#!/usr/bin/env python3
"""
Reads chart objects back out of an .xlsx file via `openpyxl` and normalizes
them into plain dicts comparable against libvisi's `Chart` struct fields
(`chart_type`, `data_range`, `title`, `xlabel`, `ylabel`, `show_legend`).

This is an independent-reader fidelity check: `openpyxl` implements OOXML
chart parsing from scratch, so it can catch chart-XML mistakes that visi's
own `parse_chart_xml` (used by the existing Rust round-trip tests) would
never notice, since a bug shared by both visi's writer and its own reader
would still "round-trip" successfully in a Rust-only test.

Relies on `worksheet._charts`, an underscore-prefixed (non-public) openpyxl
attribute -- the only way openpyxl exposes charts it has just read back from
a file. There is no supported public API for this as of openpyxl 3.1.x.

Usable standalone:
    python3 fuzz/chart_xlsx_reader.py some.xlsx
or as an importable module:
    from chart_xlsx_reader import read_charts
"""

import sys
import json

# openpyxl represents both Bar and Column charts as `BarChart` -- the
# distinction lives in `BarChart.type` ("col" vs "bar"), mirroring visi's own
# `<c:barDir val="col"|"bar">` disambiguation in `parse_chart_xml`.
_CLASS_TO_TYPE = {
    "LineChart": "Line",
    "PieChart": "Pie",
    "ScatterChart": "Scatter",
    "AreaChart": "Area",
}


def _title_text(title_obj):
    """Walks openpyxl's nested Title -> Text -> RichText -> Paragraph ->
    RegularTextRun structure down to a flat string, or None if there's no
    title (or it's a strRef-based title, which visi never writes). openpyxl
    exposes no flat-string accessor for *reading* a title, only for writing
    one via `Title.tx.rich` construction helpers -- so this must walk the
    object graph by hand. Concatenates every run's text within the first
    paragraph; visi only ever writes a single run."""
    if title_obj is None:
        return None
    tx = getattr(title_obj, "tx", None)
    if tx is None:
        return None
    rich = getattr(tx, "rich", None)
    if rich is None or not rich.p:
        return None
    first_para = rich.p[0]
    if not first_para.r:
        return None
    text = "".join(run.t or "" for run in first_para.r)
    return text if text else None


def _chart_type(chart):
    cls = type(chart).__name__
    if cls == "BarChart":
        return "Column" if chart.type == "col" else "Bar"
    return _CLASS_TO_TYPE.get(cls, cls)


def _series_range(ref):
    """`ref` is a Series' `.val` (NumDataSource) or `.cat` (AxDataSource).
    Returns the referenced range formula (`numRef.f` or `strRef.f`), or None
    if there's no series or no reference (e.g. literal/inline data)."""
    if ref is None:
        return None
    num_ref = getattr(ref, "numRef", None)
    if num_ref is not None:
        return num_ref.f
    str_ref = getattr(ref, "strRef", None)
    if str_ref is not None:
        return str_ref.f
    return None


def read_charts(xlsx_path):
    """Returns one dict per chart found across all worksheets:
    {sheet, chart_type, cat_range, val_range, title, xlabel, ylabel,
    show_legend}.

    Deliberately single-series only (reads `chart.series[0]`), matching the
    single-range model of libvisi's `Chart` struct -- multi-series charts
    are out of scope for both the engine and this fuzzing harness.
    """
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path)
    out = []
    for ws in wb.worksheets:
        for chart in ws._charts:
            series = chart.series[0] if chart.series else None
            # ScatterChart series expose `.xVal`/`.yVal` instead of the
            # `.cat`/`.val` every other chart type uses -- fall back to
            # those so scatter charts still report a category/value range.
            cat_range = None
            val_range = None
            if series is not None:
                cat_range = _series_range(getattr(series, "cat", None)) or _series_range(
                    getattr(series, "xVal", None)
                )
                val_range = _series_range(getattr(series, "val", None)) or _series_range(
                    getattr(series, "yVal", None)
                )
            # PieChart (and some others) have no x_axis/y_axis at all --
            # unlike a plain attribute holding None, accessing the name
            # itself raises AttributeError, so this must use getattr.
            x_axis = getattr(chart, "x_axis", None)
            y_axis = getattr(chart, "y_axis", None)
            out.append(
                {
                    "sheet": ws.title,
                    "chart_type": _chart_type(chart),
                    "cat_range": cat_range,
                    "val_range": val_range,
                    "title": _title_text(chart.title),
                    "xlabel": _title_text(x_axis.title) if x_axis is not None else None,
                    "ylabel": _title_text(y_axis.title) if y_axis is not None else None,
                    "show_legend": chart.legend is not None,
                }
            )
    return out


def main():
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <path-to-xlsx>", file=sys.stderr)
        sys.exit(1)
    charts = read_charts(sys.argv[1])
    print(json.dumps(charts, indent=2))


if __name__ == "__main__":
    main()
