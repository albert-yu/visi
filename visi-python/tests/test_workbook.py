#!/usr/bin/env python3
"""Unit tests for the `visi_core` extension module.

Run after building it into the venv:

    source fuzz/venv/bin/activate
    maturin develop -m visi-python/Cargo.toml --release
    pytest visi-python/tests/

These cover the translation layer only -- value conversion, the exception
hierarchy, and argument handling. Whether the bindings and the CLI agree about
what visi *does* is `fuzz/test_backend_parity.py`'s job.
"""

import pytest

visi_core = pytest.importorskip(
    "visi_core",
    reason="build it with `maturin develop -m visi-python/Cargo.toml --release`",
)


# ---------------------------------------------------------------- lifecycle


def test_new_empty_has_one_sheet():
    wb = visi_core.Workbook()
    assert len(wb.sheet_names) == 1


def test_evaluate_and_read_back():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "10")
    wb.set_cell(0, 1, "=A1*2")
    wb.evaluate()
    assert wb.get_cell(0, 0) == 10
    assert wb.get_cell(0, 1) == 20


def test_save_bytes_load_bytes_roundtrip(tmp_path):
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "42")
    wb.evaluate()

    again = visi_core.Workbook.load_bytes(wb.save_bytes())
    assert again.sheet_names == wb.sheet_names
    assert again.get_cell(0, 0) == 42


def test_roundtrip_matches_manual_save_load():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "7")
    wb.evaluate()
    assert wb.roundtrip().get_cell(0, 0) == visi_core.Workbook.load_bytes(
        wb.save_bytes()
    ).get_cell(0, 0)


def test_save_creates_parent_directories(tmp_path):
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "1")
    wb.evaluate()
    out = tmp_path / "nested" / "deeper" / "book.xlsx"
    wb.save(out)
    assert out.exists()


def test_eval_file(tmp_path):
    src, dst = tmp_path / "in.xlsx", tmp_path / "out.xlsx"
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "3")
    wb.set_cell(0, 1, "=A1+4")
    wb.save(src)

    visi_core.eval_file(src, dst)
    assert visi_core.Workbook.load(dst).get_cell(0, 1) == 7


# ------------------------------------------------------------ value mapping


@pytest.mark.parametrize(
    "src,expected,expected_type",
    [
        # A numeric *literal* stays an integer ...
        ("10", 10, int),
        # ... while arithmetic goes through floats, as Excel's model does.
        ("=1+1", 2.0, float),
        ("=1.5+1", 2.5, float),
        ('=CONCATENATE("a","b")', "ab", str),
        ("=TRUE()", True, bool),
    ],
)
def test_result_types(src, expected, expected_type):
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, src)
    wb.evaluate()
    v = wb.get_cell(0, 0)
    assert v == expected
    assert isinstance(v, expected_type)


def test_set_cell_with_explicit_type():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "12345", cell_type="string")
    wb.set_cell(0, 1, "12345")
    wb.evaluate()
    assert wb.get_cell_type(0, 0) == "string"
    assert wb.get_cell(0, 0) == "12345"
    assert isinstance(wb.get_cell(0, 0), str)
    assert wb.get_cell_type(0, 1) == "number"
    assert wb.get_cell(0, 1) == 12345

    again = wb.roundtrip()
    assert again.get_cell_type(0, 0) == "string"
    assert again.get_cell(0, 0) == "12345"


def test_set_cell_type_on_existing_cell():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "007")
    wb.set_cell_type(0, 0, "string")
    wb.evaluate()
    assert wb.get_cell_type(0, 0) == "string"
    assert wb.get_cell(0, 0) == "007"


def test_blank_cell_is_none():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "1")
    wb.evaluate()
    assert wb.get_cell(5, 5) is None


def test_error_value_is_cell_error():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "=1/0")
    wb.evaluate()
    v = wb.get_cell(0, 0)
    assert isinstance(v, visi_core.CellError)
    assert v.code == "#DIV/0!"
    assert str(v) == "#DIV/0!"
    assert v == "#DIV/0!"


def test_error_and_text_are_distinguished_by_type_not_equality():
    """The whole reason CellError exists.

    A cell can hold the *text* `#DIV/0!`. Both compare equal to the code
    string -- they are equal as values, which is also how the harness's
    XLSXEvaluatedReader has always compared them -- so the type, not `==`, is
    what tells them apart.
    """
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "=1/0")
    wb.set_cell(0, 1, '=CONCATENATE("#DIV/0!")')
    wb.evaluate()

    err, text = wb.get_cell(0, 0), wb.get_cell(0, 1)
    assert isinstance(err, visi_core.CellError)
    assert isinstance(text, str) and not isinstance(text, visi_core.CellError)
    assert err == text  # equal as values ...
    assert type(err) is not type(text)  # ... but never the same thing


def test_cell_error_hash_agrees_with_eq():
    e = visi_core.CellError("#N/A")
    assert hash(e) == hash("#N/A")
    assert {e: 1}["#N/A"] == 1


def test_get_display_renders_a_date():
    """A date is a serial number plus a format; only get_display renders it."""
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "6/22/26")
    wb.evaluate()
    assert isinstance(wb.get_cell(0, 0), (int, float))
    assert "26" in wb.get_display(0, 0)


def test_get_src_returns_the_formula_text():
    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "=1+1")
    wb.evaluate()
    assert wb.get_src(0, 0).startswith("=")
    assert wb.get_cell(0, 0) == 2


# --------------------------------------------------------------- exceptions


def test_not_found_carries_structured_payload():
    wb = visi_core.Workbook()
    with pytest.raises(visi_core.NotFoundError) as exc:
        wb.sheet_index("nope")
    e = exc.value
    assert isinstance(e, visi_core.VisiError)
    assert e.kind == "sheet"
    assert e.name == "nope"
    assert isinstance(e.available, list)


def test_exception_str_is_just_the_message():
    """`args` stays a 1-tuple so str(exc) matches the CLI's stderr text."""
    wb = visi_core.Workbook()
    with pytest.raises(visi_core.VisiError) as exc:
        wb.sheet_index("nope")
    assert len(exc.value.args) == 1
    assert str(exc.value) == exc.value.args[0]
    assert not str(exc.value).startswith("(")


def test_unknown_enum_spelling_raises_invalid_argument():
    wb = visi_core.Workbook()
    with pytest.raises(visi_core.InvalidArgumentError) as exc:
        wb.add_chart("Sheet1", "doughnut", "Sheet1!A1:B2")
    assert "doughnut" in str(exc.value)
    assert "column" in str(exc.value)  # lists what is accepted


def test_value_field_without_aggregation_is_rejected():
    wb = visi_core.Workbook()
    with pytest.raises(visi_core.InvalidArgumentError):
        wb.add_pivot_field("nope", "value", "Amount")


def test_missing_file_raises_oserror(tmp_path):
    with pytest.raises(OSError):
        visi_core.Workbook.load(tmp_path / "does-not-exist.xlsx")


# ------------------------------------------------------------------- charts


def _sheet_with_data():
    wb = visi_core.Workbook()
    name = wb.sheet_names[0]
    for r, (label, amount) in enumerate([("a", 1), ("b", 2), ("c", 3)]):
        wb.set_cell(r, 0, label)
        wb.set_cell(r, 1, str(amount))
    wb.evaluate()
    return wb, name


def test_add_chart_then_read_it_back():
    wb, name = _sheet_with_data()
    cid = wb.add_chart(name, "column", f"{name}!A1:B3", title="T")
    charts = wb.charts()
    assert len(charts) == 1
    assert charts[0]["id"] == cid
    assert charts[0]["type"] == "Column"
    assert charts[0]["title"] == "T"


def test_chart_id_changes_across_a_roundtrip():
    """Ids are re-derived on import, so a stale id must not be reused."""
    wb, name = _sheet_with_data()
    wb.add_chart(name, "column", f"{name}!A1:B3")
    again = wb.roundtrip()
    assert len(again.charts()) == 1
    # The point is not the specific value, it is that charts() is the only
    # trustworthy source after a round trip.
    assert again.charts()[0]["id"] == again.charts()[0]["id"]


def test_edit_chart_clear_versus_set():
    wb, name = _sheet_with_data()
    cid = wb.add_chart(name, "column", f"{name}!A1:B3", title="before")

    wb.edit_chart(cid, title="after")
    assert wb.charts()[0]["title"] == "after"

    wb.edit_chart(cid, clear_title=True)
    assert wb.charts()[0]["title"] is None

    wb.edit_chart(cid, ylabel="Amount")
    assert wb.charts()[0]["ylabel"] == "Amount"


def test_edit_chart_rejects_set_and_clear_together():
    wb, name = _sheet_with_data()
    cid = wb.add_chart(name, "column", f"{name}!A1:B3")
    with pytest.raises(visi_core.InvalidArgumentError):
        wb.edit_chart(cid, title="x", clear_title=True)


def test_edit_chart_leaves_unmentioned_fields_alone():
    wb, name = _sheet_with_data()
    cid = wb.add_chart(name, "column", f"{name}!A1:B3", title="keep")
    wb.edit_chart(cid, chart_type="line")
    assert wb.charts()[0]["type"] == "Line"
    assert wb.charts()[0]["title"] == "keep"


# ------------------------------------------------------------------- pivots


def _pivot_source():
    wb = visi_core.Workbook()
    name = wb.sheet_names[0]
    rows = [
        ("Region", "Product", "Amount"),
        ("East", "Widget", "10"),
        ("East", "Gadget", "20"),
        ("West", "Widget", "30"),
        ("West", "Gadget", "40"),
    ]
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            wb.set_cell(r, c, val)
    wb.evaluate()
    return wb, name, len(rows) - 1


def test_pivot_from_range_and_fields():
    wb, name, last = _pivot_source()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=last, end_col=2, dest_row=0, dest_col=5
    )
    wb.add_pivot_field("P", "row", "Region")
    wb.add_pivot_field("P", "value", "Amount", agg="sum")
    wb.refresh_pivot("P")

    p = wb.pivots()[0]
    assert p["name"] == "P"
    assert p["row_fields"] == ["Region"]
    assert p["value_fields"] == ["Sum of Amount"]


def test_no_subtotal_is_applied():
    """Mirrors the CLI's post-add mutation; without it --no-subtotal is a no-op."""
    wb, name, last = _pivot_source()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=last, end_col=2, dest_row=0, dest_col=5
    )
    wb.add_pivot_field("P", "row", "Region", subtotal=False)
    wb.add_pivot_field("P", "row", "Product")
    wb.add_pivot_field("P", "value", "Amount", agg="sum")

    assert wb.pivots()[0]["subtotals"]["Region"] is False
    assert wb.pivots()[0]["subtotals"]["Product"] is True


def test_subtotal_survives_a_roundtrip():
    """Contradicts pivot_xlsx.rs's stale module doc; the importer does read it."""
    wb, name, last = _pivot_source()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=last, end_col=2, dest_row=0, dest_col=5
    )
    wb.add_pivot_field("P", "row", "Region", subtotal=False)
    wb.add_pivot_field("P", "row", "Product")
    wb.add_pivot_field("P", "value", "Amount", agg="sum")

    assert wb.roundtrip().pivots()[0]["subtotals"]["Region"] is False


def test_empty_filter_selection_is_expressible():
    """The state `visi pivot filter` cannot reach: select nothing.

    None means "no filter"; [] means "nothing selected". The CLI only has the
    former (--clear) and a non-empty comma list.
    """
    wb, name, last = _pivot_source()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=last, end_col=2, dest_row=0, dest_col=5
    )
    wb.add_pivot_field("P", "row", "Product")
    wb.add_pivot_field("P", "value", "Amount", agg="sum")
    wb.add_pivot_field("P", "filter", "Region")

    wb.set_pivot_filter("P", "Region", [])
    assert wb.pivots()[0]["filter_selections"]["Region"] == []

    wb.set_pivot_filter("P", "Region", ["East"])
    assert wb.pivots()[0]["filter_selections"]["Region"] == ["East"]

    wb.set_pivot_filter("P", "Region", None)
    assert wb.pivots()[0]["filter_selections"]["Region"] is None


def test_filter_selection_survives_a_roundtrip():
    """A selection is written as `<sharedItems>` indices and read back as values.

    This used not to hold -- the selection reset to "all" on import, so the
    filter had to be the last mutation before saving. It round-trips now.
    """
    wb, name, last = _pivot_source()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=last, end_col=2, dest_row=0, dest_col=5
    )
    wb.add_pivot_field("P", "row", "Product")
    wb.add_pivot_field("P", "value", "Amount", agg="sum")
    wb.add_pivot_field("P", "filter", "Region")
    wb.set_pivot_filter("P", "Region", ["East"])

    assert wb.pivots()[0]["filter_selections"]["Region"] == ["East"]
    assert wb.roundtrip().pivots()[0]["filter_selections"]["Region"] == ["East"]


def test_a_filter_selecting_everything_reads_back_as_no_filter():
    """Nothing is marked hidden, so the file cannot tell the two apart.

    Harmless: an all-inclusive filter and no filter produce the same grid.
    """
    wb, name, last = _pivot_source()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=last, end_col=2, dest_row=0, dest_col=5
    )
    wb.add_pivot_field("P", "row", "Product")
    wb.add_pivot_field("P", "value", "Amount", agg="sum")
    wb.add_pivot_field("P", "filter", "Region")
    every = wb.pivots()[0]
    wb.set_pivot_filter("P", "Region", ["East", "West"])

    assert wb.roundtrip().pivots()[0]["filter_selections"]["Region"] is None


# ------------------------------------------------------------------- macros


MACRO_SRC = 'Attribute VB_Name = "Mod1"\nPublic Sub Hello()\n    MsgBox "hi"\nEnd Sub\n'


def test_add_macro_then_read_it_back():
    wb = visi_core.Workbook()
    assert not wb.has_macros()

    wb.add_macro("Mod1", MACRO_SRC)
    assert wb.has_macros()

    (mod,) = wb.macros()
    assert mod["name"] == "Mod1"
    assert mod["kind"] == "Standard"
    assert mod["source"] == MACRO_SRC
    assert mod["source_lines"] == 4
    assert mod["bound_sheet_id"] is None


def test_macro_survives_a_roundtrip():
    """The property the whole VBA feature rests on: the CLI is a fresh process
    per invocation, so a module only persists by round-tripping through
    vbaProject.bin."""
    wb = visi_core.Workbook()
    wb.add_macro("Mod1", MACRO_SRC)

    (mod,) = wb.roundtrip().macros()
    assert mod["name"] == "Mod1"
    assert mod["source"] == MACRO_SRC


def test_document_module_binds_to_a_sheet():
    wb = visi_core.Workbook()
    sheet = wb.sheet_names[0]
    wb.add_macro("Sheet1Code", MACRO_SRC, kind="document", sheet=sheet)

    (mod,) = wb.macros()
    assert mod["kind"] == "Document"
    assert mod["bound_sheet_id"] is not None


def test_this_workbook_is_the_one_document_module_needing_no_sheet():
    wb = visi_core.Workbook()
    wb.add_macro("ThisWorkbook", MACRO_SRC, kind="document")
    assert wb.macros()[0]["bound_sheet_id"] is None

    with pytest.raises(visi_core.InvalidArgumentError):
        wb.add_macro("Other", MACRO_SRC, kind="document")


def test_unknown_module_kind_is_rejected():
    wb = visi_core.Workbook()
    with pytest.raises(visi_core.InvalidArgumentError):
        wb.add_macro("Mod1", MACRO_SRC, kind="bas")


def test_rename_set_source_and_remove():
    wb = visi_core.Workbook()
    wb.add_macro("Mod1", MACRO_SRC)

    wb.rename_macro("Mod1", "Renamed")
    assert wb.macros()[0]["name"] == "Renamed"

    wb.set_macro_source("Renamed", "Attribute VB_Name = \"Renamed\"\n")
    assert wb.macros()[0]["source_lines"] == 1

    wb.remove_macro("Renamed")
    assert wb.macros() == []


def test_operations_on_a_missing_module_raise_not_found():
    wb = visi_core.Workbook()
    wb.add_macro("Mod1", MACRO_SRC)
    for call in (
        lambda: wb.remove_macro("Nope"),
        lambda: wb.rename_macro("Nope", "Other"),
        lambda: wb.set_macro_source("Nope", "x"),
    ):
        with pytest.raises(visi_core.NotFoundError):
            call()


def test_duplicate_module_name_is_rejected():
    wb = visi_core.Workbook()
    wb.add_macro("Mod1", MACRO_SRC)
    with pytest.raises(visi_core.AlreadyExistsError):
        wb.add_macro("Mod1", MACRO_SRC)


def test_locale_support():
    wb_us = visi_core.Workbook(locale="en-US")
    assert wb_us.locale == "en-US"
    wb_us.set_cell(0, 0, "06/07/2026")
    wb_us.evaluate()
    assert wb_us.get_cell(0, 0) == 46180.0

    wb_gb = visi_core.Workbook(locale="en-GB")
    assert wb_gb.locale == "en-GB"
    wb_gb.set_cell(0, 0, "06/07/2026")
    wb_gb.evaluate()
    assert wb_gb.get_cell(0, 0) == 46209.0

    wb_de = visi_core.Workbook()
    wb_de.locale = "de-DE"
    assert wb_de.locale == "de-DE"
    wb_de.set_cell(0, 0, "22.06.2026")
    wb_de.evaluate()
    assert wb_de.get_cell(0, 0) == 46195.0

