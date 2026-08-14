#!/usr/bin/env python3
"""Bindings/CLI equivalence.

The `visi_core` extension module and the `visi` CLI must be observationally
identical for every operation the fuzz harness drives. Nothing else checks
this: the two backends duplicate a little logic on purpose (`edit_chart`'s
clear-vs-set flags, `add_pivot_field`'s post-add subtotal mutation), and
without a test that duplication drifts silently and the fuzzer starts
measuring something other than what it reports.

Corpus: every saved `fuzz_results/failures/*/source.xlsx` -- real inputs that
already broke something once -- plus freshly generated workbooks at fixed
seeds, so the test is still meaningful on a clean checkout.

Comparison is over *parsed content*, never bytes: docProps/core.xml carries a
creation timestamp and chart/pivot ids come from a random `generate_unique_id`,
so two runs never produce identical archives.

    source fuzz/venv/bin/activate
    pytest fuzz/test_backend_parity.py
"""

import glob
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from chart_xlsx_reader import read_charts
from fuzz_chart import ChartFuzzGenerator
from fuzz_excel import ExcelFuzzGenerator, XLSXEvaluatedReader
from fuzz_pivot import DEST_CELL, DEST_RC, PIVOT_NAME, PivotFuzzGenerator
from visi_driver import (
    VisiChartDriver,
    VisiDriver,
    VisiPivotDriver,
    bindings_available,
)

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

pytestmark = pytest.mark.skipif(
    not bindings_available(),
    reason="build the bindings with `maturin develop -m visi-python/Cargo.toml --release`",
)


def _cli_binary_exists():
    from visi_driver import resolve_visi_binary

    p = resolve_visi_binary(None)
    return p is not None and os.path.exists(p)


requires_cli = pytest.mark.skipif(
    not _cli_binary_exists(), reason="no compiled visi binary; run `cargo build --release`"
)


def saved_failure_sources():
    return sorted(
        glob.glob(os.path.join(PROJECT_ROOT, "fuzz_results", "failures", "*", "source.xlsx"))
    )


# --------------------------------------------------------------------- eval


@requires_cli
@pytest.mark.parametrize("seed", [48291, 1, 7, 12345, 999983])
def test_eval_parity_on_generated_workbooks(seed, tmp_path):
    src = str(tmp_path / "source.xlsx")
    ExcelFuzzGenerator(seed=seed).create_fuzz_workbook(src, num_rows=10, num_cols=5)
    _assert_eval_parity(src, tmp_path)


@requires_cli
@pytest.mark.parametrize("src", saved_failure_sources())
def test_eval_parity_on_saved_failures(src, tmp_path):
    _assert_eval_parity(src, tmp_path)


def _assert_eval_parity(src, tmp_path):
    cli_out = str(tmp_path / "cli.xlsx")
    bnd_out = str(tmp_path / "bnd.xlsx")

    VisiDriver(backend="subprocess").run(src, cli_out)
    VisiDriver(backend="bindings").run(src, bnd_out)

    cli_cells = XLSXEvaluatedReader.read_evaluated_cells(cli_out)
    bnd_cells = XLSXEvaluatedReader.read_evaluated_cells(bnd_out)

    assert set(cli_cells) == set(bnd_cells), "different cells present"
    mismatches = {
        key: (cli_cells[key]["val"], bnd_cells[key]["val"])
        for key in cli_cells
        if cli_cells[key]["val"] != bnd_cells[key]["val"]
    }
    assert not mismatches, f"{len(mismatches)} value mismatch(es): {list(mismatches.items())[:5]}"


def test_run_returns_the_bytes_it_wrote(tmp_path):
    """The main loop parses the return value instead of re-reading the file."""
    src = str(tmp_path / "source.xlsx")
    out = str(tmp_path / "out.xlsx")
    ExcelFuzzGenerator(seed=4242).create_fuzz_workbook(src, num_rows=5, num_cols=3)

    data = VisiDriver(backend="bindings").run(src, out)
    with open(out, "rb") as f:
        assert data == f.read()
    assert XLSXEvaluatedReader.read_evaluated_cells_bytes(
        data
    ) == XLSXEvaluatedReader.read_evaluated_cells(out)


# ------------------------------------------------------------------- charts


@requires_cli
@pytest.mark.parametrize("seed", range(6))
def test_chart_parity(seed, tmp_path):
    src = str(tmp_path / "source.xlsx")
    range_str, add_config, edit_config = ChartFuzzGenerator(seed=seed).generate(
        src, num_rows=6
    )

    cli_out = str(tmp_path / "cli.xlsx")
    bnd_out = str(tmp_path / "bnd.xlsx")
    VisiChartDriver(backend="subprocess").run(
        src, range_str, add_config, edit_config, cli_out
    )
    VisiChartDriver(backend="bindings").run(
        src, range_str, add_config, edit_config, bnd_out
    )

    assert read_charts(cli_out) == read_charts(bnd_out)


# ------------------------------------------------------------------- pivots


@requires_cli
@pytest.mark.parametrize("seed", range(10))
@pytest.mark.parametrize("use_table", [False, True])
def test_pivot_parity(seed, use_table, tmp_path):
    src = str(tmp_path / "source.xlsx")
    config = PivotFuzzGenerator(seed=seed).generate(src, num_rows=8, use_table=use_table)

    if _wants_empty_filter_selection(config):
        pytest.skip("empty filter selection is bindings-only; see test below")

    cli_out = str(tmp_path / "cli.xlsx")
    bnd_out = str(tmp_path / "bnd.xlsx")
    VisiPivotDriver(backend="subprocess").run(
        src, config, cli_out, PIVOT_NAME, DEST_CELL, DEST_RC
    )
    VisiPivotDriver(backend="bindings").run(
        src, config, bnd_out, PIVOT_NAME, DEST_CELL, DEST_RC
    )

    cli_cells = XLSXEvaluatedReader.read_evaluated_cells(cli_out)
    bnd_cells = XLSXEvaluatedReader.read_evaluated_cells(bnd_out)
    assert cli_cells.keys() == bnd_cells.keys()
    for key in cli_cells:
        assert cli_cells[key]["val"] == bnd_cells[key]["val"], key


def _wants_empty_filter_selection(config):
    ff = config.get("filter_field")
    return bool(ff) and not ff.get("values")


def test_empty_filter_selection_is_bindings_only(tmp_path):
    """The one deliberate divergence, asserted rather than left implicit.

    `visi pivot filter` takes a non-empty comma list or --clear, with no verb
    for "select nothing", so the subprocess backend leaves such a field
    unfiltered. The bindings can express it. If the CLI ever grows the verb,
    this test should start failing and the skip in test_pivot_parity can go.
    """
    import visi_core

    wb = visi_core.Workbook()
    rows = [
        ("Region", "Amount"),
        ("East", "10"),
        ("West", "20"),
    ]
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            wb.set_cell(r, c, val)
    wb.evaluate()
    wb.add_pivot_from_range(
        "P", start_row=0, start_col=0, end_row=2, end_col=1, dest_row=0, dest_col=4
    )
    wb.add_pivot_field("P", "value", "Amount", agg="sum")
    wb.add_pivot_field("P", "filter", "Region")

    wb.set_pivot_filter("P", "Region", [])
    assert wb.pivots()[0]["filter_selections"]["Region"] == []

    wb.set_pivot_filter("P", "Region", None)
    assert wb.pivots()[0]["filter_selections"]["Region"] is None


# --------------------------------------------------------------- error text


@requires_cli
def test_binding_error_text_matches_the_cli(tmp_path):
    """str(exc) must equal the CLI's stderr minus its "Error: " prefix."""
    import subprocess

    import visi_core

    from visi_driver import resolve_visi_binary

    src = str(tmp_path / "source.xlsx")
    visi_core.Workbook().save(src)

    with pytest.raises(visi_core.NotFoundError) as exc:
        visi_core.Workbook.load(src).refresh_pivot("no-such-pivot")

    res = subprocess.run(
        [resolve_visi_binary(None), "pivot", "refresh", src, "--name", "no-such-pivot"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    assert res.returncode != 0
    cli_msg = res.stderr.strip().removeprefix("Error: ")
    assert str(exc.value) == cli_msg
