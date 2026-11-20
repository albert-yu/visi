from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

import pytest
import visi_core

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def resolve_visi_binary() -> str | None:
    candidates = [
        PROJECT_ROOT / "target" / flavor / "visi"
        for flavor in ("release", "debug")
    ]
    existing = [p for p in candidates if p.is_file()]
    if not existing:
        return None
    return str(max(existing, key=lambda p: p.stat().st_mtime))


requires_cli = pytest.mark.skipif(
    resolve_visi_binary() is None,
    reason="no compiled visi binary; run `cargo build --release`",
)

MACRO_SRC = (
    'Attribute VB_Name = "Mod1"\n'
    "Public Sub Hello()\n"
    '    Range("A1").Value = 1\n'
    "End Sub\n"
)

RUN_MACRO_SRC = (
    'Attribute VB_Name = "Runner"\n'
    "Public Function Go() As Variant\n"
    '    Range("A1").Value = 4\n'
    '    Range("B1").Formula = "=A1*2"\n'
    '    Range("C1").Value = Application.WorksheetFunction.Sum(Range("A1:B1"))\n'
    '    Go = Range("C1").Value\n'
    "End Function\n"
    "Public Function Peek() As Variant\n"
    "    Peek = ThisWorkbook.Worksheets.Count\n"
    "End Function\n"
)


@requires_cli
def test_eval_parity(tmp_path: Path) -> None:
    src = tmp_path / "source.xlsx"
    cli_out = tmp_path / "cli.xlsx"
    bnd_out = tmp_path / "bnd.xlsx"

    base = visi_core.Workbook()
    base.set_cell(0, 0, "10")
    base.set_cell(0, 1, "20")
    base.set_cell(0, 2, "=A1+B1")
    base.set_cell(1, 0, "=1/0")
    base.save(str(src))

    res = subprocess.run(
        [resolve_visi_binary(), "eval", str(src), "--output", str(cli_out), "--quiet"],
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode == 0, res.stderr

    wb = visi_core.Workbook.load(str(src))
    wb.evaluate()
    wb.save(str(bnd_out))

    cli_wb = visi_core.Workbook.load(str(cli_out))
    bnd_wb = visi_core.Workbook.load(str(bnd_out))
    for r, c in [(0, 0), (0, 1), (0, 2), (1, 0)]:
        assert cli_wb.get_display(r, c) == bnd_wb.get_display(r, c)


@requires_cli
def test_xlsb_eval_parity(tmp_path: Path) -> None:
    src = tmp_path / "source.xlsx"
    cli_out = tmp_path / "cli.xlsb"
    bnd_out = tmp_path / "bnd.xlsb"

    wb = visi_core.Workbook()
    wb.set_cell(0, 0, "4")
    wb.set_cell(0, 1, "=A1+6")
    wb.evaluate()
    wb.save(str(src))

    res = subprocess.run(
        [resolve_visi_binary(), "eval", str(src), "--output", str(cli_out), "--quiet"],
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode == 0, res.stderr

    loaded = visi_core.Workbook.load(str(src))
    loaded.evaluate()
    loaded.save(str(bnd_out))

    cli_wb = visi_core.Workbook.load(str(cli_out))
    bnd_wb = visi_core.Workbook.load(str(bnd_out))
    assert cli_wb.get_cell(0, 0) == bnd_wb.get_cell(0, 0) == 4
    assert cli_wb.get_cell(0, 1) == bnd_wb.get_cell(0, 1) == 10


@requires_cli
def test_chart_edit_title_clear_parity(tmp_path: Path) -> None:
    src = tmp_path / "source.xlsx"
    cli_out = tmp_path / "cli.xlsx"
    bnd_out = tmp_path / "bnd.xlsx"

    base = visi_core.Workbook()
    for r, (label, val) in enumerate([("Cat", "Val"), ("A", "10"), ("B", "20")]):
        base.set_cell(r, 0, label)
        base.set_cell(r, 1, val)
    base.add_chart("Sheet1", "column", "Sheet1!A1:B3", title="Initial")
    base.save(str(src))

    chart_id = visi_core.Workbook.load(str(src)).charts()[0]["id"]

    res = subprocess.run(
        [
            resolve_visi_binary(),
            "chart",
            "edit",
            str(src),
            "--id",
            str(chart_id),
            "--clear-title",
            "--output",
            str(cli_out),
            "--quiet",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode == 0, res.stderr

    wb = visi_core.Workbook.load(str(src))
    wb.edit_chart(chart_id, clear_title=True)
    wb.save(str(bnd_out))

    assert visi_core.Workbook.load(str(cli_out)).charts() == visi_core.Workbook.load(
        str(bnd_out)
    ).charts()


@requires_cli
def test_pivot_add_field_parity(tmp_path: Path) -> None:
    src = tmp_path / "source.xlsx"
    cli_out = tmp_path / "cli.xlsx"
    bnd_out = tmp_path / "bnd.xlsx"

    base = visi_core.Workbook()
    rows = [("Region", "Amount"), ("East", "10"), ("West", "20"), ("East", "15")]
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            base.set_cell(r, c, val)
    base.evaluate()
    base.add_pivot_from_range(
        "P1", start_row=0, start_col=0, end_row=3, end_col=1, dest_row=0, dest_col=4
    )
    base.save(str(src))

    for cmd in [
        [
            resolve_visi_binary(),
            "pivot",
            "add-field",
            str(src),
            "--name",
            "P1",
            "--area",
            "row",
            "--column",
            "Region",
            "--output",
            str(cli_out),
            "--quiet",
        ],
        [
            resolve_visi_binary(),
            "pivot",
            "add-field",
            str(cli_out),
            "--name",
            "P1",
            "--area",
            "value",
            "--column",
            "Amount",
            "--agg",
            "sum",
            "--label",
            "Total Amount",
            "--in-place",
            "--quiet",
        ],
        [
            resolve_visi_binary(),
            "pivot",
            "refresh",
            str(cli_out),
            "--name",
            "P1",
            "--in-place",
            "--quiet",
        ],
    ]:
        res = subprocess.run(cmd, capture_output=True, text=True, check=False)
        assert res.returncode == 0, res.stderr

    wb = visi_core.Workbook.load(str(src))
    wb.add_pivot_field("P1", "row", "Region")
    wb.add_pivot_field("P1", "value", "Amount", agg="sum", label="Total Amount")
    wb.refresh_pivot("P1")
    wb.save(str(bnd_out))

    cli_wb = visi_core.Workbook.load(str(cli_out))
    bnd_wb = visi_core.Workbook.load(str(bnd_out))
    for r in range(5):
        for c in range(4, 7):
            assert cli_wb.get_display(r, c) == bnd_wb.get_display(r, c)


def test_empty_filter_selection_is_bindings_only() -> None:
    wb = visi_core.Workbook()
    rows = [("Region", "Amount"), ("East", "10"), ("West", "20")]
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


@requires_cli
def test_binding_error_text_matches_the_cli(tmp_path: Path) -> None:
    src = tmp_path / "source.xlsx"
    visi_core.Workbook().save(str(src))

    with pytest.raises(visi_core.NotFoundError) as exc:
        visi_core.Workbook.load(str(src)).refresh_pivot("no-such-pivot")

    res = subprocess.run(
        [resolve_visi_binary(), "pivot", "refresh", str(src), "--name", "no-such-pivot"],
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode != 0
    cli_msg = res.stderr.strip().removeprefix("Error: ")
    assert str(exc.value) == cli_msg


@requires_cli
@pytest.mark.parametrize("kind,sheet", [("standard", None), ("document", "Sheet1")])
def test_macro_add_parity(kind: str, sheet: str | None, tmp_path: Path) -> None:
    src = tmp_path / "source.xlsx"
    bas = tmp_path / "mod1.bas"
    cli_out = tmp_path / "cli.xlsm"
    bindings_out = tmp_path / "bindings.xlsm"

    base = visi_core.Workbook()
    if sheet is not None:
        sheet = base.sheet_names[0]
    base.save(str(src))
    bas.write_text(MACRO_SRC)

    cmd = [
        resolve_visi_binary(),
        "macro",
        "add",
        str(src),
        "--name",
        "Mod1",
        "--kind",
        kind,
        "--source-file",
        str(bas),
        "--output",
        str(cli_out),
        "--quiet",
    ]
    if sheet is not None:
        cmd += ["--sheet", sheet]
    res = subprocess.run(cmd, capture_output=True, text=True, check=False)
    assert res.returncode == 0, res.stderr

    wb = visi_core.Workbook.load(str(src))
    wb.add_macro("Mod1", MACRO_SRC, kind=kind, sheet=sheet)
    wb.save(str(bindings_out))

    def modules(path: Path):
        return [
            (m["name"], m["kind"], m["source"], m["bound_sheet_id"] is not None)
            for m in visi_core.Workbook.load(str(path)).macros()
        ]

    assert modules(cli_out) == modules(bindings_out)


@requires_cli
def test_macro_list_parity(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsm"
    wb = visi_core.Workbook()
    wb.add_macro("Mod1", MACRO_SRC)
    wb.add_macro(
        "ThisWorkbook", 'Attribute VB_Name = "ThisWorkbook"\n', kind="document"
    )
    wb.save(str(path))

    res = subprocess.run(
        [resolve_visi_binary(), "macro", "list", str(path), "--json"],
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode == 0, res.stderr
    cli = json.loads(res.stdout)

    got = visi_core.Workbook.load(str(path)).macros()
    assert [(m["name"], m["kind"], m["source_lines"]) for m in cli] == [
        (m["name"], m["kind"], m["source_lines"]) for m in got
    ]


@requires_cli
def test_macro_run_parity(tmp_path: Path) -> None:
    src = tmp_path / "source.xlsm"
    cli_out = tmp_path / "cli.xlsm"
    bindings_out = tmp_path / "bindings.xlsm"

    base = visi_core.Workbook()
    base.add_macro("Runner", RUN_MACRO_SRC)
    base.save(str(src))

    res = subprocess.run(
        [
            resolve_visi_binary(),
            "macro",
            "run",
            str(src),
            "--name",
            "Go",
            "--output",
            str(cli_out),
            "--json",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert res.returncode == 0, res.stderr
    cli = json.loads(res.stdout)

    wb = visi_core.Workbook.load(str(src))
    type_name, value, mutated = wb.run_macro("Go")
    wb.save(str(bindings_out))

    assert (cli["type"], cli["value"], cli["mutated"]) == (type_name, value, mutated)
    assert mutated is True

    def cells(path: Path):
        w = visi_core.Workbook.load(str(path))
        return [w.get_display(0, c) for c in range(3)]

    assert cells(cli_out) == cells(bindings_out)


@requires_cli
def test_macro_run_without_a_write_target_is_an_error_only_when_it_mutated(
    tmp_path: Path,
) -> None:
    src = tmp_path / "source.xlsm"
    base = visi_core.Workbook()
    base.add_macro("Runner", RUN_MACRO_SRC)
    base.save(str(src))

    def run(proc: str):
        return subprocess.run(
            [resolve_visi_binary(), "macro", "run", str(src), "--name", proc],
            capture_output=True,
            text=True,
            check=False,
        )

    mutating = run("Go")
    assert mutating.returncode != 0
    assert "--output" in mutating.stderr

    reading = run("Peek")
    assert reading.returncode == 0, reading.stderr
    assert "Running VBA procedure" in reading.stderr

    assert visi_core.Workbook.load(str(src)).run_macro("Peek")[2] is False
    assert visi_core.Workbook.load(str(src)).run_macro("Go")[2] is True
