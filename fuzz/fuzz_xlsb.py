#!/usr/bin/env python3
import argparse
import datetime
import json
import os
import random
import shutil
import sys
import tempfile

from fuzz_excel import (
    DifferentialComparator,
    ExcelDriver,
    ExcelFuzzGenerator,
    XLSXEvaluatedReader,
)
from visi_driver import VisiDriver, add_backend_arg, bindings_hint


PROFILES = [
    "basic",
    "formula-core",
    "formula-modern",
    "styles-layout",
    "tables",
    "charts",
    "pivots",
    "vba",
    "full",
]


def create_basic_workbook(path, seed, formulas=False):
    import openpyxl

    rng = random.Random(seed)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    rows = [
        ["Text", "Integer", "Float", "Boolean", "Date", "Error"],
        ["alpha", rng.randint(-1000, 1000), 3.25, True, datetime.date(2024, 2, 29), "#DIV/0!"],
        ["", 0, -9.5, False, datetime.datetime(2025, 1, 2, 3, 4, 5), "#N/A"],
    ]
    for row in rows:
        sheet.append(row)
    sheet["E2"].number_format = "mm-dd-yy"
    sheet["E3"].number_format = "mm-dd-yy"
    sheet["F2"].data_type = "e"
    sheet["F3"].data_type = "e"
    if formulas:
        sheet["H1"] = "Formula"
        sheet["H2"] = "=B2+C2"
        sheet["H3"] = "=SUM(B2:B3)"
        sheet["H4"] = "=IF(D2,H2,0)"
        sheet["H5"] = '=A2&"-"&B2'
        sheet["H6"] = "=AVERAGE(B2:C3)"
    second = workbook.create_sheet("Data")
    second.append(["Name", "Value"])
    second.append(["x", 10])
    second.append(["y", 20])
    workbook.save(path)


def create_styles_layout_workbook(path, seed):
    import openpyxl

    create_basic_workbook(path, seed, formulas=True)
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Sheet1"]
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["H"].width = 24
    sheet.row_dimensions[1].height = 24
    sheet.row_dimensions[2].height = 30
    sheet["E2"].number_format = "mm-dd-yy"
    sheet["E3"].number_format = "mm-dd-yy"
    workbook.save(path)


def create_source(path, seed, profile):
    if profile == "basic":
        create_basic_workbook(path, seed, formulas=False)
    elif profile == "formula-core":
        create_basic_workbook(path, seed, formulas=True)
    elif profile == "styles-layout":
        create_styles_layout_workbook(path, seed)
    else:
        ExcelFuzzGenerator(seed=seed).create_fuzz_workbook(path)


def formula_snapshot(cells):
    return {
        key: {
            "formula": (cell.get("formula") or "").replace("$", ""),
            "type": DifferentialComparator.canonical_type(
                cell.get("type"), cell.get("val")
            ),
            "val": cell.get("val"),
        }
        for key, cell in cells.items()
    }


def layout_snapshot(path):
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=False, read_only=False)
    snapshot = {}
    for sheet in workbook.worksheets:
        widths = {
            column: round(dimension.width, 6)
            for column, dimension in sheet.column_dimensions.items()
            if dimension.width is not None
        }
        heights = {
            str(row): round(dimension.height, 6)
            for row, dimension in sheet.row_dimensions.items()
            if dimension.height is not None
        }
        formats = {
            cell.coordinate: cell.number_format
            for row in sheet.iter_rows()
            for cell in row
            if cell.value is not None and cell.number_format != "General"
        }
        snapshot[sheet.title] = {
            "widths": widths,
            "heights": heights,
            "formats": formats,
        }
    return snapshot


def save_failure(output_dir, iteration, seed, work_dir, details):
    failure_dir = os.path.join(
        output_dir, "failures", f"xlsb_iter_{iteration}_seed_{seed}"
    )
    os.makedirs(failure_dir, exist_ok=True)
    for name in [
        "source.xlsx",
        "source.xlsb",
        "visi_semantic.xlsx",
        "visi_out.xlsb",
        "visi_reload.xlsx",
        "excel_oracle.xlsx",
        "excel_visi_out.xlsx",
    ]:
        source = os.path.join(work_dir, name)
        if os.path.exists(source):
            shutil.copy2(source, os.path.join(failure_dir, name))
    with open(os.path.join(failure_dir, "seed.json"), "w", encoding="utf-8") as file:
        json.dump({"seed": seed, **details}, file, indent=2, default=str)
    return failure_dir


def run_iteration(seed, profile, excel, visi, comparator, work_dir):
    paths = {
        name: os.path.join(work_dir, name)
        for name in [
            "source.xlsx",
            "source.xlsb",
            "visi_semantic.xlsx",
            "visi_out.xlsb",
            "visi_reload.xlsx",
            "excel_oracle.xlsx",
            "excel_visi_out.xlsx",
        ]
    }
    create_source(paths["source.xlsx"], seed, profile)
    excel.convert(paths["source.xlsx"], paths["source.xlsb"], recalculate=True)
    visi.run(paths["source.xlsb"], paths["visi_semantic.xlsx"])
    visi.run(paths["source.xlsb"], paths["visi_out.xlsb"])
    excel.recalculate(paths["source.xlsb"], paths["excel_oracle.xlsx"])
    excel.recalculate(paths["visi_out.xlsb"], paths["excel_visi_out.xlsx"])
    visi.run(paths["visi_out.xlsb"], paths["visi_reload.xlsx"])

    semantic = XLSXEvaluatedReader.read_evaluated_cells(paths["visi_semantic.xlsx"])
    oracle = XLSXEvaluatedReader.read_evaluated_cells(paths["excel_oracle.xlsx"])
    excel_output = XLSXEvaluatedReader.read_evaluated_cells(
        paths["excel_visi_out.xlsx"]
    )
    reload = XLSXEvaluatedReader.read_evaluated_cells(paths["visi_reload.xlsx"])
    semantic_ok, semantic_mismatches = comparator.compare(semantic, oracle)
    output_ok, output_mismatches = comparator.compare(excel_output, oracle)
    reload_ok, reload_mismatches = comparator.compare(reload, excel_output)
    formula_ok = formula_snapshot(reload) == formula_snapshot(excel_output)
    layout_ok = True
    if profile == "styles-layout":
        semantic_layout = layout_snapshot(paths["visi_semantic.xlsx"])
        oracle_layout = layout_snapshot(paths["excel_oracle.xlsx"])
        output_layout = layout_snapshot(paths["excel_visi_out.xlsx"])
        reload_layout = layout_snapshot(paths["visi_reload.xlsx"])
        layout_ok = (
            semantic_layout == oracle_layout
            and output_layout == oracle_layout
            and reload_layout == output_layout
        )
    return semantic_ok and output_ok and reload_ok and formula_ok and layout_ok, {
        "profile": profile,
        "semantic_mismatches": semantic_mismatches[:100],
        "output_mismatches": output_mismatches[:100],
        "reload_mismatches": reload_mismatches[:100],
        "formula_snapshot_equal": formula_ok,
        "layout_snapshot_equal": layout_ok,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Differential XLSB import, evaluation, writer, and reload fuzzing"
    )
    parser.add_argument("--iterations", type=int, default=20)
    parser.add_argument("--seed", type=int)
    parser.add_argument("--profile", choices=PROFILES, default="basic")
    parser.add_argument("--driver", choices=["auto", "applescript", "win32com", "mock", "cli"], default="auto")
    parser.add_argument("--excel-path")
    parser.add_argument("--output-dir", default="fuzz_results/xlsb")
    parser.add_argument("--strict-error-class", action="store_true", default=True)
    add_backend_arg(parser)
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    excel = ExcelDriver(args.excel_path, args.driver)
    try:
        visi = VisiDriver(backend=args.backend)
    except RuntimeError:
        print(bindings_hint(), file=sys.stderr)
        raise
    comparator = DifferentialComparator(strict_error_class=args.strict_error_class)
    root_rng = random.Random(args.seed)
    failures = 0
    for iteration in range(1, args.iterations + 1):
        seed = args.seed if args.seed is not None and args.iterations == 1 else root_rng.randrange(1 << 31)
        with tempfile.TemporaryDirectory(prefix="visi-xlsb-") as work_dir:
            try:
                ok, details = run_iteration(
                    seed, args.profile, excel, visi, comparator, work_dir
                )
            except Exception as error:
                ok = False
                details = {"profile": args.profile, "exception": repr(error)}
            verdict = "PASS" if ok else "FAIL"
            print(f"[{iteration}/{args.iterations}] seed={seed} profile={args.profile} {verdict}")
            if not ok:
                failures += 1
                failure_dir = save_failure(
                    args.output_dir, iteration, seed, work_dir, details
                )
                print(f"  artifacts: {failure_dir}")
    print(f"Passed: {args.iterations - failures}  Failed: {failures}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
