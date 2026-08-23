#!/usr/bin/env python3
"""
Differential comparison harness for pre-saved Excel (.xlsx) files.
=================================================================
Runs pre-saved .xlsx files (by default from fuzz/presaved/) through
both visi and real Microsoft Excel, comparing evaluated cell values.

Usage:
    source fuzz/venv/bin/activate

    # Run all files in fuzz/presaved/
    python fuzz/fuzz_presaved.py

    # Run a specific file
    python fuzz/fuzz_presaved.py fuzz/presaved/NVDA_Put_Implied_Volatility.xlsx

    # Pass explicit Excel path or driver
    python fuzz/fuzz_presaved.py --excel-path "/Applications/Microsoft Excel.app"
"""

import argparse
import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fuzz_excel import (
    DifferentialComparator,
    ExcelDriver,
    XLSXEvaluatedReader,
    bindings_hint,
)
from visi_driver import VisiDriver, add_backend_arg

DEFAULT_PRESAVED_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "presaved"
)


def find_presaved_files(path: str) -> list[str]:
    """Resolves a file or directory path to a list of .xlsx files."""
    if os.path.isfile(path):
        return [path]
    if os.path.isdir(path):
        files = []
        for root, _, filenames in os.walk(path):
            for fn in filenames:
                if fn.endswith(".xlsx") and not fn.startswith("~$"):
                    files.append(os.path.join(root, fn))
        files.sort()
        return files
    return []


def compare_presaved_file(
    file_path: str,
    excel_path: str = None,
    driver_type: str = "auto",
    backend: str = "auto",
    visi_path: str = None,
    output_dir: str = None,
    strict_error_class: bool = False,
    sheet_filter: str = None,
    cell_filter: str = None,
    verbose: bool = False,
) -> tuple[bool, list[dict], dict]:
    """
    Opens `file_path` in both visi and Microsoft Excel, recalculates,
    and returns (is_match, mismatches, stats).
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

    temp_dir = tempfile.mkdtemp(prefix="fuzz_presaved_")
    cleanup_temp = False
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        work_dir = output_dir
    else:
        work_dir = temp_dir
        cleanup_temp = True

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    source_xlsx = os.path.join(work_dir, f"{base_name}_source.xlsx")
    visi_out_xlsx = os.path.join(work_dir, f"{base_name}_visi_out.xlsx")
    excel_out_xlsx = os.path.join(work_dir, f"{base_name}_excel_out.xlsx")

    try:
        shutil.copyfile(file_path, source_xlsx)

        # 1. Evaluate with visi
        visi_driver = VisiDriver(binary_path=visi_path, backend=backend)
        if backend == "auto" and visi_driver.backend != "bindings":
            print(bindings_hint(), file=sys.stderr)

        start_visi = time.time()
        visi_bytes = visi_driver.run(source_xlsx, visi_out_xlsx)
        visi_duration = time.time() - start_visi

        # 2. Evaluate with Excel
        excel_driver = ExcelDriver(excel_path=excel_path, driver_type=driver_type)
        start_excel = time.time()
        excel_driver.run(source_xlsx, excel_out_xlsx)
        excel_duration = time.time() - start_excel

        # 3. Read evaluated cells
        visi_cells = XLSXEvaluatedReader.read_evaluated_cells_bytes(
            visi_bytes, source=visi_out_xlsx
        )
        excel_cells = XLSXEvaluatedReader.read_evaluated_cells(excel_out_xlsx)

        if sheet_filter:
            visi_cells = {k: v for k, v in visi_cells.items() if k[0] == sheet_filter}
            excel_cells = {k: v for k, v in excel_cells.items() if k[0] == sheet_filter}
        if cell_filter:
            visi_cells = {k: v for k, v in visi_cells.items() if k[1] == cell_filter}
            excel_cells = {k: v for k, v in excel_cells.items() if k[1] == cell_filter}

        # 4. Compare cell values
        comparator = DifferentialComparator(strict_error_class=strict_error_class)
        is_match, mismatches = comparator.compare(visi_cells, excel_cells)

        stats = {
            "file": file_path,
            "total_cells": len(set(visi_cells.keys()).union(set(excel_cells.keys()))),
            "visi_cells_count": len(visi_cells),
            "excel_cells_count": len(excel_cells),
            "mismatch_count": len(mismatches),
            "visi_duration_sec": visi_duration,
            "excel_duration_sec": excel_duration,
            "visi_backend": visi_driver.describe(),
            "excel_driver": excel_driver.driver_type,
            "work_dir": work_dir,
        }

        return is_match, mismatches, stats

    finally:
        if cleanup_temp and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="Compare evaluation of pre-saved Excel files across visi and Microsoft Excel."
    )
    parser.add_argument(
        "target",
        nargs="?",
        default=DEFAULT_PRESAVED_DIR,
        help=f"Path to .xlsx file or directory of files (default: {DEFAULT_PRESAVED_DIR})",
    )
    parser.add_argument(
        "--excel-path",
        default=None,
        help="Path to Excel executable/application (defaults to auto-detected system Excel)",
    )
    parser.add_argument(
        "--driver",
        choices=["auto", "applescript", "win32com", "mock"],
        default="auto",
        help="Excel driver type",
    )
    add_backend_arg(parser)
    parser.add_argument("--visi-path", default=None, help="Custom path to visi binary")
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Directory to save generated comparison artifacts",
    )
    parser.add_argument(
        "--strict-error-class",
        action="store_true",
        help="Flag error class mismatches as failures",
    )
    parser.add_argument("--sheet", default=None, help="Compare only the specified sheet name")
    parser.add_argument("--cell", default=None, help="Compare only the specified cell (e.g. B9)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")

    args = parser.parse_args()

    files = find_presaved_files(args.target)
    if not files:
        print(f"[ERROR] No .xlsx files found at: {args.target}", file=sys.stderr)
        sys.exit(1)

    print("=====================================================================")
    print("      visi vs. Microsoft Excel Pre-Saved Workbook Comparison        ")
    print("=====================================================================")
    print(f" Target      : {args.target} ({len(files)} file(s))")
    print(f" Excel Path  : {args.excel_path or 'Auto-detect'}")
    print(f" Driver      : {args.driver}")
    print(f" Backend     : {args.backend}")
    print("=====================================================================\n")

    overall_passed = True
    total_files = len(files)

    for idx, file_path in enumerate(files, 1):
        rel_name = os.path.relpath(file_path, os.getcwd())
        print(f"[{idx}/{total_files}] Testing {rel_name} ...", end=" ", flush=True)

        try:
            is_match, mismatches, stats = compare_presaved_file(
                file_path=file_path,
                excel_path=args.excel_path,
                driver_type=args.driver,
                backend=args.backend,
                visi_path=args.visi_path,
                output_dir=args.output_dir,
                strict_error_class=args.strict_error_class,
                sheet_filter=args.sheet,
                cell_filter=args.cell,
                verbose=args.verbose,
            )
        except Exception as e:
            print(f"[ERROR] {e}")
            overall_passed = False
            continue

        if is_match:
            print(
                f"[OK] ({stats["total_cells"]} cells matched, "
                f"visi: {stats["visi_duration_sec"]:.3f}s, excel: {stats["excel_duration_sec"]:.3f}s)"
            )
        else:
            overall_passed = False
            print(f"[FAILED] ({len(mismatches)} mismatch(es))")
            for m_idx, m in enumerate(mismatches[:20], 1):
                sheet, cell = m["key"]
                formula = f" (Formula: ={m["formula"]})" if m.get("formula") else ""
                print(f"    {m_idx:2d}. Sheet: {sheet!r} Cell: {cell}{formula}")
                print(f"        visi : {m["visi"]!r}")
                print(f"        Excel: {m["excel"]!r}")
                print(f"        Diff : {m["reason"]}")
            if len(mismatches) > 20:
                print(f"    ... and {len(mismatches) - 20} more mismatches.")

    print("\n=====================================================================")
    if overall_passed:
        print("[SUMMARY] All pre-saved workbooks passed clean against Excel.")
        sys.exit(0)
    else:
        print("[SUMMARY] Mismatches detected across tested workbooks.")
        sys.exit(1)


if __name__ == "__main__":
    main()
