#!/usr/bin/env python3
import argparse
import os
import tempfile

from fuzz_excel import ExcelDriver, XLSXEvaluatedReader


def main():
    parser = argparse.ArgumentParser(
        description="Probe whether manual-calculation XLSB conversion preserves formula caches"
    )
    parser.add_argument("--driver", choices=["auto", "applescript", "win32com"], default="auto")
    parser.add_argument("--excel-path")
    args = parser.parse_args()

    import openpyxl

    excel = ExcelDriver(args.excel_path, args.driver)
    with tempfile.TemporaryDirectory(prefix="xlsb-cache-probe-") as directory:
        source = os.path.join(directory, "source.xlsx")
        calculated = os.path.join(directory, "calculated.xlsx")
        binary = os.path.join(directory, "manual.xlsb")
        roundtrip = os.path.join(directory, "roundtrip.xlsx")
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet["A1"] = 40
        sheet["B1"] = "=A1+2"
        workbook.save(source)
        excel.recalculate(source, calculated)
        excel.convert(
            calculated,
            binary,
            recalculate=False,
            manual_calculation=True,
        )
        excel.convert(
            binary,
            roundtrip,
            recalculate=False,
            manual_calculation=True,
        )
        before = XLSXEvaluatedReader.read_evaluated_cells(calculated)
        after = XLSXEvaluatedReader.read_evaluated_cells(roundtrip)
        before_value = before[("Sheet", "B1")]["val"]
        after_value = after[("Sheet", "B1")]["val"]
        print(f"before={before_value!r} after={after_value!r}")
        if before_value != 42 or after_value != before_value:
            raise SystemExit(1)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
