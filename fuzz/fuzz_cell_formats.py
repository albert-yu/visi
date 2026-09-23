#!/usr/bin/env python3
import argparse
import datetime as dt
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import openpyxl
except ImportError as exc:
    raise SystemExit(
        "openpyxl is required: source fuzz/venv/bin/activate && pip install -r fuzz/requirements.txt"
    ) from exc

try:
    import visi_core
except ImportError as exc:
    raise SystemExit(
        "visi_core bindings are required for cell-format fuzzing. Build them with:\n"
        "  source fuzz/venv/bin/activate && maturin develop -m visi-python/Cargo.toml --release"
    ) from exc


NUMBER_FORMATS = [
    "General",
    "0",
    "0.00",
    "0.###",
    "#,##0",
    "#,##0.00",
    "0%",
    "0.0%",
    "0.00%",
    "$#,##0.00",
    '"USD "0.00',
    '0.00" kg"',
    "0.00E+00",
]

SECTION_FORMATS = [
    "0.00;[Red]-0.00;0.00",
    '0.00;[Red](0.00);"zero"',
    '$#,##0.00;[Red]($#,##0.00);"-"',
    '0%;[Red]-0%;"flat"',
]

DATE_FORMATS = [
    "m/d/yy",
    "m/d/yyyy",
    "yyyy-mm-dd",
    "yyyy/mm/dd",
    "d-mmm-yyyy",
    "mmm-d-yyyy",
]

TEXT_FORMATS = [
    "@",
    '"ID-"@',
    '@"!"',
    'General;General;General;"text:"@',
]

TEXT_VALUES = [
    "alpha",
    "Z9",
    "north",
    "A-42",
    "x y",
    "plain text",
]


def col_name(idx):
    out = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def cell_ref(row, col):
    return f"{col_name(col + 1)}{row + 1}"


def excel_serial(date):
    return (date - dt.date(1899, 12, 30)).days


def random_number(rng):
    if rng.randrange(8) == 0:
        return 0
    if rng.randrange(4) == 0:
        return rng.randint(-9999, 9999)
    places = rng.randint(1, 4)
    return round(rng.uniform(-9999, 9999), places)


def random_date_serial(rng):
    base = dt.date(2018, 1, 1)
    return excel_serial(base + dt.timedelta(days=rng.randint(0, 5000)))


def make_case(kind, source, fmt):
    return {"kind": kind, "source": source, "format": fmt}


def build_source_workbook(path, rng):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    rows = rng.randint(4, 9)
    cols = 6
    cases = []

    for one_based_col in range(1, cols + 1):
        ws.column_dimensions[col_name(one_based_col)].width = 32

    for row in range(rows):
        excel_row = row + 1
        numeric_a = random_number(rng)
        numeric_b = random_number(rng)
        date_value = random_date_serial(rng)
        formula_factor = rng.choice([2, 3, 10, 0.5])
        date_delta = rng.randint(1, 31)
        text = rng.choice(TEXT_VALUES)

        row_cases = [
            make_case("number", numeric_a, rng.choice(NUMBER_FORMATS)),
            make_case("number", numeric_b, rng.choice(SECTION_FORMATS)),
            make_case(
                "formula", f"=A{excel_row}*{formula_factor}", rng.choice(NUMBER_FORMATS)
            ),
            make_case("date", date_value, rng.choice(DATE_FORMATS)),
            make_case(
                "formula", f"=D{excel_row}+{date_delta}", rng.choice(DATE_FORMATS)
            ),
            make_case("text", text, rng.choice(TEXT_FORMATS)),
        ]
        for col, case in enumerate(row_cases):
            cell = ws.cell(row=row + 1, column=col + 1)
            cell.value = case["source"]
            cell.number_format = case["format"]
            cases.append({"cell": cell_ref(row, col), **case})

    wb.save(path)
    return cases


class ExcelDisplayDriver:
    def __init__(self, excel_path=None, driver_type="auto", timeout=30):
        self.excel_path = excel_path
        self.timeout = timeout
        self.driver_type = driver_type
        if driver_type == "auto":
            if sys.platform == "darwin":
                self.driver_type = "applescript"
            elif sys.platform == "win32":
                self.driver_type = "win32com"
            else:
                self.driver_type = "mock"

    def app_name(self):
        name = self.excel_path or "Microsoft Excel"
        if name.endswith(".app"):
            name = os.path.splitext(os.path.basename(name))[0]
        return name

    def restart(self):
        if sys.platform == "darwin":
            subprocess.run(
                ["killall", "Microsoft Excel"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
            time.sleep(1.0)
        elif sys.platform == "win32":
            subprocess.run(
                ["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
            time.sleep(1.0)

    def run(self, input_file, output_file, cases):
        shutil.copyfile(input_file, output_file)
        abs_output = os.path.abspath(output_file)
        if self.driver_type == "mock":
            print(
                "[ExcelDisplayDriver Warning] Running in mock mode (Excel not invoked)."
            )
            return {}
        if self.driver_type == "applescript":
            return self.run_applescript(abs_output, cases)
        if self.driver_type == "win32com":
            return self.run_win32com(abs_output, cases)
        raise RuntimeError(f"unknown Excel driver {self.driver_type!r}")

    def run_applescript(self, abs_output, cases):
        read_lines = [
            f'    set acc to acc & "{case["cell"]}=" & (string value of range "{case["cell"]}" of ws) & linefeed'
            for case in cases
        ]
        script = "\n".join(
            [
                'set acc to ""',
                f'tell application "{self.app_name()}"',
                "    set display alerts to false",
                "    try",
                "        close workbooks saving no",
                "    end try",
                f'    open POSIX file "{abs_output}"',
                "    calculate",
                "    set wb to active workbook",
                '    set ws to worksheet "Sheet1" of wb',
                *read_lines,
                "    save wb",
                "    close wb saving no",
                "end tell",
                "return acc",
            ]
        )
        last_err = None
        for attempt in range(3):
            try:
                res = subprocess.run(
                    ["osascript", "-e", script],
                    capture_output=True,
                    text=True,
                    timeout=self.timeout,
                    check=False,
                )
            except subprocess.TimeoutExpired as exc:
                last_err = exc
                self.restart()
                continue
            if res.returncode == 0:
                return parse_display_lines(res.stdout)
            last_err = RuntimeError(res.stderr.strip())
            self.restart()
        raise RuntimeError(f"Excel AppleScript failed: {last_err}")

    def run_win32com(self, abs_output, cases):
        try:
            import win32com.client
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 (win32com) is required for Excel automation on Windows."
            ) from exc

        last_err = None
        for attempt in range(3):
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(abs_output)
                excel.Calculate()
                ws = wb.Worksheets("Sheet1")
                out = {case["cell"]: str(ws.Range(case["cell"]).Text) for case in cases}
                wb.Save()
                wb.Close(False)
                return out
            except Exception as exc:
                last_err = exc
            finally:
                try:
                    excel.Quit()
                except Exception:
                    pass
            self.restart()
        raise last_err


def parse_display_lines(text):
    out = {}
    for line in text.splitlines():
        if "=" not in line:
            continue
        addr, _, value = line.partition("=")
        out[addr.strip()] = value.rstrip("\r")
    return out


def read_visi_displays(path, cases):
    wb = visi_core.Workbook.load(path)
    wb.evaluate()
    return {case["cell"]: wb.get_display(*a1_to_rc(case["cell"])) for case in cases}, wb


def a1_to_rc(addr):
    letters = ""
    digits = ""
    for ch in addr:
        if ch.isalpha():
            letters += ch.upper()
        elif ch.isdigit():
            digits += ch
    col = 0
    for ch in letters:
        col = col * 26 + ord(ch) - ord("A") + 1
    return int(digits) - 1, col - 1


def compare_displays(visi_displays, excel_displays, cases):
    failures = []
    for case in cases:
        addr = case["cell"]
        visi_text = visi_displays.get(addr, "")
        excel_text = excel_displays.get(addr, "")
        if visi_text != excel_text:
            failures.append(
                {
                    "cell": addr,
                    "kind": case["kind"],
                    "source": case["source"],
                    "format": case["format"],
                    "visi": visi_text,
                    "excel": excel_text,
                }
            )
    return failures


def copy_failure_artifacts(dest, *paths):
    os.makedirs(dest, exist_ok=True)
    for path in paths:
        if path and os.path.exists(path):
            shutil.copyfile(path, os.path.join(dest, os.path.basename(path)))


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, sort_keys=True)
        f.write("\n")


def main():
    parser = argparse.ArgumentParser(
        description="Fuzz-compare visi get_display output with Excel's displayed text for formatted cells."
    )
    parser.add_argument("--excel-path", help="Path to Microsoft Excel app/binary.")
    parser.add_argument(
        "--driver", choices=["auto", "applescript", "win32com", "mock"], default="auto"
    )
    parser.add_argument("--iterations", type=int, default=20)
    parser.add_argument("--seed", type=int)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument("--output-dir", default="./fuzz_results")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    failures_dir = os.path.join(args.output_dir, "cell_format_failures")
    os.makedirs(failures_dir, exist_ok=True)

    excel_driver = ExcelDisplayDriver(
        excel_path=args.excel_path, driver_type=args.driver, timeout=args.timeout
    )
    smoke_mode = excel_driver.driver_type == "mock"

    print("=====================================================================")
    print("        visi Displayed Cell Format Fuzzing Harness")
    print("=====================================================================")
    print(f" Iterations : {args.iterations}")
    print(f" Excel Driver: {excel_driver.driver_type} ({args.excel_path or 'Default'})")
    if smoke_mode:
        print(
            " Mock mode: Excel displayed-text oracle is skipped; only visi rendering is checked."
        )
    print("=====================================================================\n")

    passed = 0
    failed = 0
    start = time.time()

    for i in range(1, args.iterations + 1):
        iter_seed = (
            (args.seed + i) if args.seed is not None else random.randint(1, 1_000_000)
        )
        rng = random.Random(iter_seed)
        temp_dir = tempfile.mkdtemp(prefix=f"cell_fmt_fuzz_{i}_")
        source_xlsx = os.path.join(temp_dir, "source.xlsx")
        visi_out_xlsx = os.path.join(temp_dir, "visi_out.xlsx")
        excel_out_xlsx = os.path.join(temp_dir, "excel_out.xlsx")

        try:
            cases = build_source_workbook(source_xlsx, rng)
            visi_displays, visi_wb = read_visi_displays(source_xlsx, cases)
            visi_wb.save(visi_out_xlsx)
            failures = []
            if not smoke_mode:
                excel_displays = excel_driver.run(source_xlsx, excel_out_xlsx, cases)
                failures = compare_displays(visi_displays, excel_displays, cases)

            if failures:
                failed += 1
                fail_dir = os.path.join(failures_dir, f"fail_iter_{i}_seed_{iter_seed}")
                copy_failure_artifacts(
                    fail_dir, source_xlsx, visi_out_xlsx, excel_out_xlsx
                )
                write_json(os.path.join(fail_dir, "cases.json"), cases)
                write_json(os.path.join(fail_dir, "failures.json"), failures)
                print(
                    f" Iteration {i:3d}/{args.iterations} [FAILED] (Seed: {iter_seed})"
                )
                for item in failures[:10]:
                    print(
                        f"   - {item['cell']}: source={item['source']!r} "
                        f"format={item['format']!r} visi={item['visi']!r} excel={item['excel']!r}"
                    )
                if len(failures) > 10:
                    print(f"   ... {len(failures) - 10} more")
                print(f"   artifacts: {fail_dir}")
            else:
                passed += 1
                print(
                    f" Iteration {i:3d}/{args.iterations} [PASSED] (Seed: {iter_seed})"
                )
        except Exception as exc:
            failed += 1
            fail_dir = os.path.join(failures_dir, f"fail_iter_{i}_seed_{iter_seed}")
            copy_failure_artifacts(fail_dir, source_xlsx, visi_out_xlsx, excel_out_xlsx)
            with open(
                os.path.join(fail_dir, "exception.txt"), "w", encoding="utf-8"
            ) as f:
                f.write(repr(exc))
                f.write("\n")
            print(f" Iteration {i:3d}/{args.iterations} [ERROR] (Seed: {iter_seed})")
            print(f"   {exc!r}")
            print(f"   artifacts: {fail_dir}")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    elapsed = time.time() - start
    print("\n=====================================================================")
    print(f" Fuzzing Completed in {elapsed:.2f}s")
    print(f" Passed : {passed}/{args.iterations}")
    print(f" Failed : {failed}/{args.iterations}")
    print("=====================================================================")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
