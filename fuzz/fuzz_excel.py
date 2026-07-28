#!/usr/bin/env python3
"""
visi vs. Microsoft Excel Differential Fuzzing Test Harness
===========================================================
Generates random .xlsx workbooks containing values and formulas, evaluates them
using both `visi` and actual Microsoft Excel, and performs a cell-by-cell semantic
comparison of evaluated results.

Usage:
    python3 fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20
"""

import argparse
import math
import os
import random
import shutil
import string
import subprocess
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET

# Namespace definitions for OpenXML spreadsheet reading
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

# -----------------------------------------------------------------------------
# 1. Formula & Data Generator
# -----------------------------------------------------------------------------

class ExcelFuzzGenerator:
    """Generates random data grids and formula trees for Excel compatibility testing."""

    FUNCTIONS_SINGLE_NUM = ["ABS", "INT", "SQRT", "ROUND", "ROUNDUP", "ROUNDDOWN"]
    FUNCTIONS_MULTI_NUM = ["SUM", "AVERAGE", "MIN", "MAX", "PRODUCT"]
    FUNCTIONS_LOGIC = ["IF", "AND", "OR", "NOT"]
    FUNCTIONS_TEXT = ["CONCATENATE", "LEFT", "RIGHT", "LEN", "UPPER", "LOWER"]

    def __init__(self, seed=None):
        if seed is not None:
            random.seed(seed)

    def _col_name(self, col_idx):
        """Converts 1-based column index to A1 column letter (1 -> A, 2 -> B, 27 -> AA)."""
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def generate_random_value(self):
        """Generates a random cell input value (number, string, boolean, date, edge case)."""
        choice = random.random()
        if choice < 0.35:
            # Integers
            return random.randint(-100, 100)
        elif choice < 0.60:
            # Floating point numbers (including edge cases)
            if random.random() < 0.1:
                return 0.0
            return round(random.uniform(-500.0, 500.0), random.randint(0, 4))
        elif choice < 0.75:
            # Short strings
            chars = string.ascii_letters + " 123"
            return "".join(random.choice(chars) for _ in range(random.randint(1, 8)))
        elif choice < 0.85:
            # Booleans
            return random.choice([True, False])
        elif choice < 0.95:
            # Empty / None
            return None
        else:
            # Small integers for range indexes
            return random.randint(1, 10)

    def generate_formula(self, current_row, current_col, max_row, max_col):
        """Generates a random formula string referencing existing cells or constants."""
        def random_cell_ref():
            r = random.randint(1, max(1, current_row - 1)) if current_row > 1 else 1
            c = random.randint(1, max_col)
            return f"{self._col_name(c)}{r}"

        def random_range_ref():
            r1 = random.randint(1, max(1, current_row - 1)) if current_row > 1 else 1
            r2 = random.randint(r1, max(1, current_row - 1)) if current_row > 1 else 1
            c1 = random.randint(1, max_col)
            c2 = random.randint(c1, max_col)
            return f"{self._col_name(c1)}{r1}:{self._col_name(c2)}{r2}"

        def gen_expr(depth=0):
            if depth >= 2 or random.random() < 0.4:
                # Leaf node: cell ref or scalar constant
                if random.random() < 0.7:
                    return random_cell_ref()
                else:
                    return str(random.randint(-50, 50))

            fn_type = random.choice(["binary", "multi_num", "single_num", "logic", "text"])

            if fn_type == "binary":
                op = random.choice(["+", "-", "*", "/", "^"])
                left = gen_expr(depth + 1)
                right = gen_expr(depth + 1)
                return f"({left} {op} {right})"

            elif fn_type == "multi_num":
                fn = random.choice(self.FUNCTIONS_MULTI_NUM)
                if random.random() < 0.6:
                    arg = random_range_ref()
                else:
                    arg = f"{gen_expr(depth + 1)}, {gen_expr(depth + 1)}"
                return f"{fn}({arg})"

            elif fn_type == "single_num":
                fn = random.choice(self.FUNCTIONS_SINGLE_NUM)
                arg = gen_expr(depth + 1)
                if fn in ["ROUND", "ROUNDUP", "ROUNDDOWN"]:
                    digits = random.randint(0, 2)
                    return f"{fn}({arg}, {digits})"
                return f"{fn}({arg})"

            elif fn_type == "logic":
                if random.random() < 0.5:
                    cond = f"({gen_expr(depth + 1)} > {gen_expr(depth + 1)})"
                    val_true = gen_expr(depth + 1)
                    val_false = gen_expr(depth + 1)
                    return f"IF({cond}, {val_true}, {val_false})"
                else:
                    fn = random.choice(["AND", "OR"])
                    return f"{fn}({gen_expr(depth + 1)} > 0, {gen_expr(depth + 1)} < 100)"

            elif fn_type == "text":
                fn = random.choice(self.FUNCTIONS_TEXT)
                if fn in ["LEFT", "RIGHT"]:
                    return f'{fn}("{gen_expr(depth+1)}", {random.randint(1, 5)})'
                elif fn == "LEN":
                    return f'LEN("{gen_expr(depth+1)}")'
                elif fn in ["UPPER", "LOWER"]:
                    return f'{fn}("{gen_expr(depth+1)}")'
                else:
                    return f'CONCATENATE("{gen_expr(depth+1)}", "{gen_expr(depth+1)}")'

        return "=" + gen_expr(0)

    def create_fuzz_workbook(self, file_path, num_rows=10, num_cols=5):
        """Creates a workbook with a mixture of raw values and formulas."""
        try:
            import openpyxl
        except ImportError:
            print("Error: 'openpyxl' is required for generating .xlsx test files.")
            print("Please run: pip install openpyxl")
            sys.exit(1)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # Populate top rows with raw values
        value_rows = max(2, num_rows // 2)
        for r in range(1, value_rows + 1):
            for c in range(1, num_cols + 1):
                val = self.generate_random_value()
                if val is not None:
                    ws.cell(row=r, column=c, value=val)

        # Populate bottom rows with formulas referencing earlier rows
        for r in range(value_rows + 1, num_rows + 1):
            for c in range(1, num_cols + 1):
                if random.random() < 0.85:
                    formula = self.generate_formula(r, c, num_rows, num_cols)
                    ws.cell(row=r, column=c, value=formula)
                else:
                    val = self.generate_random_value()
                    if val is not None:
                        ws.cell(row=r, column=c, value=val)

        wb.save(file_path)


# -----------------------------------------------------------------------------
# 2. Execution Drivers (visi & Microsoft Excel)
# -----------------------------------------------------------------------------

class VisiDriver:
    """Invokes the `visi` executable to recalculate formulas and save the updated workbook."""
    def __init__(self, binary_path):
        self.binary_path = binary_path
        if not os.path.exists(self.binary_path):
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            rel_path = os.path.join(project_root, "target", "release", "visi")
            dbg_path = os.path.join(project_root, "target", "debug", "visi")
            if os.path.exists(rel_path) and os.path.exists(dbg_path):
                if os.path.getmtime(dbg_path) > os.path.getmtime(rel_path):
                    self.binary_path = dbg_path
                else:
                    self.binary_path = rel_path
            elif os.path.exists(rel_path):
                self.binary_path = rel_path
            elif os.path.exists(dbg_path):
                self.binary_path = dbg_path

    def run(self, input_file, output_file):
        """Runs `visi eval input.xlsx --output output.xlsx`."""
        cmd = [self.binary_path, "eval", input_file, "--output", output_file]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if res.returncode != 0:
            raise RuntimeError(f"visi failed with code {res.returncode}:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}")


class ExcelDriver:
    """Invokes Microsoft Excel to recalculate and save a workbook."""
    def __init__(self, excel_path=None, driver_type="auto"):
        self.excel_path = excel_path
        self.driver_type = driver_type
        if driver_type == "auto":
            if sys.platform == "darwin":
                self.driver_type = "applescript"
            elif sys.platform == "win32":
                self.driver_type = "win32com"
            else:
                self.driver_type = "mock"

    def run(self, input_file, output_file):
        # First copy input_file to output_file so Excel modifies output_file
        shutil.copyfile(input_file, output_file)
        abs_output = os.path.abspath(output_file)

        if self.driver_type == "mock":
            # For testing the harness when Excel is not available
            print("[ExcelDriver Warning] Running in mock mode (Excel not invoked).")
            return

        elif self.driver_type == "applescript":
            # macOS AppleScript Excel recalculation driver
            app_name = self.excel_path if self.excel_path else "Microsoft Excel"
            script = f'''
            set posixPath to "{abs_output}"

            -- use Finder to set the file so that we skip the "Grant access" dialog
            tell application "Finder"
                set theFile to (POSIX file posixPath) as alias
            end tell

            tell application "{app_name}"
                set display alerts to false
                try
                    open file theFile
                    calculate
                    save active workbook
                    close active workbook saving no
                on error errText
                    try
                        close active workbook saving no
                    end try
                    error errText
                end try
            end tell
            '''
            res = None
            for attempt in range(5):
                time.sleep(0.5)
                res = subprocess.run(["osascript", "-e", script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                if res.returncode == 0:
                    break
            if res is not None and res.returncode != 0:
                raise RuntimeError(f"Excel AppleScript failed:\nSTDERR: {res.stderr}")

        elif self.driver_type == "win32com":
            # Windows COM Excel driver
            try:
                import win32com.client
            except ImportError:
                raise RuntimeError("pywin32 (win32com) is required for Excel automation on Windows.")

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(abs_output)
                excel.Calculate()
                wb.Save()
                wb.Close()
            finally:
                excel.Quit()

        elif self.driver_type == "cli":
            # Custom CLI executable driver specified by user
            cmd = [self.excel_path, abs_output]
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if res.returncode != 0:
                raise RuntimeError(f"Excel CLI command failed:\nSTDERR: {res.stderr}")


# -----------------------------------------------------------------------------
# 3. OpenXML Evaluated Content Extractor & Evaluated Value Reader
# -----------------------------------------------------------------------------

class XLSXEvaluatedReader:
    """Parses raw OpenXML structure directly from .xlsx zip files to read cached evaluated cell values."""

    @staticmethod
    def read_evaluated_cells(file_path):
        """
        Returns a dict mapping (sheet_name, cell_ref) -> dict of:
        {
            'raw_value': str,
            'formula': str or None,
            'type': str ('n', 's', 'b', 'e', 'str'),
            'normalized_val': parsed python object
        }
        """
        results = {}

        if not os.path.exists(file_path):
            return results

        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                # 1. Load shared strings table if present
                shared_strings = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    with z.open('xl/sharedStrings.xml') as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        for si in root.findall('.//main:si', NS):
                            # String value can be direct <t> or nested in <r><t>
                            t_elems = si.findall('.//main:t', NS)
                            text = "".join(t.text or "" for t in t_elems)
                            shared_strings.append(text)

                # 2. Iterate sheet XML files
                sheet_files = [name for name in z.namelist() if name.startswith('xl/worksheets/sheet') and name.endswith('.xml')]
                for sheet_file in sorted(sheet_files):
                    sheet_name = os.path.basename(sheet_file).replace('.xml', '')
                    with z.open(sheet_file) as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        for cell in root.findall('.//main:c', NS):
                            ref = cell.attrib.get('r')
                            cell_type = cell.attrib.get('t', 'n')  # default numeric

                            f_elem = cell.find('main:f', NS)
                            formula = f_elem.text if f_elem is not None else None

                            v_elem = cell.find('main:v', NS)
                            raw_val = v_elem.text if v_elem is not None else None

                            normalized_val = XLSXEvaluatedReader._normalize_cell(cell_type, raw_val, shared_strings)

                            results[(sheet_name, ref)] = {
                                'cell_ref': ref,
                                'sheet': sheet_name,
                                'type': cell_type,
                                'raw_value': raw_val,
                                'formula': formula,
                                'val': normalized_val
                            }
        except Exception as e:
            print(f"Warning: Failed to read OpenXML from {file_path}: {e}")

        return results

    @staticmethod
    def _normalize_cell(cell_type, raw_val, shared_strings):
        if raw_val is None:
            return None

        if cell_type == 's':
            # Shared string index
            try:
                idx = int(raw_val)
                if 0 <= idx < len(shared_strings):
                    return shared_strings[idx]
                return f"[Unknown string index {idx}]"
            except ValueError:
                return raw_val

        elif cell_type == 'b':
            # Boolean
            return raw_val == '1' or raw_val.lower() == 'true'

        elif cell_type == 'e':
            # Error code (e.g. #DIV/0!, #VALUE!, #N/A, #REF!, #NUM!, #NAME?, #NULL!)
            return raw_val.upper()

        elif cell_type == 'str' or cell_type == 'inlineStr':
            # Inline string
            return raw_val

        else:
            # Numeric (int or float)
            try:
                f_val = float(raw_val)
                if f_val.is_integer():
                    return int(f_val)
                return f_val
            except ValueError:
                return raw_val


# -----------------------------------------------------------------------------
# 4. Semantic Comparison Engine
# -----------------------------------------------------------------------------

class DifferentialComparator:
    """Compares evaluated cell contents between visi output and Excel output."""

    EXCEL_ERRORS = {"#DIV/0!", "#VALUE!", "#N/A", "#REF!", "#NUM!", "#NAME?", "#NULL!", "#CALC!", "#SPILL!"}

    def __init__(self, float_rel_tol=1e-7, float_abs_tol=1e-7):
        self.float_rel_tol = float_rel_tol
        self.float_abs_tol = float_abs_tol

    def compare(self, visi_cells, excel_cells):
        """
        Compares two cell dictionaries.
        Returns (is_match, mismatches)
        """
        all_keys = set(visi_cells.keys()).union(set(excel_cells.keys()))
        mismatches = []

        for key in sorted(all_keys):
            v_cell = visi_cells.get(key)
            e_cell = excel_cells.get(key)

            if v_cell is None and e_cell is not None:
                if e_cell['val'] is not None:
                    mismatches.append({
                        'key': key,
                        'reason': 'Missing in visi output',
                        'visi': None,
                        'excel': e_cell['val'],
                        'formula': e_cell.get('formula')
                    })
                continue

            if e_cell is None and v_cell is not None:
                if v_cell['val'] is not None:
                    mismatches.append({
                        'key': key,
                        'reason': 'Missing in Excel output',
                        'visi': v_cell['val'],
                        'excel': None,
                        'formula': v_cell.get('formula')
                    })
                continue

            v_val = v_cell['val']
            e_val = e_cell['val']
            formula = v_cell.get('formula') or e_cell.get('formula')

            if not self.values_equal(v_val, e_val):
                mismatches.append({
                    'key': key,
                    'reason': f"Value mismatch ({type(v_val).__name__} vs {type(e_val).__name__})",
                    'visi': v_val,
                    'excel': e_val,
                    'formula': formula
                })

        return len(mismatches) == 0, mismatches

    def values_equal(self, v1, v2):
        """Checks equality between two evaluated values with floating-point tolerance."""
        if v1 is None and v2 is None:
            return True
        if v1 is None or v2 is None:
            if (v1 is None and isinstance(v2, str) and not v2.strip()) or \
               (v2 is None and isinstance(v1, str) and not v1.strip()):
                return True
            return False

        # If both are numbers (float or int)
        if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
            return math.isclose(float(v1), float(v2), rel_tol=self.float_rel_tol, abs_tol=self.float_abs_tol)

        # If both are error strings or text strings
        if isinstance(v1, str) and isinstance(v2, str):
            if v1.upper() in self.EXCEL_ERRORS or v2.upper() in self.EXCEL_ERRORS:
                return v1.upper() == v2.upper()
            return v1.strip() == v2.strip()

        # Booleans vs strings/numbers (e.g. True vs 1, "TRUE" vs True, "FALSE" vs False)
        if isinstance(v1, bool) or isinstance(v2, bool):
            def to_b(v):
                if isinstance(v, bool):
                    return v
                if isinstance(v, str):
                    return v.upper() in ("TRUE", "1")
                if isinstance(v, (int, float)):
                    return v != 0
                return bool(v)
            return to_b(v1) == to_b(v2)

        return str(v1) == str(v2)


# -----------------------------------------------------------------------------
# 5. CLI & Test Runner Orchestrator
# -----------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Differential fuzzing test harness for visi vs Microsoft Excel.")
    parser.add_argument("--excel-path", help="Path to Microsoft Excel binary or application bundle (e.g. '/Applications/Microsoft Excel.app').")
    parser.add_argument("--driver", choices=["auto", "applescript", "win32com", "cli", "mock"], default="auto", help="Excel execution driver.")
    parser.add_argument("--visi-path", default="./target/release/visi", help="Path to compiled visi binary.")
    parser.add_argument("--iterations", type=int, default=10, help="Number of fuzz iterations to run.")
    parser.add_argument("--rows", type=int, default=10, help="Number of rows per sheet.")
    parser.add_argument("--cols", type=int, default=5, help="Number of columns per sheet.")
    parser.add_argument("--seed", type=int, default=None, help="Random seed for reproducible fuzzing.")
    parser.add_argument("--output-dir", default="./fuzz_results", help="Directory to store test outputs and failure artifacts.")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    failures_dir = os.path.join(args.output_dir, "failures")
    os.makedirs(failures_dir, exist_ok=True)

    generator = ExcelFuzzGenerator(seed=args.seed)
    visi_driver = VisiDriver(binary_path=args.visi_path)
    excel_driver = ExcelDriver(excel_path=args.excel_path, driver_type=args.driver)
    comparator = DifferentialComparator()

    print("=====================================================================")
    print("        visi vs. Microsoft Excel Differential Fuzzing Harness       ")
    print("=====================================================================")
    print(f" Iterations : {args.iterations}")
    print(f" Grid Size  : {args.rows}x{args.cols}")
    print(f" Visi Path  : {visi_driver.binary_path}")
    print(f" Excel Driver: {excel_driver.driver_type} ({args.excel_path or 'Default'})")
    print("=====================================================================\n")

    passed_count = 0
    failed_count = 0
    start_time = time.time()

    for i in range(1, args.iterations + 1):
        iter_seed = (args.seed + i) if args.seed is not None else random.randint(1, 1000000)
        iter_gen = ExcelFuzzGenerator(seed=iter_seed)

        temp_dir = tempfile.mkdtemp(prefix=f"fuzz_iter_{i}_")
        source_xlsx = os.path.join(temp_dir, "source.xlsx")
        visi_out_xlsx = os.path.join(temp_dir, "visi_out.xlsx")
        excel_out_xlsx = os.path.join(temp_dir, "excel_out.xlsx")

        try:
            # 1. Generate source workbook
            iter_gen.create_fuzz_workbook(source_xlsx, num_rows=args.rows, num_cols=args.cols)

            # 2. Evaluate with visi
            visi_driver.run(source_xlsx, visi_out_xlsx)

            # 3. Evaluate with Excel
            excel_driver.run(source_xlsx, excel_out_xlsx)

            # 4. Compare evaluated cell contents
            visi_cells = XLSXEvaluatedReader.read_evaluated_cells(visi_out_xlsx)
            excel_cells = XLSXEvaluatedReader.read_evaluated_cells(excel_out_xlsx)

            is_match, mismatches = comparator.compare(visi_cells, excel_cells)

            if is_match:
                passed_count += 1
                print(f" Iteration {i:3d}/{args.iterations} [PASSED] (Seed: {iter_seed})")
            else:
                failed_count += 1
                print(f"\n Iteration {i:3d}/{args.iterations} [FAILED] (Seed: {iter_seed})")
                print(f"   Found {len(mismatches)} cell mismatch(es):")
                for m in mismatches[:5]:  # Print first 5 mismatches
                    print(f"   - Cell {m['key'][1]} on {m['key'][0]}: visi={m['visi']} | Excel={m['excel']} (Formula: {m['formula']})")

                # Save failure artifact
                fail_case_dir = os.path.join(failures_dir, f"fail_iter_{i}_seed_{iter_seed}")
                shutil.copytree(temp_dir, fail_case_dir)
                print(f"   Saved failure reproducing files to: {fail_case_dir}\n")

        except Exception as err:
            failed_count += 1
            print(f"\n Iteration {i:3d}/{args.iterations} [ERROR]: {err}")
            fail_case_dir = os.path.join(failures_dir, f"error_iter_{i}_seed_{iter_seed}")
            if os.path.exists(temp_dir):
                shutil.copytree(temp_dir, fail_case_dir)

        finally:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

    duration = time.time() - start_time
    print("\n=====================================================================")
    print(f" Fuzzing Completed in {duration:.2f}s")
    print(f" Passed : {passed_count}/{args.iterations}")
    print(f" Failed : {failed_count}/{args.iterations}")
    print("=====================================================================")

    if failed_count > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
