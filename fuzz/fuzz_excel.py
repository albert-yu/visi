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

    FUNCTIONS_SINGLE_NUM = [
        "ABS", "INT", "SQRT", "ROUND", "ROUNDUP", "ROUNDDOWN",
        "GAUSS", "PHI", "FISHER", "FISHERINV", "GAMMALN", "GAMMA",
        "NORM.S.DIST", "NORM.S.INV", "ACOSH", "ACOT", "ACOTH",
        "ASINH", "ATANH", "COSH", "COT", "COTH", "CSC", "CSCH",
        "DEGREES", "EVEN", "FACT", "FACTDOUBLE", "ODD", "RADIANS",
        "SEC", "SECH", "SIGN", "SINH", "SQRTPI", "TANH"
    ]
    FUNCTIONS_MULTI_NUM = [
        "SUM", "AVERAGE", "MIN", "MAX", "PRODUCT",
        "AVEDEV", "AVERAGEA", "DEVSQ", "GEOMEAN", "HARMEAN",
        "MEDIAN", "MODE.SNGL", "VAR.S", "VAR.P", "VARA", "VARPA",
        "STDEV.S", "STDEV.P", "STDEVA", "STDEVPA", "SKEW", "SKEW.P",
        "KURT", "MAXA", "MINA", "GCD", "LCM", "MULTINOMIAL", "SUMSQ"
    ]
    FUNCTIONS_STAT_BIVARIATE = [
        "CORREL", "PEARSON", "SLOPE", "INTERCEPT", "RSQ", "STEYX",
        "COVARIANCE.P", "COVARIANCE.S"
    ]
    FUNCTIONS_LOGIC = ["IF", "AND", "OR", "NOT"]
    FUNCTIONS_TEXT = [
        "CONCATENATE", "LEFT", "RIGHT", "LEN", "UPPER", "LOWER",
        "ASC", "CLEAN", "CODE", "DBCS", "EXACT", "FIND", "FINDB",
        "LEFTB", "LENB", "MIDB", "REPT", "RIGHTB", "SEARCH", "SEARCHB",
        "SUBSTITUTE", "T", "TEXTAFTER", "TEXTBEFORE", "UNICHAR", "UNICODE"
    ]
    FUNCTIONS_DATE = [
        "DATE", "DAY", "DAYS", "DAYS360", "EDATE", "EOMONTH", "HOUR", "MINUTE",
        "MONTH", "SECOND", "TIME", "WEEKDAY", "WEEKNUM", "YEAR", "YEARFRAC"
    ]
    FUNCTIONS_ENGINEERING = [
        "BIN2DEC", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DELTA", "GESTEP",
        "HEX2DEC", "OCT2DEC", "BITAND", "BITOR", "BITXOR"
    ]
    FUNCTIONS_INFO = ["ISEVEN", "ISODD", "ISLOGICAL", "ISNONTEXT", "TYPE", "XOR"]

    # Scalar-argument TVM/depreciation functions. Unlike the generic
    # FUNCTIONS_* lists above, financial functions can't have arbitrary
    # sub-expressions substituted into their arguments (a rate must stay
    # small and positive, a period must stay within [1, nper], etc.), so
    # they get their own generator methods below instead of feeding into
    # gen_expr's recursive substitution.
    FINANCIAL_FUNCTIONS = [
        "PV", "FV", "PMT", "NPER", "RATE", "IPMT", "PPMT", "CUMIPMT",
        "CUMPRINC", "NPV", "IRR", "MIRR", "XNPV", "XIRR", "SLN", "SYD",
        "DB", "DDB", "VDB", "EFFECT", "NOMINAL", "DOLLARDE", "DOLLARFR",
        "FVSCHEDULE", "RRI", "PDURATION", "ISPMT",
    ]

    def __init__(self, seed=None):
        if seed is not None:
            random.seed(seed)
        # Populated by create_fuzz_workbook() once it defines an Excel Table
        # over the value grid; used by generate_formula() to emit structured
        # references. Left unset (no structured refs generated) if a caller
        # invokes generate_formula() without going through
        # create_fuzz_workbook() first.
        self._table_name = None
        self._table_cols = []
        # Populated by create_fuzz_workbook()'s financial data block; used
        # by generate_financial_formula() for the array-argument functions
        # (NPV/IRR/MIRR/XNPV/XIRR/FVSCHEDULE). visi's parser has no `{...}`
        # array-literal syntax, so those functions must reference real
        # ranges rather than inline arrays.
        self._fin_cash_range = None
        self._fin_date_range = None
        self._fin_schedule_range = None

    def _col_name(self, col_idx):
        """Converts 1-based column index to A1 column letter (1 -> A, 2 -> B, 27 -> AA)."""
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _has_table(self):
        return bool(self._table_name and self._table_cols)

    def _random_structured_col_ref(self):
        """A structured reference to one table column's whole data body,
        e.g. `Sheet1[A]`. Evaluates to an array, exactly like a range
        reference (A1:A10), so it's only ever used where a range reference
        would also be valid (as a whole argument to an aggregate function).

        visi treats a structured column reference as spanning the *entire*
        column of the sheet, not just a table's declared row range -- so if
        table columns could also hold formulas, this could create a
        dependency cycle (directly, or transitively through another
        formula). create_fuzz_workbook keeps the table's columns disjoint
        from the columns formulas are written into and never puts a
        formula in a table column, which rules that out entirely rather
        than just making it unlikely.
        """
        col_name = random.choice(self._table_cols)
        return f"{self._table_name}[{col_name}]"

    def _random_structured_header_ref(self):
        """A structured reference to a single column's header text, e.g.
        `Sheet1[[#Headers],[A]]`. Evaluates to a plain scalar string, so
        (unlike a bare column reference) it's safe to use anywhere a cell
        reference or constant would be used."""
        col_name = random.choice(self._table_cols)
        return f"{self._table_name}[[#Headers],[{col_name}]]"

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

    def generate_formula(self, current_row, current_col, max_row, max_col, min_col=1):
        """Generates a random formula string referencing existing cells or constants."""
        def random_cell_ref():
            r = random.randint(1, max(1, current_row - 1)) if current_row > 1 else 1
            c = random.randint(min_col, max_col)
            return f"{self._col_name(c)}{r}"

        def random_range_ref():
            r1 = random.randint(1, max(1, current_row - 1)) if current_row > 1 else 1
            r2 = random.randint(r1, max(1, current_row - 1)) if current_row > 1 else 1
            c1 = random.randint(min_col, max_col)
            c2 = random.randint(c1, max_col)
            return f"{self._col_name(c1)}{r1}:{self._col_name(c2)}{r2}"

        def gen_expr(depth=0):
            if depth >= 2 or random.random() < 0.4:
                # Leaf node: cell ref, scalar constant, or (if a table is
                # present) a structured reference to a column's header text.
                # The header ref evaluates to a plain scalar string, just
                # like a cell ref would, so it can drop in anywhere.
                roll = random.random()
                if self._has_table() and roll < 0.15:
                    return self._random_structured_header_ref()
                remaining = (roll - 0.15) / 0.85 if self._has_table() else roll
                if remaining < 0.7:
                    return random_cell_ref()
                else:
                    return str(random.randint(-50, 50))

            fn_type = random.choice(["binary", "multi_num", "single_num", "logic", "text", "stat_bivariate"])

            if fn_type == "binary":
                op = random.choice(["+", "-", "*", "/", "^"])
                left = gen_expr(depth + 1)
                right = gen_expr(depth + 1)
                return f"({left} {op} {right})"

            elif fn_type == "multi_num":
                fn = random.choice(self.FUNCTIONS_MULTI_NUM)
                roll = random.random()
                if self._has_table() and roll < 0.3:
                    # Single-column structured reference, e.g. SUM(Sheet1[A]).
                    arg = self._random_structured_col_ref()
                elif roll < 0.70:
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
                elif fn in ["NORM.S.DIST"]:
                    return f"{fn}({arg}, TRUE)"
                return f"{fn}({arg})"

            elif fn_type == "stat_bivariate":
                fn = random.choice(self.FUNCTIONS_STAT_BIVARIATE)
                r1 = random_range_ref()
                r2 = random_range_ref()
                return f"{fn}({r1}, {r2})"

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
                    return f'{fn}({gen_expr(depth+1)}, {random.randint(1, 5)})'
                elif fn == "LEN":
                    return f'LEN({gen_expr(depth+1)})'
                elif fn in ["UPPER", "LOWER"]:
                    return f'{fn}({gen_expr(depth+1)})'
                else:
                    return f'CONCATENATE({gen_expr(depth+1)}, {gen_expr(depth+1)})'

        return "=" + gen_expr(0)

    # -- Financial function argument helpers -----------------------------
    def _fin_rate(self, lo=0.001, hi=0.03):
        """Per-period rate. Deliberately realistic (0.1%-3%), not the
        0.5%-20% range an earlier version used: at high per-period rates
        compounded over hundreds of periods (nper goes up to 360 below),
        (1+rate)^nper explodes into territory where computing the
        amortizing payment itself loses essentially all f64 precision --
        confirmed against arbitrary-precision decimal arithmetic while
        chasing real mismatches this generator surfaced against actual
        Excel (see finance.rs's ppmt test for the worked example). That's
        a real f64 floor no closed-form or iterative rewrite escapes, not
        a bug -- so the fix here is to stop generating inputs no real
        financial instrument would ever have anyway.
        """
        return f"{round(random.uniform(lo, hi), 4)}"

    def _fin_money(self, lo=100, hi=50000, allow_negative=True):
        v = round(random.uniform(lo, hi), 2)
        if allow_negative and random.random() < 0.5:
            v = -v
        return f"{v}"

    def _fin_int(self, lo, hi):
        return str(random.randint(lo, hi))

    def _fin_type01(self):
        return str(random.choice([0, 1]))

    def _fin_money_value(self, lo=100, hi=20000, allow_negative=True):
        v = round(random.uniform(lo, hi), 2)
        if allow_negative and random.random() < 0.5:
            v = -v
        return v

    def generate_financial_formula(self):
        """Generates a single self-contained financial-function formula
        with semantically valid inputs (small positive rates, periods
        within range, etc.) rather than composing arbitrary sub-expressions
        the way generate_formula() does -- most financial arguments have a
        specific meaning (a rate, a period count) that a random
        sub-expression would violate far too often to be a useful test."""
        fn = random.choice(self.FINANCIAL_FUNCTIONS)

        if fn in ("PV", "FV", "PMT"):
            rate = self._fin_rate()
            nper = self._fin_int(1, 360)
            middle = self._fin_money(10, 5000)
            other = self._fin_money(0, 5000) if random.random() < 0.6 else "0"
            typ = self._fin_type01()
            return f"={fn}({rate}, {nper}, {middle}, {other}, {typ})"

        if fn == "NPER":
            rate = self._fin_rate()
            pmt = self._fin_money(10, 2000, allow_negative=False)
            pv = self._fin_money(1000, 20000, allow_negative=False)
            typ = self._fin_type01()
            return f"=NPER({rate}, -{pmt}, {pv}, 0, {typ})"

        if fn == "RATE":
            nper = self._fin_int(6, 360)
            pmt = self._fin_money(10, 2000, allow_negative=False)
            pv = self._fin_money(1000, 20000, allow_negative=False)
            typ = self._fin_type01()
            return f"=RATE({nper}, -{pmt}, {pv}, 0, {typ})"

        if fn in ("IPMT", "PPMT"):
            rate = self._fin_rate()
            nper = self._fin_int(2, 360)
            per = self._fin_int(1, int(nper))
            pv = self._fin_money(1000, 100000, allow_negative=False)
            typ = self._fin_type01()
            return f"={fn}({rate}, {per}, {nper}, {pv}, 0, {typ})"

        if fn in ("CUMIPMT", "CUMPRINC"):
            rate = self._fin_rate()
            nper = int(self._fin_int(12, 360))
            start = random.randint(1, nper)
            end = random.randint(start, nper)
            pv = self._fin_money(1000, 100000, allow_negative=False)
            typ = self._fin_type01()
            return f"={fn}({rate}, {nper}, {pv}, {start}, {end}, {typ})"

        # NPV/IRR/MIRR/XNPV/XIRR/FVSCHEDULE take an array argument. visi's
        # parser has no `{...}` array-literal syntax, so these always
        # reference the financial data block create_fuzz_workbook lays out
        # before calling this method -- there is no standalone-call path
        # for this generator, so no inline-array fallback is needed.
        if fn == "NPV":
            assert self._fin_cash_range
            return f"=NPV({self._fin_rate()}, {self._fin_cash_range})"

        if fn == "IRR":
            assert self._fin_cash_range
            return f"=IRR({self._fin_cash_range})"

        if fn == "MIRR":
            assert self._fin_cash_range
            return f"=MIRR({self._fin_cash_range}, {self._fin_rate()}, {self._fin_rate()})"

        if fn == "XNPV":
            assert self._fin_cash_range and self._fin_date_range
            return f"=XNPV({self._fin_rate()}, {self._fin_cash_range}, {self._fin_date_range})"

        if fn == "XIRR":
            assert self._fin_cash_range and self._fin_date_range
            return f"=XIRR({self._fin_cash_range}, {self._fin_date_range})"

        if fn in ("SLN",):
            cost = self._fin_money(1000, 50000, allow_negative=False)
            salvage = self._fin_money(0, 999, allow_negative=False)
            life = self._fin_int(1, 30)
            return f"=SLN({cost}, {salvage}, {life})"

        if fn == "SYD":
            cost = self._fin_money(1000, 50000, allow_negative=False)
            salvage = self._fin_money(0, 999, allow_negative=False)
            life = int(self._fin_int(1, 30))
            per = self._fin_int(1, life)
            return f"=SYD({cost}, {salvage}, {life}, {per})"

        if fn == "DB":
            cost = self._fin_money(1000, 50000, allow_negative=False)
            salvage = self._fin_money(1, 999, allow_negative=False)
            life = int(self._fin_int(1, 20))
            period = self._fin_int(1, life)
            month = self._fin_int(1, 12)
            return f"=DB({cost}, {salvage}, {life}, {period}, {month})"

        if fn == "DDB":
            cost = self._fin_money(1000, 50000, allow_negative=False)
            salvage = self._fin_money(1, 999, allow_negative=False)
            life = int(self._fin_int(1, 20))
            period = self._fin_int(1, life)
            factor = round(random.uniform(1.0, 3.0), 2)
            return f"=DDB({cost}, {salvage}, {life}, {period}, {factor})"

        if fn == "VDB":
            cost = self._fin_money(1000, 50000, allow_negative=False)
            salvage = self._fin_money(1, 999, allow_negative=False)
            life = int(self._fin_int(1, 20))
            start = random.randint(0, life - 1) if life > 1 else 0
            end = random.randint(start + 1, life)
            factor = round(random.uniform(1.0, 3.0), 2)
            no_switch = random.choice(["TRUE", "FALSE"])
            return f"=VDB({cost}, {salvage}, {life}, {start}, {end}, {factor}, {no_switch})"

        if fn == "EFFECT":
            nominal_rate = self._fin_rate()
            npery = self._fin_int(1, 12)
            return f"=EFFECT({nominal_rate}, {npery})"

        if fn == "NOMINAL":
            effect_rate = self._fin_rate()
            npery = self._fin_int(1, 12)
            return f"=NOMINAL({effect_rate}, {npery})"

        if fn in ("DOLLARDE", "DOLLARFR"):
            dollar = round(random.uniform(1.0, 100.0), random.choice([1, 2]))
            fraction = random.choice([2, 4, 8, 16, 32, 64])
            return f"={fn}({dollar}, {fraction})"

        if fn == "FVSCHEDULE":
            assert self._fin_schedule_range
            principal = self._fin_money(1000, 50000, allow_negative=False)
            return f"=FVSCHEDULE({principal}, {self._fin_schedule_range})"

        # RRI and PDURATION were both added in Excel 2013; real Excel's own
        # OOXML writer always stores post-2007 functions with an `_xlfn.`
        # prefix in the underlying formula text (invisible in the formula
        # bar) so older parsers degrade gracefully instead of choking on an
        # unknown name -- confirmed as the actual cause of a real `#NAME?`
        # mismatch this generator produced: openpyxl writes plain
        # `RRI(...)`/`PDURATION(...)` with no prefix, and Excel refuses to
        # recognize the bare name on load. visi already strips a leading
        # `_xlfn.` before dispatch (see evaluate_function in sheet.rs), so
        # writing the prefix here is what a real xlsx producer would do and
        # keeps both sides consistent.
        if fn == "RRI":
            nper = self._fin_int(1, 30)
            pv = self._fin_money(1000, 50000, allow_negative=False)
            fv = self._fin_money(1000, 100000, allow_negative=False)
            return f"=_xlfn.RRI({nper}, {pv}, {fv})"

        if fn == "PDURATION":
            rate = self._fin_rate()
            pv = self._fin_money(1000, 50000, allow_negative=False)
            fv = self._fin_money(1000, 100000, allow_negative=False)
            return f"=_xlfn.PDURATION({rate}, {pv}, {fv})"

        if fn == "ISPMT":
            rate = self._fin_rate()
            nper = int(self._fin_int(2, 360))
            per = self._fin_int(1, nper)
            pv = self._fin_money(1000, 100000, allow_negative=False)
            return f"=ISPMT({rate}, {per}, {nper}, {pv})"

        raise AssertionError(f"no generator wired up for financial function {fn}")

    def create_fuzz_workbook(self, file_path, num_rows=10, num_cols=5):
        """Creates a workbook with a mixture of raw values and formulas, plus
        a real Excel Table for structured references to resolve against.

        The sheet is laid out as two disjoint blocks of `num_cols` columns
        side by side: a "table" block (pure random values, every row) and a
        "formula" block (values on top, generated formulas below, exactly
        like the original single-block layout). They're kept disjoint
        because visi treats a structured reference to a table column as
        spanning that column's *entire* height, not just the table's
        declared row range -- if a table column could also hold a formula,
        a structured reference into it could create a dependency cycle
        (directly, or transitively through another formula). Since the
        table block never contains a formula, that's impossible by
        construction rather than merely unlikely.
        """
        try:
            import openpyxl
            from openpyxl.worksheet.table import Table, TableStyleInfo
        except ImportError:
            print("Error: 'openpyxl' is required for generating .xlsx test files.")
            print("Please run: pip install openpyxl")
            sys.exit(1)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # --- Table block: columns 1..num_cols. Row 1 holds deterministic
        # column headers ("A", "B", ...); visi has no concept of an Excel
        # Table distinct from the worksheet itself, and on import always
        # names each column by its plain spreadsheet letter
        # (col_idx_to_letters) -- using that same scheme as the real header
        # text means `Sheet1[A]` resolves to the same column in both visi
        # (worksheet/column-name lookup) and Excel (Table lookup). Every
        # other row is a plain random value -- never a formula.
        header_names = [self._col_name(c) for c in range(1, num_cols + 1)]
        for c, name in enumerate(header_names, start=1):
            ws.cell(row=1, column=c, value=name)
        for r in range(2, num_rows + 1):
            for c in range(1, num_cols + 1):
                val = self.generate_random_value()
                if val is not None:
                    ws.cell(row=r, column=c, value=val)

        table_name = ws.title
        self._table_name = table_name
        self._table_cols = header_names
        table_ref = f"A1:{self._col_name(num_cols)}{num_rows}"
        table = Table(displayName=table_name, ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(table)

        # --- Formula block: a second block of num_cols columns immediately
        # to the right of the table. Top rows hold raw random values (for
        # plain cell/range refs to point at); the rest hold generated
        # formulas referencing earlier rows in this same block, optionally
        # including structured references into the table block (see
        # generate_formula / _random_structured_col_ref /
        # _random_structured_header_ref).
        value_rows = max(2, num_rows // 2)
        min_col = num_cols + 1
        max_col = num_cols * 2
        for r in range(1, value_rows + 1):
            for c in range(min_col, max_col + 1):
                val = self.generate_random_value()
                if val is not None:
                    ws.cell(row=r, column=c, value=val)

        for r in range(value_rows + 1, num_rows + 1):
            for c in range(min_col, max_col + 1):
                if random.random() < 0.85:
                    formula = self.generate_formula(r, c, num_rows, max_col, min_col=min_col)
                    ws.cell(row=r, column=c, value=formula)
                else:
                    val = self.generate_random_value()
                    if val is not None:
                        ws.cell(row=r, column=c, value=val)

        # --- Financial data block: a small "cashflow" column, an aligned
        # ascending "date" column (plain Excel serial numbers, since visi
        # has no DATE() function to build one from parts), and a "schedule"
        # column of small rates, all plain values (never formulas). These
        # back the array-argument financial functions (NPV/IRR/MIRR/XNPV/
        # XIRR/FVSCHEDULE) via real ranges, since visi's parser doesn't
        # support `{...}` array literals the way generate_financial_formula
        # falls back to when this block hasn't been laid out.
        fin_cash_col = max_col + 2
        fin_date_col = fin_cash_col + 1
        fin_schedule_col = fin_cash_col + 2
        fin_formula_col = fin_cash_col + 4

        cash_rows = 6
        ws.cell(row=1, column=fin_cash_col, value=-round(random.uniform(5000, 50000), 2))
        for r in range(2, cash_rows + 1):
            ws.cell(row=r, column=fin_cash_col, value=self._fin_money_value())

        date_serial = random.randint(40000, 45000)
        for r in range(1, cash_rows + 1):
            date_serial += random.randint(15, 90)
            ws.cell(row=r, column=fin_date_col, value=date_serial)

        schedule_rows = 3
        for r in range(1, schedule_rows + 1):
            ws.cell(row=r, column=fin_schedule_col, value=round(random.uniform(0.01, 0.15), 4))

        self._fin_cash_range = f"{self._col_name(fin_cash_col)}1:{self._col_name(fin_cash_col)}{cash_rows}"
        self._fin_date_range = f"{self._col_name(fin_date_col)}1:{self._col_name(fin_date_col)}{cash_rows}"
        self._fin_schedule_range = (
            f"{self._col_name(fin_schedule_col)}1:{self._col_name(fin_schedule_col)}{schedule_rows}"
        )

        financial_formula_rows = max(6, num_rows)
        for r in range(1, financial_formula_rows + 1):
            ws.cell(row=r, column=fin_formula_col, value=self.generate_financial_formula())

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
            if app_name.endswith(".app"):
                app_name = os.path.splitext(os.path.basename(app_name))[0]
            script = f'''
            tell application "{app_name}"
                set display alerts to false
                try
                    close workbooks saving no
                end try
                try
                    set targetFile to POSIX file "{abs_output}"
                    open targetFile
                    calculate
                    save active workbook
                    close active workbook saving no
                on error errText number errNum
                    try
                        close workbooks saving no
                    end try
                    error errText number errNum
                end try
            end tell
            '''
            res = None
            for attempt in range(5):
                time.sleep(0.5)
                try:
                    res = subprocess.run(["osascript", "-e", script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=15)
                    if res.returncode == 0:
                        break
                except subprocess.TimeoutExpired:
                    subprocess.run(["killall", "Microsoft Excel"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    time.sleep(1.0)
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

                            if cell_type == 'inlineStr':
                                # Inline strings live in <is><t>...</t></is> (or
                                # <is><r><t>...</t></r>...</is> for rich-text runs),
                                # not <v> -- real Excel writes these for some string
                                # cells instead of using the shared-strings table
                                # (observed via fuzz_pivot.py's AppleScript/VBA
                                # macro save path).
                                is_elem = cell.find('main:is', NS)
                                if is_elem is not None:
                                    t_elems = is_elem.findall('.//main:t', NS)
                                    raw_val = "".join(t.text or "" for t in t_elems)
                                else:
                                    raw_val = None
                            else:
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

        # If one is a number and the other is a numeric string (e.g. 0.0394 vs ".0394", 8 vs "08")
        if isinstance(v1, (int, float)) and isinstance(v2, str):
            try:
                return math.isclose(float(v1), float(v2), rel_tol=self.float_rel_tol, abs_tol=self.float_abs_tol)
            except ValueError:
                pass
        if isinstance(v2, (int, float)) and isinstance(v1, str):
            try:
                return math.isclose(float(v1), float(v2), rel_tol=self.float_rel_tol, abs_tol=self.float_abs_tol)
            except ValueError:
                pass

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
                shutil.copytree(temp_dir, fail_case_dir, dirs_exist_ok=True)
                print(f"   Saved failure reproducing files to: {fail_case_dir}\n")

        except Exception as err:
            failed_count += 1
            print(f"\n Iteration {i:3d}/{args.iterations} [ERROR]: {err}")
            fail_case_dir = os.path.join(failures_dir, f"error_iter_{i}_seed_{iter_seed}")
            if os.path.exists(temp_dir):
                shutil.copytree(temp_dir, fail_case_dir, dirs_exist_ok=True)

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
