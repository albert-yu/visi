#!/usr/bin/env python3
import argparse
import math
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

EXCEL_APP = "Microsoft Excel"

try:
    import openpyxl
except ImportError:
    sys.exit(
        "openpyxl is required: source fuzz/venv/bin/activate && pip install -r fuzz/requirements.txt"
    )

try:
    import visi_core
except ImportError:
    sys.exit(
        "the visi_core bindings are required: "
        "maturin develop -m visi-python/Cargo.toml --release"
    )


VAR_NAMES = ["va", "vb", "vc", "vd", "ve"]


HOST_SHEET_VAR = "wsh"
HOST_CELL_VAR = "vk"
HOST_RANGE_VAR = "vq"
_HOST_VARS = (HOST_SHEET_VAR, HOST_CELL_VAR, HOST_RANGE_VAR)
assert not set(_HOST_VARS) & set(VAR_NAMES + ["vi", "vn"]), (
    f"host locals {_HOST_VARS} collide with generated variables"
)


NUM_LITERALS = [
    "0",
    "1",
    "2",
    "3",
    "7",
    "10",
    "-1",
    "-7",
    "255",
    "32767",
    "32768",
    "-32768",
    "100000",
    "2147483647",
    "1.5",
    "-2.5",
    "0.1",
    "3.75",
    "1E3",
    "0.0001",
    "1&",
    "1%",
    "2!",
    "3#",
    "&HFF",
    "&O17",
]
STR_LITERALS = ['""', '"a"', '"abc"', '"1"', '"12"', '"1.5"', '"  3  "', '"Z"']
BOOL_LITERALS = ["True", "False"]
SPECIAL_LITERALS = ["Empty", "Null"]

ARITH_OPS = ["+", "-", "*", "/", "\\", "Mod", "^"]
COMPARE_OPS = ["=", "<>", "<", ">", "<=", ">="]
LOGICAL_OPS = ["And", "Or", "Xor", "Eqv", "Imp"]


UNARY_FUNCS = [
    "CStr",
    "CInt",
    "CLng",
    "CDbl",
    "CSng",
    "CBool",
    "Abs",
    "Sgn",
    "Int",
    "Fix",
    "UCase",
    "LCase",
    "Trim",
    "TypeName",
    "IsNumeric",
    "IsEmpty",
    "IsNull",
    "Val",
    "StrReverse",
]


GRID_ROWS, GRID_COLS = 4, 6
GRID = [
    (1, 1, "1"),
    (2, 1, "2"),
    (3, 1, "3"),
    (4, 1, "4"),
    (1, 2, "10"),
    (2, 2, "20"),
    (3, 2, "30"),
    (4, 2, "40"),
    (1, 3, "=A1*2"),
    (2, 3, "=A2+B2"),
    (1, 4, "hi"),
    (2, 4, True),
]


SCRATCH_COLS = (5, 6)


WRITABLE_FORMULAS = ['"=A1+B1"', '"=SUM(A1:B2)"', '"=A1*3"', '"=COUNT(A1:B4)"']


GRID_FUNCTIONS = ["Sum", "Count", "CountA", "Min", "Max", "Average"]


FONT_NAMES = ['"Calibri"', '"Arial"', '"Courier New"', '"Times New Roman"']
NUMBER_FORMATS = ['"General"', '"0"', '"0.00"', '"@"', '"m/d/yy"']
TABLE_COLUMNS = ['"Region"', '"Amount"', '"Flag"']
TABLE_COLUMN_NAMES = ['"Region2"', '"Amount2"', '"Flag2"', '"Total"']
ADDRESS_ARGS = ["", "(False, False)", "(True, False)", "(False, True)", "(True, True)"]


class VbaGenerator:
    def __init__(self, seed=None, host_surface="basic"):
        self.rng = random.Random(seed)
        self.host_surface = host_surface

    def literal(self):
        bucket = self.rng.random()
        if bucket < 0.55:
            return self.rng.choice(NUM_LITERALS)
        if bucket < 0.78:
            return self.rng.choice(STR_LITERALS)
        if bucket < 0.92:
            return self.rng.choice(BOOL_LITERALS)
        return self.rng.choice(SPECIAL_LITERALS)

    def expr(self, depth, vars_in_scope):
        if depth <= 0:
            if vars_in_scope and self.rng.random() < 0.35:
                return self.rng.choice(vars_in_scope)
            return self.literal()

        kind = self.rng.random()
        if kind < 0.34:
            op = self.rng.choice(ARITH_OPS)
            return f"({self.expr(depth - 1, vars_in_scope)} {op} {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.46:
            op = self.rng.choice(COMPARE_OPS)
            return f"({self.expr(depth - 1, vars_in_scope)} {op} {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.56:
            op = self.rng.choice(LOGICAL_OPS)
            return f"({self.expr(depth - 1, vars_in_scope)} {op} {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.64:
            return f"({self.expr(depth - 1, vars_in_scope)} & {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.72:
            return f"(Not {self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.80:
            return f"(-{self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.90:
            fn = self.rng.choice(UNARY_FUNCS)
            return f"{fn}({self.expr(depth - 1, vars_in_scope)})"
        if kind < 0.94:
            return f"Len(CStr({self.expr(depth - 1, vars_in_scope)}))"

        fn = self.rng.choice(["Left", "Right", "String", "Space", "InStr", "Mid"])
        if fn in ("Space",):
            return f"Space({self.rng.randint(0, 4)})"
        if fn == "String":
            return f'String({self.rng.randint(0, 4)}, "x")'
        if fn == "InStr":
            return f"InStr({self.expr(depth - 1, vars_in_scope)}, {self.rng.choice(STR_LITERALS)})"
        if fn == "Mid":
            return f"Mid({self.expr(depth - 1, vars_in_scope)}, {self.rng.randint(1, 4)}, {self.rng.randint(0, 4)})"
        return f"{fn}({self.expr(depth - 1, vars_in_scope)}, {self.rng.randint(0, 4)})"

    def statements(self, count, vars_in_scope, depth):
        out = []
        for _ in range(count):
            out.extend(self.statement(vars_in_scope, depth))
        return out

    def statement(self, vars_in_scope, depth):
        kind = self.rng.random()
        target = self.rng.choice(vars_in_scope)

        if kind < 0.28:
            return [f"{target} = {self.expr(depth, vars_in_scope)}"]

        if kind < 0.39:
            return [
                f"If {self.expr(depth - 1, vars_in_scope)} Then",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "Else",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "End If",
            ]

        if kind < 0.49:
            lo, hi = 1, self.rng.randint(1, 4)
            step = self.rng.choice(["", " Step 2", " Step -1"])
            if step == " Step -1":
                lo, hi = hi, 1
            return [
                f"For vi = {lo} To {hi}{step}",
                f"    {target} = {self.expr(depth - 1, vars_in_scope + ['vi'])}",
                "Next vi",
            ]

        if kind < 0.56:
            return [
                "vn = 0",
                f"Do While vn < {self.rng.randint(1, 3)}",
                "    vn = vn + 1",
                f"    {target} = {self.expr(depth - 1, vars_in_scope + ['vn'])}",
                "Loop",
            ]

        if kind < 0.63:
            subject = self.expr(depth - 1, vars_in_scope)
            return [
                f"Select Case {subject}",
                "Case 0, 1",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "Case 2 To 5",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "Case Else",
                f"    {target} = {self.expr(depth - 1, vars_in_scope)}",
                "End Select",
            ]

        if kind < 0.70:
            return [
                "On Error Resume Next",
                f"{target} = {self.expr(depth, vars_in_scope)}",
                "On Error GoTo 0",
            ]

        return self.host_statement(target, vars_in_scope, depth)

    def cell(self, scratch=False):
        row = self.rng.randint(1, GRID_ROWS)
        col = (
            self.rng.choice(SCRATCH_COLS) if scratch else self.rng.randint(1, GRID_COLS)
        )
        return row, col

    def a1(self, row, col):
        return f"{chr(ord('A') + col - 1)}{row}"

    def rgb(self):
        return f"RGB({self.rng.randint(0, 255)}, {self.rng.randint(0, 255)}, {self.rng.randint(0, 255)})"

    def address_of(self, obj):
        return f"{obj}.Address{self.rng.choice(ADDRESS_ARGS)}"

    def range_addr(self, scratch=False):
        row, col = self.cell(scratch=scratch)
        row2, col2 = self.cell(scratch=scratch)
        r0, r1 = sorted((row, row2))
        c0, c1 = sorted((col, col2))
        return f"{self.a1(r0, c0)}:{self.a1(r1, c1)}", r0, c0, r1, c1

    def host_statement(self, target, vars_in_scope, depth):
        if self.host_surface == "extended" and self.rng.random() < 0.45:
            return self.extended_host_statement(target, vars_in_scope, depth)

        kind = self.rng.random()
        row, col = self.cell()
        srow, scol = self.cell(scratch=True)

        if kind < 0.22:
            return [
                f"{HOST_SHEET_VAR}.Cells({srow}, {scol}).Value = {self.expr(depth - 1, vars_in_scope)}",
                f"{target} = {HOST_SHEET_VAR}.Cells({srow}, {scol}).Value",
            ]

        if kind < 0.34:
            return [
                f"{HOST_SHEET_VAR}.Cells({self.rng.randint(1, 2)}, 1).Value = {self.rng.choice(NUM_LITERALS)}",
                f'{target} = {HOST_SHEET_VAR}.Range("C1").Value',
            ]

        if kind < 0.44:
            return [f"{target} = {HOST_SHEET_VAR}.Cells({row}, {col}).Value"]

        if kind < 0.52:
            prop = self.rng.choice(
                ["Value2", "Formula", "Text", "Address", "Row", "Column"]
            )
            return [f"{target} = {HOST_SHEET_VAR}.Cells({row}, {col}).{prop}"]

        if kind < 0.60:
            r2 = self.rng.randint(row, GRID_ROWS)
            c2 = self.rng.randint(col, GRID_COLS)
            rng = f'{HOST_SHEET_VAR}.Range("{self.a1(row, col)}:{self.a1(r2, c2)}")'
            prop = self.rng.choice(["Count", "Address", "Row", "Column"])
            return [f"{target} = {rng}.{prop}"]

        if kind < 0.70:
            fn = self.rng.choice(GRID_FUNCTIONS)
            r2 = self.rng.randint(row, GRID_ROWS)
            c2 = self.rng.randint(col, GRID_COLS)
            rng = f'{HOST_SHEET_VAR}.Range("{self.a1(row, col)}:{self.a1(r2, c2)}")'

            via = self.rng.choice(["Application.WorksheetFunction", "Application"])
            return [f"{target} = {via}.{fn}({rng})"]

        if kind < 0.78:
            return [
                (
                    f"{target} = Application.WorksheetFunction.{self.rng.choice(GRID_FUNCTIONS)}"
                    f"({self.rng.choice(NUM_LITERALS)}, {self.rng.choice(NUM_LITERALS)})"
                )
            ]

        if kind < 0.86:
            r2 = self.rng.randint(row, GRID_ROWS)
            return [
                f'For Each {HOST_CELL_VAR} In {HOST_SHEET_VAR}.Range("{self.a1(row, col)}:{self.a1(r2, col)}")',
                f'    {target} = {HOST_CELL_VAR}.Address & "/" & {target}',
                f"Next {HOST_CELL_VAR}",
            ]

        if kind < 0.92:
            return [
                f"{HOST_SHEET_VAR}.Cells({srow}, {scol}).Formula = {self.rng.choice(WRITABLE_FORMULAS)}",
                f"{target} = {HOST_SHEET_VAR}.Cells({srow}, {scol}).Value",
            ]

        if kind < 0.96:
            dr, dc = self.rng.randint(0, 1), self.rng.randint(0, 1)
            return [
                f'With {HOST_SHEET_VAR}.Range("{self.a1(row, col)}")',
                f'    {target} = .Offset({dr}, {dc}).Address & "/" & CStr(.Value)',
                "End With",
            ]

        return [
            f'Set {HOST_RANGE_VAR} = {HOST_SHEET_VAR}.Range("{self.a1(row, col)}")',
            f'{target} = CStr({HOST_RANGE_VAR} Is {HOST_SHEET_VAR}.Range("{self.a1(row, col)}")) & "/" & TypeName({HOST_RANGE_VAR})',
        ]

    def extended_host_statement(self, target, vars_in_scope, depth):
        kind = self.rng.random()
        row, col = self.cell()
        srow, scol = self.cell(scratch=True)
        range_addr, r0, c0, r1, c1 = self.range_addr()
        range_obj = f'{HOST_SHEET_VAR}.Range("{range_addr}")'

        if kind < 0.18:
            style_kind = self.rng.random()
            if style_kind < 0.32:
                target_obj = self.rng.choice(["Interior", "Font"])
                prop = "Color"
                value = self.rgb()
                read_obj = self.rng.choice(
                    [
                        f"{HOST_SHEET_VAR}.Cells({row}, {col})",
                        range_obj,
                    ]
                )
                return [
                    f"{range_obj}.{target_obj}.{prop} = {value}",
                    f"{target} = CStr({read_obj}.{target_obj}.{prop}) & \"/\" & TypeName({read_obj}.{target_obj}.{prop})",
                ]
            if style_kind < 0.50:
                target_obj = self.rng.choice(["Interior", "Font"])
                prop = "ColorIndex"
                value = self.rng.choice(["1", "2", "3", "4", "5", "6", "xlNone"])
                read_obj = self.rng.choice(
                    [
                        f"{HOST_SHEET_VAR}.Cells({row}, {col})",
                        range_obj,
                    ]
                )
                return [
                    f"{range_obj}.{target_obj}.{prop} = {value}",
                    f"{target} = CStr({read_obj}.{target_obj}.{prop}) & \"/\" & TypeName({read_obj}.{target_obj}.{prop})",
                ]
            if style_kind < 0.70:
                prop = self.rng.choice(["Bold", "Italic"])
                value = self.rng.choice(BOOL_LITERALS)
                read_obj = self.rng.choice(
                    [
                        f"{HOST_SHEET_VAR}.Cells({row}, {col})",
                        range_obj,
                    ]
                )
                return [
                    f"{range_obj}.Font.{prop} = {value}",
                    f"{target} = CStr({read_obj}.Font.{prop}) & \"/\" & TypeName({read_obj}.Font.{prop})",
                ]
            if style_kind < 0.84:
                prop = self.rng.choice(["Size", "Name"])
                value = (
                    self.rng.choice(["8", "10.5", "11", "14", "20"])
                    if prop == "Size"
                    else self.rng.choice(FONT_NAMES)
                )
                read_obj = self.rng.choice(
                    [
                        f"{HOST_SHEET_VAR}.Cells({row}, {col})",
                        range_obj,
                    ]
                )
                return [
                    f"{range_obj}.Font.{prop} = {value}",
                    f"{target} = CStr({read_obj}.Font.{prop}) & \"/\" & TypeName({read_obj}.Font.{prop})",
                ]
            value = self.rng.choice(NUMBER_FORMATS)
            read_obj = self.rng.choice(
                [
                    f"{HOST_SHEET_VAR}.Cells({row}, {col})",
                    range_obj,
                ]
            )
            return [
                f"{range_obj}.NumberFormat = {value}",
                f"{target} = {read_obj}.NumberFormat & \"/\" & {read_obj}.Text",
            ]

        if kind < 0.33:
            axis = self.rng.choice(["Rows", "Columns"])
            at = self.rng.randint(1, GRID_ROWS if axis == "Rows" else GRID_COLS)
            op = self.rng.choice(["Insert", "Delete"])
            addr = self.a1(row, col)
            access = self.rng.choice(
                [
                    f"{HOST_RANGE_VAR}.Address{self.rng.choice(ADDRESS_ARGS)}",
                    f"TypeName({HOST_RANGE_VAR})",
                    f"CStr({HOST_RANGE_VAR} Is Nothing)",
                    f"CStr({HOST_RANGE_VAR}.Row) & \"/\" & CStr({HOST_RANGE_VAR}.Column)",
                    f"CStr({HOST_RANGE_VAR}.Value)",
                ]
            )
            if self.rng.random() < 0.55:
                edit_line = f"{HOST_SHEET_VAR}.{axis}({at}).{op}"
            else:
                whole = "EntireRow" if axis == "Rows" else "EntireColumn"
                edit_line = f'{HOST_SHEET_VAR}.Range("{self.a1(at if axis == "Rows" else row, col if axis == "Rows" else at)}").{whole}.{op}'
            return [
                f'Set {HOST_RANGE_VAR} = {HOST_SHEET_VAR}.Range("{addr}")',
                edit_line,
                f'{target} = {access} & "/" & TypeName({HOST_RANGE_VAR})',
            ]

        if kind < 0.46:
            addr = self.a1(row, col)
            dr = self.rng.randint(0, GRID_ROWS - row)
            dc = self.rng.randint(0, GRID_COLS - col)
            height = self.rng.randint(1, GRID_ROWS - row - dr + 1)
            width = self.rng.randint(1, GRID_COLS - col - dc + 1)
            obj = f'{HOST_RANGE_VAR}.Offset({dr}, {dc}).Resize({height}, {width})'
            read = self.rng.choice(
                [
                    self.address_of(obj),
                    f"CStr({obj}.Count)",
                    f"CStr({obj}.Row) & \"/\" & CStr({obj}.Column)",
                    f"TypeName({obj}.ListObject)",
                ]
            )
            return [
                f'Set {HOST_RANGE_VAR} = {HOST_SHEET_VAR}.Range("{addr}")',
                f'{target} = {read} & "/" & CStr({HOST_RANGE_VAR} Is {HOST_SHEET_VAR}.Range("{addr}"))',
            ]

        if kind < 0.58:
            prop = self.rng.choice(["Value", "Value2", "Formula"])
            value = (
                self.rng.choice(WRITABLE_FORMULAS)
                if prop == "Formula"
                else self.expr(depth - 1, vars_in_scope)
            )
            read_prop = self.rng.choice(["Value", "Value2", "Formula", "Text"])
            return [
                f"{HOST_SHEET_VAR}.Cells({srow}, {scol}).{prop} = {value}",
                f"{target} = CStr({HOST_SHEET_VAR}.Cells({srow}, {scol}).{read_prop}) & \"/\" & TypeName({HOST_SHEET_VAR}.Cells({srow}, {scol}).{read_prop})",
            ]

        if kind < 0.72:
            member = self.rng.choice(
                [
                    "Name",
                    "Range.Address",
                    "Range.Address(False, False)",
                    "HeaderRowRange.Address",
                    "HeaderRowRange.Address(False, False)",
                    "DataBodyRange.Address",
                    "DataBodyRange.Address(False, False)",
                    "ListRows.Count",
                    "ListColumns.Count",
                    f"ListRows({self.rng.randint(1, 3)}).Range.Address(False, False)",
                    f"ListRows({self.rng.randint(1, 3)}).Index",
                    f"ListColumns({self.rng.randint(1, 3)}).Name",
                    f"ListColumns({self.rng.randint(1, 3)}).Index",
                    f"ListColumns({self.rng.randint(1, 3)}).Range.Address(False, False)",
                    f"ListColumns({self.rng.randint(1, 3)}).DataBodyRange.Address(False, False)",
                    f"ListColumns({self.rng.choice(TABLE_COLUMNS)}).Range.Address(False, False)",
                    "ShowTotals",
                    "ShowHeaders",
                ]
            )
            return [f'{target} = {HOST_SHEET_VAR}.ListObjects("Sales").{member}']

        if kind < 0.84:
            op = self.rng.random()
            if op < 0.38:
                arg = self.rng.choice(["", f"({self.rng.randint(1, 4)})"])
                return [
                    f'Set {HOST_RANGE_VAR} = {HOST_SHEET_VAR}.ListObjects("Sales").ListRows.Add{arg}.Range',
                    f'{target} = {HOST_RANGE_VAR}.Address(False, False) & "/" & CStr({HOST_SHEET_VAR}.ListObjects("Sales").ListRows.Count)',
                ]
            if op < 0.62:
                idx = self.rng.randint(1, 3)
                return [
                    f'{HOST_SHEET_VAR}.ListObjects("Sales").ListRows({idx}).Delete',
                    f'{target} = TypeName({HOST_SHEET_VAR}.ListObjects("Sales").DataBodyRange) & "/" & CStr({HOST_SHEET_VAR}.ListObjects("Sales").ListRows.Count)',
                ]
            if op < 0.80:
                value = self.rng.choice(BOOL_LITERALS)
                return [
                    f'{HOST_SHEET_VAR}.ListObjects("Sales").ShowTotals = {value}',
                    f'{target} = CStr({HOST_SHEET_VAR}.ListObjects("Sales").ShowTotals) & "/" & {HOST_SHEET_VAR}.ListObjects("Sales").Range.Address(False, False)',
                ]
            body_row = self.rng.randint(1, 3)
            body_col = self.rng.randint(1, 3)
            value = self.expr(depth - 1, vars_in_scope)
            return [
                f'{HOST_SHEET_VAR}.ListObjects("Sales").DataBodyRange.Cells({body_row}, {body_col}).Value = {value}',
                f'{target} = CStr({HOST_SHEET_VAR}.ListObjects("Sales").DataBodyRange.Cells({body_row}, {body_col}).Value) & "/" & TypeName({HOST_SHEET_VAR}.ListObjects("Sales").DataBodyRange.Cells({body_row}, {body_col}).Value)',
            ]

        if kind < 0.93:
            col_idx = self.rng.randint(1, 3)
            new_col = TABLE_COLUMN_NAMES[self.rng.randrange(len(TABLE_COLUMN_NAMES))]
            return [
                f'{HOST_SHEET_VAR}.ListObjects("Sales").ListColumns({col_idx}).Name = {new_col}',
                f'{target} = {HOST_SHEET_VAR}.ListObjects("Sales").ListColumns({col_idx}).Name & "/" & {HOST_SHEET_VAR}.ListObjects("Sales").HeaderRowRange.Cells(1, {col_idx}).Value',
            ]

        new_name = f"Sales_{self.rng.randint(1, 999)}"
        return [
            f'{HOST_SHEET_VAR}.ListObjects("Sales").Name = "{new_name}"',
            f'{target} = {HOST_SHEET_VAR}.ListObjects("{new_name}").Name & "/" & TypeName({HOST_SHEET_VAR}.ListObjects("{new_name}"))',
        ]

    def module(self, index):
        depth = self.rng.randint(1, 3)
        n_vars = self.rng.randint(2, len(VAR_NAMES))
        vars_in_scope = VAR_NAMES[:n_vars]
        body = self.statements(self.rng.randint(1, 4), vars_in_scope, depth)
        result = self.rng.choice(vars_in_scope)

        lines = [f"Private Function Gen{index}()"]
        lines.append("    Dim " + ", ".join(vars_in_scope) + ", vi, vn")

        lines.append(
            f"    Dim {HOST_SHEET_VAR} As Worksheet, "
            f"{HOST_CELL_VAR} As Range, {HOST_RANGE_VAR} As Range"
        )
        lines.append(f'    Set {HOST_SHEET_VAR} = ThisWorkbook.Worksheets("Data")')
        for v in vars_in_scope:
            lines.append(f"    {v} = {self.rng.choice(NUM_LITERALS)}")
        lines.extend("    " + s for s in body)
        lines.append(f"    Gen{index} = {result}")
        lines.append("End Function")
        return "\n".join(lines)


HARNESS_TEMPLATE = """Public Function Harness{i}() As String
    On Error GoTo Failed
    Dim r
    r = Gen{i}()
    Harness{i} = "OK|" & TypeName(r) & "|" & CStr(r)
    Exit Function
Failed:
    Harness{i} = "ERR|" & CStr(Err.Number)
End Function"""


GRID_HARNESS_TEMPLATE = """Public Function Harness{i}() As String
    On Error GoTo Failed
    ResetGrid
    Dim r
    r = Gen{i}()
    Harness{i} = "OK|" & TypeName(r) & "|" & CStr(r) & "|" & GridState()
    Exit Function
Failed:
    Harness{i} = "ERR|" & CStr(Err.Number) & "|" & GridState()
End Function"""


GRID_HELPERS = """Public Sub ResetGrid()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    ws.Range("A1:F4").Value = Empty
{writes}
End Sub

Private Function CellText(ByVal c As Range) As String
    On Error GoTo Unrenderable
    CellText = TypeName(c.Value) & ":" & CStr(c.Value2)
    Exit Function
Unrenderable:
    CellText = TypeName(c.Value) & ":?" & CStr(Err.Number)
End Function

Private Function GridState() As String
    Dim ws As Worksheet, s As String, r As Long, c As Long
    Set ws = ThisWorkbook.Worksheets("Data")
    For r = 1 To {rows}
        For c = 1 To {cols}
            s = s & CellText(ws.Cells(r, c)) & ","
        Next c
    Next r
    GridState = s
End Function"""


def grid_helpers():
    writes = []
    for row, col, value in GRID:
        if value is True:
            literal = "True"
        elif isinstance(value, str) and value.startswith("="):
            writes.append(f'    ws.Cells({row}, {col}).Formula = "{value}"')
            continue
        elif (
            isinstance(value, str)
            and not value.lstrip("-").replace(".", "", 1).isdigit()
        ):
            literal = f'"{value}"'
        else:
            literal = str(value)
        writes.append(f"    ws.Cells({row}, {col}).Value = {literal}")
    return GRID_HELPERS.format(writes="\n".join(writes), rows=GRID_ROWS, cols=GRID_COLS)


def check_no_duplicate_dims(source):
    proc, declared = None, set()
    for line in source.splitlines():
        stripped = line.strip()
        lowered = stripped.lower()
        if lowered.startswith(
            ("private function", "public function", "private sub", "public sub")
        ):
            proc, declared = stripped, set()
        elif lowered.startswith("dim "):
            for part in stripped[4:].split(","):
                name = part.strip().split()[0] if part.strip() else ""
                if not name:
                    continue
                if name.lower() in declared:
                    raise AssertionError(
                        f"generated source declares {name!r} twice in {proc!r}; "
                        "Excel would answer this with a modal compile error"
                    )
                declared.add(name.lower())
    return source


def build_module(cases):
    parts = ['Attribute VB_Name = "M"', grid_helpers()]
    for i, src in cases:
        parts.append(src)
        parts.append(GRID_HARNESS_TEMPLATE.format(i=i))
    return check_no_duplicate_dims("\n\n".join(parts) + "\n")


def build_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    headers = ["Region", "Amount", "Flag"]
    rows = [["East", 10, True], ["West", 20, False], ["East", 30, True]]
    for c, value in enumerate(headers, start=8):
        ws.cell(1, c, value)
    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=8):
            ws.cell(r, c, value)
    table = openpyxl.worksheet.table.Table(displayName="Sales", ref="H1:J4")
    table.tableStyleInfo = openpyxl.worksheet.table.TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    wb.save(path)


def visi_result(source, proc, workbook=None, harness=False):
    try:
        if workbook is None:
            type_name, value = visi_core.run_macro(source, proc)
        else:
            wb = visi_core.Workbook.load(workbook)
            wb.add_macro("M", source)
            type_name, value, _mutated = wb.run_macro(proc)
    except visi_core.VbaRuntimeError as e:
        return f"ERR|{getattr(e, 'number', '?')}"
    except visi_core.VbaSyntaxError as e:
        return f"SYNTAX|{e}"
    except visi_core.VisiError as e:
        return f"SYNTAX|{type(e).__name__}: {e}"
    if value is None:
        return "ERR|94"
    if harness:
        return value
    return f"OK|{type_name}|{value}"


FLOAT_REL_TOL = 1e-7
FLOAT_ABS_TOL = 1e-7


def _numeric(text):
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def fields_match(mine, theirs):
    if mine == theirs:
        return True
    a, b = mine.split("|"), theirs.split("|")
    if len(a) != len(b):
        return False
    for x, y in zip(a, b):
        if x == y:
            continue

        xs, ys = x.split(","), y.split(",")
        if len(xs) != len(ys):
            return False
        for xi, yi in zip(xs, ys):
            if xi == yi:
                continue
            xt, _, xv = xi.rpartition(":")
            yt, _, yv = yi.rpartition(":")
            if xt != yt:
                return False
            xn, yn = _numeric(xv), _numeric(yv)
            if xn is None or yn is None:
                return False
            if not math.isclose(xn, yn, rel_tol=FLOAT_REL_TOL, abs_tol=FLOAT_ABS_TOL):
                return False
    return True


_WIN32COM_VBA_RUNNER = """
import sys
import win32com.client

xlsm_path = sys.argv[1]
indices = [int(a) for a in sys.argv[2:]]

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
# msoAutomationSecurityForceDisable would silently skip running any macro
# at all; msoAutomationSecurityLow (1) runs macros without the "Enable
# Content" prompt that would otherwise hang this exactly like a compile
# error does.
excel.AutomationSecurity = 1
try:
    wb = excel.Workbooks.Open(xlsm_path)
    try:
        for i in indices:
            try:
                result = excel.Run("Harness{}".format(i))
            except Exception as e:
                result = "ERR|COM:{}".format(e)
            print("{}={}".format(i, result))
    finally:
        wb.Close(False)
finally:
    excel.Quit()
"""


class ExcelDriver:
    def __init__(self, excel_path=None, driver_type="auto", timeout=60):
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
        self.restarts = 0

    def app_name(self):
        name = self.excel_path or EXCEL_APP
        if name.endswith(".app"):
            name = os.path.splitext(os.path.basename(name))[0]
        return name

    def restart(self):
        self.restarts += 1
        subprocess.run(
            ["killall", EXCEL_APP],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        time.sleep(1.0)
        pgrep = subprocess.run(
            ["pgrep", "-x", EXCEL_APP],
            stdout=subprocess.PIPE,
            text=True,
            check=False,
        )
        for pid in pgrep.stdout.split():
            subprocess.run(
                ["kill", "-9", pid],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        time.sleep(1.0)
        subprocess.run(
            ["open", "-a", EXCEL_APP],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        time.sleep(4.0)

    def restart_windows(self):
        self.restarts += 1
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        time.sleep(1.0)

    def _run_win32com_batch(self, xlsm, indices):
        indices_args = [str(i) for i in indices]
        res = subprocess.run(
            [sys.executable, "-u", "-c", _WIN32COM_VBA_RUNNER, xlsm, *indices_args],
            capture_output=True,
            text=True,
            timeout=self.timeout,
            check=False,
        )
        out = {}
        for line in res.stdout.splitlines():
            if "=" in line:
                k, _, v = line.partition("=")
                if k.strip().isdigit():
                    out[int(k.strip())] = v.rstrip("\r")
        return res.returncode, out

    def run_batch(self, xlsm, indices):
        if self.driver_type == "win32com":
            for attempt in range(2):
                try:
                    returncode, out = self._run_win32com_batch(xlsm, indices)
                except subprocess.TimeoutExpired:
                    self.restart_windows()
                    if attempt == 0:
                        continue
                    return {}
                if returncode == 0:
                    return out
                self.restart_windows()
                if attempt == 0:
                    continue
                return {}
            return {}
        if self.driver_type == "mock":
            return {}

        calls = "\n".join(
            f'    set acc to acc & "{i}=" & (run VB macro "Harness{i}") & linefeed'
            for i in indices
        )
        script = "\n".join(
            [
                f'tell application "{self.app_name()}"',
                "    set display alerts to false",
                "    try",
                "        close workbooks saving no",
                "    end try",
                f'    open POSIX file "{os.path.abspath(xlsm)}"',
                "    set wb to active workbook",
                '    set acc to ""',
                calls,
                "    close wb saving no",
                "    return acc",
                "end tell",
            ]
        )
        for attempt in range(2):
            try:
                res = subprocess.run(
                    ["osascript", "-e", script],
                    capture_output=True,
                    text=True,
                    timeout=self.timeout,
                    check=False,
                )
            except subprocess.TimeoutExpired:
                self.restart()
                if attempt == 0:
                    continue
                return {}
            if res.returncode == 0:
                out = {}
                for line in res.stdout.splitlines():
                    if "=" in line:
                        k, _, v = line.partition("=")
                        if k.strip().isdigit():
                            out[int(k.strip())] = v.rstrip("\r")
                return out
            self.restart()
            if attempt == 0:
                continue
            return {}
        return {}


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--excel-path")
    ap.add_argument(
        "--driver", choices=["auto", "applescript", "win32com", "mock"], default="auto"
    )
    ap.add_argument("--iterations", type=int, default=20)
    ap.add_argument(
        "--batch",
        type=int,
        default=None,
        help="Cases per Excel round trip (default: 20 for basic, 1 for extended). The round trip dominates cost.",
    )
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument(
        "--restart-every",
        type=int,
        default=4,
        help="Restart Excel every N batches, before its automation "
        "bridge degrades (0 to only restart on failure).",
    )
    ap.add_argument("--output-dir", default="./fuzz_results")
    ap.add_argument(
        "--host-surface",
        choices=["basic", "extended"],
        default="basic",
        help="Host object model surface to generate. 'extended' adds styles, structural edits and Excel Table operations.",
    )
    args = ap.parse_args()
    if args.batch is None:
        args.batch = 1 if args.host_surface == "extended" else 20

    gen = VbaGenerator(args.seed, host_surface=args.host_surface)
    driver = ExcelDriver(args.excel_path, args.driver, args.timeout)
    failures_dir = os.path.join(args.output_dir, "failures")
    os.makedirs(failures_dir, exist_ok=True)

    print("=" * 69)
    print(
        "    visi vs. Microsoft Excel VBA Execution Differential Fuzzer    ".center(69)
    )
    print("=" * 69)
    print(f" Cases       : {args.iterations} in batches of {args.batch}")
    print(f" Host surface: {args.host_surface}")
    print(f" Excel driver: {driver.driver_type} ({args.excel_path or 'default'})")
    if driver.driver_type == "mock":
        print(" MOCK DRIVER -- visi runs, Excel is not consulted, nothing is compared.")
    print("=" * 69 + "\n")

    passed = failed = skipped = 0
    workdir = tempfile.mkdtemp(prefix="vba_exec_fuzz_")
    start = time.time()

    try:
        for batch_no, batch_start in enumerate(range(0, args.iterations, args.batch)):
            if (
                args.restart_every
                and batch_no
                and batch_no % args.restart_every == 0
                and driver.driver_type != "mock"
            ):
                if driver.driver_type == "win32com":
                    driver.restart_windows()
                else:
                    driver.restart()
            n = min(args.batch, args.iterations - batch_start)
            cases = [
                (batch_start + i + 1, gen.module(batch_start + i + 1)) for i in range(n)
            ]
            source = build_module(cases)
            indices = [i for i, _ in cases]

            base = os.path.join(workdir, "base.xlsx")
            build_workbook(base)

            excel = {}
            if driver.driver_type != "mock":
                xlsm = os.path.join(workdir, f"batch_{batch_start}.xlsm")
                wb = visi_core.Workbook.load(base)
                wb.add_macro("M", source)
                wb.save(xlsm)
                excel = driver.run_batch(xlsm, indices)

            for i, _ in cases:
                mine = visi_result(source, f"Harness{i}", workbook=base, harness=True)
                theirs = excel.get(i)
                if theirs is None:
                    skipped += 1
                    continue
                if fields_match(mine, theirs):
                    passed += 1
                    continue
                failed += 1
                print(f" case {i:<5} [MISMATCH]")
                print(f"   visi : {mine}")
                print(f"   excel: {theirs}")
                out = os.path.join(failures_dir, f"vba_exec_case_{i}")
                os.makedirs(out, exist_ok=True)
                with open(os.path.join(out, "source.bas"), "w") as f:
                    f.write(source)
                with open(os.path.join(out, "verdicts.txt"), "w") as f:
                    f.write(f"procedure: Gen{i}\nvisi:  {mine}\nexcel: {theirs}\n")
                print(f"   saved: {out}")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    total = passed + failed
    print("\n" + "=" * 69)
    print(
        f" Completed in {time.time() - start:.1f}s ({driver.restarts} Excel restarts)"
    )
    print(f" Agreed   : {passed}/{total}" if total else " Agreed   : n/a")
    print(f" Mismatch : {failed}")
    if skipped:
        print(f" Skipped  : {skipped} (Excel gave no answer)")
    print("=" * 69)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
