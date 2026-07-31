#!/usr/bin/env python3
"""
visi vs. Microsoft Excel Differential Fuzzing: Pivot Tables
=============================================================
Generates a random source workbook (a mix of low-cardinality categorical
columns and numeric columns, suited to exercising pivot grouping), builds a
matching pivot table configuration in both `visi` (via the `visi pivot` CLI)
and real Microsoft Excel (by driving Excel's own PivotTable object model via
AppleScript/COM automation -- Excel does not auto-generate a pivot table from
XML the way it recalculates formulas, so this is a fundamentally different
execution pipeline from `fuzz_excel.py`'s "generate formulas -> recalculate"
flow, which is why this lives in its own script), then compares the two
engines' rendered pivot output cell-for-cell.

This reuses (imports, does not duplicate) `XLSXEvaluatedReader` and
`DifferentialComparator` from `fuzz_excel.py` -- both engines materialize
pivot output as plain literal cell values, so the existing generic cell
reader/comparator works unchanged.

IMPORTANT -- the AppleScript path does NOT build the pivot table via Excel's
native AppleScript pivot dictionary. `make new pivot cache at wb` is declared
in Excel.sdef but fails with a generic "Parameter error (-50)" for every
variant tried against real Excel (bare, with properties, range object vs.
text source data, different containers) -- confirmed to be a real
functional gap in Mac Excel's AppleScript support, not a syntax mistake.
Instead, the AppleScript driver opens a copy of `fuzz/BuildFuzzPivot.bas`'s
one-time, human-created macro-enabled workbook (`fuzz/pivot_macro_template.xlsm`,
NOT checked in -- see "Setup" below) and invokes its `BuildFuzzPivot` macro
via the `run VB macro` command, which *is* a working, standard AppleScript
command -- the macro itself does the real work through VBA's PivotCaches.Create
/ CreatePivotTable, the same well-documented object model win32com already
uses directly.

Setup (one-time, macOS only -- skip if you only use win32com on Windows):
    1. Open Excel, create a new blank workbook.
    2. Alt+F11 (or Tools > Macro > Visual Basic Editor) to open the VBA editor.
    3. Insert > Module, paste the contents of fuzz/BuildFuzzPivot.bas.
    4. Save As... fuzz/pivot_macro_template.xlsm (Excel Macro-Enabled Workbook).
    5. Ensure macros are allowed to run for this file (Excel Preferences >
       Security, or Trust Center on Windows) -- this only needs to permit
       *running* an already-embedded macro, not any VBA-project scripting
       permission.

STATUS -- piloted against real Excel, works end-to-end (see fuzz/README.md's
"Known caveats" section for full detail, this is just the summary):
  - FIXED: visi's pivot output used to never render filter/page fields as
    header rows the way Excel does. `libvisi/src/core/pivot.rs`'s
    `compute_pivot`/`PivotGrid` now builds a `filter_rows` block (one
    "FieldName | (All)"/"(Multiple Items)" row per filter field plus a
    blank spacer, matching Excel's own convention empirically verified
    against real Excel) that `visi/src/engine.rs` materializes above the
    grid and `pivot_xlsx.rs`'s native XML accounts for via `rowPageCount`.
  - Still open: a *separate* layout gap surfaced while verifying the fix
    above -- whenever row/col fields are present, Excel's header caption is
    literally "Row Labels"/"Column Labels", not the field name, unlike
    visi's grid. `BuildFuzzPivot.bas` already sets `LayoutForm = xlTabular`
    per field, which should prevent this, so either that isn't taking
    effect via this VBA/AppleScript path or something resets it afterward.
    Not yet root-caused -- this, not an unrelated bug, is what was behind
    most of the previously-reported "wide-grid"/large-mismatch iterations.
  - One tiny-edge-case config (a column used as both a col field and a
    filter field, with the filter selecting zero values, over a 1-row
    source) made the VBA macro itself throw an outright Excel "Parameter
    error (-50)" -- also not yet root-caused.

Usage:
    python3 fuzz/fuzz_pivot.py --driver mock --iterations 5
    python3 fuzz/fuzz_pivot.py --excel-path "/Applications/Microsoft Excel.app" --iterations 1 --seed 1
"""

import argparse
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fuzz_excel import XLSXEvaluatedReader, DifferentialComparator  # noqa: E402


def _fix_openpyxl_table_rels(xlsx_path):
    """openpyxl writes worksheet->table relationship `Target` attributes as
    absolute package paths (e.g. "/xl/tables/table1.xml"), which is valid
    per the OPC spec and something real Excel and rust_xlsxwriter-authored
    files never do -- they use "../tables/table1.xml"-style relative paths.
    calamine's relationship resolver (used by `visi`'s xlsx importer) only
    special-cases the "../" relative form; an absolute leading-"/" target
    silently resolves to no zip entry, so `workbook.load_tables()` succeeds
    but finds zero tables. This is a real interop gap between openpyxl and
    calamine (confirmed independent of pivot tables entirely -- plain `visi
    table list` on an untouched openpyxl-authored table also reports none),
    not something in scope to fix in visi's core import path here, so this
    rewrites the relationship XML in place after openpyxl saves the file.
    """
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    changed = False
    for name in names:
        if not (name.startswith("xl/worksheets/_rels/") and name.endswith(".rels")):
            continue
        text = contents[name].decode("utf-8")
        new_text = re.sub(r'Target="/xl/', 'Target="../', text)
        if new_text != text:
            contents[name] = new_text.encode("utf-8")
            changed = True
    if not changed:
        return

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])


# -----------------------------------------------------------------------------
# 1. Source workbook + pivot configuration generator
# -----------------------------------------------------------------------------


class PivotFuzzGenerator:
    """Generates a random source workbook plus a matching pivot table
    configuration (as a plain dict, not XML) that both `VisiPivotDriver` and
    `ExcelPivotDriver` build a real pivot table from.

    Columns are fixed and deliberately low-cardinality where it matters for
    grouping, unlike `fuzz_excel.py`'s fully-random grid -- a pivot fuzzer
    that used high-cardinality random text for row/col fields would turn
    every group into a singleton and never exercise subtotal/grand-total
    logic at all.
    """

    COL_NAMES = ["Cat", "Mixed", "NumStr", "Amount", "Rate", "Flag"]
    CATEGORICAL_COLS = [0, 1, 2]  # Cat, Mixed, NumStr -- candidates for row/col fields
    NUMERIC_COLS = [3, 4]  # Amount, Rate -- candidates for value fields
    FILTERABLE_COLS = [0, 1, 2, 5]  # any column can be a filter field

    CATEGORIES = ["Alpha", "Beta", "Gamma", "Delta", "Epsilon"]
    # Same values in different cases, to probe case-insensitive grouping
    # parity between visi and Excel.
    CASE_VARIANTS = ["East", "east", "WEST", "west", "North"]
    AGGREGATIONS = ["sum", "count", "count-numbers", "average", "max", "min"]

    def __init__(self, seed=None):
        if seed is not None:
            random.seed(seed)

    def _random_numstr(self):
        """A quoted (forced-text), possibly numeric-looking string, or a
        blank -- probes the sort/group-key numeric-vs-text ambiguity in
        visi's `sort_group_entries` (pivot.rs)."""
        roll = random.random()
        if roll < 0.25:
            return None
        if roll < 0.5:
            return f"0{random.randint(0, 9)}"  # e.g. "08"
        if roll < 0.75:
            return f".0{random.randint(0, 999)}"  # e.g. ".0394"
        return str(random.randint(-50, 50))

    def generate(self, source_path, num_rows, use_table):
        """Builds `source_path` and returns a pivot configuration dict:
        {source_range, table_name, row_fields, col_fields, value_fields,
         filter_field, grand_totals_row, grand_totals_col}.
        `table_name` is None when `use_table` is False (raw-range source).
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
        for c, name in enumerate(self.COL_NAMES, start=1):
            ws.cell(row=1, column=c, value=name)

        # Track each filterable column's actual distinct (blank-normalized)
        # values as we generate, so the filter field picked below can select
        # a real subset instead of guessing at what exists.
        distinct = {i: set() for i in self.FILTERABLE_COLS}
        for r in range(2, num_rows + 2):
            cat = "" if random.random() < 0.1 else random.choice(self.CATEGORIES)
            ws.cell(row=r, column=1, value=cat)
            distinct[0].add(cat if cat else "(blank)")

            mixed = random.choice(self.CASE_VARIANTS)
            ws.cell(row=r, column=2, value=mixed)
            distinct[1].add(mixed)

            numstr = self._random_numstr()
            ws.cell(row=r, column=3, value=numstr)
            distinct[2].add(numstr if numstr else "(blank)")

            ws.cell(row=r, column=4, value=random.randint(-100, 100))
            ws.cell(row=r, column=5, value=round(random.uniform(-500.0, 500.0), 4))

            flag = random.choice([True, False])
            ws.cell(row=r, column=6, value=flag)
            distinct[5].add("TRUE" if flag else "FALSE")

        source_range = f"A1:F{num_rows + 1}"
        table_name = None
        if use_table:
            table_name = "FuzzTable"
            table = Table(displayName=table_name, ref=source_range)
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            ws.add_table(table)

        wb.save(source_path)
        if use_table:
            _fix_openpyxl_table_rels(source_path)

        pool = self.CATEGORICAL_COLS[:]
        random.shuffle(pool)
        n_row = random.randint(0, min(2, len(pool)))
        row_cols, pool = pool[:n_row], pool[n_row:]
        n_col = random.randint(0, min(2, len(pool)))
        col_cols, pool = pool[:n_col], pool[n_col:]

        row_fields = [
            {"column": self.COL_NAMES[i], "subtotal": random.random() < 0.7} for i in row_cols
        ]
        col_fields = [
            {"column": self.COL_NAMES[i], "subtotal": random.random() < 0.7} for i in col_cols
        ]

        n_value = random.randint(1, 2)
        value_fields = [
            {
                "column": self.COL_NAMES[random.choice(self.NUMERIC_COLS)],
                "agg": random.choice(self.AGGREGATIONS),
            }
            for _ in range(n_value)
        ]

        filter_field = None
        if random.random() < 0.5:
            fcol = random.choice(self.FILTERABLE_COLS)
            values = sorted(distinct[fcol])
            # May legitimately come out empty -> filters out every record.
            selected = [v for v in values if random.random() < 0.5]
            filter_field = {"column": self.COL_NAMES[fcol], "values": selected}

        return {
            "source_range": source_range,
            "table_name": table_name,
            "row_fields": row_fields,
            "col_fields": col_fields,
            "value_fields": value_fields,
            "filter_field": filter_field,
            "grand_totals_row": random.random() < 0.7,
            "grand_totals_col": random.random() < 0.7,
        }


# -----------------------------------------------------------------------------
# 2. Execution drivers
# -----------------------------------------------------------------------------

PIVOT_NAME = "FuzzPivot"
DEST_CELL = "H1"  # two columns clear of the source block (A:F)
# One-time, human-created macro-enabled workbook containing BuildFuzzPivot.bas
# -- see fuzz_pivot.py's module docstring for setup steps. Not checked in
# (it's a compiled-macro binary asset a human authors once, not generated).
MACRO_TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pivot_macro_template.xlsm")

# visi's `Count` aggregation counts any non-blank value -- matches Excel's
# `xlCount` (the *default* summary function Excel picks for a text field,
# labeled plain "Count" in the value-field-settings UI); visi's
# `CountNumbers` counts only numeric values -- matches Excel's `xlCountNums`.
# There is no separate "CountA" member in Excel's `XlConsolidationFunction`
# enum (confirmed via Excel.sdef) -- an earlier draft of this mapping
# assumed one existed and was wrong. `BuildFuzzPivot.bas` (invoked by the
# AppleScript path) matches these same keys directly in its `Select Case`,
# so only the win32com path needs its own VBA constant-name mapping.
AGG_TO_WIN32COM_FUNCTION = {
    "sum": "xlSum",
    "count": "xlCount",
    "count-numbers": "xlCountNums",
    "average": "xlAverage",
    "max": "xlMax",
    "min": "xlMin",
}


class VisiPivotDriver:
    """Builds a matching pivot table via the `visi pivot` CLI: `create`, then
    repeated `add-field`/`filter`. Each mutating call auto-refreshes (see
    `WorkbookManager::add_pivot_field`'s doc comment), so no explicit
    `refresh` is needed -- this mirrors how a real user would build up a
    pivot table incrementally.
    """

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

    def _run(self, args):
        cmd = [self.binary_path, "pivot"] + args
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if res.returncode != 0:
            raise RuntimeError(
                f"visi pivot {' '.join(args)} failed:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}"
            )

    def run(self, source_file, config, output_file):
        shutil.copyfile(source_file, output_file)

        create_args = ["create", output_file, "--name", PIVOT_NAME, "--dest", DEST_CELL, "-i"]
        if config["table_name"]:
            create_args += ["--source-table", config["table_name"]]
        else:
            create_args += ["--source-range", config["source_range"]]
        self._run(create_args)

        for f in config["row_fields"]:
            self._run(
                ["add-field", output_file, "--name", PIVOT_NAME, "--area", "row",
                 "--column", f["column"], "-i"]
            )
        for f in config["col_fields"]:
            self._run(
                ["add-field", output_file, "--name", PIVOT_NAME, "--area", "column",
                 "--column", f["column"], "-i"]
            )
        for f in config["value_fields"]:
            self._run(
                ["add-field", output_file, "--name", PIVOT_NAME, "--area", "value",
                 "--column", f["column"], "--agg", f["agg"], "-i"]
            )
        if config["filter_field"]:
            self._run(
                ["add-field", output_file, "--name", PIVOT_NAME, "--area", "filter",
                 "--column", config["filter_field"]["column"], "-i"]
            )
            values = config["filter_field"]["values"]
            if values:
                self._run(
                    ["filter", output_file, "--name", PIVOT_NAME,
                     "--column", config["filter_field"]["column"],
                     "--values", ",".join(values), "-i"]
                )
            # else: the config wants "select nothing", but the CLI's
            # `filter` command has no verb for an empty selection (only a
            # comma list or `--clear`) -- leave the field unfiltered rather
            # than attempting a config the CLI can't express. This is a gap
            # in the CLI surface, not the fuzzer; Layer 1 (pivot.rs's Rust
            # invariant fuzzer) already covers the empty-selection case
            # directly against `PivotTable`/`compute_pivot`.


class ExcelPivotDriver:
    """Drives Microsoft Excel's own PivotTable object model to build a
    matching pivot table, then saves. Unlike `fuzz_excel.py`'s `ExcelDriver`
    (which only needs `calculate` over cells visi already computed), Excel
    must *construct* a live PivotTable here -- there's no XML shortcut.

    The win32com path (Windows) uses the standard, well-documented VBA
    object model directly and is straightforward.

    The AppleScript path (macOS) cannot do the same directly: `make new
    pivot cache at wb` is declared in Excel.sdef (extracted straight from
    `Microsoft Excel.app/Contents/Resources/Excel.sdef`, since the `sdef`
    CLI tool needs a full Xcode install this environment doesn't have) but
    fails with a generic "Parameter error (-50)" for every variant tried
    against real Excel -- bare, with properties, range object vs. text
    source data, different containers. That's a real functional gap in Mac
    Excel's AppleScript support (the dictionary documents a capability the
    implementation doesn't back), not a syntax mistake -- other pivot
    field/property names below (`pivot field orientation`, `orient as row
    field`, `layout form`, `layout subtotal location`, `set subtotals`,
    `range object` of a list object) were all verified against the same
    .sdef and do work once a PivotTable already exists.

    So instead, the AppleScript path builds the pivot through a one-time,
    human-authored macro (`fuzz/BuildFuzzPivot.bas`, pasted into
    `fuzz/pivot_macro_template.xlsm` -- see this module's docstring for
    setup) invoked via `run VB macro`, which *is* a working AppleScript
    command. The macro itself uses the same PivotCaches.Create /
    CreatePivotTable / PivotFields object model as the win32com path below,
    just reached through VBA instead of AppleScript's broken pivot-cache
    creation.
    """

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

    def run(self, source_file, config, output_file):
        if self.driver_type == "mock":
            shutil.copyfile(source_file, output_file)
            print("[ExcelPivotDriver Warning] Running in mock mode (Excel not invoked).")
            return
        elif self.driver_type == "applescript":
            # Excel needs a macro-enabled (.xlsm) copy to run BuildFuzzPivot
            # from; build one alongside the requested output path, then
            # copy its bytes back so callers don't need to know about the
            # extension difference.
            macro_file = os.path.splitext(output_file)[0] + ".xlsm"
            self._prepare_macro_workbook(source_file, config, macro_file)
            self._run_applescript_macro(os.path.abspath(macro_file), config)
            shutil.copyfile(macro_file, output_file)
        elif self.driver_type == "win32com":
            shutil.copyfile(source_file, output_file)
            self._run_win32com(os.path.abspath(output_file), config)
        else:
            raise RuntimeError(f"Unsupported pivot driver type: {self.driver_type}")

    # -- AppleScript (macOS) --------------------------------------------

    def _applescript_str(self, s):
        escaped = s.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'

    def _prepare_macro_workbook(self, source_file, config, macro_file):
        if not os.path.exists(MACRO_TEMPLATE_PATH):
            raise RuntimeError(
                f"Missing {MACRO_TEMPLATE_PATH} -- see fuzz_pivot.py's module "
                "docstring for the one-time setup steps to create it (paste "
                "fuzz/BuildFuzzPivot.bas into a macro-enabled workbook saved "
                "at that path)."
            )
        import openpyxl
        from openpyxl.worksheet.table import Table, TableStyleInfo

        src_wb = openpyxl.load_workbook(source_file)
        src_ws = src_wb["Sheet1"]

        macro_wb = openpyxl.load_workbook(MACRO_TEMPLATE_PATH, keep_vba=True)
        macro_ws = macro_wb.active
        macro_ws.title = "Sheet1"
        for row in src_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    macro_ws.cell(row=cell.row, column=cell.column, value=cell.value)

        if config["table_name"]:
            table = Table(displayName=config["table_name"], ref=config["source_range"])
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            macro_ws.add_table(table)

        macro_wb.save(macro_file)

    def _build_applescript_macro_call(self, output_file, config):
        app_name = self.excel_path if self.excel_path else "Microsoft Excel"
        if app_name.endswith(".app"):
            app_name = os.path.splitext(os.path.basename(app_name))[0]

        row_fields_csv = ";".join(
            f'{f["column"]}:{"1" if f["subtotal"] else "0"}' for f in config["row_fields"]
        )
        col_fields_csv = ";".join(
            f'{f["column"]}:{"1" if f["subtotal"] else "0"}' for f in config["col_fields"]
        )
        value_fields_csv = ";".join(f'{f["column"]}:{f["agg"]}' for f in config["value_fields"])
        if config["filter_field"]:
            filter_spec = config["filter_field"]["column"] + "|" + ",".join(config["filter_field"]["values"])
        else:
            filter_spec = ""
        source_is_table = "1" if config["table_name"] else "0"
        source_ref = config["table_name"] if config["table_name"] else config["source_range"]

        # Labeled command parameters in this dictionary use space-separated
        # syntax (`label value`), not `label:value` colons -- confirmed via
        # the same real-Excel trial-and-error that found `create pivot
        # table`'s working form (see class docstring); colons here produce
        # a plain syntax error, not a runtime one.
        run_macro_line = (
            f'        run VB macro "BuildFuzzPivot" '
            f'arg1 {self._applescript_str(row_fields_csv)} '
            f'arg2 {self._applescript_str(col_fields_csv)} '
            f'arg3 {self._applescript_str(value_fields_csv)} '
            f'arg4 {self._applescript_str(filter_spec)} '
            f'arg5 {self._applescript_str(DEST_CELL)} '
            f'arg6 {self._applescript_str("1" if config["grand_totals_row"] else "0")} '
            f'arg7 {self._applescript_str("1" if config["grand_totals_col"] else "0")} '
            f'arg8 {self._applescript_str(source_is_table)} '
            f'arg9 {self._applescript_str(source_ref)}'
        )

        lines = [
            f'tell application "{app_name}"',
            "    set display alerts to false",
            "    try",
            "        close workbooks saving no",
            "    end try",
            "    try",
            f'        set targetFile to POSIX file "{output_file}"',
            "        open targetFile",
            "        set wb to active workbook",
            run_macro_line,
            "        close wb saving no",
            "    on error errText number errNum",
            "        try",
            "            close workbooks saving no",
            "        end try",
            "        error errText number errNum",
            "    end try",
            "end tell",
        ]
        return "\n".join(lines)

    def _run_applescript_macro(self, abs_output, config):
        script = self._build_applescript_macro_call(abs_output, config)
        res = None
        for attempt in range(5):
            time.sleep(0.5)
            try:
                res = subprocess.run(
                    ["osascript", "-e", script],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=20,
                )
                if res.returncode == 0:
                    break
            except subprocess.TimeoutExpired:
                subprocess.run(["killall", "Microsoft Excel"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                time.sleep(1.0)
        if res is not None and res.returncode != 0:
            raise RuntimeError(f"Excel pivot AppleScript failed:\nSTDERR: {res.stderr}\nScript:\n{script}")

    # -- win32com (Windows) -----------------------------------------------

    def _run_win32com(self, abs_output, config):
        try:
            import win32com.client
            from win32com.client import constants as c
        except ImportError:
            raise RuntimeError("pywin32 (win32com) is required for Excel automation on Windows.")

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(abs_output)
            ws = wb.Sheets("Sheet1")
            if config["table_name"]:
                src_range = ws.ListObjects(config["table_name"]).Range
            else:
                src_range = ws.Range(config["source_range"])

            pc = wb.PivotCaches().Create(SourceType=c.xlDatabase, SourceData=src_range)
            pt = pc.CreatePivotTable(TableDestination=ws.Range(DEST_CELL), TableName=PIVOT_NAME)

            for f in config["row_fields"]:
                pf = pt.PivotFields(f["column"])
                pf.Orientation = c.xlRowField
                subtotals = [False] * 12
                subtotals[0] = f["subtotal"]
                pf.Subtotals = subtotals
            for f in config["col_fields"]:
                pf = pt.PivotFields(f["column"])
                pf.Orientation = c.xlColumnField
                subtotals = [False] * 12
                subtotals[0] = f["subtotal"]
                pf.Subtotals = subtotals
            for f in config["value_fields"]:
                pf = pt.PivotFields(f["column"])
                pf.Orientation = c.xlDataField
                pf.Function = getattr(c, AGG_TO_WIN32COM_FUNCTION[f["agg"]])
            if config["filter_field"]:
                col = config["filter_field"]["column"]
                values = set(config["filter_field"]["values"])
                pf = pt.PivotFields(col)
                pf.Orientation = c.xlPageField
                for item in pf.PivotItems():
                    item.Visible = item.Name in values

            pt.RowAxisLayout(c.xlTabularRow)
            pt.SubtotalLocation(c.xlAtBottom)
            pt.HasAutoFormat = False
            pt.ColumnGrand = config["grand_totals_col"]
            pt.RowGrand = config["grand_totals_row"]
            pt.RefreshTable()
            wb.Save()
            wb.Close()
        finally:
            excel.Quit()


# -----------------------------------------------------------------------------
# 3. CLI & test runner orchestrator
# -----------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Differential fuzzing test harness for visi vs Microsoft Excel pivot tables."
    )
    parser.add_argument("--excel-path", help="Path to Microsoft Excel binary or application bundle.")
    parser.add_argument(
        "--driver", choices=["auto", "applescript", "win32com", "mock"], default="auto",
        help="Excel execution driver.",
    )
    parser.add_argument("--visi-path", default="./target/release/visi", help="Path to compiled visi binary.")
    parser.add_argument("--iterations", type=int, default=10, help="Number of fuzz iterations to run.")
    parser.add_argument("--rows", type=int, default=30, help="Max source data rows per iteration.")
    parser.add_argument(
        "--source-mode", choices=["table", "range", "both"], default="both",
        help="Whether the pivot source is an Excel Table, a raw range, or alternate between both.",
    )
    parser.add_argument("--seed", type=int, default=None, help="Random seed for reproducible fuzzing.")
    parser.add_argument("--output-dir", default="./fuzz_results", help="Directory to store test outputs and failure artifacts.")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    failures_dir = os.path.join(args.output_dir, "failures")
    os.makedirs(failures_dir, exist_ok=True)

    visi_driver = VisiPivotDriver(binary_path=args.visi_path)
    excel_driver = ExcelPivotDriver(excel_path=args.excel_path, driver_type=args.driver)
    comparator = DifferentialComparator()

    print("=====================================================================")
    print("      visi vs. Microsoft Excel Pivot Table Differential Fuzzer       ")
    print("=====================================================================")
    print(f" Iterations  : {args.iterations}")
    print(f" Max rows    : {args.rows}")
    print(f" Source mode : {args.source_mode}")
    print(f" Visi Path   : {visi_driver.binary_path}")
    print(f" Excel Driver: {excel_driver.driver_type} ({args.excel_path or 'Default'})")
    print("=====================================================================\n")

    passed_count = 0
    failed_count = 0
    start_time = time.time()

    for i in range(1, args.iterations + 1):
        iter_seed = (args.seed + i) if args.seed is not None else random.randint(1, 1000000)
        generator = PivotFuzzGenerator(seed=iter_seed)

        temp_dir = tempfile.mkdtemp(prefix=f"fuzz_pivot_iter_{i}_")
        source_xlsx = os.path.join(temp_dir, "source.xlsx")
        visi_out_xlsx = os.path.join(temp_dir, "visi_out.xlsx")
        excel_out_xlsx = os.path.join(temp_dir, "excel_out.xlsx")

        try:
            num_rows = random.randint(1, max(1, args.rows))
            if args.source_mode == "table":
                use_table = True
            elif args.source_mode == "range":
                use_table = False
            else:
                use_table = i % 2 == 0

            config = generator.generate(source_xlsx, num_rows=num_rows, use_table=use_table)

            visi_driver.run(source_xlsx, config, visi_out_xlsx)
            excel_driver.run(source_xlsx, config, excel_out_xlsx)

            visi_cells = XLSXEvaluatedReader.read_evaluated_cells(visi_out_xlsx)
            excel_cells = XLSXEvaluatedReader.read_evaluated_cells(excel_out_xlsx)

            is_match, mismatches = comparator.compare(visi_cells, excel_cells)

            if is_match:
                passed_count += 1
                print(f" Iteration {i:3d}/{args.iterations} [PASSED] (Seed: {iter_seed}, rows: {num_rows}, table: {use_table})")
            else:
                failed_count += 1
                print(f"\n Iteration {i:3d}/{args.iterations} [FAILED] (Seed: {iter_seed})")
                print(f"   Found {len(mismatches)} cell mismatch(es):")
                for m in mismatches[:5]:
                    print(f"   - Cell {m['key'][1]} on {m['key'][0]}: visi={m['visi']} | Excel={m['excel']}")

                fail_case_dir = os.path.join(failures_dir, f"pivot_fail_iter_{i}_seed_{iter_seed}")
                shutil.copytree(temp_dir, fail_case_dir, dirs_exist_ok=True)
                print(f"   Saved failure reproducing files to: {fail_case_dir}\n")

        except Exception as err:
            failed_count += 1
            print(f"\n Iteration {i:3d}/{args.iterations} [ERROR]: {err}")
            fail_case_dir = os.path.join(failures_dir, f"pivot_error_iter_{i}_seed_{iter_seed}")
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
