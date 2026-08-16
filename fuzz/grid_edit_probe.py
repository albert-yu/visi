#!/usr/bin/env python3
"""What real Excel does to a formula when a row or column is inserted or deleted.

The reference-shifting rules in `visi-core/src/core/grid_edit.rs` are the kind
of thing everyone thinks they remember and nobody has checked: does `$A$3` move
when a row is inserted above it? Does inserting at a range's *first* row move
the range or grow it? Does deleting one row of a three-row range shrink it or
break it? This asks Excel, one case at a time, and prints its answer next to
visi's.

    python fuzz/grid_edit_probe.py                    # every case
    python fuzz/grid_edit_probe.py -k absolute        # cases matching a substring
    python fuzz/grid_edit_probe.py --excel-path "/Applications/Microsoft Excel.app"

Unlike the VBA harnesses in this directory there is no macro involved, so the
compile-error hang described in `CLAUDE.md` does not apply -- but Excel is still
driven through `osascript`, so every call has a timeout and a `killall` fallback.

Exit status is non-zero if any case disagrees, which makes this usable as a
check rather than only as an exploration.
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VISI = os.path.join(REPO, "target", "debug", "visi")
OSASCRIPT_TIMEOUT = 60

# Each case is: a starting grid, one structural edit, and the cells whose
# formula text we compare afterwards. Addresses in `probe` are where the
# formula lives *after* the edit.
CASES = [
    {
        "name": "insert row above everything",
        "cells": {"A1": "1", "A2": "2", "A3": "3",
                  "C1": "=A3", "C2": "=SUM(A1:A3)", "C3": "=$A$3*10"},
        "edit": ("insert_row", 1),
        "probe": ["C2", "C3", "C4"],
    },
    {
        "name": "insert row at a range's first row moves the range",
        "cells": {"A2": "1", "A3": "2", "A4": "3", "C1": "=SUM(A2:A4)"},
        "edit": ("insert_row", 2),
        "probe": ["C1"],
    },
    {
        "name": "insert row inside a range grows the range",
        "cells": {"A2": "1", "A3": "2", "A4": "3", "C1": "=SUM(A2:A4)"},
        "edit": ("insert_row", 3),
        "probe": ["C1"],
    },
    {
        "name": "insert row below a range leaves it alone",
        "cells": {"A2": "1", "A3": "2", "A4": "3", "C1": "=SUM(A2:A4)"},
        "edit": ("insert_row", 5),
        "probe": ["C1"],
    },
    {
        "name": "delete the row a reference names",
        "cells": {"A1": "1", "A2": "2", "A3": "3", "C1": "=A3", "C2": "=A3+1"},
        "edit": ("delete_row", 3),
        "probe": ["C1", "C2"],
    },
    {
        "name": "delete one row of a multi-row range shrinks it",
        "cells": {"A1": "1", "A2": "2", "A3": "3", "C1": "=SUM(A1:A3)"},
        "edit": ("delete_row", 2),
        "probe": ["C1"],
    },
    {
        "name": "delete the only row of a single-row range",
        "cells": {"A2": "1", "C1": "=SUM(A2:A2)"},
        "edit": ("delete_row", 2),
        "probe": ["C1"],
    },
    {
        "name": "delete a row above a range slides it up",
        # The formula sits below the deleted row so it survives to be read;
        # putting it in row 1 would delete the formula along with the row.
        "cells": {"A2": "1", "A3": "2", "A4": "3", "C6": "=SUM(A2:A4)"},
        "edit": ("delete_row", 1),
        "probe": ["C5"],
    },
    {
        "name": "an absolute reference is not pinned against a row insert",
        "cells": {"A3": "3", "C1": "=$A$3", "C2": "=A$3", "C3": "=$A3"},
        "edit": ("insert_row", 1),
        "probe": ["C2", "C3", "C4"],
    },
    {
        "name": "insert column shifts column references",
        "cells": {"A1": "1", "B1": "2", "C1": "3",
                  "E1": "=C1", "E2": "=SUM(A1:C1)"},
        "edit": ("insert_col", "B"),
        "probe": ["F1", "F2"],
    },
    {
        "name": "delete the column a reference names",
        "cells": {"A1": "1", "B1": "2", "C1": "3",
                  "E1": "=B1", "E2": "=SUM(A1:C1)"},
        "edit": ("delete_col", "B"),
        "probe": ["D1", "D2"],
    },
    {
        "name": "a whole-column reference under a column insert",
        "cells": {"A1": "1", "B1": "2", "C1": "3",
                  "E1": "=SUM(B:B)", "E2": "=SUM(A:C)"},
        "edit": ("insert_col", "A"),
        "probe": ["F1", "F2"],
    },
    {
        "name": "a whole-column reference under a row insert",
        "cells": {"A1": "1", "B1": "2", "E1": "=SUM(B:B)", "E2": "=SUM(A:C)"},
        "edit": ("insert_row", 1),
        "probe": ["E2", "E3"],
    },
    {
        "name": "deleting the column a whole-column reference names",
        "cells": {"A1": "1", "B1": "2", "C1": "3", "E1": "=SUM(B:B)"},
        "edit": ("delete_col", "B"),
        "probe": ["D1"],
    },
    {
        "name": "a cross-sheet reference follows an edit on the other sheet",
        "cells": {"C1": "=Data!A3", "C2": "=SUM(Data!A1:A3)"},
        "other_sheet": ("Data", {"A1": "10", "A2": "20", "A3": "30"}),
        "edit": ("insert_row", 1, "Data"),
        "probe": ["C1", "C2"],
    },
]


def run(cmd):
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"{' '.join(cmd)} failed:\n{res.stderr}")
    return res.stdout


def build_workbook(case, path):
    """Writes the case's starting grid to `path` using the visi CLI."""
    pairs = []
    for addr, src in case["cells"].items():
        pairs += ["-S", f"{addr}={src}"]
    run([VISI, "set", path, "-o", path, "-q"] + pairs)
    if "other_sheet" in case:
        name, cells = case["other_sheet"]
        run([VISI, "sheet", "add", path, "--name", name, "-i", "-q"])
        pairs = []
        for addr, src in cells.items():
            pairs += ["-S", f"{addr}={src}"]
        run([VISI, "set", path, "-s", name, "-i", "-q"] + pairs)


def visi_edit(case, path):
    """Applies the case's edit with visi, and reads back the probed formulas."""
    kind, index = case["edit"][0], case["edit"][1]
    sheet = case["edit"][2] if len(case["edit"]) > 2 else None
    noun, verb = ("row", "insert") if kind == "insert_row" else \
                 ("row", "delete") if kind == "delete_row" else \
                 ("col", "insert") if kind == "insert_col" else ("col", "delete")
    cmd = [VISI, noun, verb, path, "-x", str(index), "-i", "-q"]
    if sheet:
        cmd += ["-s", sheet]
    run(cmd)
    return read_formulas(path, case["probe"])


def read_formulas(path, addrs):
    """The raw source of each address, straight out of the saved file."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    out = []
    for addr in addrs:
        value = ws[addr].value
        out.append("" if value is None else str(value))
    return out


def excel_script(app, path, case):
    kind, index = case["edit"][0], case["edit"][1]
    sheet = case["edit"][2] if len(case["edit"]) > 2 else None
    target = f'worksheet "{sheet}"' if sheet else "worksheet 1"

    if kind == "insert_row":
        edit = f'insert into range (entire row of range "A{index}" of ews) shift shift down'
    elif kind == "delete_row":
        edit = f'delete range (entire row of range "A{index}" of ews) shift shift up'
    elif kind == "insert_col":
        # The horizontal constants are `shift to right` / `shift to left`, not
        # `shift right` / `shift left` -- the short spellings parse as a
        # parameter name and fail with a bare syntax error.
        edit = f'insert into range (entire column of range "{index}1" of ews) shift shift to right'
    else:
        edit = f'delete range (entire column of range "{index}1" of ews) shift shift to left'

    reads = "\n".join(
        f'set out to out & (get formula of range "{a}" of ws) & "\\n"'
        for a in case["probe"]
    )
    return f'''
    tell application "{app}"
        set display alerts to false
        try
            close workbooks saving no
        end try
        set out to ""
        try
            open POSIX file "{path}"
            set ws to worksheet 1 of active workbook
            set ews to {target} of active workbook
            {edit}
            {reads}
            close active workbook saving no
        on error errText number errNum
            try
                close workbooks saving no
            end try
            error errText number errNum
        end try
        return out
    end tell
    '''


def excel_edit(app, case, path):
    script = excel_script(app, os.path.abspath(path), case)
    try:
        res = subprocess.run(["osascript", "-e", script], stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE, text=True, timeout=OSASCRIPT_TIMEOUT)
    except subprocess.TimeoutExpired:
        subprocess.run(["killall", "Microsoft Excel"], stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        raise RuntimeError("Excel did not respond; killed it")
    if res.returncode != 0:
        raise RuntimeError(f"AppleScript failed: {res.stderr.strip()}")
    return res.stdout.rstrip("\n").split("\n")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel-path", default="Microsoft Excel")
    ap.add_argument("-k", "--filter", default="", help="only cases whose name contains this")
    args = ap.parse_args()

    app = args.excel_path
    if app.endswith(".app"):
        app = os.path.splitext(os.path.basename(app))[0]

    if not os.path.exists(VISI):
        sys.exit(f"build the CLI first: cargo build ({VISI} not found)")

    cases = [c for c in CASES if args.filter in c["name"]]
    tmp = tempfile.mkdtemp(prefix="grid_edit_probe_")
    failures = 0

    for case in cases:
        base = os.path.join(tmp, "base.xlsx")
        for stale in (base, base + ".visi.xlsx", base + ".excel.xlsx"):
            if os.path.exists(stale):
                os.remove(stale)
        build_workbook(case, base)

        v_path, e_path = base + ".visi.xlsx", base + ".excel.xlsx"
        shutil.copyfile(base, v_path)
        shutil.copyfile(base, e_path)

        got = visi_edit(case, v_path)
        try:
            want = excel_edit(app, case, e_path)
        except RuntimeError as exc:
            print(f"  ?? {case['name']}: {exc}")
            failures += 1
            continue

        agree = len(got) == len(want) and all(
            g.strip().upper() == w.strip().upper() for g, w in zip(got, want)
        )
        mark = "ok" if agree else "XX"
        print(f"  {mark} {case['name']}")
        for addr, g, w in zip(case["probe"], got, want + [""] * len(got)):
            flag = "" if g.strip().upper() == w.strip().upper() else "   <-- differs"
            print(f"       {addr:>4}  visi {g!r:<22} excel {w!r}{flag}")
        if not agree:
            failures += 1

    print(f"\n{len(cases) - failures}/{len(cases)} cases agree with Excel")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
