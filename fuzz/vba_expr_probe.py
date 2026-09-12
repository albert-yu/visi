#!/usr/bin/env python3
import argparse
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

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

from fuzz_vba import HARNESS_TEMPLATE, ExcelDriver, visi_result

PREAMBLE = ["Dim va, vb, vc, vd, ve, vi, vn, a, b, c, n", "n = Null"]


def parse_case(text):
    """`"setup :: expr"` -> (list of setup statements, expression).

    A literal `\\n` in the setup starts a new line, which is the only way to
    probe a block statement: VBA's `:` separator carries several statements on
    one line but will not open a `Select Case` or a multi-line `If`.
    """
    setup, sep, expr = text.rpartition("::")
    if not sep:
        return [], text.strip()
    return [s.strip() for s in setup.split("\\n") if s.strip()], expr.strip()


def build_module(cases):
    parts = ['Attribute VB_Name = "P"']
    for i, (setup, expr) in enumerate(cases, start=1):
        body = "\n".join(f"    {s}" for s in PREAMBLE + setup)
        parts.append(
            f"Private Function Gen{i}()\n{body}\n    Gen{i} = {expr}\nEnd Function"
        )
        parts.append(HARNESS_TEMPLATE.format(i=i))
    return "\n\n".join(parts) + "\n"


def read_cases(path):
    out = []
    with open(path) as f:
        for line in f:
            line = (
                line.split("#", 1)[0].strip()
                if line.lstrip().startswith("#")
                else line.strip()
            )
            if line:
                out.append(line)
    return out


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument(
        "-e",
        "--expr",
        action="append",
        default=[],
        help="A case, `expr` or `setup :: expr`. Repeatable.",
    )
    ap.add_argument("-f", "--file", help="File of cases, one per line; # comments.")
    ap.add_argument("--excel-path")
    ap.add_argument("--driver", choices=["auto", "applescript", "mock"], default="auto")
    ap.add_argument("--timeout", type=int, default=120)
    args = ap.parse_args()

    raw = list(args.expr) + (read_cases(args.file) if args.file else [])
    if not raw:
        ap.error("give at least one case with -e or -f")
    cases = [parse_case(c) for c in raw]
    source = build_module(cases)

    excel = {}
    if args.driver != "mock":
        driver = ExcelDriver(args.excel_path, args.driver, args.timeout)
        with tempfile.TemporaryDirectory(prefix="vba_expr_probe_") as workdir:
            base = os.path.join(workdir, "base.xlsx")
            xlsm = os.path.join(workdir, "probe.xlsm")
            openpyxl.Workbook().save(base)
            wb = visi_core.Workbook.load(base)
            wb.add_macro("P", source)
            wb.save(xlsm)
            excel = driver.run_batch(xlsm, list(range(1, len(cases) + 1)))
        if not excel:
            print(
                "Excel returned nothing: a case is probably a compile error.\n",
                file=sys.stderr,
            )

    width = max(len(c) for c in raw)
    print(f"{'case'.ljust(width)}  {'visi'.ljust(24)}  excel")
    print(f"{'-' * width}  {'-' * 24}  {'-' * 24}")
    disagreed = 0
    for i, case in enumerate(raw, start=1):
        mine = visi_result(source, f"Gen{i}")
        theirs = excel.get(i, "-")
        mark = ""
        if theirs not in ("-", mine):
            mark = "   <-- differs"
            disagreed += 1
        print(f"{case.ljust(width)}  {mine.ljust(24)}  {theirs}{mark}")
    if excel:
        print(f"\n{len(cases) - disagreed}/{len(cases)} agree")
    return 0


if __name__ == "__main__":
    sys.exit(main())
