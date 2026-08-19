#!/usr/bin/env python3
"""
What does Excel's VBA *compiler* accept? -- a measurement probe.

`fuzz_vba_parse.py` generates random source and reports an aggregate
agree/disagree tally. This asks the same question -- does Excel compile
this? -- but about snippets *you name*, one at a time, so a disagreement
found by the fuzzer can be minimized down to the single line responsible.

That distinction matters because a generated case is typically 5-15 lines
and a compile error names none of them: Excel's only signal is the modal
dialog that hangs the automation bridge (see `fuzz_vba_parse.py`'s
docstring, point 1). So "which line did it?" is not readable from a fuzz
failure -- it has to be bisected, which is what this is for.

It is also the instrument for the *name-resolution* question specifically,
which is what issue #78 turns on. Phase 0 checking cannot tell

    a genuinely undeclared name (Excel: compile error)
    a real VBA intrinsic or another module's procedure (Excel: fine)

apart without knowing what VBA's built-in surface actually contains. Every
`builtin:*` case below is one probe of that surface -- an unadorned call to
a name in statement position -- and the answers are what
`core/vba/resolve.rs`'s registry is built from. Guessing at that list is
exactly the mistake this codebase keeps re-learning not to make.

Usage:
    python fuzz/vba_compile_probe.py                     # the built-in case list
    python fuzz/vba_compile_probe.py --only undeclared   # cases whose label contains this
    python fuzz/vba_compile_probe.py -e 'x = arr(1, 2)'  # one ad-hoc snippet
    python fuzz/vba_compile_probe.py --list              # print cases, run nothing
    python fuzz/vba_compile_probe.py -e 'Dim x As Long' --sig 'ByVal x As Long' --call-args ' 1'

Cost: an *accepted* snippet is one fast round trip; a *rejected* one costs
the driver's full timeout twice over plus an Excel restart (~35s), since a
hang is the only rejection signal there is. Keep case lists short and
targeted rather than sweeping.
"""

import argparse
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
import visi_core

from fuzz_vba_parse import ExcelVerdictDriver, build_module, visi_verdict


# Each entry is (label, snippet) or (label, snippet, sig, args). The
# snippet is spliced into a dead `If False Then` branch inside `Gen`, so it
# is compiled and never run -- see `fuzz_vba_parse.MODULE_TEMPLATE`. A
# `Helper(Optional a, b, c)` is in scope; nothing else is declared, so every
# other bare name is genuinely undeclared unless the snippet declares it.
#
# The four-element form gives `Gen` a parameter list, which is the only way
# to ask Excel about anything a *parameter* does -- `sig` is spliced into
# `Sub Gen(...)` and `args` into the `Gen` call in `Harness`, which has to
# keep invoking it or the procedure is never compiled at all.
CASES = [
    # -- Minimizations of the issue #78 false negatives that survived the
    #    first resolve.rs pass. The hypothesis under test is that each is a
    #    genuinely-undeclared name rather than some unrelated syntax defect,
    #    which was assumed but never measured.
    ("undeclared:call-2-args", "x = arr(1, 2)"),
    ("undeclared:call-1-arg", "x = arr(1)"),
    ("undeclared:implicit-call-date", "d #1/1/2000#"),
    ("undeclared:implicit-call-kw-arg", "x Function = 1 + 2 * 3"),
    ("undeclared:for-each-in", "For Each c In rng\n    x = 1\nNext"),
    ("undeclared:bare-operand", "x = a _\n    + b"),
    ("undeclared:redim-preserve", "ReDim Preserve arr(1 To 5)"),
    ("undeclared:leading-continuation", "Select Case x\nCase 1\n    _ y = 2\nEnd Select"),

    # -- Follow-ups isolating the two rejections above that are NOT about
    #    name resolution, established by the first round: `ReDim Preserve`
    #    and a leading `_` continuation both reject with no undeclared call
    #    in sight.
    ("redim:preserve-undeclared", "ReDim Preserve arr(1 To 5)"),
    ("redim:plain-undeclared", "ReDim arr(1 To 5)"),
    ("redim:preserve-after-dim", "Dim arr()\nReDim Preserve arr(1 To 5)"),
    ("redim:plain-after-dim", "Dim arr()\nReDim arr(1 To 5)"),
    ("continuation:leading-underscore", "_ y = 2"),
    ("continuation:trailing-underscore", "y = 1 + _\n    2"),

    # -- A bare identifier used as a whole statement, with no arguments and
    #    no parentheses. Left unchecked in the first pass on the grounds
    #    that it had never been measured; a later fuzz run then found Excel
    #    rejecting the undeclared case, so these settle it. The open
    #    question is the *declared* one: `x` alone is still call syntax as
    #    far as VBA is concerned, so a non-callable name may fail too.
    ("bare:undeclared", "x"),
    ("bare:declared-scalar", "Dim x As Long\nx"),
    ("bare:assigned-scalar", "x = 1\nx"),
    ("bare:declared-sub", "Helper"),
    ("bare:builtin-no-args", "Beep"),

    # -- Declaring a name that already exists in the same procedure. VBA's
    #    "Duplicate declaration in current scope". The interesting half is
    #    whether an *implicitly* created variable (no `Dim`, just assigned)
    #    counts as already-existing -- if it does, the order of the two
    #    lines matters and a checker has to track it.
    ("dup:dim-twice", "Dim x As Long\nDim x As Long"),
    ("dup:assign-then-dim", "x = 1\nDim x As Long"),
    ("dup:dim-then-assign", "Dim x As Long\nx = 1"),
    ("dup:call-then-dim", "x = Helper(1)\nDim x As Long"),

    # -- The other routes a name might take into procedure scope, left
    #    unmeasured by the round above (issue #80). The parameter is the one
    #    that needed the four-element case form; the rest are here because
    #    they are the same question and cost one round trip each.
    #    `dup:param-control` is the harness check: if a signature alone
    #    breaks the wrapper, every other verdict in this group is worthless.
    ("dup:param-control", "y = x", "ByVal x As Long", " 1"),
    ("dup:param-then-dim", "Dim x As Long", "ByVal x As Long", " 1"),
    ("dup:for-counter-then-dim", "For x = 1 To 3\n    y = x\nNext x\nDim x As Long"),
    ("dup:foreach-elem-then-dim", "For Each x In rng\n    y = 1\nNext\nDim x As Long"),
    # `As Object`, not `As Long`: a `Set` onto a non-object *declared* type
    # is its own compile error, which would make a rejection unreadable.
    ("dup:set-then-dim", "Set x = New Collection\nDim x As Object"),
    ("dup:redim-then-dim", "ReDim arr(1 To 5)\nDim arr()"),
    # The last four routes above without their trailing `Dim`, so a
    # rejection can be pinned on the duplicate rather than on the route
    # statement being unacceptable on its own -- `dup:param-control` is the
    # parameter's, and `redim:plain-after-dim` is `ReDim`'s, being
    # `dup:redim-then-dim`'s two lines the other way round.
    ("dup:for-counter-only", "For x = 1 To 3\n    y = x\nNext x"),
    ("dup:foreach-elem-only", "For Each x In rng\n    y = 1\nNext"),
    ("dup:set-only", "Set x = New Collection"),

    # -- Controls. If one of these disagrees the probe itself is suspect,
    #    not the engine.
    ("control:declared-array-index", "Dim arr(5)\nx = arr(1)"),
    ("control:declared-proc-call", "x = Helper(1)"),
    ("control:plain-assignment", "x = 1"),

    # -- The built-in surface. A registry that omits any name Excel accepts
    #    here turns a working macro into a `macro check` failure, which the
    #    project's docs call the worse failure mode -- so these are measured,
    #    not recalled.
    ("builtin:MsgBox", 'MsgBox "hi"'),
    ("builtin:Debug.Print", "Debug.Print 1"),
    ("builtin:Randomize", "Randomize 1"),
    ("builtin:Beep", "Beep"),
    ("builtin:Err.Raise", "Err.Raise 5"),
    ("builtin:Application.Run", 'Application.Run "Nope"'),
    # Expression position, which is the broad surface: the registry has to
    # carry every intrinsic that can be *called for a value*, not just the
    # statement-shaped ones above.
    ("builtin:expr-MsgBox", 'x = MsgBox("hi")'),
    ("builtin:expr-Split", 'x = Split("a,b", ",")'),
    ("builtin:expr-Rnd", "x = Rnd()"),
    ("builtin:expr-Array", "x = Array(1, 2, 3)"),
    ("builtin:expr-Format", 'x = Format(1, "0.00")'),
    ("builtin:expr-CreateObject", 'x = CreateObject("Scripting.Dictionary")'),
    ("builtin:expr-Range", 'x = Range("A1")'),
    ("builtin:expr-Worksheets", "x = Worksheets(1)"),

    # -- Issue #81: a `Print` output list. Not an argument list -- `;` is a
    #    separator here and nowhere else in the grammar, and the item and
    #    the separator are each independently optional. The last three are
    #    the boundary: `;` does not generalize to other bare-argument
    #    statements, and unqualified `Print` is a statement only before a
    #    `#`. Every one measured; `core/vba/parser.rs`'s
    #    `parse_print_output_list` is built from these answers.
    ("print:semicolon", 'Debug.Print "a"; 1'),
    ("print:semicolon-chain", 'Debug.Print "a"; "b"; 1'),
    ("print:trailing-semicolon", 'Debug.Print "a";'),
    ("print:trailing-comma", 'Debug.Print "a",'),
    ("print:leading-comma", 'Debug.Print , "a"'),
    ("print:leading-semicolon", 'Debug.Print ; "a"'),
    ("print:doubled-semicolon", 'Debug.Print "a";; "b"'),
    ("print:no-separator", 'Debug.Print "a" "b"'),
    ("print:spc-and-tab", 'Debug.Print Spc(3); "a"; Tab(10); "b"'),
    ("print:not-only-debug", 'x.Print "a"; 1'),
    ("print:trailing-then-else", 'If True Then Debug.Print "a"; Else Debug.Print "b"'),
    ("print:semicolon-is-not-general", 'MsgBox "a"; 1'),
    ("print:semicolon-is-not-debugs", 'Debug.Assert "a"; 1'),
    ("print:unqualified-needs-a-hash", 'Print "a"; 1'),
]


def probe(driver, case, workdir):
    """(excel_ok, excel_detail, visi_ok, visi_detail) for one case."""
    label, snippet, sig, call_args = case
    source = build_module(snippet, sig, call_args)
    visi_ok, visi_detail = visi_verdict(source)

    if driver.driver_type == "mock":
        return None, "mock driver: Excel not invoked", visi_ok, visi_detail

    base = os.path.join(workdir, "base.xlsx")
    safe = "".join(c if c.isalnum() else "_" for c in label)
    xlsm = os.path.join(workdir, f"{safe}.xlsm")
    openpyxl.Workbook().save(base)
    wb = visi_core.Workbook.load(base)
    # Verbatim, deliberately including source that does not parse -- an
    # unparseable module is a legitimate thing to ask Excel about here.
    wb.add_macro("M", source)
    wb.save(xlsm)
    excel_ok, excel_detail = driver.verdict(xlsm)
    return excel_ok, excel_detail, visi_ok, visi_detail


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-e", "--expr", action="append", default=[],
                    help="an ad-hoc snippet to probe (repeatable); replaces the built-in list")
    ap.add_argument("--only", help="run only cases whose label contains this substring")
    ap.add_argument("--sig", default="",
                    help="parameter list for the Gen wrapper, e.g. 'ByVal x As Long'; "
                         "applies to every case that runs")
    ap.add_argument("--call-args", default="", dest="call_args",
                    help="arguments Harness passes to Gen, leading space included, e.g. ' 1'")
    ap.add_argument("--list", action="store_true", help="print the cases and exit")
    ap.add_argument("--driver", choices=["auto", "applescript", "win32com", "mock"], default="auto")
    ap.add_argument("--excel-path")
    ap.add_argument("--timeout", type=int, default=15)
    args = ap.parse_args()

    if args.expr:
        cases = [(f"expr:{i + 1}", e) for i, e in enumerate(args.expr)]
    else:
        cases = CASES
    # (label, snippet) and (label, snippet, sig, args) both normalize to four.
    cases = [c if len(c) == 4 else (c[0], c[1], "", "") for c in cases]
    if args.sig or args.call_args:
        cases = [(c[0], c[1], args.sig, args.call_args) for c in cases]
    if args.only:
        cases = [c for c in cases if args.only in c[0]]
    if not cases:
        print("no cases matched", file=sys.stderr)
        return 2

    if args.list:
        for label, snippet, sig, call_args in cases:
            print(label)
            if sig:
                print(f"    Sub Gen({sig})   called as: Gen{call_args}")
            print("    " + snippet.replace("\n", "\n    "))
        return 0

    driver = ExcelVerdictDriver(args.excel_path, args.driver, args.timeout)

    print("=" * 72)
    print(" VBA compile probe -- what does Excel's compiler accept?")
    print(f" Excel driver: {driver.driver_type} ({args.excel_path or 'default'})")
    print(f" Timeout     : {args.timeout}s (a hang is how Excel reports a compile error)")
    print(f" Cases       : {len(cases)}")
    if driver.driver_type == "mock":
        print(" MOCK DRIVER -- visi's parser runs, Excel is not consulted.")
    print("=" * 72 + "\n")

    workdir = tempfile.mkdtemp(prefix="vba_compile_probe_")
    disagreed = 0
    try:
        for case in cases:
            label, snippet, sig, call_args = case
            excel_ok, excel_detail, visi_ok, visi_detail = probe(driver, case, workdir)
            excel_s = "accept" if excel_ok else "REJECT" if excel_ok is False else "n/a"
            visi_s = "accept" if visi_ok else "REJECT"
            agree = "" if excel_ok is None or excel_ok == visi_ok else "   <-- DISAGREE"
            if agree:
                disagreed += 1
            print(f" {label:<34} excel={excel_s:<6} visi={visi_s}{agree}")
            if sig:
                print(f"     Sub Gen({sig})   called as: Gen{call_args}")
            print(f"     {snippet.replace(chr(10), chr(10) + '     ')}")
            if not visi_ok:
                print(f"     visi : {visi_detail}")
            if excel_ok is False:
                print(f"     excel: {excel_detail}")
            print()
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    print("=" * 72)
    print(f" {len(cases) - disagreed}/{len(cases)} agreed ({driver.restarts} Excel restarts)")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
