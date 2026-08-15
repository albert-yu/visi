# VBA macro support: feasibility and test plan

Investigation for [issue #46](https://github.com/albert-yu/visi/issues/46).

Last updated: 2026-08-14. Findings verified against Microsoft Excel for Mac
16.112 on macOS 26.6.1.

Two questions:

1. What would it take to actually *execute* VBA macros, rather than only
   store them?
2. Can that execution be differentially fuzzed against real Excel through
   AppleScript, the way `fuzz/` already tests formulas, charts, and pivots —
   or does it need a Windows COM host?

Question 2 is settled and the answer is **AppleScript, no COM required**;
the evidence is in [Part 2](#part-2-testing-against-real-excel), including a
working end-to-end round trip. Question 1 is a real but bounded engineering
project, scoped in [Part 1](#part-1-the-interpreter).

---

## Where things stand today

The *storage* half of VBA support is already done and shipped:

| Piece | Module | What it does |
| --- | --- | --- |
| Data model | `core/vba.rs` | `VbaProject` / `VbaModule` (Standard, Class, Document), workbook-level like `Chart` and `PivotTable` |
| Compression codec | `core/ovba.rs` | MS-OVBA compressed-container encode/decode |
| xlsx I/O | `core/vba_xlsx.rs` | Reads and patches `vbaProject.bin` inside the zip |
| From-scratch synthesis | `core/vba_synth.rs` | Builds a valid `dir` / `_VBA_PROJECT` skeleton and module p-code prefix with no donor file |
| CLI | `visi macro {list,add,remove,rename,set-source}` | Module CRUD |

So `visi` can author a `.xlsm` containing arbitrary VBA source that real
Excel accepts, loads, and runs — verified below. What it cannot do is run
that source itself. `visi macro` is a text editor for a binary blob.

That existing capability turns out to be exactly what makes the test plan
work, so the two halves of this issue are more connected than they look:
**the fuzz harness needs to inject arbitrary VBA into a workbook, and Excel
itself offers no automation path to do that on macOS. `visi` is the only
reason it's possible.**

---

## Part 1: the interpreter

### What "VBA support" would actually mean

VBA is not an extension of the formula language. It is a separate,
statement-oriented, Turing-complete language with its own value model,
control flow, user-defined procedures, and an object model bound to the
host application. Almost none of `core/parser.rs` or the `Expr` evaluator is
reusable:

| | Formula language | VBA |
| --- | --- | --- |
| Shape | one expression | statements, blocks, procedures, modules |
| Values | `ResultData` (7 variants) | `Variant` over ~12 subtypes, plus `Object` and `Nothing` |
| Identity | cell refs by `sheet_id`/`col_id` | named variables with declared scope and lifetime (`Dim`, `Static`, `Public`) |
| Operators | `+ - * / ^ & = <> < > <= >=` | adds `\` (integer divide), `Mod`, `Like`, `Is`, `Eqv`, `Imp`, `Xor`, `Not`, `And`, `Or` with a different precedence table |
| Errors | a `String` error code like `#VALUE!` | a runtime error object (`Err.Number`, `Err.Description`) with `On Error` handlers and resumption |
| Effects | pure, returns a value | arbitrary mutation of the workbook, plus `MsgBox`, files, other applications |

What *is* reusable, and it is the expensive part of any spreadsheet
language, is the function library: `Application.WorksheetFunction.X` is
`evaluate_function`'s `X` in almost every case, so `stats`, `math_trig`,
`text`, `date_fn`, `finance`, and `engineering` all come for free behind a
thin adapter.

### Proposed shape

Four new `pub(crate)` modules under `core/vba/`, keeping the existing
`vba.rs` data model as the storage layer beneath them:

```
core/vba/lexer.rs    tokens; line continuations (_), comments (' and Rem),
                     case-insensitive keywords, line numbers/labels
core/vba/ast.rs      Module → Procedure → Statement → Expr
core/vba/value.rs    Variant: the value model and VBA's coercion rules
core/vba/interp.rs   tree-walking interpreter over the AST
core/vba/host.rs     the object model bridge onto WorkbookManager
```

A tree-walking interpreter is the right call here, not a bytecode VM. Macro
workloads are dominated by host round trips (`Range(...).Value`), not by
interpretation overhead, and the AST maps directly onto the language.

### The host object model is where the real risk lives

The language itself is a known quantity — a few thousand lines for a
competent lexer, parser, and evaluator. The unbounded part is the object
model, because "supporting VBA" colloquially means supporting everything a
macro might touch. A deliberately narrow, explicitly-enumerated surface is
the only way to keep this finite:

**Phase 1 (worth building):**
`Application.WorksheetFunction.*` · `ThisWorkbook` / `ActiveWorkbook` ·
`Worksheets(name|index)`, `.Name`, `.Count` · `Range("A1")`, `Range("A1:B2")`,
`Cells(r, c)` · `.Value`, `.Value2`, `.Formula`, `.Text`, `.Row`, `.Column`,
`.Count`, `.Address` · `.Offset`, `.Resize` · `For Each` over a Range ·
built-in functions (`CStr`, `CLng`, `CDbl`, `Len`, `Mid`, `InStr`, `Format`,
`IsNumeric`, `Array`, `UBound`, …)

**Phase 2 (plausible later):**
`.Interior.Color` / `.Font.Bold` onto `CellStyle` · `ListObjects` onto
`ExcelTable` · `PivotTables` onto `PivotTable` (`.RefreshTable` maps onto
`refresh_pivot_table`) · `Rows`/`Columns` insert and delete · `Sort`, `Find`

**Explicitly out of scope, and this must be a documented refusal rather than
a silent no-op:** UserForms and all of MSForms, `MsgBox` / `InputBox` and
anything else modal, `CreateObject` and COM automation of other
applications, `Declare` and Win32 API calls, file and network I/O,
`Application.OnTime` and event-driven execution, `Shell`.

The refusal matters more than the feature list. A macro that silently skips
a `CreateObject` line and then reports success has produced a wrong answer
in the most dangerous way available. Every unsupported construct should
raise a distinct, catchable interpreter error naming what it was.

### Sequencing

A phase is only worth starting once the one before it is fuzz-clean against
real Excel, using the harness in Part 2.

| Phase | Scope | Fuzzable when done? |
| --- | --- | --- |
| 0 | Lexer + parser + AST; `visi macro check` reports syntax errors, no execution | Parse-only: does Excel agree this compiles? |
| 1 | `Variant` model, expressions, `If`/`For`/`Do`/`Select Case`, `Sub`/`Function`, `On Error` | Yes — the bulk of the differential value |
| 2 | Range and Worksheet object model, `WorksheetFunction` bridge | Yes |
| 3 | Styles, tables, pivots, row/column edits | Yes |
| — | Events (`Worksheet_Change`, `Workbook_Open`), classes | Separate design; ordering is observable and Excel's is subtle |

Phase 0 alone has standalone value: `visi macro add` currently accepts
source that Excel will reject at compile time, and there is no way to find
out short of opening the file in Excel.

### Security posture

Executing a macro from an untrusted workbook is a materially different act
from reading one. Even with I/O out of scope, a macro can rewrite the
workbook arbitrarily, and `visi` is a CLI that scripts and CI will point at
files they did not author. Execution must be opt-in per invocation
(`visi macro run FILE --name Foo`), never implicit in `eval`, never
triggered by opening a file, and never wired to `Workbook_Open` without a
separate explicit flag. `--quiet` must not suppress the notice that a macro
ran.

---

## Part 2: testing against real Excel

### The question, and the answer

Issue #46 asks whether AppleScript suffices or a COM API is required.
**AppleScript suffices.** No Windows host, no COM, no `pywin32`.

This is not a foregone conclusion, and the reason it works is specific.
Excel for Mac's AppleScript dictionary exposes **no** VBProject object — you
cannot create, edit, or import a VBA module through automation on macOS.
(The Windows equivalent, `VBProject.VBComponents.Import`, additionally
requires the user to tick "Trust access to the VBA project object model",
which has no macOS counterpart.) So the obvious harness design — automate
Excel into containing the macro you want to test — is unavailable.

`visi macro add` sidesteps this entirely by writing the module into
`vbaProject.bin` at the file-format level, before Excel ever sees the file.
Excel then loads it as a perfectly ordinary macro. The `run VB macro`
AppleScript command, which *is* supported and is already load-bearing for
`fuzz/fuzz_pivot.py`, invokes it.

### Verified end-to-end

Every step below was executed, not reasoned about, and the whole sequence is
checked in as a repeatable script:

```bash
cargo build --release
source fuzz/venv/bin/activate
python fuzz/vba_probe.py          # 4 checks, ~15s against real Excel
```

Re-run it after any change to `vba_xlsx.rs` / `vba_synth.rs`, or after an
Excel update, before trusting anything below.

```bash
# 1. Author a .xlsm with a macro, using only visi — no donor file, no Excel
visi macro add base.xlsx --name VisiProbe --kind standard \
    --source-file mod1.bas --output probe.xlsm
```

```vb
' mod1.bas — arithmetic, a loop, string ops, a cell read, a cell write
Attribute VB_Name = "VisiProbe"
Public Sub RunProbe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Dim i As Long, total As Double
    For i = 1 To 5
        total = total + i * i
    Next i
    ws.Range("C1").Value = total
    ws.Range("C2").Value = "hello " & CStr(Len("abcd"))
    ws.Range("C3").Value = ws.Range("A1").Value + ws.Range("A2").Value
    ThisWorkbook.Save
End Sub
```

```applescript
-- 2. Real Excel opens it and runs it
tell application "Microsoft Excel"
    set display alerts to false
    open POSIX file "/…/probe.xlsm"
    run VB macro "RunProbe"
    close active workbook saving no
end tell
```

Results, read back with `openpyxl` and then again with `visi read`:
`C1 = 55`, `C2 = "hello 4"`, `C3 = 3` — all correct. `visi macro list` on the
Excel-saved file still lists `VisiProbe` (alongside a `ThisWorkbook` document
module Excel added itself), so the round trip survives Excel's own save.

Three further properties were confirmed, each of which shapes the harness:

**Macros can return values directly to AppleScript.** `run VB macro "Eval"
arg1 "21"` returned `OK|Double|42` as an AppleScript string. The harness does
not need to route results through cells and a file read — it can get a
typed result back in-process. (Cells remain the right channel for testing
*mutations*.)

**Trapped runtime errors come back as structured data.** A `1/0` under
`On Error GoTo` returned `ERR|11|Division by zero`. Excel's error numbers and
descriptions are therefore directly comparable against the interpreter's,
which makes error *behaviour* fuzzable, not just success values — the same
way `#VALUE!` vs `#DIV/0!` is already compared for formulas.

**An *un*trapped runtime error hangs the automation bridge.** This one is
the sharp edge. `CLng("not a number")` with no handler pops a modal
run-time-error dialog that `set display alerts to false` does not suppress;
the `osascript` call never returns and Excel must be SIGKILLed. A fuzzer
generating random VBA will hit this constantly.

The mitigation is verified to work: wrapping the generated procedure in a
harness entry point catches errors raised anywhere down the call stack.

```vb
Public Function Harness() As String
    On Error GoTo Failed
    Dim r As Variant
    r = Generated()                      ' the fuzzer's generated code
    Harness = "OK|" & TypeName(r) & "|" & CStr(r)
    Exit Function
Failed:
    Harness = "ERR|" & CStr(Err.Number) & "|" & Err.Description
End Function
```

With that wrapper, the same unhandled `CLng("not a number")` inside
`Generated` returned `ERR|13|Type mismatch` instead of hanging. **Every
generated macro must go through this wrapper — it is a correctness
requirement of the harness, not a nicety.** The generator must also never
emit `On Error GoTo 0` or `Resume` in a position that could escape the
wrapper, and the AppleScript call needs a hard timeout with the
force-restart recovery `fuzz_pivot.py` already implements.

### Harness design

`fuzz/fuzz_vba.py`, modelled on `fuzz_excel.py` and reusing
`fuzz_pivot.py`'s AppleScript session management verbatim:

```
        ┌──────────────────────────────────┐
        │  fuzz_vba.py — VBA generator     │
        │  (typed AST → source text)       │
        └────────────────┬─────────────────┘
                         │  wrapped in Harness(); visi macro add
                         ▼
                   source.xlsm  (base data + generated module)
            ┌────────────┴────────────┐
            ▼                         ▼
   visi VBA interpreter        Microsoft Excel
   (visi_core bindings)        (AppleScript run VB macro)
            │                         │
            ▼                         ▼
  "OK|Double|42" / "ERR|11|…"   "OK|Double|42" / "ERR|11|…"
     + mutated cell values         + mutated cell values
            └────────────┬────────────┘
                         ▼
             compare return string, error
             number, and every touched cell
```

**Generate a typed AST, not text.** The lesson already recorded in
`fuzz/README.md` and in the fuzz-harness memory notes — a generator that
emits raw strings produces files Excel silently corrupts or refuses, and the
failure is attributed to the engine. Generating from a typed AST with
declared variable types lets the generator guarantee, by construction, that
every variable is `Dim`'d before use, that names avoid VBA's reserved words,
that `For` loops terminate, and that no procedure recurses without a depth
bound.

**Grow the grammar with the phases.** Phase 1 needs only expressions over
`Long`/`Double`/`String`/`Boolean` variables, arithmetic and comparison
operators, `If`/`For`/`Do`/`Select Case`, and procedure calls — no object
model at all, with the result read purely from `Harness()`'s return string.
That subset alone will find a large fraction of the coercion and precedence
bugs, and it needs zero workbook interaction, which makes the Excel side
fast and its failure modes simple.

**Comparison rules.** Reuse `fuzz_excel.py`'s float tolerance and its
Excel-quirk exclusions wholesale — VBA arithmetic is the same IEEE double
arithmetic with the same 15-digit display rules, and `CDbl`/`CStr` round-trip
through the same formatter. Two additions: compare `Err.Number` exactly (it
is a small enumerated set), and compare `TypeName(r)` — VBA's *static* type
is observable and independently wrong-able, so an interpreter that computes
`42` as a `Double` where Excel says `Long` has a real bug worth catching.

**Cost per iteration.** Each iteration is one Excel `open` → `run VB macro`
→ `close`, which `fuzz_pivot.py` measures in the low seconds. That is 3–4
orders of magnitude slower than the in-process side, so the harness should
batch many generated procedures into one workbook and one Excel session,
invoking `Harness()` once per procedure — the AppleScript round trip, not
the macro, is the cost.

### Known limits of this approach

- **Session degradation is real and already documented.** `run VB macro`
  degrades into a config-independent "Parameter error (-50)" after enough
  consecutive invocations against one long-lived Excel process. The fix is
  `fuzz_pivot.py::_restart_excel` (SIGKILL by PID, not `killall` alone,
  which Excel can intercept); a VBA fuzzer must adopt it from day one, since
  it will drive far more invocations per run than the pivot fuzzer does.
- **Modal dialogs are unsuppressable in general.** The wrapper handles
  runtime errors, but a *compile*-time error in generated source, or a
  `MsgBox` the generator emits by accident, produces the same hang. The
  generator must never emit `MsgBox`, and a hard `osascript` timeout is the
  backstop for everything else.
- **macOS-only, as with every other driver in `fuzz/`.** A COM driver
  remains worth adding eventually for the same reason `fuzz_pivot.py` has
  one — Windows Excel is a genuinely different implementation and has
  disagreed before — but it is not a prerequisite.
- **The Excel-authored `ThisWorkbook` module appears on save.** Harmless,
  but the comparator must not treat the module list as an invariant.

### Side finding: `fuzz_pivot.py`'s manual setup step is now removable

`fuzz_pivot.py` currently requires a one-time, human-performed ritual —
open Excel, open the VBA editor, paste `fuzz/BuildFuzzPivot.bas` into a
module, Save As `fuzz/pivot_macro_template.xlsm` — because when it was
written there was no way to get a macro into a workbook programmatically.
There is now. `visi macro add base.xlsx --name BuildFuzzPivot --source-file
fuzz/BuildFuzzPivot.bas --output template.xlsm` produces the same file with
no human in the loop, which would let the template be built on demand
instead of being an uncheckable, uncheck-in-able prerequisite. Worth doing
independently of any interpreter work; not done here to keep this change
to investigation only.

---

## Recommendation

Split this issue. The investigation it asked for is complete; what remains
is implementation, and it is too large for one issue.

1. **Do Phase 0 first** (lexer, parser, `visi macro check`). It is
   self-contained, immediately useful even with no interpreter behind it,
   and it de-risks the grammar before any semantics depend on it.
2. **Build `fuzz/fuzz_vba.py` against Phase 1, not after it.** The
   AppleScript path is proven; standing the harness up early means the
   `Variant` coercion rules — historically where the Excel-parity bugs
   concentrate, per `docs/excel-discrepancies.md` and the coercion work in
   `core/engine` — get differential coverage from their first commit.
3. **Keep the object model on a published allow-list**, with an explicit
   error for everything outside it.
4. **Treat execution as opt-in and privileged** from the first line of code
   that can run a statement, not as a hardening pass afterwards.
