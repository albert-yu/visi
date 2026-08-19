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
| 0 ✅ | Lexer + parser + AST; `visi macro check` reports syntax errors, no execution | Parse-only: does Excel agree this compiles? |
| 1 ✅ | `Variant` model, expressions, `If`/`For`/`Do`/`Select Case`, `Sub`/`Function`, `On Error` | Yes — the bulk of the differential value |
| 2 ✅ | Range and Worksheet object model, `WorksheetFunction` bridge | Yes |
| 3 | Styles, tables, pivots, row/column edits | Yes |
| — | Events (`Worksheet_Change`, `Workbook_Open`), classes | Separate design; ordering is observable and Excel's is subtle |

Phase 0 alone has standalone value: `visi macro add` accepts source that
Excel will reject at compile time, and before `visi macro check` there was no
way to find out short of opening the file in Excel.

### Phase 0, as built

Landed as `core/vba/{lexer,ast,parser}.rs`, the `visi macro check` subcommand,
`visi_core.check_syntax` in the bindings, a `vba_parse` cargo-fuzz target, and
`fuzz/fuzz_vba_parse.py`. Three findings from building it are worth carrying
forward, because they constrain the later phases too.

**VBA's operator precedence was pinned against real Excel, not documentation.**
Two cases discriminate between plausible tables and were run to get the
answer: `False Imp False Eqv False` is `True` only if `Eqv` binds tighter than
`Imp`, and `True Xor True Eqv False` likewise for `Xor` over `Eqv`. The one
most likely to be got wrong is `^`: `2 ^ 3 ^ 2` is **64**, so it is
*left*-associative, unlike almost every other language with an exponent
operator. `-2 ^ 2` is `-4`, so `^` also binds tighter than unary minus. The
table and its confirming cases are documented at the top of `parser.rs`, and
each has a unit test naming the Excel result.

**Excel compiles VBA lazily, strictly per invoked procedure.** This was found
while building the differential harness and is the single most awkward fact
about testing any of this. A trivial probe procedure in the same module as
broken code compiles and runs happily; so does a procedure that *references*
the broken one from a dead branch. Only invoking a procedure compiles it. The
harness gets around it by putting generated source inside
`If False Then ... End If` in the procedure it invokes — Excel must compile
the body and cannot execute it (verified with an `Err.Raise` inside the dead
branch, which never fired). A consequence worth stating plainly: **there is no
way to ask Excel "does this arbitrary module compile?" without also asking it
to run something.** `fuzz_vba_parse.py --corpus` is therefore parser-only by
design rather than by omission.

**Diagnostics blame the opener, matching VBA.** An unclosed block reports at
the line that opened it, with VBA's own wording (`Block If without End If`,
`For without Next`, `Expected End Sub`, ...), rather than at whatever token
turned up where the closer was due — which is usually a perfectly correct
line, and in a long procedure can be hundreds of lines from the defect. The
wording and position follow VBA's documented compile errors; they could not be
read back from Excel directly, since a compile error surfaces only as a modal
dialog that is unreadable to the AppleScript bridge and, in this environment,
to UI scripting and screenshots as well. Worth re-checking by hand in the VBE
if these strings ever matter for more than readability.

**A compile error is observable only as a hang**, as predicted in Part 2 —
and, unlike a runtime error, the `On Error` wrapper does not help, because a
compile error is not trappable. The harness reads a timeout as Excel's
rejection, and confirms it survives an Excel restart before believing it, so
session degradation is not mistaken for a verdict.

The harness's first real run found a genuine disagreement, and it is a useful
one to have on record: `x = f(1, , 3)` — an omitted middle argument — parses
fine and is valid VBA, but Excel rejects it at compile time when it cannot
resolve `f` to a procedure with an `Optional` parameter in that position.
That is a *semantic* check requiring name resolution, which Phase 0
deliberately does not do; the parser cannot even tell a procedure call from an
array index without a symbol table. So it marks the boundary of what
parse-only checking can see rather than a parser bug, and the generator now
declares a real callee so the omitted-argument syntax stays under test.

#### The symbol table, as built (issue #78)

Porting the harness to Windows (`--driver win32com`) put that boundary under
real measurement for the first time and found **12 false negatives in 100**
generated cases. `core/vba/resolve.rs` and `core/vba/builtin_names.rs` are the
answer, and the headline is that the issue's own diagnosis — "essentially all
of them trace to this one limitation" — was **wrong about half of them**.
Minimizing each case with the new `fuzz/vba_compile_probe.py` split them four
ways:

| Cause | Cases |
| --- | --- |
| an undeclared name used with call syntax | iter_22, 40, 50, 64, 93 |
| a plain local (`Dim x`, or `x = 1`) used as a call target | iter_12, 15, 46 |
| `ReDim Preserve` on an undeclared name | iter_14, 57 |
| a name starting with `_` | iter_24 |
| a built-in type keyword as a declared name (fixed earlier) | iter_4 |

Only the first two groups are name resolution. The last two were assumed to be
and are not: **plain `ReDim arr(1 To 5)` declares the array and compiles with
no `Dim` anywhere, while adding `Preserve` makes the identical line a compile
error**, and a leading `_` is a pure lexer defect (VBA has no identifier
starting with an underscore, so `_ y = 2` was silently becoming an
implicit-call statement on a name spelled `_`). Two cases that looked
implicated — `For Each c In rng` and `x = a _ + b`, both with everything
undeclared — turned out to **compile fine** and were never the cause; with
`Option Explicit` off a bare name is just an implicit Variant. **Only call
syntax forces Excel to resolve anything.**

The design risk the issue flagged is real and is handled structurally rather
than by care: Excel compiles a *project*, so a checker holding one module of
several genuinely cannot tell an undeclared name from a cross-module call.
`resolve::Scope::complete_project` gates the whole undeclared-name rule —
`check_syntax` sets it (a standalone `.bas` is the whole story), the new
`VbaProject::check_modules` sets it after unioning every module's declared
names, and `VbaModule::check_syntax` deliberately does not, having no project
to consult. `visi macro check` on a workbook goes through `check_modules`, so
a call into a sibling module resolves.

One exposure is left, and it is a deliberate choice rather than an oversight:
`visi macro check some_module.bas` treats that file as the whole project, so a
call into a module the file does not contain is reported. That is right for the
differential harness (whose generated module *is* self-contained) and for a
genuinely standalone `.bas`, and wrong for one file cut out of a larger
project — where the answer is to check the workbook instead, which resolves
project-wide. If this turns out to bite in practice the fix is a flag selecting
the scope, not a weakening of the default.

The built-in registry is deliberately **over-broad**, and the asymmetry is
worth stating because it is what makes a hand-maintained list of a few hundred
names defensible at all: omitting a real built-in rejects working code, while
including a non-built-in merely lets one mistake through — which is exactly
today's behaviour and so no regression. Adding a plausible name there needs no
justification; removing one does. The registry is locked to `builtins.rs` from
both ends by a pair of tests, because auditing the two lists against each other
found `Hex`, `Oct` and `Val` implemented as intrinsics but missing here — three
names that would each have rejected working code.

**Validation: two unseen seeds, 60 cases each — 0 false positives on both**
(4242 and 91177), with false negatives down from 12-in-100. Zero is the number
that matters: the undeclared-name rule is the first thing here that can reject a
module for a reason other than its own syntax, and a single false positive would
make `macro check` untrustworthy in the way the plan has warned about
throughout.

The differential harness cannot see the whole risk, though, and that is worth
recording. The generated grammar never emits a **type-declaration suffix**, so
it could not catch the one real false positive this work introduced: the lexer
folds `$` into an identifier's spelling, so resolving `Trim$` looked up
`"trim$"`, found nothing, and rejected it — along with every other `$` string
intrinsic (`Left$`, `Mid$`, `Format$`, `UCase$`). It took hand-written
real-world VBA to surface, and `resolve::norm` now strips a suffix on the
declaring and referencing side alike. The lesson generalises: a fuzzer bounds
the false-positive rate *over the shapes it generates*, which is narrower than
the shapes real modules contain.

The two survivors were both *new* shapes, neither in the original 12:

- **A bare identifier alone as a statement** (`x`, no parentheses, no
  arguments). Deliberately left unchecked in the first pass on the grounds that
  it had never been measured — the fuzzer then measured it, and the probe
  settled every variant: **a bare statement is a call, and resolves by exactly
  the same rule**. `x` is rejected whether undeclared, `Dim`'d as a `Long`, or
  created by assignment; `Helper` (a declared `Sub`) and `Beep` (a built-in)
  are accepted. Now implemented. Note this holds in *statement* position only —
  a bare `x` inside an expression is an ordinary implicit-Variant read, which
  is why `x = a + b` with nothing declared still compiles.
- **Duplicate declaration**: `x = Helper(1)` creates `x` implicitly and a later
  `Dim x As Long` in the same procedure redeclares it. On seed 91177 **all five**
  remaining false negatives were this one shape, not five different gaps —
  ordinary code declares first, while the generator emits its `Dim` last, so the
  shape dominates the residual rate. Now implemented, and it is the one
  order-sensitive rule in the pass: measured, `Dim x` then `x = 1` compiles and
  `x = 1` then `Dim x` does not, so counting names is not enough — the checker
  has to know which came first.

  The first pass counted only two routes into scope — an explicit declaration
  and a plain assignment — because a parameter could not be put to Excel at
  all: the probe splices its snippet into a parameterless `Sub`. Issue #80
  closed that by giving `build_module` a signature (`sig`/`args`, threaded
  through to `Sub Gen(...)` and its call in `Harness`, which has to keep
  invoking it or the procedure is never compiled). With that, all five
  remaining routes measured as **compile errors** — a parameter, a `For`
  counter, a `For Each` element, a `Set` target, and a plain `ReDim` — each
  confirmed against a control running the same route *without* the trailing
  `Dim`, so the rejection is the duplicate and not the route statement itself.
  All five are now flagged. Only a `Dim`/`Static`/`Const` reports: a `ReDim` of
  a name already in scope is the ordinary resize and compiles, so the routes
  add to the scope set without ever colliding themselves. One route is still
  unmeasured — the procedure's own name, assignable inside a `Function` — since
  the harness compiles one fixed `Sub`.

A harness bug found on the way: `fuzz_vba_parse.py` saved failures to
`fuzz_results/failures/vba_parse_iter_<N>` with no seed in the path, so a second
run silently overwrote the first's reproductions — which is what happened to the
issue #78 corpus while it was in use as a reference set. It also passed
`--seed`'s `None` straight to `random.Random`, seeding from entropy and making
any failure it found unreproducible. Both fixed: the seed is now resolved up
front, printed, and included in every saved failure's directory name, matching
`fuzz_excel.py`'s existing `fail_iter_<N>_seed_<SEED>` convention.

### Phase 1, as built

Landed as `core/vba/{value,interp,builtins}.rs`, `visi macro run`,
`visi_core.run_macro` in the bindings, and `fuzz/fuzz_vba.py`. The interpreter
covers expressions, `If`/`For`/`Do`/`While`/`Select Case`, `Sub`/`Function`
calls with recursion, `On Error` in all its forms, and ~45 host-free
intrinsics. There is **no host object model** — anything touching a workbook
raises error 438 naming what it was, rather than being skipped.

**The `Variant` rules were measured, not assumed.** `fuzz/vba_variant_probe.bas`
returns `TypeName(v) & "|" & CStr(v)` for 64 expressions; every rule in
`value.rs` cites the case it came from. Several are the opposite of the
obvious guess: `"1" + 1` is a `Double` but `"1" + "2"` is a `String`; `7.6 \ 2`
is `4` and typed `Long`, because the operands round before dividing;
`CLng(2.5)` is `2` because every conversion is banker's rounding; `Null`
propagates through arithmetic but `&` skips it.

**The fuzzer immediately corrected a rule the probe had got wrong**, which is
the clearest possible argument for building it alongside the phase rather
than after. The probe showed `32767 + 1` raising error 6 and I encoded
"Variant arithmetic never promotes". The fuzzer disagreed, and a follow-up
probe showed why: **only arithmetic between two compile-time constants uses
fixed widths and overflows.** With a variable on either side it widens —
`a = 32767 : a + 1` is the `Long` 32768, `a = 100000 : a * a` is the `Double`
1e10. `value::ArithMode` now carries which of the two applies, decided by
whether both operand expressions are literal.

Two smaller corrections came from the same run: Excel's `CStr` renders
negative zero as `-0` (normalising it away was wrong), and an empty string is
*not* a zero — `"" = 0`, `"" < 0` and `Not ""` are all error 13, unlike
`Empty`, which genuinely is zero.

**Where it stands:** 500 generated procedures, **493 agreeing** with Excel on
value, subtype and error number together — up from 54 of 60 when the harness
first ran. Every divergence family found in that first run has been resolved,
and each fix is a measured rule with a test naming the Excel result:

| Was diverging | What Excel actually does |
| --- | --- |
| `And` / `Or` / `Imp` with `Null` | Three-valued. A *falsy* operand decides `And` and a *truthy* one decides `Or`, and the deciding operand is returned — converted to the integer type the bitwise operator works in, so `vb Or Null` with `vb = 0.1 - 2147483647` is the `Long` `-2147483647`. `Imp` needed no table at all: evaluating it as `Not a Or b` gets this for free, and fixed a case the hand-rolled table got wrong (`255 Imp Null` is `-256`, not `Null`). |
| string vs number comparison | Four rules, split by constant-ness exactly as arithmetic is. Both constant → numeric, error 13 if the string will not parse. Numeric constant → numeric, falling back rather than erroring. String constant → string comparison. **Both variables → a number always sorts before a string**, which is why `a = "1.5"`, `b = 1.5` makes `a = b` `False` even though they are equal both numerically and as text. |
| `For` counter after the loop | Left at the value that failed the test (`For i = 1 To 3` leaves `4`; `Step 2` leaves `5`), or at the current value on `Exit For`. Assigning the counter *before* the test rather than after is the whole fix. |
| `Space` / `String` / `Mid` counts | Rounded, not truncated: `Space(2.6)` is three spaces. |
| `(-1) ^ 1.5` | Error 5, not a quiet `NaN`. |
| `Select Case Null` | `Case 2 To 5` **matches** a `Null` subject, while `Case 0, 1` and `Case Is > 2` do not — and nothing about the comparisons predicts it, since `Null >= 2` is `Null`. An Excel quirk, matched deliberately. |
| infinity | `255 ^ 255` is `INF` and negating it gives `-INF`, but `+`, `-` and `*` raise error 6 if either side is infinite. `CStr` renders it `"INF"`. |
| `CStr(Null)` | Error 94, not a propagated `Null`. Propagating it silently poisoned callers who had it under `On Error Resume Next` expecting the assignment to be skipped. |
| `Single` with `Long` | Widens past both to `Double`, since a `Single` cannot hold every `Long` — though `Single` with `Integer` stays `Single`. |
| `"abc" - Null` | Error 13, not `Null`: the operand is coerced *before* the `Null` short-circuits. |
| `Val(255)` | An `Integer`. `Val` types its result like a literal rather than always returning a `Double`. |

Seven long-tail cases remain, all with saved reproductions under
`fuzz_results/failures/`. None is a wrong *value* — they are disagreements
about which of two errors surfaces, or about whether `Null` propagates. All seven were
root-caused and fixed in [`vba-error-ordering.md`](vba-error-ordering.md) —
one of them turned out to be a bug in the fuzz driver rather than in either
engine. That document is also where the honest number lives: on *unseen*
seeds the harness reports 294–297 of 300, not the 300/300 the tuned seed
shows, and the gap between those two figures is worth reading before trusting
any of these counts.

### Phase 2, as built

Landed as `core/vba/host.rs`, `WorkbookManager::run_macro`, the `--output` /
`--in-place` half of `visi macro run`, `Workbook.run_macro` in the bindings,
and a `fuzz/fuzz_vba.py` that compares cells as well as return values. The
allow-list is the one this document proposed and nothing beyond it; everything
else still raises 438 naming the construct.

`fuzz/vba_host_probe.py` is the measurement record — a hundred-odd fixed
questions run against a workbook with data in it — and every rule below cites
it. The inline tests in `host.rs` assert the exact strings it got back from
Excel, so an expectation there can be read straight off a probe run.

**An object is a value, not a pointer.** `ObjRef` holds ids and coordinates,
never a borrow, so the interpreter can hold `&mut WorkbookManager` for the
whole run without fighting the borrow checker at every statement. A `Range` is
`(sheet_id, row, col, height, width)` and a `Worksheet` is a `sheet_id` — an
id rather than an index or a name, for the same reason compiled formulas store
one: a macro can rename or reorder sheets mid-run.

**`Is` compares an identity token, and this was the one design the plan got
wrong.** The issue proposed comparing the coordinate tuples. Excel disagrees:
`ws.Range("A1") Is ws.Range("A1")` is **False**, because each call constructs
a fresh object, while `Set r = ws.Range("A1") : Set q = r` makes `q Is r`
True. So a `Range` carries a token that copying preserves and reconstructing
does not. Worksheets go the other way — `ws Is wb.Worksheets(1)` is True —
because Excel hands out a cached object per sheet, so a worksheet's identity
really is its id.

**A read recalculates only if a write is outstanding.** Excel in automatic
mode recalculates after every assignment and it is observable: writing `A1`
and then reading a `D1` holding `=A1*2` gives the new value. Doing that
literally would mean running `WorkbookManager::evaluate` — three passes over
every sheet — once per assignment, which a loop writing a thousand cells
cannot afford. A write instead sets a `stale` flag and the next read that
could observe it pays for one recalculation, so a run of consecutive writes
costs one rather than one each. Same observable behaviour, different timing.
The deeper staleness risk is `evaluate`'s own fixed three-pass limit, which a
macro-driven write makes observable where it previously was not.

**Default members are resolved in the interpreter, and are easy to get
silently wrong.** `x = ws.Range("A1")` reads the cell while
`Set r = ws.Range("A1")` binds the object, and the parser sees the same
expression for both. Every context wanting a scalar funnels through one
`Interpreter::scalar`; `Set`, `Is`, `TypeName`, a `With` subject and a user
procedure's arguments deliberately skip it. Getting this wrong does not raise
— it produces the wrong kind of value.

**Dates come back as `Date`, and that needed a real subtype.** A cell whose
style carries a date number format reads back through `.Value` as a `Date`
variant and through `.Value2` as a plain `Double` — measured, including for a
*fractional* serial, which is still a `Date` and whose `CStr` shows the time.
`Variant::Date` is a VBA-side type only; the engine keeps having no date value
type, exactly as `core/date.rs` argues. Writing a `Date` back writes the
serial *and* the number format, since a date in this model is both.

Other measurements worth knowing before touching this:

| Behaviour | Measured |
| --- | --- |
| `For Each` over a range | **row-major** — `A1 B1 A2 B2` over `A1:B2` |
| a numeric cell's `.Value` | always `Double`, never `Integer`, even for `1` |
| a multi-cell `.Value` | a 2-D array indexed `(row, column)` from 1; `CStr` of it is error 13 |
| an array assigned to one cell | writes its first element |
| `Worksheets("nope")`, `Worksheets(5)` | error 9 |
| a bad address, an off-sheet `Offset`, `Resize(0, 1)` | error 1004 |
| `ws.Cells.Count` | **error 6** — the property is `Long` and the grid does not fit |
| `ws.Cells.Address` | `$1:$1048576`, the whole-row form |
| `.Value = "=G1*3"` | makes a *formula*; `.Value = "6/22/2026"` makes a *date* |
| a cell holding `=1/0`, read | an error `Variant`; `CLng` of it is 2007 |
| `WorksheetFunction.VLookup` failing | raises 1004 |
| `Application.VLookup` failing | returns an error `Variant` that `IsError` detects |
| `Sum` over a range | skips text and booleans; `Sum("1", 2)` coerces them |

That last pair is the whole reason `Application` and
`Application.WorksheetFunction` are two objects over one implementation.

**The bridge is a bridge.** `Application.WorksheetFunction.X` builds formula
argument expressions — a `Range` becomes a real range reference — and hands
them to `Sheet::call_worksheet_function`, a `pub(crate)` entry point onto the
existing `evaluate_function`. That is what makes the range-versus-scalar
coercion split above fall out rather than being written twice. The engine's
own non-Excel functions (`GET`, `GET_COL`, `GET_COL_IDX`, `SLICE`, `STR`) are
blocked: a macro using one would work here and fail in Excel, which is the one
direction of divergence a differential harness cannot catch, since it only
generates what Excel accepts.

**Two limits are ours, not Excel's.** Excel's grid is sparse and `visi`'s
`Sheet` is a dense `Vec` per column, so `ws.Range("XFD1048576").Value = 1`
would ask for seventeen billion cells. It reports error 7 ("Out of memory")
past four million instead — a number VBA itself produces, so a macro sees
something plausible. Reading or iterating a range that large is capped the
same way.

**Where it stands.** On eight seeds never used while developing,
**1599 of 1600** generated procedures agree with Excel on value, subtype,
error number *and* every cell in the data grid, and a later round of seventeen
more unseen seeds stands at **3394 of 3400**. Every case left standing is
attributed: Excel misreporting a fault as an overflow once a `^` has produced
an infinity, two error-ordering cases, and one gap in the intrinsic-typing
lists — each left for its own round of measurement. The per-seed numbers and the full accounting are in
[`vba-error-ordering.md`](vba-error-ordering.md) §16–§26.

The fuzzer earned its keep several times over, and the findings split into
three kinds worth distinguishing:

- **Two cell divergences no return value could have exposed.** A macro
  assigning an infinity left `-inf` in a cell where Excel leaves `#NUM!`; one
  assigning `"  3  "` left text where Excel leaves the number 3. Both engines
  returned the same value in each case.
- **A Phase 1 rule that had been wrong since it was written.** §11's "a
  runtime String against a static Boolean compares as text" was derived from
  two cases that could not distinguish it from the truth, which is that the
  string converts with `CBool`. Two more rules (§13's static typing, §7's
  `Select Case` conversion) turned out to stop one level too early.
- **Two ordinary bugs** in `InStr` and in negating the `Long` minimum.

The process lesson is in §16's own history: the corrected rule was itself
wrong on the first attempt, in the same way, and only a seed added afterwards
caught it. Each of the eight seeds kept finding something until the last —
and a fourth round later corrected the same corner twice more (§24, §27),
having declared it settled three times.

### Security posture

Executing a macro from an untrusted workbook is a materially different act
from reading one. Even with I/O out of scope, a macro can rewrite the
workbook arbitrarily, and `visi` is a CLI that scripts and CI will point at
files they did not author. Execution must be opt-in per invocation
(`visi macro run FILE --name Foo`), never implicit in `eval`, never
triggered by opening a file, and never wired to `Workbook_Open` without a
separate explicit flag. `--quiet` must not suppress the notice that a macro
ran.

As of Phase 2 that "can rewrite the workbook arbitrarily" is literal rather
than prospective, and the posture holds unchanged and is enforced in code:

- `WorkbookManager::run_macro` is the only entry point that binds a workbook,
  and nothing calls it implicitly — not `load_bytes`, not `evaluate`, not the
  bindings' `roundtrip`.
- `visi macro run` takes `--output` / `--in-place` like every other write
  command, and a macro that **changed the workbook with neither is an error**,
  not a silent discard. Reporting a cheerful return value while throwing the
  writes away is the failure mode this whole feature exists to avoid.
- The "Running VBA procedure ..." notice goes to stderr and is not
  suppressible. Now that a macro can write cells it matters more, not less.
- A macro that raises **after** writing keeps its writes, matching Excel,
  which does not roll back. Reporting the error while quietly reverting would
  be a subtler kind of wrong.

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

### Side finding, since fixed: `fuzz_pivot.py`'s manual setup step is gone

`fuzz_pivot.py` used to require a one-time, human-performed ritual — open
Excel, open the VBA editor, paste `fuzz/BuildFuzzPivot.bas` into a module,
Save As `fuzz/pivot_macro_template.xlsm` — because when it was written there
was no way to get a macro into a workbook programmatically. The same
file-format-level injection this document is built on removes that step:
`ExcelPivotDriver._ensure_macro_template` now generates the template on first
use, and regenerates it whenever the `.bas` is newer, which also closes the
stale-template failure mode the manual flow invited (edit the macro, forget
to rebuild, blame the resulting mismatches on the engine).

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
