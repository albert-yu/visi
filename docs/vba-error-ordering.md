# VBA error ordering and `Null` handling

Scoping for the divergences left after Phase 1 of
[`vba-macro-support.md`](vba-macro-support.md). Findings measured against
Microsoft Excel for Mac 16.112 on macOS 26.6.1, with
[`fuzz/vba_ordering_probe.bas`](../fuzz/vba_ordering_probe.bas) for §2–§6 and
[`fuzz/vba_expr_probe.py`](../fuzz/vba_expr_probe.py) for §7 onwards.

Phase 1 landed at 493 of 500 generated procedures agreeing with Excel on
value, subtype and error number together. This document scoped the remainder,
and **all of it has since been implemented** — each section below records the
measurement, and each fix has a unit test naming the Excel result it came
from.

Sections 2–6 are the first round. Sections 7–13 are the second, which closed
eleven of the twelve cases that round left behind; §14 is the twelfth, which
turned out to be Excel misreporting an error and is deliberately not matched.

Sections 16–22 are a third round, found while building Phase 2 — and worth
noting for *how* they were found rather than what they say. §16 overturns
§11, which had been derived from two cases that could not tell two rules
apart and had stood since Phase 1. §17 is a pair of divergences that a return
value cannot expose at all: the macro returned the right number and left the
wrong thing in a cell. §18 and §21 are both *existing* rules (§13 and §7)
that stopped one level too early. §19 and §20 are outright bugs. §22 is
Excel's, and is not matched.

§23–§28 are a fourth round, on seeds never used before, and five of the six
are the same shape: **a rule that asked whether an operand was a compile-time
*constant* where Excel asks whether the compiler knows its *type***. §24
replaces §16's 4×4 table with one question asked once per side and explains
away the cell §16 admitted it could not; §27 is §23's split applied to the
one branch it had not reached; §28 is the overflow rule, which turns out to
key on the same predicate. §26 is an ordinary bug. §25 is Excel's, is not
matched, and subsumes §22.

Most of them came out of `fuzz_vba.py` once it started comparing the data
grid, which is the argument for that change in one sentence — and out of
simply adding more seeds, which is the argument for
[Overfitting](#overfitting-is-real-here) in another.

**Status:** see [What is left](#what-is-left) for the current per-seed
numbers, and read [Overfitting](#overfitting-is-real-here) first — the gap
between a tuned seed and an unseen one is the most useful thing in this
document.

---

## Overfitting is real here

Iterating against one seed until it reads 500/500 does **not** mean the
engine agrees with Excel. It means the engine agrees with Excel on the 500
programs that seed happens to generate. When this work first hit 500/500 on
seed 1, fresh seeds immediately showed ~2% divergence.

One family in that gap was systematic and worth fixing at once: `CStr` was
rendering 16 significant digits in exponent form and 14 in fixed form, where
Excel shows exactly 15 of each. It appeared on *every* unseen seed and on
none of seed 1. The fix derives everything from a 15-significant-digit
rendering rather than from `log10().floor()` and a decimal count, which is
what was getting it wrong in both directions.

**Always re-run on seeds you have not tuned against** before believing a
number. The remaining ~1–2% is a long tail of error-code disagreements of the
same shape as §2–§4 — which of two faults in one expression surfaces first —
and each needs its own measurement.

## 1. One of the seven was the harness's own fault

`fuzz/fuzz_vba.py` reported `"  3  "` vs `"  3"` on a procedure whose body was
just `vc = "  3  "`. Both engines were right; the driver was calling `.strip()`
on the value it parsed out of the AppleScript reply, eating the padding.

Fixed in this branch (`rstrip("\r")` instead). Worth stating plainly because
it is the failure mode a differential harness is most prone to: **a harness
bug looks exactly like an engine bug, and the engine gets the blame.** Six
real divergences remain.

---

## 2. `0 / 0` is a different error from `x / 0`

| Expression | Excel | visi, before this |
| --- | --- | --- |
| `1 / 0` | error 11 | 11 ✅ |
| `-1 / 0`, `1.5 / 0` | error 11 | 11 ✅ |
| **`0 / 0`** | **error 6** (Overflow) | 11 ❌ |
| **`False / 0`** | **error 6** | 11 ❌ |
| `0 \ 0` | error 11 | 11 ✅ |
| `0 Mod 0` | error 11 | 11 ✅ |

Only floating-point `/` distinguishes them, and only when the numerator is
zero too. `\` and `Mod` raise 11 for `0 \ 0` like any other division by zero,
so this is specific to `/`.

**Done.** `value::div` raises `VbaError::overflow()` when the dividend is
also zero and `VbaError::div_by_zero()` otherwise, covered by
`zero_divided_by_zero_is_overflow_not_division_by_zero`.

Accounts for fuzz cases 398 and 402.

---

## 3. Both operands are coerced before the divisor is tested

| Expression | Excel | visi, before this |
| --- | --- | --- |
| `"xxxx" / 0` | error 13 | 11 ❌ |
| `0 / "xxxx"` | error 13 | 13 ✅ |
| `"" / 0` | error 13 | 11 ❌ |

`value::div` used to read the divisor, test it for zero, and only then
touch the dividend — so a type mismatch on the left is masked by a
division-by-zero on the right. Excel coerces both first, and the type
mismatch wins.

**Done.** `value::div` coerces both operands before testing the divisor,
covered by `division_coerces_both_operands_before_testing_the_divisor`.

Accounts for fuzz case 251.

---

## 4. `^` overflows between constants and yields `INF` at runtime

| Expression | Excel |
| --- | --- |
| `3.75 ^ 32767` (literals) | **error 6** |
| `a = 3.75 : a ^ 32767` | `Double` `INF` |
| `255 ^ 255` (literals) | **error 6** |
| `a = 255 : a ^ 255` | `Double` `INF` |

This is the *same* constant-vs-runtime split already implemented for `+`, `-`
and `*` as [`value::ArithMode`](../visi-core/src/core/vba/value.rs) — `pow`
simply never got wired to it. Phase 1 measured the runtime half (infinity is
a value) and missed that the constant half overflows.

**Done.** `value::pow` takes an `ArithMode` like its siblings and raises
error 6 on a non-finite result in `Constant` mode, covered by
`pow_overflows_between_constants_and_yields_infinity_at_runtime`.

Accounts for fuzz case 187.

---

## 5. Intrinsics split on `Null` — and the obvious probe gets it backwards

This is the one with a methodological trap worth recording.

A first probe asked for `TypeName(v) & "|" & CStr(v)` — the same shape every
other probe in this project uses — and reported that **all nine** intrinsics
tested raised error 94 on `Null`. That conclusion is wrong, and the probe
could not have discovered it: `CStr(Null)` *is itself* error 94, so
stringifying the result cannot distinguish "the function raised 94" from
"the function returned `Null` and `CStr` raised 94".

Re-run with `IsNull(...)` instead, the answer is a split — and it has to be a
measured table, not a rule.

**Done.** The heuristic is replaced by two explicit tables in
`builtins::call`, from a sweep of **all 46** intrinsics rather than the 13
this document originally listed. The full sweep changed the answer in three
places the smaller sample would have got wrong: `CVar` propagates although
every other `C*` conversion rejects, and `Chr` and `Cos` reject although the
old heuristic exempted them.

| Propagate `Null` | Raise error 94 | Inspect it |
| --- | --- | --- |
| `CVar`, `Abs`, `Int`, `Fix`, `Round`, `Len`, `UCase`, `LCase`, `Trim`, `LTrim`, `RTrim`, `Hex`, `Oct`, `Left`, `Right`, `Mid`, `InStr`, `String`, `StrComp` | `CStr`, `CInt`, `CLng`, `CDbl`, `CSng`, `CBool`, `CCur`, `Val`, `Sgn`, `Sqr`, `Exp`, `Log`, `Sin`, `Cos`, `Tan`, `Atn`, `Space`, `StrReverse`, `Chr`, `Asc`, `Replace` | `IsNull`, `IsEmpty`, `IsNumeric`, `IsDate`, `IsObject`, `TypeName`, `VarType`, `IIf` |

There is still no principle visible: `Hex` and `Oct` propagate while `Chr` and
`Asc` reject; `String` propagates while `Space` rejects; `Trim` propagates
while `StrReverse` rejects. `every_intrinsic_handles_null_the_way_excel_does`
enumerates the whole table.

## 6. `"True"` and `"False"` coerce on the integer path only

Measured rather than guessed at, and the answer is a clean split the
hypothesis in the original scoping did not predict:

| Expression | Excel |
| --- | --- |
| `"True" Xor 1` | `Integer` `-2` |
| `"False" Xor 1` | `Integer` `1` |
| `"True" \ 1`, `"True" Mod 2` | `Integer` `-1` |
| `CBool("True")` | `Boolean` `True` |
| `Not "True"` | `Boolean` `False` |
| **`"True" + 1`, `"True" * 2`, `CDbl("True")`** | **error 13** |
| `IsNumeric("True")` | `False` |

So the *integer* conversion path accepts the words as `-1` and `0`,
case-insensitively, while the *floating-point* path has never heard of them.

**Done.** `value::bool_word` and `value::logical_operand` fold such a string
to a `Boolean` on the integer path only — which also explains why `Not "True"`
stays a `Boolean` while `"True" Xor 1` becomes an `Integer`: the first stays
inside `not`'s boolean branch, the second falls into the bitwise one because
only one side is a `Boolean`. Covered by
`the_words_true_and_false_coerce_on_the_integer_path_only`.

## 7. `Select Case` converts its cases to the subject's *static* type

The distinction is invisible in the value: `Select Case CBool(a)` matches
`Case 1`, while `Select Case a` with `a = True` does not, though both subjects
are `True` at run time. When the compiler knows the subject is `Boolean` — a
literal, a folded constant expression, or a call to `CBool`/`IsNumeric`/
`IsNull`/`IsEmpty`/`IsDate`/`IsObject` — **every case value is converted with
`CBool` and the comparison then runs on the Booleans**. A Variant that merely
holds a Boolean gets the ordinary numeric comparison, where `True` is `-1`.

One rule, and it produces a table that looks like six:

| Subject | Case | Excel | Why |
| --- | --- | --- | --- |
| static `True` | `Case 1` | match | `CBool(1)` is True |
| static `True` | `Case 0` | miss | `CBool(0)` is False |
| static `True` | `Case 2 To 5` | **match** | both ends become True |
| static `True` | `Case 0 To 1` | **miss** | the range is False To True, i.e. `0 To -1`, which is empty |
| static `True` | `Case Is < 0` | match | True is -1 |
| static `True` | `Case Null` | **error 94** | `CBool(Null)` raises it |
| Variant `True` | `Case 0, 1` | miss | -1 is neither |
| Variant `True` | `Case 2 To 5` | miss | -1 is outside |

The `Case 0 To 1` and `Case Null` rows are the ones that rule out the simpler
"compare the case value as a Boolean": the conversion happens *first*, and
everything after it is the ordinary comparison.

This also corrects a Phase 1 measurement that this document previously
recorded as fact — "the range form does not do this" — which was taken with a
Variant subject and generalised to both.

**Done.** `interp::is_statically_boolean` plus a conversion in
`case_matches`, covered by
`a_statically_boolean_select_subject_converts_its_cases_with_cbool`.

Accounts for fuzz cases `s2024/52`, `s2024/171`, `s2024/299` and `s31337/119`.

## 8. `/` overflows rather than returning an infinity

`^` is the one operator that hands back an infinity. Every other one raises
error 6 when its result is not finite, *or when an operand already is* — and
`/` was the last one still returning `INF`:

| Expression | Excel |
| --- | --- |
| `1E308 / 1E-308` | error 6 |
| `a = 1E308 : b = 1E-308 : a / b` | error 6 |
| `a = 3.75 : b = a ^ 32767 : b / 2` | error 6 |
| `a = 3.75 : b = a ^ 32767 : b * 2`, `b + 1`, `b \ 2` | error 6 (already matched) |
| `a = 3.75 : b = a ^ 32767 : -b`, `b ^ 1`, `Abs(b)`, `b & "x"` | the infinity, unchanged |

**Done.** `value::div` applies the same finiteness check `value::arith` does,
covered by `division_overflows_rather_than_returning_an_infinity`.

Accounts for `s777/298`.

## 9. `Empty + <String>` concatenates

The lead the previous round left unimplemented, and the probe confirmed it —
with a sharper edge than the hypothesis had: it is not only the non-numeric
strings that concatenate.

| Expression | Excel | visi, before this |
| --- | --- | --- |
| `Empty + "a"`, `"a" + Empty` | `String` `"a"` | error 13 |
| `Empty + "1"` | **`String` `"1"`** | `Double` `1` |
| `Empty + ""` | `String` `""` | error 13 |
| `Empty + CStr(1)` | `String` `"1"` | `Double` `1` |
| `Empty - "a"`, `Empty * "a"`, `Empty / "a"` | error 13 | error 13 ✅ |
| `TypeName(Empty + Empty)` | **`Integer`** | `Empty` |

`Empty` takes the other operand's type, the same way it does against a number
— and only for `+`. The `Empty + Empty` row is a correction to a Phase 1
measurement: reading that result back through the fuzz harness cannot
distinguish `Empty` from the `Integer` 0, since both render as `0` once
assigned onward. Asking `TypeName` *inside* VBA can.

**Done.** In `value::add`, covered by `empty_is_both_zero_and_the_empty_string`.

Accounts for `s2024/4`.

## 10. A trailing sign is part of the number

VBA's string-to-number conversion accepts a sign *after* the digits:

| Expression | Excel |
| --- | --- |
| `CDbl("1-")`, `CInt("1-")`, `CDbl("1 -")` | `-1` |
| `CDbl("1+")` | `1` |
| `CDbl("2.5-")` | `-2.5` |
| `CDbl("1E2-")` | `-100` |
| `IsNumeric("1-")` | `True` |
| `"1-" + 1` | `Double` `0` |
| `CDbl("-1-")`, `CDbl("1--")` | error 13 |
| `Val("1-")` | `1` — `Val` stops at the sign, and always did |

**Done.** In `value::parse_vba_number`, covered by
`string_to_number_accepts_what_vba_accepts`.

Accounts for `s2024/162`, where `If Left("1-2.5", 2) Then` — that is,
`If "1-"` — decides which branch runs, and visi was raising error 13 on the
condition where Excel takes the `Then` branch and reaches a division by zero.

## 11. A runtime String against a static Boolean compares as text

| Expression | Excel | Note |
| --- | --- | --- |
| `a = "011" : a < False` | `True` | "011" sorts before "False" |
| `a = "011" : a > True` | `False` | |
| `a = "011" : a < CBool(0)`, `a < IsNull(32768)` | `True` | any static Boolean |
| `a = "true" : a = True` | `True` | the *words* still convert with `CBool` |
| `a = "011" : b = False : a < b` | `False` | a Boolean **variable** is not static: the runtime rule applies and the number sorts first |
| `a = "011" : a < 0` | `False` | a numeric partner is unaffected |

**Superseded by §16, which is the same measurements under a rule that also
explains the rows this one could not.** Left here because the two are easy to
confuse: every row in the table above holds under both readings, which is
exactly why the wrong one survived a whole phase.

Accounts for `s2024/145`.

## 12. The `"True"`/`"False"` fold stops when both sides are statically typed

Section 6 established that the words coerce on the integer path. The
open question was when a `Boolean` partner suppresses that, and the answer is
the same static-typing distinction as §7 and §11:

| Expression | Excel | Why |
| --- | --- | --- |
| `True Eqv "True"` | error 13 | Boolean literal, String literal |
| `True Eqv CStr(True)` | error 13 | `CStr` is declared `As String` |
| `a = 3.75 : IsNumeric(a) Eqv CStr(True)` | error 13 | both declared |
| `LCase("TRUE") Eqv True` | `True` | **`LCase` returns a `Variant`** |
| `LCase(False) Eqv IsNull(True)` | `True` | |
| `a = True : a Eqv "True"` | `True` | `a` is a Variant |
| `a = "false" : a Eqv False` | `True` | same, other way round |
| `a = "true" : b = False : a Eqv b` | `False` | |

The `CStr`/`LCase` pair is what pins it down, and the split is not arbitrary:
`CStr` returns `String`, while `LCase`, `UCase`, `Left` and the rest return
`Variant` — it is their `$`-suffixed forms that are typed. visi had the
suppression keyed on "the partner is a Boolean" alone, which matched the four
all-constant rows and made every runtime row error.

**Done.** `value::logical_pair` takes the operand kinds, covered by
`the_words_true_and_false_coerce_on_the_integer_path_only`.

Accounts for `s2024/103`.

## 13. Strictness needs a constant string — and `Len`, `Val` and `Sgn` are static

Two corrections to the table below, both from cases that hold one side fixed
while the other varies.

**`Len`, `Val` and `Sgn` are statically typed**, like the `C*` conversions.
Against the constant string `(-32768 & -2.5)` all four raise error 13, while
`Int(a)`, `Abs(a)` and a bare `a` do not. The earlier round put `Len` in on
the strength of its `As Long` signature, tested it against a *runtime* string,
saw no error and took it out — but nothing is strict against a runtime string,
so the case could not discriminate. `Int` and `Abs` stay out for a reason
visible in their signatures: they return the type they were handed, so a
Variant argument makes them Variant.

**Strictness only applies to a constant string.** `CLng(a) < ("abc" & a)` is
`True` in Excel, not error 13: a runtime string that fails to parse falls back
to the ordering rule, exactly as it does against a numeric constant. It still
compares numerically when it *does* parse — `a = 5 : b = "1" : CLng(a) < b` is
`False`.

**Done.** `interp::STATICALLY_NUMERIC` and `value::compare_ctx`, covered by
`a_statically_typed_numeric_partner_is_strict_only_against_a_constant_string`.

Accounts for `s31337/123` and `s31337/147`.

## String against number, in full

This was the last hard question, and it took four attempts. The rule:

| Numeric side | String side | Behaviour |
| --- | --- | --- |
| statically typed (`CLng` and the other numeric `C*` conversions) | anything | the **whole** string must parse; error 13 otherwise |
| constant | constant | the string's numeric **prefix** is used, as `Val` takes it; error 13 if there is no prefix at all |
| constant | runtime | same prefix rule, but no prefix falls back to the ordering below rather than erroring |
| runtime | constant | plain **string** comparison, the number rendered with `CStr` |
| runtime | runtime | a number sorts **before** any string, whatever the values |

Prefix coercion is what finally separated the pair that had defeated three
earlier models:

| Expression | Excel | Why |
| --- | --- | --- |
| `(Not 2!) <= ("1.5" & False)` | `True` | `"1.5False"` has the numeric prefix `1.5` |
| `(-True) <> (True & &HFF)` | error 13 | `"True255"` has no numeric prefix |

Structurally those two are identical -- same operator family, same side for
the string, same literal-versus-constant-expression shape on each side, both
strings built by `&` from two literals. Only the *content* of the string
separates them, which is not something a structural model could ever have
reached.

One supporting rule: **`Null` is not foldable**, so nothing containing it is
constant. `(False & Null) = (0.1 / -2.5)` is `False` rather than an error
because the string side is not constant and therefore falls back.

### Failed models, kept so they are not retried

1. *`&` is never constant-folded.* Net negative on its own: 2093 against 2096
   across seven seeds, fixing one case and breaking three.
2. ~~*Any statically-typed numeric partner is strict, including `Len`.*~~
   **Rehabilitated — see §13.** `Len`, `Val` and `Sgn` really are strict; the
   case that rejected the model (`("abc" & va) <> Len(CStr("Z"))`) held the
   *string* side runtime, where nothing is strict, so it could not
   discriminate. A model rejected on a confounded case is not rejected.
3. *`^` with a String operand is strict.* `"255" ^ 255` and
   `StrReverse(255) ^ 1.5E54` are both `INF`, exactly as their numeric
   equivalents. The case that looked like evidence had a base of `"1E+2923"`,
   where the real rule was that the *conversion* overflows.
4. *Strictness depends on literal versus constant expression.* This scored
   best of the structural models and is still wrong; it was replaced by
   prefix coercion, which is about the string's content instead.
5. *A String compares against a Boolean as text.* True only when the string is
   a **runtime** value (§11). Applied to constants as well it fixed one case
   and broke six across four seeds — 1192 to 1190 — and was narrowed rather
   than kept. The tell was that every regression had a constant string.
6. *The `"True"`/`"False"` fold is suppressed whenever the partner is a
   Boolean.* Right for the four all-constant cases it was measured on, wrong
   for every runtime one (§12).

## 14. `s555/219` is Excel misreporting, and is not matched

The twelfth case turned out not to be a rule at all. Once a procedure has
produced an infinity — `^` is the only operator that can — Excel raises error
6 on the **next string-to-number conversion anywhere in that procedure**,
whatever string it is given:

```vba
a = 3.75
b = a ^ 32767        ' Double INF, no error
c = CDbl("1.5")      ' Excel: error 6.  visi: 1.5
```

It survives unrelated intervening statements but clears once the error has
been reported, and `Val` — a different parser — is immune. That is a sticky
floating-point status flag being read by the conversion routine, not
behaviour worth reproducing. Full table in
[`excel-discrepancies.md` §15](excel-discrepancies.md).

## 15. A static String over a `Null` is error 94

Found on a seed never used during this work, which is the point of running
those:

| Expression | Excel |
| --- | --- |
| `"  3  " Imp Null`, `"3" And Null`, `"1.5" Or Null`, `"0" Or Null` | error 94 |
| `("  " & "3") Or Null`, `CStr(3) Or Null` | error 94 |
| `a = Null : "  3  " Or a` | error 94 — the *Null* may be a variable |
| `a = "  3  " : a Imp Null` | no error — the **String** may not |
| `Null Or "  3  "`, `Null And "  3  "`, `Null Xor "  3  "` | no error — left-specific |
| `"abc" Imp Null`, `"True" Or Null` | error 13 — the conversion is checked first |
| `3 Imp Null`, `255 Imp Null` | no error — the operand must be a String |

**Done.** `interp::null_on_the_right`, covered by
`a_static_string_over_a_null_is_invalid_use_of_null`.

## 16. It was never a text comparison — the string converts with `CBool`

§11 read "a runtime String against a static Boolean compares as text", derived
from two cases that do not discriminate: `a = "011" : a < False` is `True`
whether "011" sorts before "False" or `CBool("011")` is `True` (-1) and sorts
below `False` (0). Every case that *does* discriminate says conversion:

| Expression | Excel | Text says | `CBool` says |
| --- | --- | --- | --- |
| `a = "-1" : a = True` | `True` | `False` | `True` |
| `a = "011" : a = True` | `True` | `False` | `True` |
| `a = "0" : a = False` | `True` | `False` | `True` |
| `a = "1.5" : a = True` | `True` | `False` | `True` |
| `a = "abc" : a = True` | `False` | `False` | error 13 |
| `a = "" : a = False` | `False` | `False` | error 13 |

So: **convert the string with `CBool`, compare the two Booleans as numbers**
(`True` is -1, so it sorts *below* `False`), and fall back to a text
comparison only when the conversion fails. The last two rows are what force
the fallback rather than an error.

Two axes then decide the details, and the combination is not guessable — this
took three rounds of measurement, the middle one of which was wrong in a way
the fuzzer caught.

**The table below is superseded by §24**, which is these same two axes
measured with a string that discriminates them (`"13"`, where converting and
comparing numerically give opposite answers). Five of its cells are wrong,
and the "one cell that resists explanation" was an artifact of reading
`(3# >= Empty)` as a folded Boolean when `Empty` makes it a `Variant`. Left
here because it is the third of four readings of this corner, and every one
was confirmed by every case then available. The `CBool` finding above stands.

**Which comparison runs depends on how well the compiler knows each side.**
The full table, where `cv` is the `CBool` conversion, `txt` a text comparison
and `num` the numeric rules of the section below:

| | Bool literal | Bool static | Bool folded | Bool variable |
| --- | --- | --- | --- | --- |
| **String runtime** | cv | cv | cv | numeric rules |
| **String literal** | cv | cv | cv | numeric rules |
| **String static** | cv | cv | **txt** | numeric rules |
| **String folded** | num | num | num | numeric rules |

Every cell is measured. The one that resists explanation is *static string
against folded Boolean*, and three rows put it there:

| Expression | Excel | The conversion would say |
| --- | --- | --- |
| `CStr(0) >= (3# >= Empty)` | `False` | `True` |
| `TypeName(0) >= (3# >= Empty)` | `False` | error 13 |
| `(3# >= Empty) >= TypeName(0)` | `True` | — (not side-specific) |

while a *literal* or *runtime* string against the same folded partner does
convert — `("000" < ("1" >= -7))` and `a = "000" : a < ("1" >= -7)` are both
`False`, which only `CBool` gives. It is written in `value.rs` as the
exception it is rather than dressed up as a rule.

`(Not True)` counts as a **literal** here and `(3# >= Empty)` does not, which
is the subtlest part: VBA folds a unary chain over a literal but not a
comparison. `operand_kind` had to become recursive to tell them apart — it
previously looked one level down and put `(Not True)` in the wrong bucket.

A Boolean held in a **variable** is not statically known at all, so none of
the table applies and the ordinary runtime rule takes over:
`a = "011" : b = False : a < b` is `False`, where every static partner makes
it `True`.

And on a failed conversion, what happens depends on how well the compiler
knows the string's type — the same static/runtime axis as §7, §11 and §13:

| String operand | Failed conversion | Example |
| --- | --- | --- |
| runtime `Variant` | ~~compares as text~~ **orders above the number** — see §23 | `a = TypeName(32767) : a >= (Not True)` is `True` |
| declared `String` (`CStr`, `TypeName`) | **error 13** | `TypeName(32767) >= False` |
| literal | **error 13** | `("abc" < True)` |
| folded constant expression | numeric path, unchanged | `((Empty & "1") <= ("" <> Empty))` is `False` |

Two consequences worth calling out.

**`TypeName` had to be added to `STATICALLY_NUMERIC`'s string counterpart.**
`TypeName(32767) >= False` is error 13 while `LCase("Integer") >= (Not True)`
is `True`, and the only difference is that `TypeName` is declared `As String`
where `LCase` returns `Variant`. Measured, not read off the signature.

**This closed `("011" < False)`**, the row §11 and *What is left* both carried
as unexplained. It is not a text comparison after all: `CBool("011")` is
`True`, i.e. -1, which really is less than `False`. The earlier attempt failed
because it tried to make literals compare as *text*, which is a different rule
and regressed six cases.

**Found by the Phase 2 fuzzer**, on a case whose visible symptom was a *cell*
holding the wrong value — `fuzz_vba.py` now compares the data grid, and the
return value alone agreed.

**Done**, and verified as a whole: 23 cases spanning every cell of the table
agree with Excel. `value::compare_ctx` and `interp::operand_kind`, covered by
`a_string_converts_with_cbool_against_a_static_boolean`.

It took three rounds to get here, and the middle one was wrong in a way worth
recording: "the Boolean side decides, a literal converts and everything else
compares as text" fit every case measured at the time and was refuted by the
next fuzz run, which turned up a *runtime* string against a folded Boolean
converting. Two cases that cannot tell two rules apart look exactly like
confirmation of whichever rule you already believe.

## 17. Excel has no infinities in *cells*, and trims numeric text on entry

Two divergences the Phase 2 cell comparison found, both invisible in a return
value and both about what ends up in the saved file:

| A macro assigns | Excel stores | visi stored |
| --- | --- | --- |
| `(-2.5 ^ va)` with `va = 1000`, i.e. `-INF` | `#NUM!` | the `Double` `-inf` |
| the string `"  3  "` | the `Double` `3` | the `String` `"  3  "` |
| the string `"#NUM!"` | the error `#NUM!` | the `String` `"#NUM!"` |

`^` is the one VBA operator that yields an infinity rather than raising error
6 (§4), so a macro genuinely reaches the first row. The other two are ordinary
cell-entry parsing: entering `  3  ` gives the number and entering `#NUM!`
gives the error, in Excel and through `Range.Value` alike.

All three are fixed in `Sheet::commit`'s literal parsing rather than in the
VBA layer, since they are properties of *entering a value into a cell*, not of
macros — with `xlsx::text_cell_src` quoting the two new ambiguous shapes so an
imported text cell that spells one still round-trips as text.

**Done.** Covered by `writing_an_infinity_stores_the_num_error_excel_stores`
and `writing_whitespace_padded_numeric_text_stores_a_number`.

## 18. Static typing propagates through arithmetic

§13 established that a constant string compared against a *statically typed*
number has to parse whole, and that `Len`, `Val` and `Sgn` count as
statically typed alongside the `C*` conversions. What it missed is that the
property survives arithmetic: `Len(CStr(a)) / 2` is a `Double` as surely as
`Len(CStr(a))` is a `Long`, because every operand's type is known. One
`Variant` operand loses it for the whole expression.

| Expression, with `a = -3` | Excel | Why |
| --- | --- | --- |
| `Len(CStr(a)) = "-7False"` | error 13 | bare call — §13, already implemented |
| `(Len(CStr(a)) / 2) = "-7False"` | error 13 | propagated through `/` |
| `(Len(CStr(a)) + 1) = "-7False"` | error 13 | propagated through `+` |
| `(CLng(a) / 2) = "abc"` | error 13 | likewise |
| `(Len(CStr(a)) + a) = "-7False"` | `False` | a `Variant` operand — text |
| `(a / (-32768)) = "-7False"` | `False` | no static operand at all |
| `(CLng(a) * 2) = "-6.0"` | `True` | the positive half: **numeric**, not text |

That last row matters as much as the errors. Against a statically typed
number a string that *does* parse compares numerically, where a `Variant`
partner would compare it as text and answer `False`.

Only arithmetic propagates here. Comparison and `&` are left out because
nothing measured covers them, not because they are known not to.

**Done.** `interp::is_statically_typed`, covered by
`static_typing_propagates_through_arithmetic`. Found by the fuzzer on an
unseen seed; the *rule* was already implemented and tested, and all that was
missing was that it stopped at the top-level call.

## 19. Negating the `Long` minimum wraps to itself, between constants

`-(Not 2147483647)` is arithmetically 2147483648. Excel gives back the `Long`
**-2147483648** — plain two's complement, and wrong.

| Expression | Excel |
| --- | --- |
| `-(Not 2147483647)` | the `Long` `-2147483648` |
| `-(Not 32767)` | error 6 — the `Integer` minimum does **not** do it |
| `a = 2147483647 : -(Not a)` | the `Double` `2147483648` — at run time it widens |

Narrow enough to be matched deliberately rather than treated as Excel being
wrong (§14): it is deterministic, and a macro doing this should behave the
same way here.

**Done.** `value::neg`, covered by
`negating_the_long_minimum_between_constants_wraps_to_itself`.

## 20. `InStr` with an empty haystack is 0

`InStr("", "")` is `0` while `InStr("a", "")` is `1`. An empty needle matches
at the start position only when there is a string to match in; this used to
report `1` for the empty/empty pair, on the reasoning that an empty needle
always matches. `Empty` reaches the same path as `""` and behaves the same.

**Done.** `builtins::instr`, covered by `instr_of_an_empty_haystack_is_zero`.

## 21. `Select Case` sees `Not` of a Boolean as statically Boolean

§7 established that `Select Case` converts its case values to the subject's
*static* type. That property carries through `Not`, because `Not` of a
Boolean is a Boolean:

| Subject | Excel takes | Why |
| --- | --- | --- |
| `(Not IsEmpty("Z"))` | `Case 0, 1` | statically Boolean → the cases convert with `CBool` |
| `(Not IsEmpty(""))` | `Case 0, 1` | same, with the subject `False` |
| `(Not (IsEmpty("Z")))` | `Case 0, 1` | parentheses do not change it |
| `(Not CBool(0))` | `Case 0, 1` | any statically Boolean operand |
| `(Not 5)` | `Case Else` | `Not` of a *number* is a number (-6), so no case matches |

This one is worth noting for its blast radius rather than its subtlety: the
missing case sent the fuzzer's procedure down a `Case Else` arm that raised
on an expression the correct arm never evaluates, so the symptom was an error
number, not a wrong branch.

**Done.** `interp::is_statically_boolean`, covered by
`select_case_sees_not_of_a_boolean_as_statically_boolean`.

## 22. An infinity poisons a later fractional `If` literal — Excel only

This is the second §14: Excel misreporting, deliberately not matched.

**§25 is the same artifact, seen more broadly.** What is described here as an
`If` condition being poisoned turns out to be one shape of a general rule:
once a `^` has produced an infinity anywhere in a procedure, Excel misreports
later, unrelated faults as overflows. Read the two together.

A procedure that produces an infinity with `^` and *later* tests a
non-integral numeric **literal** as an `If` condition raises error 6 in
Excel. Neither half alone does it:

| Procedure | Excel |
| --- | --- |
| `va = 2147483647 : va = (3# ^ va) : If 0.0001 Then ...` | error 6 |
| `va = 2147483647 : If 0.0001 Then ...` | fine — no infinity |
| `va = 2147483647 : va = (3# ^ va) : If 2 Then ...` | fine — an *integral* literal |
| `va = 2147483647 : va = (3# ^ va) : vb = 0.0001 : If vb Then ...` | fine — not a literal |
| `va = 2147483647 : vb = (3# ^ va) : If 0.0001 Then ...` | error 6 — a *different* variable |
| `va = 2147483647 : va = (3# ^ va) : va = 1 : If 0.0001 Then ...` | error 6 — even overwritten |
| `va = 2147483647 : va = (3# ^ va) : vc = 0.0001` | fine — an assignment, not an `If` |

The last two are what settle it. The error survives the infinity being
overwritten, and it depends on the *statement kind* rather than on any value,
so it cannot be a rule about what anything holds — it is an artifact of
Excel's compiler or its optimiser. Matching it would mean tracking whether
any `^` in a procedure ever produced an infinity and then poisoning unrelated
`If` statements, which is not a behaviour worth reproducing.

Accounts for `s77/112`, and is the one case left on eight seeds.

## 23. The failed-conversion fallback orders, and *declared* `String` is what makes a comparison strict

Two corrections to §16, from one `fuzz_vba.py` mismatch. The visible symptom
was a procedure raising **error 11** in visi and **13** in Excel with an
identical data grid: Excel stopped at a comparison in the first loop, while
visi took that comparison as `True`, entered the other branch of a later `If`,
and divided by zero somewhere Excel never reached.

Reduced: `StrReverse(False) > (Not False)` — error 13 in Excel, `True` in visi.

**`StrReverse`, `Replace` and `Join` are declared `As String`.** They are the
members of the string family with no `$` form, so the plain name *is* the
typed one — where `LCase`, `UCase`, `Left` and `Trim` return `Variant` and it
is `LCase$`/`Left$` that are typed. Measured, not read off a signature:

| Expression | Excel |
| --- | --- |
| `StrReverse("abc") > True` | error 13 |
| `Replace("abc", "a", "z") > True` | error 13 |
| `Join(Array("a", "b")) > True` | error 13 |
| `Trim("abc") > True` | `True` |
| `LTrim("abc") > True` | `True` |

**The fallback was never a text comparison either.** §16 corrected §11 from
"text" to `CBool`, and left text standing as what happens when the conversion
fails. It does not: the string simply **orders above the number**, which is
the same runtime rule as everywhere else.

Every case §16 had available agrees with both readings, because `"abc"`,
`"Integer"` and `""` all sort on the same side of `"True"`/`"False"` as the
ordering rule puts them. A string that does *not* separates them:

| Expression | Excel | Text says | Ordering says |
| --- | --- | --- | --- |
| `a = "ABC" : a > True` | `True` | `False` | `True` |
| `a = "ABC" : a < True` | `False` | `True` | `False` |
| `a = "ABC" : a >= False` | `True` | `False` | `True` |
| `Chr(65) > True` | `True` | `False` | `True` |
| `Chr(65) > False` | `True` | `False` | `True` |
| `Hex(255) > True` | `True` | `False` | `True` |
| `Space(2) > True` | `True` | `False` | `True` |
| `StrConv("abc", 1) > True` | `True` | `False` | `True` |

**And the same "declared, not constant" split was missing on the numeric
path**, where it had been wrong for `CStr` and `TypeName` since before
`StrReverse` joined them — strictness keyed off the string being *constant*,
so a declared-`String` call against a number ordered instead of raising:

| Expression | Excel | visi was |
| --- | --- | --- |
| `CStr("abc") > 5` | error 13 | `True` |
| `TypeName(1) > 5` | error 13 | `True` |
| `CStr("abc") > CLng(1)` | error 13 | `True` |
| `StrReverse("abc") > 5` | error 13 | `True` |
| `Trim("abc") > 5` | `True` | `True` |
| `CStr("11") > 5` | `True` — it converts | `True` |

So both branches now turn on the same question — *does the compiler know this
is a `String`?* — rather than one asking that and the other asking whether the
value is constant.

This is the third time this corner has been re-measured, and the third time
the previous reading was one that every then-available case confirmed. §16's
own closing paragraph predicted it: *two cases that cannot tell two rules
apart look exactly like confirmation of whichever rule you already believe.*
The discriminating case has to be constructed deliberately — here, a string
whose first letter sorts below `T` and `F`.

`value::compare_ctx` and `interp::STATICALLY_STRING`, covered by
`an_unconvertible_runtime_string_sorts_above_a_static_boolean` and
`statically_string_intrinsics_are_strict_against_a_boolean`.

`Join` is listed in `STATICALLY_STRING` though it is unimplemented (the call
raises 35 first), for the reason `IsArray` is listed in `STATICALLY_BOOLEAN`.

## 24. §16's table is one question asked twice, and its unexplained cell was `Empty`

§16 measured a 4×4 table of string-kind against Boolean-kind and closed by
admitting one cell it could not explain: a *static* string against a *folded*
Boolean compared as text where every neighbouring cell converted. That cell
was an artifact of how the Boolean side was classified. There is no table.

**A comparison is statically `Boolean` only when both its operands are
statically typed**, because a `Variant` operand could make the result `Null`.
`(3# >= Empty)` is a compile-time *constant* — `Empty` is a literal — and is
still not a compile-time `Boolean`, because `Empty` is a `Variant`. §16 read
it as folded-and-therefore-known, and had to write the consequence down as an
exception.

Once the Boolean side is classified by static *type* rather than by
constness, the whole thing is two independent questions, one per side:

| | statically `Boolean` partner | not statically `Boolean` |
| --- | --- | --- |
| **statically `String`** | convert with `CBool`; error 13 if it will not | text, and never an error |
| **not statically `String`** | convert with `CBool`; §23 ordering if it will not | numeric rules |

The Boolean side, holding the string fixed at the literal `"0"` so only the
partner varies. Convert says `True` (`CBool("0")` is 0, and `0 >= -1`); text
says `False` (`"0"` sorts below `"True"`):

| Expression | Excel | Why |
| --- | --- | --- |
| `"0" >= (3# >= CDbl(0))` | `True` | every operand statically typed |
| `"0" >= (Len(CStr(0)) >= 1)` | `True` | likewise |
| `"0" >= ("1" >= -7)` | `True` | a string *literal* is statically typed |
| `"0" >= (2 >= 1)` | `True` | — |
| `"0" >= (3# >= Empty)` | `False` | `Empty` is a `Variant` |
| `"0" >= (Empty = Empty)` | `False` | likewise |
| `b = 1 : "0" >= (3# >= b)` | `False` | a variable, likewise |
| `"0" >= IsEmpty(Empty)` | `True` | **declared** `Boolean`, `Empty` and all |
| `"0" >= CBool(Empty)` | `True` | likewise |

The last two are what make this about the static type rather than about
`Empty` appearing somewhere in the expression.

**And the same propagation was missing on the String side**, which is the half
§18 explicitly left open — *"comparison and `&` are left out because nothing
measured covers them, not because they are known not to."* They propagate.
Against a Boolean that is not statically known, where a statically typed
`String` compares as text and a `Variant` takes the numeric rules:

| Expression | Excel | visi was |
| --- | --- | --- |
| `"13" <= ("" <> Empty)` | `True` | `True` |
| `("1" + "3") <= ("" <> Empty)` | `True` | `False` |
| `("1" & "3") <= ("" <> Empty)` | `True` | `False` |
| `CStr(13) <= ("" <> Empty)` | `True` | `True` |
| `a = "13" : a <= ("" <> Empty)` | `False` | `True` |
| `(Empty & "13") <= ("" <> Empty)` | `False` | `False` |

A fold of two string literals is a `String` as surely as a literal is; a fold
over `Empty` is not, and behaves exactly as a variable does. The strictness
follows with it — `("1" + "  3  ") <= False` is error 13 where
`(Empty & "1  3  ") <= False` orders (§23).

`fuzz_vba.py` found this on seed 271828 as `("1" + "  3  ") > (2! >= -1)`,
which visi took as `True` and Excel raised 13 on. Reducing it turned up the
`Empty` distinction, and re-measuring §16's grid with a string that
*discriminates* — `"13"`, where converting and comparing numerically give
opposite answers — corrected five of its cells and dissolved its exception.

**This is the fourth pass over this corner, and the third time the previous
reading was one that every then-available case confirmed.** §16's own cases
put `"000"` and `"0"` against folded Booleans, and both sort below `"True"`
exactly where the conversion puts them; nothing in that set could tell the
two apart. The discriminating case has to be built on purpose.

**Done.** `interp::is_statically_typed` (now propagating through `&` and the
six comparison operators, with `Empty`/`Null` excluded as `Variant`),
`interp::operand_kind` and `value::compare_ctx`, covered by
`static_typing_propagates_through_comparison_and_concatenation`.

## 25. An infinity poisons later error reporting in the same procedure — Excel only

The third §14, and documented rather than matched. **§22 is the same
artifact**, seen through one statement kind: once `^` has produced an infinity
somewhere in a procedure, Excel misreports later, unrelated faults as
overflows.

Two shapes beyond §22's, both measured. The plainest is that a later **type
mismatch** is reported as error 6:

| Procedure | Excel | visi |
| --- | --- | --- |
| `True * "abc"` | error 13 | error 13 |
| `vb = (2 ^ 3) : True * "abc"` | error 13 | error 13 |
| `vb = ((1 & 1) ^ Abs(32768)) : True * "abc"` | **error 6** | error 13 |
| `vb = ((1 & 1) ^ Abs(32768)) : 1 + "abc"` | **error 6** | error 13 |

The fault is a type mismatch in every row, the string never parses in any of
them, and the only thing that changes is whether an *earlier* `^` overflowed.
A finite `^` in the same position leaves the error alone. This is what makes
it Excel's bookkeeping rather than a rule about values — and it is why the
case it came from (seed 31459 case 114) looked at first like the numeric-prefix
scan below: the string there was `"-1.00003051850948-32768"`, which invites
that reading and has nothing to do with it.

The sharpest demonstration needs no fault at all — only a **later statement**.
These two differ by one trailing assignment that touches nothing:

```text
vb = 32767 : va = 5623.41325190349 : va = (CInt(vb) ^ (vb Mod va)) : CStr(va)
    INF

vb = 32767 : va = 5623.41325190349 : va = (CInt(vb) ^ (vb Mod va)) :
    vb = ("  3  " & 3.75) : CStr(va)
    error 6
```

Same infinity, same value returned, and appending an unrelated concatenation
turns it into an overflow. Found on seed 314159, and it had to be told apart
from §28 — the same case reduced to *two* findings, one Excel's and one
visi's, which is why reducing to the smallest expression matters more than
reproducing the original.

The second shape is a **comparison**. Once `^` has produced an infinity (§4 —
it is the one operator that does), comparing it raises error 6 when the
*other* operand is a computed expression of floating type, and does not when
the same value arrives as a literal, a `Long`, or a variable:

| Right-hand operand, with `vc = 255` and the left `(1E3 ^ vc)` | Excel |
| --- | --- |
| `1000`, `1E3`, `1000#` | `True` — a literal is fine |
| `CLng(1000)`, `CInt(1000)`, `Len("abcd")` | `True` |
| `(1000 + 1)`, `(1000 And -1)`, `(CLng(1000) + 0)` | `True` |
| `vd`, where `vd = 1000` | `True` — a variable is fine |
| `CDbl(1000)`, `CDbl(0)`, `CSng(1000)` | **error 6** |
| `(1E3 + 0)`, `(1E3 * 1)`, `(1000 / 1)` | **error 6** |
| `(1E3 And -1)`, `(1E3 Eqv -1)` | **error 6** |
| `CLng(1E3)` | **error 6** — a `Long`, but folded from a `Double` |

Three things put this outside any rule about values. `CLng(1000)` is fine and
`CLng(1E3)` is not, though both are the `Long` 1000. Swapping the operands —
`CDbl(1000) < (1E3 ^ vc)` — is fine, though comparison is symmetric here. And
a *finite* left-hand side is fine against every one of these.

So it depends on the syntactic shape of the other operand and on which side
the infinity is, not on what either holds. `inf > 1000` is `True`, which is
what visi answers and what the same comparison answers in Excel when the 1000
is spelled differently.

Matching any of this would mean tracking whether any `^` in a procedure ever
produced an infinity and then poisoning unrelated statements — §22's
conclusion, now with two more shapes arguing for it.

Accounts for seed 999983 case 163 and seed 31459 case 114.

## 26. `IsNumeric(Empty)` is True

`IsNumeric(Empty)` is **True** and `IsNumeric(Null)` is False. `Empty`
answers as the 0 it coerces to; `Null` answers for nothing. `IsNumeric("")`
is False, though `""` and `Empty` compare equal — the same asymmetry §16
records for `"" = 0` (error 13) against `Empty = 0` (True).

`fuzz_vba.py` found it on seed 862021 through `(Not vc) Xor IsNumeric(Empty)`,
which is `1` with the operand False and `-2` with it True, so the wrong
answer surfaced as an ordinary wrong number rather than as an error.

**Done.** `builtins::is_numeric`, covered by
`is_numeric_of_empty_is_true_and_of_null_is_false`.

## 27. §23's split was missing on the runtime-number branch too

§23 established that a **declared** `String` is what makes a comparison
strict, and fixed the branch where the number is statically typed:
`CStr("abc") > 5` is error 13 where `Trim("abc") > 5` is True. The branch
where the number is a *runtime* `Variant` kept asking whether the string was
constant, so a declared `String` fell through to the ordering rule instead of
comparing as text.

With `a` a variable throughout, so the number is never static and §13's
strictness never applies — these differ only in how well the compiler knows
the **string**:

| Expression | Excel | visi was |
| --- | --- | --- |
| `a = 5 : a < "10"` | `False` | `False` — text, `"5"` sorts above `"1"` |
| `a = 5 : a < CStr(10)` | `False` | `True` |
| `a = 5 : a < (CStr(1) & "0")` | `False` | `True` |
| `b = 10 : a = 5 : a < CStr(b)` | `False` | `True` |
| `a = 5 : a < Trim("10")` | `True` | `True` — a `Variant` orders |
| `a = -2 : a < CStr("")` | `False` | `True` |
| `a = -2 : a < StrReverse("")` | `False` | `True` |
| `a = -2 : a < CStr("abc")` | `True` | `True` |

Note the fourth row: `CStr(b)` is a declared `String` whose *value* is only
known at run time, and it still compares as text. As everywhere else in this
document, the question is what the compiler knows about the **type**.

The two unconvertible rows are also what say this branch never raises — text,
not error 13, and not the ordering either.

`fuzz_vba.py` found it on seed 987654, in a case that only reached this branch
because §24 had just made `(x & y)` statically `String`: the two changes are
independent but the second exposed the first.

**Done.** `value::compare_ctx`, covered by
`a_statically_string_value_compares_as_text_against_a_runtime_number`.

## 28. Overflow "between constants" is really between *statically typed* operands

The rule behind [`ArithMode`](../visi-core/src/core/vba/value.rs) — `32767 + 1`
written with two literals is error 6, while `a = 32767 : a + 1` is the `Long`
32768 — was keyed on whether both operands were compile-time **constants**.
It is static typing, and the two come apart in both directions:

| Expression | Excel | visi was |
| --- | --- | --- |
| `CInt(32767) + 1` | error 6 | `32768` |
| `CInt(32767) * 2` | error 6 | `65534` |
| `CInt(32767) + CInt(1)` | error 6 | `32768` |
| `Sgn(1) + 32767` | error 6 | `32768` |
| `CLng(2147483647) + 1` | error 6 | `2147483648` |
| `CInt(32767) ^ 4652` | error 6 | `INF` |
| `CDbl(32767) ^ 4652` | error 6 | `INF` |
| `(Empty + 32767) + 1` | **`32768`** | **error 6** |

The last row is the other direction and the one that pins the rule down: it
*is* a constant expression, and it promotes anyway, because `Empty` is a
`Variant`. Constness cannot explain both halves; static typing explains both,
and it is the same predicate §24 and §27 turn on.

Two rows that look like exceptions and are not. `Len("abcde") + 32763` is
`32768` because `Len` is declared `Long` — the expression is statically typed,
there is simply nothing to overflow at that width. And `^` still yields an
infinity rather than raising (§4) whenever the expression is *not* statically
typed, which is why `vb = 4652 : 32767 ^ vb` is `INF` while
`CInt(32767) ^ 4652` is error 6.

`fuzz_vba.py` found it on seed 314159, in a case that also demonstrated §25 —
see there for why the two had to be separated before either could be read.

**Done.** `interp`'s `ArithMode` selection, covered by
`overflow_between_constants_is_really_between_statically_typed_operands`.

## What is left

| Seed | Before | After |
| --- | --- | --- |
| 2024 | 293 | **300** |
| 31337 | 297 | **300** |
| 555 | 299 | 299 |
| 777 | 299 | **300** |
| 293 (never used) | — | **300** |
| 297 (never used) | — | **300** |

`s555/219` is the one remaining, and it is §14 — Excel reporting an overflow
on a conversion that cannot overflow. It is left in place rather than
suppressed, so that the harness keeps saying 299 and the reason stays visible.

Phase 2's harness generates different programs — roughly a third of the
statements now touch the workbook — so its numbers are not comparable with the
table above and are kept separate. On seeds never used while developing, after
§16–§21:

| Seed | Cases | Agreed |
| --- | --- | --- |
| 909 | 200 | **200** |
| 4242 | 200 | **200** |
| 55555 | 200 | **200** |
| 13579 | 200 | **200** |
| 31337 | 200 | **200** |
| 2024 | 200 | **200** |
| 8675309 | 200 | **200** |
| 77 | 200 | 199 |

**1599 of 1600**, and the one remaining is §22 — Excel raising an overflow
that depends on the *statement kind* rather than on any value. It is left in
place rather than suppressed, so the harness keeps saying 199 and the reason
stays visible, exactly as `s555/219` is for §14.

A **fourth round** then ran fifteen more seeds never used before, which is
where §23–§28 came from. After them:

| Seed | Cases | Agreed | The one left |
| --- | --- | --- | --- |
| 60622, 8161, 424243, 5150 | 200 each | **200** each | |
| 77777, 20260816, 606060 | 200 each | **200** each | |
| 20250101, 8675309, 424242 | 200 each | **200** each | |
| 271828 | 200 | **200** | was §24 |
| 862021 | 200 | **200** | was §26 |
| 987654 | 200 | **200** | was §27 |
| 999983 | 200 | 199 | §25 |
| 314159 | 200 | 199 | §25 (§28 was the other half of the same case) |
| 31459 | 200 | 198 | §25, and the ordering case below |
| 19990101 | 200 | 198 | the two cases below |

**3394 of 3400**, and all six remaining are attributed: three are §25, two are
the error-ordering family, and one is the intrinsic-typing gap — the last
three all listed under *Known and unfixed*.

Worth reading alongside [Overfitting](#overfitting-is-real-here): five of the
six rules in §16–§21 came out of seeds added *after* the previous round
reported clean, and each new seed kept finding something until the eighth. The
fourth round then did it again — 271828 found §24 on a corner that four
previous rounds had each declared settled. A clean run on the seeds you
already have says less than one more seed does.

Four entries this list used to carry are gone: `("011" < False)` is §16,
`s13579/141` is §18, `s77/112` is §22, and §16's unexplained cell is §24 —
reduced, understood, and either fixed or attributed to Excel.

Known and unfixed:

- **A comparison's static type mismatch surfaces before its right operand
  runs.** `TypeName(True) = (0.0001 \ False)` is error **13** in Excel and
  **11** here, though both agree that `0.0001 \ False` alone is 11 and that
  `TypeName(True) = 0` alone is 13. So Excel evaluates the left operand,
  finds a declared `String` that will not convert against a statically
  numeric partner, and raises without ever running the division; visi
  evaluates both operands and the division's error wins. The rule is
  understood — it is §13's strictness applied one step earlier — but
  implementing it moves evaluation order in the comparison path, which is
  where §2–§4's measured orderings live, so it wants its own round of
  measurement rather than a change made alongside §24. Accounts for seed
  31459 case 2.
- **The `STATICALLY_*` lists are incomplete, and `Abs`/`Int` need a third
  category.** Measured against a string **literal**, which §24 says is the
  strict partner: `InStr(1, "a") = "abc"`, `InStr("ab", "b") = "abc"` and
  `Asc("a") = "abc"` are all error 13 in Excel and `False` here, so `InStr`
  and `Asc` belong in `STATICALLY_NUMERIC`. But `Abs(-1) = "abc"` is *also*
  error 13, and this file already records `Abs(a)` as **not** strict against
  `(-32768 & -2.5)`. Both are right: `Abs` and `Int` return the type they
  are handed, so they are statically typed exactly when their **argument**
  is — a third category alongside "declared type" and "Variant", and one
  that needs its own sweep of all 46 intrinsics rather than a guess per
  name. Adding `InStr` and `Asc` alone would leave the list wrong in the
  more interesting direction. Accounts for seed 19990101 case 171.
- **Which of two faults surfaces first, again.**
  `((Not 100000) ^ CStr(0.1)) + Val(100000 / False)` is error **11** in Excel
  and **5** here: both agree that `(-100001) ^ "0.1"` alone is 5 and that
  `Val(100000 / False)` alone is 11, so the two engines disagree only about
  which operand of the `+` is reached first. Same family as §2–§4 and as the
  comparison case above. Accounts for seed 19990101 case 98.
- `("1-2.5" = 1)` and `("1-2.5" = -1)` are both error 13 in Excel, where
  visi's numeric-prefix scan reads a `1`. Trailing-sign parsing (§10) is
  presumably what makes the prefix ambiguous, but the scan's exact rule was
  not measured. Note that seed 31459 case 114 *looked* like this and was not:
  see §25.

`Join`, `Format`, `StrConv`, `Left$` and `UCase$` are unimplemented (error 35),
noticed while measuring §23 — the `$`-suffixed forms matter beyond
completeness, since they are the typed-`String` half of the split §23 turns
on. `STATICALLY_STRING` already lists `Join` for that reason.

`IsArray` and `IsError` are also unimplemented (error 35); Excel treats both
as statically Boolean in `Select Case`, and `STATICALLY_BOOLEAN` already lists
them so they arrive with the right behaviour.

### The workflow that has resolved every one so far

1. Run the fuzzer on a seed you have **not** tuned against.
2. Reduce to the smallest expression that reproduces it. If that stalls,
   instrument the generated procedure statement by statement and diff the
   intermediate `TypeName|CStr` values against Excel's -- that is how the
   prefix rule was finally cornered, after inspection had failed.
3. Probe with the neighbouring cases that would discriminate between
   plausible rules, and check whether the probe's own rendering can confound
   the answer. Observe a possibly-`Null` result with `IsNull`, never `CStr`:
   `CStr(Null)` is itself error 94, and that confound produced two separate
   false conclusions in this work.
4. Implement **one change at a time**. Bundling three regressed six of seven
   seeds and had to be unpicked afterwards; split up, one of the three turned
   out to be net-negative on its own.
5. Re-run on several seeds, including ones not used while developing the fix.
   A fix that improves the tuned seeds and leaves fresh ones unchanged is
   still a real fix; one that only moves the tuned seeds is overfitting.
