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

Sections 16–17 are a third round, found while building Phase 2 — and worth
noting for *how* they were found rather than what they say. §16 overturns
§11, which had been derived from two cases that could not tell two rules
apart and had stood since Phase 1. §17 is a pair of divergences that a return
value cannot expose at all: the macro returned the right number and left the
wrong thing in a cell. Both came out of `fuzz_vba.py` once it started
comparing the data grid, which is the argument for that change in one
sentence.

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
[`excel-discrepancies.md` §16](excel-discrepancies.md).

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

**The Boolean side decides which comparison runs.** A literal converts;
anything else statically known compares as text and never errors:

| Expression | Excel | Why |
| --- | --- | --- |
| `TypeName(0) >= False` | error 13 | literal partner → `CBool("Double")` raises |
| `TypeName(0) >= (3# >= Empty)` | `False` | folded partner → text, `"Double" < "True"` |
| `CStr(0) >= (3# >= Empty)` | `False` | likewise; the conversion would say `True` |
| `(3# >= Empty) >= TypeName(0)` | `True` | the same, reversed — it is not side-specific |

`(Not True)` counts as a **literal** here and `(3# >= Empty)` does not, which
is the subtlest part: VBA folds a unary chain over a literal but not a
comparison. `operand_kind` had to become recursive to tell them apart —
it previously looked one level down and put `(Not True)` in the wrong bucket.

**The String side decides whether either applies.** A folded constant
expression takes neither and stays on the numeric path, which is the only way
`((Empty & "1") <= ("" <> Empty))` comes out `False`.

And on a failed conversion, what happens depends on how well the compiler
knows the string's type — the same static/runtime axis as §7, §11 and §13:

| String operand | Failed conversion | Example |
| --- | --- | --- |
| runtime `Variant` | compares as text | `a = TypeName(32767) : a >= (Not True)` is `True` |
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

**Done**, and verified as a whole: 27 cases spanning every cell of both axes
agree with Excel. `value::compare_ctx` and `interp::operand_kind`, covered by
`a_string_converts_with_cbool_against_a_static_boolean`.

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
§16 and §17:

| Seed | Cases | Agreed |
| --- | --- | --- |
| 909 | 200 | **200** |
| 4242 | 200 | **200** |
| 55555 | 200 | **200** |
| 77 | 200 | 199 |
| 13579 | 200 | 199 |

998 of 1000. The two remaining are in the list below; both are pure error
ordering, with every cell agreeing.

`("011" < False)` is no longer among these — see §16, which explains it.

Known and unfixed:

- `("1-2.5" = 1)` and `("1-2.5" = -1)` are both error 13 in Excel, where
  visi's numeric-prefix scan reads a `1`. Trailing-sign parsing (§10) is
  presumably what makes the prefix ambiguous, but the scan's exact rule was
  not measured.
- `a = -3 : (Len(CStr(a)) / (-32768)) = ((-7) & (0 > "1.5"))` is error 13 in
  Excel and `False` here. A *constant* string against a runtime number is
  documented (§"String against number, in full") as comparing as text, and
  `"2" > b` with `b = 10` being `True` is what established that — but a
  constant string that does not parse numerically appears to be error 13
  rather than an unequal text comparison. The candidate rule is "coerce
  first, then compare as text", the same shape as `"abc" - Null` being error
  13; it needs two or three more measurements before it is worth
  implementing. Found by the Phase 2 fuzzer (`s13579/141`).
- `s77/112` returns `String|Z` here and error 6 in Excel, from a procedure
  whose sub-expressions all agree individually (`3# ^ 2147483647` is `INF` on
  both sides, `Len(CStr(INF))` is 3 on both). Not reduced; the reproduction is
  saved under `fuzz_results/failures/`.

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
