# VBA error ordering and `Null` handling

Scoping for the divergences left after Phase 1 of
[`vba-macro-support.md`](vba-macro-support.md). Findings measured against
Microsoft Excel for Mac 16.112 on macOS 26.6.1, with
[`fuzz/vba_ordering_probe.bas`](../fuzz/vba_ordering_probe.bas).

Phase 1 landed at 493 of 500 generated procedures agreeing with Excel on
value, subtype and error number together. This document scoped the remainder,
and **all of it has since been implemented** — each section below records the
measurement, and each fix has a unit test naming the Excel result it came
from.

**Status:** every item here is done. On the seed this was developed against,
the harness now reports 300/300. On *unseen* seeds it reports 294–297 of 300,
which is the honest number — see [Overfitting](#overfitting-is-real-here)
below, because the gap between those two figures is the most useful thing in
this document.

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

## What is left

The long tail is error-code disagreements of the same shape as §2–§4: both
engines fault on one expression and report different numbers, because they
coerce its subexpressions in a different order. Examples from unseen seeds:
`ERR|94` where Excel says `ERR|13`, `ERR|6` where Excel says `ERR|5`,
`ERR|13` where Excel says `ERR|94`.

Each needs its own probe — which is tractable but not derivable, and the
pattern so far is that guessing at one of these gets it backwards about half
the time. The workflow that has worked every round:

1. Run `python fuzz/fuzz_vba.py --iterations 300 --batch 25 --seed <unseen>`.
2. Reduce a mismatch to the smallest expression that reproduces it.
3. Put that expression, and the neighbouring cases that would discriminate
   between plausible rules, into a probe modelled on
   [`vba_ordering_probe.bas`](../fuzz/vba_ordering_probe.bas) — and check
   whether the probe's own rendering can confound the answer, as §5's did.
4. Implement, add a unit test naming the Excel result, re-run on several
   seeds.
