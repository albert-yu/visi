# VBA error ordering and `Null` handling: scope for the next pass

Scoping for the divergences left after Phase 1 of
[`vba-macro-support.md`](vba-macro-support.md). Findings measured against
Microsoft Excel for Mac 16.112 on macOS 26.6.1, with
[`fuzz/vba_ordering_probe.bas`](../fuzz/vba_ordering_probe.bas).

Phase 1 landed at 493 of 500 generated procedures agreeing with Excel on
value, subtype and error number together. This document is about the
remainder. **Every one of them is now root-caused**, so what follows is a
work plan rather than an investigation.

The headline: none of these is a wrong *value*. They are all disagreements
about **which error surfaces**, or about **whether `Null` propagates**, and
each turns out to be a small, local change.

---

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

| Expression | Excel | visi today |
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

**Change:** in `value::div`, when the divisor is zero, raise
`VbaError::overflow()` if the dividend is also zero and
`VbaError::div_by_zero()` otherwise. Two lines.

Accounts for fuzz cases 398 and 402.

---

## 3. Both operands are coerced before the divisor is tested

| Expression | Excel | visi today |
| --- | --- | --- |
| `"xxxx" / 0` | error 13 | 11 ❌ |
| `0 / "xxxx"` | error 13 | 13 ✅ |
| `"" / 0` | error 13 | 11 ❌ |

`value::div` currently reads the divisor, tests it for zero, and only then
touches the dividend — so a type mismatch on the left is masked by a
division-by-zero on the right. Excel coerces both first, and the type
mismatch wins.

**Change:** coerce the dividend before testing the divisor. One line moved.

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

**Change:** give `value::pow` an `ArithMode` parameter, as its siblings have,
and raise error 6 on a non-finite result in `Constant` mode. The interpreter
already computes the mode at every binary operator, so the call site needs
one extra argument.

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

Re-run with `IsNull(...)` instead, the answer is a split:

| Propagate `Null` | Raise error 94 |
| --- | --- |
| `Abs`, `Int`, `Fix`, `Len`, `UCase`, `Trim`, `Left`, `Mid`, `InStr` | `Sgn`, `Val`, `Sqr`, `Replace` |

There is no obvious principle behind the split — `Abs` propagates but `Sgn`
raises; `Len` and `Left` propagate but `Replace` raises — so it has to be a
table, and the table has to be measured rather than reasoned about.

`builtins::call` currently propagates `Null` for everything outside a small
inspection list, so the four in the right-hand column are wrong.

**Change:** add those four to the existing `rejects_null` set in
`builtins::call`, and extend the probe to cover the rest of the intrinsic
library before trusting the propagating column beyond what is listed.

Accounts for fuzz case 442.

**Open:** only 13 of ~45 intrinsics have been measured. The rest should be
swept before this is called done — that sweep is most of the remaining work
in this document.

---

## 6. Still unexplained

Fuzz case 233 (`visi: ERR|13`, `excel: OK|Double|-2.5`) is not covered by any
of the above. Its trigger looks like `Trim(<Boolean>) Xor <number>` —
`Trim` renders the boolean as `"False"`, and visi then fails to coerce that
string to a number, while Excel produces a value. The hypothesis is that
VBA's numeric coercion accepts `"True"` and `"False"` as `-1` and `0`, which
`value::parse_vba_number` does not.

**Not yet measured.** It needs its own probe — `CDbl("True")`, `"True" + 1`,
`"False" Xor 1` — before anything is changed, because guessing at coercion
rules is exactly what this project has been burned by twice already.

---

## Suggested order

1. §2, §3 and §4 together — three small, fully-determined changes in
   `value.rs`, each with a measured test. Should close fuzz cases 187, 251,
   398 and 402.
2. §5's four-function correction, then the full intrinsic sweep.
3. §6's probe, then whatever it says.

After each, re-run `python fuzz/fuzz_vba.py --iterations 500 --batch 25
--seed 1` and check the count moves the right way. Expect a handful of new
long-tail cases to surface as the current ones clear — that has happened at
every round so far, and it is the harness working, not regressing.
