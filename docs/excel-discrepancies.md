# Known discrepancies with Microsoft Excel

Last updated: 2026-08-31

Cases where `visi-core` and real Microsoft Excel (verified against 16.111.3 on
macOS) disagree, and which are therefore **excluded from the differential
fuzz harness** in `fuzz/fuzz_excel.py`.

Everything here has been reduced to a specific, reproducible case. Nothing is
listed as "known" merely because it was inconvenient — an item earns a place
here only once the root cause is understood well enough to say *which* engine
is right, or to say precisely why the question has no stable answer.

Three kinds of entry:

| Kind | Meaning |
| --- | --- |
| **Excel is wrong** | visi is measurably more accurate. Do **not** change visi to match. |
| **visi gap** | A real shortfall in visi. Fixing it is worthwhile; the exclusion stops it drowning out new regressions until then. |
| **No stable answer** | Excel's behaviour is a heuristic or is internally inconsistent, so no independent implementation can agree with it by construction. |

---

## 1. BESSELI / BESSELJ / BESSELK / BESSELY — *Excel is wrong*

Excel's Bessel routines lose accuracy well before visi's do. Arbitrated
against 60-significant-digit references (a `decimal` evaluation of the same
ascending series):

| Call | True value (18 s.f.) | visi rel. error | Excel rel. error |
| --- | --- | --- | --- |
| `BESSELJ(9.59, 1)` | 0.141754162508486557 | 1.6e-14 | **1.8e-6** |
| `BESSELJ(8.72, 2)` | 0.0795586089024345565 | 6.5e-15 | **1.3e-6** |
| `BESSELJ(2.95, 3)` | 0.300141005800689674 | 2.5e-16 | **3.8e-7** |

Excel's error grows with the *order* as well as the argument, so no argument
cap keeps it usable as a reference — capping x at 8 was tried first and still
failed at order 3. The whole family is excluded from the harness; visi's own
accuracy is pinned directly against those references by
`test_besselj_stays_accurate_where_excel_does_not`.

## 2. XIRR on series with no internal rate of return — *Excel is wrong*

Given a cashflow series that has no root, Excel does not report `#NUM!` — it
returns a non-answer:

| Excel's XIRR | XNPV at that rate |
| --- | --- |
| -0.92945409 | -184430.99 |
| 2.98e-09 | -34415.90 |
| -0.89982008 | -8804.04 |

For the first two, XNPV has no sign change anywhere in `(-0.999, 10)`: no root
exists to find. visi's `#NUM!` is correct. Rather than exclude XIRR, the
harness now generates *well-posed* cashflows (one outlay followed by returns
that more than repay it), so IRR/XIRR/MIRR always have a unique root and the
comparison stays meaningful.

## 3. QUOTIENT past 2^53 — *Excel is wrong*

Above the exactly-representable integer range Excel's `QUOTIENT` drifts while
visi's stays correct:

```
QUOTIENT(123456789012345678, -49)
  exact  -2519526306374401.59...
  visi   -2519526306374401        error  0.59
  Excel  -2519526306374387        error 14.59
```

Arbitrated with 60-significant-digit `decimal` arithmetic, so this is not a
matter of which rounding convention to prefer — Excel is 24x further from the
true quotient, and in the same direction as its other precision limits (see
section 1). visi is left alone.

This is not excluded: it needs operands above 2^53 to show up at all, so it
costs roughly one cell per few hundred iterations, and the same generator
range is what exercises the number-to-text formatting rules that *did* turn up
real bugs.

## 4. FORECAST.ETS.SEASONALITY — *No stable answer*

Excel's automatic season-length detection does not report the series' true
period, and its answer turns on the *arrangement* of the seasonal offsets
rather than their magnitude. Over 16 points with slope 2 and the same four
offsets merely permuted:

| Offsets | Excel |
| --- | --- |
| `[8, -2, -8, 2]` | 4 |
| `[11, -11, 2, -2]` | 2 |
| `[2, -2, 11, -11]` | 2 |
| `[-2, -11, 11, 2]` | **0** (for a series that is exactly period-4) |

Trend strength is not the trigger — holding the offsets fixed and sweeping the
slope from 0 to 4 leaves the answer at 4 throughout.

`FORECAST.ETS.SEASONALITY` is excluded, and every other ETS call in the
harness is handed an **explicit** season length, so what gets compared is the
model (timeline handling, the Holt-Winters recurrences, extrapolation) rather
than the detector. Detection is covered by visi's unit tests on the patterns
where Excel's answer *is* the true period.

Related, and the reason the ETS series are exactly-modelled: Excel fits
alpha/beta/gamma with a proprietary optimizer and reports them to three
decimals. On noisy data no independent implementation lands on the same
digits, so `FORECAST.ETS.STAT` types **1-3** (the fitted parameters) are
excluded too. Types 4-8 (MASE/SMAPE/MAE/RMSE/step) are well-defined and are
fuzzed.

## 5. DATEDIF `"YD"` — *No stable answer*

Excel's `"YD"` is internally inconsistent. Tested against 8 real-Excel data
points, no candidate rule fits:

| Rule | Fits |
| --- | --- |
| Anchor the start date in the end date's year | 5/8 |
| Anchor the end date in the start date's year | 5/8 (a *different* 5) |
| `total days - 365 * whole years` | 0/8 |
| Remap both into 1900/1901 (Excel's phantom leap day) | 3/8 |

Microsoft documents DATEDIF's unit codes as only partially supported. visi
keeps the defensible definition — days since the most recent anniversary of
the start date — and `"YD"` is excluded from the harness. The other units
(`"Y"`, `"M"`, `"D"`, `"MD"`, `"YM"`) agree with Excel and stay fuzzed.

## 6. Odd-coupon bond functions — *visi gap*

`ODDFPRICE` and `ODDFYIELD` disagree on odd-first-coupon configurations where
Excel returns `#NUM!` and visi computes a value, e.g.

```
ODDFPRICE(DATE(1995,6,6)+57, EDATE(DATE(1995,6,6)+57,30), DATE(1995,6,6), …)
  visi 100.2656…   Excel #NUM!
```

Excel rejects certain settlement/issue/first-coupon orderings that visi
accepts. The exact admissibility condition has not been pinned down. Both are
excluded pending that work; the regular-coupon functions (`PRICE`, `YIELD`,
`COUPDAYBS`, `COUPNCD`, `COUPPCD`, `COUPNUM`, …) agree and stay fuzzed.

## 7. AMORDEGRC — *visi gap*

The French declining-balance depreciation still disagrees on some schedules,
sometimes by one unit and sometimes substantially:

```
AMORDEGRC(48665.34, DATE(1998,12,22), EDATE(…,11), 7901.84, 13, 0.05, …)
  visi 777   Excel 1085
```

The running balance now carries full precision (fixed — that accounted for the
off-by-one cases), but the coefficient brackets and the switch to straight
line at the end of life are not fully reverse-engineered. Excluded pending
that work.

## 8. ACCRINT from a February month-end — *visi gap*

With an issue date on a February month-end, ACCRINT accrues slightly less
than Excel does:

```
ACCRINT(2003-02-28, +6mo, +24mo, 0.0171, 34973.86, 2, 0, FALSE)
  visi 1192.783   Excel 1196.106      (Excel = exactly 4 coupons)
ACCRINT(2004-02-29, +6mo, +18mo, 0.05, 10000, 2, 0, FALSE)
  visi 745.833    Excel 748.611
```

Both Excel answers equal `par * rate * NASD-30/360 days(issue, settlement)
/ 360` — i.e. the whole span counted once, not summed period by period.
But that model is not what ACCRINT does in general, because the result
*does* depend on `frequency`: for one span, Excel gives 608.33 at
frequency 1 and 2 and 483.33 at frequency 4, and 483.33 corresponds to
accruing from the second quasi-coupon date rather than from the issue.

So the whole-span model fits the February cases and contradicts the
frequency cases, and the period-walk model fits the frequency cases and
contradicts the February ones. A "a period spanned end to end is worth
exactly one coupon" rule was tried and fixed the first case above while
breaking the second. Excel's actual schedule rule is not understood, so
this stays a gap rather than a guess; the harness avoids February
month-end issue dates for ACCRINT and everything else about it agrees.

Note this is *not* the DAYS360/YEARFRAC divergence found alongside it —
that one turned out to be real and is now implemented. Excel's `DAYS360`
function and its `YEARFRAC` basis 0 use genuinely different 30/360 rules
(`DAYS360(2003-02-28, 2005-02-28, FALSE)` is 718 while
`YEARFRAC(...) * 360` is 720), verified over twelve date pairs and
covered by `test_days360_and_yearfrac_use_different_thirty_360_rules`.

## 9. QUARTILE.EXC — *visi gap*

```
QUARTILE.EXC(F1:G5, 3)   visi #NUM!   Excel 53
```

visi's exclusive-quartile interpolation rejects some quart/sample-size
combinations Excel accepts. `QUARTILE.INC` and the `PERCENTILE.*` family
agree.

## 10. RATE — *No stable answer*

Excel's `RATE` iterates from its `guess` (default 0.1) and gives up with
`#NUM!` on series where a root demonstrably exists — handed a guess near that
root it finds it:

```
RATE(275, -1582.09, 12951.19, 0, 0)         Excel #NUM!
RATE(275, -1582.09, 12951.19, 0, 0, 0.12)   Excel 0.12215788664979609
```

So the `#NUM!` is a statement about Excel's iteration from its default
starting point, not about the problem having no answer. visi's solver is more
robust and reports the root. (The genuinely-degenerate direction *was* a visi
bug and is fixed: for an annuity-due with `fv = 0` the payment term carries a
factor of `(1 + r)`, so `r = -1` satisfies the equation for any inputs, and
the iteration used to slide into that basin and report ~-0.9999 as if it were
an answer.)

Excluded because "did the other engine's iteration happen to converge from
0.1" is not a property worth asserting. `IRR`, `XIRR`, `MIRR`, `NPV` and the
rest of the TVM family stay fuzzed.

## 11. FREQUENCY with non-numeric bins — *visi gap*

When `bins_array` contains blanks, booleans or text, visi and Excel disagree
on both the bucket contents and the *length* of the result. visi drops
non-numeric bins entirely (so the bin count collapses and every bucket
shifts); Excel keeps some of them.

Excel's exact rule is not understood. Probing data `{-78, -393.28, 54, "I",
36}` against bins `{<blank>, "fpiijWIx", "ST"}` returns a **two**-element
result `{2, 2}` — consistent with dropping the two text bins but keeping the
blank one as 0. Implementing that reading, however, made agreement *worse*
across a 40-iteration run (3 mismatches became 8, in both directions), so it
is not the rule either. Excluded until it is pinned down properly rather than
guessed at.

An all-numeric `bins_array` agrees, including the non-obvious part that Excel
sorts the bins internally but reports each count back at that bin's original
position — that is covered by a regression test.

## 12. Error-class precedence in composed expressions — *tolerated by the comparator*

When several sub-expressions of one formula each produce a *different* error,
visi and Excel sometimes surface different ones:

```
FISHER(AND(H5 > 0, I7 < 100))                     visi #DIV/0!   Excel #VALUE!
(SUMX2PY2(F6:G6, I7:I7) ^ AVEDEV(I5:J7))          visi #DIV/0!   Excel #VALUE!
IFERROR(FACT(RSQ(…)), (FTEST(…) - MOD(…)))        visi #VALUE!   Excel #DIV/0!
```

Which error wins depends on Excel's internal evaluation order, and it differs
per operator and per function. It cannot be excluded by dropping a function
from the generator, because it is emergent from the random expression trees
rather than attached to any one function — and those trees are where a lot of
the harness's value lies.

It is therefore handled in the **comparator** instead: a disagreement where
*both* engines produced an error, differing only in class, is counted
separately rather than as a failure. Crucially it is still counted and
printed in the run summary:

```
 Tolerated: 28 cell(s) where both engines errored with different error classes
```

so it can never quietly hide a regression. Pass `--strict-error-class` to
treat these as failures again — worth doing periodically, since strict
comparison is exactly what surfaced genuine bugs like `TYPE(error)` and
`ERROR.TYPE` returning the wrong value, `LOG(n, 1)` being `#NUM!` rather than
`#DIV/0!`, and CHITEST's `#N/A` cases.

Each individual case is cheap to investigate — the failure artifacts under
`fuzz_results/failures/` carry the source workbook alongside both engines'
output.

---

## 13. Empty-string cell vs. blank cell — *fixed in the comparator*

Excel distinguishes a cell holding the empty string from a cell holding
nothing; visi does not, deliberately and consistently — `ISBLANK("")` is
TRUE, `COUNTA` skips it, and `rust_xlsxwriter` collapses an empty string to
a blank cell on write (`store_string` turns `""` into `write_blank`).

The harness generates short strings from an alphabet that includes a space,
so it sometimes produces a cell containing only whitespace. OOXML strips
whitespace-only `<t>` content unless the element carries
`xml:space="preserve"`, which `openpyxl` does not emit, so *neither* engine
recovers the space: Excel keeps an empty-string cell (a `<t/>` entry in the
shared strings) and visi writes no cell at all. Both mean "nothing here".

This surfaced as a spurious failure:

```
Cell C8 on sheet1: visi=None | Excel= (Formula: None)
```

The comparator already had the right rule — `values_equal` treats `None` and
a whitespace-only string as equal — but `compare` never reached it when the
cell was *absent* from visi's output rather than present-and-empty. That path
guarded on `val is not None`, so a blank-equivalent value was reported as
"Missing in visi output" while the identical disagreement between two
*present* cells was tolerated. The two paths now apply the same rule.

Note this is narrow: it only excuses a value that is blank-equivalent. A cell
genuinely missing from either side still fails, so real data loss on import
or export is still caught.

A later fuzz pass found one more place where the same OOXML whitespace stripping
is observable: `ARRAYTOTEXT` over a one-cell whitespace string. Excel still has
an empty-string cell and returns an empty string; visi imports it as blank and
applies the real blank-cell rule (`#VALUE!`). The random generator now avoids
whitespace-only source strings rather than treating an xlsx encoding artifact as
a formula semantics failure.

---

## 14. MOD with a divisor far larger than the dividend — *Excel is wrong*

When `|d|` is enormous relative to `|n|` and the signs differ, Excel returns
`0` where the remainder is not zero:

```
MOD(36, POWER(-327.3, 69))    visi -3.3984E+173   Excel 0
MOD(1, -10^37)                visi -1E+37         Excel 0
```

Both engines agree on the same shape at ordinary magnitudes — `MOD(5, -3)` is
`-1` in both, and so is `MOD(5, -1E10)` = `-9999999995` — so this is not a
disagreement about the definition. `MOD(n, d) = n - d*INT(n/d)` gives
`INT(n/d) = -1` for every one of these, hence a remainder of `d + n`, which
sits inside `(d, 0]` exactly as a remainder must.

A remainder of `0` would require `d` to divide `n` exactly, and it cannot:
`0 < |n| < |d|`. Excel's answer is unreachable from its own documented
formula, so this is Excel losing the small operand rather than visi being
imprecise, and visi keeps the mathematically correct value.

Note this is *not* the same as the deliberate `#NUM!` cutoff already in
`MOD` for large **quotients** (`MOD(28^31, 3)`), which is a real Excel
behavior visi reproduces. Here the quotient is vanishingly small; it is the
result that is large.

A second, unrelated mechanism lands in the same "Excel's `INT(n/d)` is off"
bucket at ordinary magnitudes: `MOD(-47, 47 / -13)`. `-47 / (47/-13)` is
exactly `13` mathematically (`47` cancels), and stays exactly `13.0` even
carried through the actual `f64` division of the two doubles — no rounding
residue at all — so `INT` of it is unambiguously `13` and the true remainder
is `0`, which is what visi returns. Excel instead returns
`-3.615384615384615`, the divisor itself, as if its own `INT(n/d)` had
landed on `12`. Unlike the large-divisor case above this isn't a magnitude
problem — every value involved is an ordinary double — so it looks like
Excel's own division/`INT` sequence for this particular ratio rounds down
one step early. The same shape came back through the new `PERCENTOF` generator:
`MOD(-14, PERCENTOF(-14, 37))` should be exactly zero, while Excel returns the
divisor (`-0.378378...`). The generator now avoids building `MOD` operands out
of `PERCENTOF`, for the same reason it already avoids `POWER`/`^` there; visi's
correct behavior remains pinned by
`test_fuzz_mod_stays_exact_at_an_integer_quotient_boundary`.

## 15. VBA: an infinity poisons the next string-to-number conversion — *Excel is wrong*

VBA's `^` is the one operator that returns an infinity rather than raising
overflow (`a = 3.75 : a ^ 32767` is the `Double` `INF`, measured). Once a
procedure has produced one, the **next string-to-number conversion anywhere in
that procedure raises error 6**, whatever string it is given:

```vba
a = 3.75
b = a ^ 32767        ' Double INF, no error
c = CDbl("1.5")      ' Excel: error 6.  visi: 1.5
```

Nothing about `"1.5"` overflows. The same statement on its own, or after
`b = 1`, converts fine. What the probe shows is a status flag rather than a
rule:

| After `b = a ^ 32767` | Excel |
| --- | --- |
| `CDbl("1.5")`, `CSng("1.5")`, `"1.5" * 1`, `"1.5" <> 0` | error 6 |
| `Val("1.5")` | `1.5` — a different parser |
| `CDbl("abc")` | error 13 — the type check comes first |
| `1 <> 0`, `"abc" & 1` | fine — no conversion involved |
| an intervening `b = 1` or `c = 1 + 1`, then `CDbl("1.5")` | still error 6 |
| a *first* conversion swallowed by `On Error Resume Next`, then another | fine |

The last two rows are the tell: the condition is not cleared by unrelated
work, but *is* cleared by being reported once — the behaviour of a sticky
floating-point exception flag that the conversion routine reads and clears.
`Val`, which does not consult it, is unaffected.

visi raises the error where the overflow actually happens and nowhere else.
Matching this would mean modelling an FPU status word across statement
boundaries, to reproduce an error on a conversion that cannot overflow.

Measured with `fuzz/vba_expr_probe.py`; it is why `fuzz_vba.py` case
`s555/219` is left as a known divergence rather than chased.

---

## 16. VBA: `Err.Number` on a `Range` whose cells were deleted — *No stable answer*

Excel's `Range` objects track a structural edit: `Set r = ws.Range("A5")`
followed by `ws.Rows(1).Insert` leaves `r` reading `$A$6`, and still holding
the value that was in `A5`. visi matches that, by interning ranges. But when
the edit deletes *every* cell a range covered, the object enters a state with
no reproducible error number:

```vba
Set r = ws.Range("A5")
ws.Rows(5).Delete
s = r.Address              ' raises
```

What is stable, and what visi reproduces exactly:

| Probe | Excel |
| --- | --- |
| `r Is Nothing` | `False` — it is still an object |
| `TypeName(r)` | `"Range"` |
| `r.Address` | raises, `Method 'Address' of object 'Range' failed` |
| `r.Value` | raises, `Method 'Value' of object 'Range' failed` |

What is not stable is `Err.Number`. The *same* case returned `-1667945984` on
one run and `-1667949824` on the next; two different members returned the same
number on one run and different numbers on another. These are raw automation
error values from Excel for Mac, not the documented VBA error codes, and they
do not survive a re-run.

visi raises **1004** with Excel's description text. 1004 is what Excel on
Windows documents for `Method '<name>' of object '<object>' failed`, and it is
the number `vba/host.rs` already uses for the rest of the object-defined error
family (a bad address, an out-of-sheet `Offset`, a failing
`WorksheetFunction`). Pinning Excel for Mac's number would be pinning noise.

Measured with `fuzz/vba_range_tracking_probe.py`, which also establishes that
the *geometry* of the tracking — move vs. grow vs. shrink — is identical to
`core::grid_edit`'s rules for a formula's range reference, case for case.

---

## 17. A tiny power underflows to exactly 0 in Excel — *Excel is wrong*

`FACT(15) ^ -26` (`1307674368000 ^ -26`) is a legitimate positive subnormal,
about `9.35e-316`. visi computes it via `f64::powf` and gets that value, so
`(FACT(15) ^ -26) > 0` is `TRUE`. Real Excel's `^` returns exactly `0.0` for
this call instead, so the same comparison is `FALSE`.

Both engines agree at ordinary magnitudes; this is specifically an
extreme-magnitude base raised to a large negative exponent, well past where
the true result underflows the *normal* `f64` range but is still representable
as a subnormal. Excel's own `^` implementation evidently doesn't handle
subnormals here (likely computing `exp(y * ln(x))` without a path back down
into subnormal territory), while Rust's `powf` does. The true mathematical
value is unambiguously nonzero, so visi's answer is the accurate one — the
same shape as the BESSEL entries at the top of this document.

Found via fuzz/fuzz_excel.py, seed 795107:
`XOR((FACT(15) ^ -26) > 0, ...)` disagreed only because of this one operand.
The generator now avoids `^`/`POWER` calls whose result would underflow to a
subnormal, rather than chasing Excel's underflow threshold.

## 18. Negative base raised to a tiny fractional exponent — *No stable answer*

`-2 ^ POWER(-15, -6)` — precedence-wise this is `(-2) ^ (POWER(-15, -6))`,
not `-(2 ^ POWER(-15, -6))`: `-2^2` really is `4` in Excel's formula
language, unary minus binding *tighter* than `^` there (the opposite of
VBA's `^`, and the opposite of most languages' convention) — confirmed
directly (`=-2^2` is `4` in real Excel), so visi's existing precedence
here was already correct and needed no fix. `POWER(-15, -6)` is
`8.779...e-08`, a tiny positive fraction, so the outer call is a negative
base raised to a non-integer exponent — mathematically undefined over the
reals. visi's `#NUM!` for that is the principled answer, and matches real
Excel for most such exponents: `(-2)^0.5`, `(-2)^0.1`, `(-2)^0.01`,
`(-2)^1e-6`, `(-2)^1e-8` are all `#NUM!` too. But not every one:
`(-2)^0.2` is a real number in real Excel, `-1.148698354997035`, and so —
unpredictably — is `(-2)^POWER(-15,-6)` (`-1.0000000608524293`, matching
neither `-(2^y)` nor any other describable transform of `2^y` we could
find). Sweeping exponents from `1e-8` to `0.5` in fine steps found no
pattern separating the handful of exponents that return a real number
from the great majority that return `#NUM!` — this looks like numerical
noise in Excel's own `^` implementation (most plausibly an internal
complex-domain evaluation whose imaginary part fails to cancel to exactly
zero for almost every input, and does for a few), not a rule. No
independent implementation could reproduce it without hard-coding the
exact bit patterns that happen to trigger it, so visi's consistent
`#NUM!` stands.

Found via fuzz/fuzz_excel.py, seed 245879: `TYPE((-2 ^ POWER(-15, -6)))`
disagreed (visi `16`, an error code; Excel `1`, a number) purely because
of this. The generator now avoids `^`/`POWER` calls that could combine a
possibly-negative base with a fractional exponent close to zero.

## 19. VBA: a never-executed statement can change which error a procedure raises — *Under investigation*

`fuzz/fuzz_vba.py` first ran against real Windows Excel this session (a
`win32com` driver alongside the existing macOS AppleScript one), and one
case out of 200 landed here rather than being understood well enough to fix
or fully document:

```vba
Private Function Gen1()
    Dim va, vb, vi, vn
    Dim wsh As Worksheet, vk As Range, vq As Range
    Set wsh = ThisWorkbook.Worksheets("Data")
    va = 7
    vb = 2147483647
    vb = ((Left(32767, 3) & Len(CStr(3))) Eqv ((Not Null) \ ("1.5" = &HFF)))
    wsh.Cells(3, 5).Value = (("Z" & "abc") >= (Not vb))
    va = wsh.Cells(3, 5).Value
    va = (((Not "a") Eqv ("12" Mod vb)) Mod (Val(100000) <> (&HFF Eqv va)))
    Gen1 = va
End Function
```

Real Excel raises **13** (Type Mismatch) calling this; visi raises **94**
(Invalid use of Null). Both numbers are individually explicable —
`(Not Null) \ (...)` on line 7 is 94 on both engines when isolated, and
`Not "a"` on the last line is a `Not` of a non-numeric string literal, which
is a plausible source of 13 — the puzzle is *which one wins*, and why.

Measured directly (win32com, real Windows Excel), holding everything else
fixed and varying only how much of the function survives:

| Body kept | Excel's error |
| --- | --- |
| Through line 7 only (the `Not Null` line, `Gen1`/`Harness` inlined into one function) | **94** — matches visi |
| The full function above, called as `Gen1()` from a separate `Harness` | **13** |

Line 7 executes and raises before line 10 (`Not "a"`) is ever reached — VBA
does not roll a statement back or re-run the function once an error fires.
So Excel's 13 cannot come from *executing* line 10; the only thing that
changed between the two rows of that table is whether line 10 (and 8, 9)
exist **anywhere in the compiled procedure**, executed or not. That points
at something in Excel's own compile step for the procedure — the same
"compiles lazily, once, per invoked procedure" step `fuzz_vba_parse.py` is
built around — evaluating (or type-checking) a constant sub-expression like
`Not "a"` and having that outcome override which runtime error the
*procedure* is later reported as raising, rather than the error simply
propagating up from whichever statement actually executed first.

Not fixed and not excluded: the mechanism above is a hypothesis from a
single case, not a rule confirmed against a systematic sweep (varying which
kind of never-executed statement, where it sits in the procedure, and
whether the current statement's own error survives). Chasing that sweep is
future work; guessing at a rule from one data point risks encoding a
coincidence into the interpreter. `visi`'s own error stays principled in
the meantime (94 is genuinely what `(Not Null) \ (...)` raises), so this is
left as a known open question rather than either "Excel is wrong" or a
"visi gap" verdict.

## 20. VBA: an extreme `^` exponent may raise Overflow directly rather than giving infinity — *Under investigation*

Section 15 and `writing_an_infinity_stores_the_num_error_excel_stores`
(`vba/host.rs`) establish, with real-Excel measurements, that `^` never
raises Overflow -- `3.75 ^ 32767` and `-2.5 ^ 1000` both give a `Double`
infinity with no trapped error, and assigning that infinity to a
`Range.Value` leaves a `#NUM!` error *cell* rather than raising anything
in the macro. `fuzz/fuzz_vba.py`'s new win32com driver found a case that
looks like the same shape but isn't: `va` holding a cell-read `33`
(`va = wsh.Cells(1, 5).Value`, itself `=SUM(A1:B2)`), then
`wsh.Cells(2, 6).Value = (va ^ 2147483647)`. visi computes it exactly like
the smaller-exponent cases (infinity, no trap, cell becomes `#NUM!`, the
macro finishes and returns `va` unaffected); real Excel raises a trapped
**Overflow (6)** partway through the same statement.

Not yet isolated to a single cause -- the exponent here (`2147483647`,
`Long.MaxValue`) is far larger than the already-measured `32767`/`1000`
cases, so this may be an exponent-magnitude threshold specific to `^`
itself (distinct from the infinity-assignment behavior in section 15, not
a contradiction of it), or may depend on `va` arriving from a cell read
rather than a literal the way `ArithMode::Constant` vs `::Promote` already
distinguishes for plain arithmetic overflow. Left open rather than guessed
at; a systematic sweep of exponent magnitude (and literal- vs
variable-sourced base) is future work.

## 21. COUPDAYS basis 1 on some quarterly schedules — *visi gap*

Windows Excel's `COUPDAYS(..., basis=1)` does not always equal the actual
calendar length between the surrounding coupon dates, even though that is the
rule visi currently implements. A fuzz case found:

```
COUPDAYS(DATE(2000,11,28), EDATE(DATE(2000,11,28),54), 4, 1)
  visi 92   Excel 91
```

The affected pattern is not yet pinned down. Nearby quarterly schedules can be
92, 91, or 90 depending on the anchor date/year, and the neighbouring coupon
functions still agree. Until the actual Excel schedule rule is reverse-
engineered, the harness avoids basis 1 for `COUPDAYS` only; other bases and the
other coupon-date functions remain fuzzed.

## 22. SORT/SORTBY Unicode text collation — *No stable answer*

`SORT` and `SORTBY` use Windows' locale-sensitive text collation for strings.
The differential generator used a small set of non-ASCII sample strings and hit
contradictory-looking orderings around `"ümlaut"`: it sorts below uppercase
ASCII strings such as `"VQAUU"`/`"Usyqz2oQ"`, but above lowercase generated
strings such as `"gcMdU"`/`"eY"`/`"kDjumH"` in descending `SORT` cases. This is
a Windows collation-table question, not a spreadsheet-calculation rule, and the
answer can vary with locale.

visi keeps its simple deterministic string ordering. The formula fuzzer now
keeps random generated cell text ASCII-only (with punctuation still included)
so dynamic-array sort tests exercise spreadsheet behavior without depending on
locale-specific Unicode collation.

## 23. PRICEMAT/YIELDMAT basis 0 issue-anchored 30/360 schedules — *visi gap*

`PRICEMAT` on some basis-0 schedules whose settlement/maturity are generated by
`EDATE(issue, n)` still disagrees slightly with Windows Excel, for example:

```
PRICEMAT(EDATE(DATE(2002,3,28),6), EDATE(DATE(2002,3,28),11),
         DATE(2002,3,28), 0.0224, 0.04, 0)
  visi 99.25062937062937   Excel 99.26032786885246
```

The existing `PRICEMAT`/`YIELDMAT` special day-count code covers the known
February month-end cases, but this shows another NASD-30/360 leg rule that is
not yet reverse-engineered. The harness avoids basis 0 for `PRICEMAT` and
`YIELDMAT` pending that work; the other bases remain fuzzed.

## 24. MULTINOMIAL returns just below exact integers — *Excel is wrong*

Excel's `MULTINOMIAL` can return a value just below an exact integer even when
all inputs truncate to ordinary non-negative integers:

```
MULTINOMIAL(0.1, 40)      exact 1    Excel 0.9999999999999769
MULTINOMIAL(1.2, 40)      exact 41   Excel 40.9999999999989
```

The worksheet displays these as `1` and `41`, but wrappers that observe the raw
number expose the drift: `INT(MULTINOMIAL(0.1,40))` is `0` in Excel and `1` in
visi. The exact combinatorial definition is unambiguous, so visi keeps the
integer result. The random formula fuzzer no longer generates `MULTINOMIAL`,
while direct coercion/domain behavior remains covered by Rust tests.


## 25. COTH near negative saturation changes integer wrappers � *Excel is wrong*

Excel rounds `COTH` to exactly `-1` for some moderately large negative
arguments where the true value is still just below `-1`. That changes wrappers
that observe the integer boundary:

```
ISODD(INT(COTH(-19)))      visi FALSE   Excel TRUE
```

A high-precision decimal evaluation of `coth(x) = (exp(2x)+1)/(exp(2x)-1)` gives
`COTH(-19) = -1.000000000000000062782655841...`, so `INT` is `-2` and
`ISODD(-2)` is `FALSE`. Excel's `TRUE` requires first rounding the hyperbolic
cotangent to exactly `-1`, which is farther from the mathematical value and
mirrors its known tendency to saturate extreme hyperbolic results too early.
