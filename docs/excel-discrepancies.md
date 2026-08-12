# Known discrepancies with Microsoft Excel

Last updated: 2026-08-12

Cases where `libvisi` and real Microsoft Excel (verified against 16.111.3 on
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

## 4. Rounding ties in the incomplete beta — *No stable answer*

The F and t distributions are computed from a continued-fraction
incomplete beta, accurate to a few ULP. That is enough to agree with Excel
on the value, but not always on the 15th digit Excel *prints*, because
some inputs land almost exactly on a rounding tie:

```
FTEST over one fuzzed pair
  true value  0.94171633283387507291   (40-digit mpmath)
  Excel       0.941716332833875
  visi        0.941716332833876
```

The true value sits **0.7 ULP** above the point where the 15th digit flips
from 5 to 6. Rounding it the way Excel does requires better-than-ULP
accuracy, which no double-precision implementation can promise, so this
particular input has no answer an independent engine can reliably
reproduce.

Not excluded, because it is rare (roughly one cell per 60 iterations) and
indistinguishable at generation time from the ordinary cases the same
functions cover. Worth checking rather than assuming, though: visi's
incomplete beta is now within a few ULP throughout, and on
`F.DIST.RT(120.02429320013077, 2, 4)` it is **37x closer to the true value
than Excel** (3.5e-16 relative against 1.3e-14) — so a disagreement here
is at least as likely to be Excel's error as visi's. Arbitrate with
`mpmath` before changing anything.

## 5. FORECAST.ETS.SEASONALITY — *No stable answer*

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

## 6. DATEDIF `"YD"` — *No stable answer*

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

## 7. Odd-coupon bond functions — *visi gap*

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

## 8. AMORDEGRC — *visi gap*

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

## 9. ACCRINT from a February month-end — *visi gap*

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

## 10. QUARTILE.EXC — *visi gap*

```
QUARTILE.EXC(F1:G5, 3)   visi #NUM!   Excel 53
```

visi's exclusive-quartile interpolation rejects some quart/sample-size
combinations Excel accepts. `QUARTILE.INC` and the `PERCENTILE.*` family
agree.

## 11. RATE — *No stable answer*

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

## 12. FREQUENCY with non-numeric bins — *visi gap*

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

## 13. Error-class precedence in composed expressions — *tolerated by the comparator*

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

## Fixed, not excluded

For contrast, these looked like Excel divergences during investigation and
turned out to be real visi bugs, now fixed and covered by regression tests:
the whole `IM*` complex family (an unimplemented stub), `erf`/`erfc`
precision, the `inv_incbeta` solver behind `BETA.INV`/`F.INV`, `AGGREGATE`
aggregating its own options argument, `FREQUENCY` bin handling, `DATEDIF`'s
other units, `ROMAN`'s concise forms, `GAMMALN`'s domain, `QUOTIENT`'s
boolean rejection, end-of-month coupon schedules, the number→text formatting
range, and the actual/actual (basis 1) day-count denominator.

A later pass added: `DAYS360`'s US method (the February month-end rules were
missing, so it silently returned the European answer for both spellings);
supplied-but-blank arguments defaulting instead of counting as 0, which made
`LOG(1, <blank>)` return 0 where Excel says `#NUM!`; the direct-argument text
rule across the whole statistical family (`DEVSQ("abc", 3, 4, 5)` answered 2
instead of `#VALUE!`); `ERF`/`ERFC` rejecting numeric text they should coerce;
`SUMPRODUCT` erroring on non-numeric entries that Excel counts as zero;
`POWER(0, 0)` returning 1 rather than `#NUM!`; `CHITEST` rejecting negative
expected values per element when Excel only rejects a negative total; and
number→text keeping more than Excel's 15 significant digits. See `git log`.
