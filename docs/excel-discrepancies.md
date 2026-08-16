# Known discrepancies with Microsoft Excel

Last updated: 2026-08-15

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

## 4. Incomplete-beta accuracy — *closed*

The F and t distributions come from a single continued-fraction incomplete
beta, and its accuracy decides how often visi's 15th displayed digit
matches Excel's. Three changes took it from the worst source of
disagreement in the harness to a non-issue:

1. Converge the continued fraction to `f64::EPSILON`, not the textbook
   `1e-15` — that threshold leaves error exactly where the 15th digit sits.
2. Compute the prefactor as `x^a (1-x)^b / (a·B(a,b))` from `tgamma`,
   rather than `exp(a·ln x + b·ln(1-x) − lbeta)`. The log form routes
   everything through one exponential, so the *absolute* error of its
   argument becomes the *relative* error of the result; `lgamma(a+b)`
   alone can put ~5e-16 in. `Γ` for the integer and half-integer arguments
   these families produce is built by recurrence from `sqrt(pi)`, which
   stays near half an ULP where this crate's `tgamma` drifts 2.6 ULP at
   1.5.
3. Correct `(1-x)^b` for the rounding of `1-x`. That subtraction rounds,
   and the power multiplies the slip by `b`: at b = 50 a half-ULP became
   15 ULP. Recovering the exact residual brings the same case to 0.3 ULP.

Measured against 50-digit `mpmath` over 60 random `F.DIST.RT(x, df1, df2)`,
**separating the algorithm's own error from the conditioning of the
input**:

| | median | p90 | max |
| --- | --- | --- | --- |
| algorithm, before | ~12.7 ULP | ~51.6 | ~76.8 |
| algorithm, after | **2.2 ULP** | 7.2 | 16.5 |
| input rounding alone | 0.8 ULP | 10.4 | 19.1 |

That second row is the part visi controls. The third is not a defect in
either engine: `F.DIST.RT` has to form `y = df2/(df1·x + df2)` before it
can call the beta, and `I_y(a,b)` can be steeply sensitive to `y` — at
`df1=30, df2=60` the relative condition number is 22.5, so the unavoidable
half-ULP rounding of `y` is worth ~22 ULP in the answer on its own. Excel
pays exactly the same cost.

**Two earlier versions of this section were wrong**, and the way they were
wrong is the useful part:

- The first called the last failing case an unfixable sub-ULP rounding
  tie. It measured the distance from the true value to the *nearer*
  15-digit candidate instead of to the **tie midpoint** between the two
  candidates. Those differ by half a digit-step — several ULP. The true
  value actually sat 3.85 ULP from the tie while visi was 10 ULP out: a
  plain accuracy bug.
- The second attributed the remaining spread to the continued fraction.
  It did not separate conditioning from algorithm error. Once separated,
  the fraction converges in 2–13 iterations against a cap of 200, and the
  real culprits were the prefactor's `(1-x)^b` amplification and the
  input conditioning above.

So: measure the distance to the **tie midpoint**, and hold the input fixed
when attributing error to an algorithm. And check which engine is wrong
before assuming it is visi — on `F.DIST.RT(120.02429320013077, 2, 4)` visi
is **37x closer to the true value than Excel** (3.5e-16 relative against
1.3e-14). Arbitrate with `mpmath`.

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

## 14. Empty-string cell vs. blank cell — *fixed in the comparator*

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

---

## 15. MOD with a divisor far larger than the dividend — *Excel is wrong*

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

## 16. VBA: an infinity poisons the next string-to-number conversion — *Excel is wrong*

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

## 17. VBA: `Err.Number` on a `Range` whose cells were deleted — *No stable answer*

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
