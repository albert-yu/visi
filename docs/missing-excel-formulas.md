# Missing Excel Formulas in visi-core

A tracking list of standard Microsoft Excel functions (as documented in Excel
formula reference) and their implementation status in `visi-core`.

Last updated: 2026-08-31

---

## Summary

- Microsoft-documented functions: **522**
- Genuinely implemented in visi-core: **507**
- Missing or non-functional: **15**

## Caveats

- This is a name-level diff, not a semantics/argument-compatibility check.
  A handful of implemented functions likely don't cover every argument form
  Excel supports (array forms, optional args, etc.) — that's out of scope
  for this list.
- `SHEET()` is a known, documented approximation: it always returns `1`,
  since this engine has no notion of a sheet's true ordinal position within
  the workbook. Bare `COLUMN()`/`ROW()` (no argument, meaning "the current
  formula cell") are similarly not implemented — both need the engine to
  know which cell is currently being evaluated in a context where that
  isn't threaded through; the argument form (`COLUMN(A1)`, `ROW(A1:A5)`) is
  fully implemented.
- `TRANSPOSE` is genuinely implemented and hand-verified correct (see
  `visi-core/src/core/engine/tests/new_functions.rs`), but could not be
  confirmed against real Microsoft Excel via this repo's differential
  fuzzing harness: every authoring variant tried (bare, `_xlfn.`,
  `_xlfn._xlws.`, standalone, nested in `SUM`/`INDEX`) gave real Excel
  `#VALUE!` when the formula was written by `openpyxl` rather than Excel
  itself — `TRANSPOSE` predates dynamic arrays and has always required the
  legacy CSE (Ctrl+Shift+Enter) `t="array"` formula flag, which openpyxl's
  plain string assignment never produces. This is a test-authoring
  limitation of the harness, not a suspected implementation bug.
- `DETECTLANGUAGE`, `TRANSLATE`, and `PHONETIC` *are* genuinely implemented
  (real local logic in `core/text.rs`, not a placeholder) and stay counted
  here, even though the differential fuzz harness (`fuzz/fuzz_excel.py`)
  deliberately excludes them since real Excel's versions depend on
  Microsoft's cloud translation/language-detection services and aren't
  comparable to a local implementation.
- `GETPIVOTDATA` recomputes its pivot table's grid fresh on every
  evaluation (no caching) via `core/pivot.rs::getpivotdata`, and resolves
  its `pivot_table_ref` argument by scanning `Context.pivot_tables` for
  whichever pivot table's rendered destination range contains that cell —
  it does not take the pivot table by name. Row/column criteria that don't
  specify every field down to the innermost one match that branch's
  subtotal group (matching real Excel); an unresolvable field, item, or
  data-field name is `#REF!`/`#VALUE!`, also matching real Excel.
- `RTD` and `STOCKHISTORY` always return `#N/A` — this matches what real
  Excel itself shows once its equivalent live connection (a registered
  Windows COM `IRtdServer`, Microsoft's stock-data cloud service) is
  unavailable, rather than the misleading echo-the-last-argument
  placeholder these used to fall through to. Neither has a real local
  implementation; see "Recognized but not functional" below for why.

## Implemented (507)

ABS, ACCRINT, ACCRINTM, ACOS, ACOSH, ACOT, ACOTH, ADDRESS, AGGREGATE,
AMORDEGRC, AMORLINC, AND, ARABIC, AREAS, ARRAYTOTEXT, ASC, ASIN, ASINH,
ATAN, ATAN2, ATANH, AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS,
BAHTTEXT, BASE, BESSELI, BESSELJ, BESSELK, BESSELY, BETA.DIST, BETADIST,
BETA.INV, BETAINV, BIN2DEC, BIN2HEX, BIN2OCT, BINOM.DIST, BINOMDIST,
BINOM.DIST.RANGE, BINOM.INV, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR,
BYCOL, BYROW, CEILING, CEILING.MATH, CEILING.PRECISE, CELL, CHAR,
CHISQ.DIST, CHIDIST, CHISQ.DIST.RT, CHISQ.INV, CHIINV, CHISQ.INV.RT,
CHISQ.TEST, CHITEST, CHOOSE, CHOOSECOLS, CHOOSEROWS, CLEAN, CODE, COLUMN,
COLUMNS, COMBIN, COMBINA, COMPLEX, CONCAT, CONCATENATE, CONFIDENCE,
CONFIDENCE.NORM, CONFIDENCE.T, CONVERT, CORREL, COS, COSH, COT, COTH,
COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COUPDAYBS, COUPDAYS,
COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, COVAR, COVARIANCE.P, COVARIANCE.S,
CRITBINOM, CSC, CSCH, CUMIPMT, CUMPRINC, DATE, DATEDIF, DATEVALUE,
DAVERAGE, DAY, DAYS, DAYS360, DB, DBCS, DCOUNT, DCOUNTA, DEC2BIN, DEC2HEX,
DEC2OCT, DECIMAL, DDB, DEGREES, DELTA, DETECTLANGUAGE, DEVSQ, DGET, DISC,
DMAX, DMIN, DOLLAR, DOLLARDE, DOLLARFR, DPRODUCT, DROP, DSTDEV, DSTDEVP,
DSUM, DURATION, DVAR, DVARP, EDATE, EFFECT, ENCODEURL, EOMONTH, ERF,
ERF.PRECISE, ERFC, ERFC.PRECISE, ERROR.TYPE, EUROCONVERT, EXACT, EVEN, EXP,
EXPAND, EXPON.DIST, EXPONDIST, F.DIST, FDIST, F.DIST.RT, F.INV, FINV,
F.INV.RT, F.TEST, FTEST, FACT, FACTDOUBLE, FALSE, FILTER, FILTERXML, FIND,
FINDB, FISHER, FISHERINV, FIXED, FLOOR, FLOOR.MATH, FLOOR.PRECISE,
FORECAST, FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY,
FORECAST.ETS.STAT, FORECAST.LINEAR, FORMULATEXT, FREQUENCY, FV, FVSCHEDULE,
GAMMA, GAMMA.DIST, GAMMADIST, GAMMA.INV, GAMMAINV, GAMMALN,
GAMMALN.PRECISE, GAUSS, GCD, GEOMEAN, GESTEP, GETPIVOTDATA, GROWTH,
HARMEAN, HEX2BIN, HEX2DEC, HEX2OCT, HLOOKUP, HOUR, HSTACK, HYPERLINK,
HYPGEOM.DIST, HYPGEOMDIST, IF, IFERROR, IFNA, IFS, IMABS, IMAGINARY,
IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP,
IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN,
IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, INDEX, INDIRECT, INFO, INT, INTERCEPT,
INTRATE, IPMT, IRR, IS.CEILING, ISBLANK, ISERR, ISEVEN, ISERROR, ISFORMULA,
ISLOGICAL, ISOMITTED, ISNA, ISNONTEXT, ISNUMBER, ISO.CEILING, ISODD, ISPMT,
ISREF, ISOWEEKNUM, ISTEXT, JIS, KURT, LAMBDA, LARGE, LCM, LEFT, LEFTB, LEN,
LENB, LET, LINEST, LN, LOG, LOG10, LOGEST, LOGINV, LOGNORM.DIST,
LOGNORMDIST, LOGNORM.INV, LOOKUP, LOWER, MAKEARRAY, MAP, MATCH, MAX, MAXA,
MAXIFS, MDETERM, MDURATION, MEDIAN, MID, MIDB, MIN, MINA, MINIFS, MINVERSE,
MIRR, MINUTE, MMULT, MOD, MODE, MODE.MULT, MODE.SNGL, MONTH, MROUND,
MULTINOMIAL, MUNIT, N, NA, NEGBINOM.DIST, NEGBINOMDIST, NETWORKDAYS,
NETWORKDAYS.INTL, NOMINAL, NORM.DIST, NORMDIST, NORM.INV, NORMINV,
NORM.S.DIST, NORMSDIST, NORM.S.INV, NORMSINV, NOT, NOW, NPER, NPV,
NUMBERVALUE, OCT2BIN, OCT2DEC, OCT2HEX, ODD, ODDFPRICE, ODDFYIELD,
ODDLPRICE, ODDLYIELD, OFFSET, OR, PEARSON, PDURATION, PERCENTILE,
PERCENTILE.EXC, PERCENTILE.INC, PERCENTOF, PERCENTRANK, PERCENTRANK.EXC,
PERCENTRANK.INC, PERMUT, PERMUTATIONA, PHI, PHONETIC, PI, PMT, POISSON,
POISSON.DIST, POWER, PPMT, PRICE, PRICEDISC, PRICEMAT, PROB, PRODUCT,
PROPER, PV, QUARTILE, QUOTIENT, QUARTILE.EXC, QUARTILE.INC, RADIANS, RAND,
RANDARRAY, RANDBETWEEN, RANK, RANK.AVG, RANK.EQ, RATE, RECEIVED, REDUCE,
REGEXEXTRACT, REGEXREPLACE, REGEXTEST, REPLACE, REPLACEB, REPT, RIGHT,
RIGHTB, ROMAN, ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, RRI, RSQ, SCAN,
SEARCH, SEARCHB, SEC, SECOND, SECH, SEQUENCE, SERIESSUM, SHEET, SHEETS,
SIGN, SIN, SINH, SKEW, SKEW.P, SLN, SLOPE, SMALL, SORT, SORTBY, SQRT,
SQRTPI, STANDARDIZE, STDEV, STDEV.P, STDEV.S, STDEVA, STDEVP, STDEVPA,
STEYX, SUBSTITUTE, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ,
SUMX2MY2, SUMX2PY2, SUMXMY2, SWITCH, SYD, T, TAKE, T.DIST, TDIST,
T.DIST.2T, T.DIST.RT, T.INV, TINV, T.INV.2T, T.TEST, TTEST, TAN, TANH,
TBILLEQ, TBILLPRICE, TBILLYIELD, TEXT, TEXTAFTER, TEXTBEFORE, TEXTJOIN,
TEXTSPLIT, TIME, TIMEVALUE, TOCOL, TODAY, TOROW, TRANSLATE, TRANSPOSE,
TREND, TRIMMEAN, TRIM, TRIMRANGE, TRUNC, TRUE, TYPE, UNICHAR, UNICODE,
UNIQUE, UPPER, VALUE, VALUETOTEXT, VAR, VAR.P, VAR.S, VARA, VARP, VARPA,
VDB, VLOOKUP, VSTACK, WEIBULL, WEIBULL.DIST, WEEKDAY, WEEKNUM, WORKDAY,
WORKDAY.INTL, WRAPCOLS, WRAPROWS, XIRR, XLOOKUP, XMATCH, XNPV, XOR, YEAR,
YEARFRAC, YIELD, YIELDDISC, YIELDMAT, Z.TEST, ZTEST

## Missing, by category

### Financial (0/28 missing)

All previously-missing financial day-count/bond-pricing functions are now
implemented (`ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`,
`COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `DISC`,
`DURATION`, `INTRATE`, `MDURATION`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`,
`ODDLYIELD`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `RECEIVED`, `TBILLEQ`,
`TBILLPRICE`, `TBILLYIELD`, `YIELD`, `YIELDDISC`, `YIELDMAT`) — see
`visi-core/src/core/finance.rs`.

### Add-in / legacy (2 — intentionally unsupported, not missing work)

CALL, REGISTER.ID — neither is dispatched at all, so calling either
already falls through to the generic "unknown function" error (`#NAME?`),
which is exactly what real Excel does when the XLL/DLL add-in that defines
them isn't loaded. These invoke arbitrary native add-in code; there is no
meaningful local implementation, and current behavior already matches
Excel's own (add-in-less) behavior. (`EUROCONVERT` is implemented — see
`core/finance.rs`.)

### Recognized but not functional (13 — parsed and dispatched, but the
implementation is a stub that just echoes an argument back rather than
computing the real result, or reports the data source as unavailable)

- **Cube (7)**: CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY,
  CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE — in real Excel these
  all query a live OLAP (Analysis Services) connection using MDX tuple/set
  expressions (e.g. `[Product].[Category].[Bikes]`); visi-core has no OLAP
  connection concept and no MDX parser, so there's no local data these
  formulas could compute a real result from. Genuinely implementing them
  would mean inventing visi-core-specific semantics (e.g. treating the
  `connection` argument as an `ExcelTable` name and member expressions as
  column/value references) rather than matching real Excel/SSAS behavior —
  a scope decision intentionally left unmade rather than guessed at. Left
  as stubs; not attempted.
- **Web service (1)**: WEBSERVICE — needs a blocking HTTP client, which
  conflicts with `visi-core/core`'s "no IO assumptions, wasm-targetable"
  rule; would need a non-default Cargo feature gating an HTTP dependency
  rather than an unconditional one. Not attempted. (`FILTERXML` is now
  implemented locally — see `core/xml.rs` — since XPath-over-a-literal-
  string needs no network access.)
- **External/live data (2)**: RTD, STOCKHISTORY — `RTD` needs a registered
  Windows COM `IRtdServer`, which is the one genuinely Windows-COM-only
  case in this whole list; `STOCKHISTORY` is backed by Microsoft's
  internal "Data Types" cloud service with no public, unauthenticated REST
  API. Neither is implementable without either a Windows-only COM client
  or reverse-engineering an undocumented, likely-authenticated Microsoft
  service. Both now return `#N/A` (matching real Excel's own display once
  its equivalent live connection is unavailable) rather than echoing an
  argument back.
- **Pivot / grouping / image (3)**: GROUPBY, PIVOTBY, IMAGE — `GROUPBY` and
  `PIVOTBY` take raw array arguments rather than an existing `PivotTable`
  object, so they'd need new array-level grouping/aggregation logic (unlike
  `GETPIVOTDATA`, now implemented — see below); `IMAGE` needs a new
  `IMAGE` needs a new `ResultData` variant plus real xlsx image-embedding support.

`GETPIVOTDATA` is now genuinely implemented
(`core/pivot.rs::getpivotdata`, dispatched via
`Sheet::evaluate_getpivotdata`) — see the Caveats section above.

Everything else previously listed here — all 12 `D*` database functions,
all of `AREAS`, `CHOOSECOLS`, `CHOOSEROWS`, `COLUMN`, `COLUMNS`, `DROP`,
`EXPAND`, `FILTER`, `FORMULATEXT`, `HSTACK`, `HYPERLINK`, `INDIRECT`,
`LOOKUP`, `OFFSET`, `ROW`, `ROWS`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`,
`TOROW`, `TRANSPOSE`, `TRIMRANGE`, `UNIQUE`, `VSTACK`, `WRAPCOLS`,
`WRAPROWS`, `XMATCH`, `GETPIVOTDATA`, and all of `CELL`, `INFO`,
`ISFORMULA`, `ISOMITTED`, `ISREF`, `SHEETS`, `BYCOL`, `BYROW`, `LAMBDA`,
`MAKEARRAY`, `MAP`, `REDUCE`, `SCAN` — is now genuinely implemented (see
the "Caveats" section above for `SHEET`'s approximation and `TRANSPOSE`'s
fuzz-harness limitation, both counted as implemented here).

### Not dispatched at all (0)

`CHOOSE` was referenced in one internal arg-classification check but had no
match arm, so calling it fell through to the generic "unknown function"
error; it's now wired up and fully implemented.
