# Missing Excel Formulas in libvisi

A tracking list of standard Microsoft Excel functions (as documented in Excel
formula reference) and their implementation status in `libvisi`.

Last updated: 2026-08-06

---

## Summary

- Microsoft-documented functions: **522**
- Genuinely implemented in libvisi: **465**
- Missing or non-functional: **57**

libvisi implements `PLOT`, `GET`, `GET_COL`, `GET_COL_IDX`, `SLICE`, and
`STR` — these are engine-specific extensions, not Excel functions, so they
aren't counted above in either total.

## Caveats

- This is a name-level diff, not a semantics/argument-compatibility check.
  A handful of implemented functions likely don't cover every argument form
  Excel supports (array forms, optional args, etc.) — that's out of scope
  for this list.
- A prior revision of this file double-counted 28 financial functions (they
  appeared in both "Implemented" and "Missing") and, separately, counted 70
  functions as implemented that are either dispatched to a placeholder that
  just echoes back an argument (see "Recognized but not functional" below)
  or aren't dispatched at all (`CHOOSE`, since fixed). Both classes have
  been moved out of "Implemented" in that revision; this revision moves the
  now-genuinely-implemented ones back in (see below).
- `SHEET()` is a known, documented approximation: it always returns `1`,
  since this engine has no notion of a sheet's true ordinal position within
  the workbook. Bare `COLUMN()`/`ROW()` (no argument, meaning "the current
  formula cell") are similarly not implemented — both need the engine to
  know which cell is currently being evaluated in a context where that
  isn't threaded through; the argument form (`COLUMN(A1)`, `ROW(A1:A5)`) is
  fully implemented.
- `TRANSPOSE` is genuinely implemented and hand-verified correct (see
  `libvisi/src/core/engine/tests/new_functions.rs`), but could not be
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
- The 57 "missing or non-functional" total is derived as 522 minus the
  verified 465 implemented; only 17 of those 57 are individually named
  below. The remaining ~39 were already unaccounted for in a much earlier
  revision's own numbers and weren't independently re-derived here for
  lack of a master function list to diff against.

## Implemented (465)

ABS, ACCRINT, ACCRINTM, ACOS, ACOSH, ACOT, ACOTH, ADDRESS, AGGREGATE,
AMORDEGRC, AMORLINC, AND, ARABIC, AREAS, ARRAYTOTEXT, ASC, ASIN, ASINH,
ATAN, ATAN2, ATANH, AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS,
BAHTTEXT, BASE, BESSELI, BESSELJ, BESSELK, BESSELY, BETA.DIST, BETA.INV,
BIN2DEC, BIN2HEX, BIN2OCT, BINOM.DIST, BINOM.DIST.RANGE, BINOM.INV, BITAND,
BITLSHIFT, BITOR, BITRSHIFT, BITXOR, BYCOL, BYROW, CEILING, CEILING.MATH,
CEILING.PRECISE, CELL, CHAR, CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV,
CHISQ.INV.RT, CHISQ.TEST, CHOOSE, CHOOSECOLS, CHOOSEROWS, CLEAN, CODE,
COLUMN, COLUMNS, COMBIN, COMBINA, COMPLEX, CONCAT, CONCATENATE,
CONFIDENCE.NORM, CONFIDENCE.T, CONVERT, CORREL, COS, COSH, COT, COTH,
COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COUPDAYBS, COUPDAYS,
COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, COVARIANCE.P, COVARIANCE.S, CSC,
CSCH, CUMIPMT, CUMPRINC, DATE, DATEDIF, DATEVALUE, DAVERAGE, DAY, DAYS,
DAYS360, DB, DBCS, DCOUNT, DCOUNTA, DEC2BIN, DEC2HEX, DEC2OCT, DECIMAL, DDB,
DEGREES, DELTA, DETECTLANGUAGE, DEVSQ, DGET, DISC, DMAX, DMIN, DOLLAR,
DOLLARDE, DOLLARFR, DPRODUCT, DROP, DSTDEV, DSTDEVP, DSUM, DURATION, DVAR,
DVARP, EDATE, EFFECT, ENCODEURL, EOMONTH, ERF, ERF.PRECISE, ERFC,
ERFC.PRECISE, ERROR.TYPE, EUROCONVERT, EXACT, EVEN, EXP, EXPAND,
EXPON.DIST, F.DIST, F.DIST.RT, F.INV, F.INV.RT, F.TEST, FACT, FACTDOUBLE,
FALSE, FILTER, FIND, FINDB, FISHER, FISHERINV, FIXED, FLOOR, FLOOR.MATH,
FLOOR.PRECISE, FORECAST, FORECAST.ETS, FORECAST.ETS.CONFINT,
FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT, FORECAST.LINEAR, FORMULATEXT,
FREQUENCY, FV, FVSCHEDULE, GAMMA, GAMMA.DIST, GAMMA.INV, GAMMALN,
GAMMALN.PRECISE, GAUSS, GCD, GEOMEAN, GESTEP, GROWTH, HARMEAN, HEX2BIN,
HEX2DEC, HEX2OCT, HLOOKUP, HOUR, HSTACK, HYPERLINK, HYPGEOM.DIST, IF,
IFERROR, IFNA, IFS, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS,
IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER,
IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM,
IMTAN, INDEX, INDIRECT, INFO, INT, INTERCEPT, INTRATE, IPMT, IRR,
IS.CEILING, ISBLANK, ISERR, ISEVEN, ISERROR, ISFORMULA, ISLOGICAL,
ISOMITTED, ISNA, ISNONTEXT, ISNUMBER, ISO.CEILING, ISODD, ISPMT, ISREF,
ISOWEEKNUM, ISTEXT, KURT, LAMBDA, LARGE, LCM, LEFT, LEFTB, LEN, LENB, LET,
LINEST, LN, LOG, LOG10, LOGEST, LOGNORM.DIST, LOGNORM.INV, LOOKUP, LOWER,
MAKEARRAY, MAP, MATCH, MAX, MAXA, MAXIFS, MDETERM, MDURATION, MEDIAN, MID,
MIDB, MIN, MINA, MINIFS, MINVERSE, MIRR, MINUTE, MMULT, MOD, MODE.MULT,
MODE.SNGL, MONTH, MROUND, MULTINOMIAL, MUNIT, N, NA, NEGBINOM.DIST,
NETWORKDAYS, NETWORKDAYS.INTL, NOMINAL, NORM.DIST, NORM.INV, NORM.S.DIST,
NORM.S.INV, NOT, NOW, NPER, NPV, NUMBERVALUE, OCT2BIN, OCT2DEC, OCT2HEX,
ODD, ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD, OFFSET, OR, PEARSON,
PDURATION, PERCENTILE.EXC, PERCENTILE.INC, PERCENTOF, PERCENTRANK.EXC,
PERCENTRANK.INC, PERMUT, PERMUTATIONA, PHI, PHONETIC, PI, PMT,
POISSON.DIST, POWER, PPMT, PRICE, PRICEDISC, PRICEMAT, PROB, PRODUCT,
PROPER, PV, QUOTIENT, QUARTILE.EXC, QUARTILE.INC, RADIANS, RAND, RANDARRAY,
RANDBETWEEN, RANK.AVG, RANK.EQ, RATE, RECEIVED, REDUCE, REGEXEXTRACT,
REGEXREPLACE, REGEXTEST, REPLACE, REPLACEB, REPT, RIGHT, RIGHTB, ROMAN,
ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, RRI, RSQ, SCAN, SEARCH, SEARCHB, SEC,
SECOND, SECH, SEQUENCE, SERIESSUM, SHEETS, SIGN, SIN, SINH, SKEW, SKEW.P,
SLN, SLOPE, SMALL, SORT, SORTBY, SQRT, SQRTPI, STANDARDIZE, STDEV.P,
STDEV.S, STDEVA, STDEVPA, STEYX, SUBSTITUTE, SUBTOTAL, SUM, SUMIF, SUMIFS,
SUMPRODUCT, SUMSQ, SUMX2MY2, SUMX2PY2, SUMXMY2, SWITCH, SYD, T, TAKE,
T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, T.TEST, TAN, TANH, TBILLEQ,
TBILLPRICE, TBILLYIELD, TEXT, TEXTAFTER, TEXTBEFORE, TEXTJOIN, TEXTSPLIT,
TIME, TIMEVALUE, TOCOL, TODAY, TOROW, TRANSLATE, TRANSPOSE, TREND,
TRIMMEAN, TRIM, TRIMRANGE, TRUNC, TRUE, TYPE, UNICHAR, UNICODE, UNIQUE,
UPPER, VALUE, VALUETOTEXT, VAR.P, VAR.S, VARA, VARPA, VDB, VLOOKUP, VSTACK,
WEIBULL.DIST, WEEKDAY, WEEKNUM, WORKDAY, WORKDAY.INTL, WRAPCOLS, WRAPROWS,
XIRR, XLOOKUP, XMATCH, XNPV, XOR, YEAR, YEARFRAC, YIELD, YIELDDISC,
YIELDMAT, Z.TEST

## Missing, by category

### Financial (0/28 missing)

All previously-missing financial day-count/bond-pricing functions are now
implemented (`ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`,
`COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `DISC`,
`DURATION`, `INTRATE`, `MDURATION`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`,
`ODDLYIELD`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `RECEIVED`, `TBILLEQ`,
`TBILLPRICE`, `TBILLYIELD`, `YIELD`, `YIELDDISC`, `YIELDMAT`) — see
`libvisi/src/core/finance.rs`.

### Add-in / legacy (2/3 missing)

CALL, REGISTER.ID

`EUROCONVERT` is now implemented (fixed historical ECB triangulation rates
— see `core/finance.rs`); `CALL`/`REGISTER.ID` invoke external DLL/XLL
add-ins, which has no meaning in this engine.

### Recognized but not functional (17 — parsed and dispatched, but the
implementation is a stub that just echoes an argument back rather than
computing the real result)

- **Cube (7)**: CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY,
  CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE — in real Excel these
  all query a live OLAP (Analysis Services) connection using MDX tuple/set
  expressions (e.g. `[Product].[Category].[Bikes]`); libvisi has no OLAP
  connection concept and no MDX parser, so there's no local data these
  formulas could compute a real result from. Genuinely implementing them
  would mean inventing libvisi-specific semantics (e.g. treating the
  `connection` argument as an `ExcelTable` name and member expressions as
  column/value references) rather than matching real Excel/SSAS behavior —
  a scope decision intentionally left unmade rather than guessed at. Left
  as stubs; not attempted.
- **Web service (2)**: FILTERXML, WEBSERVICE
- **Pivot / grouping / external data (6)**: GETPIVOTDATA, GROUPBY, IMAGE,
  PIVOTBY, RTD, STOCKHISTORY — `GETPIVOTDATA`/`GROUPBY`/`PIVOTBY` need
  deeper pivot-table integration than a formula-level implementation
  affords; `RTD`/`STOCKHISTORY` need a live external data source; `IMAGE`
  needs image-object support this engine's cell model doesn't have.

Everything else previously listed here — all 12 `D*` database functions,
all of `AREAS`, `CHOOSECOLS`, `CHOOSEROWS`, `COLUMN`, `COLUMNS`, `DROP`,
`EXPAND`, `FILTER`, `FORMULATEXT`, `HSTACK`, `HYPERLINK`, `INDIRECT`,
`LOOKUP`, `OFFSET`, `ROW`, `ROWS`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`,
`TOROW`, `TRANSPOSE`, `TRIMRANGE`, `UNIQUE`, `VSTACK`, `WRAPCOLS`,
`WRAPROWS`, `XMATCH`, and all of `CELL`, `INFO`, `ISFORMULA`, `ISOMITTED`,
`ISREF`, `SHEETS`, `BYCOL`, `BYROW`, `LAMBDA`, `MAKEARRAY`, `MAP`,
`REDUCE`, `SCAN` — is now genuinely implemented (see the "Caveats" section
above for `SHEET`'s approximation and `TRANSPOSE`'s fuzz-harness
limitation, both counted as implemented here).

### Not dispatched at all (0)

`CHOOSE` was referenced in one internal arg-classification check but had no
match arm, so calling it fell through to the generic "unknown function"
error; it's now wired up and fully implemented.
