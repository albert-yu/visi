# Missing Excel Formulas in libvisi

A tracking list of standard Microsoft Excel functions (as documented in Excel
formula reference) and their implementation status in `libvisi`.

Last updated: 2026-08-03

---

## Summary

- Microsoft-documented functions: **522**
- Genuinely implemented in libvisi: **382**
- Missing or non-functional: **140**

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
  or aren't dispatched at all (`CHOOSE`). Both classes have been moved out
  of "Implemented" in this revision. `DETECTLANGUAGE`, `TRANSLATE`, and
  `PHONETIC` *are* genuinely implemented (real local logic in `core/text.rs`,
  not a placeholder) and stay counted here, even though the differential
  fuzz harness (`fuzz/fuzz_excel.py`) deliberately excludes them since real
  Excel's versions depend on Microsoft's cloud translation/language-detection
  services and aren't comparable to a local implementation. `LET` was one of
  the 70 placeholder functions at the time of that revision, but has since
  gained real variable-binding support (`LetScope` in
  `libvisi/src/core/engine/sheet.rs`) and moved back to "Implemented".
- The 140 "missing or non-functional" total is derived as 522 minus the
  verified 382 implemented; only 100 of those 140 are individually named
  below (28 financial + 3 add-in/legacy + 68 recognized-but-non-functional
  + `CHOOSE` = 100). The remaining ~40 were already unaccounted for in the
  prior revision's own numbers (it claimed 70 missing but only ever named
  31) — that gap predates this revision and wasn't independently
  re-derived here for lack of a master function list to diff against.

## Implemented (382)

ABS, ACOS, ACOSH, ACOT, ACOTH, ADDRESS, AGGREGATE, AND, ARABIC, ARRAYTOTEXT, ASC, ASIN, ASINH, ATAN, ATAN2, ATANH, AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS, BAHTTEXT, BASE, BESSELI, BESSELJ, BESSELK, BESSELY, BETA.DIST, BETA.INV, BIN2DEC, BIN2HEX, BIN2OCT, BINOM.DIST, BINOM.DIST.RANGE, BINOM.INV, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, CEILING, CEILING.MATH, CEILING.PRECISE, CHAR, CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT, CHISQ.TEST, CLEAN, CODE, COMBIN, COMBINA, COMPLEX, CONCAT, CONCATENATE, CONFIDENCE.NORM, CONFIDENCE.T, CONVERT, CORREL, COS, COSH, COT, COTH, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COVARIANCE.P, COVARIANCE.S, CSC, CSCH, CUMIPMT, CUMPRINC, DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, DB, DBCS, DEC2BIN, DEC2HEX, DEC2OCT, DECIMAL, DDB, DEGREES, DELTA, DETECTLANGUAGE, DEVSQ, DOLLAR, DOLLARDE, DOLLARFR, EDATE, EFFECT, ENCODEURL, EOMONTH, ERF, ERF.PRECISE, ERFC, ERFC.PRECISE, ERROR.TYPE, EXACT, EVEN, EXP, EXPON.DIST, F.DIST, F.DIST.RT, F.INV, F.INV.RT, F.TEST, FACT, FACTDOUBLE, FALSE, FIND, FINDB, FISHER, FISHERINV, FIXED, FLOOR, FLOOR.MATH, FLOOR.PRECISE, FORECAST, FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT, FORECAST.LINEAR, FREQUENCY, FV, FVSCHEDULE, GAMMA, GAMMA.DIST, GAMMA.INV, GAMMALN, GAMMALN.PRECISE, GAUSS, GCD, GEOMEAN, GESTEP, GROWTH, HARMEAN, HEX2BIN, HEX2DEC, HEX2OCT, HLOOKUP, HOUR, HYPGEOM.DIST, IF, IFERROR, IFNA, IFS, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, INDEX, INT, INTERCEPT, IPMT, IRR, IS.CEILING, ISBLANK, ISERR, ISEVEN, ISERROR, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISO.CEILING, ISODD, ISPMT, ISOWEEKNUM, ISTEXT, KURT, LARGE, LCM, LEFT, LEFTB, LEN, LENB, LET, LINEST, LN, LOG, LOG10, LOGEST, LOGNORM.DIST, LOGNORM.INV, LOWER, MATCH, MAX, MAXA, MAXIFS, MDETERM, MEDIAN, MID, MIDB, MIN, MINA, MINIFS, MINVERSE, MIRR, MINUTE, MMULT, MOD, MODE.MULT, MODE.SNGL, MONTH, MROUND, MULTINOMIAL, MUNIT, N, NA, NEGBINOM.DIST, NETWORKDAYS, NETWORKDAYS.INTL, NOMINAL, NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, NOT, NOW, NPER, NPV, NUMBERVALUE, OCT2BIN, OCT2DEC, OCT2HEX, ODD, OR, PEARSON, PDURATION, PERCENTILE.EXC, PERCENTILE.INC, PERCENTOF, PERCENTRANK.EXC, PERCENTRANK.INC, PERMUT, PERMUTATIONA, PHI, PHONETIC, PI, PMT, POISSON.DIST, POWER, PPMT, PROB, PRODUCT, PROPER, PV, QUOTIENT, QUARTILE.EXC, QUARTILE.INC, RADIANS, RAND, RANDARRAY, RANDBETWEEN, RANK.AVG, RANK.EQ, RATE, REGEXEXTRACT, REGEXREPLACE, REGEXTEST, REPLACE, REPLACEB, REPT, RIGHT, RIGHTB, ROMAN, ROUND, ROUNDDOWN, ROUNDUP, RRI, RSQ, SEARCH, SEARCHB, SEC, SECOND, SECH, SEQUENCE, SERIESSUM, SIGN, SIN, SINH, SKEW, SKEW.P, SLN, SLOPE, SMALL, SQRT, SQRTPI, STANDARDIZE, STDEV.P, STDEV.S, STDEVA, STDEVPA, STEYX, SUBSTITUTE, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ, SUMX2MY2, SUMX2PY2, SUMXMY2, SWITCH, SYD, T, T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, T.TEST, TAN, TANH, TEXT, TEXTAFTER, TEXTBEFORE, TEXTJOIN, TEXTSPLIT, TIME, TIMEVALUE, TODAY, TRANSLATE, TREND, TRIMMEAN, TRIM, TRUNC, TRUE, TYPE, UNICHAR, UNICODE, UPPER, VALUE, VALUETOTEXT, VAR.P, VAR.S, VARA, VARPA, VDB, VLOOKUP, WEIBULL.DIST, WEEKDAY, WEEKNUM, WORKDAY, WORKDAY.INTL, XIRR, XLOOKUP, XNPV, XOR, YEAR, YEARFRAC, Z.TEST

## Missing, by category

### Financial (28/55 missing — the day-count/bond-pricing half)

ACCRINT, ACCRINTM, AMORDEGRC, AMORLINC, COUPDAYBS, COUPDAYS, COUPDAYSNC,
COUPNCD, COUPNUM, COUPPCD, DISC, DURATION, INTRATE, MDURATION, ODDFPRICE,
ODDFYIELD, ODDLPRICE, ODDLYIELD, PRICE, PRICEDISC, PRICEMAT, RECEIVED,
TBILLEQ, TBILLPRICE, TBILLYIELD, YIELD, YIELDDISC, YIELDMAT

### Add-in / legacy (3/3 missing)

CALL, EUROCONVERT, REGISTER.ID

### Recognized but not functional (68 — parsed and dispatched, but the
implementation is a stub that just echoes an argument back rather than
computing the real result)

- **Database (12)**: DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT,
  DSTDEV, DSTDEVP, DSUM, DVAR, DVARP
- **Cube / web service (9)**: CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY,
  CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE, FILTERXML, WEBSERVICE
- **Dynamic arrays, lookup, and range introspection (32)**: AREAS,
  CHOOSECOLS, CHOOSEROWS, COLUMN, COLUMNS, DROP, EXPAND, FILTER,
  FORMULATEXT, GETPIVOTDATA, GROUPBY, HSTACK, HYPERLINK, IMAGE, INDIRECT,
  LOOKUP, OFFSET, PIVOTBY, ROW, ROWS, RTD, SORT, SORTBY, TAKE, TOCOL, TOROW,
  TRANSPOSE, TRIMRANGE, UNIQUE, VSTACK, WRAPCOLS, WRAPROWS, XMATCH
- **LAMBDA family / workbook metadata (15)**: CELL, INFO, ISFORMULA,
  ISOMITTED, ISREF, SHEET, SHEETS, STOCKHISTORY, BYCOL, BYROW, LAMBDA,
  MAKEARRAY, MAP, REDUCE, SCAN. `LAMBDA` needs the same parameter-binding
  mechanism `LET` now uses (see `LetScope`) but hasn't been wired up to use
  it yet.

### Not dispatched at all (1)

CHOOSE — referenced in one internal arg-classification check but has no
match arm; calling it falls through to the generic "unknown function" error.
