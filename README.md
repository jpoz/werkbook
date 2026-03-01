# Werkbook

A Go library for reading and writing Excel XLSX files. No external dependencies beyond the Go standard library.

```go
import "github.com/jpoz/werkbook"
```

## Excel Function Support

**147 of 500+ Excel functions supported** | **1,371 tests passing**

### Math & Trigonometry

| Function | Supported | Tests |
|----------|-----------|-------|
| ABS | :white_check_mark: | 5 |
| ACOS | :white_check_mark: | 3 |
| ACOSH | :white_check_mark: | 6 |
| ASIN | :white_check_mark: | 3 |
| ASINH | :white_check_mark: | 4 |
| ATAN | :white_check_mark: | 2 |
| ATAN2 | :white_check_mark: | 2 |
| ATANH | :white_check_mark: | 7 |
| CEILING | :white_check_mark: | 2 |
| COMBIN | :white_check_mark: | 6 |
| COS | :white_check_mark: | 1 |
| DEGREES | :white_check_mark: | 3 |
| EVEN | :white_check_mark: | 6 |
| EXP | :white_check_mark: | 2 |
| FACT | :white_check_mark: | 6 |
| FLOOR | :white_check_mark: | 2 |
| GCD | :white_check_mark: | 7 |
| INT | :white_check_mark: | 2 |
| LCM | :white_check_mark: | 6 |
| LN | :white_check_mark: | 3 |
| LOG | :white_check_mark: | 4 |
| LOG10 | :white_check_mark: | 1 |
| MOD | :white_check_mark: | 16 |
| MROUND | :white_check_mark: | 6 |
| ODD | :white_check_mark: | 6 |
| PI | :white_check_mark: | 2 |
| POWER | :white_check_mark: | 2 |
| PRODUCT | :white_check_mark: | 3 |
| QUOTIENT | :white_check_mark: | 5 |
| RADIANS | :white_check_mark: | 3 |
| RAND | :white_check_mark: | 3 |
| RANDBETWEEN | :white_check_mark: | 2 |
| ROUND | :white_check_mark: | 3 |
| ROUNDDOWN | :white_check_mark: | 2 |
| ROUNDUP | :white_check_mark: | 2 |
| SIGN | :white_check_mark: | 1 |
| SIN | :white_check_mark: | 1 |
| SQRT | :white_check_mark: | 3 |
| SUBTOTAL | :white_check_mark: | 10 |
| SUM | :white_check_mark: | 37 |
| SUMIF | :white_check_mark: | 9 |
| SUMIFS | :white_check_mark: | 5 |
| SUMPRODUCT | :white_check_mark: | 6 |
| SUMSQ | :white_check_mark: | 8 |
| TAN | :white_check_mark: | 1 |
| TRUNC | :white_check_mark: | 1 |

Not yet supported: ACOT, ACOTH, AGGREGATE, ARABIC, BASE, CEILING.MATH, CEILING.PRECISE, COMBINA, COSH, COT, COTH, CSC, CSCH, DECIMAL, FACTDOUBLE, FLOOR.MATH, FLOOR.PRECISE, ISO.CEILING, MDETERM, MINVERSE, MMULT, MULTINOMIAL, MUNIT, RANDARRAY, ROMAN, SEC, SECH, SEQUENCE, SERIESSUM, SINH, SQRTPI, SUMX2MY2, SUMX2PY2, SUMXMY2, TANH

### Statistical

| Function | Supported | Tests |
|----------|-----------|-------|
| AVERAGE | :white_check_mark: | 6 |
| AVERAGEIF | :white_check_mark: | 4 |
| AVERAGEIFS | :white_check_mark: | 4 |
| COUNT | :white_check_mark: | 2 |
| COUNTA | :white_check_mark: | 4 |
| COUNTBLANK | :white_check_mark: | 9 |
| COUNTIF | :white_check_mark: | 10 |
| COUNTIFS | :white_check_mark: | 3 |
| LARGE | :white_check_mark: | 4 |
| MAX | :white_check_mark: | 5 |
| MAXIFS | :white_check_mark: | 15 |
| MEDIAN | :white_check_mark: | 3 |
| MIN | :white_check_mark: | 3 |
| MINIFS | :white_check_mark: | 3 |
| MODE | :white_check_mark: | 7 |
| PERCENTILE | :white_check_mark: | 9 |
| PERMUT | :white_check_mark: | 13 |
| RANK | :white_check_mark: | 8 |
| SMALL | :white_check_mark: | 2 |
| STDEV | :white_check_mark: | 7 |
| STDEVP | :white_check_mark: | 6 |
| VAR | :white_check_mark: | 7 |
| VARP | :white_check_mark: | 6 |

Not yet supported: AVEDEV, AVERAGEA, BETA.DIST, BETA.INV, BINOM.DIST, BINOM.DIST.RANGE, BINOM.INV, CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT, CHISQ.TEST, CONFIDENCE.NORM, CONFIDENCE.T, CORREL, COVARIANCE.P, COVARIANCE.S, DEVSQ, EXPON.DIST, F.DIST, F.DIST.RT, F.INV, F.INV.RT, F.TEST, FISHER, FISHERINV, FORECAST, FORECAST.ETS, FORECAST.LINEAR, FREQUENCY, GAMMA, GAMMA.DIST, GAMMA.INV, GAMMALN, GAMMALN.PRECISE, GAUSS, GEOMEAN, GROWTH, HARMEAN, HYPGEOM.DIST, INTERCEPT, KURT, LINEST, LOGEST, LOGNORM.DIST, LOGNORM.INV, MAXA, MINA, MODE.MULT, MODE.SNGL, NEGBINOM.DIST, NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, PEARSON, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK.EXC, PERCENTRANK.INC, PERMUTATIONA, PHI, POISSON.DIST, PROB, QUARTILE.EXC, QUARTILE.INC, RANK.AVG, RANK.EQ, RSQ, SKEW, SKEW.P, SLOPE, STANDARDIZE, STDEV.P, STDEV.S, STDEVA, STDEVPA, STEYX, T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, T.TEST, TREND, TRIMMEAN, VAR.P, VAR.S, VARA, VARPA, WEIBULL.DIST, Z.TEST

### Logical

| Function | Supported | Tests |
|----------|-----------|-------|
| AND | :white_check_mark: | 8 |
| IF | :white_check_mark: | 36 |
| IFERROR | :white_check_mark: | 11 |
| IFNA | :white_check_mark: | 12 |
| IFS | :white_check_mark: | 13 |
| NOT | :white_check_mark: | 6 |
| OR | :white_check_mark: | 2 |
| SWITCH | :white_check_mark: | 14 |
| XOR | :white_check_mark: | 12 |

Not yet supported: BYCOL, BYROW, FALSE, LAMBDA, LET, MAKEARRAY, MAP, REDUCE, SCAN, TRUE

### Text

| Function | Supported | Tests |
|----------|-----------|-------|
| CHAR | :white_check_mark: | 8 |
| CLEAN | :white_check_mark: | 2 |
| CODE | :white_check_mark: | 4 |
| CONCAT | :white_check_mark: | 1 |
| CONCATENATE | :white_check_mark: | 7 |
| EXACT | :white_check_mark: | 6 |
| FIND | :white_check_mark: | 8 |
| FIXED | :white_check_mark: | 9 |
| LEFT | :white_check_mark: | 5 |
| LEN | :white_check_mark: | 3 |
| LOWER | :white_check_mark: | 1 |
| MID | :white_check_mark: | 6 |
| NUMBERVALUE | :white_check_mark: | 13 |
| PROPER | :white_check_mark: | 7 |
| REPLACE | :white_check_mark: | 6 |
| REPT | :white_check_mark: | 3 |
| RIGHT | :white_check_mark: | 5 |
| SEARCH | :white_check_mark: | 5 |
| SUBSTITUTE | :white_check_mark: | 14 |
| T | :white_check_mark: | 5 |
| TEXT | :white_check_mark: | 14 |
| TEXTJOIN | :white_check_mark: | 10 |
| TRIM | :white_check_mark: | 6 |
| UPPER | :white_check_mark: | 1 |
| VALUE | :white_check_mark: | 10 |

Not yet supported: ARRAYTOTEXT, ASC, BAHTTEXT, DBCS, DOLLAR, FINDB, LEFTB, LENB, MIDB, PHONETIC, REPLACEB, RIGHTB, SEARCHB, TEXTAFTER, TEXTBEFORE, TEXTSPLIT, UNICHAR, UNICODE, VALUETOTEXT

### Date & Time

| Function | Supported | Tests |
|----------|-----------|-------|
| DATE | :white_check_mark: | 15 |
| DATEDIF | :white_check_mark: | 9 |
| DATEVALUE | :white_check_mark: | 5 |
| DAY | :white_check_mark: | 2 |
| DAYS | :white_check_mark: | 9 |
| EDATE | :white_check_mark: | 5 |
| EOMONTH | :white_check_mark: | 5 |
| HOUR | :white_check_mark: | 1 |
| ISOWEEKNUM | :white_check_mark: | 7 |
| MINUTE | :white_check_mark: | 1 |
| MONTH | :white_check_mark: | 2 |
| NETWORKDAYS | :white_check_mark: | 5 |
| NOW | :white_check_mark: | 4 |
| SECOND | :white_check_mark: | 1 |
| TIME | :white_check_mark: | 4 |
| TODAY | :white_check_mark: | 4 |
| WEEKDAY | :white_check_mark: | 17 |
| WEEKNUM | :white_check_mark: | 6 |
| WORKDAY | :white_check_mark: | 301 |
| YEAR | :white_check_mark: | 2 |
| YEARFRAC | :white_check_mark: | 8 |

Not yet supported: DAYS360, NETWORKDAYS.INTL, TIMEVALUE, WORKDAY.INTL

### Lookup & Reference

| Function | Supported | Tests |
|----------|-----------|-------|
| ADDRESS | :white_check_mark: | 14 |
| CHOOSE | :white_check_mark: | 4 |
| COLUMN | :white_check_mark: | 5 |
| COLUMNS | :white_check_mark: | 4 |
| HLOOKUP | :white_check_mark: | 5 |
| INDEX | :white_check_mark: | 11 |
| LOOKUP | :white_check_mark: | 5 |
| MATCH | :white_check_mark: | 14 |
| ROW | :white_check_mark: | 7 |
| ROWS | :white_check_mark: | 4 |
| SORT | :white_check_mark: | 6 |
| VLOOKUP | :white_check_mark: | 19 |
| XLOOKUP | :white_check_mark: | 9 |

Not yet supported: AREAS, CHOOSECOLS, CHOOSEROWS, DROP, EXPAND, FILTER, FORMULATEXT, GETPIVOTDATA, GROUPBY, HSTACK, HYPERLINK, INDIRECT, OFFSET, PIVOTBY, RTD, SORTBY, TAKE, TOCOL, TOROW, TRANSPOSE, TRIMRANGE, UNIQUE, VSTACK, WRAPCOLS, WRAPROWS, XMATCH

### Information

| Function | Supported | Tests |
|----------|-----------|-------|
| ERROR.TYPE | :white_check_mark: | 9 |
| ISBLANK | :white_check_mark: | 2 |
| ISERR | :white_check_mark: | 7 |
| ISERROR | :white_check_mark: | 6 |
| ISEVEN | :white_check_mark: | 9 |
| ISLOGICAL | :white_check_mark: | 9 |
| ISNA | :white_check_mark: | 5 |
| ISNONTEXT | :white_check_mark: | 6 |
| ISNUMBER | :white_check_mark: | 9 |
| ISODD | :white_check_mark: | 8 |
| ISTEXT | :white_check_mark: | 2 |
| N | :white_check_mark: | 9 |
| NA | :white_check_mark: | 4 |
| TYPE | :white_check_mark: | 19 |

Not yet supported: CELL, INFO, ISFORMULA, ISOMITTED, ISREF, SHEET, SHEETS, STOCKHISTORY

### Financial (not yet supported)

ACCRINT, ACCRINTM, AMORDEGRC, AMORLINC, COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, CUMIPMT, CUMPRINC, DB, DDB, DISC, DOLLARDE, DOLLARFR, DURATION, EFFECT, FV, FVSCHEDULE, INTRATE, IPMT, IRR, ISPMT, MDURATION, MIRR, NOMINAL, NPER, NPV, ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD, PDURATION, PMT, PPMT, PRICE, PRICEDISC, PRICEMAT, PV, RATE, RECEIVED, RRI, SLN, SYD, TBILLEQ, TBILLPRICE, TBILLYIELD, VDB, XIRR, XNPV, YIELD, YIELDDISC, YIELDMAT

### Engineering (not yet supported)

BESSELI, BESSELJ, BESSELK, BESSELY, BIN2DEC, BIN2HEX, BIN2OCT, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, COMPLEX, CONVERT, DEC2BIN, DEC2HEX, DEC2OCT, DELTA, ERF, ERF.PRECISE, ERFC, ERFC.PRECISE, GESTEP, HEX2BIN, HEX2DEC, HEX2OCT, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, OCT2BIN, OCT2DEC, OCT2HEX

### Database (not yet supported)

DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP

### Web (not yet supported)

ENCODEURL, FILTERXML, WEBSERVICE

### Cube (not yet supported)

CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE
