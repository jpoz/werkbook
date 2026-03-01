# Werkbook

A Go library for reading and writing Excel XLSX files. No external dependencies beyond the Go standard library.

```go
import "github.com/jpoz/werkbook"
```

## Excel Function Support

**144 of 500+ Excel functions supported** | **1,163 tests passing**

### Math & Trigonometry

| Function | Supported | Tests |
|----------|-----------|-------|
| ABS | :white_check_mark: | 5 |
| ACOS | :white_check_mark: | 3 |
| ACOSH | :x: | - |
| ACOT | :x: | - |
| ACOTH | :x: | - |
| AGGREGATE | :x: | - |
| ARABIC | :x: | - |
| ASIN | :white_check_mark: | 3 |
| ASINH | :x: | - |
| ATAN | :white_check_mark: | 2 |
| ATAN2 | :white_check_mark: | 2 |
| ATANH | :x: | - |
| BASE | :x: | - |
| CEILING | :white_check_mark: | 2 |
| CEILING.MATH | :x: | - |
| CEILING.PRECISE | :x: | - |
| COMBIN | :white_check_mark: | 6 |
| COMBINA | :x: | - |
| COS | :white_check_mark: | 1 |
| COSH | :x: | - |
| COT | :x: | - |
| COTH | :x: | - |
| CSC | :x: | - |
| CSCH | :x: | - |
| DECIMAL | :x: | - |
| DEGREES | :white_check_mark: | 3 |
| EVEN | :white_check_mark: | 6 |
| EXP | :white_check_mark: | 2 |
| FACT | :white_check_mark: | 6 |
| FACTDOUBLE | :x: | - |
| FLOOR | :white_check_mark: | 2 |
| FLOOR.MATH | :x: | - |
| FLOOR.PRECISE | :x: | - |
| GCD | :white_check_mark: | 7 |
| INT | :white_check_mark: | 2 |
| ISO.CEILING | :x: | - |
| LCM | :white_check_mark: | 6 |
| LN | :white_check_mark: | 3 |
| LOG | :white_check_mark: | 4 |
| LOG10 | :white_check_mark: | 1 |
| MDETERM | :x: | - |
| MINVERSE | :x: | - |
| MMULT | :x: | - |
| MOD | :white_check_mark: | 16 |
| MROUND | :white_check_mark: | 6 |
| MULTINOMIAL | :x: | - |
| MUNIT | :x: | - |
| ODD | :white_check_mark: | 6 |
| PI | :white_check_mark: | 2 |
| POWER | :white_check_mark: | 2 |
| PRODUCT | :white_check_mark: | 3 |
| QUOTIENT | :white_check_mark: | 5 |
| RADIANS | :white_check_mark: | 3 |
| RAND | :white_check_mark: | 3 |
| RANDARRAY | :x: | - |
| RANDBETWEEN | :white_check_mark: | 2 |
| ROMAN | :x: | - |
| ROUND | :white_check_mark: | 3 |
| ROUNDDOWN | :white_check_mark: | 2 |
| ROUNDUP | :white_check_mark: | 2 |
| SEC | :x: | - |
| SECH | :x: | - |
| SERIESSUM | :x: | - |
| SEQUENCE | :x: | - |
| SIGN | :white_check_mark: | 1 |
| SIN | :white_check_mark: | 1 |
| SINH | :x: | - |
| SQRT | :white_check_mark: | 3 |
| SQRTPI | :x: | - |
| SUBTOTAL | :white_check_mark: | 10 |
| SUM | :white_check_mark: | 37 |
| SUMIF | :white_check_mark: | 9 |
| SUMIFS | :white_check_mark: | 5 |
| SUMPRODUCT | :white_check_mark: | 6 |
| SUMSQ | :white_check_mark: | 8 |
| SUMX2MY2 | :x: | - |
| SUMX2PY2 | :x: | - |
| SUMXMY2 | :x: | - |
| TAN | :white_check_mark: | 1 |
| TANH | :x: | - |
| TRUNC | :white_check_mark: | 1 |

### Statistical

| Function | Supported | Tests |
|----------|-----------|-------|
| AVEDEV | :x: | - |
| AVERAGE | :white_check_mark: | 6 |
| AVERAGEA | :x: | - |
| AVERAGEIF | :white_check_mark: | 4 |
| AVERAGEIFS | :white_check_mark: | 4 |
| BETA.DIST | :x: | - |
| BETA.INV | :x: | - |
| BINOM.DIST | :x: | - |
| BINOM.DIST.RANGE | :x: | - |
| BINOM.INV | :x: | - |
| CHISQ.DIST | :x: | - |
| CHISQ.DIST.RT | :x: | - |
| CHISQ.INV | :x: | - |
| CHISQ.INV.RT | :x: | - |
| CHISQ.TEST | :x: | - |
| CONFIDENCE.NORM | :x: | - |
| CONFIDENCE.T | :x: | - |
| CORREL | :x: | - |
| COUNT | :white_check_mark: | 2 |
| COUNTA | :white_check_mark: | 4 |
| COUNTBLANK | :white_check_mark: | 9 |
| COUNTIF | :white_check_mark: | 10 |
| COUNTIFS | :white_check_mark: | 3 |
| COVARIANCE.P | :x: | - |
| COVARIANCE.S | :x: | - |
| DEVSQ | :x: | - |
| EXPON.DIST | :x: | - |
| F.DIST | :x: | - |
| F.DIST.RT | :x: | - |
| F.INV | :x: | - |
| F.INV.RT | :x: | - |
| F.TEST | :x: | - |
| FISHER | :x: | - |
| FISHERINV | :x: | - |
| FORECAST | :x: | - |
| FORECAST.ETS | :x: | - |
| FORECAST.LINEAR | :x: | - |
| FREQUENCY | :x: | - |
| GAMMA | :x: | - |
| GAMMA.DIST | :x: | - |
| GAMMA.INV | :x: | - |
| GAMMALN | :x: | - |
| GAMMALN.PRECISE | :x: | - |
| GAUSS | :x: | - |
| GEOMEAN | :x: | - |
| GROWTH | :x: | - |
| HARMEAN | :x: | - |
| HYPGEOM.DIST | :x: | - |
| INTERCEPT | :x: | - |
| KURT | :x: | - |
| LARGE | :white_check_mark: | 4 |
| LINEST | :x: | - |
| LOGEST | :x: | - |
| LOGNORM.DIST | :x: | - |
| LOGNORM.INV | :x: | - |
| MAX | :white_check_mark: | 5 |
| MAXA | :x: | - |
| MAXIFS | :white_check_mark: | 15 |
| MEDIAN | :white_check_mark: | 3 |
| MIN | :white_check_mark: | 3 |
| MINA | :x: | - |
| MINIFS | :white_check_mark: | 3 |
| MODE | :white_check_mark: | 7 |
| MODE.MULT | :x: | - |
| MODE.SNGL | :x: | - |
| NEGBINOM.DIST | :x: | - |
| NORM.DIST | :x: | - |
| NORM.INV | :x: | - |
| NORM.S.DIST | :x: | - |
| NORM.S.INV | :x: | - |
| PEARSON | :x: | - |
| PERCENTILE | :white_check_mark: | 9 |
| PERCENTILE.EXC | :x: | - |
| PERCENTILE.INC | :x: | - |
| PERCENTRANK.EXC | :x: | - |
| PERCENTRANK.INC | :x: | - |
| PERMUT | :white_check_mark: | 13 |
| PERMUTATIONA | :x: | - |
| PHI | :x: | - |
| POISSON.DIST | :x: | - |
| PROB | :x: | - |
| QUARTILE.EXC | :x: | - |
| QUARTILE.INC | :x: | - |
| RANK | :white_check_mark: | 8 |
| RANK.AVG | :x: | - |
| RANK.EQ | :x: | - |
| RSQ | :x: | - |
| SKEW | :x: | - |
| SKEW.P | :x: | - |
| SLOPE | :x: | - |
| SMALL | :white_check_mark: | 2 |
| STANDARDIZE | :x: | - |
| STDEV | :white_check_mark: | 7 |
| STDEV.P | :x: | - |
| STDEV.S | :x: | - |
| STDEVA | :x: | - |
| STDEVP | :white_check_mark: | 6 |
| STDEVPA | :x: | - |
| STEYX | :x: | - |
| T.DIST | :x: | - |
| T.DIST.2T | :x: | - |
| T.DIST.RT | :x: | - |
| T.INV | :x: | - |
| T.INV.2T | :x: | - |
| T.TEST | :x: | - |
| TREND | :x: | - |
| TRIMMEAN | :x: | - |
| VAR | :white_check_mark: | 7 |
| VAR.P | :x: | - |
| VAR.S | :x: | - |
| VARA | :x: | - |
| VARP | :white_check_mark: | 6 |
| VARPA | :x: | - |
| WEIBULL.DIST | :x: | - |
| Z.TEST | :x: | - |

### Logical

| Function | Supported | Tests |
|----------|-----------|-------|
| AND | :white_check_mark: | 8 |
| BYCOL | :x: | - |
| BYROW | :x: | - |
| FALSE | :x: | - |
| IF | :white_check_mark: | 36 |
| IFERROR | :white_check_mark: | 11 |
| IFNA | :white_check_mark: | 12 |
| IFS | :white_check_mark: | 13 |
| LAMBDA | :x: | - |
| LET | :x: | - |
| MAKEARRAY | :x: | - |
| MAP | :x: | - |
| NOT | :white_check_mark: | 6 |
| OR | :white_check_mark: | 2 |
| REDUCE | :x: | - |
| SCAN | :x: | - |
| SWITCH | :white_check_mark: | 14 |
| TRUE | :x: | - |
| XOR | :white_check_mark: | 12 |

### Text

| Function | Supported | Tests |
|----------|-----------|-------|
| ARRAYTOTEXT | :x: | - |
| ASC | :x: | - |
| BAHTTEXT | :x: | - |
| CHAR | :white_check_mark: | 8 |
| CLEAN | :white_check_mark: | 2 |
| CODE | :white_check_mark: | 4 |
| CONCAT | :white_check_mark: | 1 |
| CONCATENATE | :white_check_mark: | 7 |
| DBCS | :x: | - |
| DOLLAR | :x: | - |
| EXACT | :white_check_mark: | 6 |
| FIND | :white_check_mark: | 8 |
| FINDB | :x: | - |
| FIXED | :white_check_mark: | 9 |
| LEFT | :white_check_mark: | 5 |
| LEFTB | :x: | - |
| LEN | :white_check_mark: | 3 |
| LENB | :x: | - |
| LOWER | :white_check_mark: | 1 |
| MID | :white_check_mark: | 6 |
| MIDB | :x: | - |
| NUMBERVALUE | :white_check_mark: | 13 |
| PHONETIC | :x: | - |
| PROPER | :white_check_mark: | 7 |
| REPLACE | :white_check_mark: | 6 |
| REPLACEB | :x: | - |
| REPT | :white_check_mark: | 3 |
| RIGHT | :white_check_mark: | 5 |
| RIGHTB | :x: | - |
| SEARCH | :white_check_mark: | 5 |
| SEARCHB | :x: | - |
| SUBSTITUTE | :white_check_mark: | 14 |
| T | :white_check_mark: | 5 |
| TEXT | :white_check_mark: | 14 |
| TEXTAFTER | :x: | - |
| TEXTBEFORE | :x: | - |
| TEXTJOIN | :white_check_mark: | 10 |
| TEXTSPLIT | :x: | - |
| TRIM | :white_check_mark: | 6 |
| UNICHAR | :x: | - |
| UNICODE | :x: | - |
| UPPER | :white_check_mark: | 1 |
| VALUE | :white_check_mark: | 10 |
| VALUETOTEXT | :x: | - |

### Date & Time

| Function | Supported | Tests |
|----------|-----------|-------|
| DATE | :white_check_mark: | 15 |
| DATEDIF | :white_check_mark: | 9 |
| DATEVALUE | :white_check_mark: | 5 |
| DAY | :white_check_mark: | 2 |
| DAYS | :white_check_mark: | 9 |
| DAYS360 | :x: | - |
| EDATE | :white_check_mark: | 5 |
| EOMONTH | :white_check_mark: | 5 |
| HOUR | :white_check_mark: | 1 |
| ISOWEEKNUM | :white_check_mark: | 7 |
| MINUTE | :white_check_mark: | 1 |
| MONTH | :white_check_mark: | 2 |
| NETWORKDAYS | :white_check_mark: | 5 |
| NETWORKDAYS.INTL | :x: | - |
| NOW | :white_check_mark: | 4 |
| SECOND | :white_check_mark: | 1 |
| TIME | :white_check_mark: | 4 |
| TIMEVALUE | :x: | - |
| TODAY | :white_check_mark: | 4 |
| WEEKDAY | :white_check_mark: | 17 |
| WEEKNUM | :white_check_mark: | 6 |
| WORKDAY | :white_check_mark: | 301 |
| WORKDAY.INTL | :x: | - |
| YEAR | :white_check_mark: | 2 |
| YEARFRAC | :white_check_mark: | 8 |

### Lookup & Reference

| Function | Supported | Tests |
|----------|-----------|-------|
| ADDRESS | :white_check_mark: | 14 |
| AREAS | :x: | - |
| CHOOSE | :white_check_mark: | 4 |
| CHOOSECOLS | :x: | - |
| CHOOSEROWS | :x: | - |
| COLUMN | :white_check_mark: | 5 |
| COLUMNS | :white_check_mark: | 4 |
| DROP | :x: | - |
| EXPAND | :x: | - |
| FILTER | :x: | - |
| FORMULATEXT | :x: | - |
| HLOOKUP | :white_check_mark: | 5 |
| HSTACK | :x: | - |
| HYPERLINK | :x: | - |
| INDEX | :white_check_mark: | 11 |
| INDIRECT | :x: | - |
| LOOKUP | :white_check_mark: | 5 |
| MATCH | :white_check_mark: | 14 |
| OFFSET | :x: | - |
| ROW | :white_check_mark: | 7 |
| ROWS | :white_check_mark: | 4 |
| SORT | :white_check_mark: | 6 |
| SORTBY | :x: | - |
| TAKE | :x: | - |
| TOCOL | :x: | - |
| TOROW | :x: | - |
| TRANSPOSE | :x: | - |
| UNIQUE | :x: | - |
| VLOOKUP | :white_check_mark: | 19 |
| VSTACK | :x: | - |
| XLOOKUP | :white_check_mark: | 9 |
| XMATCH | :x: | - |

### Information

| Function | Supported | Tests |
|----------|-----------|-------|
| CELL | :x: | - |
| ERROR.TYPE | :white_check_mark: | 9 |
| INFO | :x: | - |
| ISBLANK | :white_check_mark: | 2 |
| ISERR | :white_check_mark: | 7 |
| ISERROR | :white_check_mark: | 6 |
| ISEVEN | :white_check_mark: | 9 |
| ISFORMULA | :x: | - |
| ISLOGICAL | :white_check_mark: | 9 |
| ISNA | :white_check_mark: | 5 |
| ISNONTEXT | :white_check_mark: | 6 |
| ISNUMBER | :white_check_mark: | 9 |
| ISODD | :white_check_mark: | 8 |
| ISREF | :x: | - |
| ISTEXT | :white_check_mark: | 2 |
| N | :white_check_mark: | 9 |
| NA | :white_check_mark: | 4 |
| SHEET | :x: | - |
| SHEETS | :x: | - |
| TYPE | :white_check_mark: | 19 |

### Financial

| Function | Supported | Tests |
|----------|-----------|-------|
| ACCRINT | :x: | - |
| ACCRINTM | :x: | - |
| AMORDEGRC | :x: | - |
| AMORLINC | :x: | - |
| COUPDAYBS | :x: | - |
| COUPDAYS | :x: | - |
| COUPDAYSNC | :x: | - |
| COUPNCD | :x: | - |
| COUPNUM | :x: | - |
| COUPPCD | :x: | - |
| CUMIPMT | :x: | - |
| CUMPRINC | :x: | - |
| DB | :x: | - |
| DDB | :x: | - |
| DISC | :x: | - |
| DOLLARDE | :x: | - |
| DOLLARFR | :x: | - |
| DURATION | :x: | - |
| EFFECT | :x: | - |
| FV | :x: | - |
| FVSCHEDULE | :x: | - |
| INTRATE | :x: | - |
| IPMT | :x: | - |
| IRR | :x: | - |
| ISPMT | :x: | - |
| MDURATION | :x: | - |
| MIRR | :x: | - |
| NOMINAL | :x: | - |
| NPER | :x: | - |
| NPV | :x: | - |
| ODDFPRICE | :x: | - |
| ODDFYIELD | :x: | - |
| ODDLPRICE | :x: | - |
| ODDLYIELD | :x: | - |
| PDURATION | :x: | - |
| PMT | :x: | - |
| PPMT | :x: | - |
| PRICE | :x: | - |
| PRICEDISC | :x: | - |
| PRICEMAT | :x: | - |
| PV | :x: | - |
| RATE | :x: | - |
| RECEIVED | :x: | - |
| RRI | :x: | - |
| SLN | :x: | - |
| SYD | :x: | - |
| TBILLEQ | :x: | - |
| TBILLPRICE | :x: | - |
| TBILLYIELD | :x: | - |
| VDB | :x: | - |
| XIRR | :x: | - |
| XNPV | :x: | - |
| YIELD | :x: | - |
| YIELDDISC | :x: | - |
| YIELDMAT | :x: | - |

### Engineering

| Function | Supported | Tests |
|----------|-----------|-------|
| BESSELI | :x: | - |
| BESSELJ | :x: | - |
| BESSELK | :x: | - |
| BESSELY | :x: | - |
| BIN2DEC | :x: | - |
| BIN2HEX | :x: | - |
| BIN2OCT | :x: | - |
| BITAND | :x: | - |
| BITLSHIFT | :x: | - |
| BITOR | :x: | - |
| BITRSHIFT | :x: | - |
| BITXOR | :x: | - |
| COMPLEX | :x: | - |
| CONVERT | :x: | - |
| DEC2BIN | :x: | - |
| DEC2HEX | :x: | - |
| DEC2OCT | :x: | - |
| DELTA | :x: | - |
| ERF | :x: | - |
| ERF.PRECISE | :x: | - |
| ERFC | :x: | - |
| ERFC.PRECISE | :x: | - |
| GESTEP | :x: | - |
| HEX2BIN | :x: | - |
| HEX2DEC | :x: | - |
| HEX2OCT | :x: | - |
| IMABS | :x: | - |
| IMAGINARY | :x: | - |
| IMARGUMENT | :x: | - |
| IMCONJUGATE | :x: | - |
| IMCOS | :x: | - |
| IMCOSH | :x: | - |
| IMCOT | :x: | - |
| IMCSC | :x: | - |
| IMCSCH | :x: | - |
| IMDIV | :x: | - |
| IMEXP | :x: | - |
| IMLN | :x: | - |
| IMLOG10 | :x: | - |
| IMLOG2 | :x: | - |
| IMPOWER | :x: | - |
| IMPRODUCT | :x: | - |
| IMREAL | :x: | - |
| IMSEC | :x: | - |
| IMSECH | :x: | - |
| IMSIN | :x: | - |
| IMSINH | :x: | - |
| IMSQRT | :x: | - |
| IMSUB | :x: | - |
| IMSUM | :x: | - |
| IMTAN | :x: | - |
| OCT2BIN | :x: | - |
| OCT2DEC | :x: | - |
| OCT2HEX | :x: | - |

### Database

| Function | Supported | Tests |
|----------|-----------|-------|
| DAVERAGE | :x: | - |
| DCOUNT | :x: | - |
| DCOUNTA | :x: | - |
| DGET | :x: | - |
| DMAX | :x: | - |
| DMIN | :x: | - |
| DPRODUCT | :x: | - |
| DSTDEV | :x: | - |
| DSTDEVP | :x: | - |
| DSUM | :x: | - |
| DVAR | :x: | - |
| DVARP | :x: | - |

### Web

| Function | Supported | Tests |
|----------|-----------|-------|
| ENCODEURL | :x: | - |
| FILTERXML | :x: | - |
| WEBSERVICE | :x: | - |

### Cube

| Function | Supported | Tests |
|----------|-----------|-------|
| CUBEKPIMEMBER | :x: | - |
| CUBEMEMBER | :x: | - |
| CUBEMEMBERPROPERTY | :x: | - |
| CUBERANKEDMEMBER | :x: | - |
| CUBESET | :x: | - |
| CUBESETCOUNT | :x: | - |
| CUBEVALUE | :x: | - |
