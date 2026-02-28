package fuzz

import "strings"

// ComplexityConfig defines the parameters for a given complexity level.
type ComplexityConfig struct {
	Level      int
	MinCells   int
	MaxCells   int
	MinChecks  int
	MaxChecks  int
	MaxSheets  int
	MaxNesting int      // Formula nesting depth
	CrossSheet bool     // Allow cross-sheet references
	Functions  []string // Which function categories to use
	EdgeCases  string   // "none", "basic", "aggressive"
}

// Levels defines the built-in complexity progression.
var Levels = []ComplexityConfig{
	{Level: 1, MinCells: 3, MaxCells: 8, MinChecks: 2, MaxChecks: 4, MaxSheets: 1, MaxNesting: 1, Functions: []string{"math"}, EdgeCases: "none"},
	{Level: 2, MinCells: 5, MaxCells: 12, MinChecks: 3, MaxChecks: 6, MaxSheets: 1, MaxNesting: 2, Functions: []string{"math", "logic"}, EdgeCases: "basic"},
	{Level: 3, MinCells: 8, MaxCells: 15, MinChecks: 4, MaxChecks: 8, MaxSheets: 1, MaxNesting: 3, Functions: []string{"math", "logic", "text"}, EdgeCases: "basic"},
	{Level: 4, MinCells: 10, MaxCells: 20, MinChecks: 5, MaxChecks: 10, MaxSheets: 2, MaxNesting: 3, CrossSheet: true, Functions: []string{"math", "logic", "text", "stat"}, EdgeCases: "basic"},
	{Level: 5, MinCells: 12, MaxCells: 25, MinChecks: 6, MaxChecks: 12, MaxSheets: 3, MaxNesting: 4, CrossSheet: true, Functions: []string{"math", "logic", "text", "stat", "lookup"}, EdgeCases: "aggressive"},
	{Level: 6, MinCells: 15, MaxCells: 35, MinChecks: 8, MaxChecks: 15, MaxSheets: 4, MaxNesting: 5, CrossSheet: true, Functions: []string{"math", "logic", "text", "stat", "lookup", "date", "info"}, EdgeCases: "aggressive"},
	{Level: 7, MinCells: 18, MaxCells: 40, MinChecks: 10, MaxChecks: 18, MaxSheets: 4, MaxNesting: 5, CrossSheet: true, Functions: []string{"math", "logic", "text", "stat", "lookup", "date", "info", "financial"}, EdgeCases: "aggressive"},
	{Level: 8, MinCells: 20, MaxCells: 50, MinChecks: 12, MaxChecks: 20, MaxSheets: 5, MaxNesting: 6, CrossSheet: true, Functions: []string{"math", "logic", "text", "stat", "lookup", "date", "info", "financial", "engineering", "database"}, EdgeCases: "aggressive"},
}

// GetLevel returns the ComplexityConfig for the given level.
// Levels above the max reuse the highest config but scale cell/check counts.
func GetLevel(level int) *ComplexityConfig {
	if level < 1 {
		level = 1
	}
	if level <= len(Levels) {
		cfg := Levels[level-1]
		return &cfg
	}
	// Beyond max: scale up from the last defined level.
	cfg := Levels[len(Levels)-1]
	cfg.Level = level
	extra := level - len(Levels)
	cfg.MinCells += extra * 3
	cfg.MaxCells += extra * 5
	cfg.MinChecks += extra * 2
	cfg.MaxChecks += extra * 3
	return &cfg
}

// categoryFunctions maps category names to their function lists.
var categoryFunctions = map[string][]string{
	"math": {
		"ABS", "ACOSH", "ACOT", "ACOTH", "ARABIC", "ASINH", "ATANH",
		"BASE", "CEILING", "CEILING.MATH", "CEILING.PRECISE",
		"COMBIN", "COMBINA", "COS", "COSH", "COT", "COTH", "CSC", "CSCH",
		"DECIMAL", "DEGREES", "EVEN", "EXP",
		"FACT", "FACTDOUBLE", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE",
		"GCD", "INT", "LCM", "LN", "LOG", "LOG10",
		"MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND", "MULTINOMIAL", "MUNIT",
		"ODD", "PI", "POWER", "PRODUCT", "QUOTIENT",
		"RADIANS", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP",
		"SEC", "SECH", "SERIESSUM", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI",
		"SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT",
		"SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2",
		"TAN", "TANH", "TRUNC",
	},
	"logic": {
		"AND", "FALSE", "IF", "IFERROR", "IFNA", "IFS",
		"ISBLANK", "ISERR", "ISERROR", "ISNA", "ISNUMBER", "ISTEXT",
		"NOT", "OR", "SWITCH", "TRUE", "XOR",
	},
	"text": {
		"CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR",
		"EXACT", "FIND", "FIXED", "LEFT", "LEN", "LOWER", "MID",
		"NUMBERVALUE", "PROPER", "REPLACE", "REPT", "RIGHT",
		"SEARCH", "SUBSTITUTE", "T", "TEXT", "TEXTJOIN",
		"TRIM", "UNICHAR", "UNICODE", "UPPER", "VALUE",
	},
	"stat": {
		"AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS",
		"BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
		"CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
		"CONFIDENCE.NORM", "CONFIDENCE.T", "CORREL", "COUNT", "COUNTA", "COUNTBLANK",
		"COUNTIF", "COUNTIFS", "COVARIANCE.P", "COVARIANCE.S",
		"DEVSQ", "EXPON.DIST",
		"F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
		"FISHER", "FISHERINV", "FORECAST.LINEAR", "FREQUENCY",
		"GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN", "GAMMALN.PRECISE",
		"GAUSS", "GEOMEAN", "GROWTH",
		"HARMEAN", "HYPGEOM.DIST",
		"INTERCEPT", "KURT",
		"LARGE", "LINEST", "LOGEST", "LOGNORM.DIST", "LOGNORM.INV",
		"MAX", "MAXA", "MAXIFS", "MEDIAN", "MIN", "MINA", "MINIFS",
		"MODE.MULT", "MODE.SNGL",
		"NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
		"PEARSON", "PERCENTILE.EXC", "PERCENTILE.INC",
		"PERCENTRANK.EXC", "PERCENTRANK.INC",
		"PERMUT", "PERMUTATIONA", "PHI", "POISSON.DIST", "PROB",
		"QUARTILE.EXC", "QUARTILE.INC",
		"RANK.AVG", "RANK.EQ", "RSQ",
		"SKEW", "SKEW.P", "SLOPE", "SMALL", "STANDARDIZE",
		"STDEV.P", "STDEV.S", "STDEVA", "STDEVPA", "STEYX",
		"T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST",
		"TREND", "TRIMMEAN",
		"VAR.P", "VAR.S", "VARA", "VARPA",
		"WEIBULL.DIST", "Z.TEST",
	},
	"lookup": {
		"ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS",
		"FORMULATEXT", "HLOOKUP", "HYPERLINK",
		"INDEX", "LOOKUP", "MATCH",
		"ROW", "ROWS", "TRANSPOSE", "VLOOKUP",
	},
	"date": {
		"DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS", "DAYS360",
		"EDATE", "EOMONTH", "HOUR", "ISOWEEKNUM",
		"MINUTE", "MONTH",
		"NETWORKDAYS", "NETWORKDAYS.INTL",
		"SECOND", "TIME", "TIMEVALUE",
		"WEEKDAY", "WEEKNUM",
		"WORKDAY", "WORKDAY.INTL",
		"YEAR", "YEARFRAC",
	},
	"info": {
		"ERROR.TYPE", "ISBLANK", "ISERR", "ISERROR",
		"ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA", "ISNONTEXT",
		"ISNUMBER", "ISODD", "ISREF", "ISTEXT",
		"N", "NA", "SHEET", "SHEETS", "TYPE",
	},
	"financial": {
		"ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC",
		"COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD",
		"CUMIPMT", "CUMPRINC",
		"DB", "DDB", "DISC", "DOLLARDE", "DOLLARFR", "DURATION",
		"EFFECT", "FV", "FVSCHEDULE",
		"INTRATE", "IPMT", "IRR", "ISPMT",
		"MDURATION", "MIRR",
		"NOMINAL", "NPER", "NPV",
		"ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD",
		"PDURATION", "PMT", "PPMT",
		"PRICE", "PRICEDISC", "PRICEMAT", "PV",
		"RATE", "RECEIVED", "RRI",
		"SLN", "SYD",
		"TBILLEQ", "TBILLPRICE", "TBILLYIELD",
		"VDB",
		"XIRR", "XNPV",
		"YIELD", "YIELDDISC", "YIELDMAT",
	},
	"engineering": {
		"BESSELI", "BESSELJ", "BESSELK", "BESSELY",
		"BIN2DEC", "BIN2HEX", "BIN2OCT",
		"BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR",
		"COMPLEX", "CONVERT",
		"DEC2BIN", "DEC2HEX", "DEC2OCT",
		"DELTA", "ERF", "ERF.PRECISE", "ERFC", "ERFC.PRECISE",
		"GESTEP",
		"HEX2BIN", "HEX2DEC", "HEX2OCT",
		"IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE",
		"IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH",
		"IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2",
		"IMPOWER", "IMPRODUCT", "IMREAL",
		"IMSEC", "IMSECH", "IMSIN", "IMSINH", "IMSQRT",
		"IMSUB", "IMSUM", "IMTAN",
		"OCT2BIN", "OCT2DEC", "OCT2HEX",
	},
	"database": {
		"DAVERAGE", "DCOUNT", "DCOUNTA", "DGET",
		"DMAX", "DMIN", "DPRODUCT",
		"DSTDEV", "DSTDEVP", "DSUM",
		"DVAR", "DVARP",
	},
}

// FilteredFunctions returns the list of functions allowed at this complexity level.
func (c *ComplexityConfig) FilteredFunctions() []string {
	seen := make(map[string]bool)
	var funcs []string
	for _, cat := range c.Functions {
		for _, fn := range categoryFunctions[strings.ToLower(cat)] {
			if !seen[fn] {
				seen[fn] = true
				funcs = append(funcs, fn)
			}
		}
	}
	if len(funcs) == 0 {
		return KnownFunctionsList
	}
	return funcs
}
