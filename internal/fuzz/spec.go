package fuzz

import (
	"encoding/json"
	"fmt"
	"os"
	"strings"
)

// TestSpec defines a complete fuzz test case.
type TestSpec struct {
	Name   string      `json:"name"`
	Sheets []SheetSpec `json:"sheets"`
	Checks []CheckSpec `json:"checks"`
}

// SheetSpec defines one sheet within a test spec.
type SheetSpec struct {
	Name  string     `json:"name"`
	Cells []CellSpec `json:"cells"`
}

// CellSpec defines a single cell in the spec.
type CellSpec struct {
	Ref     string `json:"ref"`
	Value   any    `json:"value,omitempty"`
	Type    string `json:"type,omitempty"` // "number", "string", "bool"
	Formula string `json:"formula,omitempty"`
}

// CheckSpec defines an expected result to verify.
type CheckSpec struct {
	Ref      string `json:"ref"`      // "Sheet1!B1" or "B1" (defaults to first sheet)
	Expected string `json:"expected"`
	Type     string `json:"type"` // "number", "string", "bool", "error", "empty"
}

// commonExcludedFunctions lists functions excluded for all oracles.
var commonExcludedFunctions = map[string]bool{
	// Non-deterministic
	"RAND":        true,
	"RANDBETWEEN": true,
	"RANDARRAY":   true,
	"NOW":         true,
	"TODAY":       true,

	// Volatile/indirect resolution
	"INDIRECT": true,

	// Require external data / environment
	"WEBSERVICE":     true,
	"FILTERXML":      true,
	"ENCODEURL":      true,
	"RTD":            true,
	"GETPIVOTDATA":   true,
	"STOCKHISTORY":   true,
	"DETECTLANGUAGE": true,
	"TRANSLATE":      true,
	"REGEXEXTRACT":   true,
	"REGEXREPLACE":   true,
	"REGEXTEST":      true,

	// Environment-dependent
	"CELL": true,
	"INFO": true,

	// Cube functions (require OLAP connection)
	"CUBEKPIMEMBER":      true,
	"CUBEMEMBER":         true,
	"CUBEMEMBERPROPERTY": true,
	"CUBERANKEDMEMBER":   true,
	"CUBESET":            true,
	"CUBESETCOUNT":       true,
	"CUBEVALUE":          true,

	// User-defined/add-in
	"CALL":        true,
	"EUROCONVERT": true,
	"REGISTER.ID": true,
}

// LibreOfficeExcludedFunctions lists functions that should not appear in specs
// when using LibreOffice as the oracle. Includes all common exclusions plus
// Excel 365+ functions that LibreOffice cannot evaluate.
var LibreOfficeExcludedFunctions = func() map[string]bool {
	m := make(map[string]bool)
	for k, v := range commonExcludedFunctions {
		m[k] = v
	}
	// Excel 365+ dynamic array functions (not in LibreOffice)
	for _, fn := range []string{
		"CONCAT", "XLOOKUP", "XMATCH", "SORT", "SORTBY",
		"FILTER", "UNIQUE", "SEQUENCE", "LET", "LAMBDA",
		"MAP", "REDUCE", "SCAN", "BYCOL", "BYROW",
		"MAKEARRAY", "CHOOSECOLS", "CHOOSEROWS", "DROP", "TAKE",
		"EXPAND", "HSTACK", "VSTACK", "WRAPCOLS", "WRAPROWS",
		"TOCOL", "TOROW", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
		"VALUETOTEXT", "ARRAYTOTEXT", "ISOMITTED", "IMAGE",
		"GROUPBY", "PIVOTBY", "PERCENTOF", "TRIMRANGE",
	} {
		m[fn] = true
	}
	return m
}()

// ExcelOnlineExcludedFunctions lists functions that should not appear in specs
// when using Excel Online (MS Graph) as the oracle. Much smaller than LibreOffice
// since Excel 365 supports XLOOKUP, CONCAT, dynamic arrays, LET, LAMBDA, etc.
var ExcelOnlineExcludedFunctions = func() map[string]bool {
	m := make(map[string]bool)
	for k, v := range commonExcludedFunctions {
		m[k] = v
	}
	// Excel Online still can't do these:
	for _, fn := range []string{
		"IMAGE",    // requires desktop Excel
		"GROUPBY",  // very new, may not be in web API
		"PIVOTBY",  // very new, may not be in web API
		"PERCENTOF", // very new
		"TRIMRANGE", // very new
	} {
		m[fn] = true
	}
	return m
}()

// LocalExcelExcludedFunctions lists functions that should not appear in specs
// when using local desktop Excel as the oracle. Desktop Excel supports all modern
// functions, so only the common exclusions apply.
var LocalExcelExcludedFunctions = func() map[string]bool {
	m := make(map[string]bool)
	for k, v := range commonExcludedFunctions {
		m[k] = v
	}
	return m
}()

// ExcludedFunctions is the default excluded functions list (LibreOffice).
// Kept for backward compatibility.
var ExcludedFunctions = LibreOfficeExcludedFunctions

// LoadSpec reads and validates a test spec from a JSON file.
// If excluded is provided, it is used for validation instead of the default.
func LoadSpec(path string, excluded ...map[string]bool) (*TestSpec, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("read spec: %w", err)
	}
	var spec TestSpec
	if err := json.Unmarshal(data, &spec); err != nil {
		return nil, fmt.Errorf("parse spec: %w", err)
	}
	if err := ValidateSpec(&spec, excluded...); err != nil {
		return nil, err
	}
	return &spec, nil
}

// SaveSpec writes a test spec to a JSON file.
func SaveSpec(path string, spec *TestSpec) error {
	data, err := json.MarshalIndent(spec, "", "  ")
	if err != nil {
		return fmt.Errorf("marshal spec: %w", err)
	}
	return os.WriteFile(path, data, 0644)
}

// ValidateSpec checks that a spec is well-formed.
// If excluded is nil, the default ExcludedFunctions (LibreOffice) is used.
func ValidateSpec(spec *TestSpec, excluded ...map[string]bool) error {
	excl := ExcludedFunctions
	if len(excluded) > 0 && excluded[0] != nil {
		excl = excluded[0]
	}

	if spec.Name == "" {
		return fmt.Errorf("spec missing name")
	}
	if len(spec.Sheets) == 0 {
		return fmt.Errorf("spec has no sheets")
	}
	if len(spec.Checks) == 0 {
		return fmt.Errorf("spec has no checks")
	}
	for i, sheet := range spec.Sheets {
		if sheet.Name == "" {
			return fmt.Errorf("sheet %d missing name", i)
		}
		for j, cell := range sheet.Cells {
			if cell.Ref == "" {
				return fmt.Errorf("sheet %q cell %d missing ref", sheet.Name, j)
			}
			if cell.Formula == "" && cell.Value == nil {
				return fmt.Errorf("sheet %q cell %q has neither value nor formula", sheet.Name, cell.Ref)
			}
			// Check for excluded functions in formulas.
			if cell.Formula != "" {
				upper := strings.ToUpper(cell.Formula)
				for fn := range excl {
					if strings.Contains(upper, fn+"(") {
						return fmt.Errorf("sheet %q cell %q uses excluded function %s", sheet.Name, cell.Ref, fn)
					}
				}
			}
		}
	}
	for i, check := range spec.Checks {
		if check.Ref == "" {
			return fmt.Errorf("check %d missing ref", i)
		}
		if check.Type == "" {
			return fmt.Errorf("check %d missing type", i)
		}
	}
	return nil
}

// ParseCheckRef splits "Sheet1!B1" into (sheet, cellRef). If no sheet prefix, returns "".
func ParseCheckRef(ref string) (sheet, cellRef string) {
	if idx := strings.Index(ref, "!"); idx >= 0 {
		return ref[:idx], ref[idx+1:]
	}
	return "", ref
}

// SeedCategory maps user-provided seed to a prompt category hint.
func SeedCategory(seed string) string {
	switch strings.ToLower(seed) {
	case "math":
		return "math (ABS, CEILING, CEILING.MATH, COMBIN, COS, COSH, DEGREES, EVEN, EXP, FACT, FACTDOUBLE, FLOOR, FLOOR.MATH, GCD, INT, LCM, LN, LOG, LOG10, MDETERM, MOD, MROUND, MULTINOMIAL, ODD, PI, POWER, PRODUCT, QUOTIENT, RADIANS, ROUND, ROUNDDOWN, ROUNDUP, SIGN, SIN, SINH, SQRT, SQRTPI, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ, TAN, TANH, TRUNC)"
	case "text":
		return "text (CHAR, CLEAN, CODE, CONCATENATE, DOLLAR, EXACT, FIND, FIXED, LEFT, LEN, LOWER, MID, NUMBERVALUE, PROPER, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, T, TEXT, TEXTJOIN, TRIM, UNICHAR, UNICODE, UPPER, VALUE)"
	case "lookup":
		return "lookup and reference (ADDRESS, AREAS, CHOOSE, COLUMN, COLUMNS, FORMULATEXT, HLOOKUP, INDEX, LOOKUP, MATCH, ROW, ROWS, TRANSPOSE, VLOOKUP)"
	case "logic":
		return "logical (AND, FALSE, IF, IFERROR, IFNA, IFS, ISBLANK, ISERR, ISERROR, ISNA, ISNUMBER, ISTEXT, NOT, OR, SWITCH, TRUE, XOR)"
	case "stat":
		return "statistical (AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS, BINOM.DIST, CHISQ.DIST, CONFIDENCE.NORM, CORREL, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COVARIANCE.P, EXPON.DIST, F.DIST, FISHER, FORECAST.LINEAR, GAMMA.DIST, GEOMEAN, HARMEAN, LARGE, MAX, MAXA, MAXIFS, MEDIAN, MIN, MINA, MINIFS, MODE.SNGL, NORM.DIST, NORM.INV, PEARSON, PERCENTILE.INC, PERMUT, POISSON.DIST, RANK.AVG, RANK.EQ, SMALL, STDEV.P, STDEV.S, T.DIST, TRIMMEAN, VAR.P, VAR.S, WEIBULL.DIST, Z.TEST)"
	case "date":
		return "date/time (DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS, SECOND, TIME, TIMEVALUE, WEEKDAY, WEEKNUM, WORKDAY, YEAR, YEARFRAC)"
	case "info":
		return "information (ERROR.TYPE, ISBLANK, ISERR, ISERROR, ISEVEN, ISFORMULA, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISODD, ISREF, ISTEXT, N, NA, TYPE)"
	case "financial":
		return "financial (ACCRINT, CUMIPMT, CUMPRINC, DB, DDB, DISC, DOLLARDE, DOLLARFR, DURATION, EFFECT, FV, FVSCHEDULE, INTRATE, IPMT, IRR, ISPMT, MDURATION, MIRR, NOMINAL, NPER, NPV, PDURATION, PMT, PPMT, PRICE, PRICEDISC, PV, RATE, RECEIVED, RRI, SLN, SYD, TBILLEQ, TBILLPRICE, TBILLYIELD, VDB, XIRR, XNPV, YIELD, YIELDDISC, YIELDMAT)"
	case "engineering":
		return "engineering (BESSELI, BESSELJ, BIN2DEC, BIN2HEX, BIN2OCT, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, COMPLEX, CONVERT, DEC2BIN, DEC2HEX, DEC2OCT, DELTA, ERF, ERFC, GESTEP, HEX2BIN, HEX2DEC, HEX2OCT, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSIN, IMSQRT, IMSUB, IMSUM, OCT2BIN, OCT2DEC, OCT2HEX)"
	case "database":
		return "database (DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP)"
	case "mixed", "":
		return ""
	default:
		return seed
	}
}
