// Script to regenerate FORMULAS.md from the source of truth:
//  1. Supported functions = all Register() calls in formula/*.go
//  2. Test counts = formula strings + TestFUNC_ patterns in *_test.go
//  3. Unsupported functions = those in the known function list but not registered
//
// Usage: go run scripts/count_formula_tests.go [--check]
//
//	(default) regenerate FORMULAS.md
//	--check   exit 1 if FORMULAS.md would change (for CI)
package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// willNotSupport lists functions that depend on the spreadsheet application's
// runtime environment, have side effects, or require locale-specific behavior
// that cannot be reproduced in a server-side library.
var willNotSupport = map[string]string{
	"WEBSERVICE": "Makes HTTP requests from formulas — security risk, side effects",
	"FILTERXML":  "Paired with WEBSERVICE; fetches/parses external XML",
	"INFO":       "Returns application environment info (OS version, memory, app version)",
	"CELL":       "Many modes return application-specific environment info (filename, format codes from the UI)",
	"PHONETIC":   "Returns Japanese furigana metadata from the IME — not stored in XLSX files",
	"ASC":        "Full-width → half-width conversion; behavior depends on the application's MBCS locale setting",
	"DBCS":       "Half-width → full-width conversion; behavior depends on the application's MBCS locale setting",
	"BAHTTEXT":   "Converts numbers to Thai Baht text — extremely locale-specific",
	"FINDB":      "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
	"LEFTB":      "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
	"LENB":       "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
	"MIDB":       "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
	"REPLACEB":   "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
	"RIGHTB":     "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
	"SEARCHB":    "Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS)",
}

// categoryMap maps source file suffix to display category.
var categoryMap = map[string]string{
	"date":    "Date & Time",
	"eng":     "Engineering",
	"finance": "Financial",
	"info":    "Information",
	"logic":   "Logical",
	"lookup":  "Lookup & Reference",
	"math":    "Math & Trig",
	"stat":    "Statistical",
	"text":    "Text",
}

type funcInfo struct {
	Name        string
	Description string
	Category    string
	Tests       int
}

func main() {
	check := len(os.Args) > 1 && os.Args[1] == "--check"

	// Step 1: discover supported functions from Register() calls
	supported := discoverSupported()

	// Step 2: count tests
	testCounts := countTests(supported)
	for name := range testCounts {
		if fi, ok := supported[name]; ok {
			fi.Tests = testCounts[name]
			supported[name] = fi
		}
	}

	// Step 3: load known functions and compute unsupported
	unsupported := computeUnsupported(supported)

	// Step 4: generate FORMULAS.md content
	content := generateMarkdown(supported, unsupported)

	if check {
		existing, err := os.ReadFile("FORMULAS.md")
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error reading FORMULAS.md: %v\n", err)
			os.Exit(1)
		}
		if string(existing) == content {
			fmt.Println("FORMULAS.md is up to date.")
			os.Exit(0)
		}
		fmt.Println("FORMULAS.md is out of date. Run: go run scripts/count_formula_tests.go")
		os.Exit(1)
	}

	if err := os.WriteFile("FORMULAS.md", []byte(content), 0644); err != nil {
		fmt.Fprintf(os.Stderr, "Error writing FORMULAS.md: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Updated FORMULAS.md: %d supported, %d unsupported\n",
		len(supported), len(unsupported))
}

// parserSupported lists functions that are implemented at the parser level
// rather than via Register(). These are handled specially (e.g. desugared
// during parsing) and won't appear in Register() calls.
var parserSupported = map[string]string{
	"LAMBDA": "Logical",
	"LET":    "Logical",
	"MAP":    "Logical",
	"REDUCE": "Logical",
	"SCAN":   "Logical",
	"BYROW":  "Logical",
	"BYCOL":     "Logical",
	"MAKEARRAY": "Logical",
}

// descriptionMap holds locally authored, concise descriptions for formulas
// that are awkward to derive from their names alone. Straightforward families
// are handled by the rule-based helpers below.
var descriptionMap = map[string]string{
	"ACCRINT":                  "Returns accrued interest for a security that pays periodic interest.",
	"ACCRINTM":                 "Returns accrued interest for a security that pays interest at maturity.",
	"ADDRESS":                  "Builds a cell reference from row and column numbers.",
	"AGGREGATE":                "Applies an aggregate calculation with options to ignore selected values.",
	"AMORDEGRC":                "Returns depreciation for each accounting period using a French declining balance method.",
	"AMORLINC":                 "Returns depreciation for each accounting period using a linear method.",
	"ANCHORARRAY":              "Returns the full spilled array produced by a dynamic array formula.",
	"AND":                      "Returns TRUE when every supplied condition is TRUE.",
	"AREAS":                    "Counts the number of areas in a reference.",
	"ARRAYTOTEXT":              "Converts an array into a text representation.",
	"ASC":                      "Converts full-width characters to half-width characters.",
	"AVEDEV":                   "Returns the average absolute deviation from the mean.",
	"AVERAGE":                  "Returns the arithmetic mean of the supplied values.",
	"AVERAGEA":                 "Returns the average of supplied values, including logical values and text coercions.",
	"AVERAGEIF":                "Returns the average of values that match one condition.",
	"AVERAGEIFS":               "Returns the average of values that match all supplied conditions.",
	"BAHTTEXT":                 "Converts a number to Thai Baht text.",
	"CELL":                     "Returns information about a cell's contents, location, or formatting.",
	"CHAR":                     "Returns the character for a numeric code.",
	"CHOOSE":                   "Returns the value at a position from a supplied list of choices.",
	"CHOOSECOLS":               "Returns selected columns from an array.",
	"CHOOSEROWS":               "Returns selected rows from an array.",
	"CLEAN":                    "Removes non-printing characters from text.",
	"CODE":                     "Returns the numeric code for the first character in text.",
	"COLUMN":                   "Returns the column number of a reference.",
	"COLUMNS":                  "Returns the number of columns in a reference or array.",
	"COMPLEX":                  "Builds a complex number from real and imaginary parts.",
	"CONCAT":                   "Joins text from multiple values or ranges.",
	"CONCATENATE":              "Joins multiple text values into one string.",
	"CONFIDENCE.NORM":          "Returns a normal-distribution confidence interval half-width.",
	"CONFIDENCE.T":             "Returns a Student's t-distribution confidence interval half-width.",
	"CONVERT":                  "Converts a number from one measurement system to another.",
	"CORREL":                   "Returns the correlation coefficient between two data sets.",
	"COUNT":                    "Counts numeric values.",
	"COUNTA":                   "Counts non-empty values.",
	"COUNTBLANK":               "Counts blank cells in a range.",
	"COUNTIF":                  "Counts values that match one condition.",
	"COUNTIFS":                 "Counts values that match all supplied conditions.",
	"COUPDAYBS":                "Returns the number of days from the coupon period start to settlement.",
	"COUPDAYS":                 "Returns the number of days in the coupon period containing settlement.",
	"COUPDAYSNC":               "Returns the number of days from settlement to the next coupon date.",
	"COUPNCD":                  "Returns the next coupon date after settlement.",
	"COUPNUM":                  "Returns the number of coupons payable between settlement and maturity.",
	"COUPPCD":                  "Returns the previous coupon date before settlement.",
	"COVAR":                    "Returns the population covariance of two data sets.",
	"COVARIANCE.P":             "Returns the population covariance of two data sets.",
	"COVARIANCE.S":             "Returns the sample covariance of two data sets.",
	"CUMIPMT":                  "Returns cumulative interest paid over a span of payment periods.",
	"CUMPRINC":                 "Returns cumulative principal paid over a span of payment periods.",
	"DATE":                     "Builds a date from year, month, and day numbers.",
	"DATEDIF":                  "Returns the difference between two dates in requested units.",
	"DATEVALUE":                "Converts date text to a serial date number.",
	"DAVERAGE":                 "Returns the average of matching records in a database range.",
	"DAY":                      "Returns the day of the month from a date.",
	"DAYS":                     "Returns the number of days between two dates.",
	"DAYS360":                  "Returns the number of days between two dates using a 360-day year.",
	"DB":                       "Returns depreciation using the fixed-declining balance method.",
	"DBCS":                     "Converts half-width characters to full-width characters.",
	"DCOUNT":                   "Counts numeric values in matching records of a database range.",
	"DCOUNTA":                  "Counts non-empty values in matching records of a database range.",
	"DDB":                      "Returns depreciation using the double-declining balance method or another factor.",
	"DEVSQ":                    "Returns the sum of squared deviations from the sample mean.",
	"DGET":                     "Returns a single value from matching records in a database range.",
	"DISC":                     "Returns the discount rate for a security.",
	"DMAX":                     "Returns the maximum value from matching records in a database range.",
	"DMIN":                     "Returns the minimum value from matching records in a database range.",
	"DOLLAR":                   "Formats a number as currency text.",
	"DOLLARDE":                 "Converts a fractional dollar price to a decimal price.",
	"DOLLARFR":                 "Converts a decimal dollar price to a fractional price.",
	"DPRODUCT":                 "Returns the product of matching records in a database range.",
	"DROP":                     "Drops rows or columns from the start or end of an array.",
	"DSTDEV":                   "Returns the sample standard deviation of matching database records.",
	"DSTDEVP":                  "Returns the population standard deviation of matching database records.",
	"DSUM":                     "Returns the sum of matching records in a database range.",
	"DURATION":                 "Returns the Macauley duration of a security.",
	"DVAR":                     "Returns the sample variance of matching database records.",
	"DVARP":                    "Returns the population variance of matching database records.",
	"EDATE":                    "Shifts a date by a number of whole months.",
	"EFFECT":                   "Returns the effective annual interest rate for a nominal rate.",
	"ENCODEURL":                "Encodes text so it can be safely used in a URL.",
	"EOMONTH":                  "Returns the last day of a month offset from a start date.",
	"ERF.PRECISE":              "Returns the error function integrated from 0 to a limit.",
	"ERFC.PRECISE":             "Returns the complementary error function integrated from a limit to infinity.",
	"ERROR.TYPE":               "Returns a number identifying an error value.",
	"EXACT":                    "Returns TRUE when two text values match exactly.",
	"EXPAND":                   "Pads an array to a larger size with a fill value.",
	"F.TEST":                   "Returns the result of an F-test for two arrays.",
	"FILTER":                   "Filters an array to rows or columns that meet a Boolean include mask.",
	"FILTERXML":                "Extracts data from XML by XPath.",
	"FIND":                     "Returns the position of one text value inside another, case-sensitive.",
	"FINDB":                    "Returns the byte position of one text value inside another.",
	"FISHER":                   "Returns the Fisher transformation of a correlation coefficient.",
	"FISHERINV":                "Returns the inverse Fisher transformation.",
	"FIXED":                    "Formats a number as text with a fixed number of decimals.",
	"FORECAST":                 "Predicts a future value using linear regression.",
	"FORECAST.ETS":             "Predicts a future value using exponential smoothing.",
	"FORECAST.ETS.CONFINT":     "Returns a confidence interval for an exponential smoothing forecast.",
	"FORECAST.ETS.SEASONALITY": "Returns the detected seasonality length for exponential smoothing.",
	"FORECAST.ETS.STAT":        "Returns forecast statistics from exponential smoothing.",
	"FORECAST.LINEAR":          "Predicts a future value using linear regression.",
	"FORMULATEXT":              "Returns the formula text from a referenced cell.",
	"FREQUENCY":                "Returns a frequency distribution for numeric bins.",
	"FV":                       "Returns the future value of an investment or annuity.",
	"FVSCHEDULE":               "Returns a future value after applying a schedule of interest rates.",
	"GAUSS":                    "Returns the probability that a standard normal variable lies between the mean and a value.",
	"GEOMEAN":                  "Returns the geometric mean of positive values.",
	"GESTEP":                   "Returns 1 when a number meets or exceeds a threshold, otherwise 0.",
	"GETPIVOTDATA":             "Returns a value from a pivot table by field and item labels.",
	"GROWTH":                   "Fits or projects values along an exponential trend.",
	"HARMEAN":                  "Returns the harmonic mean of positive values.",
	"HLOOKUP":                  "Looks up a value across the top row of a table and returns a value from a specified row.",
	"HOUR":                     "Returns the hour from a time or datetime value.",
	"HSTACK":                   "Stacks arrays horizontally.",
	"HYPERLINK":                "Creates a clickable hyperlink.",
	"IF":                       "Returns one value when a condition is TRUE and another when it is FALSE.",
	"IFERROR":                  "Returns a fallback value when a formula returns an error.",
	"IFNA":                     "Returns a fallback value when a formula returns #N/A.",
	"IFS":                      "Evaluates conditions in order and returns the value for the first TRUE condition.",
	"INDEX":                    "Returns a value or subrange at a row and column position.",
	"INDIRECT":                 "Turns text into a reference and returns its value.",
	"INTERCEPT":                "Returns the y-intercept of a linear regression line.",
	"INTRATE":                  "Returns the interest rate for a fully invested security.",
	"IPMT":                     "Returns the interest portion of a payment for a given period.",
	"IRR":                      "Returns the internal rate of return for periodic cash flows.",
	"ISBLANK":                  "Returns TRUE when a value is blank.",
	"ISERR":                    "Returns TRUE for any error except #N/A.",
	"ISERROR":                  "Returns TRUE for any error value.",
	"ISEVEN":                   "Returns TRUE when a number is even.",
	"ISFORMULA":                "Returns TRUE when a referenced cell contains a formula.",
	"ISLOGICAL":                "Returns TRUE when a value is logical.",
	"ISNA":                     "Returns TRUE when a value is #N/A.",
	"ISNONTEXT":                "Returns TRUE when a value is not text.",
	"ISNUMBER":                 "Returns TRUE when a value is numeric.",
	"ISODD":                    "Returns TRUE when a number is odd.",
	"ISOWEEKNUM":               "Returns the ISO week number for a date.",
	"ISPMT":                    "Returns the interest paid during a period for straight-line amortization.",
	"ISREF":                    "Returns TRUE when a value is a reference.",
	"ISTEXT":                   "Returns TRUE when a value is text.",
	"INFO":                     "Returns environment information about the current workbook session.",
	"KURT":                     "Returns the kurtosis of a data set.",
	"LAMBDA":                   "Defines a reusable custom formula from parameters and an expression.",
	"LARGE":                    "Returns the k-th largest value in a data set.",
	"LEFT":                     "Returns characters from the left side of a text value.",
	"LEFTB":                    "Returns bytes from the left side of a text value.",
	"LEN":                      "Returns the number of characters in text.",
	"LENB":                     "Returns the number of bytes in text.",
	"LET":                      "Assigns names to intermediate values within a formula.",
	"LINEST":                   "Returns linear regression statistics.",
	"LOGEST":                   "Returns exponential regression statistics.",
	"LOOKUP":                   "Looks up a value in a one-dimensional range and returns the matching result.",
	"LOWER":                    "Converts text to lowercase.",
	"MAKEARRAY":                "Builds an array by applying a lambda to row and column indexes.",
	"MAP":                      "Applies a lambda element-by-element across one or more arrays.",
	"MATCH":                    "Returns the relative position of a lookup value in a range or array.",
	"MAX":                      "Returns the largest numeric value.",
	"MAXA":                     "Returns the largest value, counting logical values and text coercions.",
	"MAXIFS":                   "Returns the maximum value that matches all supplied conditions.",
	"MDURATION":                "Returns the modified duration of a security.",
	"MEDIAN":                   "Returns the median of a data set.",
	"MID":                      "Returns characters from the middle of a text value.",
	"MIDB":                     "Returns bytes from the middle of a text value.",
	"MIN":                      "Returns the smallest numeric value.",
	"MINA":                     "Returns the smallest value, counting logical values and text coercions.",
	"MINIFS":                   "Returns the minimum value that matches all supplied conditions.",
	"MINUTE":                   "Returns the minute from a time or datetime value.",
	"MIRR":                     "Returns the modified internal rate of return.",
	"MODE":                     "Returns the most frequently occurring value.",
	"MODE.MULT":                "Returns all modes in a data set.",
	"MODE.SNGL":                "Returns a single mode from a data set.",
	"MOD":                      "Returns the remainder after division.",
	"MONTH":                    "Returns the month number from a date.",
	"N":                        "Converts a value to a number using spreadsheet coercion rules.",
	"NA":                       "Returns the #N/A error value.",
	"NETWORKDAYS":              "Counts working days between two dates.",
	"NETWORKDAYS.INTL":         "Counts working days between two dates using a custom weekend pattern.",
	"NOMINAL":                  "Returns the nominal annual interest rate for an effective rate.",
	"NOT":                      "Returns the opposite of a logical value.",
	"NOW":                      "Returns the current date and time.",
	"NPER":                     "Returns the number of periods for an investment or loan.",
	"NPV":                      "Returns the net present value of periodic cash flows.",
	"NUMBERVALUE":              "Converts locale-formatted text to a number.",
	"ODDFPRICE":                "Returns the price of a security with an odd first period.",
	"ODDFYIELD":                "Returns the yield of a security with an odd first period.",
	"ODDLPRICE":                "Returns the price of a security with an odd last period.",
	"ODDLYIELD":                "Returns the yield of a security with an odd last period.",
	"OFFSET":                   "Returns a reference offset from a starting reference.",
	"OR":                       "Returns TRUE when any supplied condition is TRUE.",
	"PDURATION":                "Returns the periods required for an investment to reach a target value.",
	"PEARSON":                  "Returns the Pearson correlation coefficient.",
	"PERCENTILE":               "Returns the percentile of a data set.",
	"PERCENTILE.EXC":           "Returns the exclusive percentile of a data set.",
	"PERCENTILE.INC":           "Returns the inclusive percentile of a data set.",
	"PERCENTRANK":              "Returns the rank of a value as a percentage of a data set.",
	"PERCENTRANK.EXC":          "Returns the exclusive percentage rank of a value in a data set.",
	"PERCENTRANK.INC":          "Returns the inclusive percentage rank of a value in a data set.",
	"PHONETIC":                 "Returns furigana text associated with a string.",
	"PMT":                      "Returns the periodic payment for a loan or annuity.",
	"PPMT":                     "Returns the principal portion of a payment for a given period.",
	"PRICE":                    "Returns the price per 100 face value of a coupon-paying security.",
	"PRICEDISC":                "Returns the price per 100 face value of a discounted security.",
	"PRICEMAT":                 "Returns the price per 100 face value of a security that pays interest at maturity.",
	"PROB":                     "Returns the probability that values fall within a range.",
	"PROPER":                   "Capitalizes words in text.",
	"PV":                       "Returns the present value of an investment or loan.",
	"QUARTILE":                 "Returns a quartile of a data set.",
	"QUARTILE.EXC":             "Returns the exclusive quartile of a data set.",
	"QUARTILE.INC":             "Returns the inclusive quartile of a data set.",
	"RANK":                     "Returns the rank of a number in a list.",
	"RANK.AVG":                 "Returns the average rank of a number in a list with ties.",
	"RANK.EQ":                  "Returns the rank of a number in a list with ties sharing the same rank.",
	"RATE":                     "Returns the interest rate per period for an annuity.",
	"RECEIVED":                 "Returns the amount received at maturity for a fully invested security.",
	"REDUCE":                   "Folds an array to a single result by repeatedly applying a lambda.",
	"REPLACE":                  "Replaces characters within text at a given position.",
	"REPLACEB":                 "Replaces bytes within text at a given position.",
	"REPT":                     "Repeats text a specified number of times.",
	"RIGHT":                    "Returns characters from the right side of a text value.",
	"RIGHTB":                   "Returns bytes from the right side of a text value.",
	"ROMAN":                    "Converts an Arabic number to Roman numeral text.",
	"ROW":                      "Returns the row number of a reference.",
	"ROWS":                     "Returns the number of rows in a reference or array.",
	"RRI":                      "Returns an equivalent interest rate for investment growth.",
	"RSQ":                      "Returns the square of the Pearson correlation coefficient.",
	"SCAN":                     "Returns the running accumulation produced by a lambda over an array.",
	"SEARCH":                   "Returns the position of one text value inside another, case-insensitive.",
	"SEARCHB":                  "Returns the byte position of one text value inside another, case-insensitive.",
	"SECOND":                   "Returns the second from a time or datetime value.",
	"SHEET":                    "Returns the sheet index of a reference.",
	"SHEETS":                   "Returns the number of sheets in a reference or workbook scope.",
	"SKEW":                     "Returns the sample skewness of a data set.",
	"SKEW.P":                   "Returns the population skewness of a data set.",
	"SLN":                      "Returns straight-line depreciation for one period.",
	"SLOPE":                    "Returns the slope of a linear regression line.",
	"SMALL":                    "Returns the k-th smallest value in a data set.",
	"SORT":                     "Sorts an array by row or column order.",
	"SORTBY":                   "Sorts an array by one or more companion arrays.",
	"STANDARDIZE":              "Returns a normalized value from a mean and standard deviation.",
	"STDEV":                    "Returns the sample standard deviation.",
	"STDEV.P":                  "Returns the population standard deviation.",
	"STDEV.S":                  "Returns the sample standard deviation.",
	"STDEVA":                   "Returns the sample standard deviation including logical values and text coercions.",
	"STDEVP":                   "Returns the population standard deviation.",
	"STDEVPA":                  "Returns the population standard deviation including logical values and text coercions.",
	"STEYX":                    "Returns the standard error of predicted y values in regression.",
	"SUBSTITUTE":               "Replaces matching text within a string.",
	"SUBTOTAL":                 "Returns a subtotal using a selected aggregate function.",
	"SUM":                      "Returns the sum of supplied numbers.",
	"SUMIF":                    "Returns the sum of values that match one condition.",
	"SUMIFS":                   "Returns the sum of values that match all supplied conditions.",
	"SUMPRODUCT":               "Returns the sum of pairwise products across arrays.",
	"SUMSQ":                    "Returns the sum of squares of the supplied values.",
	"SWITCH":                   "Matches an expression against a list of values and returns the corresponding result.",
	"SYD":                      "Returns sum-of-years'-digits depreciation for a period.",
	"TAKE":                     "Returns a requested number of rows or columns from an array.",
	"T":                        "Returns text when a value is text, otherwise an empty string.",
	"TBILLEQ":                  "Returns the bond-equivalent yield for a Treasury bill.",
	"TBILLPRICE":               "Returns the price per 100 face value for a Treasury bill.",
	"TBILLYIELD":               "Returns the yield for a Treasury bill.",
	"TEXT":                     "Formats a value as text using a number format pattern.",
	"TEXTAFTER":                "Returns the text that appears after a delimiter.",
	"TEXTBEFORE":               "Returns the text that appears before a delimiter.",
	"TEXTJOIN":                 "Joins text values with a delimiter.",
	"TEXTSPLIT":                "Splits text into rows or columns around delimiters.",
	"TIME":                     "Builds a time value from hour, minute, and second numbers.",
	"TIMEVALUE":                "Converts time text to a serial time value.",
	"TOCOL":                    "Flattens an array into a single column.",
	"TODAY":                    "Returns the current date.",
	"TOROW":                    "Flattens an array into a single row.",
	"TRANSPOSE":                "Swaps the rows and columns of an array.",
	"TREND":                    "Returns values along a linear trend.",
	"TRIM":                     "Removes extra spaces from text.",
	"TRIMMEAN":                 "Returns the mean after trimming values from both tails of a data set.",
	"TRUE":                     "Returns the logical value TRUE.",
	"TYPE":                     "Returns a numeric code describing a value's type.",
	"UNICHAR":                  "Returns the Unicode character for a code point.",
	"UNICODE":                  "Returns the Unicode code point for the first character in text.",
	"UNIQUE":                   "Returns distinct rows or columns from an array.",
	"UPPER":                    "Converts text to uppercase.",
	"VALUE":                    "Converts text that looks like a number into a numeric value.",
	"VALUETOTEXT":              "Converts a value to a text representation.",
	"VAR":                      "Returns the sample variance.",
	"VAR.P":                    "Returns the population variance.",
	"VAR.S":                    "Returns the sample variance.",
	"VARA":                     "Returns the sample variance including logical values and text coercions.",
	"VARP":                     "Returns the population variance.",
	"VARPA":                    "Returns the population variance including logical values and text coercions.",
	"VDB":                      "Returns depreciation using the variable declining balance method.",
	"VLOOKUP":                  "Looks up a value in the first column of a table and returns a value from another column.",
	"VSTACK":                   "Stacks arrays vertically.",
	"WEBSERVICE":               "Returns data from a web service.",
	"WEEKDAY":                  "Returns the day of the week for a date.",
	"WEEKNUM":                  "Returns the week number of a date.",
	"WORKDAY":                  "Returns a working day offset from a start date.",
	"WORKDAY.INTL":             "Returns a working day offset using a custom weekend pattern.",
	"WRAPCOLS":                 "Wraps a vector into a two-dimensional array by columns.",
	"WRAPROWS":                 "Wraps a vector into a two-dimensional array by rows.",
	"XIRR":                     "Returns the internal rate of return for cash flows on irregular dates.",
	"XLOOKUP":                  "Looks up a value in one array and returns the matching value from another array.",
	"XMATCH":                   "Returns the position of a lookup value with exact, wildcard, or binary search modes.",
	"XNPV":                     "Returns the net present value of cash flows on irregular dates.",
	"XOR":                      "Returns TRUE when an odd number of supplied conditions are TRUE.",
	"YEAR":                     "Returns the year from a date.",
	"YEARFRAC":                 "Returns the fraction of a year between two dates.",
	"YIELD":                    "Returns the yield of a coupon-paying security.",
	"YIELDDISC":                "Returns the annual yield of a discounted security.",
	"YIELDMAT":                 "Returns the annual yield of a security that pays interest at maturity.",
	"BYCOL":                    "Applies a lambda to each column of an array.",
	"BYROW":                    "Applies a lambda to each row of an array.",
	"Z.TEST":                   "Returns the one-tailed probability value of a z-test.",
}

var simpleDescriptionMap = map[string]string{
	"ABS":          "Returns the absolute value of a number.",
	"ACOS":         "Returns the arccosine of a number.",
	"ACOSH":        "Returns the inverse hyperbolic cosine of a number.",
	"ACOT":         "Returns the arccotangent of a number.",
	"ACOTH":        "Returns the inverse hyperbolic cotangent of a number.",
	"ARABIC":       "Converts Roman numeral text to an Arabic number.",
	"ASIN":         "Returns the arcsine of a number.",
	"ASINH":        "Returns the inverse hyperbolic sine of a number.",
	"ATAN":         "Returns the arctangent of a number.",
	"ATAN2":        "Returns the arctangent from x and y coordinates.",
	"ATANH":        "Returns the inverse hyperbolic tangent of a number.",
	"BASE":         "Converts a number to text in the requested base.",
	"CEILING":      "Rounds a number up to the nearest multiple of a significance.",
	"COMBIN":       "Returns the number of combinations for a given number of items.",
	"COMBINA":      "Returns the number of combinations with repetitions.",
	"COS":          "Returns the cosine of an angle.",
	"COSH":         "Returns the hyperbolic cosine of a number.",
	"COT":          "Returns the cotangent of an angle.",
	"COTH":         "Returns the hyperbolic cotangent of a number.",
	"CSC":          "Returns the cosecant of an angle.",
	"CSCH":         "Returns the hyperbolic cosecant of a number.",
	"DECIMAL":      "Converts a number in a given base to decimal.",
	"DEGREES":      "Converts radians to degrees.",
	"DELTA":        "Tests whether two numbers are equal and returns 1 or 0.",
	"ERF":          "Returns the error function.",
	"ERFC":         "Returns the complementary error function.",
	"EVEN":         "Rounds a number away from zero to the nearest even integer.",
	"EXP":          "Returns e raised to a power.",
	"FACT":         "Returns the factorial of a number.",
	"FACTDOUBLE":   "Returns the double factorial of a number.",
	"FLOOR":        "Rounds a number down to the nearest multiple of a significance.",
	"GAMMA":        "Returns the gamma function value.",
	"GCD":          "Returns the greatest common divisor.",
	"INT":          "Rounds a number down to the nearest integer.",
	"LCM":          "Returns the least common multiple.",
	"LN":           "Returns the natural logarithm of a number.",
	"LOG":          "Returns the logarithm of a number in a chosen base.",
	"LOG10":        "Returns the base-10 logarithm of a number.",
	"MDETERM":      "Returns the determinant of a matrix.",
	"MINVERSE":     "Returns the inverse of a matrix.",
	"MMULT":        "Returns the matrix product of two arrays.",
	"MROUND":       "Rounds a number to the nearest multiple.",
	"MULTINOMIAL":  "Returns the multinomial of a set of numbers.",
	"MUNIT":        "Returns a unit matrix of a requested size.",
	"ODD":          "Rounds a number away from zero to the nearest odd integer.",
	"PERMUT":       "Returns the number of permutations for a number of objects.",
	"PERMUTATIONA": "Returns the number of permutations with repetitions.",
	"PHI":          "Returns the standard normal density at a value.",
	"PI":           "Returns the value of pi.",
	"POWER":        "Returns a number raised to a power.",
	"PRODUCT":      "Returns the product of supplied numbers.",
	"QUOTIENT":     "Returns the integer portion of a division.",
	"RADIANS":      "Converts degrees to radians.",
	"RAND":         "Returns a random number between 0 and 1.",
	"RANDARRAY":    "Returns an array of random numbers.",
	"RANDBETWEEN":  "Returns a random integer between two bounds.",
	"ROUND":        "Rounds a number to a requested number of digits.",
	"ROUNDDOWN":    "Rounds a number toward zero.",
	"ROUNDUP":      "Rounds a number away from zero.",
	"SEC":          "Returns the secant of an angle.",
	"SECH":         "Returns the hyperbolic secant of a number.",
	"SEQUENCE":     "Returns a sequence of numbers as an array.",
	"SERIESSUM":    "Returns the sum of a power series.",
	"SIGN":         "Returns the sign of a number.",
	"SIN":          "Returns the sine of an angle.",
	"SINH":         "Returns the hyperbolic sine of a number.",
	"SQRT":         "Returns the square root of a number.",
	"SQRTPI":       "Returns the square root of a number multiplied by pi.",
	"SUMX2MY2":     "Returns the sum of the difference of squares of paired arrays.",
	"SUMX2PY2":     "Returns the sum of the sum of squares of paired arrays.",
	"SUMXMY2":      "Returns the sum of squares of differences of paired arrays.",
	"TAN":          "Returns the tangent of an angle.",
	"TANH":         "Returns the hyperbolic tangent of a number.",
	"TRUNC":        "Truncates a number to an integer or fixed precision.",
	"FALSE":        "Returns the logical value FALSE.",
}

func newFuncInfo(name, category string) funcInfo {
	return funcInfo{
		Name:        name,
		Description: descriptionFor(name, category),
		Category:    category,
	}
}

func descriptionFor(name, category string) string {
	if desc, ok := descriptionMap[name]; ok {
		return desc
	}
	if desc, ok := simpleDescriptionMap[name]; ok {
		return desc
	}
	if desc := descriptionForBaseConversion(name); desc != "" {
		return desc
	}
	if desc := descriptionForBitOperation(name); desc != "" {
		return desc
	}
	if desc := descriptionForBessel(name); desc != "" {
		return desc
	}
	if desc := descriptionForCeilingFloor(name); desc != "" {
		return desc
	}
	if desc := descriptionForDistribution(name); desc != "" {
		return desc
	}
	if desc := descriptionForComplex(name); desc != "" {
		return desc
	}
	return genericDescription(name, category)
}

func descriptionForBaseConversion(name string) string {
	re := regexp.MustCompile(`^(BIN|DEC|HEX|OCT)2(BIN|DEC|HEX|OCT)$`)
	m := re.FindStringSubmatch(name)
	if m == nil {
		return ""
	}
	baseName := map[string]string{
		"BIN": "binary",
		"DEC": "decimal",
		"HEX": "hexadecimal",
		"OCT": "octal",
	}
	return fmt.Sprintf("Converts a %s number to %s.", baseName[m[1]], baseName[m[2]])
}

func descriptionForBitOperation(name string) string {
	switch name {
	case "BITAND":
		return "Returns the bitwise AND of two integers."
	case "BITLSHIFT":
		return "Returns a number shifted left by a requested number of bits."
	case "BITOR":
		return "Returns the bitwise OR of two integers."
	case "BITRSHIFT":
		return "Returns a number shifted right by a requested number of bits."
	case "BITXOR":
		return "Returns the bitwise XOR of two integers."
	default:
		return ""
	}
}

func descriptionForBessel(name string) string {
	switch name {
	case "BESSELI":
		return "Returns the modified Bessel function I_n(x)."
	case "BESSELJ":
		return "Returns the Bessel function J_n(x)."
	case "BESSELK":
		return "Returns the modified Bessel function K_n(x)."
	case "BESSELY":
		return "Returns the Bessel function Y_n(x)."
	default:
		return ""
	}
}

func descriptionForCeilingFloor(name string) string {
	switch name {
	case "CEILING.MATH":
		return "Rounds a number up using Excel's CEILING.MATH rules."
	case "CEILING.PRECISE":
		return "Rounds a number up to the nearest significance, ignoring the sign of the significance."
	case "ISO.CEILING":
		return "Rounds a number up using ISO.CEILING rules."
	case "FLOOR.MATH":
		return "Rounds a number down using Excel's FLOOR.MATH rules."
	case "FLOOR.PRECISE":
		return "Rounds a number down to the nearest significance, ignoring the sign of the significance."
	case "FLOOR":
		return "Rounds a number down to the nearest multiple of a significance."
	default:
		return ""
	}
}

func descriptionForDistribution(name string) string {
	switch name {
	case "BETA.DIST":
		return "Returns the beta distribution."
	case "BETA.INV":
		return "Returns the inverse of the beta cumulative distribution."
	case "BINOM.DIST":
		return "Returns the binomial distribution."
	case "BINOM.DIST.RANGE":
		return "Returns the probability that a binomial result falls within a range."
	case "BINOM.INV":
		return "Returns the smallest value whose binomial cumulative distribution meets a criterion."
	case "CHISQ.DIST":
		return "Returns the chi-square distribution."
	case "CHISQ.DIST.RT":
		return "Returns the right-tailed chi-square probability."
	case "CHISQ.INV":
		return "Returns the inverse of the chi-square cumulative distribution."
	case "CHISQ.INV.RT":
		return "Returns the inverse of the right-tailed chi-square distribution."
	case "CHISQ.TEST":
		return "Returns the result of a chi-square test."
	case "EXPON.DIST":
		return "Returns the exponential distribution."
	case "F.DIST":
		return "Returns the F probability distribution."
	case "F.DIST.RT":
		return "Returns the right-tailed F probability distribution."
	case "F.INV":
		return "Returns the inverse of the F cumulative distribution."
	case "F.INV.RT":
		return "Returns the inverse of the right-tailed F distribution."
	case "GAMMA.DIST":
		return "Returns the gamma distribution."
	case "GAMMA.INV":
		return "Returns the inverse of the gamma cumulative distribution."
	case "GAMMALN":
		return "Returns the natural logarithm of the gamma function."
	case "GAMMALN.PRECISE":
		return "Returns the natural logarithm of the gamma function using the precise definition."
	case "HYPGEOM.DIST":
		return "Returns the hypergeometric distribution."
	case "LOGNORM.DIST":
		return "Returns the lognormal distribution."
	case "LOGNORM.INV":
		return "Returns the inverse of the lognormal cumulative distribution."
	case "NEGBINOM.DIST":
		return "Returns the negative binomial distribution."
	case "NORM.DIST":
		return "Returns the normal distribution."
	case "NORM.INV":
		return "Returns the inverse of the normal cumulative distribution."
	case "NORM.S.DIST":
		return "Returns the standard normal distribution."
	case "NORM.S.INV":
		return "Returns the inverse of the standard normal cumulative distribution."
	case "POISSON.DIST":
		return "Returns the Poisson distribution."
	case "T.DIST":
		return "Returns the Student's t-distribution."
	case "T.DIST.2T":
		return "Returns the two-tailed Student's t-distribution."
	case "T.DIST.RT":
		return "Returns the right-tailed Student's t-distribution."
	case "T.INV":
		return "Returns the inverse of the Student's t cumulative distribution."
	case "T.INV.2T":
		return "Returns the inverse of the two-tailed Student's t-distribution."
	case "T.TEST":
		return "Returns the probability associated with a Student's t-test."
	case "WEIBULL.DIST":
		return "Returns the Weibull distribution."
	default:
		return ""
	}
}

func descriptionForComplex(name string) string {
	switch name {
	case "IMABS":
		return "Returns the absolute value of a complex number."
	case "IMAGINARY":
		return "Returns the imaginary coefficient of a complex number."
	case "IMARGUMENT":
		return "Returns the argument of a complex number."
	case "IMCONJUGATE":
		return "Returns the complex conjugate of a complex number."
	case "IMCOS":
		return "Returns the cosine of a complex number."
	case "IMCOSH":
		return "Returns the hyperbolic cosine of a complex number."
	case "IMCOT":
		return "Returns the cotangent of a complex number."
	case "IMCSC":
		return "Returns the cosecant of a complex number."
	case "IMCSCH":
		return "Returns the hyperbolic cosecant of a complex number."
	case "IMDIV":
		return "Returns the quotient of two complex numbers."
	case "IMEXP":
		return "Returns e raised to a complex power."
	case "IMLN":
		return "Returns the natural logarithm of a complex number."
	case "IMLOG10":
		return "Returns the base-10 logarithm of a complex number."
	case "IMLOG2":
		return "Returns the base-2 logarithm of a complex number."
	case "IMPOWER":
		return "Raises a complex number to a power."
	case "IMPRODUCT":
		return "Returns the product of complex numbers."
	case "IMREAL":
		return "Returns the real coefficient of a complex number."
	case "IMSEC":
		return "Returns the secant of a complex number."
	case "IMSECH":
		return "Returns the hyperbolic secant of a complex number."
	case "IMSIN":
		return "Returns the sine of a complex number."
	case "IMSINH":
		return "Returns the hyperbolic sine of a complex number."
	case "IMSQRT":
		return "Returns the square root of a complex number."
	case "IMSUB":
		return "Subtracts one complex number from another."
	case "IMSUM":
		return "Returns the sum of complex numbers."
	case "IMTAN":
		return "Returns the tangent of a complex number."
	default:
		return ""
	}
}

func genericDescription(name, category string) string {
	switch category {
	case "Date & Time":
		return "Returns a date or time calculation related to " + name + "."
	case "Engineering":
		return "Returns an engineering result related to " + name + "."
	case "Financial":
		return "Returns a financial calculation related to " + name + "."
	case "Information":
		return "Returns information related to " + name + "."
	case "Logical":
		return "Returns a logical result related to " + name + "."
	case "Lookup & Reference":
		return "Returns a lookup or reference result related to " + name + "."
	case "Math & Trig":
		return "Returns a mathematical result related to " + name + "."
	case "Statistical":
		return "Returns a statistical result related to " + name + "."
	case "Text":
		return "Returns a text result related to " + name + "."
	default:
		return "Returns a result related to " + name + "."
	}
}

// discoverSupported scans formula/*.go (non-test) for Register("NAME", ...) calls
// and includes parser-level functions.
func discoverSupported() map[string]funcInfo {
	re := regexp.MustCompile(`Register\("([A-Z][A-Z0-9.]*)"`)
	result := map[string]funcInfo{}

	// Include parser-level functions
	for name, category := range parserSupported {
		result[name] = newFuncInfo(name, category)
	}

	files, _ := filepath.Glob("formula/functions_*.go")
	for _, path := range files {
		if strings.HasSuffix(path, "_test.go") {
			continue
		}
		// Extract category from filename: functions_math.go -> "math"
		base := filepath.Base(path)
		cat := strings.TrimPrefix(base, "functions_")
		cat = strings.TrimSuffix(cat, ".go")
		displayCat := categoryMap[cat]
		if displayCat == "" {
			displayCat = strings.ToUpper(cat[:1]) + cat[1:]
		}

		f, err := os.Open(path)
		if err != nil {
			continue
		}
		scanner := bufio.NewScanner(f)
		for scanner.Scan() {
			m := re.FindStringSubmatch(scanner.Text())
			if m != nil {
				result[m[1]] = newFuncInfo(m[1], displayCat)
			}
		}
		f.Close()
	}
	return result
}

// countTests counts test cases per formula function.
func countTests(supported map[string]funcInfo) map[string]int {
	counts := map[string]int{}

	// Strategy 1: formula strings in test files
	countFormulaStrings(counts, supported)

	// Strategy 2: TestFUNC_Name test function names
	countTestFunctions(counts, supported)

	// Strategy 3: count t.Run sub-tests inside TestFUNCNAME functions for
	// functions not yet counted (e.g. TRUE, FALSE, MODE — which are either
	// skipped by strategy 1 or use dynamic formula construction).
	countSubTests(counts, supported)

	return counts
}

func countFormulaStrings(counts map[string]int, supported map[string]funcInfo) {
	reFormula := regexp.MustCompile("[\"` ]{1}=?([A-Z][A-Z0-9.]*)" + `\(`)

	skip := map[string]bool{
		"TRUE": true, "FALSE": true, "DIV": true,
		"REF": true, "NULL": true, "NAME": true,
	}

	testFiles, _ := filepath.Glob("formula/*_test.go")
	rootFiles, _ := filepath.Glob("*_test.go")
	testFiles = append(testFiles, rootFiles...)

	for _, path := range testFiles {
		f, err := os.Open(path)
		if err != nil {
			continue
		}
		scanner := bufio.NewScanner(f)
		for scanner.Scan() {
			line := scanner.Text()
			trimmed := strings.TrimSpace(line)
			if strings.HasPrefix(trimmed, "//") {
				continue
			}

			matches := reFormula.FindAllStringSubmatch(trimmed, -1)
			if len(matches) == 0 {
				continue
			}

			isTestCase := strings.Contains(trimmed, "evalCompile") ||
				strings.Contains(trimmed, `formula:`) ||
				strings.Contains(trimmed, `formula :`) ||
				strings.Contains(trimmed, "SetFormula") ||
				(strings.HasPrefix(trimmed, "{") && (strings.Contains(trimmed, `"`) || strings.Contains(trimmed, "`"))) ||
				strings.Contains(trimmed, `name:`)

			if !isTestCase {
				continue
			}

			funcName := matches[0][1]
			if skip[funcName] || !isSupported(funcName, supported) {
				continue
			}
			counts[funcName]++
		}
		f.Close()
	}
}

func countTestFunctions(counts map[string]int, supported map[string]funcInfo) {
	reTestFunc := regexp.MustCompile(`^func\s+Test([A-Z][A-Z0-9_]*?)_\w+\s*\(`)

	testFiles, _ := filepath.Glob("formula/*_test.go")
	for _, path := range testFiles {
		f, err := os.Open(path)
		if err != nil {
			continue
		}
		scanner := bufio.NewScanner(f)
		for scanner.Scan() {
			m := reTestFunc.FindStringSubmatch(scanner.Text())
			if m == nil {
				continue
			}
			funcName := m[1]
			if isSupported(funcName, supported) {
				counts[funcName]++
				continue
			}
			dotted := strings.ReplaceAll(funcName, "_", ".")
			if isSupported(dotted, supported) {
				counts[dotted]++
			}
		}
		f.Close()
	}
}

// countSubTests finds TestFUNCNAME functions (no underscore suffix) and counts
// test cases inside them. This catches functions skipped by strategy 1
// (TRUE, FALSE) and functions that build formula strings dynamically (MODE).
// Only adds counts for functions that have zero tests from earlier strategies.
//
// It counts test cases by looking for:
//   - Table-driven test entries: lines starting with { that contain a string literal
//   - Named t.Run calls with string-literal names that directly test (not containers)
//
// A t.Run is considered a "container" (not counted) if its body includes a
// []struct test table or a for-range loop. Only leaf t.Run calls count.
//
// For functions tested together (e.g. MODE and MODE.SNGL in a single TestMODE),
// it also detects for-loop multipliers like: for _, fn := range []string{"MODE", "MODE.SNGL"}
func countSubTests(counts map[string]int, supported map[string]funcInfo) {
	reFunc := regexp.MustCompile(`^func\s+Test([A-Z][A-Z0-9]*)\s*\(`)
	// Matches for loops that iterate over function names to test multiple aliases
	reFuncLoop := regexp.MustCompile(`for\s+.+range\s+\[\]string\{([^}]+)\}`)

	testFiles, _ := filepath.Glob("formula/*_test.go")
	for _, path := range testFiles {
		data, err := os.ReadFile(path)
		if err != nil {
			continue
		}
		lines := strings.Split(string(data), "\n")

		for i := 0; i < len(lines); i++ {
			m := reFunc.FindStringSubmatch(lines[i])
			if m == nil {
				continue
			}
			testName := m[1]

			// Resolve to a supported function name
			funcName := ""
			if isSupported(testName, supported) {
				funcName = testName
			} else {
				dotted := strings.ReplaceAll(testName, "_", ".")
				if isSupported(dotted, supported) {
					funcName = dotted
				}
			}
			if funcName == "" {
				continue
			}

			// Only fill in if strategy 1+2 found nothing
			if counts[funcName] > 0 {
				continue
			}

			// Find the end of the function body using brace depth
			depth := 0
			started := false
			endLine := len(lines) - 1
			for j := i; j < len(lines); j++ {
				line := lines[j]
				depth += strings.Count(line, "{") - strings.Count(line, "}")
				if !started && depth > 0 {
					started = true
				}
				if started && depth <= 0 {
					endLine = j
					break
				}
			}

			body := lines[i : endLine+1]
			bodyStr := strings.Join(body, "\n")

			// Check if there's a function-name loop multiplier
			extraFuncs := []string{}
			if lm := reFuncLoop.FindStringSubmatch(bodyStr); lm != nil {
				names := strings.Split(lm[1], ",")
				for _, n := range names {
					n = strings.TrimSpace(n)
					n = strings.Trim(n, `"`)
					n = strings.Trim(n, "`")
					if n != "" && n != funcName {
						extraFuncs = append(extraFuncs, n)
					}
				}
			}

			total := countTestCasesInBody(body)
			if total > 0 {
				counts[funcName] += total
				for _, extra := range extraFuncs {
					if isSupported(extra, supported) && counts[extra] == 0 {
						counts[extra] += total
					}
				}
			}
		}
	}
}

// countTestCasesInBody counts test cases in a function body by walking through
// t.Run calls and test table entries. It distinguishes between:
//   - Container t.Run calls (those wrapping a test table + for loop): not counted,
//     but their table entries ARE counted
//   - Leaf t.Run calls (those containing direct test logic): counted as 1 each
//   - Table entries ({...} lines inside a []struct literal): counted as 1 each
func countTestCasesInBody(lines []string) int {
	reNamedRun := regexp.MustCompile(`t\.Run\(["\x60]`)
	reTableEntry := regexp.MustCompile(`^\s*\{["\x60]`)
	reStructDecl := regexp.MustCompile(`\[\]struct`)

	total := 0

	for idx := 0; idx < len(lines); idx++ {
		trimmed := strings.TrimSpace(lines[idx])
		if strings.HasPrefix(trimmed, "//") {
			continue
		}

		// When we find a t.Run with a string literal name, check if it's
		// a container (has a []struct inside) or a leaf test.
		if reNamedRun.MatchString(trimmed) {
			// Find the extent of this t.Run block
			blockStart := idx
			depth := 0
			blockStarted := false
			blockEnd := idx

			for j := blockStart; j < len(lines); j++ {
				line := lines[j]
				depth += strings.Count(line, "{") - strings.Count(line, "}")
				if !blockStarted && depth > 0 {
					blockStarted = true
				}
				if blockStarted && depth <= 0 {
					blockEnd = j
					break
				}
			}

			// Check if this block contains a []struct (table-driven test container)
			blockLines := lines[blockStart : blockEnd+1]
			isContainer := false
			for _, bl := range blockLines {
				if reStructDecl.MatchString(bl) {
					isContainer = true
					break
				}
			}

			if isContainer {
				// Count table entries inside this container
				for _, bl := range blockLines {
					bt := strings.TrimSpace(bl)
					if strings.HasPrefix(bt, "//") {
						continue
					}
					if reTableEntry.MatchString(bt) {
						total++
					}
				}
			} else {
				// Leaf t.Run — counts as one test case
				total++
			}

			// Skip past this block
			idx = blockEnd
			continue
		}
	}

	return total
}

func isSupported(name string, supported map[string]funcInfo) bool {
	_, ok := supported[name]
	return ok
}

func markdownRow(line string) []string {
	if !strings.HasPrefix(line, "|") {
		return nil
	}
	parts := strings.Split(line, "|")
	if len(parts) < 3 {
		return nil
	}
	var cells []string
	for _, part := range parts[1 : len(parts)-1] {
		cells = append(cells, strings.TrimSpace(part))
	}
	return cells
}

func markdownCell(s string) string {
	s = strings.TrimSpace(s)
	s = strings.ReplaceAll(s, "|", "\\|")
	s = strings.Join(strings.Fields(s), " ")
	return s
}

// computeUnsupported reads the existing FORMULAS.md unsupported list, removes
// any that are now supported, and ensures all willNotSupport entries are included.
func computeUnsupported(supported map[string]funcInfo) []funcInfo {
	seen := map[string]bool{}
	var unsupported []funcInfo

	// Always include willNotSupport entries (they are the source of truth)
	for name := range willNotSupport {
		if _, ok := supported[name]; ok {
			continue
		}
		seen[name] = true
		unsupported = append(unsupported, newFuncInfo(name, categoryForUnsupported(name)))
	}

	// Read remaining unsupported from existing FORMULAS.md
	f, err := os.Open("FORMULAS.md")
	if err != nil {
		return unsupported
	}
	defer f.Close()

	inUnsupported := false
	section := ""

	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		line := scanner.Text()
		if strings.HasPrefix(line, "# Unsupported") ||
			strings.HasPrefix(line, "# No Planned Support") ||
			strings.HasPrefix(line, "# Not Yet Implemented") {
			inUnsupported = true
			switch {
			case strings.HasPrefix(line, "# No Planned Support"):
				section = "wont"
			case strings.HasPrefix(line, "# Not Yet Implemented"):
				section = "notyet"
			default:
				section = ""
			}
			continue
		}
		if strings.HasPrefix(line, "# ") && inUnsupported {
			section = ""
			break
		}
		if !inUnsupported {
			continue
		}
		cells := markdownRow(line)
		if len(cells) == 0 {
			continue
		}
		name := cells[0]
		if matched, _ := regexp.MatchString(`^[A-Z][A-Z0-9.]*$`, name); !matched {
			continue
		}
		cat := ""
		desc := ""
		switch section {
		case "wont":
			if len(cells) >= 4 {
				desc = cells[1]
				cat = cells[2]
			} else if len(cells) >= 3 {
				cat = cells[1]
			}
		case "notyet":
			if len(cells) >= 3 {
				desc = cells[1]
				cat = cells[2]
			} else if len(cells) >= 2 {
				cat = cells[1]
			}
		default:
			continue
		}
		if cat == "" {
			cat = categoryForUnsupported(name)
		}
		if _, ok := supported[name]; ok {
			continue
		}
		if seen[name] {
			continue
		}
		seen[name] = true
		fi := newFuncInfo(name, cat)
		if fi.Description == "" {
			fi.Description = desc
		}
		unsupported = append(unsupported, fi)
	}
	return unsupported
}

// categoryForUnsupported returns the category for a known unsupported function.
var unsupportedCategories = map[string]string{
	"WEBSERVICE": "Web",
	"FILTERXML":  "Web",
	"INFO":       "Information",
	"CELL":       "Information",
	"PHONETIC":   "Text",
	"ASC":        "Text",
	"DBCS":       "Text",
	"BAHTTEXT":   "Text",
	"FINDB":      "Text",
	"LEFTB":      "Text",
	"LENB":       "Text",
	"MIDB":       "Text",
	"REPLACEB":   "Text",
	"RIGHTB":     "Text",
	"SEARCHB":    "Text",
}

func categoryForUnsupported(name string) string {
	if cat, ok := unsupportedCategories[name]; ok {
		return cat
	}
	return "Unknown"
}

func generateMarkdown(supported map[string]funcInfo, unsupported []funcInfo) string {
	var b strings.Builder

	// Sort supported by name
	var funcs []funcInfo
	for _, fi := range supported {
		funcs = append(funcs, fi)
	}
	sort.Slice(funcs, func(i, j int) bool {
		return funcs[i].Name < funcs[j].Name
	})

	b.WriteString("# Supported Formulas\n\n")
	b.WriteString(fmt.Sprintf("Werkbook supports **%d** spreadsheet formula functions.\n\n", len(funcs)))
	b.WriteString("| Function | Description | Category | Tests |\n")
	b.WriteString("|----------|-------------|----------|------:|\n")
	for _, fi := range funcs {
		tests := "-"
		if fi.Tests > 0 {
			tests = strconv.Itoa(fi.Tests)
		}
		b.WriteString(fmt.Sprintf("| %s | %s | %s | %s |\n",
			fi.Name,
			markdownCell(fi.Description),
			fi.Category,
			tests,
		))
	}

	// Split unsupported into "will not support" and "not yet implemented"
	var wontSupport, notYet []funcInfo
	for _, fi := range unsupported {
		if _, ok := willNotSupport[fi.Name]; ok {
			wontSupport = append(wontSupport, fi)
		} else {
			notYet = append(notYet, fi)
		}
	}

	if len(wontSupport) > 0 {
		sort.Slice(wontSupport, func(i, j int) bool {
			return wontSupport[i].Name < wontSupport[j].Name
		})
		b.WriteString("\n# No Planned Support\n\n")
		b.WriteString("These functions depend on the spreadsheet application's runtime environment, have side\n")
		b.WriteString("effects, or require locale-specific behavior that cannot be reproduced in a server-side library.\n\n")
		b.WriteString("| Function | Description | Category | Reason |\n")
		b.WriteString("|----------|-------------|----------|--------|\n")
		for _, fi := range wontSupport {
			reason := willNotSupport[fi.Name]
			b.WriteString(fmt.Sprintf("| %s | %s | %s | %s |\n",
				fi.Name,
				markdownCell(fi.Description),
				fi.Category,
				markdownCell(reason),
			))
		}
	}

	if len(notYet) > 0 {
		sort.Slice(notYet, func(i, j int) bool {
			return notYet[i].Name < notYet[j].Name
		})
		b.WriteString(fmt.Sprintf("\n# Not Yet Implemented\n\n"))
		b.WriteString(fmt.Sprintf("The following **%d** functions are not yet supported.\n\n", len(notYet)))
		b.WriteString("| Function | Description | Category |\n")
		b.WriteString("|----------|-------------|----------|\n")
		for _, fi := range notYet {
			b.WriteString(fmt.Sprintf("| %s | %s | %s |\n",
				fi.Name,
				markdownCell(fi.Description),
				fi.Category,
			))
		}
	}

	return b.String()
}
