// Script to regenerate FORMULAS.md from the source of truth:
//  1. Supported functions = all Register* helper calls in formula/*.go
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
	// OLAP cube functions — require a live connection to an Analysis Services cube.
	"CUBEKPIMEMBER":      "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	"CUBEMEMBER":         "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	"CUBEMEMBERPROPERTY": "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	"CUBERANKEDMEMBER":   "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	"CUBESET":            "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	"CUBESETCOUNT":       "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	"CUBEVALUE":          "Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine",
	// External-service / runtime-environment dependent.
	"RTD":            "Connects to a real-time COM automation server — requires running external program",
	"IMAGE":          "Returns an image rendered from a URL — werkbook is a calc engine, not a renderer",
	"STOCKHISTORY":   "Fetches live financial market data via Microsoft service — external network dependency",
	"DETECTLANGUAGE": "Calls Microsoft's online translation/ML service — external network dependency",
	"TRANSLATE":      "Calls Microsoft's online translation service — external network dependency",
	// DLL / VBA-era extension functions — security risk and platform-specific binary loading.
	"CALL":        "Calls procedures in dynamic-link libraries — security risk, OS-specific binary loading",
	"REGISTER.ID": "Returns the register ID of a DLL or code resource — paired with CALL, OS-specific",
	"EUROCONVERT": "Excel add-in tied to a Euro currency conversion table — locale and add-in-version specific",
}

// allExcelFunctions is the universe of Excel functions sourced from
// Microsoft's official 'Excel functions by category' reference page.
// The set of 513 functions is the source of truth for what
// FORMULAS.md may classify as 'Supported', 'No Planned Support', or
// 'Not Yet Implemented'. Functions in this map but not registered and
// not in willNotSupport are emitted under the # Not Yet Implemented
// section by the generator.
//
// To refresh: re-fetch the page and update this map. Microsoft adds
// new functions periodically (e.g. GROUPBY/PIVOTBY/PERCENTOF in 2024).
var allExcelFunctions = map[string]funcInfo{
	"ABS": {Name: "ABS", Category: "Math & Trig", Description: "Returns the absolute value of a number"},
	"ACCRINT": {Name: "ACCRINT", Category: "Financial", Description: "Returns the accrued interest for a security that pays periodic interest"},
	"ACCRINTM": {Name: "ACCRINTM", Category: "Financial", Description: "Returns the accrued interest for a security that pays interest at maturity"},
	"ACOS": {Name: "ACOS", Category: "Math & Trig", Description: "Returns the arccosine of a number"},
	"ACOSH": {Name: "ACOSH", Category: "Math & Trig", Description: "Returns the inverse hyperbolic cosine of a number"},
	"ACOT": {Name: "ACOT", Category: "Math & Trig", Description: "Returns the arccotangent of a number"},
	"ACOTH": {Name: "ACOTH", Category: "Math & Trig", Description: "Returns the hyperbolic arccotangent of a number"},
	"ADDRESS": {Name: "ADDRESS", Category: "Lookup & Reference", Description: "Returns a reference as text to a single cell in a worksheet"},
	"AGGREGATE": {Name: "AGGREGATE", Category: "Math & Trig", Description: "Returns an aggregate in a list or database"},
	"AMORDEGRC": {Name: "AMORDEGRC", Category: "Financial", Description: "Returns the depreciation for each accounting period by using a depreciation coefficient"},
	"AMORLINC": {Name: "AMORLINC", Category: "Financial", Description: "Returns the depreciation for each accounting period"},
	"AND": {Name: "AND", Category: "Logical", Description: "Returns TRUE if all of its arguments are TRUE"},
	"ARABIC": {Name: "ARABIC", Category: "Math & Trig", Description: "Converts a Roman number to Arabic, as a number"},
	"AREAS": {Name: "AREAS", Category: "Lookup & Reference", Description: "Returns the number of areas in a reference"},
	"ARRAYTOTEXT": {Name: "ARRAYTOTEXT", Category: "Text", Description: "Returns an array of text values from any specified range"},
	"ASC": {Name: "ASC", Category: "Text", Description: "Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters"},
	"ASIN": {Name: "ASIN", Category: "Math & Trig", Description: "Returns the arcsine of a number"},
	"ASINH": {Name: "ASINH", Category: "Math & Trig", Description: "Returns the inverse hyperbolic sine of a number"},
	"ATAN": {Name: "ATAN", Category: "Math & Trig", Description: "Returns the arctangent of a number"},
	"ATAN2": {Name: "ATAN2", Category: "Math & Trig", Description: "Returns the arctangent from x- and y-coordinates"},
	"ATANH": {Name: "ATANH", Category: "Math & Trig", Description: "Returns the inverse hyperbolic tangent of a number"},
	"AVEDEV": {Name: "AVEDEV", Category: "Statistical", Description: "Returns the average of the absolute deviations of data points from their mean"},
	"AVERAGE": {Name: "AVERAGE", Category: "Statistical", Description: "Returns the average of its arguments"},
	"AVERAGEA": {Name: "AVERAGEA", Category: "Statistical", Description: "Returns the average of its arguments, including numbers, text, and logical values"},
	"AVERAGEIF": {Name: "AVERAGEIF", Category: "Statistical", Description: "Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria"},
	"AVERAGEIFS": {Name: "AVERAGEIFS", Category: "Statistical", Description: "Returns the average (arithmetic mean) of all cells that meet multiple criteria"},
	"BAHTTEXT": {Name: "BAHTTEXT", Category: "Text", Description: "Converts a number to text, using the baht currency format"},
	"BASE": {Name: "BASE", Category: "Math & Trig", Description: "Converts a number into a text representation with the given radix (base)"},
	"BESSELI": {Name: "BESSELI", Category: "Engineering", Description: "Returns the modified Bessel function In(x)"},
	"BESSELJ": {Name: "BESSELJ", Category: "Engineering", Description: "Returns the Bessel function Jn(x)"},
	"BESSELK": {Name: "BESSELK", Category: "Engineering", Description: "Returns the modified Bessel function Kn(x)"},
	"BESSELY": {Name: "BESSELY", Category: "Engineering", Description: "Returns the Bessel function Yn(x)"},
	"BETA.DIST": {Name: "BETA.DIST", Category: "Statistical", Description: "Returns the beta cumulative distribution function"},
	"BETA.INV": {Name: "BETA.INV", Category: "Statistical", Description: "Returns the inverse of the cumulative distribution function for a specified beta distribution"},
	"BETADIST": {Name: "BETADIST", Category: "Statistical", Description: "Returns the beta cumulative distribution function"},
	"BETAINV": {Name: "BETAINV", Category: "Statistical", Description: "Returns the inverse of the cumulative distribution function for a specified beta distribution"},
	"BIN2DEC": {Name: "BIN2DEC", Category: "Engineering", Description: "Converts a binary number to decimal"},
	"BIN2HEX": {Name: "BIN2HEX", Category: "Engineering", Description: "Converts a binary number to hexadecimal"},
	"BIN2OCT": {Name: "BIN2OCT", Category: "Engineering", Description: "Converts a binary number to octal"},
	"BINOM.DIST": {Name: "BINOM.DIST", Category: "Statistical", Description: "Returns the individual term binomial distribution probability"},
	"BINOM.DIST.RANGE": {Name: "BINOM.DIST.RANGE", Category: "Statistical", Description: "Returns the probability of a trial result using a binomial distribution"},
	"BINOM.INV": {Name: "BINOM.INV", Category: "Statistical", Description: "Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value"},
	"BINOMDIST": {Name: "BINOMDIST", Category: "Statistical", Description: "Returns the individual term binomial distribution probability"},
	"BITAND": {Name: "BITAND", Category: "Engineering", Description: "Returns a 'Bitwise And' of two numbers"},
	"BITLSHIFT": {Name: "BITLSHIFT", Category: "Engineering", Description: "Returns a value number shifted left by shift_amount bits"},
	"BITOR": {Name: "BITOR", Category: "Engineering", Description: "Returns a bitwise OR of 2 numbers"},
	"BITRSHIFT": {Name: "BITRSHIFT", Category: "Engineering", Description: "Returns a value number shifted right by shift_amount bits"},
	"BITXOR": {Name: "BITXOR", Category: "Engineering", Description: "Returns a bitwise 'Exclusive Or' of two numbers"},
	"BYCOL": {Name: "BYCOL", Category: "Logical", Description: "Applies a LAMBDA to each column and returns an array of the results"},
	"BYROW": {Name: "BYROW", Category: "Logical", Description: "Applies a LAMBDA to each row and returns an array of the results"},
	"CALL": {Name: "CALL", Category: "User Defined", Description: "Calls a procedure in a dynamic link library or code resource"},
	"CEILING": {Name: "CEILING", Category: "Math & Trig", Description: "Rounds a number to the nearest integer or to the nearest multiple of significance"},
	"CEILING.MATH": {Name: "CEILING.MATH", Category: "Math & Trig", Description: "Rounds a number up, to the nearest integer or to the nearest multiple of significance"},
	"CEILING.PRECISE": {Name: "CEILING.PRECISE", Category: "Math & Trig", Description: "Rounds a number the nearest integer or to the nearest multiple of significance. Regardless of the sign of the number, the number is rounded up."},
	"CELL": {Name: "CELL", Category: "Information", Description: "Returns information about the formatting, location, or contents of a cell"},
	"CHAR": {Name: "CHAR", Category: "Text", Description: "Returns the character specified by the code number"},
	"CHIDIST": {Name: "CHIDIST", Category: "Statistical", Description: "Returns the one-tailed probability of the chi-squared distribution"},
	"CHIINV": {Name: "CHIINV", Category: "Statistical", Description: "Returns the inverse of the one-tailed probability of the chi-squared distribution"},
	"CHISQ.DIST": {Name: "CHISQ.DIST", Category: "Statistical", Description: "Returns the cumulative beta probability density function"},
	"CHISQ.DIST.RT": {Name: "CHISQ.DIST.RT", Category: "Statistical", Description: "Returns the one-tailed probability of the chi-squared distribution"},
	"CHISQ.INV": {Name: "CHISQ.INV", Category: "Statistical", Description: "Returns the cumulative beta probability density function"},
	"CHISQ.INV.RT": {Name: "CHISQ.INV.RT", Category: "Statistical", Description: "Returns the inverse of the one-tailed probability of the chi-squared distribution"},
	"CHISQ.TEST": {Name: "CHISQ.TEST", Category: "Statistical", Description: "Returns the test for independence"},
	"CHITEST": {Name: "CHITEST", Category: "Statistical", Description: "Returns the test for independence"},
	"CHOOSE": {Name: "CHOOSE", Category: "Lookup & Reference", Description: "Chooses a value from a list of values"},
	"CHOOSECOLS": {Name: "CHOOSECOLS", Category: "Lookup & Reference", Description: "Returns the specified columns from an array"},
	"CHOOSEROWS": {Name: "CHOOSEROWS", Category: "Lookup & Reference", Description: "Returns the specified rows from an array"},
	"CLEAN": {Name: "CLEAN", Category: "Text", Description: "Removes all nonprintable characters from text"},
	"CODE": {Name: "CODE", Category: "Text", Description: "Returns a numeric code for the first character in a text string"},
	"COLUMN": {Name: "COLUMN", Category: "Lookup & Reference", Description: "Returns the column number of a reference"},
	"COLUMNS": {Name: "COLUMNS", Category: "Lookup & Reference", Description: "Returns the number of columns in a reference"},
	"COMBIN": {Name: "COMBIN", Category: "Math & Trig", Description: "Returns the number of combinations for a given number of objects"},
	"COMBINA": {Name: "COMBINA", Category: "Math & Trig", Description: "Returns the number of combinations with repetitions for a given number of items"},
	"COMPLEX": {Name: "COMPLEX", Category: "Engineering", Description: "Converts real and imaginary coefficients into a complex number"},
	"CONCAT": {Name: "CONCAT", Category: "Text", Description: "Combines the text from multiple ranges and/or strings, but it doesn't provide the delimiter or IgnoreEmpty arguments."},
	"CONCATENATE": {Name: "CONCATENATE", Category: "Text", Description: "Joins several text items into one text item"},
	"CONFIDENCE": {Name: "CONFIDENCE", Category: "Statistical", Description: "Returns the confidence interval for a population mean"},
	"CONFIDENCE.NORM": {Name: "CONFIDENCE.NORM", Category: "Statistical", Description: "Returns the confidence interval for a population mean"},
	"CONFIDENCE.T": {Name: "CONFIDENCE.T", Category: "Statistical", Description: "Returns the confidence interval for a population mean, using a Student's t distribution"},
	"CONVERT": {Name: "CONVERT", Category: "Engineering", Description: "Converts a number from one measurement system to another"},
	"CORREL": {Name: "CORREL", Category: "Statistical", Description: "Returns the correlation coefficient between two data sets"},
	"COS": {Name: "COS", Category: "Math & Trig", Description: "Returns the cosine of a number"},
	"COSH": {Name: "COSH", Category: "Math & Trig", Description: "Returns the hyperbolic cosine of a number"},
	"COT": {Name: "COT", Category: "Math & Trig", Description: "Returns the cotangent of an angle"},
	"COTH": {Name: "COTH", Category: "Math & Trig", Description: "Returns the hyperbolic cotangent of a number"},
	"COUNT": {Name: "COUNT", Category: "Statistical", Description: "Counts how many numbers are in the list of arguments"},
	"COUNTA": {Name: "COUNTA", Category: "Statistical", Description: "Counts how many values are in the list of arguments"},
	"COUNTBLANK": {Name: "COUNTBLANK", Category: "Statistical", Description: "Counts the number of blank cells within a range"},
	"COUNTIF": {Name: "COUNTIF", Category: "Statistical", Description: "Counts the number of cells within a range that meet the given criteria"},
	"COUNTIFS": {Name: "COUNTIFS", Category: "Statistical", Description: "Counts the number of cells within a range that meet multiple criteria"},
	"COUPDAYBS": {Name: "COUPDAYBS", Category: "Financial", Description: "Returns the number of days from the beginning of the coupon period to the settlement date"},
	"COUPDAYS": {Name: "COUPDAYS", Category: "Financial", Description: "Returns the number of days in the coupon period that contains the settlement date"},
	"COUPDAYSNC": {Name: "COUPDAYSNC", Category: "Financial", Description: "Returns the number of days from the settlement date to the next coupon date"},
	"COUPNCD": {Name: "COUPNCD", Category: "Financial", Description: "Returns the next coupon date after the settlement date"},
	"COUPNUM": {Name: "COUPNUM", Category: "Financial", Description: "Returns the number of coupons payable between the settlement date and maturity date"},
	"COUPPCD": {Name: "COUPPCD", Category: "Financial", Description: "Returns the previous coupon date before the settlement date"},
	"COVAR": {Name: "COVAR", Category: "Statistical", Description: "Returns covariance, the average of the products of paired deviations"},
	"COVARIANCE.P": {Name: "COVARIANCE.P", Category: "Statistical", Description: "Returns covariance, the average of the products of paired deviations"},
	"COVARIANCE.S": {Name: "COVARIANCE.S", Category: "Statistical", Description: "Returns the sample covariance, the average of the products deviations for each data point pair in two data sets"},
	"CRITBINOM": {Name: "CRITBINOM", Category: "Statistical", Description: "Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value"},
	"CSC": {Name: "CSC", Category: "Math & Trig", Description: "Returns the cosecant of an angle"},
	"CSCH": {Name: "CSCH", Category: "Math & Trig", Description: "Returns the hyperbolic cosecant of an angle"},
	"CUBEKPIMEMBER": {Name: "CUBEKPIMEMBER", Category: "Cube", Description: "Returns a key performance indicator (KPI) property and displays the KPI name in the cell."},
	"CUBEMEMBER": {Name: "CUBEMEMBER", Category: "Cube", Description: "Returns a member or tuple from the cube."},
	"CUBEMEMBERPROPERTY": {Name: "CUBEMEMBERPROPERTY", Category: "Cube", Description: "Returns the value of a member property from the cube."},
	"CUBERANKEDMEMBER": {Name: "CUBERANKEDMEMBER", Category: "Cube", Description: "Returns the nth, or ranked, member in a set."},
	"CUBESET": {Name: "CUBESET", Category: "Cube", Description: "Defines a calculated set of members or tuples by sending a set expression to the cube on the server."},
	"CUBESETCOUNT": {Name: "CUBESETCOUNT", Category: "Cube", Description: "Returns the number of items in a set."},
	"CUBEVALUE": {Name: "CUBEVALUE", Category: "Cube", Description: "Returns an aggregated value from the cube."},
	"CUMIPMT": {Name: "CUMIPMT", Category: "Financial", Description: "Returns the cumulative interest paid between two periods"},
	"CUMPRINC": {Name: "CUMPRINC", Category: "Financial", Description: "Returns the cumulative principal paid on a loan between two periods"},
	"DATE": {Name: "DATE", Category: "Date & Time", Description: "Returns the serial number of a particular date"},
	"DATEDIF": {Name: "DATEDIF", Category: "Date & Time", Description: "Calculates the number of days, months, or years between two dates."},
	"DATEVALUE": {Name: "DATEVALUE", Category: "Date & Time", Description: "Converts a date in the form of text to a serial number"},
	"DAVERAGE": {Name: "DAVERAGE", Category: "Db", Description: "Returns the average of selected database entries"},
	"DAY": {Name: "DAY", Category: "Date & Time", Description: "Converts a serial number to a day of the month"},
	"DAYS": {Name: "DAYS", Category: "Date & Time", Description: "Returns the number of days between two dates"},
	"DAYS360": {Name: "DAYS360", Category: "Date & Time", Description: "Calculates the number of days between two dates based on a 360-day year"},
	"DB": {Name: "DB", Category: "Financial", Description: "Returns the depreciation of an asset for a specified period by using the fixed-declining balance method"},
	"DBCS": {Name: "DBCS", Category: "Text", Description: "Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters"},
	"DCOUNT": {Name: "DCOUNT", Category: "Db", Description: "Counts the cells that contain numbers in a database"},
	"DCOUNTA": {Name: "DCOUNTA", Category: "Db", Description: "Counts nonblank cells in a database"},
	"DDB": {Name: "DDB", Category: "Financial", Description: "Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify"},
	"DEC2BIN": {Name: "DEC2BIN", Category: "Engineering", Description: "Converts a decimal number to binary"},
	"DEC2HEX": {Name: "DEC2HEX", Category: "Engineering", Description: "Converts a decimal number to hexadecimal"},
	"DEC2OCT": {Name: "DEC2OCT", Category: "Engineering", Description: "Converts a decimal number to octal"},
	"DECIMAL": {Name: "DECIMAL", Category: "Math & Trig", Description: "Converts a text representation of a number in a given base into a decimal number"},
	"DEGREES": {Name: "DEGREES", Category: "Math & Trig", Description: "Converts radians to degrees"},
	"DELTA": {Name: "DELTA", Category: "Engineering", Description: "Tests whether two values are equal"},
	"DETECTLANGUAGE": {Name: "DETECTLANGUAGE", Category: "Text", Description: "Identifies the language of a specified text"},
	"DEVSQ": {Name: "DEVSQ", Category: "Statistical", Description: "Returns the sum of squares of deviations"},
	"DGET": {Name: "DGET", Category: "Db", Description: "Extracts from a database a single record that matches the specified criteria"},
	"DISC": {Name: "DISC", Category: "Financial", Description: "Returns the discount rate for a security"},
	"DMAX": {Name: "DMAX", Category: "Db", Description: "Returns the maximum value from selected database entries"},
	"DMIN": {Name: "DMIN", Category: "Db", Description: "Returns the minimum value from selected database entries"},
	"DOLLAR": {Name: "DOLLAR", Category: "Text", Description: "Converts a number to text, using the dollar currency format"},
	"DOLLARDE": {Name: "DOLLARDE", Category: "Financial", Description: "Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number"},
	"DOLLARFR": {Name: "DOLLARFR", Category: "Financial", Description: "Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction"},
	"DPRODUCT": {Name: "DPRODUCT", Category: "Db", Description: "Multiplies the values in a particular field of records that match the criteria in a database"},
	"DROP": {Name: "DROP", Category: "Lookup & Reference", Description: "Excludes a specified number of rows or columns from the start or end of an array"},
	"DSTDEV": {Name: "DSTDEV", Category: "Db", Description: "Estimates the standard deviation based on a sample of selected database entries"},
	"DSTDEVP": {Name: "DSTDEVP", Category: "Db", Description: "Calculates the standard deviation based on the entire population of selected database entries"},
	"DSUM": {Name: "DSUM", Category: "Db", Description: "Adds the numbers in the field column of records in the database that match the criteria"},
	"DURATION": {Name: "DURATION", Category: "Financial", Description: "Returns the annual duration of a security with periodic interest payments"},
	"DVAR": {Name: "DVAR", Category: "Db", Description: "Estimates variance based on a sample from selected database entries"},
	"DVARP": {Name: "DVARP", Category: "Db", Description: "Calculates variance based on the entire population of selected database entries"},
	"EDATE": {Name: "EDATE", Category: "Date & Time", Description: "Returns the serial number of the date that is the indicated number of months before or after the start date"},
	"EFFECT": {Name: "EFFECT", Category: "Financial", Description: "Returns the effective annual interest rate"},
	"ENCODEURL": {Name: "ENCODEURL", Category: "Web", Description: "Returns a URL-encoded string"},
	"EOMONTH": {Name: "EOMONTH", Category: "Date & Time", Description: "Returns the serial number of the last day of the month before or after a specified number of months"},
	"ERF": {Name: "ERF", Category: "Engineering", Description: "Returns the error function"},
	"ERF.PRECISE": {Name: "ERF.PRECISE", Category: "Engineering", Description: "Returns the error function"},
	"ERFC": {Name: "ERFC", Category: "Engineering", Description: "Returns the complementary error function"},
	"ERFC.PRECISE": {Name: "ERFC.PRECISE", Category: "Engineering", Description: "Returns the complementary ERF function integrated between x and infinity"},
	"ERROR.TYPE": {Name: "ERROR.TYPE", Category: "Information", Description: "Returns a number corresponding to an error type"},
	"EUROCONVERT": {Name: "EUROCONVERT", Category: "User Defined", Description: "Converts a number to euros, converts a number from euros to a euro member currency, or converts a number from one euro member currency to another by using the euro as an intermediary (triangulation)"},
	"EVEN": {Name: "EVEN", Category: "Math & Trig", Description: "Rounds a number up to the nearest even integer"},
	"EXACT": {Name: "EXACT", Category: "Text", Description: "Checks to see if two text values are identical"},
	"EXP": {Name: "EXP", Category: "Math & Trig", Description: "Returns _e_ raised to the power of a given number"},
	"EXPAND": {Name: "EXPAND", Category: "Lookup & Reference", Description: "Expands or pads an array to specified row and column dimensions"},
	"EXPON.DIST": {Name: "EXPON.DIST", Category: "Statistical", Description: "Returns the exponential distribution"},
	"EXPONDIST": {Name: "EXPONDIST", Category: "Statistical", Description: "Returns the exponential distribution"},
	"F.DIST": {Name: "F.DIST", Category: "Statistical", Description: "Returns the F probability distribution"},
	"F.DIST.RT": {Name: "F.DIST.RT", Category: "Statistical", Description: "Returns the F probability distribution"},
	"F.INV": {Name: "F.INV", Category: "Statistical", Description: "Returns the inverse of the F probability distribution"},
	"F.INV.RT": {Name: "F.INV.RT", Category: "Statistical", Description: "Returns the inverse of the F probability distribution"},
	"F.TEST": {Name: "F.TEST", Category: "Statistical", Description: "Returns the result of an F-test"},
	"FACT": {Name: "FACT", Category: "Math & Trig", Description: "Returns the factorial of a number"},
	"FACTDOUBLE": {Name: "FACTDOUBLE", Category: "Math & Trig", Description: "Returns the double factorial of a number"},
	"FALSE": {Name: "FALSE", Category: "Logical", Description: "Returns the logical value FALSE"},
	"FDIST": {Name: "FDIST", Category: "Statistical", Description: "Returns the F probability distribution"},
	"FILTER": {Name: "FILTER", Category: "Lookup & Reference", Description: "Filters a range of data based on criteria you define"},
	"FILTERXML": {Name: "FILTERXML", Category: "Web", Description: "Returns specific data from the XML content by using the specified XPath"},
	"FIND": {Name: "FIND", Category: "Text", Description: "Finds one text value within another (case-sensitive)"},
	"FINV": {Name: "FINV", Category: "Statistical", Description: "Returns the inverse of the F probability distribution"},
	"FISHER": {Name: "FISHER", Category: "Statistical", Description: "Returns the Fisher transformation"},
	"FISHERINV": {Name: "FISHERINV", Category: "Statistical", Description: "Returns the inverse of the Fisher transformation"},
	"FIXED": {Name: "FIXED", Category: "Text", Description: "Formats a number as text with a fixed number of decimals"},
	"FLOOR": {Name: "FLOOR", Category: "Math & Trig", Description: "Rounds a number down, toward zero"},
	"FLOOR.MATH": {Name: "FLOOR.MATH", Category: "Math & Trig", Description: "Rounds a number down, to the nearest integer or to the nearest multiple of significance"},
	"FLOOR.PRECISE": {Name: "FLOOR.PRECISE", Category: "Math & Trig", Description: "Rounds a number down to the nearest integer or to the nearest multiple of significance. Regardless of the sign of the number, the number is rounded down."},
	"FORECAST": {Name: "FORECAST", Category: "Statistical", Description: "Returns a value along a linear trend"},
	"FORECAST.ETS": {Name: "FORECAST.ETS", Category: "Statistical", Description: "Returns a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm"},
	"FORECAST.ETS.CONFINT": {Name: "FORECAST.ETS.CONFINT", Category: "Statistical", Description: "Returns a confidence interval for the forecast value at the specified target date"},
	"FORECAST.ETS.SEASONALITY": {Name: "FORECAST.ETS.SEASONALITY", Category: "Statistical", Description: "Returns the length of the repetitive pattern Excel detects for the specified time series"},
	"FORECAST.ETS.STAT": {Name: "FORECAST.ETS.STAT", Category: "Statistical", Description: "Returns a statistical value as a result of time series forecasting"},
	"FORECAST.LINEAR": {Name: "FORECAST.LINEAR", Category: "Statistical", Description: "Returns a future value based on existing values"},
	"FORMULATEXT": {Name: "FORMULATEXT", Category: "Lookup & Reference", Description: "Returns the formula at the given reference as text"},
	"FREQUENCY": {Name: "FREQUENCY", Category: "Statistical", Description: "Returns a frequency distribution as a vertical array"},
	"FTEST": {Name: "FTEST", Category: "Statistical", Description: "Returns the result of an F-test"},
	"FV": {Name: "FV", Category: "Financial", Description: "Returns the future value of an investment"},
	"FVSCHEDULE": {Name: "FVSCHEDULE", Category: "Financial", Description: "Returns the future value of an initial principal after applying a series of compound interest rates"},
	"GAMMA": {Name: "GAMMA", Category: "Statistical", Description: "Returns the Gamma function value"},
	"GAMMA.DIST": {Name: "GAMMA.DIST", Category: "Statistical", Description: "Returns the gamma distribution"},
	"GAMMA.INV": {Name: "GAMMA.INV", Category: "Statistical", Description: "Returns the inverse of the gamma cumulative distribution"},
	"GAMMADIST": {Name: "GAMMADIST", Category: "Statistical", Description: "Returns the gamma distribution"},
	"GAMMAINV": {Name: "GAMMAINV", Category: "Statistical", Description: "Returns the inverse of the gamma cumulative distribution"},
	"GAMMALN": {Name: "GAMMALN", Category: "Statistical", Description: "Returns the natural logarithm of the gamma function"},
	"GAMMALN.PRECISE": {Name: "GAMMALN.PRECISE", Category: "Statistical", Description: "Returns the natural logarithm of the gamma function"},
	"GAUSS": {Name: "GAUSS", Category: "Statistical", Description: "Returns 0.5 less than the standard normal cumulative distribution"},
	"GCD": {Name: "GCD", Category: "Math & Trig", Description: "Returns the greatest common divisor"},
	"GEOMEAN": {Name: "GEOMEAN", Category: "Statistical", Description: "Returns the geometric mean"},
	"GESTEP": {Name: "GESTEP", Category: "Engineering", Description: "Tests whether a number is greater than a threshold value"},
	"GETPIVOTDATA": {Name: "GETPIVOTDATA", Category: "Lookup & Reference", Description: "Returns data stored in a PivotTable report"},
	"GROUPBY": {Name: "GROUPBY", Category: "Lookup & Reference", Description: "Helps a user group, aggregate, sort, and filter data based on the fields you specify"},
	"GROWTH": {Name: "GROWTH", Category: "Statistical", Description: "Returns values along an exponential trend"},
	"HARMEAN": {Name: "HARMEAN", Category: "Statistical", Description: "Returns the harmonic mean"},
	"HEX2BIN": {Name: "HEX2BIN", Category: "Engineering", Description: "Converts a hexadecimal number to binary"},
	"HEX2DEC": {Name: "HEX2DEC", Category: "Engineering", Description: "Converts a hexadecimal number to decimal"},
	"HEX2OCT": {Name: "HEX2OCT", Category: "Engineering", Description: "Converts a hexadecimal number to octal"},
	"HLOOKUP": {Name: "HLOOKUP", Category: "Lookup & Reference", Description: "Looks in the top row of an array and returns the value of the indicated cell"},
	"HOUR": {Name: "HOUR", Category: "Date & Time", Description: "Converts a serial number to an hour"},
	"HSTACK": {Name: "HSTACK", Category: "Lookup & Reference", Description: "Appends arrays horizontally and in sequence to return a larger array"},
	"HYPERLINK": {Name: "HYPERLINK", Category: "Lookup & Reference", Description: "Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet"},
	"HYPGEOM.DIST": {Name: "HYPGEOM.DIST", Category: "Statistical", Description: "Returns the hypergeometric distribution"},
	"HYPGEOMDIST": {Name: "HYPGEOMDIST", Category: "Statistical", Description: "Returns the hypergeometric distribution"},
	"IF": {Name: "IF", Category: "Logical", Description: "Specifies a logical test to perform"},
	"IFERROR": {Name: "IFERROR", Category: "Logical", Description: "Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula"},
	"IFNA": {Name: "IFNA", Category: "Logical", Description: "Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression"},
	"IFS": {Name: "IFS", Category: "Logical", Description: "Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition."},
	"IMABS": {Name: "IMABS", Category: "Engineering", Description: "Returns the absolute value (modulus) of a complex number"},
	"IMAGE": {Name: "IMAGE", Category: "Lookup & Reference", Description: "Returns an image from a given source"},
	"IMAGINARY": {Name: "IMAGINARY", Category: "Engineering", Description: "Returns the imaginary coefficient of a complex number"},
	"IMARGUMENT": {Name: "IMARGUMENT", Category: "Engineering", Description: "Returns the argument theta, an angle expressed in radians"},
	"IMCONJUGATE": {Name: "IMCONJUGATE", Category: "Engineering", Description: "Returns the complex conjugate of a complex number"},
	"IMCOS": {Name: "IMCOS", Category: "Engineering", Description: "Returns the cosine of a complex number"},
	"IMCOSH": {Name: "IMCOSH", Category: "Engineering", Description: "Returns the hyperbolic cosine of a complex number"},
	"IMCOT": {Name: "IMCOT", Category: "Engineering", Description: "Returns the cotangent of a complex number"},
	"IMCSC": {Name: "IMCSC", Category: "Engineering", Description: "Returns the cosecant of a complex number"},
	"IMCSCH": {Name: "IMCSCH", Category: "Engineering", Description: "Returns the hyperbolic cosecant of a complex number"},
	"IMDIV": {Name: "IMDIV", Category: "Engineering", Description: "Returns the quotient of two complex numbers"},
	"IMEXP": {Name: "IMEXP", Category: "Engineering", Description: "Returns the exponential of a complex number"},
	"IMLN": {Name: "IMLN", Category: "Engineering", Description: "Returns the natural logarithm of a complex number"},
	"IMLOG10": {Name: "IMLOG10", Category: "Engineering", Description: "Returns the base-10 logarithm of a complex number"},
	"IMLOG2": {Name: "IMLOG2", Category: "Engineering", Description: "Returns the base-2 logarithm of a complex number"},
	"IMPOWER": {Name: "IMPOWER", Category: "Engineering", Description: "Returns a complex number raised to an integer power"},
	"IMPRODUCT": {Name: "IMPRODUCT", Category: "Engineering", Description: "Returns the product of from 2 to 255 complex numbers"},
	"IMREAL": {Name: "IMREAL", Category: "Engineering", Description: "Returns the real coefficient of a complex number"},
	"IMSEC": {Name: "IMSEC", Category: "Engineering", Description: "Returns the secant of a complex number"},
	"IMSECH": {Name: "IMSECH", Category: "Engineering", Description: "Returns the hyperbolic secant of a complex number"},
	"IMSIN": {Name: "IMSIN", Category: "Engineering", Description: "Returns the sine of a complex number"},
	"IMSINH": {Name: "IMSINH", Category: "Engineering", Description: "Returns the hyperbolic sine of a complex number"},
	"IMSQRT": {Name: "IMSQRT", Category: "Engineering", Description: "Returns the square root of a complex number"},
	"IMSUB": {Name: "IMSUB", Category: "Engineering", Description: "Returns the difference between two complex numbers"},
	"IMSUM": {Name: "IMSUM", Category: "Engineering", Description: "Returns the sum of complex numbers"},
	"IMTAN": {Name: "IMTAN", Category: "Engineering", Description: "Returns the tangent of a complex number"},
	"INDEX": {Name: "INDEX", Category: "Lookup & Reference", Description: "Uses an index to choose a value from a reference or array"},
	"INDIRECT": {Name: "INDIRECT", Category: "Lookup & Reference", Description: "Returns a reference indicated by a text value"},
	"INFO": {Name: "INFO", Category: "Information", Description: "Returns information about the current operating environment"},
	"INT": {Name: "INT", Category: "Math & Trig", Description: "Rounds a number down to the nearest integer"},
	"INTERCEPT": {Name: "INTERCEPT", Category: "Statistical", Description: "Returns the intercept of the linear regression line"},
	"INTRATE": {Name: "INTRATE", Category: "Financial", Description: "Returns the interest rate for a fully invested security"},
	"IPMT": {Name: "IPMT", Category: "Financial", Description: "Returns the interest payment for an investment for a given period"},
	"IRR": {Name: "IRR", Category: "Financial", Description: "Returns the internal rate of return for a series of cash flows"},
	"ISBLANK": {Name: "ISBLANK", Category: "Information", Description: "Returns TRUE if the value is blank"},
	"ISERR": {Name: "ISERR", Category: "Information", Description: "Returns TRUE if the value is any error value except #N/A"},
	"ISERROR": {Name: "ISERROR", Category: "Information", Description: "Returns TRUE if the value is any error value"},
	"ISEVEN": {Name: "ISEVEN", Category: "Information", Description: "Returns TRUE if the number is even"},
	"ISFORMULA": {Name: "ISFORMULA", Category: "Information", Description: "Returns TRUE if there is a reference to a cell that contains a formula"},
	"ISLOGICAL": {Name: "ISLOGICAL", Category: "Information", Description: "Returns TRUE if the value is a logical value"},
	"ISNA": {Name: "ISNA", Category: "Information", Description: "Returns TRUE if the value is the #N/A error value"},
	"ISNONTEXT": {Name: "ISNONTEXT", Category: "Information", Description: "Returns TRUE if the value is not text"},
	"ISNUMBER": {Name: "ISNUMBER", Category: "Information", Description: "Returns TRUE if the value is a number"},
	"ISO.CEILING": {Name: "ISO.CEILING", Category: "Math & Trig", Description: "Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance"},
	"ISODD": {Name: "ISODD", Category: "Information", Description: "Returns TRUE if the number is odd"},
	"ISOMITTED": {Name: "ISOMITTED", Category: "Information", Description: "Checks whether the value in a LAMBDA is missing and returns TRUE or FALSE"},
	"ISOWEEKNUM": {Name: "ISOWEEKNUM", Category: "Date & Time", Description: "Returns the number of the ISO week number of the year for a given date"},
	"ISPMT": {Name: "ISPMT", Category: "Financial", Description: "Calculates the interest paid during a specific period of an investment"},
	"ISREF": {Name: "ISREF", Category: "Information", Description: "Returns TRUE if the value is a reference"},
	"ISTEXT": {Name: "ISTEXT", Category: "Information", Description: "Returns TRUE if the value is text"},
	"KURT": {Name: "KURT", Category: "Statistical", Description: "Returns the kurtosis of a data set"},
	"LAMBDA": {Name: "LAMBDA", Category: "Logical", Description: "Create custom, reusable functions and call them by a friendly name"},
	"LARGE": {Name: "LARGE", Category: "Statistical", Description: "Returns the k-th largest value in a data set"},
	"LCM": {Name: "LCM", Category: "Math & Trig", Description: "Returns the least common multiple"},
	"LEFT": {Name: "LEFT", Category: "Text", Description: "Returns the leftmost characters from a text value"},
	"LEN": {Name: "LEN", Category: "Text", Description: "Returns the number of characters in a text string"},
	"LET": {Name: "LET", Category: "Logical", Description: "Assigns names to calculation results to allow storing intermediate calculations, values, or defining names inside a formula"},
	"LINEST": {Name: "LINEST", Category: "Statistical", Description: "Returns the parameters of a linear trend"},
	"LN": {Name: "LN", Category: "Math & Trig", Description: "Returns the natural logarithm of a number"},
	"LOG": {Name: "LOG", Category: "Math & Trig", Description: "Returns the logarithm of a number to a specified base"},
	"LOG10": {Name: "LOG10", Category: "Math & Trig", Description: "Returns the base-10 logarithm of a number"},
	"LOGEST": {Name: "LOGEST", Category: "Statistical", Description: "Returns the parameters of an exponential trend"},
	"LOGINV": {Name: "LOGINV", Category: "Statistical", Description: "Returns the inverse of the lognormal cumulative distribution function"},
	"LOGNORM.DIST": {Name: "LOGNORM.DIST", Category: "Statistical", Description: "Returns the cumulative lognormal distribution"},
	"LOGNORM.INV": {Name: "LOGNORM.INV", Category: "Statistical", Description: "Returns the inverse of the lognormal cumulative distribution"},
	"LOGNORMDIST": {Name: "LOGNORMDIST", Category: "Statistical", Description: "Returns the cumulative lognormal distribution"},
	"LOOKUP": {Name: "LOOKUP", Category: "Lookup & Reference", Description: "Looks up values in a vector or array"},
	"LOWER": {Name: "LOWER", Category: "Text", Description: "Converts text to lowercase"},
	"MAKEARRAY": {Name: "MAKEARRAY", Category: "Logical", Description: "Returns a calculated array of a specified row and column size, by applying a LAMBDA"},
	"MAP": {Name: "MAP", Category: "Logical", Description: "Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA to create a new value"},
	"MATCH": {Name: "MATCH", Category: "Lookup & Reference", Description: "Looks up values in a reference or array"},
	"MAX": {Name: "MAX", Category: "Statistical", Description: "Returns the maximum value in a list of arguments"},
	"MAXA": {Name: "MAXA", Category: "Statistical", Description: "Returns the maximum value in a list of arguments, including numbers, text, and logical values"},
	"MAXIFS": {Name: "MAXIFS", Category: "Statistical", Description: "Returns the maximum value among cells specified by a given set of conditions or criteria"},
	"MDETERM": {Name: "MDETERM", Category: "Math & Trig", Description: "Returns the matrix determinant of an array"},
	"MDURATION": {Name: "MDURATION", Category: "Financial", Description: "Returns the Macauley modified duration for a security with an assumed par value of $100"},
	"MEDIAN": {Name: "MEDIAN", Category: "Statistical", Description: "Returns the median of the given numbers"},
	"MID": {Name: "MID", Category: "Text", Description: "Returns a specific number of characters from a text string starting at the position you specify"},
	"MIN": {Name: "MIN", Category: "Statistical", Description: "Returns the minimum value in a list of arguments"},
	"MINA": {Name: "MINA", Category: "Statistical", Description: "Returns the smallest value in a list of arguments, including numbers, text, and logical values"},
	"MINIFS": {Name: "MINIFS", Category: "Statistical", Description: "Returns the minimum value among cells specified by a given set of conditions or criteria."},
	"MINUTE": {Name: "MINUTE", Category: "Date & Time", Description: "Converts a serial number to a minute"},
	"MINVERSE": {Name: "MINVERSE", Category: "Math & Trig", Description: "Returns the matrix inverse of an array"},
	"MIRR": {Name: "MIRR", Category: "Financial", Description: "Returns the internal rate of return where positive and negative cash flows are financed at different rates"},
	"MMULT": {Name: "MMULT", Category: "Math & Trig", Description: "Returns the matrix product of two arrays"},
	"MOD": {Name: "MOD", Category: "Math & Trig", Description: "Returns the remainder from division"},
	"MODE": {Name: "MODE", Category: "Statistical", Description: "Returns the most common value in a data set"},
	"MODE.MULT": {Name: "MODE.MULT", Category: "Statistical", Description: "Returns a vertical array of the most frequently occurring, or repetitive values in an array or range of data"},
	"MODE.SNGL": {Name: "MODE.SNGL", Category: "Statistical", Description: "Returns the most common value in a data set"},
	"MONTH": {Name: "MONTH", Category: "Date & Time", Description: "Converts a serial number to a month"},
	"MROUND": {Name: "MROUND", Category: "Math & Trig", Description: "Returns a number rounded to the desired multiple"},
	"MULTINOMIAL": {Name: "MULTINOMIAL", Category: "Math & Trig", Description: "Returns the multinomial of a set of numbers"},
	"MUNIT": {Name: "MUNIT", Category: "Math & Trig", Description: "Returns the unit matrix or the specified dimension"},
	"N": {Name: "N", Category: "Information", Description: "Returns a value converted to a number"},
	"NA": {Name: "NA", Category: "Information", Description: "Returns the error value #N/A"},
	"NEGBINOM.DIST": {Name: "NEGBINOM.DIST", Category: "Statistical", Description: "Returns the negative binomial distribution"},
	"NEGBINOMDIST": {Name: "NEGBINOMDIST", Category: "Statistical", Description: "Returns the negative binomial distribution"},
	"NETWORKDAYS": {Name: "NETWORKDAYS", Category: "Date & Time", Description: "Returns the number of whole workdays between two dates"},
	"NETWORKDAYS.INTL": {Name: "NETWORKDAYS.INTL", Category: "Date & Time", Description: "Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days"},
	"NOMINAL": {Name: "NOMINAL", Category: "Financial", Description: "Returns the annual nominal interest rate"},
	"NORM.DIST": {Name: "NORM.DIST", Category: "Statistical", Description: "Returns the normal cumulative distribution"},
	"NORM.INV": {Name: "NORM.INV", Category: "Statistical", Description: "Returns the inverse of the normal cumulative distribution"},
	"NORM.S.DIST": {Name: "NORM.S.DIST", Category: "Statistical", Description: "Returns the standard normal cumulative distribution"},
	"NORM.S.INV": {Name: "NORM.S.INV", Category: "Statistical", Description: "Returns the inverse of the standard normal cumulative distribution"},
	"NORMDIST": {Name: "NORMDIST", Category: "Statistical", Description: "Returns the normal cumulative distribution"},
	"NORMINV": {Name: "NORMINV", Category: "Statistical", Description: "Returns the inverse of the normal cumulative distribution"},
	"NORMSDIST": {Name: "NORMSDIST", Category: "Statistical", Description: "Returns the standard normal cumulative distribution"},
	"NORMSINV": {Name: "NORMSINV", Category: "Statistical", Description: "Returns the inverse of the standard normal cumulative distribution"},
	"NOT": {Name: "NOT", Category: "Logical", Description: "Reverses the logic of its argument"},
	"NOW": {Name: "NOW", Category: "Date & Time", Description: "Returns the serial number of the current date and time"},
	"NPER": {Name: "NPER", Category: "Financial", Description: "Returns the number of periods for an investment"},
	"NPV": {Name: "NPV", Category: "Financial", Description: "Returns the net present value of an investment based on a series of periodic cash flows and a discount rate"},
	"NUMBERVALUE": {Name: "NUMBERVALUE", Category: "Text", Description: "Converts text to number in a locale-independent manner"},
	"OCT2BIN": {Name: "OCT2BIN", Category: "Engineering", Description: "Converts an octal number to binary"},
	"OCT2DEC": {Name: "OCT2DEC", Category: "Engineering", Description: "Converts an octal number to decimal"},
	"OCT2HEX": {Name: "OCT2HEX", Category: "Engineering", Description: "Converts an octal number to hexadecimal"},
	"ODD": {Name: "ODD", Category: "Math & Trig", Description: "Rounds a number up to the nearest odd integer"},
	"ODDFPRICE": {Name: "ODDFPRICE", Category: "Financial", Description: "Returns the price per $100 face value of a security with an odd first period"},
	"ODDFYIELD": {Name: "ODDFYIELD", Category: "Financial", Description: "Returns the yield of a security with an odd first period"},
	"ODDLPRICE": {Name: "ODDLPRICE", Category: "Financial", Description: "Returns the price per $100 face value of a security with an odd last period"},
	"ODDLYIELD": {Name: "ODDLYIELD", Category: "Financial", Description: "Returns the yield of a security with an odd last period"},
	"OFFSET": {Name: "OFFSET", Category: "Lookup & Reference", Description: "Returns a reference offset from a given reference"},
	"OR": {Name: "OR", Category: "Logical", Description: "Returns TRUE if any argument is TRUE"},
	"PDURATION": {Name: "PDURATION", Category: "Financial", Description: "Returns the number of periods required by an investment to reach a specified value"},
	"PEARSON": {Name: "PEARSON", Category: "Statistical", Description: "Returns the Pearson product moment correlation coefficient"},
	"PERCENTILE": {Name: "PERCENTILE", Category: "Statistical", Description: "Returns the k-th percentile of values in a range"},
	"PERCENTILE.EXC": {Name: "PERCENTILE.EXC", Category: "Statistical", Description: "Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive"},
	"PERCENTILE.INC": {Name: "PERCENTILE.INC", Category: "Statistical", Description: "Returns the k-th percentile of values in a range"},
	"PERCENTOF": {Name: "PERCENTOF", Category: "Math & Trig", Description: "Sums the values in the subset and divides it by all the values"},
	"PERCENTRANK": {Name: "PERCENTRANK", Category: "Statistical", Description: "Returns the percentage rank of a value in a data set"},
	"PERCENTRANK.EXC": {Name: "PERCENTRANK.EXC", Category: "Statistical", Description: "Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set"},
	"PERCENTRANK.INC": {Name: "PERCENTRANK.INC", Category: "Statistical", Description: "Returns the percentage rank of a value in a data set"},
	"PERMUT": {Name: "PERMUT", Category: "Statistical", Description: "Returns the number of permutations for a given number of objects"},
	"PERMUTATIONA": {Name: "PERMUTATIONA", Category: "Statistical", Description: "Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects"},
	"PHI": {Name: "PHI", Category: "Statistical", Description: "Returns the value of the density function for a standard normal distribution"},
	"PHONETIC": {Name: "PHONETIC", Category: "Text", Description: "Extracts the phonetic (furigana) characters from a text string"},
	"PI": {Name: "PI", Category: "Math & Trig", Description: "Returns the value of pi"},
	"PIVOTBY": {Name: "PIVOTBY", Category: "Lookup & Reference", Description: "Helps a user group, aggregate, sort, and filter data based on the row and column fields that you specify"},
	"PMT": {Name: "PMT", Category: "Financial", Description: "Returns the periodic payment for an annuity"},
	"POISSON": {Name: "POISSON", Category: "Statistical", Description: "Returns the Poisson distribution"},
	"POISSON.DIST": {Name: "POISSON.DIST", Category: "Statistical", Description: "Returns the Poisson distribution"},
	"POWER": {Name: "POWER", Category: "Math & Trig", Description: "Returns the result of a number raised to a power"},
	"PPMT": {Name: "PPMT", Category: "Financial", Description: "Returns the payment on the principal for an investment for a given period"},
	"PRICE": {Name: "PRICE", Category: "Financial", Description: "Returns the price per $100 face value of a security that pays periodic interest"},
	"PRICEDISC": {Name: "PRICEDISC", Category: "Financial", Description: "Returns the price per $100 face value of a discounted security"},
	"PRICEMAT": {Name: "PRICEMAT", Category: "Financial", Description: "Returns the price per $100 face value of a security that pays interest at maturity"},
	"PROB": {Name: "PROB", Category: "Statistical", Description: "Returns the probability that values in a range are between two limits"},
	"PRODUCT": {Name: "PRODUCT", Category: "Math & Trig", Description: "Multiplies its arguments"},
	"PROPER": {Name: "PROPER", Category: "Text", Description: "Capitalizes the first letter in each word of a text value"},
	"PV": {Name: "PV", Category: "Financial", Description: "Returns the present value of an investment"},
	"QUARTILE": {Name: "QUARTILE", Category: "Statistical", Description: "Returns the quartile of a data set"},
	"QUARTILE.EXC": {Name: "QUARTILE.EXC", Category: "Statistical", Description: "Returns the quartile of the data set, based on percentile values from 0..1, exclusive"},
	"QUARTILE.INC": {Name: "QUARTILE.INC", Category: "Statistical", Description: "Returns the quartile of a data set"},
	"QUOTIENT": {Name: "QUOTIENT", Category: "Math & Trig", Description: "Returns the integer portion of a division"},
	"RADIANS": {Name: "RADIANS", Category: "Math & Trig", Description: "Converts degrees to radians"},
	"RAND": {Name: "RAND", Category: "Math & Trig", Description: "Returns a random number between 0 and 1"},
	"RANDARRAY": {Name: "RANDARRAY", Category: "Math & Trig", Description: "Returns an array of random numbers between 0 and 1. However, you can specify the number of rows and columns to fill, minimum and maximum values, and whether to return whole numbers or decimal values."},
	"RANDBETWEEN": {Name: "RANDBETWEEN", Category: "Math & Trig", Description: "Returns a random number between the numbers you specify"},
	"RANK": {Name: "RANK", Category: "Statistical", Description: "Returns the rank of a number in a list of numbers"},
	"RANK.AVG": {Name: "RANK.AVG", Category: "Statistical", Description: "Returns the rank of a number in a list of numbers"},
	"RANK.EQ": {Name: "RANK.EQ", Category: "Statistical", Description: "Returns the rank of a number in a list of numbers"},
	"RATE": {Name: "RATE", Category: "Financial", Description: "Returns the interest rate per period of an annuity"},
	"RECEIVED": {Name: "RECEIVED", Category: "Financial", Description: "Returns the amount received at maturity for a fully invested security"},
	"REDUCE": {Name: "REDUCE", Category: "Logical", Description: "Reduces an array to an accumulated value by applying a LAMBDA to each value and returning the total value in the accumulator"},
	"REGEXEXTRACT": {Name: "REGEXEXTRACT", Category: "Text", Description: "Extracts strings within the provided text that matches the pattern"},
	"REGEXREPLACE": {Name: "REGEXREPLACE", Category: "Text", Description: "Replaces strings within the provided text that matches the pattern with replacement"},
	"REGEXTEST": {Name: "REGEXTEST", Category: "Text", Description: "Determines whether any part of text matches the pattern"},
	"REGISTER.ID": {Name: "REGISTER.ID", Category: "User Defined", Description: "Returns the register ID of the specified dynamic link library (DLL) or code resource that has been previously registered"},
	"REPLACE": {Name: "REPLACE", Category: "Text", Description: "Replaces characters within text"},
	"REPT": {Name: "REPT", Category: "Text", Description: "Repeats text a given number of times"},
	"RIGHT": {Name: "RIGHT", Category: "Text", Description: "Returns the rightmost characters from a text value"},
	"ROMAN": {Name: "ROMAN", Category: "Math & Trig", Description: "Converts an Arabic numeral to Roman, as text"},
	"ROUND": {Name: "ROUND", Category: "Math & Trig", Description: "Rounds a number to a specified number of digits"},
	"ROUNDDOWN": {Name: "ROUNDDOWN", Category: "Math & Trig", Description: "Rounds a number down, toward zero"},
	"ROUNDUP": {Name: "ROUNDUP", Category: "Math & Trig", Description: "Rounds a number up, away from zero"},
	"ROW": {Name: "ROW", Category: "Lookup & Reference", Description: "Returns the row number of a reference"},
	"ROWS": {Name: "ROWS", Category: "Lookup & Reference", Description: "Returns the number of rows in a reference"},
	"RRI": {Name: "RRI", Category: "Financial", Description: "Returns an equivalent interest rate for the growth of an investment"},
	"RSQ": {Name: "RSQ", Category: "Statistical", Description: "Returns the square of the Pearson product moment correlation coefficient"},
	"RTD": {Name: "RTD", Category: "Lookup & Reference", Description: "Retrieves real-time data from a program that supports COM automation"},
	"SCAN": {Name: "SCAN", Category: "Logical", Description: "Scans an array by applying a LAMBDA to each value and returns an array that has each intermediate value"},
	"SEARCH": {Name: "SEARCH", Category: "Text", Description: "Finds one text value within another (not case-sensitive)"},
	"SEC": {Name: "SEC", Category: "Math & Trig", Description: "Returns the secant of an angle"},
	"SECH": {Name: "SECH", Category: "Math & Trig", Description: "Returns the hyperbolic secant of an angle"},
	"SECOND": {Name: "SECOND", Category: "Date & Time", Description: "Converts a serial number to a second"},
	"SEQUENCE": {Name: "SEQUENCE", Category: "Math & Trig", Description: "Generates a list of sequential numbers in an array, such as 1, 2, 3, 4"},
	"SERIESSUM": {Name: "SERIESSUM", Category: "Math & Trig", Description: "Returns the sum of a power series based on the formula"},
	"SHEET": {Name: "SHEET", Category: "Information", Description: "Returns the sheet number of the referenced sheet"},
	"SHEETS": {Name: "SHEETS", Category: "Information", Description: "Returns the number of sheets in a reference"},
	"SIGN": {Name: "SIGN", Category: "Math & Trig", Description: "Returns the sign of a number"},
	"SIN": {Name: "SIN", Category: "Math & Trig", Description: "Returns the sine of the given angle"},
	"SINH": {Name: "SINH", Category: "Math & Trig", Description: "Returns the hyperbolic sine of a number"},
	"SKEW": {Name: "SKEW", Category: "Statistical", Description: "Returns the skewness of a distribution"},
	"SKEW.P": {Name: "SKEW.P", Category: "Statistical", Description: "Returns the skewness of a distribution based on a population"},
	"SLN": {Name: "SLN", Category: "Financial", Description: "Returns the straight-line depreciation of an asset for one period"},
	"SLOPE": {Name: "SLOPE", Category: "Statistical", Description: "Returns the slope of the linear regression line"},
	"SMALL": {Name: "SMALL", Category: "Statistical", Description: "Returns the k-th smallest value in a data set"},
	"SORT": {Name: "SORT", Category: "Lookup & Reference", Description: "Sorts the contents of a range or array"},
	"SORTBY": {Name: "SORTBY", Category: "Lookup & Reference", Description: "Sorts the contents of a range or array based on the values in a corresponding range or array"},
	"SQRT": {Name: "SQRT", Category: "Math & Trig", Description: "Returns a positive square root"},
	"SQRTPI": {Name: "SQRTPI", Category: "Math & Trig", Description: "Returns the square root of (number * pi)"},
	"STANDARDIZE": {Name: "STANDARDIZE", Category: "Statistical", Description: "Returns a normalized value"},
	"STDEV": {Name: "STDEV", Category: "Statistical", Description: "Estimates standard deviation based on a sample"},
	"STDEV.P": {Name: "STDEV.P", Category: "Statistical", Description: "Calculates standard deviation based on the entire population"},
	"STDEV.S": {Name: "STDEV.S", Category: "Statistical", Description: "Estimates standard deviation based on a sample"},
	"STDEVA": {Name: "STDEVA", Category: "Statistical", Description: "Estimates standard deviation based on a sample, including numbers, text, and logical values"},
	"STDEVP": {Name: "STDEVP", Category: "Statistical", Description: "Calculates standard deviation based on the entire population"},
	"STDEVPA": {Name: "STDEVPA", Category: "Statistical", Description: "Calculates standard deviation based on the entire population, including numbers, text, and logical values"},
	"STEYX": {Name: "STEYX", Category: "Statistical", Description: "Returns the standard error of the predicted y-value for each x in the regression"},
	"STOCKHISTORY": {Name: "STOCKHISTORY", Category: "Information", Description: "Retrieves historical data about a financial instrument"},
	"SUBSTITUTE": {Name: "SUBSTITUTE", Category: "Text", Description: "Substitutes new text for old text in a text string"},
	"SUBTOTAL": {Name: "SUBTOTAL", Category: "Math & Trig", Description: "Returns a subtotal in a list or database"},
	"SUM": {Name: "SUM", Category: "Math & Trig", Description: "Adds its arguments"},
	"SUMIF": {Name: "SUMIF", Category: "Math & Trig", Description: "Adds the cells specified by a given criteria"},
	"SUMIFS": {Name: "SUMIFS", Category: "Math & Trig", Description: "Adds the cells in a range that meet multiple criteria"},
	"SUMPRODUCT": {Name: "SUMPRODUCT", Category: "Math & Trig", Description: "Returns the sum of the products of corresponding array components"},
	"SUMSQ": {Name: "SUMSQ", Category: "Math & Trig", Description: "Returns the sum of the squares of the arguments"},
	"SUMX2MY2": {Name: "SUMX2MY2", Category: "Math & Trig", Description: "Returns the sum of the difference of squares of corresponding values in two arrays"},
	"SUMX2PY2": {Name: "SUMX2PY2", Category: "Math & Trig", Description: "Returns the sum of the sum of squares of corresponding values in two arrays"},
	"SUMXMY2": {Name: "SUMXMY2", Category: "Math & Trig", Description: "Returns the sum of squares of differences of corresponding values in two arrays"},
	"SWITCH": {Name: "SWITCH", Category: "Logical", Description: "Evaluates an expression against a list of values and returns the result corresponding to the first matching value."},
	"SYD": {Name: "SYD", Category: "Financial", Description: "Returns the sum-of-years' digits depreciation of an asset for a specified period"},
	"T": {Name: "T", Category: "Text", Description: "Converts its arguments to text"},
	"T.DIST": {Name: "T.DIST", Category: "Statistical", Description: "Returns the Percentage Points (probability) for the Student t-distribution"},
	"T.DIST.2T": {Name: "T.DIST.2T", Category: "Statistical", Description: "Returns the Percentage Points (probability) for the Student t-distribution"},
	"T.DIST.RT": {Name: "T.DIST.RT", Category: "Statistical", Description: "Returns the Student's t-distribution"},
	"T.INV": {Name: "T.INV", Category: "Statistical", Description: "Returns the t-value of the Student's t-distribution as a function of the probability and the degrees of freedom"},
	"T.INV.2T": {Name: "T.INV.2T", Category: "Statistical", Description: "Returns the inverse of the Student's t-distribution"},
	"T.TEST": {Name: "T.TEST", Category: "Statistical", Description: "Returns the probability associated with a Student's t-test"},
	"TAKE": {Name: "TAKE", Category: "Lookup & Reference", Description: "Returns a specified number of contiguous rows or columns from the start or end of an array"},
	"TAN": {Name: "TAN", Category: "Math & Trig", Description: "Returns the tangent of a number"},
	"TANH": {Name: "TANH", Category: "Math & Trig", Description: "Returns the hyperbolic tangent of a number"},
	"TBILLEQ": {Name: "TBILLEQ", Category: "Financial", Description: "Returns the bond-equivalent yield for a Treasury bill"},
	"TBILLPRICE": {Name: "TBILLPRICE", Category: "Financial", Description: "Returns the price per $100 face value for a Treasury bill"},
	"TBILLYIELD": {Name: "TBILLYIELD", Category: "Financial", Description: "Returns the yield for a Treasury bill"},
	"TDIST": {Name: "TDIST", Category: "Statistical", Description: "Returns the Student's t-distribution"},
	"TEXT": {Name: "TEXT", Category: "Text", Description: "Formats a number and converts it to text"},
	"TEXTAFTER": {Name: "TEXTAFTER", Category: "Text", Description: "Returns text that occurs after given character or string"},
	"TEXTBEFORE": {Name: "TEXTBEFORE", Category: "Text", Description: "Returns text that occurs before a given character or string"},
	"TEXTJOIN": {Name: "TEXTJOIN", Category: "Text", Description: "Combines the text from multiple ranges and/or strings"},
	"TEXTSPLIT": {Name: "TEXTSPLIT", Category: "Text", Description: "Splits text strings by using column and row delimiters"},
	"TIME": {Name: "TIME", Category: "Date & Time", Description: "Returns the serial number of a particular time"},
	"TIMEVALUE": {Name: "TIMEVALUE", Category: "Date & Time", Description: "Converts a time in the form of text to a serial number"},
	"TINV": {Name: "TINV", Category: "Statistical", Description: "Returns the inverse of the Student's t-distribution"},
	"TOCOL": {Name: "TOCOL", Category: "Lookup & Reference", Description: "Returns the array in a single column"},
	"TODAY": {Name: "TODAY", Category: "Date & Time", Description: "Returns the serial number of today's date"},
	"TOROW": {Name: "TOROW", Category: "Lookup & Reference", Description: "Returns the array in a single row"},
	"TRANSLATE": {Name: "TRANSLATE", Category: "Text", Description: "Translates a text from one language to another"},
	"TRANSPOSE": {Name: "TRANSPOSE", Category: "Lookup & Reference", Description: "Returns the transpose of an array"},
	"TREND": {Name: "TREND", Category: "Statistical", Description: "Returns values along a linear trend"},
	"TRIM": {Name: "TRIM", Category: "Text", Description: "Removes spaces from text"},
	"TRIMMEAN": {Name: "TRIMMEAN", Category: "Statistical", Description: "Returns the mean of the interior of a data set"},
	"TRIMRANGE": {Name: "TRIMRANGE", Category: "Lookup & Reference", Description: "Scans in from the edges of a range or array until it finds a non-blank cell (or value), it then excludes those blank rows or columns"},
	"TRUE": {Name: "TRUE", Category: "Logical", Description: "Returns the logical value TRUE"},
	"TRUNC": {Name: "TRUNC", Category: "Math & Trig", Description: "Truncates a number to an integer"},
	"TTEST": {Name: "TTEST", Category: "Statistical", Description: "Returns the probability associated with a Student's t-test"},
	"TYPE": {Name: "TYPE", Category: "Information", Description: "Returns a number indicating the data type of a value"},
	"UNICHAR": {Name: "UNICHAR", Category: "Text", Description: "Returns the Unicode character that is references by the given numeric value"},
	"UNICODE": {Name: "UNICODE", Category: "Text", Description: "Returns the number (code point) that corresponds to the first character of the text"},
	"UNIQUE": {Name: "UNIQUE", Category: "Lookup & Reference", Description: "Returns a list of unique values in a list or range"},
	"UPPER": {Name: "UPPER", Category: "Text", Description: "Converts text to uppercase"},
	"VALUE": {Name: "VALUE", Category: "Text", Description: "Converts a text argument to a number"},
	"VALUETOTEXT": {Name: "VALUETOTEXT", Category: "Text", Description: "Returns text from any specified value"},
	"VAR": {Name: "VAR", Category: "Statistical", Description: "Estimates variance based on a sample"},
	"VAR.P": {Name: "VAR.P", Category: "Statistical", Description: "Calculates variance based on the entire population"},
	"VAR.S": {Name: "VAR.S", Category: "Statistical", Description: "Estimates variance based on a sample"},
	"VARA": {Name: "VARA", Category: "Statistical", Description: "Estimates variance based on a sample, including numbers, text, and logical values"},
	"VARP": {Name: "VARP", Category: "Statistical", Description: "Calculates variance based on the entire population"},
	"VARPA": {Name: "VARPA", Category: "Statistical", Description: "Calculates variance based on the entire population, including numbers, text, and logical values"},
	"VDB": {Name: "VDB", Category: "Financial", Description: "Returns the depreciation of an asset for a specified or partial period by using a declining balance method"},
	"VLOOKUP": {Name: "VLOOKUP", Category: "Lookup & Reference", Description: "Looks in the first column of an array and moves across the row to return the value of a cell"},
	"VSTACK": {Name: "VSTACK", Category: "Lookup & Reference", Description: "Appends arrays vertically and in sequence to return a larger array"},
	"WEBSERVICE": {Name: "WEBSERVICE", Category: "Web", Description: "Returns data from a web service"},
	"WEEKDAY": {Name: "WEEKDAY", Category: "Date & Time", Description: "Converts a serial number to a day of the week"},
	"WEEKNUM": {Name: "WEEKNUM", Category: "Date & Time", Description: "Converts a serial number to a number representing where the week falls numerically with a year"},
	"WEIBULL": {Name: "WEIBULL", Category: "Statistical", Description: "Returns the Weibull distribution"},
	"WEIBULL.DIST": {Name: "WEIBULL.DIST", Category: "Statistical", Description: "Returns the Weibull distribution"},
	"WORKDAY": {Name: "WORKDAY", Category: "Date & Time", Description: "Returns the serial number of the date before or after a specified number of workdays"},
	"WORKDAY.INTL": {Name: "WORKDAY.INTL", Category: "Date & Time", Description: "Returns the serial number of the date before or after a specified number of workdays using parameters to indicate which and how many days are weekend days"},
	"WRAPCOLS": {Name: "WRAPCOLS", Category: "Lookup & Reference", Description: "Wraps the provided row or column of values by columns after a specified number of elements"},
	"WRAPROWS": {Name: "WRAPROWS", Category: "Lookup & Reference", Description: "Wraps the provided row or column of values by rows after a specified number of elements"},
	"XIRR": {Name: "XIRR", Category: "Financial", Description: "Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic"},
	"XLOOKUP": {Name: "XLOOKUP", Category: "Lookup & Reference", Description: "Searches a range or an array, and returns an item corresponding to the first match it finds."},
	"XMATCH": {Name: "XMATCH", Category: "Lookup & Reference", Description: "Returns the relative position of an item in an array or range of cells."},
	"XNPV": {Name: "XNPV", Category: "Financial", Description: "Returns the net present value for a schedule of cash flows that is not necessarily periodic"},
	"XOR": {Name: "XOR", Category: "Logical", Description: "Returns a logical exclusive OR of all arguments"},
	"YEAR": {Name: "YEAR", Category: "Date & Time", Description: "Converts a serial number to a year"},
	"YEARFRAC": {Name: "YEARFRAC", Category: "Date & Time", Description: "Returns the year fraction representing the number of whole days between start_date and end_date"},
	"YIELD": {Name: "YIELD", Category: "Financial", Description: "Returns the yield on a security that pays periodic interest"},
	"YIELDDISC": {Name: "YIELDDISC", Category: "Financial", Description: "Returns the annual yield for a discounted security; for example, a Treasury bill"},
	"YIELDMAT": {Name: "YIELDMAT", Category: "Financial", Description: "Returns the annual yield of a security that pays interest at maturity"},
	"Z.TEST": {Name: "Z.TEST", Category: "Statistical", Description: "Returns the one-tailed probability-value of a z-test"},
	"ZTEST": {Name: "ZTEST", Category: "Statistical", Description: "Returns the one-tailed probability-value of a z-test"},
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
	"regex":   "Text",
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
	"REGEXEXTRACT":             "Extracts one or more substrings that match a regular expression.",
	"REGEXREPLACE":             "Replaces substrings that match a regular expression.",
	"REGEXTEST":                "Returns TRUE if a string matches a regular expression.",
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
	// Fallback to Microsoft's official description when we have one.
	// Curated descriptionMap / simpleDescriptionMap entries above always win
	// (so werkbook's voice is preserved); this only fills in functions we
	// haven't hand-described yet (e.g. GROUPBY, ISOMITTED, TRIMRANGE).
	if info, ok := allExcelFunctions[name]; ok && info.Description != "" {
		desc := info.Description
		if !strings.HasSuffix(desc, ".") {
			desc += "."
		}
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
	// Match any RegisterX(...) helper call — Register, RegisterWithMeta,
	// RegisterScalarLifted, RegisterWithSpec, RegisterWithMetaAndSpec,
	// RegisterScalarLiftedUnarySpec, etc. New helpers shouldn't require
	// updating this regex.
	re := regexp.MustCompile(`Register[A-Za-z]*\("([A-Z][A-Z0-9.]*)"`)
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

// computeUnsupported drives the unsupported list off allExcelFunctions (the
// universe sourced from Microsoft's official function reference). Anything
// in the universe that is not supported and not in willNotSupport ends up
// in # Not Yet Implemented; willNotSupport entries always end up in
// # No Planned Support. The category and description come from the
// allExcelFunctions entry when available, falling back to categoryForUnsupported
// for category and the descriptionFor() chain for description.
func computeUnsupported(supported map[string]funcInfo) []funcInfo {
	seen := map[string]bool{}
	var unsupported []funcInfo

	// Always include willNotSupport entries — they are emitted under
	// # No Planned Support by generateMarkdown using the willNotSupport map
	// for the reason. Category prefers allExcelFunctions, falls back to
	// the hand-coded unsupportedCategories map.
	for name := range willNotSupport {
		if _, ok := supported[name]; ok {
			continue
		}
		seen[name] = true
		cat := categoryForUnsupported(name)
		if info, ok := allExcelFunctions[name]; ok {
			cat = info.Category
		}
		unsupported = append(unsupported, newFuncInfo(name, cat))
	}

	// Then add every Excel function in the universe that werkbook does not
	// implement and that we have not explicitly chosen not to support.
	for name, info := range allExcelFunctions {
		if _, ok := supported[name]; ok {
			continue
		}
		if _, ok := willNotSupport[name]; ok {
			continue
		}
		if seen[name] {
			continue
		}
		seen[name] = true
		unsupported = append(unsupported, newFuncInfo(name, info.Category))
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
