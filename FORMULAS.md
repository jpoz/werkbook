# Supported Formulas

Werkbook supports **450** spreadsheet formula functions.

| Function | Description | Category | Tests |
|----------|-------------|----------|------:|
| ABS | Returns the absolute value of a number. | Math & Trig | 27 |
| ACCRINT | Returns accrued interest for a security that pays periodic interest. | Financial | 19 |
| ACCRINTM | Returns accrued interest for a security that pays interest at maturity. | Financial | 15 |
| ACOS | Returns the arccosine of a number. | Math & Trig | 32 |
| ACOSH | Returns the inverse hyperbolic cosine of a number. | Math & Trig | 32 |
| ACOT | Returns the arccotangent of a number. | Math & Trig | 29 |
| ACOTH | Returns the inverse hyperbolic cotangent of a number. | Math & Trig | 32 |
| ADDRESS | Builds a cell reference from row and column numbers. | Lookup & Reference | 44 |
| AGGREGATE | Applies an aggregate calculation with options to ignore selected values. | Statistical | 98 |
| AMORDEGRC | Returns depreciation for each accounting period using a French declining balance method. | Financial | 12 |
| AMORLINC | Returns depreciation for each accounting period using a linear method. | Financial | 14 |
| ANCHORARRAY | Returns the full spilled array produced by a dynamic array formula. | Lookup & Reference | 9 |
| AND | Returns TRUE when every supplied condition is TRUE. | Logical | 31 |
| ARABIC | Converts Roman numeral text to an Arabic number. | Math & Trig | 38 |
| AREAS | Counts the number of areas in a reference. | Information | 23 |
| ARRAYTOTEXT | Converts an array into a text representation. | Text | 50 |
| ASIN | Returns the arcsine of a number. | Math & Trig | 30 |
| ASINH | Returns the inverse hyperbolic sine of a number. | Math & Trig | 26 |
| ATAN | Returns the arctangent of a number. | Math & Trig | 25 |
| ATAN2 | Returns the arctangent from x and y coordinates. | Math & Trig | 30 |
| ATANH | Returns the inverse hyperbolic tangent of a number. | Math & Trig | 31 |
| AVEDEV | Returns the average absolute deviation from the mean. | Statistical | 23 |
| AVERAGE | Returns the arithmetic mean of the supplied values. | Statistical | 45 |
| AVERAGEA | Returns the average of supplied values, including logical values and text coercions. | Statistical | 24 |
| AVERAGEIF | Returns the average of values that match one condition. | Statistical | 43 |
| AVERAGEIFS | Returns the average of values that match all supplied conditions. | Statistical | 48 |
| BASE | Converts a number to text in the requested base. | Math & Trig | 31 |
| BESSELI | Returns the modified Bessel function I_n(x). | Engineering | 29 |
| BESSELJ | Returns the Bessel function J_n(x). | Engineering | 30 |
| BESSELK | Returns the modified Bessel function K_n(x). | Engineering | 32 |
| BESSELY | Returns the Bessel function Y_n(x). | Engineering | 33 |
| BETA.DIST | Returns the beta distribution. | Statistical | 43 |
| BETA.INV | Returns the inverse of the beta cumulative distribution. | Statistical | 42 |
| BIN2DEC | Converts a binary number to decimal. | Engineering | 38 |
| BIN2HEX | Converts a binary number to hexadecimal. | Engineering | 49 |
| BIN2OCT | Converts a binary number to octal. | Engineering | 45 |
| BINOM.DIST | Returns the binomial distribution. | Statistical | 30 |
| BINOM.DIST.RANGE | Returns the probability that a binomial result falls within a range. | Statistical | 66 |
| BINOM.INV | Returns the smallest value whose binomial cumulative distribution meets a criterion. | Statistical | 56 |
| BITAND | Returns the bitwise AND of two integers. | Math & Trig | 28 |
| BITLSHIFT | Returns a number shifted left by a requested number of bits. | Math & Trig | 27 |
| BITOR | Returns the bitwise OR of two integers. | Math & Trig | 27 |
| BITRSHIFT | Returns a number shifted right by a requested number of bits. | Math & Trig | 27 |
| BITXOR | Returns the bitwise XOR of two integers. | Math & Trig | 27 |
| BYCOL | Applies a lambda to each column of an array. | Logical | 31 |
| BYROW | Applies a lambda to each row of an array. | Logical | 28 |
| CEILING | Rounds a number up to the nearest multiple of a significance. | Math & Trig | 51 |
| CEILING.MATH | Rounds a number up using Excel's CEILING.MATH rules. | Math & Trig | 37 |
| CEILING.PRECISE | Rounds a number up to the nearest significance, ignoring the sign of the significance. | Math & Trig | 37 |
| CHAR | Returns the character for a numeric code. | Text | 56 |
| CHISQ.DIST | Returns the chi-square distribution. | Statistical | 39 |
| CHISQ.DIST.RT | Returns the right-tailed chi-square probability. | Statistical | 29 |
| CHISQ.INV | Returns the inverse of the chi-square cumulative distribution. | Statistical | 36 |
| CHISQ.INV.RT | Returns the inverse of the right-tailed chi-square distribution. | Statistical | 32 |
| CHISQ.TEST | Returns the result of a chi-square test. | Statistical | 23 |
| CHOOSE | Returns the value at a position from a supplied list of choices. | Text | 49 |
| CHOOSECOLS | Returns selected columns from an array. | Lookup & Reference | 29 |
| CHOOSEROWS | Returns selected rows from an array. | Lookup & Reference | 29 |
| CLEAN | Removes non-printing characters from text. | Text | 25 |
| CODE | Returns the numeric code for the first character in text. | Text | 13 |
| COLUMN | Returns the column number of a reference. | Information | 28 |
| COLUMNS | Returns the number of columns in a reference or array. | Information | 38 |
| COMBIN | Returns the number of combinations for a given number of items. | Math & Trig | 40 |
| COMBINA | Returns the number of combinations with repetitions. | Math & Trig | 31 |
| COMPLEX | Builds a complex number from real and imaginary parts. | Engineering | 51 |
| CONCAT | Joins text from multiple values or ranges. | Text | 35 |
| CONCATENATE | Joins multiple text values into one string. | Text | 45 |
| CONFIDENCE.NORM | Returns a normal-distribution confidence interval half-width. | Statistical | 47 |
| CONFIDENCE.T | Returns a Student's t-distribution confidence interval half-width. | Statistical | 50 |
| CONVERT | Converts a number from one measurement system to another. | Engineering | 185 |
| CORREL | Returns the correlation coefficient between two data sets. | Statistical | 51 |
| COS | Returns the cosine of an angle. | Math & Trig | 28 |
| COSH | Returns the hyperbolic cosine of a number. | Math & Trig | 26 |
| COT | Returns the cotangent of an angle. | Math & Trig | 26 |
| COTH | Returns the hyperbolic cotangent of a number. | Math & Trig | 26 |
| COUNT | Counts numeric values. | Statistical | 47 |
| COUNTA | Counts non-empty values. | Statistical | 25 |
| COUNTBLANK | Counts blank cells in a range. | Statistical | 41 |
| COUNTIF | Counts values that match one condition. | Statistical | 80 |
| COUNTIFS | Counts values that match all supplied conditions. | Statistical | 51 |
| COUPDAYBS | Returns the number of days from the coupon period start to settlement. | Financial | 17 |
| COUPDAYS | Returns the number of days in the coupon period containing settlement. | Financial | 17 |
| COUPDAYSNC | Returns the number of days from settlement to the next coupon date. | Financial | 17 |
| COUPNCD | Returns the next coupon date after settlement. | Financial | 19 |
| COUPNUM | Returns the number of coupons payable between settlement and maturity. | Financial | 38 |
| COUPPCD | Returns the previous coupon date before settlement. | Financial | 19 |
| COVAR | Returns the population covariance of two data sets. | Statistical | 34 |
| COVARIANCE.P | Returns the population covariance of two data sets. | Statistical | 21 |
| COVARIANCE.S | Returns the sample covariance of two data sets. | Statistical | 7 |
| CSC | Returns the cosecant of an angle. | Math & Trig | 27 |
| CSCH | Returns the hyperbolic cosecant of a number. | Math & Trig | 28 |
| CUMIPMT | Returns cumulative interest paid over a span of payment periods. | Financial | 56 |
| CUMPRINC | Returns cumulative principal paid over a span of payment periods. | Financial | 53 |
| DATE | Builds a date from year, month, and day numbers. | Date & Time | 72 |
| DATEDIF | Returns the difference between two dates in requested units. | Date & Time | 72 |
| DATEVALUE | Converts date text to a serial date number. | Date & Time | 39 |
| DAVERAGE | Returns the average of matching records in a database range. | Db | 5 |
| DAY | Returns the day of the month from a date. | Date & Time | 28 |
| DAYS | Returns the number of days between two dates. | Date & Time | 17 |
| DAYS360 | Returns the number of days between two dates using a 360-day year. | Date & Time | 80 |
| DB | Returns depreciation using the fixed-declining balance method. | Financial | 29 |
| DCOUNT | Counts numeric values in matching records of a database range. | Db | 6 |
| DCOUNTA | Counts non-empty values in matching records of a database range. | Db | 5 |
| DDB | Returns depreciation using the double-declining balance method or another factor. | Financial | 27 |
| DEC2BIN | Converts a decimal number to binary. | Engineering | 37 |
| DEC2HEX | Converts a decimal number to hexadecimal. | Engineering | 45 |
| DEC2OCT | Converts a decimal number to octal. | Engineering | 45 |
| DECIMAL | Converts a number in a given base to decimal. | Math & Trig | 38 |
| DEGREES | Converts radians to degrees. | Math & Trig | 25 |
| DELTA | Tests whether two numbers are equal and returns 1 or 0. | Engineering | 24 |
| DEVSQ | Returns the sum of squared deviations from the sample mean. | Statistical | 24 |
| DGET | Returns a single value from matching records in a database range. | Db | 2 |
| DISC | Returns the discount rate for a security. | Financial | 14 |
| DMAX | Returns the maximum value from matching records in a database range. | Db | 6 |
| DMIN | Returns the minimum value from matching records in a database range. | Db | 5 |
| DOLLAR | Formats a number as currency text. | Text | 39 |
| DOLLARDE | Converts a fractional dollar price to a decimal price. | Financial | 18 |
| DOLLARFR | Converts a decimal dollar price to a fractional price. | Financial | 17 |
| DPRODUCT | Returns the product of matching records in a database range. | Db | 67 |
| DROP | Drops rows or columns from the start or end of an array. | Lookup & Reference | 27 |
| DSTDEV | Returns the sample standard deviation of matching database records. | Db | 3 |
| DSTDEVP | Returns the population standard deviation of matching database records. | Db | 2 |
| DSUM | Returns the sum of matching records in a database range. | Db | 10 |
| DURATION | Returns the Macauley duration of a security. | Financial | 27 |
| DVAR | Returns the sample variance of matching database records. | Db | 3 |
| DVARP | Returns the population variance of matching database records. | Db | 3 |
| EDATE | Shifts a date by a number of whole months. | Date & Time | 40 |
| EFFECT | Returns the effective annual interest rate for a nominal rate. | Financial | 18 |
| ENCODEURL | Encodes text so it can be safely used in a URL. | Text | 44 |
| EOMONTH | Returns the last day of a month offset from a start date. | Date & Time | 45 |
| ERF | Returns the error function. | Math & Trig | 22 |
| ERF.PRECISE | Returns the error function integrated from 0 to a limit. | Math & Trig | 20 |
| ERFC | Returns the complementary error function. | Math & Trig | 18 |
| ERFC.PRECISE | Returns the complementary error function integrated from a limit to infinity. | Math & Trig | 19 |
| ERROR.TYPE | Returns a number identifying an error value. | Information | 19 |
| EVEN | Rounds a number away from zero to the nearest even integer. | Math & Trig | 34 |
| EXACT | Returns TRUE when two text values match exactly. | Text | 32 |
| EXP | Returns e raised to a power. | Math & Trig | 29 |
| EXPAND | Pads an array to a larger size with a fill value. | Lookup & Reference | 28 |
| EXPON.DIST | Returns the exponential distribution. | Statistical | 28 |
| F.DIST | Returns the F probability distribution. | Statistical | 42 |
| F.DIST.RT | Returns the right-tailed F probability distribution. | Statistical | 30 |
| F.INV | Returns the inverse of the F cumulative distribution. | Statistical | 36 |
| F.INV.RT | Returns the inverse of the right-tailed F distribution. | Statistical | 34 |
| F.TEST | Returns the result of an F-test for two arrays. | Statistical | 42 |
| FACT | Returns the factorial of a number. | Math & Trig | 28 |
| FACTDOUBLE | Returns the double factorial of a number. | Math & Trig | 30 |
| FALSE | Returns the logical value FALSE. | Logical | 11 |
| FILTER | Filters an array to rows or columns that meet a Boolean include mask. | Lookup & Reference | 111 |
| FIND | Returns the position of one text value inside another, case-sensitive. | Text | 48 |
| FISHER | Returns the Fisher transformation of a correlation coefficient. | Statistical | 15 |
| FISHERINV | Returns the inverse Fisher transformation. | Statistical | 15 |
| FIXED | Formats a number as text with a fixed number of decimals. | Text | 47 |
| FLOOR | Rounds a number down to the nearest multiple of a significance. | Math & Trig | 50 |
| FLOOR.MATH | Rounds a number down using Excel's FLOOR.MATH rules. | Math & Trig | 37 |
| FLOOR.PRECISE | Rounds a number down to the nearest significance, ignoring the sign of the significance. | Math & Trig | 36 |
| FORECAST | Predicts a future value using linear regression. | Statistical | 35 |
| FORECAST.LINEAR | Predicts a future value using linear regression. | Statistical | 30 |
| FORMULATEXT | Returns the formula text from a referenced cell. | Information | 39 |
| FREQUENCY | Returns a frequency distribution for numeric bins. | Statistical | 69 |
| FV | Returns the future value of an investment or annuity. | Financial | 13 |
| FVSCHEDULE | Returns a future value after applying a schedule of interest rates. | Financial | 9 |
| GAMMA | Returns the gamma function value. | Math & Trig | 31 |
| GAMMA.DIST | Returns the gamma distribution. | Statistical | 41 |
| GAMMA.INV | Returns the inverse of the gamma cumulative distribution. | Statistical | 34 |
| GAMMALN | Returns the natural logarithm of the gamma function. | Statistical | 23 |
| GAMMALN.PRECISE | Returns the natural logarithm of the gamma function using the precise definition. | Statistical | 30 |
| GAUSS | Returns the probability that a standard normal variable lies between the mean and a value. | Statistical | 28 |
| GCD | Returns the greatest common divisor. | Math & Trig | 37 |
| GEOMEAN | Returns the geometric mean of positive values. | Statistical | 24 |
| GESTEP | Returns 1 when a number meets or exceeds a threshold, otherwise 0. | Engineering | 29 |
| GROWTH | Fits or projects values along an exponential trend. | Statistical | 13 |
| HARMEAN | Returns the harmonic mean of positive values. | Statistical | 21 |
| HEX2BIN | Converts a hexadecimal number to binary. | Engineering | 38 |
| HEX2DEC | Converts a hexadecimal number to decimal. | Engineering | 39 |
| HEX2OCT | Converts a hexadecimal number to octal. | Engineering | 35 |
| HLOOKUP | Looks up a value across the top row of a table and returns a value from a specified row. | Lookup & Reference | 102 |
| HOUR | Returns the hour from a time or datetime value. | Date & Time | 16 |
| HSTACK | Stacks arrays horizontally. | Lookup & Reference | 22 |
| HYPERLINK | Creates a clickable hyperlink. | Lookup & Reference | 46 |
| HYPGEOM.DIST | Returns the hypergeometric distribution. | Statistical | 64 |
| IF | Returns one value when a condition is TRUE and another when it is FALSE. | Logical | 41 |
| IFERROR | Returns a fallback value when a formula returns an error. | Logical | 85 |
| IFNA | Returns a fallback value when a formula returns #N/A. | Information | 45 |
| IFS | Evaluates conditions in order and returns the value for the first TRUE condition. | Logical | 32 |
| IMABS | Returns the absolute value of a complex number. | Engineering | 38 |
| IMAGINARY | Returns the imaginary coefficient of a complex number. | Engineering | 38 |
| IMARGUMENT | Returns the argument of a complex number. | Engineering | 47 |
| IMCONJUGATE | Returns the complex conjugate of a complex number. | Engineering | 47 |
| IMCOS | Returns the cosine of a complex number. | Engineering | 30 |
| IMCOSH | Returns the hyperbolic cosine of a complex number. | Engineering | 27 |
| IMCOT | Returns the cotangent of a complex number. | Engineering | 25 |
| IMCSC | Returns the cosecant of a complex number. | Engineering | 25 |
| IMCSCH | Returns the hyperbolic cosecant of a complex number. | Engineering | 25 |
| IMDIV | Returns the quotient of two complex numbers. | Engineering | 34 |
| IMEXP | Returns e raised to a complex power. | Engineering | 30 |
| IMLN | Returns the natural logarithm of a complex number. | Engineering | 30 |
| IMLOG10 | Returns the base-10 logarithm of a complex number. | Engineering | 30 |
| IMLOG2 | Returns the base-2 logarithm of a complex number. | Engineering | 31 |
| IMPOWER | Raises a complex number to a power. | Engineering | 39 |
| IMPRODUCT | Returns the product of complex numbers. | Engineering | 36 |
| IMREAL | Returns the real coefficient of a complex number. | Engineering | 38 |
| IMSEC | Returns the secant of a complex number. | Engineering | 29 |
| IMSECH | Returns the hyperbolic secant of a complex number. | Engineering | 25 |
| IMSIN | Returns the sine of a complex number. | Engineering | 30 |
| IMSINH | Returns the hyperbolic sine of a complex number. | Engineering | 27 |
| IMSQRT | Returns the square root of a complex number. | Engineering | 32 |
| IMSUB | Subtracts one complex number from another. | Engineering | 34 |
| IMSUM | Returns the sum of complex numbers. | Engineering | 37 |
| IMTAN | Returns the tangent of a complex number. | Engineering | 28 |
| INDEX | Returns a value or subrange at a row and column position. | Lookup & Reference | 143 |
| INDIRECT | Turns text into a reference and returns its value. | Lookup & Reference | 63 |
| INT | Rounds a number down to the nearest integer. | Math & Trig | 27 |
| INTERCEPT | Returns the y-intercept of a linear regression line. | Statistical | 39 |
| INTRATE | Returns the interest rate for a fully invested security. | Financial | 11 |
| IPMT | Returns the interest portion of a payment for a given period. | Financial | 4 |
| IRR | Returns the internal rate of return for periodic cash flows. | Financial | 19 |
| ISBLANK | Returns TRUE when a value is blank. | Information | 23 |
| ISERR | Returns TRUE for any error except #N/A. | Information | 22 |
| ISERROR | Returns TRUE for any error value. | Information | 22 |
| ISEVEN | Returns TRUE when a number is even. | Information | 33 |
| ISFORMULA | Returns TRUE when a referenced cell contains a formula. | Information | 34 |
| ISLOGICAL | Returns TRUE when a value is logical. | Information | 27 |
| ISNA | Returns TRUE when a value is #N/A. | Information | 30 |
| ISNONTEXT | Returns TRUE when a value is not text. | Information | 26 |
| ISNUMBER | Returns TRUE when a value is numeric. | Information | 25 |
| ISO.CEILING | Rounds a number up using ISO.CEILING rules. | Math & Trig | 48 |
| ISODD | Returns TRUE when a number is odd. | Information | 33 |
| ISOWEEKNUM | Returns the ISO week number for a date. | Date & Time | 26 |
| ISPMT | Returns the interest paid during a period for straight-line amortization. | Financial | 17 |
| ISREF | Returns TRUE when a value is a reference. | Information | 26 |
| ISTEXT | Returns TRUE when a value is text. | Information | 25 |
| KURT | Returns the kurtosis of a data set. | Statistical | 22 |
| LAMBDA | Defines a reusable custom formula from parameters and an expression. | Logical | 14 |
| LARGE | Returns the k-th largest value in a data set. | Statistical | 31 |
| LCM | Returns the least common multiple. | Math & Trig | 39 |
| LEFT | Returns characters from the left side of a text value. | Text | 42 |
| LEN | Returns the number of characters in text. | Text | 37 |
| LET | Assigns names to intermediate values within a formula. | Logical | 48 |
| LINEST | Returns linear regression statistics. | Statistical | 3 |
| LN | Returns the natural logarithm of a number. | Math & Trig | 32 |
| LOG | Returns the logarithm of a number in a chosen base. | Math & Trig | 35 |
| LOG10 | Returns the base-10 logarithm of a number. | Math & Trig | 31 |
| LOGEST | Returns exponential regression statistics. | Statistical | 2 |
| LOGNORM.DIST | Returns the lognormal distribution. | Statistical | 31 |
| LOGNORM.INV | Returns the inverse of the lognormal cumulative distribution. | Statistical | 31 |
| LOOKUP | Looks up a value in a one-dimensional range and returns the matching result. | Lookup & Reference | 24 |
| LOWER | Converts text to lowercase. | Text | 48 |
| MAKEARRAY | Builds an array by applying a lambda to row and column indexes. | Logical | 33 |
| MAP | Applies a lambda element-by-element across one or more arrays. | Logical | 27 |
| MATCH | Returns the relative position of a lookup value in a range or array. | Lookup & Reference | 81 |
| MAX | Returns the largest numeric value. | Statistical | 42 |
| MAXA | Returns the largest value, counting logical values and text coercions. | Statistical | 16 |
| MAXIFS | Returns the maximum value that matches all supplied conditions. | Statistical | 46 |
| MDETERM | Returns the determinant of a matrix. | Math & Trig | 33 |
| MDURATION | Returns the modified duration of a security. | Financial | 22 |
| MEDIAN | Returns the median of a data set. | Statistical | 27 |
| MID | Returns characters from the middle of a text value. | Text | 50 |
| MIN | Returns the smallest numeric value. | Statistical | 37 |
| MINA | Returns the smallest value, counting logical values and text coercions. | Statistical | 16 |
| MINIFS | Returns the minimum value that matches all supplied conditions. | Statistical | 50 |
| MINUTE | Returns the minute from a time or datetime value. | Date & Time | 16 |
| MINVERSE | Returns the inverse of a matrix. | Math & Trig | 58 |
| MIRR | Returns the modified internal rate of return. | Financial | 65 |
| MMULT | Returns the matrix product of two arrays. | Math & Trig | 67 |
| MOD | Returns the remainder after division. | Math & Trig | 60 |
| MODE | Returns the most frequently occurring value. | Statistical | 19 |
| MODE.MULT | Returns all modes in a data set. | Statistical | 22 |
| MODE.SNGL | Returns a single mode from a data set. | Statistical | 19 |
| MONTH | Returns the month number from a date. | Date & Time | 45 |
| MROUND | Rounds a number to the nearest multiple. | Math & Trig | 38 |
| MULTINOMIAL | Returns the multinomial of a set of numbers. | Math & Trig | 36 |
| MUNIT | Returns a unit matrix of a requested size. | Math & Trig | 38 |
| N | Converts a value to a number using spreadsheet coercion rules. | Information | 19 |
| NA | Returns the #N/A error value. | Information | 17 |
| NEGBINOM.DIST | Returns the negative binomial distribution. | Statistical | 62 |
| NETWORKDAYS | Counts working days between two dates. | Date & Time | 53 |
| NETWORKDAYS.INTL | Counts working days between two dates using a custom weekend pattern. | Date & Time | 44 |
| NOMINAL | Returns the nominal annual interest rate for an effective rate. | Financial | 17 |
| NORM.DIST | Returns the normal distribution. | Statistical | 29 |
| NORM.INV | Returns the inverse of the normal cumulative distribution. | Statistical | 28 |
| NORM.S.DIST | Returns the standard normal distribution. | Statistical | 27 |
| NORM.S.INV | Returns the inverse of the standard normal cumulative distribution. | Statistical | 27 |
| NOT | Returns the opposite of a logical value. | Logical | 45 |
| NOW | Returns the current date and time. | Date & Time | 2 |
| NPER | Returns the number of periods for an investment or loan. | Financial | 14 |
| NPV | Returns the net present value of periodic cash flows. | Financial | 25 |
| NUMBERVALUE | Converts locale-formatted text to a number. | Text | 27 |
| OCT2BIN | Converts a octal number to binary. | Engineering | 38 |
| OCT2DEC | Converts a octal number to decimal. | Engineering | 43 |
| OCT2HEX | Converts a octal number to hexadecimal. | Engineering | 36 |
| ODD | Rounds a number away from zero to the nearest odd integer. | Math & Trig | 40 |
| OFFSET | Returns a reference offset from a starting reference. | Lookup & Reference | 29 |
| OR | Returns TRUE when any supplied condition is TRUE. | Logical | 51 |
| PDURATION | Returns the periods required for an investment to reach a target value. | Financial | 23 |
| PEARSON | Returns the Pearson correlation coefficient. | Statistical | 47 |
| PERCENTILE | Returns the percentile of a data set. | Statistical | 33 |
| PERCENTILE.EXC | Returns the exclusive percentile of a data set. | Statistical | 31 |
| PERCENTILE.INC | Returns the inclusive percentile of a data set. | Statistical | 46 |
| PERCENTRANK | Returns the rank of a value as a percentage of a data set. | Statistical | 27 |
| PERCENTRANK.EXC | Returns the exclusive percentage rank of a value in a data set. | Statistical | 29 |
| PERCENTRANK.INC | Returns the inclusive percentage rank of a value in a data set. | Statistical | 52 |
| PERMUT | Returns the number of permutations for a number of objects. | Math & Trig | 39 |
| PERMUTATIONA | Returns the number of permutations with repetitions. | Statistical | 41 |
| PHI | Returns the standard normal density at a value. | Statistical | 35 |
| PI | Returns the value of pi. | Math & Trig | 13 |
| PMT | Returns the periodic payment for a loan or annuity. | Financial | 5 |
| POISSON.DIST | Returns the Poisson distribution. | Statistical | 26 |
| POWER | Returns a number raised to a power. | Math & Trig | 49 |
| PPMT | Returns the principal portion of a payment for a given period. | Financial | 2 |
| PRICE | Returns the price per 100 face value of a coupon-paying security. | Financial | 24 |
| PRICEDISC | Returns the price per 100 face value of a discounted security. | Financial | 31 |
| PRICEMAT | Returns the price per 100 face value of a security that pays interest at maturity. | Financial | 32 |
| PROB | Returns the probability that values fall within a range. | Statistical | 27 |
| PRODUCT | Returns the product of supplied numbers. | Math & Trig | 37 |
| PROPER | Capitalizes words in text. | Text | 54 |
| PV | Returns the present value of an investment or loan. | Financial | 4 |
| QUARTILE | Returns a quartile of a data set. | Statistical | 37 |
| QUARTILE.EXC | Returns the exclusive quartile of a data set. | Statistical | 27 |
| QUARTILE.INC | Returns the inclusive quartile of a data set. | Statistical | 65 |
| QUOTIENT | Returns the integer portion of a division. | Math & Trig | 68 |
| RADIANS | Converts degrees to radians. | Math & Trig | 30 |
| RAND | Returns a random number between 0 and 1. | Math & Trig | 13 |
| RANDARRAY | Returns an array of random numbers. | Math & Trig | 51 |
| RANDBETWEEN | Returns a random integer between two bounds. | Math & Trig | 26 |
| RANK | Returns the rank of a number in a list. | Statistical | 34 |
| RANK.AVG | Returns the average rank of a number in a list with ties. | Statistical | 21 |
| RANK.EQ | Returns the rank of a number in a list with ties sharing the same rank. | Statistical | 46 |
| RATE | Returns the interest rate per period for an annuity. | Financial | 6 |
| RECEIVED | Returns the amount received at maturity for a fully invested security. | Financial | 11 |
| REDUCE | Folds an array to a single result by repeatedly applying a lambda. | Logical | 26 |
| REGEXEXTRACT | Extracts one or more substrings that match a regular expression. | Text | 19 |
| REGEXREPLACE | Replaces substrings that match a regular expression. | Text | 20 |
| REGEXTEST | Returns TRUE if a string matches a regular expression. | Text | 12 |
| REPLACE | Replaces characters within text at a given position. | Text | 82 |
| REPT | Repeats text a specified number of times. | Text | 26 |
| RIGHT | Returns characters from the right side of a text value. | Text | 41 |
| ROMAN | Converts an Arabic number to Roman numeral text. | Text | 38 |
| ROUND | Rounds a number to a requested number of digits. | Math & Trig | 42 |
| ROUNDDOWN | Rounds a number toward zero. | Math & Trig | 36 |
| ROUNDUP | Rounds a number away from zero. | Math & Trig | 36 |
| ROW | Returns the row number of a reference. | Information | 42 |
| ROWS | Returns the number of rows in a reference or array. | Information | 44 |
| RRI | Returns an equivalent interest rate for investment growth. | Financial | 21 |
| RSQ | Returns the square of the Pearson correlation coefficient. | Statistical | 42 |
| SCAN | Returns the running accumulation produced by a lambda over an array. | Logical | 26 |
| SEARCH | Returns the position of one text value inside another, case-insensitive. | Text | 51 |
| SEC | Returns the secant of an angle. | Math & Trig | 26 |
| SECH | Returns the hyperbolic secant of a number. | Math & Trig | 28 |
| SECOND | Returns the second from a time or datetime value. | Date & Time | 16 |
| SEQUENCE | Returns a sequence of numbers as an array. | Math & Trig | 116 |
| SERIESSUM | Returns the sum of a power series. | Math & Trig | 27 |
| SHEET | Returns the sheet index of a reference. | Information | 14 |
| SHEETS | Returns the number of sheets in a reference or workbook scope. | Information | 8 |
| SIGN | Returns the sign of a number. | Math & Trig | 30 |
| SIN | Returns the sine of an angle. | Math & Trig | 33 |
| SINGLE | Returns a lookup or reference result related to SINGLE. | Lookup & Reference | 2 |
| SINH | Returns the hyperbolic sine of a number. | Math & Trig | 23 |
| SKEW | Returns the sample skewness of a data set. | Statistical | 23 |
| SKEW.P | Returns the population skewness of a data set. | Statistical | 27 |
| SLN | Returns straight-line depreciation for one period. | Financial | 53 |
| SLOPE | Returns the slope of a linear regression line. | Statistical | 39 |
| SMALL | Returns the k-th smallest value in a data set. | Statistical | 29 |
| SORT | Sorts an array by row or column order. | Logical | 52 |
| SORTBY | Sorts an array by one or more companion arrays. | Logical | 8 |
| SQRT | Returns the square root of a number. | Math & Trig | 29 |
| SQRTPI | Returns the square root of a number multiplied by pi. | Math & Trig | 25 |
| STANDARDIZE | Returns a normalized value from a mean and standard deviation. | Statistical | 34 |
| STDEV | Returns the sample standard deviation. | Statistical | 28 |
| STDEV.P | Returns the population standard deviation. | Statistical | 32 |
| STDEV.S | Returns the sample standard deviation. | Statistical | 45 |
| STDEVA | Returns the sample standard deviation including logical values and text coercions. | Statistical | 28 |
| STDEVP | Returns the population standard deviation. | Statistical | 11 |
| STDEVPA | Returns the population standard deviation including logical values and text coercions. | Statistical | 30 |
| STEYX | Returns the standard error of predicted y values in regression. | Statistical | 41 |
| SUBSTITUTE | Replaces matching text within a string. | Text | 69 |
| SUBTOTAL | Returns a subtotal using a selected aggregate function. | Math & Trig | 76 |
| SUM | Returns the sum of supplied numbers. | Statistical | 124 |
| SUMIF | Returns the sum of values that match one condition. | Statistical | 77 |
| SUMIFS | Returns the sum of values that match all supplied conditions. | Statistical | 65 |
| SUMPRODUCT | Returns the sum of pairwise products across arrays. | Statistical | 88 |
| SUMSQ | Returns the sum of squares of the supplied values. | Statistical | 30 |
| SUMX2MY2 | Returns the sum of the difference of squares of paired arrays. | Math & Trig | 55 |
| SUMX2PY2 | Returns the sum of the sum of squares of paired arrays. | Math & Trig | 57 |
| SUMXMY2 | Returns the sum of squares of differences of paired arrays. | Math & Trig | 57 |
| SWITCH | Matches an expression against a list of values and returns the corresponding result. | Logical | 28 |
| SYD | Returns sum-of-years'-digits depreciation for a period. | Financial | 27 |
| T | Returns text when a value is text, otherwise an empty string. | Text | 21 |
| T.DIST | Returns the Student's t-distribution. | Statistical | 41 |
| T.DIST.2T | Returns the two-tailed Student's t-distribution. | Statistical | 32 |
| T.DIST.RT | Returns the right-tailed Student's t-distribution. | Statistical | 27 |
| T.INV | Returns the inverse of the Student's t cumulative distribution. | Statistical | 33 |
| T.INV.2T | Returns the inverse of the two-tailed Student's t-distribution. | Statistical | 37 |
| T.TEST | Returns the probability associated with a Student's t-test. | Statistical | 56 |
| TAKE | Returns a requested number of rows or columns from an array. | Lookup & Reference | 31 |
| TAN | Returns the tangent of an angle. | Math & Trig | 23 |
| TANH | Returns the hyperbolic tangent of a number. | Math & Trig | 27 |
| TBILLEQ | Returns the bond-equivalent yield for a Treasury bill. | Financial | 34 |
| TBILLPRICE | Returns the price per 100 face value for a Treasury bill. | Financial | 36 |
| TBILLYIELD | Returns the yield for a Treasury bill. | Financial | 35 |
| TEXT | Formats a value as text using a number format pattern. | Text | 134 |
| TEXTAFTER | Returns the text that appears after a delimiter. | Text | 34 |
| TEXTBEFORE | Returns the text that appears before a delimiter. | Text | 34 |
| TEXTJOIN | Joins text values with a delimiter. | Text | 69 |
| TEXTSPLIT | Splits text into rows or columns around delimiters. | Text | 89 |
| TIME | Builds a time value from hour, minute, and second numbers. | Date & Time | 47 |
| TIMEVALUE | Converts time text to a serial time value. | Date & Time | 48 |
| TOCOL | Flattens an array into a single column. | Lookup & Reference | 23 |
| TODAY | Returns the current date. | Date & Time | 2 |
| TOROW | Flattens an array into a single row. | Lookup & Reference | 47 |
| TRANSPOSE | Swaps the rows and columns of an array. | Lookup & Reference | 58 |
| TREND | Returns values along a linear trend. | Statistical | 10 |
| TRIM | Removes extra spaces from text. | Text | 44 |
| TRIMMEAN | Returns the mean after trimming values from both tails of a data set. | Statistical | 23 |
| TRUE | Returns the logical value TRUE. | Logical | 10 |
| TRUNC | Truncates a number to an integer or fixed precision. | Math & Trig | 34 |
| TYPE | Returns a numeric code describing a value's type. | Information | 25 |
| UNICHAR | Returns the Unicode character for a code point. | Text | 48 |
| UNICODE | Returns the Unicode code point for the first character in text. | Text | 36 |
| UNIQUE | Returns distinct rows or columns from an array. | Lookup & Reference | 56 |
| UPPER | Converts text to uppercase. | Text | 62 |
| VALUE | Converts text that looks like a number into a numeric value. | Text | 31 |
| VALUETOTEXT | Converts a value to a text representation. | Text | 47 |
| VAR | Returns the sample variance. | Statistical | 26 |
| VAR.P | Returns the population variance. | Statistical | 26 |
| VAR.S | Returns the sample variance. | Statistical | 31 |
| VARA | Returns the sample variance including logical values and text coercions. | Statistical | 24 |
| VARP | Returns the population variance. | Statistical | 30 |
| VARPA | Returns the population variance including logical values and text coercions. | Statistical | 22 |
| VDB | Returns depreciation using the variable declining balance method. | Financial | 33 |
| VLOOKUP | Looks up a value in the first column of a table and returns a value from another column. | Lookup & Reference | 89 |
| VSTACK | Stacks arrays vertically. | Lookup & Reference | 19 |
| WEEKDAY | Returns the day of the week for a date. | Date & Time | 59 |
| WEEKNUM | Returns the week number of a date. | Date & Time | 39 |
| WEIBULL.DIST | Returns the Weibull distribution. | Statistical | 33 |
| WORKDAY | Returns a working day offset from a start date. | Date & Time | 42 |
| WORKDAY.INTL | Returns a working day offset using a custom weekend pattern. | Date & Time | 34 |
| WRAPCOLS | Wraps a vector into a two-dimensional array by columns. | Lookup & Reference | 21 |
| WRAPROWS | Wraps a vector into a two-dimensional array by rows. | Lookup & Reference | 20 |
| XIRR | Returns the internal rate of return for cash flows on irregular dates. | Financial | 38 |
| XLOOKUP | Looks up a value in one array and returns the matching value from another array. | Lookup & Reference | 114 |
| XMATCH | Returns the position of a lookup value with exact, wildcard, or binary search modes. | Lookup & Reference | 81 |
| XNPV | Returns the net present value of cash flows on irregular dates. | Financial | 42 |
| XOR | Returns TRUE when an odd number of supplied conditions are TRUE. | Logical | 34 |
| YEAR | Returns the year from a date. | Date & Time | 26 |
| YEARFRAC | Returns the fraction of a year between two dates. | Date & Time | 38 |
| YIELD | Returns the yield of a coupon-paying security. | Financial | 19 |
| YIELDDISC | Returns the annual yield of a discounted security. | Financial | 23 |
| YIELDMAT | Returns the annual yield of a security that pays interest at maturity. | Financial | 27 |
| Z.TEST | Returns the one-tailed probability value of a z-test. | Statistical | 38 |

# No Planned Support

These functions depend on the spreadsheet application's runtime environment, have side
effects, or require locale-specific behavior that cannot be reproduced in a server-side library.

| Function | Description | Category | Reason |
|----------|-------------|----------|--------|
| ASC | Converts full-width characters to half-width characters. | Text | Full-width → half-width conversion; behavior depends on the application's MBCS locale setting |
| BAHTTEXT | Converts a number to Thai Baht text. | Text | Converts numbers to Thai Baht text — extremely locale-specific |
| CALL | Calls a procedure in a dynamic link library or code resource. | User Defined | Calls procedures in dynamic-link libraries — security risk, OS-specific binary loading |
| CELL | Returns information about a cell's contents, location, or formatting. | Information | Many modes return application-specific environment info (filename, format codes from the UI) |
| CUBEKPIMEMBER | Returns a key performance indicator (KPI) property and displays the KPI name in the cell. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| CUBEMEMBER | Returns a member or tuple from the cube. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| CUBEMEMBERPROPERTY | Returns the value of a member property from the cube. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| CUBERANKEDMEMBER | Returns the nth, or ranked, member in a set. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| CUBESET | Defines a calculated set of members or tuples by sending a set expression to the cube on the server. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| CUBESETCOUNT | Returns the number of items in a set. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| CUBEVALUE | Returns an aggregated value from the cube. | Cube | Requires a live connection to an OLAP cube data source — not applicable to a server-side calc engine |
| DBCS | Converts half-width characters to full-width characters. | Text | Half-width → full-width conversion; behavior depends on the application's MBCS locale setting |
| DETECTLANGUAGE | Identifies the language of a specified text. | Text | Calls Microsoft's online translation/ML service — external network dependency |
| EUROCONVERT | Converts a number to euros, converts a number from euros to a euro member currency, or converts a number from one euro member currency to another by using the euro as an intermediary (triangulation). | User Defined | Excel add-in tied to a Euro currency conversion table — locale and add-in-version specific |
| FILTERXML | Extracts data from XML by XPath. | Web | Paired with WEBSERVICE; fetches/parses external XML |
| FINDB | Returns the byte position of one text value inside another. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| IMAGE | Returns an image from a given source. | Lookup & Reference | Returns an image rendered from a URL — werkbook is a calc engine, not a renderer |
| INFO | Returns environment information about the current workbook session. | Information | Returns application environment info (OS version, memory, app version) |
| LEFTB | Returns bytes from the left side of a text value. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| LENB | Returns the number of bytes in text. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| MIDB | Returns bytes from the middle of a text value. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| PHONETIC | Returns furigana text associated with a string. | Text | Returns Japanese furigana metadata from the IME — not stored in XLSX files |
| REGISTER.ID | Returns the register ID of the specified dynamic link library (DLL) or code resource that has been previously registered. | User Defined | Returns the register ID of a DLL or code resource — paired with CALL, OS-specific |
| REPLACEB | Replaces bytes within text at a given position. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| RIGHTB | Returns bytes from the right side of a text value. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| RTD | Retrieves real-time data from a program that supports COM automation. | Lookup & Reference | Connects to a real-time COM automation server — requires running external program |
| SEARCHB | Returns the byte position of one text value inside another, case-insensitive. | Text | Byte-position text function; behavior depends on the application's default language setting (MBCS vs SBCS) |
| STOCKHISTORY | Retrieves historical data about a financial instrument. | Information | Fetches live financial market data via Microsoft service — external network dependency |
| TRANSLATE | Translates a text from one language to another. | Text | Calls Microsoft's online translation service — external network dependency |
| WEBSERVICE | Returns data from a web service. | Web | Makes HTTP requests from formulas — security risk, side effects |

# Not Yet Implemented

The following **42** functions are not yet supported.

| Function | Description | Category |
|----------|-------------|----------|
| BETADIST | Returns the beta cumulative distribution function. | Statistical |
| BETAINV | Returns the inverse of the cumulative distribution function for a specified beta distribution. | Statistical |
| BINOMDIST | Returns the individual term binomial distribution probability. | Statistical |
| CHIDIST | Returns the one-tailed probability of the chi-squared distribution. | Statistical |
| CHIINV | Returns the inverse of the one-tailed probability of the chi-squared distribution. | Statistical |
| CHITEST | Returns the test for independence. | Statistical |
| CONFIDENCE | Returns the confidence interval for a population mean. | Statistical |
| CRITBINOM | Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value. | Statistical |
| EXPONDIST | Returns the exponential distribution. | Statistical |
| FDIST | Returns the F probability distribution. | Statistical |
| FINV | Returns the inverse of the F probability distribution. | Statistical |
| FORECAST.ETS | Predicts a future value using exponential smoothing. | Statistical |
| FORECAST.ETS.CONFINT | Returns a confidence interval for an exponential smoothing forecast. | Statistical |
| FORECAST.ETS.SEASONALITY | Returns the detected seasonality length for exponential smoothing. | Statistical |
| FORECAST.ETS.STAT | Returns forecast statistics from exponential smoothing. | Statistical |
| FTEST | Returns the result of an F-test. | Statistical |
| GAMMADIST | Returns the gamma distribution. | Statistical |
| GAMMAINV | Returns the inverse of the gamma cumulative distribution. | Statistical |
| GETPIVOTDATA | Returns a value from a pivot table by field and item labels. | Lookup & Reference |
| GROUPBY | Helps a user group, aggregate, sort, and filter data based on the fields you specify. | Lookup & Reference |
| HYPGEOMDIST | Returns the hypergeometric distribution. | Statistical |
| ISOMITTED | Checks whether the value in a LAMBDA is missing and returns TRUE or FALSE. | Information |
| LOGINV | Returns the inverse of the lognormal cumulative distribution function. | Statistical |
| LOGNORMDIST | Returns the cumulative lognormal distribution. | Statistical |
| NEGBINOMDIST | Returns the negative binomial distribution. | Statistical |
| NORMDIST | Returns the normal cumulative distribution. | Statistical |
| NORMINV | Returns the inverse of the normal cumulative distribution. | Statistical |
| NORMSDIST | Returns the standard normal cumulative distribution. | Statistical |
| NORMSINV | Returns the inverse of the standard normal cumulative distribution. | Statistical |
| ODDFPRICE | Returns the price of a security with an odd first period. | Financial |
| ODDFYIELD | Returns the yield of a security with an odd first period. | Financial |
| ODDLPRICE | Returns the price of a security with an odd last period. | Financial |
| ODDLYIELD | Returns the yield of a security with an odd last period. | Financial |
| PERCENTOF | Sums the values in the subset and divides it by all the values. | Math & Trig |
| PIVOTBY | Helps a user group, aggregate, sort, and filter data based on the row and column fields that you specify. | Lookup & Reference |
| POISSON | Returns the Poisson distribution. | Statistical |
| TDIST | Returns the Student's t-distribution. | Statistical |
| TINV | Returns the inverse of the Student's t-distribution. | Statistical |
| TRIMRANGE | Scans in from the edges of a range or array until it finds a non-blank cell (or value), it then excludes those blank rows or columns. | Lookup & Reference |
| TTEST | Returns the probability associated with a Student's t-test. | Statistical |
| WEIBULL | Returns the Weibull distribution. | Statistical |
| ZTEST | Returns the one-tailed probability-value of a z-test. | Statistical |
