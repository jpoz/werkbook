# Plan: STDEV Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0

**Syntax:** `STDEV(number1, [number2], ...)`

**Description:** Returns the standard deviation of a sample. Uses the "n-1" method (sample standard deviation).

**Arguments:**
- `number1` (required) — first number argument corresponding to a sample
- `number2, ...` (optional) — number arguments 2 to 255 corresponding to a sample; can also use arrays/references

**Remarks:**
- STDEV assumes arguments are a sample (use STDEVP for entire population)
- Arguments can be numbers, names, arrays, or references containing numbers
- Logical values and text representations of numbers typed directly into the argument list are counted
- If an argument is an array or reference, only numbers are counted; empty cells, logical values, text, and error values in arrays/references are ignored
- Arguments that are error values or text that cannot be translated into numbers cause errors
- Formula: `sqrt(Σ(x - x̄)² / (n - 1))` where x̄ is the sample mean and n is the sample size

**Examples:**
- `=STDEV(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → `27.46`

## Implementation Steps

1. Add `"STDEV"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnSTDEV(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate at least 1 arg
   - Collect all numeric values from args using `iterateNumeric` (or similar traversal)
   - Compute mean of collected values
   - Compute sum of squared deviations from mean
   - Divide by (n - 1) and take square root
   - If n < 2, return #DIV/0!
3. Add `case "STDEV":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `STDEV(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → approximately `27.46`
- `STDEV(2, 4, 4, 4, 5, 5, 7, 9)` → approximately `2.138`
- `STDEV(5)` → #DIV/0! (need at least 2 values)
- `STDEV(3, 3, 3)` → `0`
