# Plan: VAR Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/var-function-1f2b7ab2-954d-4e17-ba2c-9e58b15a7da2

**Syntax:** `VAR(number1, [number2], ...)`

**Description:** Returns the variance of a sample (estimates variance based on a sample).

**Arguments:**
- `number1` (required) — first number argument corresponding to a sample
- `number2, ...` (optional) — number arguments 2 to 255 corresponding to a sample

**Remarks:**
- VAR assumes arguments are a sample (use VARP for entire population)
- Arguments can be numbers, names, arrays, or references containing numbers
- Logical values and text representations typed directly are counted
- Arrays/references: only numbers counted; empty cells, logical values, text, error values are ignored
- Arguments that are error values or text that cannot be translated into numbers cause errors
- Formula: `Σ(x - x̄)² / (n - 1)` where x̄ is the sample mean and n is the sample size

**Examples:**
- `=VAR(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → `754.27`

## Implementation Steps

1. Add `"VAR"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnVAR(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate at least 1 arg
   - Collect all numeric values from args
   - Compute mean
   - Compute sum of squared deviations
   - Divide by (n - 1)
   - If n < 2, return #DIV/0!
3. Add `case "VAR":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `VAR(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → approximately `754.27`
- `VAR(2, 4, 4, 4, 5, 5, 7, 9)` → approximately `4.571`
- `VAR(5)` → #DIV/0! (need at least 2 values)
- `VAR(3, 3, 3)` → `0`
