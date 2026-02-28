# Plan: STDEVP Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/stdevp-function-1f7c1c88-1bec-4422-8242-e9f7dc8bb195

**Syntax:** `STDEVP(number1, [number2], ...)`

**Description:** Returns the standard deviation of an entire population. Uses the "n" method (population standard deviation).

**Arguments:**
- `number1` (required) — first number argument corresponding to the population
- `number2, ...` (optional) — number arguments 2 to 255 corresponding to the population; can also use arrays/references

**Remarks:**
- STDEVP assumes arguments are the entire population (use STDEV for a sample)
- Arguments can be numbers, names, arrays, or references containing numbers
- Logical values and text representations typed directly are counted
- Arrays/references: only numbers counted; empty cells, logical values, text, error values are ignored
- Arguments that are error values or text that cannot be translated into numbers cause errors
- Formula: `sqrt(Σ(x - x̄)² / n)` where x̄ is the mean and n is the population size

**Examples:**
- `=STDEVP(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → `26.05`

## Implementation Steps

1. Add `"STDEVP"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnSTDEVP(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate at least 1 arg
   - Collect all numeric values from args
   - Compute mean of collected values
   - Compute sum of squared deviations from mean
   - Divide by n (not n-1) and take square root
   - If n < 1, return #DIV/0!
3. Add `case "STDEVP":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `STDEVP(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → approximately `26.05`
- `STDEVP(2, 4, 4, 4, 5, 5, 7, 9)` → approximately `2.0`
- `STDEVP(5)` → `0` (single value, zero deviation)
- `STDEVP(3, 3, 3)` → `0`
