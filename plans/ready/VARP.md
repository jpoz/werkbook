# Plan: VARP Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/varp-function-26a541c4-ecee-464d-a731-bd4c575b1a6b

**Syntax:** `VARP(number1, [number2], ...)`

**Description:** Returns the variance of an entire population.

**Arguments:**
- `number1` (required) — first number argument corresponding to a population
- `number2, ...` (optional) — number arguments 2 to 255 corresponding to a population

**Remarks:**
- VARP assumes arguments are the entire population (use VAR for a sample)
- Arguments can be numbers, names, arrays, or references containing numbers
- Logical values and text representations typed directly are counted
- Arrays/references: only numbers counted; empty cells, logical values, text, error values are ignored
- Arguments that are error values or text that cannot be translated into numbers cause errors
- Formula: `Σ(x - x̄)² / n` where x̄ is the mean and n is the population size

**Examples:**
- `=VARP(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → `678.84`

## Implementation Steps

1. Add `"VARP"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnVARP(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate at least 1 arg
   - Collect all numeric values from args
   - Compute mean
   - Compute sum of squared deviations
   - Divide by n (not n-1)
   - If n < 1, return #DIV/0!
3. Add `case "VARP":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `VARP(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)` → approximately `678.84`
- `VARP(2, 4, 4, 4, 5, 5, 7, 9)` → approximately `4.0`
- `VARP(5)` → `0` (single value, zero variance)
- `VARP(3, 3, 3)` → `0`
