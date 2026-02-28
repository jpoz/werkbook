# Plan: MODE Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/mode-function-e45192ce-9122-4980-82ed-4bdc34973120

**Syntax:** `MODE(number1, [number2], ...)`

**Description:** Returns the most frequently occurring (mode) value in a dataset.

**Arguments:**
- `number1` (required) — first number argument for mode calculation
- `number2, ...` (optional) — number arguments 2 to 255; can also use arrays/references

**Remarks:**
- Arguments can be numbers, names, arrays, or references containing numbers
- If array/reference contains text, logical values, or empty cells, those are ignored; cells with zero are included
- Arguments that are error values or text that can't be translated into numbers cause errors
- If the data set contains no duplicate data points, MODE returns #N/A

**Examples:**
- `=MODE(5.6, 4, 4, 3, 2, 4)` → `4`

## Implementation Steps

1. Add `"MODE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnMODE(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate at least 1 arg
   - Collect all numeric values from args (expanding arrays/ranges)
   - Count frequency of each value using a map
   - Find value with highest frequency (must be >= 2, i.e., at least one duplicate)
   - If no duplicates exist, return #N/A
   - If tie, return the value that appears first in the data (Excel behavior)
3. Add `case "MODE":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `MODE(5.6, 4, 4, 3, 2, 4)` → `4`
- `MODE(1, 2, 3, 4, 5)` → #N/A (no duplicates)
- `MODE(1, 2, 2, 3, 3)` → `2` (tie goes to first encountered)
- `MODE(7, 7, 7)` → `7`
- `MODE(1.5, 1.5, 2.5, 2.5, 2.5)` → `2.5`
