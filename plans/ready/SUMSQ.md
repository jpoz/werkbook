# Plan: SUMSQ Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/sumsq-function-e3313c02-51cc-4963-aae6-31442d9ec307

**Syntax:** `SUMSQ(number1, [number2], ...)`

**Description:** Returns the sum of the squares of the arguments.

**Arguments:**
- `number1` (required) — first value to square and sum
- `number2, ...` (optional) — up to 255 additional values
- Accepts numbers, arrays, and ranges

**Remarks:**
- In arrays/references, only numeric values are counted
- Text that cannot convert to numbers causes errors
- Logical values and text in references are ignored

**Examples:**
- `=SUMSQ(3, 4)` → `25` (9 + 16)
- `=SUMSQ(1, 2, 3)` → `14` (1 + 4 + 9)

## Implementation Steps

1. Add `"SUMSQ"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnSUMSQ(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate at least 1 arg
   - Flatten arrays, coerce to numbers (follow SUM pattern for handling arrays)
   - Square each value and accumulate sum
3. Add `case "SUMSQ":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `SUMSQ(3, 4)` → `25`
- `SUMSQ(1, 2, 3)` → `14`
- `SUMSQ(0)` → `0`
- `SUMSQ(-3, 4)` → `25`
- `SUMSQ(5)` → `25`
