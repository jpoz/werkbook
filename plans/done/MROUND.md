# Plan: MROUND Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/mround-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427

**Syntax:** `MROUND(number, multiple)`

**Description:** Returns a number rounded to the desired multiple.

**Arguments:**
- `number` (required) — the value to round
- `multiple` (required) — the multiple to round to

**Remarks:**
- Rounds away from zero when remainder ≥ half the multiple
- Number and multiple must have the same sign, else #NUM!
- If multiple is 0, returns 0

**Examples:**
- `=MROUND(10, 3)` → `9`
- `=MROUND(-10, -3)` → `-9`
- `=MROUND(1.3, 0.2)` → `1.4`
- `=MROUND(5, -2)` → #NUM!

## Implementation Steps

1. Add `"MROUND"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnMROUND(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 2 args
   - Coerce both to numbers
   - If multiple is 0, return 0
   - If signs differ, return #NUM!
   - Compute: `math.Round(number / multiple) * multiple`
3. Add `case "MROUND":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `MROUND(10, 3)` → `9`
- `MROUND(-10, -3)` → `-9`
- `MROUND(1.3, 0.2)` → `1.4`
- `MROUND(5, -2)` → #NUM!
- `MROUND(5, 0)` → `0`
- `MROUND(7.5, 5)` → `10`
