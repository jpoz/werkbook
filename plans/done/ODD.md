# Plan: ODD Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/odd-function-deae64eb-e08a-4c88-8b40-6d0b42575c98

**Syntax:** `ODD(number)`

**Description:** Rounds a number up to the nearest odd integer. Rounds away from zero.

**Arguments:**
- `number` (required) — the value to round

**Remarks:**
- Non-numeric input returns #VALUE!
- Rounds away from zero: ODD(-2) → -3
- Already-odd integers unchanged: ODD(3) → 3

**Examples:**
- `=ODD(1.5)` → `3`
- `=ODD(3)` → `3`
- `=ODD(2)` → `3`
- `=ODD(-1)` → `-1`
- `=ODD(-2)` → `-3`
- `=ODD(0)` → `1`

## Implementation Steps

1. Add `"ODD"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnODD(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number, round away from zero to nearest odd
   - Special case: ODD(0) → 1
   - For positive: if already odd integer, keep; else ceil to next odd
   - For negative: if already odd integer, keep; else floor to next odd (away from zero)
3. Add `case "ODD":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `ODD(1.5)` → `3`
- `ODD(3)` → `3`
- `ODD(2)` → `3`
- `ODD(-1)` → `-1`
- `ODD(-2)` → `-3`
- `ODD(0)` → `1`
- `ODD("text")` → #VALUE!
