# Plan: EVEN Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/even-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9

**Syntax:** `EVEN(number)`

**Description:** Rounds a number up to the nearest even integer. Rounds away from zero.

**Arguments:**
- `number` (required) — the value to round

**Remarks:**
- Non-numeric input returns #VALUE!
- Rounds away from zero: EVEN(-1) → -2
- Already-even integers unchanged: EVEN(2) → 2

**Examples:**
- `=EVEN(1.5)` → `2`
- `=EVEN(3)` → `4`
- `=EVEN(2)` → `2`
- `=EVEN(-1)` → `-2`
- `=EVEN(0)` → `0`

## Implementation Steps

1. Add `"EVEN"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnEVEN(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number, round away from zero to nearest even
   - For positive: `math.Ceil(n/2) * 2`
   - For negative: `math.Floor(n/2) * 2`
3. Add `case "EVEN":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `EVEN(1.5)` → `2`
- `EVEN(3)` → `4`
- `EVEN(2)` → `2`
- `EVEN(-1)` → `-2`
- `EVEN(-2)` → `-2`
- `EVEN(0)` → `0`
- `EVEN("text")` → #VALUE!
