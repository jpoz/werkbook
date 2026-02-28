# Plan: QUOTIENT Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/quotient-function-9f7bf099-2a18-4282-8fa4-65290cc99dee

**Syntax:** `QUOTIENT(numerator, denominator)`

**Description:** Returns the integer portion of a division (truncates toward zero).

**Arguments:**
- `numerator` (required) — the dividend
- `denominator` (required) — the divisor

**Remarks:**
- Non-numeric args return #VALUE!
- Division by zero returns #DIV/0!

**Examples:**
- `=QUOTIENT(5, 2)` → `2`
- `=QUOTIENT(4.5, 3.1)` → `1`
- `=QUOTIENT(-10, 3)` → `-3`

## Implementation Steps

1. Add `"QUOTIENT"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnQUOTIENT(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 2 args
   - Coerce both to numbers
   - If denominator is 0, return #DIV/0!
   - Return `math.Trunc(numerator / denominator)`
3. Add `case "QUOTIENT":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `QUOTIENT(5, 2)` → `2`
- `QUOTIENT(4.5, 3.1)` → `1`
- `QUOTIENT(-10, 3)` → `-3`
- `QUOTIENT(10, 0)` → #DIV/0!
- `QUOTIENT(7, 7)` → `1`
- `QUOTIENT("a", 2)` → #VALUE!
