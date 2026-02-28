# Plan: LCM Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/lcm-function-7152b67a-8bb5-4075-ae5c-06ede5563c94

**Syntax:** `LCM(number1, [number2], ...)`

**Description:** Returns the least common multiple of integers.

**Arguments:**
- `number1` (required) — first value
- `number2, ...` (optional) — up to 255 values; non-integers truncated

**Remarks:**
- Non-numeric args return #VALUE!
- Negative args return #NUM!
- Results exceeding 2^53 return #NUM!
- LCM(n, 0) = 0

**Examples:**
- `=LCM(5, 2)` → `10`
- `=LCM(24, 36)` → `72`

## Implementation Steps

1. Add `"LCM"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnLCM(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate at least 1 arg
   - Coerce each arg to number, truncate to integer
   - Return #NUM! if any arg is negative
   - Compute: `lcm(a, b) = |a*b| / gcd(a, b)`; iterate across all args
   - If any arg is 0, result is 0
3. Add `case "LCM":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `LCM(5, 2)` → `10`
- `LCM(24, 36)` → `72`
- `LCM(3, 4, 5)` → `60`
- `LCM(5, 0)` → `0`
- `LCM(-5, 2)` → #NUM!
- `LCM(7)` → `7`
