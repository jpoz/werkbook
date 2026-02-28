# Plan: RADIANS Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/radians-function-ac409508-3d48-45f5-ac02-1497c92de5bf

**Syntax:** `RADIANS(angle)`

**Description:** Converts degrees to radians.

**Arguments:**
- `angle` (required) — angle in degrees to convert

**Formula:** `radians = angle * π / 180`

**Examples:**
- `=RADIANS(270)` → `4.712389` (3π/2)
- `=RADIANS(180)` → `3.141593` (π)

## Implementation Steps

1. Add `"RADIANS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnRADIANS(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 1 arg
   - Support array lifting via `liftUnary`
   - Coerce to number, return `n * math.Pi / 180`
3. Add `case "RADIANS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `RADIANS(180)` → `≈3.14159265`
- `RADIANS(270)` → `≈4.71238898`
- `RADIANS(0)` → `0`
- `RADIANS(360)` → `≈6.28318530` (2π)
- `RADIANS("text")` → #VALUE!
