# Plan: ISNONTEXT Function

## Category
Information (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665

**Syntax:** `ISNONTEXT(value)`

**Description:** Returns TRUE if value refers to any item that is not text. Returns TRUE for blank cells.

**Arguments:**
- `value` (required) — the value to test

**Remarks:**
- Returns TRUE for numbers, booleans, errors, and blank/empty cells
- Returns FALSE only for text values
- This is the logical inverse of ISTEXT, except both return TRUE for blanks (ISTEXT returns FALSE for blanks)

**Examples:**
- `=ISNONTEXT(123)` → `TRUE`
- `=ISNONTEXT("hello")` → `FALSE`
- `=ISNONTEXT(TRUE)` → `TRUE`
- `=ISNONTEXT("")` → `FALSE`

## Implementation Steps

1. Add `"ISNONTEXT"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnISNONTEXT(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - Check if arg[0].Type != TypeString
   - Return BoolValue(true) if not string type, BoolValue(false) if string
   - Empty cells (TypeEmpty) return TRUE
3. Add `case "ISNONTEXT":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `ISNONTEXT(123)` → `TRUE`
- `ISNONTEXT("hello")` → `FALSE`
- `ISNONTEXT(TRUE)` → `TRUE`
- `ISNONTEXT(1/0)` → `TRUE` (error is non-text)
- `ISNONTEXT("")` → `FALSE` (empty string is still text type)
