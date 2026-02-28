# Plan: T Function

## Category
Text (formula/functions_text.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/t-function-fb83aeec-45e7-4924-af95-53e073541228

**Syntax:** `T(value)`

**Description:** Returns the text if value is text, otherwise returns empty string.

**Arguments:**
- `value` (required) — the value to test

**Conversion rules:**
- Text → returns the text unchanged
- Number, Boolean, Error, Empty → returns "" (empty string)

**Examples:**
- `=T("Rainfall")` → `"Rainfall"`
- `=T(19)` → `""`
- `=T(TRUE)` → `""`

## Implementation Steps

1. Add `"T"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnT(args []Value) (Value, error)` in `formula/functions_text.go`:
   - Validate exactly 1 arg
   - If value type is TypeString, return as-is
   - Otherwise return `StringVal("")`
3. Add `case "T":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_text_test.go`

## Test Cases
- `T("hello")` → `"hello"`
- `T(123)` → `""`
- `T(TRUE)` → `""`
- `T("")` → `""`
