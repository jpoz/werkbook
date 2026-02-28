# Plan: ERROR.TYPE Function

## Category
Information (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/error-type-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa

**Syntax:** `ERROR.TYPE(error_val)`

**Description:** Returns a number corresponding to one of the error values in Excel, or returns #N/A if no error exists.

**Arguments:**
- `error_val` (required) — the error value whose identifying number you want to find

**Return value mapping:**
- `#NULL!` → 1
- `#DIV/0!` → 2
- `#VALUE!` → 3
- `#REF!` → 4
- `#NAME?` → 5
- `#NUM!` → 6
- `#N/A` → 7
- Anything else → #N/A

**Examples:**
- `=ERROR.TYPE(#NULL!)` → `1`
- `=ERROR.TYPE(#DIV/0!)` → `2`
- `=ERROR.TYPE(123)` → `#N/A`

## Implementation Steps

1. Add `"ERROR.TYPE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnERRORTYPE(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - If arg[0].Type != TypeError, return #N/A
   - Map the error string to its number:
     - `"#NULL!"` → 1, `"#DIV/0!"` → 2, `"#VALUE!"` → 3
     - `"#REF!"` → 4, `"#NAME?"` → 5, `"#NUM!"` → 6, `"#N/A"` → 7
   - If error string doesn't match known errors, return #N/A
3. Add `case "ERROR.TYPE":` dispatch in `formula/eval.go`
   - Note: the function name contains a dot, ensure the parser/compiler handles `ERROR.TYPE` correctly
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `ERROR.TYPE(1/0)` → `2` (#DIV/0!)
- `ERROR.TYPE(#N/A)` → `7` (via NA())
- `ERROR.TYPE(#VALUE!)` → `3`
- `ERROR.TYPE(123)` → #N/A (not an error)
- `ERROR.TYPE("hello")` → #N/A (not an error)
