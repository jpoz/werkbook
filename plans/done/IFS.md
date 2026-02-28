# Plan: IFS Function

## Category
Logic (formula/functions_logic.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/ifs-function-36329a26-37b2-467c-972b-4a39bd951d45

**Syntax:** `IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)`

**Description:** Checks one or more conditions and returns the value corresponding to the first TRUE condition. Replaces nested IF statements.

**Arguments:**
- `logical_test1` (required) — condition evaluating to TRUE or FALSE
- `value_if_true1` (required) — result returned if logical_test1 is TRUE
- Up to 127 condition/value pairs

**Remarks:**
- Args must come in pairs (condition, value); odd arg count is an error
- If no TRUE condition found, returns #N/A
- Non-boolean evaluations return #VALUE!
- Use `TRUE` as final condition for a default/else value

**Examples:**
- `=IFS(1>2, "a", 3>2, "b")` → `"b"`
- `=IFS(FALSE, "a", FALSE, "b")` → #N/A
- `=IFS(TRUE, "always")` → `"always"`

## Implementation Steps

1. Add `"IFS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnIFS(args []Value) (Value, error)` in `formula/functions_logic.go`:
   - Validate even number of args ≥ 2, else #VALUE!
   - Iterate pairs: coerce condition to bool
   - Return value for first TRUE condition
   - If no TRUE found, return #N/A
3. Add `case "IFS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_logic_test.go`

## Test Cases
- `IFS(TRUE, "yes")` → `"yes"`
- `IFS(FALSE, "a", TRUE, "b")` → `"b"`
- `IFS(FALSE, "a", FALSE, "b")` → #N/A
- `IFS(1>2, "a", 3>2, "b", TRUE, "c")` → `"b"`
- `IFS(TRUE)` → #VALUE! (odd arg count)
