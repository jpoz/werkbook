# Plan: SWITCH Function

## Category
Logic (formula/functions_logic.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/switch-function-47ab33c0-28ce-4530-8a45-d532ec4aa25e

**Syntax:** `SWITCH(expression, value1, result1, [default or value2, result2], ...)`

**Description:** Evaluates an expression against a list of values and returns the result for the first match. Optional default value if no match.

**Arguments:**
- `expression` (required) — the value being compared
- `value1...value126` — values to compare against
- `result1...result126` — results returned on match
- `default` (optional) — returned if no match; identified by being the final unpaired argument

**Remarks:**
- Up to 126 value/result pairs (254 arg limit)
- If no match and no default, returns #N/A
- Default is the last arg when total args (excluding expression) is odd

**Examples:**
- `=SWITCH(2, 1, "Sun", 2, "Mon", 3, "Tue")` → `"Mon"`
- `=SWITCH(99, 1, "Sun", 2, "Mon")` → #N/A
- `=SWITCH(99, 1, "Sun", "No match")` → `"No match"` (default)

## Implementation Steps

1. Add `"SWITCH"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnSWITCH(args []Value) (Value, error)` in `formula/functions_logic.go`:
   - Validate at least 3 args
   - First arg is the expression to match
   - Remaining args: if even count → value/result pairs; if odd count → last is default
   - Compare expression against each value (use value equality)
   - Return matched result or default or #N/A
3. Add `case "SWITCH":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_logic_test.go`

## Test Cases
- `SWITCH(2, 1, "a", 2, "b", 3, "c")` → `"b"`
- `SWITCH(99, 1, "a", 2, "b")` → #N/A
- `SWITCH(99, 1, "a", "default")` → `"default"`
- `SWITCH("x", "x", 1, "y", 2)` → `1`
- `SWITCH(1, 1, "match")` → `"match"`
