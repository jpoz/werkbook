# Plan: TYPE Function

## Category
Info (formula/functions_info.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/type-function-45b4e688-4bc3-48b3-a105-ffa892995899

**Syntax:** `TYPE(value)`

**Description:** Returns the type of a value as a number.

**Arguments:**
- `value` (required) — any Excel value

**Return values:**
| Value type | TYPE returns |
|---|---|
| Number | 1 |
| Text | 2 |
| Logical | 4 |
| Error | 16 |
| Array | 64 |

**Remarks:**
- Returns type of the displayed/resulting value, not whether it's a formula
- TYPE("Smith") → 2

**Examples:**
- `=TYPE(1)` → `1`
- `=TYPE("text")` → `2`
- `=TYPE(TRUE)` → `4`
- `=TYPE(NA())` → `16`
- `=TYPE({1,2;3,4})` → `64`

## Implementation Steps

1. Add `"TYPE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnTYPE(args []Value) (Value, error)` in `formula/functions_info.go`:
   - Validate exactly 1 arg
   - Switch on value type:
     - TypeNumber → return 1
     - TypeString → return 2
     - TypeBool → return 4
     - TypeError → return 16
     - ValueArray → return 64
     - TypeEmpty → return 1 (empty treated as number 0)
3. Add `case "TYPE":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_info_test.go`

## Test Cases
- `TYPE(1)` → `1`
- `TYPE("text")` → `2`
- `TYPE(TRUE)` → `4`
- `TYPE(1/0)` → `16` (error)
- `TYPE({1,2,3})` → `64` (array)
