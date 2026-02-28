# Plan: COMBIN Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/combin-function-12a3f276-0a21-423a-8de6-06990aaf638a

**Syntax:** `COMBIN(number, number_chosen)`

**Description:** Returns the number of combinations for a given number of items. C(n,k) = n! / (k!(n-k)!)

**Arguments:**
- `number` (required) — total count of items; truncated to integer
- `number_chosen` (required) — items per combination; truncated to integer

**Remarks:**
- Non-numeric args return #VALUE!
- Returns #NUM! if number < 0, number_chosen < 0, or number < number_chosen
- Combinations ignore internal ordering (unlike permutations)

**Examples:**
- `=COMBIN(8, 2)` → `28`
- `=COMBIN(5, 0)` → `1`
- `=COMBIN(5, 5)` → `1`

## Implementation Steps

1. Add `"COMBIN"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnCOMBIN(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 2 args
   - Coerce to numbers, truncate to integers
   - Return #NUM! if n < 0 or k < 0 or n < k
   - Compute using multiplicative formula to avoid overflow: `C(n,k) = ∏(i=1..k) (n-k+i)/i`
3. Add `case "COMBIN":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `COMBIN(8, 2)` → `28`
- `COMBIN(5, 0)` → `1`
- `COMBIN(5, 5)` → `1`
- `COMBIN(10, 3)` → `120`
- `COMBIN(3, 5)` → #NUM!
- `COMBIN(-1, 2)` → #NUM!
