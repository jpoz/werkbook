# Plan: PERMUT Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/permut-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3

**Syntax:** `PERMUT(number, number_chosen)`

**Description:** Returns the number of permutations for a given number of objects that can be selected from number objects. P(n,k) = n! / (n-k)!

**Arguments:**
- `number` (required) — total number of objects; truncated to integer
- `number_chosen` (required) — number of objects in each permutation; truncated to integer

**Remarks:**
- Both arguments are truncated to integers
- If number or number_chosen is non-numeric, returns #VALUE!
- If number <= 0 or number_chosen < 0, returns #NUM!
- If number < number_chosen, returns #NUM!

**Examples:**
- `=PERMUT(100, 3)` → `970200`
- `=PERMUT(3, 2)` → `6`

## Implementation Steps

1. Add `"PERMUT"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnPERMUT(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate exactly 2 args
   - Coerce to numbers, truncate to integers
   - Return #NUM! if n <= 0 or k < 0 or n < k
   - Compute using multiplicative formula to avoid overflow: `P(n,k) = ∏(i=0..k-1) (n - i)`
3. Add `case "PERMUT":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `PERMUT(100, 3)` → `970200`
- `PERMUT(3, 2)` → `6`
- `PERMUT(5, 0)` → `1`
- `PERMUT(5, 5)` → `120`
- `PERMUT(3, 5)` → #NUM!
- `PERMUT(-1, 2)` → #NUM!
- `PERMUT(0, 0)` → #NUM! (number <= 0)
