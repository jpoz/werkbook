# Plan: RANK Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/rank-function-6a2fc49d-1831-4a03-9d8c-c279cf99f723

**Syntax:** `RANK(number, ref, [order])`

**Description:** Returns the rank of a number in a list of numbers. The rank is its size relative to other values in the list. Equivalent to RANK.EQ.

**Arguments:**
- `number` (required) — the number whose rank you want to find
- `ref` (required) — a reference to a list of numbers; non-numeric values are ignored
- `order` (optional) — how to rank: 0 or omitted = descending (largest is rank 1); any nonzero value = ascending (smallest is rank 1)

**Remarks:**
- Duplicate numbers receive the same rank
- Presence of duplicates affects subsequent ranks (e.g., two values tied at rank 5 means rank 6 is skipped; next is rank 7)
- If number is not found in ref, returns #N/A

**Examples:**
- `=RANK(3.5, {7, 3.5, 3.5, 1, 2})` → `3` (descending)
- `=RANK(7, {7, 3.5, 3.5, 1, 2}, 1)` → `5` (ascending)

## Implementation Steps

1. Add `"RANK"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnRANK(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate 2-3 args
   - Coerce arg[0] to number (the value to rank)
   - Flatten arg[1] to get the list of numeric values
   - Determine order from arg[2] (default 0 = descending)
   - For descending: rank = 1 + count of values strictly greater than number
   - For ascending: rank = 1 + count of values strictly less than number
   - If number not found in the list, return #N/A
3. Add `case "RANK":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `RANK(3.5, {7, 3.5, 3.5, 1, 2})` → `3` (descending, default)
- `RANK(7, {7, 3.5, 3.5, 1, 2})` → `1` (descending, largest)
- `RANK(1, {7, 3.5, 3.5, 1, 2})` → `5` (descending, smallest)
- `RANK(1, {7, 3.5, 3.5, 1, 2}, 1)` → `1` (ascending, smallest)
- `RANK(7, {7, 3.5, 3.5, 1, 2}, 1)` → `5` (ascending, largest)
- `RANK(99, {1, 2, 3})` → #N/A (not in list)
