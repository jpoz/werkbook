# Plan: PERCENTILE Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/percentile-function-91b43a53-543c-4708-93de-d626debdddca

**Syntax:** `PERCENTILE(array, k)`

**Description:** Returns the k-th percentile of values in a range. Equivalent to PERCENTILE.INC.

**Arguments:**
- `array` (required) — the array or range of data that defines relative standing
- `k` (required) — the percentile value in the range 0 to 1, inclusive

**Remarks:**
- If k is non-numeric, returns #VALUE!
- If k < 0 or k > 1, returns #NUM!
- If k is not a multiple of 1/(n-1), PERCENTILE interpolates to determine the value at the k-th percentile

**Algorithm (interpolation):**
1. Sort data ascending
2. Compute rank = k * (n - 1)
3. integer_part = floor(rank), frac = rank - integer_part
4. If frac == 0, result = data[integer_part]
5. Else result = data[integer_part] + frac * (data[integer_part + 1] - data[integer_part])

**Examples:**
- `=PERCENTILE({1, 2, 3, 4}, 0.3)` → `1.9`

## Implementation Steps

1. Add `"PERCENTILE"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnPERCENTILE(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Validate exactly 2 args
   - Flatten arg[0] to get numeric values, sort them ascending
   - Coerce arg[1] to number (k)
   - Return #NUM! if k < 0 or k > 1
   - Return #NUM! if array is empty
   - Compute interpolated percentile value
3. Add `case "PERCENTILE":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `PERCENTILE({1, 2, 3, 4}, 0.3)` → `1.9`
- `PERCENTILE({1, 2, 3, 4}, 0)` → `1` (minimum)
- `PERCENTILE({1, 2, 3, 4}, 1)` → `4` (maximum)
- `PERCENTILE({1, 2, 3, 4}, 0.5)` → `2.5` (median)
- `PERCENTILE({1, 2, 3, 4}, -0.1)` → #NUM!
- `PERCENTILE({1, 2, 3, 4}, 1.1)` → #NUM!
