# Plan: SUBTOTAL Function

## Category
Math (formula/functions_math.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/subtotal-function-7b027003-f060-4ade-9040-e478765b9939

**Syntax:** `SUBTOTAL(function_num, ref1, [ref2], ...)`

**Description:** Returns a subtotal in a list or database. Applies a specified aggregate function (AVERAGE, COUNT, MAX, etc.) to the given references.

**Arguments:**
- `function_num` (required) — number 1-11 or 101-111 specifying the aggregate function:
  - 1/101: AVERAGE
  - 2/102: COUNT
  - 3/103: COUNTA
  - 4/104: MAX
  - 5/105: MIN
  - 6/106: PRODUCT
  - 7/107: STDEV
  - 8/108: STDEVP
  - 9/109: SUM
  - 10/110: VAR
  - 11/111: VARP
  (1-11 includes manually hidden rows; 101-111 excludes them. In our implementation, both behave the same since we don't track row visibility.)
- `ref1` (required) — first range/reference for the subtotal
- `ref2, ...` (optional) — additional ranges (up to 254)

**Remarks:**
- Nested SUBTOTAL calls within refs are ignored to avoid double counting (not applicable in our formula engine since we don't track which cells contain SUBTOTAL)
- If function_num is not in 1-11 or 101-111, returns #VALUE!

**Examples:**
- `=SUBTOTAL(9, {120, 10, 150, 23})` → `303` (SUM)
- `=SUBTOTAL(1, {120, 10, 150, 23})` → `75.75` (AVERAGE)

## Implementation Steps

1. Add `"SUBTOTAL"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnSUBTOTAL(args []Value) (Value, error)` in `formula/functions_math.go`:
   - Validate at least 2 args
   - Coerce arg[0] to number, truncate to integer for function_num
   - Normalize function_num: if 101-111, subtract 100 to get 1-11
   - Return #VALUE! if not in range 1-11
   - Pass remaining args (args[1:]) to the appropriate existing function:
     - 1 → fnAVERAGE, 2 → fnCOUNT, 3 → fnCOUNTA, 4 → fnMAX
     - 5 → fnMIN, 6 → fnPRODUCT, 7 → fnSTDEV, 8 → fnSTDEVP
     - 9 → fnSUM, 10 → fnVAR, 11 → fnVARP
   - **Dependency: STDEV, STDEVP, VAR, VARP must be implemented first**
3. Add `case "SUBTOTAL":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_math_test.go`

## Test Cases
- `SUBTOTAL(9, {120, 10, 150, 23})` → `303` (SUM)
- `SUBTOTAL(109, {120, 10, 150, 23})` → `303` (SUM, ignoring hidden)
- `SUBTOTAL(1, {120, 10, 150, 23})` → `75.75` (AVERAGE)
- `SUBTOTAL(4, {120, 10, 150, 23})` → `150` (MAX)
- `SUBTOTAL(5, {120, 10, 150, 23})` → `10` (MIN)
- `SUBTOTAL(2, {120, 10, 150, 23})` → `4` (COUNT)
- `SUBTOTAL(0, {1, 2})` → #VALUE! (invalid function_num)
- `SUBTOTAL(12, {1, 2})` → #VALUE! (invalid function_num)
