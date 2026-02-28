# Plan: MAXIFS Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883

**Syntax:** `MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Description:** Returns the maximum value among cells specified by a given set of conditions.

**Arguments:**
- `max_range` (required) — range where maximum is determined
- `criteria_range1` (required) — cells to evaluate with criteria
- `criteria1` (required) — criteria as number, expression, or text
- Additional range/criteria pairs (optional) — up to 126 pairs

**Remarks:**
- max_range and criteria_range must be same size/shape, else #VALUE!
- Same criteria format as SUMIFS and COUNTIFS (already implemented)
- Returns 0 when no cells match all criteria
- Empty criteria cells treated as zero

**Examples:**
- `=MAXIFS(A2:A7, B2:B7, 1)` → max where B=1
- `=MAXIFS(A2:A7, B2:B7, "b", D2:D7, ">100")` → max where B="b" AND D>100

## Implementation Steps

1. Add `"MAXIFS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnMAXIFS(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Follow the existing `fnSUMIFS` pattern closely
   - Validate at least 3 args, odd count (range + pairs)
   - Parse criteria using existing `matchesCriteria` helper
   - Track maximum value (instead of sum) among matching cells
   - Return 0 if no matches
3. Add `case "MAXIFS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `MAXIFS({10,20,30}, {1,2,1}, 1)` → `30`
- `MAXIFS({10,20,30}, {1,2,1}, 2)` → `20`
- `MAXIFS({10,20,30}, {1,2,1}, 3)` → `0` (no match)
- `MAXIFS({5,15,25}, {"a","b","a"}, "a", {1,2,3}, ">1")` → `25`
