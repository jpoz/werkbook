# Plan: MINIFS Function

## Category
Statistics (formula/functions_stat.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599

**Syntax:** `MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Description:** Returns the minimum value among cells specified by a given set of conditions.

**Arguments:**
- `min_range` (required) — range where minimum is determined
- `criteria_range1` (required) — cells to evaluate with criteria
- `criteria1` (required) — criteria as number, expression, or text
- Additional range/criteria pairs (optional) — up to 126 pairs

**Remarks:**
- min_range and criteria_range must be same size/shape, else #VALUE!
- Same criteria format as SUMIFS, COUNTIFS, MAXIFS
- Returns 0 when no cells match all criteria
- Empty criteria cells treated as zero

**Examples:**
- `=MINIFS(A2:A7, B2:B7, 1)` → min where B=1

## Implementation Steps

1. Add `"MINIFS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnMINIFS(args []Value) (Value, error)` in `formula/functions_stat.go`:
   - Follow the existing `fnSUMIFS` / `fnMAXIFS` pattern
   - Validate at least 3 args, odd count
   - Parse criteria using existing `matchesCriteria` helper
   - Track minimum value among matching cells
   - Return 0 if no matches
3. Add `case "MINIFS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_stat_test.go`

## Test Cases
- `MINIFS({10,20,30}, {1,2,1}, 1)` → `10`
- `MINIFS({10,20,30}, {1,2,1}, 2)` → `20`
- `MINIFS({10,20,30}, {1,2,1}, 3)` → `0` (no match)
- `MINIFS({5,15,25}, {"a","b","a"}, "a")` → `5`
