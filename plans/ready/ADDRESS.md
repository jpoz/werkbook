# Plan: ADDRESS Function

## Category
Lookup/Reference (formula/functions_lookup.go)

## Excel Documentation
URL: https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89

**Syntax:** `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])`

**Description:** Returns the address of a cell as text, given specified row and column numbers.

**Arguments:**
- `row_num` (required) — row number for the cell reference
- `column_num` (required) — column number for the cell reference
- `abs_num` (optional) — type of reference to return:
  - 1 or omitted: Absolute (`$C$2`)
  - 2: Absolute row, relative column (`C$2`)
  - 3: Relative row, absolute column (`$C2`)
  - 4: Relative (`C2`)
- `a1` (optional) — TRUE or omitted for A1-style, FALSE for R1C1-style
- `sheet_text` (optional) — worksheet name for external reference

**Examples:**
- `=ADDRESS(2,3)` → `"$C$2"`
- `=ADDRESS(2,3,2)` → `"C$2"`
- `=ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")` → `"'[Book1]Sheet1'!R2C3"`

## Implementation Steps

1. Add `"ADDRESS"` to `knownFunctions` in `formula/compiler.go`
2. Implement `fnADDRESS(args []Value) (Value, error)` in `formula/functions_lookup.go`:
   - Validate 2-5 args
   - Coerce arg[0] to number for row_num
   - Coerce arg[1] to number for column_num
   - Default abs_num to 1, a1 to true, sheet_text to ""
   - Use existing `CoordinatesToCellName()` or column number → letter conversion from `coords.go`
   - For A1 style: build string with `$` signs per abs_num
   - For R1C1 style: build `R{row}C{col}` string with brackets for relative refs
   - Prepend sheet name if provided (with quotes if contains spaces)
3. Add `case "ADDRESS":` dispatch in `formula/eval.go`
4. Add tests in `formula/functions_lookup_test.go`

## Test Cases
- `ADDRESS(2,3)` → `"$C$2"`
- `ADDRESS(2,3,2)` → `"C$2"`
- `ADDRESS(2,3,3)` → `"$C2"`
- `ADDRESS(2,3,4)` → `"C2"`
- `ADDRESS(2,3,1,FALSE)` → `"R2C3"`
- `ADDRESS(2,3,4,FALSE)` → `"R[2]C[3]"`
- `ADDRESS(1,1,1,TRUE,"Sheet2")` → `"Sheet2!$A$1"`
- `ADDRESS(1,1,1,TRUE,"My Sheet")` → `"'My Sheet'!$A$1"`
