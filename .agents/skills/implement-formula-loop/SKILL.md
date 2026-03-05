---
name: implement-formula-loop
description: Iteratively implement unimplemented Excel formula functions. Finds which formulas are missing, picks one, implements it with thorough tests, and loops. Use this skill when the user mentions "implement formulas", "add formulas", "formula loop", "missing functions", "unimplemented functions", or wants to add new Excel function support to the library.
---

# Implement Formula Loop

You are implementing missing Excel formula functions in the werkbook community Go library. You will iteratively discover unimplemented functions, pick one, implement it with extensive tests, and repeat.

## Loop Protocol

Run this loop until the user stops you or you've implemented a satisfying number of functions:

### 1. Discover unimplemented functions

Run the following to get the list of all registered functions:

```bash
grep 'Register("' /Users/jpoz/Developer/werkbook/community/formula/functions_*.go | sed 's/.*Register("\([^"]*\)".*/\1/' | sort
```

Then check what common Excel functions are NOT in that list. Prioritize functions that are:
- Commonly used in spreadsheets
- Referenced in existing plan documents at `/Users/jpoz/Developer/werkbook/pro/plans/ready/`
- Dependencies of other functions already implemented
- Simple enough to implement correctly in one pass

You can use the `exceldoc` tool to look up Excel function documentation:

```bash
exceldoc <FUNCTION_NAME>
```

### 2. Pick one function

Choose **one** function to implement. State which function you're implementing and why you chose it. Prefer functions that:
- Have a plan document in `../pro/plans/ready/` (these have detailed specs)
- Are commonly used (e.g., RANK, TRIM, PERCENTRANK before obscure engineering functions)
- Fit naturally into an existing category file (math, stat, text, date, logic, lookup, info, finance)

### 3. Delegate implementation to a sub-agent

Use the **Agent tool** (`subagent_type: "general-purpose"`) to implement the function. Pass the sub-agent a prompt that includes:

- The function name and which category file it belongs in
- The Excel documentation for the function (use `exceldoc` output or the plan document)
- Instructions to:

  1. **Research** — Read the relevant `functions_<category>.go` file to understand existing patterns. Read the corresponding test file `functions_<category>_test.go` to understand test patterns. If a plan document exists at `../pro/plans/ready/<FUNCTION>.md`, read it for detailed specs.

  2. **Look up Excel docs** — Run `exceldoc <FUNCTION_NAME>` to get the official Excel documentation for exact syntax, behavior, and edge cases.

  3. **Implement** — Add the function implementation to the appropriate `functions_<category>.go` file in `/Users/jpoz/Developer/werkbook/community/formula/`. Follow these patterns:
     - Register the function in the file's `init()` block using `Register("<NAME>", NoCtx(fn<Name>))` (or `Register("<NAME>", fn<Name>)` if it needs EvalContext)
     - Validate argument count first
     - Handle `ValueArray` inputs with `LiftUnary`/`LiftBinary` or manual array iteration where appropriate
     - Propagate errors from arguments
     - Use `CoerceNum()`, `CoerceString()`, `IsTruthy()` for type coercion
     - Return results using `NumberVal()`, `StringVal()`, `BoolVal()`, `ErrorVal()`

  4. **Write EXTENSIVE tests** — This is critical. Add tests to the corresponding `functions_<category>_test.go` file. Write **at least 15-25 test cases** covering:
     - Basic/happy path usage
     - Multiple valid argument combinations
     - Edge cases (empty cells, zero, negative numbers, very large numbers)
     - Error handling (wrong number of args, wrong types, #DIV/0!, #VALUE!, #NUM!, #N/A)
     - Boundary conditions (empty strings, single element arrays, max/min values)
     - Array/range inputs where applicable
     - String coercion to numbers where applicable
     - Boolean coercion where applicable
     - Mixed types in ranges
     - Comparison with known Excel results

     Use table-driven tests with `t.Run()` sub-tests for clean organization. Example pattern:
     ```go
     func TestFUNCNAME(t *testing.T) {
         tests := []struct {
             name    string
             formula string
             want    Value
         }{
             {"basic", `FUNCNAME(1)`, NumberVal(1)},
             // ... many more cases
         }
         for _, tt := range tests {
             t.Run(tt.name, func(t *testing.T) {
                 cf := evalCompile(t, tt.formula)
                 got, err := Eval(cf, nil, nil)
                 if err != nil {
                     t.Fatalf("Eval: %v", err)
                 }
                 assertValueEqual(t, tt.want, got)
             })
         }
     }
     ```

     If the function needs cell references, set up a `mockResolver` with test data.

  5. **Run tests** — Execute `cd /Users/jpoz/Developer/werkbook/community && gotestsum ./...` to verify everything passes.

  6. **Commit** — If tests pass: `git add -A && git commit -m "feat: implement <FUNCTION_NAME> function"`

  7. **If tests fail** — Diagnose and fix before committing. Do not commit broken code.

Wait for the sub-agent to complete. Review its result summary.

### 4. Validate against Excel

After the implementation sub-agent finishes and tests pass, launch a **second sub-agent** (`subagent_type: "general-purpose"`) to validate the implementation against real Excel. Pass it a prompt that includes:

- The function name that was just implemented
- Instructions to:

  1. **Generate a comprehensive test XLSX file** — Write a Go program at `/Users/jpoz/Developer/werkbook/testdata/cmd/gen-formula-tests/main.go` (or append to it if it already exists) that generates `/Users/jpoz/Developer/werkbook/testdata/formula/<FUNCTION_NAME>.xlsx`. The program should:
     - Use the `github.com/jpoz/werkbook` library to create the workbook
     - Create a "Data" sheet with test input data (numbers, strings, booleans, empty cells, error values, dates — whatever the function accepts)
     - Create a "Results" sheet with formulas testing every aspect of the function:
       - Basic/happy path cases
       - Multiple argument combinations
       - Edge cases: empty cells, zero, negative numbers, very large/small numbers
       - Error cases: wrong types, #DIV/0!, #VALUE!, #NUM!, #N/A, #REF!
       - Boundary conditions: empty strings, single elements, max/min values
       - Range inputs vs direct scalar inputs (booleans in ranges are ignored by many functions but coerced as direct args)
       - String-to-number coercion behavior
       - Boolean coercion behavior
       - Mixed types in ranges
     - **CRITICAL: Do NOT include formulas with wrong argument counts** (e.g., `SYD(1,2,3)` when SYD requires 4 args, or `VARA()` when VARA requires 1+ args). Excel treats these as **syntactically invalid** — they corrupt the XLSX file and trigger a repair dialog, even when wrapped in `IFERROR()`. This is different from runtime errors like `1/0` or `SYD(1,2,0,1)` which are valid syntax that evaluates to an error. Wrong-arg-count edge cases should only be tested in Go unit tests, never in generated XLSX files.
     - Use column A for test labels (descriptions) and column B for formulas
     - Aim for **30-50 formula cells** covering comprehensive test coverage
     - Call `wb.Recalculate()` then `wb.SaveAs(path)` to write the file
     - Make sure the `formula` directory exists: `os.MkdirAll("/Users/jpoz/Developer/werkbook/testdata/formula", 0755)`

     Example pattern (from gen-edge-tests):
     ```go
     package main

     import (
         "fmt"
         "os"
         werkbook "github.com/jpoz/werkbook"
     )

     func main() {
         os.MkdirAll("formula", 0755)
         wb := werkbook.New(werkbook.FirstSheet("Data"))
         data := wb.Sheet("Data")

         // Set up test data
         data.SetValue("A1", 10)
         data.SetValue("A2", 20)
         // ... more data ...

         res, _ := wb.NewSheet("Results")
         res.SetValue("A1", "Test")
         res.SetValue("B1", "Result")
         row := 2

         // Basic usage
         res.SetValue(fmt.Sprintf("A%d", row), "basic case")
         res.SetFormula(fmt.Sprintf("B%d", row), "FUNCNAME(Data!A1)")
         row++
         // ... many more formula cells ...

         wb.Recalculate()
         wb.SaveAs("formula/FUNCNAME.xlsx")
     }
     ```

  2. **Run the generator** — Execute:
     ```bash
     cd /Users/jpoz/Developer/werkbook/testdata && go run ./cmd/gen-formula-tests/main.go
     ```
     Verify the xlsx file was created at `testdata/formula/<FUNCTION_NAME>.xlsx`.

  3. **Open in Excel for ground truth** — Run this AppleScript to have Excel recalculate and save:
     ```bash
     ABS_PATH=$(cd /Users/jpoz/Developer/werkbook/testdata/formula && pwd)/<FUNCTION_NAME>.xlsx
     osascript -e "
     tell application \"Microsoft Excel\"
         activate
         delay 1
         open POSIX file \"$ABS_PATH\"
         delay 2
         calculate
         save active workbook
         close active workbook saving no
     end tell
     "
     ```

  4. **Compare Excel vs werkbook** — Run the excel-audit tool on just this file:
     ```bash
     cd /Users/jpoz/Developer/werkbook/testdata && go run ./cmd/excel-audit -excel formula/<FUNCTION_NAME>.xlsx
     ```
     This will generate `formula/<FUNCTION_NAME>_spec.json` (Excel's values) and `formula/<FUNCTION_NAME>_issues.json` (any discrepancies).

  5. **If there are discrepancies** — Read the issues JSON, diagnose each mismatch, and fix the implementation in `/Users/jpoz/Developer/werkbook/community/formula/functions_<category>.go`. For each fix:
     - Understand what Excel returns vs what werkbook returns
     - Fix the function implementation to match Excel's behavior
     - Update or add unit tests to cover the corrected behavior
     - Re-run `cd /Users/jpoz/Developer/werkbook/community && gotestsum ./...` to verify
     - Re-run the audit: `cd /Users/jpoz/Developer/werkbook/testdata && go run ./cmd/excel-audit formula/<FUNCTION_NAME>.xlsx`
     - Repeat until there are zero issues (or only acceptable tolerance differences for iterative functions)

  6. **Commit fixes** — If any fixes were made: `cd /Users/jpoz/Developer/werkbook/community && git add -A && git commit -m "fix: <FUNCTION_NAME> match Excel behavior"`

  7. **Commit test file** — Add the generated xlsx and spec files: `cd /Users/jpoz/Developer/werkbook/testdata && git add formula/<FUNCTION_NAME>.xlsx formula/<FUNCTION_NAME>_spec.json && git commit -m "test: add Excel validation for <FUNCTION_NAME>"`

Wait for the validation sub-agent to complete.

### 5. Verify and repeat

After the validation sub-agent finishes, verify everything:

```bash
cd /Users/jpoz/Developer/werkbook/community && gotestsum ./...
```

Send a notification:

```bash
tg "<FUNCTION_NAME> implemented and validated against Excel with N tests"
```

If tests pass, go back to step 1 to pick the next function.

## Completion

When stopping (user request or target reached):

1. Run the full test suite: `cd /Users/jpoz/Developer/werkbook/community && gotestsum ./...`
2. Print a summary of all functions implemented (list each commit message)
3. Notify the user using `tg`

## Rules

- **One function per iteration.** Implement and commit one function at a time.
- **Tests are mandatory.** Every function must have at least 15 test cases. More is better. Aim for comprehensive coverage.
- **Always run tests before committing.** Never commit code that doesn't pass `gotestsum ./...`.
- **Commit after each function.** Clean git history with one commit per function.
- **Follow existing patterns.** Match the code style and conventions in the existing formula files.
- **Use `exceldoc`** to look up exact Excel behavior. Match Excel's behavior precisely, including edge cases.
- **Always validate against Excel.** Every newly implemented function must be validated via the Excel XLSX audit step (step 4). Fix any discrepancies before moving on.
- **If a plan document exists** in `../pro/plans/ready/`, use it as the primary specification.
- **If you get stuck** on a function after two attempts, skip it and move to the next one. Note it for the user.
- **Do not modify existing functions** unless the new function requires it (e.g., adding a helper that's shared).
- **No external dependencies.** The community edition uses only the Go standard library.
- **No invalid-syntax formulas in XLSX files.** Never put formulas with wrong argument counts (e.g., `FUNC()` when it requires args, or `FUNC(a,b)` when it requires 3+) into generated XLSX test files. Excel treats these as file corruption, not runtime errors. Test wrong-arg-count cases only in Go unit tests.
