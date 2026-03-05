---
name: formula-coverage-loop
description: Iteratively improve test coverage for formula functions that have insufficient tests. Reads FORMULAS.md to find functions with low or zero test counts, then adds comprehensive tests. Use this skill when the user mentions "formula coverage", "test coverage", "increase coverage", "more tests", "undertested formulas", "formula tests", or wants to improve test coverage for existing formula functions.
---

# Formula Coverage Loop

You are improving test coverage for existing formula functions in the werkbook community Go library. You will read FORMULAS.md to find functions with insufficient test coverage, pick one, write comprehensive tests for it, and repeat.

## Loop Protocol

Run this loop until the user stops you or all supported functions have adequate coverage (15+ tests each):

### 1. Find undertested functions

Read `/Users/jpoz/Developer/werkbook/community/FORMULAS.md` and parse the supported formulas table. Identify functions that are undertested, prioritizing:

1. **No tests at all** (`-` in the Tests column) — highest priority
2. **Very few tests** (1-5 tests) — high priority
3. **Below threshold** (6-14 tests) — medium priority

From the candidates, pick **one** function to improve. Prefer functions that are:
- Commonly used in spreadsheets (e.g., ROUND, LEFT, RIGHT before obscure functions)
- More complex (more edge cases to cover)
- In categories that have low overall coverage

### 2. Pick one function

Choose **one** function to add tests for. State which function you're testing and its current test count.

### 3. Delegate test writing to a sub-agent

Use the **Agent tool** (`subagent_type: "general-purpose"`) to write the tests. Pass the sub-agent a prompt that includes:

- The function name and its current test count
- Which category file it belongs in (e.g., `functions_math.go` / `functions_math_test.go`)
- Instructions to:

  1. **Research the function** — Read the implementation in the relevant `functions_<category>.go` file in `/Users/jpoz/Developer/werkbook/community/formula/`. Understand exactly what it does, what arguments it takes, and how it handles edge cases.

  2. **Read existing tests** — Read `functions_<category>_test.go` to see what tests already exist for this function (if any) and understand the test patterns used.

  3. **Look up Excel docs** — Run `exceldoc <FUNCTION_NAME>` to get the official Excel documentation for exact syntax, behavior, and edge cases.

  4. **Write comprehensive tests** — Add tests to the corresponding `functions_<category>_test.go` file. If the function already has some tests, ADD to them — do not delete existing tests. Write enough tests to bring the total to **at least 15-25 test cases** covering:
     - Basic/happy path usage with various inputs
     - Multiple valid argument combinations (optional args, different arg counts)
     - Edge cases (empty cells, zero, negative numbers, very large numbers, very small numbers)
     - Error handling (wrong number of args, wrong types, #DIV/0!, #VALUE!, #NUM!, #N/A, #REF!)
     - Boundary conditions (empty strings, single element arrays, max/min values)
     - Array/range inputs where applicable
     - String-to-number coercion where applicable
     - Boolean coercion where applicable
     - Mixed types in ranges
     - Results that match known Excel behavior

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

  6. **Commit** — If tests pass: `git add -A && git commit -m "test: add comprehensive tests for <FUNCTION_NAME>"`

  7. **If tests fail** — Investigate whether the test expectation is wrong or the implementation has a bug.
     - If the test expectation is wrong, fix the test to match correct Excel behavior.
     - If the implementation has a bug, fix the implementation AND the test. Use commit message: `fix: correct <FUNCTION_NAME> behavior and add tests`

Wait for the sub-agent to complete. Review its result summary.

### 4. Update FORMULAS.md

After tests are added, update the test count in `/Users/jpoz/Developer/werkbook/community/FORMULAS.md` for the function. Count the actual number of test cases by looking at the test file. Commit:

```bash
git add FORMULAS.md && git commit -m "docs: update test count for <FUNCTION_NAME> in FORMULAS.md"
```

### 5. Verify and repeat

After the sub-agent finishes, verify the full test suite still passes:

```bash
cd /Users/jpoz/Developer/werkbook/community && gotestsum ./...
```

Send a notification:

```bash
tg "<FUNCTION_NAME> now has N tests (was M)"
```

If there are more undertested functions, go back to step 1.

## Completion

When stopping (user request or all functions have 15+ tests):

1. Run the full test suite: `cd /Users/jpoz/Developer/werkbook/community && gotestsum ./...`
2. Print a summary of all functions that had coverage improved (function name, old count, new count)
3. Notify the user using `tg`

## Rules

- **One function per iteration.** Test and commit one function at a time.
- **Never delete existing tests.** Only add new test cases.
- **At least 15 test cases per function.** If a function already has 10 tests, add at least 5 more to reach 15+.
- **Always run tests before committing.** Never commit code that doesn't pass `gotestsum ./...`.
- **Commit after each function.** Clean git history with one commit per function.
- **Follow existing test patterns.** Match the test style and conventions in the existing test files.
- **Use `exceldoc`** to verify expected Excel behavior. Tests should assert what Excel actually does.
- **If a test reveals an implementation bug**, fix the bug in the same iteration. Use a descriptive commit message that mentions both the fix and the tests.
- **If you get stuck** on a function after two attempts, skip it and move to the next one. Note it for the user.
- **Prioritize breadth.** It's better to bring 10 functions from 0 to 15 tests than to bring 1 function from 10 to 100 tests.
- **No external dependencies.** The community edition uses only the Go standard library.
