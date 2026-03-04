---
name: excel-audit-loop
description: Iteratively fix discrepancies between our Go library and Excel by running the excel-audit tool in a loop. Use this skill whenever the user mentions "excel audit", "excel errors", "audit loop", "fix excel discrepancies", "excel parity", or wants to bring our library's calculations in line with Excel's output. Also trigger when the user asks to "ralph loop" or "run the audit" on the calculation engine.
---

# Excel Audit Loop

You are fixing discrepancies between our Go library and Microsoft Excel by iteratively running an audit tool and resolving errors one at a time.

## Loop Protocol

Run this loop until there are zero errors remaining:

### 1. Run the audit

```bash
cd ../testdata && go run ./cmd/excel-audit
```

Capture the full output. If the exit code is 0 and there are no errors, you're done — skip to the Completion step.

### 2. Triage

Review the list of errors. Pick **one** error to fix — prefer whichever looks simplest or most isolated so early fixes don't destabilize later ones. State which error you're fixing and why you chose it.

### 3. Diagnose

Read the relevant source code to understand why our library produces a different result than Excel. Identify the root cause before writing any code.

### 4. Fix

Make the minimal code change that resolves the discrepancy. Do not refactor unrelated code.

### 5. Add or update tests

Write or extend unit tests that cover the case you just fixed. Run the tests to make sure they pass:

```bash
go test ./...
```

### 6. Verify

Re-run the audit to confirm your fix resolved the error:

```bash
cd ../testdata && go run ./cmd/excel-audit
```

If the error you targeted is gone, commit the fix with a clear message:

```bash
git add -A && git commit -m "fix: <short description of what was wrong>"
```

If the error persists or new errors appeared because of your change, diagnose and fix before moving on.

### 7. Repeat

Go back to step 1.

## Completion

When the audit reports zero errors:

1. Run the full test suite one final time: `go test ./...`
2. Print a summary of all fixes made (list each commit message).
3. Tell the user you're done.

## Rules

- **One error per iteration.** Do not batch fixes — it makes diagnosis harder if something breaks.
- **Always run tests before committing.** Never commit code that doesn't pass `go test ./...`.
- **Commit after each fix.** This gives a clean rollback point per change.
- **If you get stuck** on a particular error after two attempts, skip it, note it for the user, and move to the next one. Come back to skipped errors after all others are resolved.
- **Do not modify test fixture files** (e.g., `.xlsx` files in testdata) unless that is explicitly the right fix.
