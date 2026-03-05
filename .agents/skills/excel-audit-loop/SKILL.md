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

### 3. Delegate the fix to a sub-agent

Use the **Agent tool** (`subagent_type: "general-purpose"`) to fix the error. Pass the sub-agent a prompt that includes:

- The specific error output from the audit (copy/paste the relevant lines)
- Which error to fix and why you chose it
- Instructions to:
  1. **Diagnose** — read the relevant source code to understand the root cause
  2. **Fix** — make the minimal code change that resolves the discrepancy (no unrelated refactoring)
  3. **Add or update tests** — write or extend unit tests covering the fix, then run `go test ./...` to confirm they pass
  4. **Verify** — re-run `cd ../testdata && go run ./cmd/excel-audit` to confirm the targeted error is resolved
  5. **Commit** — if the fix works: `git add -A && git commit -m "fix: <short description>"`
  6. If the error persists or new errors appeared, diagnose and fix before committing

Wait for the sub-agent to complete. Review its result summary.

### 4. Verify and repeat

After the sub-agent finishes, re-run the audit yourself to confirm progress:

```bash
cd ../testdata && go run ./cmd/excel-audit
```

After the verification send a message using the `tg` command. Example:

```bash
tg "The SUM function was fixed successfully"
```

If errors remain, go back to step 2.

## Completion

When the audit reports zero errors:

1. Run the full test suite one final time: `go test ./...`
2. Print a summary of all fixes made (list each commit message).
3. Tell the user you're done using the `tg` command.

## Rules

- **One error per iteration.** Do not batch fixes — it makes diagnosis harder if something breaks.
- **Always run tests before committing.** Never commit code that doesn't pass `go test ./...`.
- **Commit after each fix.** This gives a clean rollback point per change.
- **If you get stuck** on a particular error after two attempts, skip it, note it for the user, and move to the next one. Come back to skipped errors after all others are resolved.
- **Do not modify test fixture files** (e.g., `.xlsx` files in testdata) unless that is explicitly the right fix.
