# Stress Test Generator — Lessons Learned

## Architecture

The stress test system has three components:

1. **Generator** (`cmd/gentestdata/`) — Creates XLSX + spec.json files from Go-defined test cases
2. **Bootstrap** (`-bootstrap` flag) — Opens each XLSX in local Excel via AppleScript, saves (triggering recalc), reads back cached values as ground truth `.expected.json`
3. **Regression test** (`stress_test.go`) — Compares werkbook formula evaluation against committed ground truth. No Excel needed.

## Formula Offset System

Test cases are stacked vertically on a single sheet. Each test case defines cells relative to row 1, and `buildSpec()` applies a row offset so they don't collide.

**Critical**: Both the cell positions AND the formula contents must be offset. A formula `SUM(A1:A5)` in test case 3 (at row offset 12) must become `SUM(A13:A17)`, not stay as `SUM(A1:A5)`.

The `offsetFormula()` function uses a regex to find and shift cell references inside formula strings. Two important edge cases:

- **Function names containing digits** (LOG10, DAYS360): The regex must not treat these as cell references. We check if the match is preceded by a letter or followed by `(`.
- **Cross-sheet references** (Data!A1): Only offset references on the current sheet. References to other sheets are left unchanged.

## AppleScript + Excel Integration

### The `open` command blocks on dialogs

When Excel encounters a file it considers problematic, it shows a recovery dialog ("We found a problem with some content..."). The AppleScript `open xlsxFile` command blocks until this dialog is dismissed — but since AppleScript is single-threaded, you can't dismiss it from the same script.

**Solution**: Use `open -a "Microsoft Excel" <path>` via Go's `exec.Command` (non-blocking), then handle dialogs via a separate AppleScript using System Events.

### Two types of repair dialogs

1. **Recovery prompt**: "We found a problem with some content. Do you want us to try to recover?" — Buttons: **Yes** / **No**
2. **Repair result**: "Excel was able to open the file by repairing or removing the unreadable content." — Buttons: **View** / **Delete**

Clicking **Delete** on the second dialog removes the workbook entirely. Click **View** to keep it open.

### `save active workbook` fails after repair

When Excel repairs a file, it may rename the workbook (appending " - Repaired"). The `save active workbook` command then fails with "Parameter error (-50)" because the original file path is stale.

**Solution**: Use `save active workbook in <path>` to save to a separate known path, then read from that path.

### Inspecting Excel state from the command line

```bash
# See all windows (unnamed = dialog)
osascript -e 'tell application "System Events" to tell process "Microsoft Excel" to get name of every window'

# See dialog buttons
osascript -e 'tell application "System Events" to tell process "Microsoft Excel" to get name of every button of (item 1 of (every window whose name is ""))'

# See full dialog text
osascript -e 'tell application "System Events" to tell process "Microsoft Excel" to get entire contents of (item 1 of (every window whose name is ""))'

# Check workbook count and paths
osascript -e 'tell application "Microsoft Excel" to get {count of workbooks, name of every workbook}'
```

### Cascading failures between files

When processing multiple files sequentially, a timeout or error on one file can leave Excel in a bad state (dialog open, workbook still loaded). Subsequent files then fail with "Parameter error".

**Solution**: Run a cleanup script before each file that dismisses any open dialogs and closes all workbooks.

## XLSX Validity

Some werkbook-generated XLSX files trigger Excel's repair mechanism. This doesn't necessarily mean the formulas are wrong — it may be a structural issue with the XML packaging. Files that trigger repair still produce correct formula results after Excel recovers them.

## Bugs Found

The stress test system successfully identified a real werkbook formula engine bug:

- `TEXT(15000, "000000")` should produce `"015000"` (with leading zero padding) but werkbook produces a different result, causing `MID(TEXT(A1*1000,"000000"),2,3)` to return `"500"` instead of the correct `"150"`.

This validates the approach — deterministic ground truth from Excel catches formula evaluation differences that unit tests might miss.
