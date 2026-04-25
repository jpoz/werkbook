# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Build & Test Commands

```bash
# Run all unit tests
gotestsum ./...

# Run a single test
gotestsum -- -run TestRoundTrip ./...
```

No external dependencies beyond the Go standard library. Module: `github.com/jpoz/werkbook`.

## Useful Tools

```bash
# Look up spreadsheet function documentation (e.g. syntax, behavior, edge cases)
exceldoc SUM
```

## Engineering Approach

- Prefer stewardship over speed. Make changes that improve the long-term shape of the library, not just the local diff.
- Prefer small, surgical patches. Keep slices narrow, reviewable, and easy to validate.
- Preserve existing public behavior unless the task explicitly calls for a semantic change. For formula behavior, Excel parity and workbook corpus parity are the source of truth.
- When touching compatibility scaffolding, move toward a clearer architecture with fewer semantic centers. Do not silently add new bridging layers when the real fix is to tighten the core model.
- Fix root causes when they are reasonably bounded. Avoid papering over architectural problems with one-off conditionals unless the file already establishes that pattern and the broader cleanup is out of scope.
- Treat edge cases as first-class work, not cleanup. When changing formula, spill, range, coercion, or serialization behavior, actively look for nearby cases involving empty values, errors, full-row/full-column refs, trimmed ranges, single-cell refs, implicit intersection, and blocked spills.
- Add targeted regression coverage with each semantic fix. Prefer tests that encode the user-visible behavior and the tricky neighboring cases that are easy to break next.
- Leave clear notes in code only when they help the next maintainer understand an invariant, compatibility constraint, or architectural direction.

## Change Strategy

- Start by identifying the architectural seam being modified and the narrowest viable slice.
- Prefer extending existing helpers and shared paths over duplicating logic in individual functions.
- Keep boundary concerns separate from runtime semantics when working across the public API, formula engine, and OOXML layers.
- Avoid broad refactors unless they are explicitly requested or already called for by the active plan.
- If the worktree is already dirty, do not overwrite adjacent work. Adapt to it and keep your change isolated.
- Use calm, precise language in plans, comments, and delegated work. Favor words like `implement`, `tighten`, `validate`, `advance`, or `finish` over aggressive framing.

## Verification Expectations

- Run the smallest relevant targeted tests first, then widen scope as confidence grows.
- For changes in the formula engine or spill behavior, add or update focused tests near the affected behavior before relying only on full-suite coverage.
- If a plan document defines an exit gate or corpus command, treat that as part of done, not as optional follow-up.

## Architecture

Werkbook is a Go library for reading and writing XLSX spreadsheet files with a built-in formula engine supporting ~170 functions.

### Two-layer design

- **Public API layer** (root package `werkbook`): `File`, `Sheet`, `Row`, `Cell`, `Value` types. All coordinates are 1-based. Cells are stored sparsely in `map[int]*Row` -> `map[int]*Cell`.
- **OOXML layer** (`ooxml/` package): Handles ZIP/XML serialization. Communicates with the public layer through the `WorkbookData`/`SheetData`/`RowData`/`CellData` intermediary types defined in `ooxml/workbook.go`. The public API never exposes OOXML internals.

### Data flow

- **Write path**: `File.buildWorkbookData()` converts `Sheet` -> `SheetData`, then `ooxml.WriteWorkbook()` serializes to ZIP. Strings use a shared string table (SST); cells with `Type: "s"` have their raw string replaced with an SST index during write.
- **Read path**: `ooxml.ReadWorkbook()` parses ZIP -> `WorkbookData`, then `fileFromData()` converts back to `File`/`Sheet`/`Cell` structures. SST indices are resolved to actual strings during read.

### Value system

`Value` is a tagged union (`value.go`) with types: `TypeEmpty`, `TypeNumber`, `TypeString`, `TypeBool`, `TypeError`. The `toValue()` function converts Go types (all int/uint/float variants, string, bool, `time.Time`, nil) to `Value`. Dates are stored as serial numbers via `timeToSerial()` in `date.go`, which accounts for the 1900 leap year bug.

### Coordinate system

`coords.go` handles A1-style references. `CellNameToCoordinates("B3")` returns `(col=2, row=3)`. Max: 1,048,576 rows, 16,384 columns (XFD).
