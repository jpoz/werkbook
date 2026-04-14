# TODO

## Spill lock-down follow-ups

Captured during the spill testing-strategy work. Tests and fixtures in
place are green; this file lists what still needs a human in the loop.

### Action items, in order

1. **Commit the spill test work.**
   - `werkbook` repo:
     - `Makefile` (new `check-spill` target)
     - `spill_regression_test.go` (matrix + cascade + transitions + full-col defined name)
     - `spill_state_test.go` (new, internal unit tests)
     - `spill_fixtures_test.go` (new, fixture parity + round-trip)
     - `spill_fixtures_gen_test.go` (new, `//go:build fixtures`, generates xlsx)
   - `testdata` repo:
     - `spill_fixtures/11_…` through `spill_fixtures/30_…` (20 new fixtures)
     - `check_config.json` (still needs `14_spill_blocker_then_cleared.xlsx` removed from `ignore_files` after the sibling repo is updated)

2. **Resave the new fixtures through real Excel.** Run
   `../testdata/resave.sh`. That populates cached formula values on the
   20 new fixtures, which is what `wb check` and
   `TestSpillFixturesCachedValueParity` compare against. Without this
   step the fixtures only verify round-trip, not Excel parity.

3. **After resave: re-enable fixture 14 in parity checks.**
   - Remove `14_spill_blocker_then_cleared.xlsx` from `testdata/check_config.json`'s `ignore_files` in the sibling repo.
   - Run `make check-spill` and `go test -run TestSpillFixturesCachedValueParity ./...`
     to confirm the Excel-cached `#SPILL!` now matches werkbook's recalc.

4. **Wire `make check-spill` into CI.** Requires a decision about how
   the testdata sibling repo gets cloned into CI (it's `werkbook/testdata`
   on GitHub and may be private). Options:
   - Add a deploy key + `actions/checkout` step that clones testdata into `../testdata`.
   - Run `make check-spill` as a separate job that `continue-on-error`s when the dir is missing.
   The Go tests in `spill_fixtures_test.go` already handle the missing-dir
   case by calling `t.Skip`, so `go test ./...` stays green either way.

### Bugs discovered while writing the tests

All three are recorded as tests in the repo so they can't silently
disappear. None are in scope for the testing work itself.

1. **Resolved in werkbook.**
   Repro: `spill_fixtures_gen_test.go::buildFixture14`. Set a dynamic
   array anchor, then `SetValue` a blocker inside its spill rect. After
   `Recalculate()`, in-memory `GetValue` correctly returns `#SPILL!`.
   The blocked anchor now round-trips through `SaveAs` and reopen in this
   repo. The sibling `testdata` ignore-list cleanup still needs to land
   before that repo's parity job can include fixture 14 again.

2. **Full-row range references return `#NAME?`.** `SUM(Spill!2:3)`
   does not parse as a valid row-only range. See
   `TestSpillFullRowRangeReference` — it's marked `t.Skip` today but
   will auto-fail (in a good way) once the parser accepts row-only syntax,
   which will then lock down range aggregation over full-row refs.
   - Likely location: `formula/` range parser. Full-column refs
     (`Spill!B:D`) work, so the fix is probably symmetric.

3. **Resolved in werkbook.** Plain range arithmetic such as
   `=SomeSheet!B2:B20*2` now uses top-level spill probing and spills when
   the semantic result is a multi-cell array. Regression coverage lives in
   `spill_regression_test.go` for range arithmetic, reciprocal range
   arithmetic, IF broadcast, scalar reductions, and non-lexical
   dynamic-array metadata serialization.

### When a new spill bug is fixed, add a test

The four test files are the regression registry. Pick the layer that
matches the fix:

- **Pure state/overlay logic** → add a subtest in `spill_state_test.go`.
- **Evaluation or range aggregation** → add a subtest in the matrix in
  `spill_regression_test.go::TestSpillRangeAggregation` or a sibling
  test. If the bug was a new entry in
  `formula/registry.go::inheritedArrayArgFuncs`, add a subtest in
  `TestSpillArrayContextInheritance` citing the PR number.
- **OOXML serialization** → add a fixture via
  `spill_fixtures_gen_test.go`, regenerate, resave, commit.
- **Cross-sheet / defined name** → extend
  `TestSpillRangeAggregationWithDefinedName` or
  `TestSpillFullColumnSiblingFilters`.

### How to regenerate fixtures

```sh
go test -tags=fixtures -run TestGenerateSpillFixtures ./...
../testdata/resave.sh
```

The first command writes uncached xlsx files. The second opens each
one in Excel and saves, which Excel uses to compute and cache formula
results. After both run, `make check-spill` becomes the ground-truth
parity check.
