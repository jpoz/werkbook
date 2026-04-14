# Spill Benchmark Slowdown Report

Date: 2026-04-13

## Summary

The spill correctness work is still slower than `HEAD`, but the worst allocation regressions have been mostly recovered by follow-up fixes.

Current slowdown versus `HEAD` is typically 5% to 23% on row-scaled spill benchmarks. The main outlier is `BenchmarkSpillManyAnchors`, especially at 50 and 100 anchors, where the current tree is 44% to 55% slower.

Allocations are now close to baseline in larger cases. The remaining gap looks less like large allocation churn and more like extra bookkeeping, dynamic dependency updates, and spill-anchor scanning.

## Method

Baseline was measured from a temporary clone of `HEAD`.

Current was measured from the working tree after the mitigation patches.

Command used for both:

```bash
env GOCACHE=/tmp/werkbook-bench-... go test -run '^$' -bench 'BenchmarkSpill' -benchmem -benchtime=200ms -count=1 .
```

These are microbenchmarks, so timing has noise. Allocation counts and bytes are more stable than single-run timing.

## Practical Timing Examples

The table below converts the benchmark deltas into rough wall-clock examples. These are extrapolations from the benchmark scenarios, not end-user workload guarantees.

| Scenario | Benchmark case | Extra time per operation | Example impact |
| --- | --- | ---: | --- |
| Small spill point lookup workbook | `PointLookup/rows=100` | 0.036 ms | 10,000 recalculations would take about 0.36 seconds longer. |
| Medium spill point lookup workbook | `PointLookup/rows=1000` | 0.289 ms | 10,000 recalculations would take about 2.89 seconds longer. |
| Large spill point lookup workbook | `PointLookup/rows=5000` | 1.77 ms | 1,000 recalculations would take about 1.77 seconds longer. |
| Medium full-column aggregate over a spill | `RangeAggregate/rows=1000` | 0.226 ms | 10,000 recalculations would take about 2.26 seconds longer. |
| Large full-column aggregate over a spill | `RangeAggregate/rows=5000` | 0.813 ms | 1,000 recalculations would take about 0.81 seconds longer. |
| 100 independent spill anchors | `ManyAnchors/anchors=100` | 0.556 ms | 10,000 recalculations would take about 5.56 seconds longer. |
| Medium edit and lazy recalc loop | `LazyAfterEdit/rows=1000` | 0.081 ms | 10,000 edits followed by reads would take about 0.81 seconds longer. |
| Large edit and lazy recalc loop | `LazyAfterEdit/rows=5000` | 0.465 ms | 10,000 edits followed by reads would take about 4.65 seconds longer. |

In practical terms: for one-off recalculation, the current slowdown is usually sub-millisecond except on the largest benchmark shapes. It becomes noticeable when the same spill-heavy path runs thousands of times, or when many independent spill anchors are involved.

## Timing Results

| Benchmark | HEAD | Current | Slowdown |
| --- | ---: | ---: | ---: |
| `BenchmarkSpillPointLookup/rows=100` | 90,820 ns/op | 127,231 ns/op | 40.1% |
| `BenchmarkSpillPointLookup/rows=1000` | 2,298,357 ns/op | 2,587,480 ns/op | 12.6% |
| `BenchmarkSpillPointLookup/rows=5000` | 37,955,465 ns/op | 39,725,486 ns/op | 4.7% |
| `BenchmarkSpillRangeAggregate/rows=100` | 99,612 ns/op | 118,413 ns/op | 18.9% |
| `BenchmarkSpillRangeAggregate/rows=1000` | 1,038,270 ns/op | 1,264,270 ns/op | 21.8% |
| `BenchmarkSpillRangeAggregate/rows=5000` | 5,853,607 ns/op | 6,666,835 ns/op | 13.9% |
| `BenchmarkSpillManyAnchors/anchors=10` | 88,699 ns/op | 95,141 ns/op | 7.3% |
| `BenchmarkSpillManyAnchors/anchors=50` | 476,623 ns/op | 738,045 ns/op | 54.9% |
| `BenchmarkSpillManyAnchors/anchors=100` | 1,265,722 ns/op | 1,821,496 ns/op | 43.9% |
| `BenchmarkSpillLazyAfterEdit/rows=100` | 47,472 ns/op | 58,608 ns/op | 23.5% |
| `BenchmarkSpillLazyAfterEdit/rows=1000` | 608,152 ns/op | 688,904 ns/op | 13.3% |
| `BenchmarkSpillLazyAfterEdit/rows=5000` | 2,772,542 ns/op | 3,237,977 ns/op | 16.8% |

## Allocation Results

| Benchmark | HEAD B/op | Current B/op | Delta | HEAD allocs/op | Current allocs/op | Delta |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `PointLookup/rows=100` | 109,880 | 119,194 | 8.5% | 830 | 849 | +19 |
| `PointLookup/rows=1000` | 1,115,739 | 1,189,794 | 6.6% | 8,037 | 8,056 | +19 |
| `PointLookup/rows=5000` | 6,247,562 | 6,609,648 | 5.8% | 40,047 | 40,066 | +19 |
| `RangeAggregate/rows=100` | 279,291 | 289,377 | 3.6% | 1,851 | 1,961 | +110 |
| `RangeAggregate/rows=1000` | 2,746,195 | 2,756,290 | 0.4% | 17,162 | 17,272 | +110 |
| `RangeAggregate/rows=5000` | 14,788,704 | 14,798,691 | 0.1% | 85,178 | 85,288 | +110 |
| `ManyAnchors/anchors=10` | 250,905 | 261,257 | 4.1% | 2,107 | 2,171 | +64 |
| `ManyAnchors/anchors=50` | 1,252,635 | 1,303,126 | 4.0% | 10,509 | 10,785 | +276 |
| `ManyAnchors/anchors=100` | 2,505,531 | 2,606,623 | 4.0% | 21,010 | 21,542 | +532 |
| `LazyAfterEdit/rows=100` | 122,558 | 125,278 | 2.2% | 762 | 798 | +36 |
| `LazyAfterEdit/rows=1000` | 1,213,418 | 1,216,136 | 0.2% | 7,069 | 7,105 | +36 |
| `LazyAfterEdit/rows=5000` | 6,714,992 | 6,717,711 | 0.0% | 35,079 | 35,115 | +36 |

## Root Causes Found

1. Spill probing was too broad.

   The initial semantic-detection change treated any formula with a range as a dynamic spill candidate. Scalar reducers such as `SUM(Spill!B:B)` and `COUNT(Spill!B:B)` ran a normal eval and then a spill-probe eval, re-materializing the same range a second time.

   Mitigation: added `CompiledFormula.NeedsSpillProbe` and narrowed `formulaShouldProbeForSpill` to formulas whose top-level result shape can actually change under spill probing.

2. Range materialization added an intermediate snapshot.

   `GetRangeValues` was collecting populated cells into `[]rangeMaterializationCell`, then passing that to `materializeRange`, which iterated it again. This increased allocation and copying around range aggregate benchmarks.

   Mitigation: added a sheet-backed `rangeGridReader` so `materializeRange` can still stay pure for tests while production code iterates sheet cells directly.

3. Spill lookup republished identical dynamic ranges.

   Point lookup repeatedly refreshed the same spill anchor and wrote identical dynamic dependency ranges.

   Mitigation: made dynamic range replacement idempotent and cached spill-anchor refs alongside the overlay state map.

## Remaining Likely Hotspots

1. Spill-anchor scanning.

   `BenchmarkSpillManyAnchors` is still the clearest outlier. The current overlay keeps anchor state in a map and a dense ref slice, but lookups still scan anchors to find whether a cell is inside a published spill.

   Next option: add a per-sheet indexed lookup structure for published spill rectangles, keyed by row or by occupied spill cell for common point lookups.

2. Dynamic dependency bookkeeping.

   Dynamic range dependency tracking is now correct for spill blockers and materialized grown spill ranges, but it still adds bookkeeping to hot eval paths.

   Next option: only track materialized dynamic ranges when the producer is dynamic or volatile range-producing behavior such as `INDIRECT` or `OFFSET`, or when the materialized/grown range is not already covered by the compiled static range subscriptions.

3. Occupancy tracking for materialized ranges.

   `materializeRange` allocates an `occupied` matrix for every materialized range. This is correct and simple, but it is expensive for large sparse ranges.

   Next option: switch to a sparse occupied set for sparse ranges, keeping the dense bool matrix only when the range is small or dense.

## Correctness Status

Validation after the mitigations:

```bash
gotestsum ./...
```

Result: `19373 tests, 1 skipped`.

```bash
make -C ../testdata check-all
```

Result: `20123 matches, 0 mismatches`.
