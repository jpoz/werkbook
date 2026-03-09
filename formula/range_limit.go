package formula

// MaxMaterializedRangeCells caps how many worksheet cells the engine will
// materialize into a dense array for range evaluation. The limit is set to
// preserve a single full-column reference while rejecting wider ranges that
// would otherwise allocate prohibitively large matrices.
const MaxMaterializedRangeCells int64 = maxExcelRows

// RangeCellCountExceedsLimit reports whether a dense rows×cols materialization
// would exceed the engine's allocation budget.
func RangeCellCountExceedsLimit(rows, cols int) bool {
	if rows <= 0 || cols <= 0 {
		return false
	}
	return int64(rows)*int64(cols) > MaxMaterializedRangeCells
}

func isRangeOverflowMatrix(rows [][]Value) bool {
	return len(rows) == 1 &&
		len(rows[0]) == 1 &&
		rows[0][0].Type == ValueError &&
		rows[0][0].Err == ErrValREF &&
		rows[0][0].RangeOverflow
}
