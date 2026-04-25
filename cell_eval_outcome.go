package werkbook

import "github.com/jpoz/werkbook/formula"

// CellEvalOutcome separates the raw evaluator result from the scalar value a
// formula anchor displays in the grid.
type CellEvalOutcome struct {
	Raw     formula.EvalValue
	Display formula.Value
	Spill   *SpillPlan
}

// SpillPlan captures the attempted and published spill bounds for an anchor.
type SpillPlan struct {
	Raw                            formula.Value
	AttemptedToCol, AttemptedToRow int
	PublishedToCol, PublishedToRow int
	Blocked                        bool
}

func newCellEvalOutcome(raw formula.Value, display formula.Value, spill *SpillPlan) CellEvalOutcome {
	return CellEvalOutcome{
		Raw:     formula.ValueToEvalValue(raw),
		Display: display,
		Spill:   spill,
	}
}

func (o CellEvalOutcome) RawValue() formula.Value {
	return formula.EvalValueToValue(o.Raw)
}

func newSpillPlan(raw formula.Value, anchorCol, anchorRow int) *SpillPlan {
	toCol, toRow, ok := spillArrayRect(raw, anchorCol, anchorRow)
	if !ok {
		return nil
	}
	return &SpillPlan{
		Raw:            raw,
		AttemptedToCol: toCol,
		AttemptedToRow: toRow,
		PublishedToCol: toCol,
		PublishedToRow: toRow,
	}
}

func formulaDisplayValue(fv formula.Value, isArrayFormula bool) formula.Value {
	switch fv.Type {
	case formula.ValueArray:
		if fv.NoSpill && !isArrayFormula {
			return formula.ErrorVal(formula.ErrValVALUE)
		}
		if len(fv.Array) > 0 && len(fv.Array[0]) > 0 {
			return formulaDisplayValue(fv.Array[0][0], isArrayFormula)
		}
		return formula.NumberVal(0)
	case formula.ValueEmpty:
		return formula.NumberVal(0)
	default:
		return fv
	}
}

func formulaDisplayValueAt(
	fv formula.Value,
	isArrayFormula bool,
	intersectRangeOrigin bool,
	currentCol, currentRow int,
) formula.Value {
	if !intersectRangeOrigin || isArrayFormula || fv.Type != formula.ValueArray || fv.RangeOrigin == nil {
		return formulaDisplayValue(fv, isArrayFormula)
	}
	ro := fv.RangeOrigin
	if ro.FromCol == ro.ToCol && ro.FromRow == ro.ToRow {
		if len(fv.Array) > 0 && len(fv.Array[0]) > 0 {
			return formulaDisplayValue(fv.Array[0][0], isArrayFormula)
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	if currentCol >= ro.FromCol && currentCol <= ro.ToCol &&
		currentRow >= ro.FromRow && currentRow <= ro.ToRow {
		rowIdx := currentRow - ro.FromRow
		colIdx := currentCol - ro.FromCol
		if rowIdx >= 0 && rowIdx < len(fv.Array) &&
			colIdx >= 0 && colIdx < len(fv.Array[rowIdx]) {
			return formulaDisplayValue(fv.Array[rowIdx][colIdx], isArrayFormula)
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	if ro.FromCol == ro.ToCol {
		if currentRow >= ro.FromRow && currentRow <= ro.ToRow {
			rowIdx := currentRow - ro.FromRow
			if rowIdx >= 0 && rowIdx < len(fv.Array) && len(fv.Array[rowIdx]) > 0 {
				return formulaDisplayValue(fv.Array[rowIdx][0], isArrayFormula)
			}
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	if ro.FromRow == ro.ToRow {
		if currentCol >= ro.FromCol && currentCol <= ro.ToCol && len(fv.Array) > 0 {
			colIdx := currentCol - ro.FromCol
			if colIdx >= 0 && colIdx < len(fv.Array[0]) {
				return formulaDisplayValue(fv.Array[0][colIdx], isArrayFormula)
			}
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	return formula.ErrorVal(formula.ErrValVALUE)
}
