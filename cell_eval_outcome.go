package werkbook

import "github.com/jpoz/werkbook/formula"

// CellEvalOutcome separates the raw evaluator result from the scalar value a
// formula anchor displays in the grid.
type CellEvalOutcome struct {
	Raw         formula.EvalValue
	rawValue    formula.Value
	hasRawValue bool
	Display     formula.Value
	Spill       *SpillPlan
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
		Raw:         formula.ValueToEvalValue(raw),
		rawValue:    raw,
		hasRawValue: true,
		Display:     display,
		Spill:       spill,
	}
}

func (o CellEvalOutcome) RawValue() formula.Value {
	if o.hasRawValue {
		return o.rawValue
	}
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
	return formulaDisplayEvalValueAt(formula.ValueToEvalValue(fv), isArrayFormula, intersectRangeOrigin, currentCol, currentRow)
}

func formulaDisplayEvalValueAt(
	ev formula.EvalValue,
	isArrayFormula bool,
	intersectRangeOrigin bool,
	currentCol, currentRow int,
) formula.Value {
	if !intersectRangeOrigin || isArrayFormula {
		return formulaDisplayEvalValue(ev, isArrayFormula)
	}

	if ev.Kind != formula.EvalRef || ev.Ref == nil {
		return formulaDisplayEvalValue(ev, isArrayFormula)
	}

	ref := ev.Ref
	if ref.FromCol == ref.ToCol && ref.FromRow == ref.ToRow {
		return formulaDisplayRefValueAt(ref, 0, 0, isArrayFormula)
	}
	if currentCol >= ref.FromCol && currentCol <= ref.ToCol &&
		currentRow >= ref.FromRow && currentRow <= ref.ToRow {
		return formulaDisplayRefValueAt(
			ref,
			currentRow-ref.FromRow,
			currentCol-ref.FromCol,
			isArrayFormula,
		)
	}
	if ref.FromCol == ref.ToCol {
		if currentRow >= ref.FromRow && currentRow <= ref.ToRow {
			return formulaDisplayRefValueAt(ref, currentRow-ref.FromRow, 0, isArrayFormula)
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	if ref.FromRow == ref.ToRow {
		if currentCol >= ref.FromCol && currentCol <= ref.ToCol {
			return formulaDisplayRefValueAt(ref, 0, currentCol-ref.FromCol, isArrayFormula)
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	return formula.ErrorVal(formula.ErrValVALUE)
}

func formulaDisplayRefValueAt(ref *formula.RefValue, rowOffset, colOffset int, isArrayFormula bool) formula.Value {
	if ref == nil || rowOffset < 0 || colOffset < 0 {
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	if ref.Materialized != nil {
		if rowOffset >= ref.Materialized.Rows() || colOffset >= ref.Materialized.Cols() {
			return formula.ErrorVal(formula.ErrValVALUE)
		}
		return formulaDisplayEvalValue(ref.Materialized.Cell(rowOffset, colOffset), isArrayFormula)
	}
	if rowOffset == 0 && colOffset == 0 && ref.Legacy != nil {
		return formulaDisplayValue(ref.Legacy.SingleCellValue, isArrayFormula)
	}
	return formula.ErrorVal(formula.ErrValVALUE)
}

func formulaDisplayEvalValue(ev formula.EvalValue, isArrayFormula bool) formula.Value {
	switch ev.Kind {
	case formula.EvalKindError:
		return formula.ErrorVal(ev.Err)
	case formula.EvalScalar:
		return formulaDisplayValue(ev.Scalar, isArrayFormula)
	case formula.EvalArray:
		if ev.Array == nil {
			return formula.NumberVal(0)
		}
		if ev.Array.SpillClass == formula.SpillScalarOnly && !isArrayFormula {
			return formula.ErrorVal(formula.ErrValVALUE)
		}
		if ev.Array.Rows == 0 || ev.Array.Cols == 0 || ev.Array.Grid == nil {
			return formula.NumberVal(0)
		}
		return formulaDisplayEvalValue(ev.Array.Grid.Cell(0, 0), isArrayFormula)
	case formula.EvalRef:
		return formulaDisplayRefValueAt(ev.Ref, 0, 0, isArrayFormula)
	default:
		return formula.NumberVal(0)
	}
}
