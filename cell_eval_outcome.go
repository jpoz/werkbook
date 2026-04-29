package werkbook

import (
	"fmt"

	"github.com/jpoz/werkbook/formula"
)

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
	fmt.Printf("DEBUG formulaDisplayEvalValueAt: kind=%v, intersectRangeOrigin=%v, isArrayFormula=%v\n",
		ev.Kind, intersectRangeOrigin, isArrayFormula)
	if ev.Kind == formula.EvalArray && ev.Array != nil {
		fmt.Printf("  EvalArray rows=%d cols=%d origin=%v spillClass=%v\n",
			ev.Array.Rows, ev.Array.Cols, ev.Array.Origin, ev.Array.SpillClass)
	}
	if !intersectRangeOrigin || isArrayFormula {
		return formulaDisplayEvalValue(ev, isArrayFormula)
	}

	if ev.Kind == formula.EvalArray && ev.Array != nil &&
		ev.Array.Origin != nil && ev.Array.Origin.Range != nil &&
		ev.Array.Grid != nil {
		return formulaDisplayArrayIntersect(ev.Array, currentCol, currentRow, isArrayFormula)
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

// formulaDisplayArrayIntersect implements legacy implicit intersection for an
// EvalArray that carries a source-range origin (e.g. INDEX(A1:D4,0,2) which
// produces a column window backed by B1:B4). Excel resolves these in
// single-cell context by intersecting the array along the formula cell's row
// or column, mirroring the EvalRef path above.
func formulaDisplayArrayIntersect(arr *formula.ArrayValue, currentCol, currentRow int, isArrayFormula bool) formula.Value {
	rng := arr.Origin.Range
	grid := arr.Grid
	rows := arr.Rows
	cols := arr.Cols
	cellAt := func(r, c int) formula.Value {
		if r < 0 || c < 0 || r >= rows || c >= cols {
			return formula.ErrorVal(formula.ErrValVALUE)
		}
		return formulaDisplayEvalValue(grid.Cell(r, c), isArrayFormula)
	}
	if rng.FromCol == rng.ToCol && rng.FromRow == rng.ToRow {
		return cellAt(0, 0)
	}
	if currentCol >= rng.FromCol && currentCol <= rng.ToCol &&
		currentRow >= rng.FromRow && currentRow <= rng.ToRow {
		return cellAt(currentRow-rng.FromRow, currentCol-rng.FromCol)
	}
	if rng.FromCol == rng.ToCol {
		if currentRow >= rng.FromRow && currentRow <= rng.ToRow {
			return cellAt(currentRow-rng.FromRow, 0)
		}
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	if rng.FromRow == rng.ToRow {
		if currentCol >= rng.FromCol && currentCol <= rng.ToCol {
			return cellAt(0, currentCol-rng.FromCol)
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
