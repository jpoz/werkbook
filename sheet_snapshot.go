package werkbook

import (
	"fmt"

	"github.com/jpoz/werkbook/formula"
)

// CellSnapshot is the stored state of one cell: its value, formula, formula
// kind and style. It is the raw state, not the evaluated state: a spill
// follower that has no stored cell gives a snapshot that does not exist.
//
// A snapshot shares its strings and its *Style with the cell it came from.
// The zero value is the snapshot of a cell that does not exist.
//
// Callers that record history take a snapshot before a write and restore it
// to undo the write. The With* methods derive the snapshot of a cell to write,
// so that copy, fill and format operations use the same write path as undo.
type CellSnapshot struct {
	exists  bool
	value   Value
	formula string
	style   *Style

	isArrayFormula    bool
	dynamicArraySpill bool
	// deriveSpill is true for a snapshot made by WithFormula. RestoreCell
	// then derives the dynamic-array flag from the formula, as SetFormula
	// does. A snapshot taken from a cell keeps the flag of that cell: an
	// imported formula without dynamic-array metadata must stay scalar.
	deriveSpill bool
	// spillFormulaRef is the OOXML formula ref of an imported dynamic-array
	// anchor. The writer uses it as a fallback when the engine cannot derive
	// a spill range.
	spillFormulaRef string
}

// Exists reports whether the snapshot is of a stored cell.
func (c CellSnapshot) Exists() bool { return c.exists }

// IsEmpty reports whether the cell has no value and no formula. A cell that
// has only a style is empty.
func (c CellSnapshot) IsEmpty() bool {
	return !c.exists || (c.formula == "" && c.value.Type == TypeEmpty)
}

// Formula returns the formula text without the leading '=', or "" if none.
func (c CellSnapshot) Formula() string { return c.formula }

// Value returns the stored value. For a formula cell this is the result that
// was cached when the snapshot was taken; it can be stale.
func (c CellSnapshot) Value() Value { return c.value }

// Style returns the style, or nil for the default style.
func (c CellSnapshot) Style() *Style { return c.style }

// IsArrayFormula reports whether the cell holds a legacy (CSE) array formula.
func (c CellSnapshot) IsArrayFormula() bool { return c.isArrayFormula }

// Equal reports whether c and o describe the same stored state. The cached
// result of a formula is ignored: it depends on when the snapshot was taken,
// not on what the cell holds.
func (c CellSnapshot) Equal(o CellSnapshot) bool {
	if c.exists != o.exists {
		return false
	}
	if !c.exists {
		return true
	}
	if c.formula != o.formula || c.style != o.style {
		return false
	}
	if c.formula == "" {
		return c.value == o.value
	}
	return c.isArrayFormula == o.isArrayFormula &&
		c.dynamicArraySpill == o.dynamicArraySpill &&
		c.deriveSpill == o.deriveSpill &&
		c.spillFormulaRef == o.spillFormulaRef
}

// WithValue returns a snapshot of a stored cell that holds v and no formula.
// The style is kept.
func (c CellSnapshot) WithValue(v Value) CellSnapshot {
	return CellSnapshot{exists: true, value: v, style: c.style}
}

// WithFormula returns a snapshot of a stored cell that holds the formula f
// (a single leading '=' is removed). The style is kept. An empty f gives an
// empty cell.
func (c CellSnapshot) WithFormula(f string) CellSnapshot {
	if len(f) > 0 && f[0] == '=' {
		f = f[1:]
	}
	return CellSnapshot{exists: true, formula: f, style: c.style, deriveSpill: f != ""}
}

// WithStyle returns a snapshot of a stored cell with the given style. Value
// and formula are kept.
func (c CellSnapshot) WithStyle(st *Style) CellSnapshot {
	c.exists = true
	c.style = st
	return c
}

// SnapshotCell returns the stored state of the cell at (col, row). It does
// not evaluate formulas. Coordinates out of the sheet give the zero snapshot.
func (s *Sheet) SnapshotCell(col, row int) CellSnapshot {
	r, ok := s.rows[row]
	if !ok {
		return CellSnapshot{}
	}
	c, ok := r.cells[col]
	if !ok {
		return CellSnapshot{}
	}
	snap := CellSnapshot{
		exists:            true,
		value:             c.value,
		formula:           c.formula,
		style:             c.style,
		isArrayFormula:    c.isArrayFormula,
		dynamicArraySpill: c.dynamicArraySpill,
	}
	if c.formula == "" {
		// These flags have no meaning without a formula.
		snap.isArrayFormula = false
		snap.dynamicArraySpill = false
	} else if state, ok := s.spillState(col, row); ok {
		snap.spillFormulaRef = state.formulaRef
	}
	return snap
}

// RestoreCell makes the stored state of the cell at (col, row) equal to snap.
// A snapshot that does not exist removes the cell, and removes its row when
// the row has no other cell, no custom height and is not hidden.
//
// It follows the same steps as SetValue and SetFormula: formulas that depend
// on the cell are invalidated, and spill ranges are derived again at the next
// evaluation. On error the cell is unchanged.
func (s *Sheet) RestoreCell(col, row int, snap CellSnapshot) error {
	if _, err := CoordinatesToCellName(col, row); err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
	// A failed restore changes nothing but still gives a new mark; that is
	// the safe direction for the check that uses it.
	s.file.changes++

	if !snap.exists {
		r, ok := s.rows[row]
		if !ok {
			return nil
		}
		c, ok := r.cells[col]
		if !ok {
			return nil
		}
		s.clearSpillState(col, row)
		if c.formula != "" {
			s.file.deps.Unregister(qc)
		}
		delete(r.cells, col)
		if len(r.cells) == 0 && r.height == 0 && !r.hidden {
			delete(s.rows, row)
		}
		s.file.calcGen++
		s.file.invalidateDependents(s.name, col, row)
		return nil
	}

	// Expansion is the only step that can fail; do it before the first change.
	var src string
	if snap.formula != "" {
		var err error
		src, err = s.file.expandFormula(snap.formula, s.name, row)
		if err != nil {
			return err
		}
	}

	c := s.ensureRow(row).ensureCell(col)
	s.clearSpillState(col, row)
	if c.formula != "" {
		s.file.deps.Unregister(qc)
	}
	c.style = snap.style
	c.formula = snap.formula
	c.compiled = nil
	c.rawValue = formula.Value{}
	c.cachedGen = 0
	c.rawCachedGen = 0

	if snap.formula == "" {
		c.value = snap.value
		c.isArrayFormula = false
		c.dynamicArraySpill = false
		c.dirty = false
	} else {
		// The cached result of the snapshot is not restored: the workbook
		// can be different now, so the formula is evaluated again.
		c.value = Value{}
		c.isArrayFormula = snap.isArrayFormula
		c.dynamicArraySpill = snap.dynamicArraySpill
		if snap.deriveSpill {
			c.dynamicArraySpill = formula.IsDynamicArrayFormula(snap.formula)
		}
		c.dirty = true
		if node, err := formula.Parse(src); err == nil {
			if cf, err := formula.Compile(src, node); err == nil {
				c.compiled = cf
				if snap.deriveSpill {
					c.dynamicArraySpill = formulaNeedsSpillAnchor(c.formula, cf)
				}
				s.file.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
			}
		}
		if c.dynamicArraySpill && !c.isArrayFormula {
			s.setSpillFormulaRef(col, row, snap.spillFormulaRef)
		}
	}

	s.file.calcGen++
	s.file.invalidateDependents(s.name, col, row)
	return nil
}
