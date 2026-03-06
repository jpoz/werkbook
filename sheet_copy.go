package werkbook

import (
	"fmt"

	"github.com/jpoz/werkbook/formula"
)

// CopySheet duplicates the named sheet within the workbook under a new name.
func (f *File) CopySheet(srcName, dstName string) (*Sheet, error) {
	src := f.Sheet(srcName)
	if src == nil {
		return nil, ErrSheetNotFound
	}
	return f.CloneSheetFrom(src, dstName)
}

// CloneSheetFrom copies a sheet from this workbook or another workbook into f.
func (f *File) CloneSheetFrom(src *Sheet, dstName string) (*Sheet, error) {
	if src == nil {
		return nil, fmt.Errorf("source sheet is nil")
	}
	if f.SheetIndex(dstName) >= 0 {
		return nil, fmt.Errorf("sheet %q already exists", dstName)
	}

	dst := f.addSheet(dstName)
	dst.visible = src.visible

	for col, width := range src.colWidths {
		dst.colWidths[col] = width
	}

	if len(src.merges) > 0 {
		dst.merges = make([]MergeRange, len(src.merges))
		copy(dst.merges, src.merges)
	}

	for rowNum, srcRow := range src.rows {
		dstRow := &Row{
			num:    rowNum,
			cells:  make(map[int]*Cell, len(srcRow.cells)),
			height: srcRow.height,
			hidden: srcRow.hidden,
		}
		for col, srcCell := range srcRow.cells {
			dstRow.cells[col] = cloneCell(srcCell)
		}
		dst.rows[rowNum] = dstRow
	}

	f.registerSheetFormulas(dst)
	return dst, nil
}

func cloneCell(src *Cell) *Cell {
	clone := &Cell{
		col:            src.col,
		value:          src.value,
		formula:        src.formula,
		isArrayFormula: src.isArrayFormula,
		formulaRef:     src.formulaRef,
		style:          cloneStyle(src.style),
	}
	if clone.formula != "" {
		clone.dirty = true
	}
	return clone
}

func (f *File) registerSheetFormulas(s *Sheet) {
	for _, r := range s.rows {
		for col, c := range r.cells {
			if c.formula == "" {
				continue
			}
			src := f.expandFormula(c.formula, s.name, r.num)
			node, err := formula.Parse(src)
			if err != nil {
				continue
			}
			cf, err := formula.Compile(src, node)
			if err != nil {
				continue
			}
			c.compiled = cf
			qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: r.num}
			f.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
		}
	}
}
