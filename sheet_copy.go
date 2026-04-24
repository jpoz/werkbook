package werkbook

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
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
	if len(src.spill.anchors) > 0 {
		dst.spill.anchors = make(map[cellKey]*spillAnchorState, len(src.spill.anchors))
		for key, state := range src.spill.anchors {
			if state == nil {
				continue
			}
			copiedKey := cellKey{sheet: dst.name, col: key.col, row: key.row}
			copiedState := *state
			dst.spill.anchors[copiedKey] = &copiedState
		}
		dst.invalidateSpillOverlay()
	}

	// Copy passthrough sheet metadata (drawings, conditional formatting, etc.).
	dst.rootAttrs = cloneRawAttrs(src.rootAttrs)
	dst.extraElements = cloneRawElements(src.extraElements)
	if src.file == f || src.file == nil {
		dst.extraRels = cloneOpaqueRels(src.extraRels)
	} else {
		dst.extraRels = f.importOpaqueEntries(src.file, src.extraRels)
	}

	_ = f.registerSheetFormulas(dst, false)
	return dst, nil
}

func cloneCell(src *Cell) *Cell {
	clone := &Cell{
		col:               src.col,
		value:             src.value,
		formula:           src.formula,
		isArrayFormula:    src.isArrayFormula,
		dynamicArraySpill: src.dynamicArraySpill,
		style:             cloneStyle(src.style),
	}
	if clone.formula != "" {
		clone.dirty = true
	}
	return clone
}

func (f *File) registerSheetFormulas(s *Sheet, strict bool) error {
	for _, r := range s.rows {
		for col, c := range r.cells {
			if c.formula == "" {
				continue
			}
			src, err := f.expandFormula(c.formula, s.name, r.num)
			if err != nil {
				if strict {
					return formulaExpansionError(s.name, col, r.num, err)
				}
				continue
			}
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
			if c.dynamicArraySpill {
				if state, ok := s.spillState(col, r.num); ok {
					s.setSpillBlockerDynamicRanges(col, r.num, state)
				}
			}
		}
	}
	return nil
}

func formulaExpansionError(sheet string, col, row int, err error) error {
	cell := fmt.Sprintf("R%dC%d", row, col)
	if ref, refErr := CoordinatesToCellName(col, row); refErr == nil {
		cell = ref
	}
	return fmt.Errorf("sheet %q cell %s: %w", sheet, cell, err)
}

// resolveRelTarget resolves a relative rel target from the standard sheet
// directory (xl/worksheets) to an absolute zip path.
func resolveRelTarget(target string) string {
	base := "xl/worksheets"
	for strings.HasPrefix(target, "../") {
		target = target[3:]
		if idx := strings.LastIndex(base, "/"); idx >= 0 {
			base = base[:idx]
		} else {
			base = ""
		}
	}
	if base == "" {
		return target
	}
	return base + "/" + target
}

// absoluteToRelTarget converts an absolute zip path back to a relative target
// from the xl/worksheets directory.
func absoluteToRelTarget(absPath string) string {
	if strings.HasPrefix(absPath, "xl/") {
		return "../" + absPath[len("xl/"):]
	}
	return "../../" + absPath
}

// uniqueOpaqueEntryPath returns path if no entry in entries has that path;
// otherwise it appends _2, _3, etc. before the extension.
func uniqueOpaqueEntryPath(entries []ooxml.OpaqueEntry, path string) string {
	taken := make(map[string]struct{}, len(entries))
	for _, e := range entries {
		taken[e.Path] = struct{}{}
	}
	if _, ok := taken[path]; !ok {
		return path
	}
	ext := ""
	base := path
	if dot := strings.LastIndex(path, "."); dot >= 0 {
		ext = path[dot:]
		base = path[:dot]
	}
	for n := 2; ; n++ {
		candidate := fmt.Sprintf("%s_%d%s", base, n, ext)
		if _, ok := taken[candidate]; !ok {
			return candidate
		}
	}
}

// importOpaqueEntries copies opaque entries referenced by srcRels from srcFile
// into f, returning updated rels for the destination sheet. Internal targets
// are resolved, matched to OpaqueEntries in srcFile, and imported (with path
// deconfliction if needed). External rels are copied as-is.
func (f *File) importOpaqueEntries(srcFile *File, srcRels []ooxml.OpaqueRel) []ooxml.OpaqueRel {
	if len(srcRels) == 0 {
		return nil
	}

	srcByPath := make(map[string]*ooxml.OpaqueEntry, len(srcFile.opaqueEntries))
	for i := range srcFile.opaqueEntries {
		srcByPath[srcFile.opaqueEntries[i].Path] = &srcFile.opaqueEntries[i]
	}

	dstByPath := make(map[string][]byte, len(f.opaqueEntries))
	for _, e := range f.opaqueEntries {
		dstByPath[e.Path] = e.Data
	}

	dstRels := make([]ooxml.OpaqueRel, len(srcRels))
	for i, rel := range srcRels {
		dstRels[i] = rel
		if rel.TargetMode == "External" {
			continue
		}
		absPath := resolveRelTarget(rel.Target)
		srcEntry, ok := srcByPath[absPath]
		if !ok {
			continue
		}
		if existing, found := dstByPath[absPath]; found {
			if bytes.Equal(existing, srcEntry.Data) {
				continue
			}
			newPath := uniqueOpaqueEntryPath(f.opaqueEntries, absPath)
			f.opaqueEntries = append(f.opaqueEntries, ooxml.OpaqueEntry{
				Path:        newPath,
				ContentType: srcEntry.ContentType,
				Data:        append([]byte(nil), srcEntry.Data...),
			})
			dstByPath[newPath] = f.opaqueEntries[len(f.opaqueEntries)-1].Data
			dstRels[i].Target = absoluteToRelTarget(newPath)
		} else {
			f.opaqueEntries = append(f.opaqueEntries, ooxml.OpaqueEntry{
				Path:        absPath,
				ContentType: srcEntry.ContentType,
				Data:        append([]byte(nil), srcEntry.Data...),
			})
			dstByPath[absPath] = f.opaqueEntries[len(f.opaqueEntries)-1].Data
		}
	}

	// Import opaque content-type defaults (e.g., .png, .vml extensions).
	if len(srcFile.opaqueDefaults) > 0 {
		existing := make(map[string]struct{}, len(f.opaqueDefaults))
		for _, d := range f.opaqueDefaults {
			existing[strings.ToLower(d.Extension)] = struct{}{}
		}
		for _, d := range srcFile.opaqueDefaults {
			if _, ok := existing[strings.ToLower(d.Extension)]; !ok {
				f.opaqueDefaults = append(f.opaqueDefaults, d)
				existing[strings.ToLower(d.Extension)] = struct{}{}
			}
		}
	}

	return dstRels
}
