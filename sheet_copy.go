package werkbook

import (
	"bytes"
	"fmt"
	"path"
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

// resolveRelTarget resolves a sheet relationship target to a zip path.
// Relative targets resolve against the standard sheet directory
// (xl/worksheets); targets with a leading "/" are already rooted at the
// package. The result never climbs above the package root.
func resolveRelTarget(target string) string {
	if !strings.HasPrefix(target, "/") {
		target = "/xl/worksheets/" + target
	}
	return strings.TrimPrefix(path.Clean(target), "/")
}

// absoluteToRelTarget converts an absolute zip path back to a relative target
// from the xl/worksheets directory.
func absoluteToRelTarget(absPath string) string {
	if strings.HasPrefix(absPath, "xl/") {
		return "../" + absPath[len("xl/"):]
	}
	return "../../" + absPath
}

// importOpaqueEntries copies the opaque parts that srcRels refer to from
// srcFile into f, and returns the relationships for the destination sheet.
//
// A part is imported together with its relationship graph: its own .rels
// file and the parts that this file refers to (for example the charts and
// images of a drawing), and their relationships. If a path is in use in the
// destination, the part gets a new name, and the relationships that refer to
// it are changed. External relationships are copied as they are.
func (f *File) importOpaqueEntries(srcFile *File, srcRels []ooxml.OpaqueRel) []ooxml.OpaqueRel {
	if len(srcRels) == 0 {
		return nil
	}

	imp := newOpaquePartImporter(f, srcFile)
	dstRels := make([]ooxml.OpaqueRel, len(srcRels))
	for i, rel := range srcRels {
		dstRels[i] = rel
		if rel.TargetMode == "External" {
			continue
		}
		srcPath := resolveRelTarget(rel.Target)
		dstPath, ok := imp.importPart(srcPath)
		if !ok || dstPath == srcPath {
			continue
		}
		dstRels[i].Target = absoluteToRelTarget(dstPath)
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

// opaquePartImporter copies opaque parts and their relationship graphs from
// one workbook into another.
type opaquePartImporter struct {
	dst       *File
	srcByPath map[string]*ooxml.OpaqueEntry
	dstByPath map[string][]byte
	// imported maps the path of a part in the source to its path in the
	// destination. A part is in the map from the start of its import, so a
	// cycle in the relationship graph stops.
	imported map[string]string
}

func newOpaquePartImporter(dst, src *File) *opaquePartImporter {
	imp := &opaquePartImporter{
		dst:       dst,
		srcByPath: make(map[string]*ooxml.OpaqueEntry, len(src.opaqueEntries)),
		dstByPath: make(map[string][]byte, len(dst.opaqueEntries)),
		imported:  make(map[string]string),
	}
	for i := range src.opaqueEntries {
		imp.srcByPath[src.opaqueEntries[i].Path] = &src.opaqueEntries[i]
	}
	for _, e := range dst.opaqueEntries {
		imp.dstByPath[e.Path] = e.Data
	}
	return imp
}

// partRelsPath returns the path of the relationships file of a part:
// xl/drawings/drawing1.xml gives xl/drawings/_rels/drawing1.xml.rels.
func partRelsPath(partPath string) string {
	dir, base := path.Split(partPath)
	return dir + "_rels/" + base + ".rels"
}

// resolvePartRelTarget resolves a relationship target of a part to a path in
// the package. A relative target is relative to the directory of the part.
func resolvePartRelTarget(partPath, target string) string {
	if !strings.HasPrefix(target, "/") {
		target = "/" + path.Dir(partPath) + "/" + target
	}
	return strings.TrimPrefix(path.Clean(target), "/")
}

// renamedRelTarget returns the target for a part that got a new name in the
// same directory. It keeps the form (relative or absolute) of the original
// target.
func renamedRelTarget(target, newPath string) string {
	if strings.HasPrefix(target, "/") {
		return "/" + newPath
	}
	return target[:strings.LastIndex(target, "/")+1] + path.Base(newPath)
}

// importPart copies the part at srcPath and its relationship graph into the
// destination. It returns the path of the part in the destination. It returns
// false if the source has no such part.
func (imp *opaquePartImporter) importPart(srcPath string) (string, bool) {
	if dstPath, ok := imp.imported[srcPath]; ok {
		return dstPath, true
	}
	srcEntry, ok := imp.srcByPath[srcPath]
	if !ok {
		return "", false
	}
	imp.imported[srcPath] = srcPath

	// Import the parts that this part refers to first. Their paths in the
	// destination decide the content of the relationships file.
	var relsData []byte
	srcRelsEntry, hasRels := imp.srcByPath[partRelsPath(srcPath)]
	if hasRels {
		relsData = imp.importPartRels(srcPath, srcRelsEntry.Data)
	}

	dstPath := srcPath
	if existing, found := imp.dstByPath[srcPath]; found {
		existingRels, existingHasRels := imp.dstByPath[partRelsPath(srcPath)]
		same := bytes.Equal(existing, srcEntry.Data) &&
			existingHasRels == hasRels &&
			bytes.Equal(existingRels, relsData)
		if same {
			// The destination has this part with the same relationships.
			return dstPath, true
		}
		dstPath = imp.uniquePartPath(srcPath)
	}

	imp.addEntry(dstPath, srcEntry.ContentType, srcEntry.Data)
	if hasRels {
		imp.addEntry(partRelsPath(dstPath), srcRelsEntry.ContentType, relsData)
	}
	imp.imported[srcPath] = dstPath
	return dstPath, true
}

// importPartRels imports the parts that the relationships file of a part
// refers to. It returns the content of the relationships file for the
// destination: the original bytes, or new bytes if a target got a new name.
func (imp *opaquePartImporter) importPartRels(partPath string, relsData []byte) []byte {
	rels, err := ooxml.ParseRelationships(relsData)
	if err != nil {
		// Not a relationships file that can be read. Copy it as it is.
		return relsData
	}
	changed := false
	for i, rel := range rels {
		if rel.TargetMode == "External" {
			continue
		}
		childPath := resolvePartRelTarget(partPath, rel.Target)
		dstChildPath, ok := imp.importPart(childPath)
		if !ok || dstChildPath == childPath {
			continue
		}
		rels[i].Target = renamedRelTarget(rel.Target, dstChildPath)
		changed = true
	}
	if !changed {
		return relsData
	}
	out, err := ooxml.MarshalRelationships(rels)
	if err != nil {
		return relsData
	}
	return out
}

// uniquePartPath returns a path that no part and no relationships file in
// the destination uses. It appends _2, _3, and so on before the extension.
func (imp *opaquePartImporter) uniquePartPath(partPath string) string {
	ext := path.Ext(partPath)
	base := strings.TrimSuffix(partPath, ext)
	for n := 2; ; n++ {
		candidate := fmt.Sprintf("%s_%d%s", base, n, ext)
		if _, ok := imp.dstByPath[candidate]; ok {
			continue
		}
		if _, ok := imp.dstByPath[partRelsPath(candidate)]; ok {
			continue
		}
		return candidate
	}
}

func (imp *opaquePartImporter) addEntry(entryPath, contentType string, data []byte) {
	copied := append([]byte(nil), data...)
	imp.dst.opaqueEntries = append(imp.dst.opaqueEntries, ooxml.OpaqueEntry{
		Path:        entryPath,
		ContentType: contentType,
		Data:        copied,
	})
	imp.dstByPath[entryPath] = copied
}
