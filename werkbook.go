package werkbook

import (
	"bytes"
	"fmt"
	"io"
	"os"
	"strconv"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// File represents an XLSX workbook.
type File struct {
	sheets       []*Sheet
	sheetNames   []string
	date1904     bool // true if the workbook uses the 1904 date system (Mac Excel)
	calcProps    CalcProperties
	calcGen      uint64            // incremented on any cell mutation; starts at 1
	evaluating   map[cellKey]bool  // tracks cells being evaluated (circular ref detection)
	deps         *formula.DepGraph // cell dependency graph for incremental recalculation
	tableDefs    []Table
	tables       []formula.TableInfo       // table definitions for structured reference expansion
	definedNames []formula.DefinedNameInfo // defined names (named ranges) for formula expansion
}

// cellKey identifies a cell across the entire workbook for circular ref detection.
type cellKey struct {
	sheet string
	col   int
	row   int
}

// Option configures a new workbook created by New.
type Option func(*options)

type options struct {
	firstSheet string
	date1904   bool
}

// FirstSheet sets the name of the initial sheet (default "Sheet1").
func FirstSheet(name string) Option {
	return func(o *options) {
		o.firstSheet = name
	}
}

// WithDate1904 configures a newly created workbook to use the 1904 date system.
func WithDate1904(enabled bool) Option {
	return func(o *options) {
		o.date1904 = enabled
	}
}

// New creates a new workbook with one empty sheet.
// By default the sheet is named "Sheet1"; use FirstSheet to override.
func New(opts ...Option) *File {
	o := options{firstSheet: "Sheet1"}
	for _, fn := range opts {
		fn(&o)
	}
	f := &File{calcGen: 1, date1904: o.date1904, deps: formula.NewDepGraph()}
	f.addSheet(o.firstSheet)
	return f
}

// Date1904 reports whether the workbook uses the 1904 date system (Mac Excel).
func (f *File) Date1904() bool {
	return f.date1904
}

// Sheet returns the sheet with the given name, or nil if not found.
func (f *File) Sheet(name string) *Sheet {
	for _, s := range f.sheets {
		if s.name == name {
			return s
		}
	}
	return nil
}

// SheetNames returns the names of all sheets in order.
func (f *File) SheetNames() []string {
	names := make([]string, len(f.sheetNames))
	copy(names, f.sheetNames)
	return names
}

// SheetIndex returns the 0-based index of the named sheet, or -1 if not found.
func (f *File) SheetIndex(name string) int {
	for i, n := range f.sheetNames {
		if n == name {
			return i
		}
	}
	return -1
}

// NewSheet adds a new empty sheet with the given name.
// Returns an error if a sheet with that name already exists.
func (f *File) NewSheet(name string) (*Sheet, error) {
	if f.SheetIndex(name) >= 0 {
		return nil, fmt.Errorf("sheet %q already exists", name)
	}
	return f.addSheet(name), nil
}

// DeleteSheet removes the sheet with the given name.
// Returns an error if the sheet does not exist or is the only sheet.
func (f *File) DeleteSheet(name string) error {
	if len(f.sheets) <= 1 {
		return fmt.Errorf("cannot delete the only sheet")
	}
	idx := f.SheetIndex(name)
	if idx < 0 {
		return ErrSheetNotFound
	}

	f.sheets = append(f.sheets[:idx], f.sheets[idx+1:]...)
	f.sheetNames = append(f.sheetNames[:idx], f.sheetNames[idx+1:]...)

	// Remove table metadata for the deleted sheet.
	filteredTables := f.tables[:0]
	for _, t := range f.tables {
		if t.SheetName != name {
			filteredTables = append(filteredTables, t)
		}
	}
	f.tables = filteredTables
	filteredTableDefs := f.tableDefs[:0]
	for _, t := range f.tableDefs {
		if t.SheetName != name {
			filteredTableDefs = append(filteredTableDefs, t)
		}
	}
	f.tableDefs = filteredTableDefs

	// Sheet-scoped names on the deleted sheet disappear; later sheet indexes shift down.
	filteredNames := f.definedNames[:0]
	for _, dn := range f.definedNames {
		switch {
		case dn.LocalSheetID == idx:
			continue
		case dn.LocalSheetID > idx:
			dn.LocalSheetID--
		}
		filteredNames = append(filteredNames, dn)
	}
	f.definedNames = filteredNames

	f.rebuildFormulaState()
	return nil
}

func (f *File) addSheet(name string) *Sheet {
	s := newSheet(name, f)
	f.sheets = append(f.sheets, s)
	f.sheetNames = append(f.sheetNames, name)
	return s
}

// SetSheetName renames a sheet.
func (f *File) SetSheetName(old, new string) error {
	if old == new {
		if f.SheetIndex(old) < 0 {
			return ErrSheetNotFound
		}
		return nil
	}
	idx := f.SheetIndex(old)
	if idx < 0 {
		return ErrSheetNotFound
	}
	if f.SheetIndex(new) >= 0 {
		return fmt.Errorf("sheet %q already exists", new)
	}

	f.sheets[idx].name = new
	f.sheetNames[idx] = new
	for i := range f.tables {
		if f.tables[i].SheetName == old {
			f.tables[i].SheetName = new
		}
	}
	for i := range f.tableDefs {
		if f.tableDefs[i].SheetName == old {
			f.tableDefs[i].SheetName = new
		}
	}

	f.rebuildFormulaState()
	return nil
}

// SetSheetVisible sets whether a sheet is visible.
func (f *File) SetSheetVisible(name string, visible bool) error {
	s := f.Sheet(name)
	if s == nil {
		return ErrSheetNotFound
	}
	s.visible = visible
	return nil
}

// SaveAs writes the workbook to the given file path.
func (f *File) SaveAs(name string) error {
	out, err := os.Create(name)
	if err != nil {
		return err
	}

	data := f.buildWorkbookData()
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		return err
	}
	return out.Close()
}

// WriteTo serializes the workbook to the given writer.
func (f *File) WriteTo(w io.Writer) error {
	return ooxml.WriteWorkbook(w, f.buildWorkbookData())
}

// Open opens an existing XLSX file for reading.
func Open(name string) (*File, error) {
	osf, err := os.Open(name)
	if err != nil {
		return nil, err
	}
	defer osf.Close()

	info, err := osf.Stat()
	if err != nil {
		return nil, err
	}

	return OpenReaderAt(osf, info.Size())
}

// OpenReader opens an XLSX from an arbitrary reader.
func OpenReader(r io.Reader) (*File, error) {
	buf, err := io.ReadAll(r)
	if err != nil {
		return nil, err
	}
	return OpenReaderAt(bytes.NewReader(buf), int64(len(buf)))
}

// OpenReaderAt opens an XLSX from a random-access reader.
func OpenReaderAt(r io.ReaderAt, size int64) (*File, error) {
	data, err := ooxml.ReadWorkbook(r, size)
	if err != nil {
		return nil, err
	}
	return fileFromData(data), nil
}

func fileFromData(data *ooxml.WorkbookData) *File {
	f := &File{
		calcGen:   1,
		date1904:  data.Date1904,
		calcProps: calcPropsFromData(data.CalcProps),
		deps:      formula.NewDepGraph(),
	}

	// Convert StyleData slice to *Style slice for assignment.
	var parsedStyles []*Style
	if len(data.Styles) > 0 {
		parsedStyles = make([]*Style, len(data.Styles))
		for i, sd := range data.Styles {
			parsedStyles[i] = styleDataToStyle(sd)
		}
	}

	for _, sd := range data.Sheets {
		s := f.addSheet(sd.Name)
		s.visible = sd.State != "hidden" && sd.State != "veryHidden"

		// Restore column widths.
		for _, cw := range sd.ColWidths {
			for col := cw.Min; col <= cw.Max; col++ {
				s.colWidths[col] = cw.Width
			}
		}

		// Restore merged cells.
		if len(sd.MergeCells) > 0 {
			s.merges = make([]MergeRange, len(sd.MergeCells))
			for i, mc := range sd.MergeCells {
				s.merges[i] = MergeRange{
					Start: mc.StartAxis,
					End:   mc.EndAxis,
				}
			}
		}

		for _, rd := range sd.Rows {
			// Restore row height and hidden state.
			if rd.Height != 0 || rd.Hidden {
				r := s.ensureRow(rd.Num)
				r.height = rd.Height
				r.hidden = rd.Hidden
			}

			for _, cd := range rd.Cells {
				col, row, err := CellNameToCoordinates(cd.Ref)
				if err != nil {
					continue
				}
				v := cellDataToValue(cd, parsedStyles, f.date1904)
				r := s.ensureRow(row)
				c := r.ensureCell(col)
				c.value = v
				c.formula = formula.StripXlfnPrefixes(cd.Formula)
				c.isArrayFormula = cd.IsArrayFormula
				c.formulaRef = cd.FormulaRef
				// Trust the file's cached value for formula cells that have one.
				if cd.Formula != "" && v.Type != TypeEmpty {
					c.cachedGen = f.calcGen
				}
				// Assign style if non-default.
				if cd.StyleIdx > 0 && cd.StyleIdx < len(parsedStyles) {
					c.style = parsedStyles[cd.StyleIdx]
				}
			}
		}
	}
	// Build table info from parsed table definitions.
	for _, td := range data.Tables {
		col1, row1, col2, row2, err := RangeToCoordinates(td.Ref)
		if err != nil {
			continue
		}
		sheetName := ""
		if td.SheetIndex >= 0 && td.SheetIndex < len(data.Sheets) {
			sheetName = data.Sheets[td.SheetIndex].Name
		}
		ti := formula.TableInfo{
			Name:            td.DisplayName,
			SheetName:       sheetName,
			Columns:         td.Columns,
			FirstCol:        col1,
			FirstRow:        row1,
			LastCol:         col2,
			LastRow:         row2,
			HeaderRows:      td.HeaderRowCount,
			TotalRows:       td.TotalsRowCount,
			HasActiveFilter: td.HasActiveFilter,
		}
		f.tableDefs = append(f.tableDefs, tableFromData(td, sheetName))
		f.tables = append(f.tables, ti)
	}

	// Build defined name info from parsed defined names.
	for _, dn := range data.DefinedNames {
		f.definedNames = append(f.definedNames, formula.DefinedNameInfo{
			Name:         dn.Name,
			Value:        dn.Value,
			LocalSheetID: dn.LocalSheetID,
		})
	}

	f.registerAllFormulas()
	return f
}

func cellDataToValue(cd ooxml.CellData, _ []*Style, date1904 bool) Value {
	switch cd.Type {
	case "s":
		// Shared-string cells are always strings. In Excel, a cell stored
		// as a shared string is text even if the content looks numeric —
		// e.g. the string "1" is not equal to the number 1 via the = operator.
		// Arithmetic operations will coerce via CoerceNum as needed.
		return Value{Type: TypeString, String: cd.Value}
	case "str", "inlineStr":
		if cd.Formula != "" && (cd.IsDynamicArray || cd.IsArrayFormula || cd.FormulaType == "array") && isFormulaErrorString(cd.Value) {
			return Value{Type: TypeError, String: cd.Value}
		}
		return Value{Type: TypeString, String: cd.Value}
	case "d":
		n, err := excelDateStringToSerial(cd.Value, date1904)
		if err == nil {
			return Value{Type: TypeNumber, Number: n}
		}
		return Value{Type: TypeString, String: cd.Value}
	case "b":
		return Value{Type: TypeBool, Bool: cd.Value == "1"}
	case "e":
		return Value{Type: TypeError, String: cd.Value}
	default:
		// Number or empty.
		if cd.Value == "" {
			return Value{Type: TypeEmpty}
		}
		n, err := strconv.ParseFloat(cd.Value, 64)
		if err != nil {
			// If we can't parse as number, treat as string.
			return Value{Type: TypeString, String: cd.Value}
		}
		return Value{Type: TypeNumber, Number: n}
	}
}

func isFormulaErrorString(v string) bool {
	switch v {
	case "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!", "#GETTING_DATA":
		return true
	default:
		return false
	}
}

func (f *File) buildWorkbookData() *ooxml.WorkbookData {
	data := &ooxml.WorkbookData{
		Date1904:  f.date1904,
		CalcProps: f.calcProps.toData(),
	}

	// Style dedup: index 0 is always the default (empty StyleData).
	styles := []ooxml.StyleData{{}}
	styleMap := map[string]int{styleKey(ooxml.StyleData{}): 0}

	for _, s := range f.sheets {
		data.Sheets = append(data.Sheets, s.toSheetData(styleMap, &styles))
	}
	data.Styles = styles

	// Preserve defined names.
	for _, dn := range f.definedNames {
		data.DefinedNames = append(data.DefinedNames, ooxml.DefinedName{
			Name:         dn.Name,
			Value:        dn.Value,
			LocalSheetID: dn.LocalSheetID,
		})
	}
	for _, td := range f.tableDefs {
		sheetIdx := f.SheetIndex(td.SheetName)
		if sheetIdx < 0 {
			continue
		}
		data.Tables = append(data.Tables, td.toData(sheetIdx))
	}

	return data
}

// registerAllFormulas iterates all cells and registers compiled formulas in
// the dependency graph. Called at the end of fileFromData.
func (f *File) rebuildFormulaState() {
	f.calcGen++
	f.deps = formula.NewDepGraph()
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula == "" {
					continue
				}
				c.compiled = nil
				c.cachedGen = 0
				c.dirty = true
			}
		}
	}
	f.registerAllFormulas()
}

// expandFormula expands table refs and defined names in a formula string.
func (f *File) expandFormula(src string, sheetName string, row int) string {
	src = formula.ExpandTableRefs(src, f.tables, row)
	src = formula.ExpandDefinedNames(src, f.definedNames, f.SheetIndex(sheetName))
	return src
}

func (f *File) registerAllFormulas() {
	for _, s := range f.sheets {
		f.registerSheetFormulas(s)
	}
}

// invalidateDependents queries the dep graph for all transitive dependents
// of the given cell and marks them dirty.
func (f *File) invalidateDependents(sheet string, col, row int) {
	changed := formula.QualifiedCell{Sheet: sheet, Col: col, Row: row}
	for _, dep := range f.deps.Dependents(changed) {
		s := f.Sheet(dep.Sheet)
		if s == nil {
			continue
		}
		r, ok := s.rows[dep.Row]
		if !ok {
			continue
		}
		c, ok := r.cells[dep.Col]
		if !ok {
			continue
		}
		c.dirty = true
	}
}

// Precedents returns the cells and ranges that the formula in the given cell reads from.
// Returns nil slices if the cell has no formula or no registered dependencies.
func (f *File) Precedents(sheet, cell string) ([]formula.QualifiedCell, []formula.RangeAddr, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return nil, nil, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	qc := formula.QualifiedCell{Sheet: sheet, Col: col, Row: row}
	points, ranges := f.deps.DependsOn(qc)
	return points, ranges, nil
}

// DirectDependents returns the cells whose formulas directly read the given cell.
func (f *File) DirectDependents(sheet, cell string) ([]formula.QualifiedCell, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	qc := formula.QualifiedCell{Sheet: sheet, Col: col, Row: row}
	return f.deps.DirectDependents(qc), nil
}

// Recalculate evaluates all dirty formula cells. Cells are evaluated lazily
// via GetValue, but this method forces evaluation of every dirty cell.
func (f *File) Recalculate() {
	f.calcGen++
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for col, c := range r.cells {
				if c.formula != "" && (c.dirty || c.cachedGen < f.calcGen) {
					c.value = s.evaluateFormula(c, col, r.num)
					c.cachedGen = f.calcGen
					c.dirty = false
				}
			}
		}
	}
}
