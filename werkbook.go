package werkbook

import (
	"fmt"
	"os"
	"strconv"

	"github.com/jpoz/werkbook/ooxml"
)

// File represents an XLSX workbook.
type File struct {
	sheets     []*Sheet
	sheetNames []string
}

// Option configures a new workbook created by New.
type Option func(*options)

type options struct {
	firstSheet string
}

// FirstSheet sets the name of the initial sheet (default "Sheet1").
func FirstSheet(name string) Option {
	return func(o *options) {
		o.firstSheet = name
	}
}

// New creates a new workbook with one empty sheet.
// By default the sheet is named "Sheet1"; use FirstSheet to override.
func New(opts ...Option) *File {
	o := options{firstSheet: "Sheet1"}
	for _, fn := range opts {
		fn(&o)
	}
	f := &File{}
	f.addSheet(o.firstSheet)
	return f
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

// NewSheet adds a new empty sheet with the given name.
// Returns an error if a sheet with that name already exists.
func (f *File) NewSheet(name string) (*Sheet, error) {
	for _, n := range f.sheetNames {
		if n == name {
			return nil, fmt.Errorf("sheet %q already exists", name)
		}
	}
	return f.addSheet(name), nil
}

// DeleteSheet removes the sheet with the given name.
// Returns an error if the sheet does not exist or is the only sheet.
func (f *File) DeleteSheet(name string) error {
	if len(f.sheets) <= 1 {
		return fmt.Errorf("cannot delete the only sheet")
	}
	for i, s := range f.sheets {
		if s.name == name {
			f.sheets = append(f.sheets[:i], f.sheets[i+1:]...)
			f.sheetNames = append(f.sheetNames[:i], f.sheetNames[i+1:]...)
			return nil
		}
	}
	return ErrSheetNotFound
}

func (f *File) addSheet(name string) *Sheet {
	s := newSheet(name)
	f.sheets = append(f.sheets, s)
	f.sheetNames = append(f.sheetNames, name)
	return s
}

// SaveAs writes the workbook to the given file path.
func (f *File) SaveAs(name string) error {
	out, err := os.Create(name)
	if err != nil {
		return err
	}
	defer out.Close()

	data := f.buildWorkbookData()
	return ooxml.WriteWorkbook(out, data)
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

	data, err := ooxml.ReadWorkbook(osf, info.Size())
	if err != nil {
		return nil, err
	}

	return fileFromData(data), nil
}

func fileFromData(data *ooxml.WorkbookData) *File {
	f := &File{}
	for _, sd := range data.Sheets {
		s := f.addSheet(sd.Name)
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				col, row, err := CellNameToCoordinates(cd.Ref)
				if err != nil {
					continue
				}
				v := cellDataToValue(cd)
				r := s.ensureRow(row)
				c := r.ensureCell(col)
				c.value = v
				c.formula = cd.Formula
			}
		}
	}
	return f
}

func cellDataToValue(cd ooxml.CellData) Value {
	switch cd.Type {
	case "s", "inlineStr":
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

func (f *File) buildWorkbookData() *ooxml.WorkbookData {
	data := &ooxml.WorkbookData{}
	for _, s := range f.sheets {
		data.Sheets = append(data.Sheets, s.toSheetData())
	}
	return data
}
