package werkbook

import (
	"bytes"
	"errors"
	"strings"
	"testing"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

func TestOpenReaderAtRejectsOversizedDefinedNameExpansion(t *testing.T) {
	data := &ooxml.WorkbookData{
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{
						Num: 1,
						Cells: []ooxml.CellData{
							{Ref: "A1", Formula: "Huge+Huge"},
						},
					},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{
				Name:         "Huge",
				Value:        strings.Repeat("1", formula.MaxExpandedFormulaBytes/2+1),
				LocalSheetID: -1,
			},
		},
	}

	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, data); err != nil {
		t.Fatalf("WriteWorkbook error: %v", err)
	}

	_, err := OpenReaderAt(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if !errors.Is(err, ErrFormulaTooLarge) {
		t.Fatalf("expected ErrFormulaTooLarge, got %v", err)
	}
}
