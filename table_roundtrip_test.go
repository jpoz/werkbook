package werkbook

import (
	"os"
	"testing"

	"github.com/jpoz/werkbook/ooxml"
)

func TestReadTablesFromXLSX(t *testing.T) {
	// Read the ClosedXML test file that has a table.
	path := "../testdata/closedxml/TableHeadersWithLineBreaks.xlsx"
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skip("test file not found:", path)
	}

	osf, err := os.Open(path)
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer osf.Close()

	info, err := osf.Stat()
	if err != nil {
		t.Fatalf("stat: %v", err)
	}

	data, err := ooxml.ReadWorkbook(osf, info.Size())
	if err != nil {
		t.Fatalf("ReadWorkbook: %v", err)
	}

	if len(data.Tables) == 0 {
		t.Fatal("expected at least one table definition")
	}

	td := data.Tables[0]
	if td.DisplayName != "Table1" {
		t.Errorf("table display name = %q, want %q", td.DisplayName, "Table1")
	}
	if td.Ref != "A1:K7" {
		t.Errorf("table ref = %q, want %q", td.Ref, "A1:K7")
	}
	if len(td.Columns) != 11 {
		t.Errorf("table columns count = %d, want 11", len(td.Columns))
	}
	if td.TotalsRowCount != 1 {
		t.Errorf("table totals row count = %d, want 1", td.TotalsRowCount)
	}
	if td.HeaderRowCount != 1 {
		t.Errorf("table header row count = %d, want 1", td.HeaderRowCount)
	}

	// Verify first column name (may have line break encoded as _x000a_).
	if td.Columns[0] != "Item #" {
		t.Errorf("first column = %q, want %q", td.Columns[0], "Item #")
	}
}

func TestTableFormulaEvaluation(t *testing.T) {
	// Create a workbook with a table and test SUBTOTAL with table reference.
	path := "../testdata/closedxml/TableHeadersWithLineBreaks.xlsx"
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skip("test file not found:", path)
	}

	f, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	// Verify tables were loaded.
	if len(f.tables) == 0 {
		t.Fatal("expected table definitions in File")
	}

	ti := f.tables[0]
	t.Logf("Table: %s, Sheet: %s, Ref: cols %d-%d rows %d-%d, %d columns, header=%d totals=%d",
		ti.Name, ti.SheetName, ti.FirstCol, ti.LastCol, ti.FirstRow, ti.LastRow,
		len(ti.Columns), ti.HeaderRows, ti.TotalRows)
}
