package formula

import (
	"testing"
)

func TestExpandTableRefs(t *testing.T) {
	// Define a sample table: Sales table on Sheet1, A1:E20, 1 header row, 0 total rows.
	// Columns: ID, Product, Sales, Cost, Profit
	tables := []TableInfo{
		{
			Name:       "Table6",
			SheetName:  "Sheet1",
			Columns:    []string{"ID", "Product", "Sales", "Cost", "Profit"},
			FirstCol:   1, // A
			FirstRow:   1,
			LastCol:    5, // E
			LastRow:    20,
			HeaderRows: 1,
			TotalRows:  0,
		},
		{
			Name:       "HedgeResults",
			SheetName:  "Data",
			Columns:    []string{"CurveType", "ShockDate", "LiabilityRandsPerPoint"},
			FirstCol:   1,
			FirstRow:   1,
			LastCol:    3,
			LastRow:    100,
			HeaderRows: 1,
			TotalRows:  0,
		},
	}

	tests := []struct {
		name    string
		formula string
		row     int
		want    string
	}{
		{
			name:    "simple column ref",
			formula: "SUM(Table6[Sales])",
			row:     5,
			want:    "SUM(Sheet1!C2:C20)",
		},
		{
			name:    "this row ref",
			formula: "Table6[[#This Row],[Sales]]",
			row:     5,
			want:    "Sheet1!C5",
		},
		{
			name:    "this row ref in arithmetic",
			formula: "Table6[[#This Row],[Sales]]+Table6[[#This Row],[Cost]]",
			row:     10,
			want:    "Sheet1!C10+Sheet1!D10",
		},
		{
			name:    "no table ref - passthrough",
			formula: "SUM(A1:A10)",
			row:     1,
			want:    "SUM(A1:A10)",
		},
		{
			name:    "string literal not expanded",
			formula: `"Table6[Sales]"`,
			row:     1,
			want:    `"Table6[Sales]"`,
		},
		{
			name:    "multiple table refs",
			formula: "SUMIFS(HedgeResults[LiabilityRandsPerPoint],HedgeResults[CurveType],\"Nominal\")",
			row:     5,
			want:    "SUMIFS(Data!C2:C100,Data!A2:A100,\"Nominal\")",
		},
		{
			name:    "unknown column passthrough",
			formula: "Table6[UnknownCol]",
			row:     5,
			want:    "Table6[UnknownCol]",
		},
		{
			name:    "unknown table passthrough",
			formula: "Unknown[Sales]",
			row:     5,
			want:    "Unknown[Sales]",
		},
		{
			// A table-name substring inside a quoted sheet name must NOT
			// trigger structured-ref expansion. Mirrors the same guard in
			// ExpandDefinedNamesBounded.
			name:    "quoted sheet name containing table name is left alone",
			formula: "'my-Table6'!A1",
			row:     5,
			want:    "'my-Table6'!A1",
		},
		{
			name:    "escaped '' inside quoted sheet name is preserved",
			formula: "'Bob''s-Table6'!A1",
			row:     5,
			want:    "'Bob''s-Table6'!A1",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ExpandTableRefs(tt.formula, tables, tt.row)
			if got != tt.want {
				t.Errorf("ExpandTableRefs(%q) = %q, want %q", tt.formula, got, tt.want)
			}
		})
	}
}

func TestExpandTableRefsWithTotals(t *testing.T) {
	tables := []TableInfo{
		{
			Name:       "Table1",
			SheetName:  "Sheet1",
			Columns:    []string{"Item", "Price"},
			FirstCol:   1,
			FirstRow:   1,
			LastCol:    2,
			LastRow:    7,
			HeaderRows: 1,
			TotalRows:  1,
		},
	}

	tests := []struct {
		name    string
		formula string
		row     int
		want    string
	}{
		{
			name:    "data column excludes totals",
			formula: "SUM(Table1[Price])",
			row:     3,
			want:    "SUM(Sheet1!B2:B6)",
		},
		{
			name:    "totals ref",
			formula: "Table1[[#Totals],[Price]]",
			row:     3,
			want:    "Sheet1!B7",
		},
		{
			name:    "headers ref",
			formula: "Table1[[#Headers],[Item]]",
			row:     3,
			want:    "Sheet1!A1",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ExpandTableRefs(tt.formula, tables, tt.row)
			if got != tt.want {
				t.Errorf("ExpandTableRefs(%q) = %q, want %q", tt.formula, got, tt.want)
			}
		})
	}
}
