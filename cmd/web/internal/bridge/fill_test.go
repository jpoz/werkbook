package bridge

import (
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

func mustFill(t *testing.T, s *Session, sheet, src, dst string, series bool) FillResult {
	t.Helper()
	res, err := s.Fill(sheet, rangeOf(t, src), rangeOf(t, dst), series)
	if err != nil {
		t.Fatalf("Fill(%s -> %s): %v", src, dst, err)
	}
	return res
}

func TestFillDirectionsAndFormulaShift(t *testing.T) {
	tests := []struct {
		name     string
		set      []string // cell, input pairs
		src, dst string
		want     map[string]string // inputs of all cells after the fill
	}{
		{
			name: "down",
			set:  []string{"B2", "=A2*2"},
			src:  "B2", dst: "B2:B5",
			want: map[string]string{"B2": "=A2*2", "B3": "=A3*2", "B4": "=A4*2", "B5": "=A5*2"},
		},
		{
			name: "up",
			set:  []string{"B5", "=A5*2"},
			src:  "B5", dst: "B3:B5",
			want: map[string]string{"B5": "=A5*2", "B4": "=A4*2", "B3": "=A3*2"},
		},
		{
			name: "right",
			set:  []string{"B2", "=B1+1"},
			src:  "B2", dst: "B2:E2",
			want: map[string]string{"B2": "=B1+1", "C2": "=C1+1", "D2": "=D1+1", "E2": "=E1+1"},
		},
		{
			name: "left",
			set:  []string{"E2", "=E1+1"},
			src:  "E2", dst: "C2:E2",
			want: map[string]string{"E2": "=E1+1", "D2": "=D1+1", "C2": "=C1+1"},
		},
		{
			name: "absolute and mixed forms, down",
			set:  []string{"C1", "=$A$1+A1+$A1+A$1"},
			src:  "C1", dst: "C1:C3",
			want: map[string]string{"C1": "=$A$1+A1+$A1+A$1", "C2": "=$A$1+A2+$A2+A$1", "C3": "=$A$1+A3+$A3+A$1"},
		},
		{
			name: "absolute and mixed forms, right",
			set:  []string{"C1", "=$A$1+A1+$A1+A$1"},
			src:  "C1", dst: "C1:E1",
			want: map[string]string{"C1": "=$A$1+A1+$A1+A$1", "D1": "=$A$1+B1+$A1+B$1", "E1": "=$A$1+C1+$A1+C$1"},
		},
		{
			name: "reference that leaves the sheet becomes #REF!",
			set:  []string{"B3", "=A1+B2"},
			src:  "B3", dst: "B1:B3",
			want: map[string]string{"B3": "=A1+B2", "B2": "=#REF!+B1", "B1": "=#REF!+#REF!"},
		},
		{
			name: "reference that leaves the sheet at the left",
			set:  []string{"C1", "=A1"},
			src:  "C1", dst: "A1:C1",
			want: map[string]string{"C1": "=A1", "B1": "=#REF!", "A1": "=#REF!"},
		},
		{
			name: "other text of the formula stays",
			set:  []string{"A1", `= if( b1>0 , "B1" , Sheet1!c1 )`},
			src:  "A1", dst: "A1:A2",
			want: map[string]string{"A1": `= if( b1>0 , "B1" , Sheet1!c1 )`, "A2": `= if( B2>0 , "B1" , Sheet1!C2 )`},
		},
		{
			name: "values copy",
			set:  []string{"A1", "text", "B1", "TRUE", "C1", "7"},
			src:  "A1:C1", dst: "A1:C3",
			want: map[string]string{
				"A1": "text", "B1": "TRUE", "C1": "7",
				"A2": "text", "B2": "TRUE", "C2": "7",
				"A3": "text", "B3": "TRUE", "C3": "7",
			},
		},
		{
			name: "pattern repeats down",
			set:  []string{"A1", "x", "A2", "=B2", "A3", "z"},
			src:  "A1:A3", dst: "A1:A8",
			want: map[string]string{
				"A1": "x", "A2": "=B2", "A3": "z",
				"A4": "x", "A5": "=B5", "A6": "z",
				"A7": "x", "A8": "=B8",
			},
		},
		{
			name: "pattern repeats up",
			set:  []string{"A5", "x", "A6", "y"},
			src:  "A5:A6", dst: "A2:A6",
			want: map[string]string{"A5": "x", "A6": "y", "A4": "y", "A3": "x", "A2": "y"},
		},
		{
			name: "empty source cell clears the destination",
			set:  []string{"A1", "x", "A3", "old", "A4", "old", "A5", "old"},
			src:  "A1:A2", dst: "A1:A5",
			want: map[string]string{"A1": "x", "A3": "x", "A5": "x"},
		},
		{
			name: "error value copies as a value; a formula error copies the formula",
			set:  []string{"A1", "=1/0", "B1", "=NA()"},
			src:  "A1:B1", dst: "A1:B2",
			want: map[string]string{"A1": "=1/0", "B1": "=NA()", "A2": "=1/0", "B2": "=NA()"},
		},
		{
			name: "each column is its own lane",
			set:  []string{"A1", "1", "A2", "2", "B1", "a", "B2", "b", "C1", "=A1", "C2", "5"},
			src:  "A1:C2", dst: "A1:C4",
			want: map[string]string{
				"A1": "1", "A2": "2", "B1": "a", "B2": "b", "C1": "=A1", "C2": "5",
				"A3": "1", "A4": "2", "B3": "a", "B4": "b", "C3": "=A3", "C4": "5",
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, tt.set...)
			before := dump(t, s, sheet1)
			depth := s.History().Depth

			res := mustFill(t, s, sheet1, tt.src, tt.dst, false)
			if res.Range != rangeOf(t, tt.dst) {
				t.Errorf("range = %+v, want %s", res.Range, tt.dst)
			}
			if want := rangeOf(t, tt.dst).cells() - rangeOf(t, tt.src).cells(); res.Changed != want || res.Skipped != 0 {
				t.Errorf("changed %d skipped %d, want %d and 0", res.Changed, res.Skipped, want)
			}
			if res.History.Depth != depth+1 || res.History.UndoLabel != "Fill" {
				t.Errorf("history = %+v", res.History)
			}
			wantInputs(t, s, sheet1, tt.want)
			checkUndoRedo(t, s, sheet1, before)
			if undo := mustUndo(t, s); *undo.Select != rangeOf(t, tt.dst) {
				t.Errorf("undo select = %v, want %s", undo.Select, tt.dst)
			}
		})
	}
}

func TestFillSeries(t *testing.T) {
	dateStyle := &werkbook.Style{NumFmtID: 14}
	tests := []struct {
		name     string
		set      []string
		styles   map[string]*werkbook.Style
		src, dst string
		series   bool
		want     map[string]string
	}{
		{
			name: "two cells continue their step",
			set:  []string{"A1", "1", "A2", "2"},
			src:  "A1:A2", dst: "A1:A6", series: true,
			want: map[string]string{"A1": "1", "A2": "2", "A3": "3", "A4": "4", "A5": "5", "A6": "6"},
		},
		{
			name: "step of 5 to the right",
			set:  []string{"A1", "10", "B1", "15"},
			src:  "A1:B1", dst: "A1:E1", series: true,
			want: map[string]string{"A1": "10", "B1": "15", "C1": "20", "D1": "25", "E1": "30"},
		},
		{
			name: "negative step",
			set:  []string{"A1", "10", "A2", "7"},
			src:  "A1:A2", dst: "A1:A5", series: true,
			want: map[string]string{"A1": "10", "A2": "7", "A3": "4", "A4": "1", "A5": "-2"},
		},
		{
			name: "decimal step has no binary noise",
			set:  []string{"A1", "0.1", "A2", "0.2"},
			src:  "A1:A2", dst: "A1:A4", series: true,
			want: map[string]string{"A1": "0.1", "A2": "0.2", "A3": "0.3", "A4": "0.4"},
		},
		{
			name: "fill up continues the series backward",
			set:  []string{"A4", "10", "A5", "12"},
			src:  "A4:A5", dst: "A1:A5", series: true,
			want: map[string]string{"A4": "10", "A5": "12", "A3": "8", "A2": "6", "A1": "4"},
		},
		{
			name: "fill left continues the series backward",
			set:  []string{"D1", "1", "E1", "2"},
			src:  "D1:E1", dst: "B1:E1", series: true,
			want: map[string]string{"D1": "1", "E1": "2", "C1": "0", "B1": "-1"},
		},
		{
			// Least-squares line through (0,1), (1,2), (2,4): y = 5/6 + 1.5x.
			// 16/3 and 41/6 have no short decimal form: the stored values
			// are the calculated float64 values, not 15-digit numbers.
			name: "three cells that are not a line continue the linear trend",
			set:  []string{"A1", "1", "A2", "2", "A3", "4"},
			src:  "A1:A3", dst: "A1:A5", series: true,
			want: map[string]string{"A1": "1", "A2": "2", "A3": "4", "A4": "5.333333333333333", "A5": "6.833333333333333"},
		},
		{
			name: "decimal step of 0.7 has no binary noise",
			set:  []string{"A1", "0.7", "A2", "1.4"},
			src:  "A1:A2", dst: "A1:A5", series: true,
			want: map[string]string{"A1": "0.7", "A2": "1.4", "A3": "2.1", "A4": "2.8", "A5": "3.5"},
		},
		{
			name: "equal numbers stay equal",
			set:  []string{"A1", "5", "A2", "5"},
			src:  "A1:A2", dst: "A1:A4", series: true,
			want: map[string]string{"A1": "5", "A2": "5", "A3": "5", "A4": "5"},
		},
		{
			name: "one number copies",
			set:  []string{"A1", "7"},
			src:  "A1", dst: "A1:A3", series: true,
			want: map[string]string{"A1": "7", "A2": "7", "A3": "7"},
		},
		{
			name:   "one date counts days",
			set:    []string{"A1", "45292"},
			styles: map[string]*werkbook.Style{"A1": dateStyle},
			src:    "A1", dst: "A1:A3", series: true,
			want: map[string]string{"A1": "45292", "A2": "45293", "A3": "45294"},
		},
		{
			name:   "one date counts days backward for a fill up",
			set:    []string{"A3", "45292"},
			styles: map[string]*werkbook.Style{"A3": dateStyle},
			src:    "A3", dst: "A1:A3", series: true,
			want: map[string]string{"A3": "45292", "A2": "45291", "A1": "45290"},
		},
		{
			name:   "series false copies numbers and dates",
			set:    []string{"A1", "1", "A2", "2", "B1", "45292"},
			styles: map[string]*werkbook.Style{"B1": dateStyle},
			src:    "A1:B2", dst: "A1:B4", series: false,
			want: map[string]string{"A1": "1", "A2": "2", "B1": "45292", "A3": "1", "A4": "2", "B3": "45292"},
		},
		{
			name: "a lane with text copies; the number lane makes a series",
			set:  []string{"A1", "1", "A2", "2", "B1", "1", "B2", "x"},
			src:  "A1:B2", dst: "A1:B4", series: true,
			want: map[string]string{"A1": "1", "A2": "2", "B1": "1", "B2": "x", "A3": "3", "A4": "4", "B3": "1", "B4": "x"},
		},
		{
			name: "a lane with a formula copies",
			set:  []string{"A1", "1", "A2", "=A1+1"},
			src:  "A1:A2", dst: "A1:A4", series: true,
			want: map[string]string{"A1": "1", "A2": "=A1+1", "A3": "1", "A4": "=A3+1"},
		},
		{
			name: "a lane with an empty cell copies",
			set:  []string{"A1", "1", "A3", "3"},
			src:  "A1:A3", dst: "A1:A6", series: true,
			want: map[string]string{"A1": "1", "A3": "3", "A4": "1", "A6": "3"},
		},
		{
			name: "text that looks like a number is text",
			set:  []string{"A1", "'1", "A2", "'2"},
			src:  "A1:A2", dst: "A1:A4", series: true,
			want: map[string]string{"A1": "'1", "A2": "'2", "A3": "'1", "A4": "'2"},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			for ref, st := range tt.styles {
				if err := s.file.Sheet(sheet1).SetStyle(ref, st); err != nil {
					t.Fatal(err)
				}
			}
			mustSet(t, s, sheet1, tt.set...)
			before := dump(t, s, sheet1)
			mustFill(t, s, sheet1, tt.src, tt.dst, tt.series)
			wantInputs(t, s, sheet1, tt.want)
			checkUndoRedo(t, s, sheet1, before)
		})
	}
}

func TestFillCopiesStyles(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	fill := &werkbook.Style{Fill: &werkbook.Fill{Color: "FFFF00"}}
	old := &werkbook.Style{Font: &werkbook.Font{Italic: true}}
	for ref, st := range map[string]*werkbook.Style{"A1": bold, "A2": fill, "A3": old, "A4": old, "A5": old} {
		if err := sh.SetStyle(ref, st); err != nil {
			t.Fatal(err)
		}
	}
	// A2 is a styled empty cell in the source.
	mustSet(t, s, sheet1, "A1", "1", "A3", "old", "A5", "old")
	before := dump(t, s, sheet1)

	mustFill(t, s, sheet1, "A1:A2", "A1:A5", true)
	for ref, want := range map[string]*werkbook.Style{"A3": bold, "A4": fill, "A5": bold} {
		if got, _ := sh.GetStyle(ref); got != want {
			t.Errorf("%s style = %+v, want the pointer of its source cell", ref, got)
		}
	}
	// The styled empty source cell clears the value and gives its style.
	cells := cellsByRef(t, s, sheet1, 1, 1, 5, 1)
	if got := cells["A4"]; got.Type != "empty" || got.Style == nil || got.Style.Fill != "#FFFF00" {
		t.Errorf("A4 = %+v, want an empty cell with the fill", got)
	}
	if got := cells["A5"]; got.Display != "1" || got.Style == nil || !got.Style.Bold {
		t.Errorf("A5 = %+v, want bold 1", got)
	}
	checkUndoRedo(t, s, sheet1, before)
	// Undo gives the old style pointers back.
	mustUndo(t, s)
	if got, _ := sh.GetStyle("A4"); got != old {
		t.Errorf("A4 style after undo = %+v, want the old pointer", got)
	}
}

func TestFillSpill(t *testing.T) {
	t.Run("source follower gives its value", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)*10")
		// B column is filled to the right from the followers A2:A3.
		res := mustFill(t, s, sheet1, "A2:A3", "A2:C3", false)
		if res.Changed != 4 || res.Skipped != 0 {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{
			"A1": "=SEQUENCE(3)*10", "A2": "20", "A3": "30",
			"B2": "20", "C2": "20", "B3": "30", "C3": "30",
		})
		if cells := cellsByRef(t, s, sheet1, 1, 1, 3, 1); !cells["A2"].Spill {
			t.Errorf("the spill is gone: %+v", cells)
		}
	})

	t.Run("destination follower is skipped", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x", "B2", "=SEQUENCE(2,2)")
		before := dump(t, s, sheet1)
		// Fill A1 right over B1; then fill row 1 down over the spill B2:C3.
		mustFill(t, s, sheet1, "A1", "A1:D1", false)
		res := mustFill(t, s, sheet1, "A1:D1", "A1:D3", false)
		// B2 is the anchor (not a follower): it is overwritten, and its
		// followers C2, B3, C3 are written too.
		if res.Skipped != 0 || res.Changed != 8 {
			t.Errorf("result = %+v, want all 8 cells written", res)
		}
		mustUndo(t, s)
		mustUndo(t, s)
		if got := dump(t, s, sheet1); got != before {
			t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
		}

		// Now a fill that covers only followers: C1 down to C3 (C2, C3 are
		// followers of B2, which is not in the destination).
		mustSet(t, s, sheet1, "C1", "y")
		res = mustFill(t, s, sheet1, "C1", "C1:C3", false)
		if res.Skipped != 2 || res.Changed != 0 {
			t.Errorf("result = %+v, want 2 skipped", res)
		}
		cells := cellsByRef(t, s, sheet1, 1, 1, 3, 3)
		if cells["B2"].Display != "1" || cells["C3"].Display != "4" || !cells["C3"].Spill {
			t.Errorf("the spill changed: %+v", cells)
		}
	})

	t.Run("source with an anchor and its followers: the copy spills", func(t *testing.T) {
		// The case of the review: A1 spills to A1:A3; fill A1:A3 to A1:A6.
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)", "A5", "old")
		before := dump(t, s, sheet1)
		res := mustFill(t, s, sheet1, "A1:A3", "A1:A6", true)
		if res.Changed != 3 || res.Skipped != 0 {
			t.Errorf("result = %+v", res)
		}
		cells := cellsByRef(t, s, sheet1, 1, 1, 6, 1)
		if got := cells["A4"]; got.Input != "=SEQUENCE(3)" || got.Display != "1" {
			t.Errorf("A4 = %+v, want the anchor formula with the value 1 (no #SPILL!)", got)
		}
		for ref, want := range map[string]string{"A5": "2", "A6": "3"} {
			if got := cells[ref]; got.Display != want || !got.Spill || got.Formula {
				t.Errorf("%s = %+v, want spill follower %s", ref, got, want)
			}
		}
		if got := cells["A2"]; got.Display != "2" || !got.Spill {
			t.Errorf("the source spill changed: A2 = %+v", got)
		}
		checkUndoRedo(t, s, sheet1, before)
	})

	t.Run("source with a 2D spill: followers in all lanes are absent", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(2,2)")
		mustFill(t, s, sheet1, "A1:B2", "A1:B4", false)
		cells := cellsByRef(t, s, sheet1, 1, 1, 4, 2)
		if cells["A3"].Input != "=SEQUENCE(2,2)" || cells["A3"].Display != "1" || cells["B4"].Display != "4" || !cells["B4"].Spill {
			t.Errorf("cells = %+v, want a second spill at A3:B4", cells)
		}
		// To the right: the anchor copy is at C1.
		mustFill(t, s, sheet1, "A1:B2", "A1:D2", false)
		cells = cellsByRef(t, s, sheet1, 1, 1, 2, 4)
		if cells["C1"].Input != "=SEQUENCE(2,2)" || cells["D2"].Display != "4" || !cells["D2"].Spill {
			t.Errorf("cells = %+v, want a second spill at C1:D2", cells)
		}
	})

	t.Run("followers with no anchor in the source still give values", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)")
		mustFill(t, s, sheet1, "A2:A3", "A2:B3", false)
		wantInputs(t, s, sheet1, map[string]string{"A1": "=SEQUENCE(3)", "A2": "2", "A3": "3", "B2": "2", "B3": "3"})
	})

	t.Run("fill from an anchor copies the formula below its spill", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(2)")
		res := mustFill(t, s, sheet1, "A1", "A1:A4", false)
		// A2 is a follower of the source anchor: read-only.
		if res.Skipped != 1 || res.Changed != 2 {
			t.Errorf("result = %+v, want 1 skipped and 2 written", res)
		}
		// The copy in A4 spills into A5; the copy in A3 is blocked by A4.
		wantInputs(t, s, sheet1, map[string]string{"A1": "=SEQUENCE(2)", "A2": "2", "A3": "=SEQUENCE(2)", "A4": "=SEQUENCE(2)", "A5": "2"})
		if got := cellsByRef(t, s, sheet1, 3, 1, 3, 1)["A3"].Display; got != "#SPILL!" {
			t.Errorf("A3 = %q, want #SPILL!", got)
		}
	})
}

func TestFillErrors(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "B1", "2", "D5", "keep")
	if err := s.file.Sheet(sheet1).MergeCell("C3", "D4"); err != nil {
		t.Fatal(err)
	}
	before := dump(t, s, sheet1)
	history := s.History()

	R := func(a1 string) Range { return rangeOf(t, a1) }
	tests := []struct {
		name     string
		sheet    string
		src, dst Range
		wantErr  string
	}{
		{"unknown sheet", "Nope", R("A1"), R("A1:A3"), "not found"},
		{"src out of range", sheet1, Range{0, 1, 1, 1}, R("A1:A3"), "out of range"},
		{"dst out of range", sheet1, R("A1"), Range{1, 1, werkbook.MaxRows + 1, 1}, "out of range"},
		{"dst does not contain src", sheet1, R("A1:A2"), R("B1:B5"), "must contain"},
		{"dst smaller than src", sheet1, R("A1:A3"), R("A1:A2"), "must contain"},
		{"dst extends on two axes", sheet1, R("A1"), R("A1:C3"), "one direction"},
		{"dst extends up and down", sheet1, R("A3"), R("A1:A5"), "one direction"},
		{"dst extends left and right", sheet1, R("C1"), R("A1:E1"), "one direction"},
		{"source in a merged range", sheet1, R("C3:D3"), R("C3:D8"), "merged"},
		{"destination crosses a merged range", sheet1, R("C1"), R("C1:C6"), "merged"},
		{"destination above the cell limit", sheet1, R("A1:B1"), Range{1, 1, werkbook.MaxRows, 2}, "200000"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := s.Fill(tt.sheet, tt.src, tt.dst, true)
			if err == nil {
				t.Fatalf("Fill gave no error")
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Errorf("error = %q, want it to contain %q", err, tt.wantErr)
			}
			if got := s.History(); got != history {
				t.Errorf("history changed: %+v", got)
			}
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("the workbook changed:\n%s", got)
			}
		})
	}

	t.Run("dst equal to src changes nothing", func(t *testing.T) {
		res, err := s.Fill(sheet1, R("A1:B1"), R("B1:A1"), true)
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 0 || res.Range != R("A1:B1") || res.History != history {
			t.Errorf("result = %+v", res)
		}
	})

	t.Run("exactly the cell limit is valid", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1")
		// 1 source cell + 200,000 destination cells.
		res, err := s.Fill(sheet1, R("A1"), Range{1, 1, maxOpCells + 1, 1}, false)
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != maxOpCells {
			t.Errorf("changed = %d, want %d", res.Changed, maxOpCells)
		}
		mustUndo(t, s)
		if info := s.Sheets()[0]; info.MaxRow != 1 {
			t.Errorf("MaxRow after undo = %d, want 1", info.MaxRow)
		}
	})
}

// Ctrl+D and Ctrl+R are fills with series false. The UI builds the ranges;
// these are the two source rules of cmd/web/PLAN.md section 4.2.
func TestFillCtrlDAndCtrlR(t *testing.T) {
	t.Run("Ctrl+D with two or more rows copies the top row", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "B1", "=A1+1", "A2", "9", "A3", "9")
		mustFill(t, s, sheet1, "A1:B1", "A1:B3", false)
		wantInputs(t, s, sheet1, map[string]string{"A1": "1", "B1": "=A1+1", "A2": "1", "B2": "=A2+1", "A3": "1", "B3": "=A3+1"})
	})
	t.Run("Ctrl+D with one row copies from the row above", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "B1", "=A1+1")
		// Selection A2:B2; src = the cells above, dst = src + selection.
		mustFill(t, s, sheet1, "A1:B1", "A1:B2", false)
		wantInputs(t, s, sheet1, map[string]string{"A1": "1", "B1": "=A1+1", "A2": "1", "B2": "=A2+1"})
	})
	t.Run("Ctrl+R with two or more columns copies the left column", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "A2", "=A1+1")
		mustFill(t, s, sheet1, "A1:A2", "A1:C2", false)
		wantInputs(t, s, sheet1, map[string]string{"A1": "1", "A2": "=A1+1", "B1": "1", "B2": "=B1+1", "C1": "1", "C2": "=C1+1"})
	})
	t.Run("Ctrl+R with one column copies from the column at the left", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "A2", "2")
		mustFill(t, s, sheet1, "A1:A2", "A1:B2", false)
		wantInputs(t, s, sheet1, map[string]string{"A1": "1", "A2": "2", "B1": "1", "B2": "2"})
	})
}
