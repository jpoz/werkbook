package formula

import (
	"testing"
)

func TestVLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: StringVal("one"),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 2, Row: 2}: StringVal("two"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 3}: StringVal("three"),
		},
	}

	cf := evalCompile(t, "VLOOKUP(2,A1:B3,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "two" {
		t.Errorf("VLOOKUP exact: got %v, want two", got)
	}

	// Not found
	cf = evalCompile(t, "VLOOKUP(5,A1:B3,2,FALSE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("VLOOKUP not found: got %v, want #N/A", got)
	}
}

func TestHLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 3, Row: 2}: StringVal("c"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(2,A1:C2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "b" {
		t.Errorf("HLOOKUP: got %v, want b", got)
	}
}

func TestINDEX(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 2, Row: 2}: NumberVal(40),
		},
	}

	cf := evalCompile(t, "INDEX(A1:B2,2,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 40 {
		t.Errorf("INDEX: got %g, want 40", got.Num)
	}
}

func TestMATCH(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "MATCH(20,A1:A3,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH: got %g, want 2", got.Num)
	}
}

func TestXLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 1, Row: 3}: StringVal("c"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `XLOOKUP("b",A1:A3,B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("XLOOKUP: got %g, want 200", got.Num)
	}

	// Not found with default
	cf = evalCompile(t, `XLOOKUP("z",A1:A3,B1:B3,"missing")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "missing" {
		t.Errorf("XLOOKUP not found: got %v, want missing", got)
	}
}
