package formula

import "testing"

func TestEvalAVERAGEIFFullColumnIgnoresLogicalBlankTail(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("A"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("E"),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: StringVal("B"),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, `AVERAGEIF(A:A,"<>E",B:B)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf(`Eval(AVERAGEIF(A:A,"<>E",B:B)): %v`, err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Fatalf(`AVERAGEIF(A:A,"<>E",B:B) = %#v, want 20`, got)
	}
}
