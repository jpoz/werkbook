package werkbook

import (
	"bytes"
	"testing"
)

func TestOpenReaderAtWithoutFormulasDefersFormulaRegistration(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("A1", 100); err != nil {
		t.Fatalf("SetValue(A1): %v", err)
	}
	if err := s.SetFormula("B1", "A1+50"); err != nil {
		t.Fatalf("SetFormula(B1): %v", err)
	}
	if v, err := s.GetValue("B1"); err != nil {
		t.Fatalf("GetValue(B1): %v", err)
	} else if v.Number != 150 {
		t.Fatalf("GetValue(B1) = %v, want 150", v.Number)
	}

	var buf bytes.Buffer
	n, err := f.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	if n != int64(buf.Len()) {
		t.Fatalf("WriteTo wrote %d bytes, buffer has %d", n, buf.Len())
	}

	f2, err := OpenReaderAt(bytes.NewReader(buf.Bytes()), int64(buf.Len()), WithoutFormulas())
	if err != nil {
		t.Fatalf("OpenReaderAt: %v", err)
	}

	s2 := f2.Sheet("Sheet1")
	if s2 == nil {
		t.Fatal("Sheet1 missing after reopen")
	}
	b1 := s2.rows[1].cells[2]
	if b1 == nil {
		t.Fatal("B1 missing after reopen")
	}
	if b1.compiled != nil {
		t.Fatal("expected WithoutFormulas to skip eager formula compilation")
	}

	deps, err := f2.DirectDependents("Sheet1", "A1")
	if err != nil {
		t.Fatalf("DirectDependents before lazy eval: %v", err)
	}
	if len(deps) != 0 {
		t.Fatalf("DirectDependents before lazy eval = %v, want none", deps)
	}

	if v, err := s2.GetValue("B1"); err != nil {
		t.Fatalf("GetValue(B1) from cache: %v", err)
	} else if v.Number != 150 {
		t.Fatalf("GetValue(B1) from cache = %v, want 150", v.Number)
	}
	if b1.compiled != nil {
		t.Fatal("expected cached formula read to avoid lazy compilation")
	}

	if err := s2.SetValue("A1", 200); err != nil {
		t.Fatalf("SetValue(A1) after reopen: %v", err)
	}
	if v, err := s2.GetValue("B1"); err != nil {
		t.Fatalf("GetValue(B1) after mutation: %v", err)
	} else if v.Number != 250 {
		t.Fatalf("GetValue(B1) after mutation = %v, want 250", v.Number)
	}
	if b1.compiled == nil {
		t.Fatal("expected GetValue to compile formulas lazily after mutation")
	}

	deps, err = f2.DirectDependents("Sheet1", "A1")
	if err != nil {
		t.Fatalf("DirectDependents after lazy eval: %v", err)
	}
	if len(deps) != 1 || deps[0].Sheet != "Sheet1" || deps[0].Col != 2 || deps[0].Row != 1 {
		t.Fatalf("DirectDependents after lazy eval = %v, want [Sheet1!B1]", deps)
	}
}
