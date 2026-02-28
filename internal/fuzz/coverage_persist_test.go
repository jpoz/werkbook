package fuzz

import (
	"os"
	"path/filepath"
	"testing"
)

func TestSaveAndLoadCoverage(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "coverage.json")

	fc := NewFunctionCoverage()
	fc.Tested["SUM"] = &FunctionStats{Passed: 10, Failed: 2}
	fc.Tested["IF"] = &FunctionStats{Passed: 5, Failed: 0}
	fc.Broken["WORKDAY"] = 2

	if err := SaveCoverage(path, fc); err != nil {
		t.Fatalf("SaveCoverage: %v", err)
	}

	loaded, err := LoadCoverage(path)
	if err != nil {
		t.Fatalf("LoadCoverage: %v", err)
	}
	if loaded == nil {
		t.Fatal("LoadCoverage returned nil")
	}

	if loaded.Tested["SUM"].Passed != 10 || loaded.Tested["SUM"].Failed != 2 {
		t.Errorf("SUM: got %+v", loaded.Tested["SUM"])
	}
	if loaded.Tested["IF"].Passed != 5 {
		t.Errorf("IF: got %+v", loaded.Tested["IF"])
	}
	if loaded.Broken["WORKDAY"] != 2 {
		t.Errorf("Broken WORKDAY: got %d", loaded.Broken["WORKDAY"])
	}
}

func TestLoadCoverageNotFound(t *testing.T) {
	fc, err := LoadCoverage("/nonexistent/path/coverage.json")
	if err != nil {
		t.Fatalf("expected nil error, got: %v", err)
	}
	if fc != nil {
		t.Fatalf("expected nil coverage, got: %+v", fc)
	}
}

func TestMergeCoverage(t *testing.T) {
	dst := NewFunctionCoverage()
	dst.Tested["SUM"] = &FunctionStats{Passed: 5, Failed: 1}

	src := NewFunctionCoverage()
	src.Tested["SUM"] = &FunctionStats{Passed: 10, Failed: 2}
	src.Tested["IF"] = &FunctionStats{Passed: 3, Failed: 0}
	src.Broken["MOD"] = 3

	MergeCoverage(dst, src)

	if dst.Tested["SUM"].Passed != 15 || dst.Tested["SUM"].Failed != 3 {
		t.Errorf("SUM: got %+v, want Passed=15 Failed=3", dst.Tested["SUM"])
	}
	if dst.Tested["IF"].Passed != 3 {
		t.Errorf("IF: got %+v", dst.Tested["IF"])
	}
	if dst.Broken["MOD"] != 3 {
		t.Errorf("Broken MOD: got %d", dst.Broken["MOD"])
	}
}

func TestSaveCoverageAtomic(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "coverage.json")

	fc := NewFunctionCoverage()
	fc.Tested["ABS"] = &FunctionStats{Passed: 1}

	if err := SaveCoverage(path, fc); err != nil {
		t.Fatalf("SaveCoverage: %v", err)
	}

	// Verify no .tmp file left behind.
	if _, err := os.Stat(path + ".tmp"); !os.IsNotExist(err) {
		t.Error("temp file was not cleaned up")
	}
}
