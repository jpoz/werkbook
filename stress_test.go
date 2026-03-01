package werkbook_test

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func TestStressFormulas(t *testing.T) {
	stressDir := "testdata/stress"

	// Check if the stress directory exists
	if _, err := os.Stat(stressDir); os.IsNotExist(err) {
		t.Skip("testdata/stress/ not found; run 'go run ./cmd/gentestdata' to generate")
	}

	// Discover all .spec.json files
	entries, err := os.ReadDir(stressDir)
	if err != nil {
		t.Fatalf("read stress dir: %v", err)
	}

	found := false
	for _, entry := range entries {
		name := entry.Name()
		if !strings.HasSuffix(name, ".spec.json") {
			continue
		}

		base := strings.TrimSuffix(name, ".spec.json")
		expectedPath := filepath.Join(stressDir, base+".expected.json")

		// Skip if no expected file
		if _, err := os.Stat(expectedPath); os.IsNotExist(err) {
			continue
		}

		found = true
		t.Run(base, func(t *testing.T) {
			// Load spec
			specPath := filepath.Join(stressDir, name)
			spec, err := fuzz.LoadSpec(specPath, fuzz.LocalExcelExcludedFunctions)
			if err != nil {
				t.Fatalf("load spec: %v", err)
			}

			// Load expected results
			expectedData, err := os.ReadFile(expectedPath)
			if err != nil {
				t.Fatalf("read expected: %v", err)
			}
			var expected []fuzz.CellResult
			if err := json.Unmarshal(expectedData, &expected); err != nil {
				t.Fatalf("parse expected: %v", err)
			}

			// Build XLSX and evaluate with werkbook
			tmpDir := t.TempDir()
			_, wbResults, err := fuzz.BuildXLSX(spec, tmpDir)
			if err != nil {
				t.Fatalf("build xlsx: %v", err)
			}

			// Compare
			mismatches := fuzz.CompareResults(spec.Checks, wbResults, expected, spec)
			if len(mismatches) > 0 {
				t.Errorf("%d mismatches in %s:\n%s", len(mismatches), base,
					fuzz.FormatMismatches(mismatches, "excel"))
			}
		})
	}

	if !found {
		t.Skip("no .expected.json files found; run 'go run ./cmd/gentestdata -bootstrap' to generate")
	}
}
