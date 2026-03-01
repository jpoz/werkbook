package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func allScenarios() []ScenarioSet {
	return []ScenarioSet{
		{Name: "math", Cases: mathScenarios()},
		{Name: "stat", Cases: statScenarios()},
		{Name: "text", Cases: textScenarios()},
		{Name: "date", Cases: dateScenarios()},
		{Name: "lookup", Cases: lookupScenarios()},
		{Name: "logic", Cases: logicScenarios()},
		{Name: "info", Cases: infoScenarios()},
		{Name: "arrays", Cases: arrayScenarios()},
		{Name: "nested", Cases: nestedScenarios()},
		crosssheetScenarios(), // returns a ScenarioSet directly with Sheets
	}
}

func main() {
	outdir := flag.String("outdir", "testdata/stress", "Output directory")
	bootstrap := flag.Bool("bootstrap", false, "Run Excel to capture ground truth")
	oracle := flag.String("oracle", "local-excel", "Oracle: local-excel, libreoffice, excel")
	only := flag.String("only", "", "Only process this scenario (e.g. -only arrays)")
	flag.Parse()

	if err := os.MkdirAll(*outdir, 0755); err != nil {
		log.Fatalf("create output dir: %v", err)
	}

	scenarios := allScenarios()

	for _, ss := range scenarios {
		if *only != "" && ss.Name != *only {
			continue
		}
		spec := buildSpec(ss)

		// Save spec.json
		specPath := filepath.Join(*outdir, ss.Name+".spec.json")
		if err := fuzz.SaveSpec(specPath, spec); err != nil {
			log.Fatalf("save spec %s: %v", ss.Name, err)
		}
		fmt.Printf("Generated %s\n", specPath)

		// Build XLSX + evaluate with werkbook
		xlsxPath, _, err := fuzz.BuildXLSX(spec, *outdir)
		if err != nil {
			log.Fatalf("build xlsx %s: %v", ss.Name, err)
		}
		fmt.Printf("Generated %s\n", xlsxPath)

		// Bootstrap: run oracle to get ground truth
		if *bootstrap {
			eval, err := fuzz.NewEvaluator(*oracle)
			if err != nil {
				log.Fatalf("create evaluator: %v", err)
			}

			results, err := eval.Eval(xlsxPath, spec.Checks)
			if err != nil {
				fmt.Fprintf(os.Stderr, "WARNING: evaluate %s failed: %v (skipping)\n", ss.Name, err)
				continue
			}

			expectedPath := filepath.Join(*outdir, ss.Name+".expected.json")
			data, err := json.MarshalIndent(results, "", "  ")
			if err != nil {
				log.Fatalf("marshal expected %s: %v", ss.Name, err)
			}
			if err := os.WriteFile(expectedPath, data, 0644); err != nil {
				log.Fatalf("write expected %s: %v", ss.Name, err)
			}
			fmt.Printf("Bootstrapped %s\n", expectedPath)
		}
	}

	coverageReport(scenarios)
}
