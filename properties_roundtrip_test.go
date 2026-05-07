package werkbook_test

import (
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/jpoz/werkbook"
)

func TestCorePropertiesRoundTrip(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	if _, err := f.NewSheet("Summary"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	created := time.Date(2026, 3, 6, 12, 34, 56, 0, time.FixedZone("MST", -7*60*60))
	modified := created.Add(95 * time.Minute)
	f.SetCoreProperties(werkbook.CoreProperties{
		Title:          "Budget",
		Subject:        "Q1",
		Creator:        "Alice",
		Description:    "Ops workbook",
		Identifier:     "urn:budget:q1",
		Language:       "en-US",
		Keywords:       "budget,finance",
		Category:       "Planning",
		ContentStatus:  "Draft",
		Version:        "1.0",
		Revision:       "3",
		LastModifiedBy: "Bob",
		Created:        created,
		Modified:       modified,
	})

	dir := t.TempDir()
	path := filepath.Join(dir, "core-props.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	rootRels := string(readSheetXML(t, path, "_rels/.rels"))
	for _, want := range []string{
		`Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"`,
		`Target="docProps/core.xml"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"`,
		`Target="docProps/app.xml"`,
	} {
		if !strings.Contains(rootRels, want) {
			t.Fatalf("expected root relationships to contain %s: %s", want, rootRels)
		}
	}

	contentTypes := string(readSheetXML(t, path, "[Content_Types].xml"))
	for _, want := range []string{
		`PartName="/docProps/core.xml"`,
		`application/vnd.openxmlformats-package.core-properties+xml`,
		`PartName="/docProps/app.xml"`,
		`application/vnd.openxmlformats-officedocument.extended-properties+xml`,
	} {
		if !strings.Contains(contentTypes, want) {
			t.Fatalf("expected content types to contain %s: %s", want, contentTypes)
		}
	}

	coreXML := string(readSheetXML(t, path, "docProps/core.xml"))
	for _, want := range []string{
		`<dc:title>Budget</dc:title>`,
		`<dc:creator>Alice</dc:creator>`,
		`<cp:lastModifiedBy>Bob</cp:lastModifiedBy>`,
		`xsi:type="dcterms:W3CDTF"`,
		created.UTC().Format(time.RFC3339),
		modified.UTC().Format(time.RFC3339),
	} {
		if !strings.Contains(coreXML, want) {
			t.Fatalf("expected core.xml to contain %s: %s", want, coreXML)
		}
	}

	appXML := string(readSheetXML(t, path, "docProps/app.xml"))
	for _, want := range []string{
		`<Application>Werkbook</Application>`,
		`<vt:lpstr>Worksheets</vt:lpstr>`,
		`<vt:i4>2</vt:i4>`,
		`<vt:lpstr>Data</vt:lpstr>`,
		`<vt:lpstr>Summary</vt:lpstr>`,
	} {
		if !strings.Contains(appXML, want) {
			t.Fatalf("expected app.xml to contain %s: %s", want, appXML)
		}
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	got := f2.CoreProperties()
	if got.Title != "Budget" || got.Subject != "Q1" || got.Creator != "Alice" || got.Description != "Ops workbook" ||
		got.Identifier != "urn:budget:q1" || got.Language != "en-US" || got.Keywords != "budget,finance" ||
		got.Category != "Planning" || got.ContentStatus != "Draft" || got.Version != "1.0" ||
		got.Revision != "3" || got.LastModifiedBy != "Bob" {
		t.Fatalf("CoreProperties = %#v", got)
	}
	if !got.Created.Equal(created) {
		t.Fatalf("Created = %v, want %v", got.Created, created)
	}
	if !got.Modified.Equal(modified) {
		t.Fatalf("Modified = %v, want %v", got.Modified, modified)
	}
}

