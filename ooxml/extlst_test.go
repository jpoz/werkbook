package ooxml

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"strings"
	"testing"
)

const (
	testSlicerExt = `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:slicerCaches/></ext>`
	testExtA      = `<ext uri="{AAAAAAAA-0000-0000-0000-000000000001}"><a/></ext>`
	testExtB      = `<ext uri="{BBBBBBBB-0000-0000-0000-000000000002}"><b/></ext>`
)

var testExts = []extensionXML{
	{URI: "{AAAAAAAA-0000-0000-0000-000000000001}", XML: testExtA},
	{URI: "{BBBBBBBB-0000-0000-0000-000000000002}", XML: testExtB},
}

func TestAddMissingExts(t *testing.T) {
	tests := []struct {
		name string
		in   string
		want string
	}{
		{
			name: "unrelated extension is kept and the two are added",
			in:   `<extLst>` + testSlicerExt + `</extLst>`,
			want: `<extLst>` + testSlicerExt + testExtA + testExtB + `</extLst>`,
		},
		{
			name: "one is present",
			in:   `<extLst>` + testExtB + testSlicerExt + `</extLst>`,
			want: `<extLst>` + testExtB + testSlicerExt + testExtA + `</extLst>`,
		},
		{
			name: "all are present",
			in:   `<extLst>` + testExtA + testExtB + `</extLst>`,
			want: `<extLst>` + testExtA + testExtB + `</extLst>`,
		},
		{
			name: "uri match ignores case",
			in:   `<extLst><ext uri="{aaaaaaaa-0000-0000-0000-000000000001}"><a/></ext></extLst>`,
			want: `<extLst><ext uri="{aaaaaaaa-0000-0000-0000-000000000001}"><a/></ext>` + testExtB + `</extLst>`,
		},
		{
			name: "empty element",
			in:   `<extLst></extLst>`,
			want: `<extLst>` + testExtA + testExtB + `</extLst>`,
		},
		{
			name: "self-closing element",
			in:   `<extLst/>`,
			want: `<extLst>` + testExtA + testExtB + `</extLst>`,
		},
		{
			name: "prefixed self-closing element",
			in:   `<x:extLst />`,
			want: `<x:extLst >` + testExtA + testExtB + `</x:extLst>`,
		},
		{
			name: "prefixed element with white space",
			in:   "<x:extLst>\n  " + testSlicerExt + "\n</x:extLst>",
			want: "<x:extLst>\n  " + testSlicerExt + "\n" + testExtA + testExtB + "</x:extLst>",
		},
		{
			name: "uri of a nested element does not count",
			in:   `<extLst><ext uri="{CCCCCCCC-0000-0000-0000-000000000003}"><ext uri="{AAAAAAAA-0000-0000-0000-000000000001}"/></ext></extLst>`,
			want: `<extLst><ext uri="{CCCCCCCC-0000-0000-0000-000000000003}"><ext uri="{AAAAAAAA-0000-0000-0000-000000000001}"/></ext>` + testExtA + testExtB + `</extLst>`,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := addMissingExts(tt.in, testExts)
			if err != nil {
				t.Fatalf("addMissingExts: %v", err)
			}
			if got != tt.want {
				t.Fatalf("got  %s\nwant %s", got, tt.want)
			}
			// The result is one well-formed element.
			dec := xml.NewDecoder(strings.NewReader(got))
			for {
				if _, err := dec.RawToken(); err == io.EOF {
					break
				} else if err != nil {
					t.Fatalf("result is not well-formed: %v", err)
				}
			}
		})
	}
}

func TestAddMissingExtsRejectsBrokenXML(t *testing.T) {
	for _, in := range []string{``, `<extLst>`, `<extLst><ext></extLst>`, `<extLst/><extLst/>`} {
		if _, err := addMissingExts(in, testExts); err == nil {
			t.Errorf("addMissingExts(%q) gave no error", in)
		}
	}
}

func futureFunctionWorkbook(extras []RawElement) *WorkbookData {
	return &WorkbookData{
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{
				{Ref: "A1", Formula: "_xlfn.LET(_xlpm.x,5,_xlpm.x+1)", Value: "6"},
			}}},
		}},
		ExtraElements: extras,
	}
}

func writtenWorkbookXML(t *testing.T, data *WorkbookData) string {
	t.Helper()
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	return readZipEntry(t, zr, "xl/workbook.xml")
}

// A workbook that has an unrelated extension and gets a LET formula must get
// the future-function extensions. Without them Excel shows a repair prompt.
func TestWriteWorkbookAddsFutureFunctionExtsToExistingExtLst(t *testing.T) {
	existing := `<extLst>` + testSlicerExt + `</extLst>`
	workbookXML := writtenWorkbookXML(t, futureFunctionWorkbook([]RawElement{
		{Name: "extLst", XML: existing, OrderKey: workbookElementOrder["extLst"], Seq: 1},
	}))

	if n := strings.Count(workbookXML, "<extLst>"); n != 1 {
		t.Fatalf("workbook.xml has %d extLst elements, want 1: %s", n, workbookXML)
	}
	for _, want := range []string{
		testSlicerExt,
		futureFunctionsWorkbookPrExtXML,
		futureFunctionsCalcFeaturesExtXML,
	} {
		if !strings.Contains(workbookXML, want) {
			t.Errorf("workbook.xml does not contain %s\nxml: %s", want, workbookXML)
		}
	}
	if !strings.HasSuffix(workbookXML, "</extLst></workbook>") {
		t.Errorf("extLst is not the last child of workbook: %s", workbookXML)
	}
}

// Extensions that the source workbook has are not written two times.
func TestWriteWorkbookKeepsPresentFutureFunctionExts(t *testing.T) {
	existing := `<extLst>` + futureFunctionsCalcFeaturesExtXML + `</extLst>`
	workbookXML := writtenWorkbookXML(t, futureFunctionWorkbook([]RawElement{
		{Name: "extLst", XML: existing, OrderKey: workbookElementOrder["extLst"], Seq: 1},
	}))
	if n := strings.Count(workbookXML, futureFunctionsCalcFeaturesExtURI); n != 1 {
		t.Errorf("calcFeatures extension is written %d times, want 1", n)
	}
	if n := strings.Count(workbookXML, futureFunctionsWorkbookPrExtURI); n != 1 {
		t.Errorf("workbookPr extension is written %d times, want 1", n)
	}
}

// A workbook with no future function keeps its extLst byte for byte.
func TestWriteWorkbookLeavesExtLstWithoutFutureFunctions(t *testing.T) {
	existing := `<extLst>` + testSlicerExt + `</extLst>`
	data := futureFunctionWorkbook([]RawElement{
		{Name: "extLst", XML: existing, OrderKey: workbookElementOrder["extLst"], Seq: 1},
	})
	data.Sheets[0].Rows[0].Cells[0].Formula = "SUM(1,2)"
	workbookXML := writtenWorkbookXML(t, data)
	if !strings.Contains(workbookXML, existing) {
		t.Errorf("extLst changed: %s", workbookXML)
	}
	if strings.Contains(workbookXML, futureFunctionsCalcFeaturesExtURI) {
		t.Errorf("future-function extension was added with no future function: %s", workbookXML)
	}
}
