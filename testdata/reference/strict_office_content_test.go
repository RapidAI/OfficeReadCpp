package officeread

import (
	"path/filepath"
	"strings"
	"testing"
)

func TestStrictOfficeContentExcludesDOCXChartCache(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "LibreOffice__core__chart2_qa_extras_data_docx_data_point_inherited_color.docx")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(compatible.Text, "Column 1") {
		t.Fatalf("compatibility text lost chart cache: %q", compatible.Text)
	}
	strict, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strict.Text != "" {
		t.Fatalf("strict document content = %q, want empty", strict.Text)
	}
}

func TestStrictOfficeContentExcludesEmbeddedChartWorkbook(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "docx", "LibreOffice__core__chart2_qa_extras_data_docx_3d-bar-label.docx")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(compatible.Text, "Series 3") {
		t.Fatalf("compatibility text lost embedded chart workbook: %q", compatible.Text)
	}
	strict, err := Extract(file, Options{StrictOfficeContent: true})
	if err != nil {
		t.Fatal(err)
	}
	if strict.Text != "" {
		t.Fatalf("strict document content = %q, want empty", strict.Text)
	}
}
