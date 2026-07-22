package officeread

import (
	"strings"
	"testing"
)

func TestVisibleXMLTextSkipsNegativeZIndexVMLTextBox(t *testing.T) {
	xml := `<w:document xmlns:w="urn:x" xmlns:v="urn:v"><w:body><w:p><w:r><w:t>Visible text</w:t></w:r></w:p><w:pict><v:shape style="position:absolute;z-index:-251658752;visibility:visible"><w:txbxContent><w:p><w:r><w:t>Behind text</w:t></w:r></w:p></w:txbxContent></v:shape></w:pict></w:body></w:document>`
	got, err := visibleXMLText([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Visible text") || strings.Contains(got, "Behind text") {
		t.Fatalf("visibleXMLText() = %q", got)
	}
}

func TestVisibleXMLTextSkipsVMLTextBoxWithBehindDocumentZIndex(t *testing.T) {
	xml := `<w:document xmlns:w="urn:x" xmlns:v="urn:v"><w:body><w:p><w:r><w:t>Visible text</w:t></w:r></w:p><w:pict><v:shape style="position:absolute;z-index:251658752"><w:txbxContent><w:p><w:r><w:t>Behind document text</w:t></w:r></w:p></w:txbxContent></v:shape></w:pict></w:body></w:document>`
	got, err := visibleXMLText([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Visible text") || strings.Contains(got, "Behind document text") {
		t.Fatalf("visibleXMLText() = %q", got)
	}
}
