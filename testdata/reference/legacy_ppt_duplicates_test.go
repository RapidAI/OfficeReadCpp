package officeread

import (
	"bytes"
	"strings"
	"testing"
)

func TestLegacyPPTPreservesVisibleTextOccurrences(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptContainerRecord(0x03ee, pptRecord(0x0fa8, []byte("Repeated title"))))
	ppt.Write(pptContainerRecord(0x03ee, pptRecord(0x0fa8, []byte("Repeated title"))))
	parts := extractLegacyTextWithMetadata("sample.ppt", nil, []oleStream{{Name: "PowerPoint Document", Data: ppt.Bytes()}}, false)
	if got := strings.Count(strings.Join(parts, "\n"), "Repeated title"); got != 2 {
		t.Fatalf("visible PPT title occurrences = %d, want 2: %#v", got, parts)
	}
}
