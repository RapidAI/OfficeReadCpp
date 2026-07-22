package officeread

import "testing"

func TestCleanTextRepairsLegacyWordApostropheMojibake(t *testing.T) {
	if got := cleanText("Commission\ubb69 policy"); got != "Commission's policy" {
		t.Fatalf("cleanText() = %q", got)
	}
}
