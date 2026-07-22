package officeread

import "testing"

func TestXlsxGeneralNumberSuppressesStorageResidue(t *testing.T) {
	styles := xlsxCellStyles{"0", "82"}
	for _, style := range []int{0, 1} {
		if got := xlsxDisplayNumberForCell("42250557.5799999", style, styles); got != "42250557.58" {
			t.Fatalf("style %d display = %q, want %q", style, got, "42250557.58")
		}
	}
}
