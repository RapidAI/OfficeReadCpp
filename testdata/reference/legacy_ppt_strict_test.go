package officeread

import (
	"bytes"
	"path/filepath"
	"strings"
	"testing"
)

func TestStrictLegacyPPTWebSampleExcludesCarvedInternalImages(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000008.ppt")
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(compatible.Images) != 2 {
		t.Fatalf("compatibility PPT images = %d, want 2", len(compatible.Images))
	}
	strict, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(strict.Images) != 0 {
		t.Fatalf("strict PPT images = %#v, want none", strict.Images)
	}
}

func TestStrictLegacyPPTUsesVisiblePictureShapeOccurrences(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// PowerPoint's Slide.Shapes exposes 24 msoPicture instances. Several
	// instances deliberately share a source blip, so content de-duplication
	// would be incorrect here.
	if len(result.Images) != 24 {
		t.Fatalf("strict PPT visible picture occurrences = %d, want 24", len(result.Images))
	}
}

func TestStrictLegacyPPTUsesActiveDocumentAndSlidesOnly(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt")
	result, err := Extract(file, Options{StrictOfficeContent: true, StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"A Tiered Approach for the Identification of a Human Fecal Pollution Source",

		"Compare ENT levels at each station measured during",
	} {
		if !strings.Contains(result.Text, want) {
			t.Fatalf("missing active PPT content %q in %q", want, result.Text)
		}
	}
	for _, unwanted := range []string{"Santa Monica (I)", "Log mean ENT", "Frequency of data"} {
		if strings.Contains(result.Text, unwanted) {
			t.Fatalf("kept chart-axis text %q in %q", unwanted, result.Text)
		}
	}
}

func TestLegacyPPTDropsUnrenderedOutlinePlaceholders(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "009847.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, placeholder := range []string{"Second level", "Third level", "Fourth level", "Fifth level"} {
		if strings.Contains(result.Text, placeholder) {
			t.Fatalf("kept unrendered outline placeholder %q in %q", placeholder, result.Text)
		}
	}
}

func TestLegacyPPTDropsReplacedHiddenShapeText(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "ppt", "004607.ppt")
	result, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(result.Text, "and slow spill (septa)\nBeam pipe at F17 lowered") {
		t.Fatalf("kept hidden text runs in %q", result.Text)
	}
	if !strings.Contains(result.Text, "and slow spill (septa) to SY absorber") || !strings.Contains(result.Text, "Beam pipe at F17 c-magnet raised") {
		t.Fatalf("lost rendered replacement text in %q", result.Text)
	}
}

func TestLegacyPPTTextKeepsVisibleSlideDuplicatesAndExcludesNotes(t *testing.T) {
	var slide bytes.Buffer
	slide.Write(pptRecord(0x0fa8, []byte("Repeated visible title")))
	var deck bytes.Buffer
	deck.Write(pptContainerRecord(0x03ee, slide.Bytes()))
	deck.Write(pptContainerRecord(0x03ee, slide.Bytes()))
	deck.Write(pptContainerRecord(0x03f0, pptRecord(0x0fa8, []byte("Speaker note only"))))
	text := strings.Join(pptRecordText(deck.Bytes()), "\n")
	if strings.Count(text, "Repeated visible title") != 2 {
		t.Fatalf("visible slide text occurrences = %q", text)
	}
	if strings.Contains(text, "Speaker note only") {
		t.Fatalf("notes text leaked into slide text: %q", text)
	}
}
