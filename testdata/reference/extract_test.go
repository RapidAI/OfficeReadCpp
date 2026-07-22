package officeread

import (
	"archive/zip"
	"bytes"
	"compress/gzip"
	"encoding/base64"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"hash/crc32"
	"image"
	"image/color"
	"image/gif"
	"image/jpeg"
	"math"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"unicode/utf16"
	"unicode/utf8"
)

func TestExtractDownloadedSamples(t *testing.T) {
	samples, err := filepath.Glob(filepath.Join("testdata", "samples", "*.*"))
	if err != nil {
		t.Fatal(err)
	}
	if len(samples) < 6 {
		t.Fatalf("expected at least 6 Office samples, got %d", len(samples))
	}
	seenExt := map[string]bool{}
	allowEmpty := map[string]bool{
		"51921-Word-Crash067.doc":      true, // no visible body text after metadata is excluded
		"51921-Word-Crash067.docx":     true, // no visible body text after metadata is excluded
		"57603-seven_columns.doc":      true, // no visible body text after metadata is excluded
		"58616.xlsx":                   true, // encrypted OOXML-in-OLE package without extractable plaintext
		"59378.docx":                   true, // no visible body text after metadata is excluded
		"generated-docx-glossary.docx": true, // glossary/building-block content is not visible document text
		"missing-blip-fill.pptx":       true, // blank slide with only notes page system placeholders
		"60042.pptx":                   true, // no visible slide text or content images after package thumbnail is excluded
		"61515.pptx":                   true, // no visible slide text or content images after package thumbnail is excluded
		"63200.pptx":                   true, // no visible slide text or content images after package thumbnail is excluded
		"SmartArt.pptx":                true, // no visible slide text or content images after package thumbnail is excluded
		"Bug51944.doc":                 true, // no visible body text after metadata is excluded
		"Bug62859.docx":                true, // glossary/building-block content is not visible document text
		"42844.xls":                    true, // only printer driver settings remain after non-visible workbook internals are excluded
		"OPCCompliance_CoreProperties_OnlyOneCorePropertiesPart.docx": true, // core properties only
		"OPCCompliance_CoreProperties_SUCCESS.docx":                   true, // core properties only
		"TestPackageCoreProperiesGetters.docx":                        true, // core properties only
		"TestPackageCoreProperiesSetters.docx":                        true, // core properties only
		"TestTableColumns.docx":                                       true, // no visible body text after metadata is excluded
		"WithArtShapes.doc":                                           true, // no visible body text after metadata is excluded
		"empty.doc":                                                   true, // intentionally empty fixture
		"equation.doc":                                                true, // only embedded equation OLE control text is visible to the extractor
		"protected_passtika.xlsx":                                     true, // encrypted OOXML-in-OLE package without extractable plaintext
		"stress025.docx":                                              true, // package stress fixture without document content or metadata text parts
		"table_test.pptx":                                             true, // no visible slide text or content images after package thumbnail is excluded
		"Tika-792.docx":                                               true, // only deleted/move-from revision text remains after hidden revisions are excluded
		"vector_image.doc":                                            true, // no supported visible text or image payload
	}
	for _, sample := range samples {
		ext := strings.ToLower(filepath.Ext(sample))
		if !map[string]bool{".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true}[ext] {
			continue
		}
		res, err := Extract(sample, Options{})
		if err != nil {
			t.Fatalf("%s: %v", sample, err)
		}
		seenExt[ext] = true
		if strings.TrimSpace(res.Text) == "" && len(res.Images) == 0 && !allowEmpty[filepath.Base(sample)] {
			t.Fatalf("%s: extracted no text or images", sample)
		}
		if strings.ContainsRune(res.Text, utf8.RuneError) {
			t.Fatalf("%s: extracted text contains replacement rune in %.400q", sample, res.Text)
		}
		if bad := firstDisallowedSampleTextControl(res.Text); bad != 0 {
			t.Fatalf("%s: extracted text contains raw control U+%04X in %.400q", sample, bad, res.Text)
		}
		for _, bad := range historicalMojibakeFragments() {
			if strings.Contains(res.Text, bad) {
				t.Fatalf("%s: extracted text contains historical mojibake fragment %q in %.400q", sample, bad, res.Text)
			}
		}
	}
	for _, ext := range []string{".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx"} {
		if !seenExt[ext] {
			t.Fatalf("missing sample extension %s", ext)
		}
	}
}

func TestSampleImagesAreValidAndReferenced(t *testing.T) {
	samples, err := filepath.Glob(filepath.Join("testdata", "samples", "*.*"))
	if err != nil {
		t.Fatal(err)
	}
	imageDir := t.TempDir()
	checkedImages := 0
	checkedDocs := 0
	for i, sample := range samples {
		ext := strings.ToLower(filepath.Ext(sample))
		if !map[string]bool{".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true}[ext] {
			continue
		}
		dir := filepath.Join(imageDir, fmt.Sprintf("sample-%04d", i))
		res, err := Extract(sample, Options{ImageDir: dir})
		if err != nil {
			t.Fatalf("%s: %v", sample, err)
		}
		images := validOutputImages(res.Images)
		if len(res.Images) != len(images) {
			t.Fatalf("%s: result kept invalid image entries: got %d valid %d", sample, len(res.Images), len(images))
		}
		if len(images) == 0 {
			continue
		}
		checkedDocs++
		names := imageOutputFilenames(images)
		md := res.Markdown("images")
		for j, img := range images {
			name := names[j]
			if name == "" {
				t.Fatalf("%s: empty image output name at index %d", sample, j)
			}
			if strings.ContainsAny(name, `\/:*?"<>|`) {
				t.Fatalf("%s: unsafe image output name %q", sample, name)
			}
			if !validImageData(img.Ext, img.Data) {
				t.Fatalf("%s: invalid image data %s at index %d (%d bytes)", sample, img.Ext, j, len(img.Data))
			}
			written, err := os.ReadFile(filepath.Join(dir, name))
			if err != nil {
				t.Fatalf("%s: image %q was not written: %v", sample, name, err)
			}
			if !bytes.Equal(written, img.Data) {
				t.Fatalf("%s: written image %q differs from result image data", sample, name)
			}
			if !strings.Contains(md, "images/"+name) {
				t.Fatalf("%s: markdown does not reference image %q in:\n%s", sample, name, md)
			}
			checkedImages++
		}
	}
	if checkedDocs == 0 || checkedImages == 0 {
		t.Fatalf("expected image-bearing samples, got docs=%d images=%d", checkedDocs, checkedImages)
	}
}

func firstDisallowedSampleTextControl(s string) rune {
	for _, r := range s {
		if r < 0x20 && r != '\n' && r != '\t' {
			return r
		}
	}
	return 0
}

func historicalMojibakeFragments() []string {
	return []string{
		"\uff8e\u78c5\u890c", // Word95 CP1251 short text misread as Shift-JIS.
		"\u00ec\u00eb",
		"\u00f0\u00f3\u00e1",
		"\u8106\u8107",
		"\u8320",
		"\u9225?",
	}
}

func TestSampleMarkdownQuality(t *testing.T) {
	samples := []string{
		"word95err.doc",
		"testPictures.doc",
		"VariousPictures.docx",
		"pictures.ppt",
		"54880_chinese.ppt",
		"WithComments.ppt",
		"generated-pptx-rich-parts.pptx",
		"SimpleWithImages.xls",
		"TestUnicode.xls",
		"picture.xlsx",
	}
	seenExt := map[string]bool{}
	for _, name := range samples {
		t.Run(name, func(t *testing.T) {
			sample := filepath.Join("testdata", "samples", name)
			res, err := Extract(sample, Options{})
			if err != nil {
				t.Fatal(err)
			}
			seenExt[strings.ToLower(filepath.Ext(name))] = true
			md := res.Markdown("images")
			if strings.TrimSpace(md) == "" {
				t.Fatalf("markdown is empty for representative sample")
			}
			if strings.ContainsRune(md, utf8.RuneError) {
				t.Fatalf("markdown contains replacement rune in %.400q", md)
			}
			if bad := firstDisallowedSampleTextControl(md); bad != 0 {
				t.Fatalf("markdown contains raw control U+%04X in %.400q", bad, md)
			}
			for _, bad := range historicalMojibakeFragments() {
				if strings.Contains(md, bad) {
					t.Fatalf("markdown contains historical mojibake fragment %q in %.400q", bad, md)
				}
			}
			for _, hidden := range sampleMarkdownInternalFragments() {
				if strings.Contains(md, hidden) {
					t.Fatalf("markdown leaked internal Office fragment %q in %.400q", hidden, md)
				}
			}
		})
	}
	for _, ext := range []string{".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx"} {
		if !seenExt[ext] {
			t.Fatalf("missing representative markdown sample extension %s", ext)
		}
	}
}

func sampleMarkdownInternalFragments() []string {
	return []string{
		"[Content_Types].xml",
		"customXml/",
		"docProps/",
		"ppt/_rels/",
		"ppt/media/",
		"ppt/slides/_rels/",
		"word/_rels/",
		"word/media/",
		"xl/_rels/",
		"xl/media/",
		"xl/worksheets/_rels/",
		"application/vnd.openxmlformats",
		"Relationship Id=",
		"TargetMode=",
		"r:embed=",
	}
}

func TestComplexSampleExpectations(t *testing.T) {
	tests := []struct {
		name       string
		wantText   string
		minImages  int
		allowEmpty bool
	}{
		{name: "VariousPictures.docx", minImages: 5},
		{name: "header_image.doc", minImages: 1},
		{name: "pictures.ppt", minImages: 1},
		{name: "chart-picture-bg.pptx", minImages: 1},
		{name: "picture.xlsx", minImages: 1},
		{name: "WithDrawing.xlsx", minImages: 1},
		{name: "HeaderFooterComplexFormats.xlsx", wantText: "Header"},
		{name: "comments.xlsx", wantText: "comment"},
		{name: "ExcelTables.xlsx", wantText: "AgeGroup"},
		{name: "table_test2.pptx", wantText: "Header"},
		{name: "55406_Conditional_formatting_sample.xlsx", wantText: "Summary"},
		{name: "WithHyperlink.xls", wantText: "http://poi.apache.org/"},
		{name: "FieldCodes.docx", wantText: "16 June 2010"},
		{name: "InlineStrings.xlsx", wantText: "Inline String"},
		{name: "54084 - Greek - beyond BMP.xlsx", wantText: "𝝊"},
		{name: "smartart-rotated-text.pptx", wantText: "abc"},
		{name: "WithGIF.docx", minImages: 1},
		{name: "testPictures.doc", minImages: 3},
		{name: "word_with_embeded_ooxml.doc", minImages: 10},
		{name: "generated-altchunk-html.docx", wantText: "Generated AltChunk Heading"},
		{name: "generated-altchunk-html.docx", wantText: "Visible & decoded HTML text"},
		{name: "generated-docx-alt-attributes.docx", wantText: "Generated DOCX Picture Description Attribute"},
		{name: "generated-docx-alt-attributes.docx", wantText: "Generated DOCX Picture Title Attribute"},
		{name: "generated-docx-embedded-ooxml.docx", wantText: "Generated Embedded DOCX Inner Text", minImages: 1},
		{name: "generated-docx-rich-parts.docx", wantText: "Generated Footnote Text", minImages: 1},
		{name: "generated-docx-rich-parts.docx", wantText: "Generated DOCX Chart Text"},
		{name: "generated-docx-smartart.docx", wantText: "Generated DOCX SmartArt Text"},
		{name: "generated-docx-thumbnail.docx", wantText: "Generated DOCX Thumbnail Text"},
		{name: "generated-docx-vml.docx", wantText: "Generated DOCX VML Text"},
		{name: "generated-pptx-alt-attributes.pptx", wantText: "Generated PPTX Shape Description Attribute"},
		{name: "generated-pptx-alt-attributes.pptx", wantText: "Generated PPTX Hyperlink Tooltip Attribute"},
		{name: "generated-pptx-chart-comment.pptx", wantText: "Generated PPTX XML Comment Text"},
		{name: "generated-pptx-chart-comment.pptx", wantText: "Generated PPTX Chart XML Text"},
		{name: "generated-pptx-embedded-legacy-ole.pptx", wantText: "URL is http://poi.apache.org/"},
		{name: "generated-pptx-embedded-ooxml.pptx", wantText: "Generated Embedded XLSX Inner Text"},
		{name: "generated-pptx-rich-parts.pptx", wantText: "Generated PPTX Notes Text", minImages: 1},
		{name: "generated-pptx-rich-parts.pptx", wantText: "Generated SmartArt Data Text"},
		{name: "generated-pptx-thumbnail.pptx", wantText: "Generated PPTX Thumbnail Text"},
		{name: "generated-xlsx-alt-attributes.xlsx", wantText: "Generated XLSX Drawing Description Attribute"},
		{name: "generated-xlsx-alt-attributes.xlsx", wantText: "Generated XLSX Drawing Title Attribute"},
		{name: "generated-xlsx-sheetnames.xlsx", wantText: "Generated Visible Sheet Name"},
		{name: "generated-xlsx-threaded-comments.xlsx", wantText: "Generated XLSX Threaded Comment Text"},
		{name: "generated-xlsx-vml-drawing.xlsx", wantText: "Generated XLSX VML Drawing Text"},
		{name: "generated-xlsx-embedded-legacy-ole.xlsx", wantText: "With a comment on it"},
		{name: "generated-xlsx-embedded-ooxml.xlsx", wantText: "Generated Embedded PPTX Inner Text"},
		{name: "generated-xlsx-pivot-cache.xlsx", wantText: "GeneratedPivotTableName"},
		{name: "generated-xlsx-pivot-cache.xlsx", wantText: "Generated Pivot Field Caption"},
		{name: "generated-xlsx-pivot-cache.xlsx", wantText: "Generated Pivot Cache Field"},
		{name: "generated-xlsx-pivot-cache.xlsx", wantText: "Generated Pivot Region North"},
		{name: "generated-xlsx-slicer-cache.xlsx", wantText: "GeneratedSlicerName"},
		{name: "generated-xlsx-slicer-cache.xlsx", wantText: "Generated Slicer Caption"},
		{name: "generated-xlsx-slicer-cache.xlsx", wantText: "Generated Slicer Source Field"},
		{name: "generated-xlsx-slicer-cache.xlsx", wantText: "Generated Slicer Item North"},
		{name: "generated-xlsx-workbook-table-names.xlsx", wantText: "GeneratedDefinedName"},
		{name: "generated-xlsx-workbook-table-names.xlsx", wantText: "GeneratedTableDisplay"},
		{name: "generated-xlsx-workbook-table-names.xlsx", wantText: "Generated Column Two"},
		{name: "generated-xlsx-rich-parts.xlsx", wantText: "Generated Inline String", minImages: 1},
		{name: "generated-xlsx-rich-parts.xlsx", wantText: "Generated XLSX Comment Text"},
		{name: "generated-xlsx-rich-parts.xlsx", wantText: "Generated XLSX Chart Text"},
		{name: "generated-xlsx-thumbnail.xlsx", wantText: "Generated XLSX Thumbnail Text"},
		{name: "generated-xlsx-validation-hyperlink-attrs.xlsx", wantText: "Generated Validation Prompt Title"},
		{name: "generated-xlsx-validation-hyperlink-attrs.xlsx", wantText: "Generated Validation Error Body"},
		{name: "generated-xlsx-validation-hyperlink-attrs.xlsx", wantText: "Generated XLSX Hyperlink Display Attribute"},
		{name: "generated-xlsx-validation-hyperlink-attrs.xlsx", wantText: "Generated XLSX Hyperlink Tooltip Attribute"},
		{name: "stress025.docx", allowEmpty: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tt.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			if !tt.allowEmpty && strings.TrimSpace(res.Text) == "" && len(res.Images) == 0 {
				t.Fatal("extracted no text or images")
			}
			if tt.wantText != "" && !strings.Contains(strings.ToLower(res.Text), strings.ToLower(tt.wantText)) {
				t.Fatalf("missing text %q in %q", tt.wantText, res.Text)
			}
			if len(res.Images) < tt.minImages {
				t.Fatalf("expected at least %d images, got %d", tt.minImages, len(res.Images))
			}
		})
	}
}

func TestMetadataIsOptIn(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "47304.doc"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Just a") {
		t.Fatalf("missing visible document text in %q", res.Text)
	}
	for _, hidden := range []string{"Normal.dotm", "Jelmer Kuperus", "Microsoft Word 12.1.0"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("default output included non-visible metadata %q in %q", hidden, res.Text)
		}
	}

	withMetadata, err := Extract(filepath.Join("testdata", "samples", "47304.doc"), Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(withMetadata.Text, "Normal.dotm") {
		t.Fatalf("metadata option did not include legacy document properties in %q", withMetadata.Text)
	}

	cases := []struct {
		name string
		text string
	}{
		{name: "generated-docx-docprops.docx", text: "Generated Metadata Title"},
		{name: "generated-docx-customxml.docx", text: "Generated DOCX Custom XML Value"},
		{name: "generated-docx-relationships.docx", text: "https://example.test/generated-docx-relationship-link"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			plain, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			if strings.Contains(plain.Text, tc.text) {
				t.Fatalf("default output included metadata/custom package text %q in %q", tc.text, plain.Text)
			}
			meta, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{IncludeMetadata: true})
			if err != nil {
				t.Fatal(err)
			}
			if !strings.Contains(meta.Text, tc.text) {
				t.Fatalf("metadata option missing %q in %q", tc.text, meta.Text)
			}
		})
	}
}

func TestOOXMLRelationshipMetadataKeepsOnlyTextLinks(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible document text</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/visible-link" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/image.jpg" TargetMode="External"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="/ppt/slides/slide2.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///C:/Users/me/hidden.docx" TargetMode="External"/></Relationships>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "relationship-metadata.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	plain, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(plain.Text, "https://example.test/visible-link") {
		t.Fatalf("default output included relationship metadata in %q", plain.Text)
	}
	meta, err := Extract(file, Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(meta.Text, "Visible document text") || !strings.Contains(meta.Text, "https://example.test/visible-link") {
		t.Fatalf("metadata output missing visible text link in %q", meta.Text)
	}
	for _, hidden := range []string{"https://example.test/image.jpg", "/ppt/slides/slide2.xml", "file:///C:/Users/me/hidden.docx"} {
		if strings.Contains(meta.Text, hidden) {
			t.Fatalf("metadata output kept non-text relationship %q in %q", hidden, meta.Text)
		}
	}
}

func TestOOXMLMixedCaseMetadataIsOptIn(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "Word/Document.XML", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible mixed-case body</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "DocProps/Core.XML", `<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:title>Mixed Case Metadata Title</dc:title><dc:creator>Mixed Case Metadata Author</dc:creator></cp:coreProperties>`)
	addZip(t, zw, "CustomXML/Item1.XML", `<root><value>Mixed Case Custom XML Value</value></root>`)
	addZip(t, zw, "Word/_rels/Document.XML.RELS", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/mixed-case-link" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.test/mixed-case-image.png" TargetMode="External"/></Relationships>`)
	addZipBytes(t, zw, "DocProps/Thumbnail.PNG", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case-metadata.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	plain, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(plain.Text, "Visible mixed-case body") {
		t.Fatalf("missing visible body text in %q", plain.Text)
	}
	for _, hidden := range []string{"Mixed Case Metadata Title", "Mixed Case Custom XML Value", "https://example.test/mixed-case-link"} {
		if strings.Contains(plain.Text, hidden) {
			t.Fatalf("default output included mixed-case metadata %q in %q", hidden, plain.Text)
		}
	}
	if len(plain.Images) != 0 {
		t.Fatalf("default output included mixed-case thumbnail: %#v", plain.Images)
	}
	meta, err := Extract(file, Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible mixed-case body", "Mixed Case Metadata Title", "Mixed Case Metadata Author", "Mixed Case Custom XML Value", "https://example.test/mixed-case-link"} {
		if !strings.Contains(meta.Text, want) {
			t.Fatalf("metadata output missing %q in %q", want, meta.Text)
		}
	}
	if strings.Contains(meta.Text, "https://example.test/mixed-case-image.png") {
		t.Fatalf("metadata output kept image relationship target in %q", meta.Text)
	}
	if len(meta.Images) != 1 || meta.Images[0].Name != "Thumbnail.png" || !validImageData(".png", meta.Images[0].Data) {
		t.Fatalf("metadata output missing valid mixed-case thumbnail, got %#v", meta.Images)
	}
}

func TestOOXMLMetadataFiltersInternalPropertyValues(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible body</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "docProps/core.xml", `<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:title>Visible Metadata Title</dc:title><dc:creator>(rId7)</dc:creator><dc:description>file://server/share/hidden.docx</dc:description></cp:coreProperties>`)
	addZip(t, zw, "docProps/app.xml", `<Properties xmlns="urn:props"><Application>Visible Metadata App</Application><Template>docProps/core.xml</Template></Properties>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "metadata-internal-values.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible body", "Visible Metadata Title", "Visible Metadata App"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("metadata output missing visible value %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"rId7", "file://server/share/hidden.docx", "docProps/core.xml"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("metadata output kept internal property value %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLMetadataThumbnailsUseAllSupportedImageExts(t *testing.T) {
	wdp := testJPEGXR()
	dib := testDIB()
	bmp, ok := dibToBMP(dib)
	if !ok {
		t.Fatal("test DIB did not convert to BMP")
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible thumbnail body</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "docProps/thumbnail.dib", append(dib, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "docProps/thumbnail.wdp", append(wdp, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "docProps/thumbnail.emz", gzipBytes(t, testEMF()))
	addZipBytes(t, zw, "docProps/thumbnail.wmz", gzipBytes(t, testPlaceableWMF()))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "metadata-thumbnails.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	plain, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(plain.Images) != 0 {
		t.Fatalf("default output included metadata thumbnails: %#v", plain.Images)
	}
	outDir := filepath.Join(dir, "images")
	meta, err := Extract(file, Options{IncludeMetadata: true, ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(meta.Images) != 4 {
		t.Fatalf("expected four metadata thumbnails, got %#v", meta.Images)
	}
	want := []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "thumbnail.bmp", ext: ".bmp", data: bmp},
		{name: "thumbnail.emf", ext: ".emf", data: testEMF()},
		{name: "thumbnail.wdp", ext: ".wdp", data: wdp},
		{name: "thumbnail.wmf", ext: ".wmf", data: testPlaceableWMF()},
	}
	for i, w := range want {
		img := meta.Images[i]
		if img.Name != w.name || img.Ext != w.ext || !bytes.Equal(img.Data, w.data) || !validImageData(w.ext, img.Data) {
			t.Fatalf("expected thumbnail %+v, got %#v", w, img)
		}
		b, err := os.ReadFile(filepath.Join(outDir, w.name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(w.ext, b) {
			t.Fatalf("written metadata thumbnail is invalid: %s", w.name)
		}
	}
}

func TestNegativeSamplesDoNotPanic(t *testing.T) {
	samples, err := filepath.Glob(filepath.Join("testdata", "negative", "*.*"))
	if err != nil {
		t.Fatal(err)
	}
	if len(samples) == 0 {
		t.Fatal("expected negative samples")
	}
	for _, sample := range samples {
		t.Run(filepath.Base(sample), func(t *testing.T) {
			defer func() {
				if r := recover(); r != nil {
					t.Fatalf("panic extracting %s: %v", sample, r)
				}
			}()
			_, _ = Extract(sample, Options{})
		})
	}
}

func TestExtractedSampleTextIsClean(t *testing.T) {
	samples, err := filepath.Glob(filepath.Join("testdata", "samples", "*.*"))
	if err != nil {
		t.Fatal(err)
	}
	for _, sample := range samples {
		ext := strings.ToLower(filepath.Ext(sample))
		if !map[string]bool{".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true}[ext] {
			continue
		}
		t.Run(filepath.Base(sample), func(t *testing.T) {
			res, err := Extract(sample, Options{})
			if err != nil {
				t.Fatal(err)
			}
			assertCleanExtractedText(t, res.Text)
		})
	}
}

func TestEncryptedOOXMLContainerDoesNotEmitControlText(t *testing.T) {
	samples := []string{
		filepath.Join("testdata", "samples", "58616.xlsx"),
		filepath.Join("testdata", "samples", "protected_passtika.xlsx"),
		filepath.Join("testdata", "negative", "60320-protected.xlsx"),
	}
	for _, sample := range samples {
		t.Run(filepath.Base(sample), func(t *testing.T) {
			if _, err := os.Stat(sample); err != nil {
				t.Skip(err)
			}
			res, err := Extract(sample, Options{})
			if err != nil {
				t.Fatal(err)
			}
			forbidden := []string{
				"Microsoft.Container.DataSpaces",
				"EncryptedPackage",
				"StrongEncryptionDataSpace",
				"EncryptionTransform",
				"Microsoft Enhanced RSA and AES Cryptographic Provider",
			}
			for _, bad := range forbidden {
				if strings.Contains(strings.ToLower(res.Text), strings.ToLower(bad)) {
					t.Fatalf("kept encrypted container text %q in %q", bad, res.Text)
				}
			}
			if strings.TrimSpace(res.Text) != "" || len(res.Images) != 0 {
				t.Fatalf("expected encrypted package to return no plaintext or images, got text len %d images %d", len(res.Text), len(res.Images))
			}
		})
	}
}

func TestExtractOOXMLImageAndText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Hello Office</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "sample.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Hello Office") {
		t.Fatalf("missing text: %q", res.Text)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one image, got %d", len(res.Images))
	}
	if _, err := os.Stat(filepath.Join(outDir, "image1.png")); err != nil {
		t.Fatal(err)
	}
}

func TestOOXMLWordFieldInstructionsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:instrText>HYPERLINK "https://example.test/internal" \h</w:instrText></w:r><w:r><w:t>Visible Link</w:t></w:r></w:p><w:p><w:r><w:instrText>CREATEDATE \@ "d MMMM yyyy"</w:instrText></w:r><w:r><w:t>16 June 2010</w:t></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "fields.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Link", "16 June 2010"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible field result %q in %q", want, res.Text)
		}
	}
	for _, bad := range []string{"HYPERLINK", "CREATEDATE", "example.test/internal"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept field instruction %q in %q", bad, res.Text)
		}
	}
}

func TestOOXMLDeletedRevisionTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible before</w:t></w:r><w:del><w:r><w:delText>Deleted Secret Text</w:delText></w:r><w:r><w:t>Deleted Plain Text</w:t></w:r></w:del><w:moveFrom><w:r><w:t>Moved Away Secret</w:t></w:r></w:moveFrom><w:ins><w:r><w:t>Inserted Visible Text</w:t></w:r></w:ins><w:moveTo><w:r><w:t>Moved Here Visible Text</w:t></w:r></w:moveTo><w:r><w:t>Visible after</w:t></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "deleted-revision.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible before", "Inserted Visible Text", "Moved Here Visible Text", "Visible after"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible revision text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Deleted Secret Text", "Deleted Plain Text", "Moved Away Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden revision text %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLMoveFromRangeContentIsNotVisible(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>
<w:r><w:t>Visible before</w:t></w:r>
<w:moveFromRangeStart w:id="1"/>
<w:r><w:t>Moved Range Secret</w:t></w:r>
<p:pic><p:nvPicPr><p:cNvPr id="2" descr="Moved Range Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
<w:moveFromRangeEnd w:id="1"/>
<w:r><w:t>Visible after</w:t></w:r>
<p:pic><p:nvPicPr><p:cNvPr id="3" descr="Visible Range Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
</w:p></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHidden" Type="x" Target="media/hidden.jpg"/><Relationship Id="rIdVisible" Type="x" Target="media/visible.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/visible.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "move-from-range.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible before", "Visible after"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible move range text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Moved Range Secret", "Moved Range Picture"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden move-from range text %q in %q", hidden, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only visible move range image, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden move-from range image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible before", "Visible after", "![Visible Range Picture](images/visible.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible move range content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Moved Range Secret", "Moved Range Picture", "hidden"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden move-from range content %q in:\n%s", hidden, md)
		}
	}
}

func TestOOXMLResourcePathAttributesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:wp="urn:wp" xmlns:a="urn:a"><w:body><w:p><w:r><w:t>Visible body text</w:t></w:r></w:p><wp:docPr descr="Generated visible description" title="C:\Users\me\Pictures\hidden.jpg"/><wp:docPr descr="/ppt/slides/slide2.xml" title="file:///C:/Users/me/hidden.png"/><wp:docPr descr="ppt/media/image1.png"/></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "resource-path-attrs.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible body text", "Generated visible description"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"C:\\Users\\me\\Pictures\\hidden.jpg", "/ppt/slides/slide2.xml", "file:///C:/Users/me/hidden.png", "ppt/media/image1.png"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden resource reference %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLResourcePathTextLinesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody>
<a:p><a:r><a:t>Visible slide text</a:t></a:r></a:p>
<a:p><a:r><a:t>/ppt/slides/slide2.xml</a:t></a:r></a:p>
<a:p><a:r><a:t>ppt/media/image1.png</a:t></a:r></a:p>
<a:p><a:r><a:t>http://example.test/visible</a:t></a:r></a:p>
</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "resource-path-text.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible slide text", "http://example.test/visible"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"/ppt/slides/slide2.xml", "ppt/media/image1.png"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept resource path text %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLVMLHiddenShapesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible body text</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/drawings/vmlDrawing1.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml">
<v:shape id="visible" alt="Visible VML Alt" title="Visible VML Title"><v:textbox><div>Visible VML Text</div></v:textbox></v:shape>
<v:shape id="hidden1" style="visibility:hidden" alt="Hidden VML Alt" title="Hidden VML Title"><v:textbox><div>Hidden VML Text</div></v:textbox></v:shape>
<v:shape id="hidden2" style="display:none"><v:textbox><div>Display None VML Text</div></v:textbox></v:shape>
<v:shape id="hidden3" hidden="true"><v:textbox><div>Hidden Attr VML Text</div></v:textbox></v:shape>
</xml>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-vml.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible body text", "Visible VML Alt", "Visible VML Title", "Visible VML Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible VML text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden VML Alt", "Hidden VML Title", "Hidden VML Text", "Display None VML Text", "Hidden Attr VML Text"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden VML text %q in %q", hidden, res.Text)
		}
	}
}

func TestDOCXUnreferencedVMLIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible body text</w:t></w:r></w:p><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" r:id="rIdVML"/></w:pict></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="drawings/vmlDrawing1.vml"/></Relationships>`)
	addZip(t, zw, "word/drawings/vmlDrawing1.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="visible" alt="Visible Referenced VML Alt"><v:textbox><div>Visible Referenced VML Text</div></v:textbox></v:shape></xml>`)
	addZip(t, zw, "word/drawings/vmlDrawing2.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="internal" alt="Internal Unreferenced VML Alt"><v:textbox><div>Internal Unreferenced VML Secret</div></v:textbox></v:shape></xml>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-vml.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible body text", "Visible Referenced VML Alt", "Visible Referenced VML Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing referenced VML text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced VML Alt", "Internal Unreferenced VML Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced VML text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Visible body text", "## VML", "Visible Referenced VML Alt", "Visible Referenced VML Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing referenced VML content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced VML Alt", "Internal Unreferenced VML Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced VML content %q in:\n%s", hidden, md)
		}
	}
}

func TestOOXMLHiddenRunTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:p="urn:p"><w:body><w:p><w:r><w:t>Visible before</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden Vanish Secret</w:t><p:pic><p:nvPicPr><p:cNvPr descr="Hidden Vanish Picture Description" title="Hidden Vanish Picture Title"/></p:nvPicPr></p:pic></w:r><w:r><w:rPr><w:webHidden/></w:rPr><w:t>Hidden Web Secret</w:t><p:pic><p:nvPicPr><p:cNvPr descr="Hidden Web Picture Description" title="Hidden Web Picture Title"/></p:nvPicPr></p:pic></w:r><w:r><w:t>Visible after</w:t></w:r></w:p><w:p><w:pPr><w:vanish/></w:pPr><w:r><w:t>Hidden Paragraph Secret</w:t><p:pic><p:nvPicPr><p:cNvPr descr="Hidden Paragraph Picture Description" title="Hidden Paragraph Picture Title"/></p:nvPicPr></p:pic></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-run.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible before", "Visible after"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible run text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, hidden := range []string{"Hidden Vanish Secret", "Hidden Web Secret", "Hidden Vanish Picture Description", "Hidden Vanish Picture Title", "Hidden Web Picture Description", "Hidden Web Picture Title", "Hidden Paragraph Secret", "Hidden Paragraph Picture Description", "Hidden Paragraph Picture Title"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden run text %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden run text/attribute %q in:\n%s", hidden, md)
		}
	}
}

func TestOOXMLRelationshipIDAttributesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a"><w:body><w:p><w:r><w:t>Visible Body</w:t></w:r></w:p><p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible Picture Description" title="rId42" alt="rId7"/></p:nvPicPr><p:blipFill><a:blip/></p:blipFill></p:pic><p:pic><p:nvPicPr><p:cNvPr id="2" descr="PowerPoint.Slide.80" title="AcroExch.Document.11" alt="CLSID={00020906-0000-0000-C000-000000000046}"/></p:nvPicPr><p:blipFill><a:blip/></p:blipFill></p:pic></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "relationship-id-attrs.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Body") || !strings.Contains(res.Text, "Visible Picture Description") {
		t.Fatalf("missing visible DOCX attribute/body text in %q", res.Text)
	}
	for _, hidden := range []string{"rId42", "rId7", "PowerPoint.Slide.80", "AcroExch.Document.11", "CLSID={00020906-0000-0000-C000-000000000046}"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden/control attribute %q in text %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible Picture Description") {
		t.Fatalf("markdown missing visible DOCX attribute text:\n%s", md)
	}
	for _, hidden := range []string{"rId42", "rId7", "PowerPoint.Slide.80", "AcroExch.Document.11", "CLSID={00020906-0000-0000-C000-000000000046}"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden/control attribute %q:\n%s", hidden, md)
		}
	}
}

func TestOOXMLNamespaceAttributesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body>
<w:p><w:r><w:t>Visible namespace body</w:t></w:r></w:p>
<w:p><w:r><w:t>xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"</w:t></w:r></w:p>
<w:p><w:r><w:t>mc:Ignorable="w14 wp14"</w:t></w:r></w:p>
<w:p><w:r><w:t>xsi:schemaLocation="http://schemas.openxmlformats.org/wordprocessingml/2006/main wordprocessingml.xsd"</w:t></w:r></w:p>
<w:p><w:r><w:t>Visible schema footer</w:t></w:r></w:p>
</w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "namespace-attrs.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible namespace body", "Visible schema footer"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible namespace test text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"xmlns:", "schemas.openxmlformats.org", "mc:Ignorable", "schemaLocation", "w14 wp14"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept Office XML namespace metadata %q in text %q", hidden, res.Text)
		}
		if strings.Contains(res.Markdown("images"), hidden) {
			t.Fatalf("kept Office XML namespace metadata %q in markdown:\n%s", hidden, res.Markdown("images"))
		}
	}
}

func TestOOXMLHiddenParagraphTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible before</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>Hidden Paragraph Secret</w:t><w:tab/><w:t>Hidden Tail</w:t><w:br/><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p><w:p><w:pPr><w:rPr><w:webHidden/></w:rPr></w:pPr><w:r><w:t>Hidden Web Paragraph Secret</w:t></w:r></w:p><w:p><w:r><w:t>Visible after</w:t></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-paragraph.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible before", "Visible after"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible paragraph text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Paragraph Secret", "Hidden Tail", "Hidden Web Paragraph Secret", "✓"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden paragraph text %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLFormCheckboxesAreVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Done </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:checked/></w:checkBox></w:ffData></w:fldChar><w:t> Pending </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:checked w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t> Hidden </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:checked/></w:checkBox></w:ffData></w:fldChar></w:r></w:p><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>Hidden paragraph checkbox </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox/></w:ffData></w:fldChar></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "checkboxes.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Done ☒", "Pending ☐"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible checkbox text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden", "Hidden paragraph checkbox"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden checkbox text %q in %q", hidden, res.Text)
		}
	}
	if strings.Count(res.Text, "☒") != 1 || strings.Count(res.Text, "☐") != 1 {
		t.Fatalf("unexpected checkbox count in %q", res.Text)
	}
}

func TestOOXMLFormDropdownSelectedItemIsVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Status </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:result w:val="1"/><w:listEntry w:val="Draft"/><w:listEntry w:val="Approved"/><w:listEntry w:val="Archived"/></w:ddList></w:ffData></w:fldChar><w:t> Priority </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:listEntry w:val="Low"/><w:listEntry w:val="High"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t> Hidden </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:result w:val="1"/><w:listEntry w:val="Visible"/><w:listEntry w:val="Hidden Choice"/></w:ddList></w:ffData></w:fldChar></w:r></w:p><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>Hidden paragraph dropdown </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:listEntry w:val="Hidden Default"/></w:ddList></w:ffData></w:fldChar></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "dropdown.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Status Approved", "Priority Low"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible dropdown text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Draft", "Archived", "Hidden", "Hidden Choice", "Hidden paragraph dropdown", "Hidden Default"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden or unselected dropdown text %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLFormTextInputDefaultIsVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Name </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:type w:val="regular"/><w:default w:val="Alice Example Target: ../media/hidden.png"/><w:maxLength w:val="40"/><w:format w:val="Title Case"/></w:textInput></w:ffData></w:fldChar><w:t> Code </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:default w:val="A-123 rId77"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t> Hidden </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:default w:val="Hidden Default"/></w:textInput></w:ffData></w:fldChar></w:r></w:p><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>Hidden paragraph text input </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:default w:val="Hidden Paragraph Default"/></w:textInput></w:ffData></w:fldChar></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "text-input.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Name Alice Example", "Code A-123"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible text input default %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"regular", "Title Case", "40", "Target:", "../media/hidden.png", "rId77", "Hidden", "Hidden Default", "Hidden paragraph text input", "Hidden Paragraph Default"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden or internal text input value %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLInlineSpecialCharactersAreVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>No</w:t><w:noBreakHyphen/><w:t>Break</w:t></w:r><w:r><w:t>Soft</w:t><w:softHyphen/><w:t>Hyphen</w:t></w:r><w:r><w:t>A</w:t><w:tab/><w:t>Tab</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t><w:noBreakHyphen/><w:t>NoBreak</w:t><w:softHyphen/><w:t>SoftHyphen</w:t><w:tab/><w:t>Tab</w:t></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "inline-special.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"No\u2011Break", "SoftHyphen", "A Tab"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible inline character text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden", "NoBreak", "TabTab"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden inline character text %q in %q", hidden, res.Text)
		}
	}
}

func TestOOXMLHiddenBreaksAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a"><w:body><w:p><w:r><w:t>Line</w:t><w:br/><w:t>Break</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t><w:br/><w:t>Break</w:t></w:r><p:sp><p:nvSpPr><p:cNvPr hidden="1"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Hidden Shape</a:t><a:br/><a:t>Break</a:t></a:r></a:p></p:txBody></p:sp><w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-breaks.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Line\nBreak") {
		t.Fatalf("missing visible break in %q", res.Text)
	}
	if !strings.Contains(res.Text, "Tail") {
		t.Fatalf("missing visible tail text in %q", res.Text)
	}
	for _, hidden := range []string{"Hidden", "Hidden Shape"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden break text %q in %q", hidden, res.Text)
		}
	}
	if strings.Contains(res.Text, "\n\n\n") {
		t.Fatalf("hidden breaks left excessive blank lines in %q", res.Text)
	}
}

func TestOOXMLVisibleSymbolsAreExtractedAsText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Math </w:t><w:sym w:font="Symbol" w:char="F061"/><w:t> + </w:t><w:sym w:font="Symbol" w:char="F062"/></w:r><w:r><w:t> Check </w:t><w:sym w:font="Wingdings" w:char="F0FC"/></w:r><w:r><w:t> Arrow </w:t><w:sym w:font="Segoe UI Symbol" w:char="2192"/></w:r><w:r><w:t> Private </w:t><w:sym w:font="PrivateFont" w:char="F123"/><w:t>Tail</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "symbols.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Math α + β", "Check ✓", "Arrow →", "Private Tail"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible symbol text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Hidden") || strings.Contains(res.Text, "\uf0fc") || strings.Contains(res.Text, "\uf123") {
		t.Fatalf("kept hidden or private-use symbol text in %q", res.Text)
	}
}

func TestLegacyWordFieldInstructionsAreStrippedFromText(t *testing.T) {
	text := cleanText(`AUTHOR \* Upper \* MERGEFORMAT ANTONI
CREATEDATE \@ "d MMMM yyyy" \* MERGEFORMAT 16 June 2010
Please visit HYPERLINK "https://example.test/internal" \h Visible Link
MERGEFIELD FirstName «FirstName»
Section PAGEREF _Toc12345 \h 7
Name: FORMTEXT Done: FORMCHECKBOX
Field in EndNote. File size: FILESIZE \* MERGEFORMAT 0
Textbox: EDITTIME \* MERGEFORMAT 2
Property: DOCPROPERTY "PackageName" \* MERGEFORMAT Insert package name here.
Style: STYLEREF \n"見出し 1" \* MERGEFORMAT Chapter 3`)
	for _, want := range []string{"ANTONI", "16 June 2010", "Visible Link", "«FirstName»", "Section 7", "Name:", "Done:", "File size: 0", "Textbox: 2", "Property: Insert package name here.", "Style: Chapter 3"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"AUTHOR", "CREATEDATE", "HYPERLINK", "MERGEFIELD", "PAGEREF", "FORMTEXT", "FORMCHECKBOX", "FILESIZE", "EDITTIME", "DOCPROPERTY", "STYLEREF", "MERGEFORMAT", "Upper", "example.test/internal"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordLinkAndSymbolFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Linked value: LINK Excel.Sheet.8 "C:\Users\me\source.xlsx" "Sheet1!R1C1" \a \f 4 Visible linked cell
Remote: LINK Word.Document.8 https://example.test/internal.docx \a Remote visible text
Bullet: SYMBOL 183 \f Symbol \s 10 Visible bullet text
Symbolic prose should remain visible.`)
	for _, want := range []string{"Linked value: Visible linked cell", "Remote: Remote visible text", "Bullet: Visible bullet text", "Symbolic prose should remain visible."} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible field result %q in %q", want, text)
		}
	}
	for _, bad := range []string{"LINK", "Excel.Sheet.8", "Word.Document.8", "source.xlsx", "Sheet1!R1C1", "example.test/internal.docx", "SYMBOL 183", `\f Symbol`, `\s 10`} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept LINK/SYMBOL field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordAdditionalFieldInstructionsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Included: INCLUDETEXT "C:\Users\me\source.docx" \c HTML Included visible text
Ref doc: RD "\\server\share\appendix.doc" \f
Quote: QUOTE "Visible quoted text" \* MERGEFORMAT
Numbers: LISTNUM LegalDefault \l 2 Numbered clause
Auto: AUTOTEXT "Signature Block" Visible signature
AutoNum: AUTONUMOUT Visible outline number
The report can include text as visible prose.`)
	for _, want := range []string{
		"Included: Included visible text",
		"Ref doc:",
		"Quote: \"Visible quoted text\"",
		"Numbers: LegalDefault \\l 2 Numbered clause",
		"Auto: \"Signature Block\" Visible signature",
		"AutoNum: Visible outline number",
		"The report can include text as visible prose.",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible field result %q in %q", want, text)
		}
	}
	for _, bad := range []string{"INCLUDETEXT", "source.docx", "appendix.doc", "RD", "QUOTE", "MERGEFORMAT", "LISTNUM", "AUTOTEXT", "AUTONUMOUT"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept additional field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordIncludePicturePathsForSupportedFormatsAreStripped(t *testing.T) {
	text := cleanText(`Visible before
INCLUDEPICTURE "C:\Users\me\Pictures\diagram.webp" \* MERGEFORMATINET WebP caption
INCLUDEPICTURE "https://example.test/assets/vector.svgz" \* MERGEFORMATINET SVG caption
INCLUDEPICTURE "https://example.test/assets/diagram.eps" \* MERGEFORMATINET EPS caption
INCLUDEPICTURE "file://server/share/photo.heic" \* MERGEFORMATINET HEIC caption
INCLUDEPICTURE "C:\Users\me\Pictures\scan.jp2" \* MERGEFORMATINET JP2 caption
INCLUDEPICTURE "C:\Users\me\Pictures\tile.j2c" \* MERGEFORMATINET J2C caption
INCLUDEPICTURE "C:\Users\me\Pictures\codestream.jpc" \* MERGEFORMATINET JPC caption
INCLUDEPICTURE "C:\Users\me\Pictures\legacy.pct" \* MERGEFORMATINET PICT caption
INCLUDEPICTURE "C:\Users\me\Pictures\photo.jpe" \* MERGEFORMATINET JPE caption
INCLUDEPICTURE "C:\Users\me\Pictures\photo.jfif" \* MERGEFORMATINET JFIF caption
INCLUDEPICTURE "C:\Users\me\Pictures\photo.wdp" \* MERGEFORMATINET WDP caption
INCLUDEPICTURE "C:\Users\me\Pictures\photo.jxr" \* MERGEFORMATINET JXR caption
INCLUDEPICTURE "C:\Users\me\Pictures\photo.hdp" \* MERGEFORMATINET HDP caption
INCLUDEPICTURE "C:\Users\me\Pictures\vector.emz" \* MERGEFORMATINET EMZ caption
INCLUDEPICTURE "C:\Users\me\Pictures\vector.wmz" \* MERGEFORMATINET WMZ caption
INCLUDEPICTURE "C:\Users\me\Pictures\texture.tga" \* MERGEFORMATINET TGA caption
"media/relative-diagram.png" \* MERGEFORMATINET Relative caption
"..\media\relative-vector.emf" \* MERGEFORMATINET Parent caption
"inline-photo.jpg" \* MERGEFORMATINET Bare caption
"file:///C:/Users/me/Pictures/vector.ps" \* MERGEFORMATINET File URI caption
Visible after`)
	for _, want := range []string{"Visible before", "WebP caption", "SVG caption", "EPS caption", "HEIC caption", "JP2 caption", "J2C caption", "JPC caption", "PICT caption", "JPE caption", "JFIF caption", "WDP caption", "JXR caption", "HDP caption", "EMZ caption", "WMZ caption", "TGA caption", "Relative caption", "Parent caption", "Bare caption", "File URI caption", "Visible after"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"INCLUDEPICTURE", "diagram.webp", "vector.svgz", "diagram.eps", "photo.heic", "scan.jp2", "tile.j2c", "codestream.jpc", "legacy.pct", "photo.jpe", "photo.jfif", "photo.wdp", "photo.jxr", "photo.hdp", "vector.emz", "vector.wmz", "texture.tga", "relative-diagram.png", "relative-vector.emf", "inline-photo.jpg", "vector.ps", "MERGEFORMATINET"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept picture field instruction/path %q in %q", bad, text)
		}
	}
}

func TestLegacyWordPCTPictureFieldPathIsStripped(t *testing.T) {
	text := cleanText(`Cover "media/legacy.pct" \* MERGEFORMATINET Legacy PICT caption
Tail`)
	if !strings.Contains(text, "Cover Legacy PICT caption") || !strings.Contains(text, "Tail") {
		t.Fatalf("missing visible PCT field result in %q", text)
	}
	for _, bad := range []string{"legacy.pct", "MERGEFORMATINET"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept PCT picture field path %q in %q", bad, text)
		}
	}
}

func TestLegacyWordObjectFieldInstructionsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Embedded: EMBED "Equation" \* mergeformat
Object: EMBED Word.Document.12 \s
Clipart: EMBED MS_ClipArt_Gallery
Button: MACROBUTTON NoMacro [Document Title] MACROBUTTON GetDate [Alt+R for Date]
Equation (2) gives the force on the particle.
Package capsules in HDPE bottles.`)
	for _, want := range []string{"Embedded:", "Object:", "Clipart:", "Button: [Document Title] [Alt+R for Date]", "Equation (2) gives the force", "Package capsules"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"EMBED", "Word.Document.12", "MS_ClipArt_Gallery", "MACROBUTTON", "NoMacro", "GetDate"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept object field instruction %q in %q", bad, text)
		}
	}
}

func TestOLEClassFragmentsAreFilteredWithoutDroppingProse(t *testing.T) {
	dropped := []string{
		"Equation.3",
		"Equation.DSMT4",
		"Microsoft Equation",
		"Microsoft Equation 3.0",
		"MathType Equation",
		"MathType 5.0 Equation",
		"Word.Document.8",
		"PowerPoint.Show.8",
		"PowerPoint.Slide.12",
		"PowerPoint.Presentation.12",
		"PowerPoint.Template.8",
		"PowerPoint Presentation",
		"PowerPoint Slide",
		"CompObj",
		"ObjInfo",
		"Ole10Native",
		"OlePres000",
		"Relationships",
		"ContentType",
		"PartName",
		"TargetMode",
		"MS_ClipArt_Gallery.5",
		"MSPhotoEd.3",
		"MSGraph.Chart.8",
		"OrgPlusWOPX.4",
		"Photoshop.Image.8",
		"Excel.Chart.5",
		"SmartDraw.2",
		"CorelDRAW.Graphic.10",
		"WordPad.Document.1",
		"RichEdit.Document.1",
		"MediaPlayer.MediaPlayer.1",
		"ShockwaveFlash.ShockwaveFlash.9",
		"Forms.TextBox.1",
		"Forms.CommandButton.1",
		"MSForms.CheckBox.1",
		"MSForms.ComboBox.1",
		"Shell.Explorer",
		"Shell.Explorer.2",
		"HTMLDocument",
		"htmlfile",
		"Internet Explorer_Server",
		"Outlook.FileAttach",
		"Outlook.FileAttach.1",
		"Outlook.Message.1",
		"MSComctlLib.ListViewCtrl.2",
		"MSComctlLib.TreeCtrl.2",
		"MSComctlLib.ImageListCtrl.2",
		"MSComCtl2.DTPicker.2",
		"WMPlayer.OCX.7",
		"SmartDraw",
		"SmartDraw Drawing",
		"Bitmap Image",
		"Paint.Picture",
		"CorelDRAW",
		"CorelDRAW 10.0 Graphic",
		"WordPad Document",
		"RichEdit Document",
		"Media Clip",
		"Windows Media Player",
		"Shockwave Flash Object",
		"Macromedia Flash Factory Object",
		"Microsoft Forms 2.0 TextBox",
		"Microsoft Forms 2.0 CommandButton",
		"HTML Document",
		"Shell Explorer",
		"Current User",
		"CacheLastModifiedFactor.1",
		"Microsoft Word Document",
		"Microsoft Excel Worksheet",
		"Microsoft Excel 97-2003 Worksheet",
		"Microsoft Excel 2007 Worksheet",
		"Microsoft Excel 2007 Workbook",
		"Microsoft Excel Chart",
		"Microsoft Graph Chart",
		"Microsoft Graph 97 Chart",
		"Microsoft Graph 2000 Chart",
		"Microsoft PowerPoint Slide",
		"Microsoft PowerPoint 97-2003 Presentation",
		"Microsoft Office Excel Worksheet",
		"Microsoft Office Excel 97-2003 Worksheet",
		"Microsoft Office Excel 2007 Worksheet",
		"Microsoft Office Excel 2007 Workbook",
		"Microsoft Office Word Document",
		"Microsoft Office Word 97-2003 Document",
		"Microsoft Office Word 2007 Document",
		"Microsoft Word 2007 Document",
		"Microsoft PowerPoint Presentation",
		"Microsoft PowerPoint 2007 Presentation",
		"Microsoft Office PowerPoint Presentation",
		"Microsoft Office PowerPoint 97-2003 Presentation",
		"Microsoft Office PowerPoint 2007 Presentation",
		"Adobe Photoshop Image",
		"Adobe Acrobat Document",
		"Acrobat Document",
		"PDF Document",
		"AcroExch.Document",
		"AcroExch.Document.DC",
		"Visio.Drawing.11",
		"Microsoft Visio Drawing",
		"MS Org Chart",
		"MS Organization Chart 2.0",
		"Photo Editor Photo",
		"Microsoft Photo Editor 3.0 Photo",
		"Package",
		"Package Object",
		"Packager Shell Object",
		"Sheet1!Object 1",
	}
	for _, s := range dropped {
		if !looksLikeBinaryControlFragment(s) {
			t.Fatalf("expected OLE/control fragment %q to be filtered", s)
		}
		if got := cleanText(s); got != "" {
			t.Fatalf("expected cleanText to drop OLE/control fragment %q, got %q", s, got)
		}
	}
	kept := []string{
		"Equation (2) gives the force on the particle.",
		"Package capsules in HDPE bottles.",
		"Document the worksheet review notes.",
		"The PDF document is attached for review.",
		"Visio drawing migration notes are visible.",
		"Windows media player settings are visible.",
		"RichEdit document support notes are visible.",
		"Forms are visible when users complete the survey.",
		"Microsoft PowerPoint presentation notes are visible.",
		"Microsoft Office Word 97-2003 document migration notes are visible.",
		"Microsoft Office Excel 2007 workbook migration notes are visible.",
		"Microsoft Office PowerPoint 2007 presentation notes are visible.",
		"Object pooling notes are visible.",
		"Customer relationships are visible.",
	}
	for _, s := range kept {
		if looksLikeBinaryControlFragment(s) {
			t.Fatalf("expected visible prose %q to be kept", s)
		}
		if got := cleanText(s); got != s {
			t.Fatalf("expected cleanText to keep visible prose %q, got %q", s, got)
		}
	}
}

func TestUniqueStringsDropsHiddenResourceReferences(t *testing.T) {
	got := strings.Join(uniqueStrings([]string{
		"Visible document text",
		"word/media/image1.png",
		"file:///C:/Users/me/hidden.docx",
		"(rId7)",
		"C:\\Reports\\Q1 is visible user text",
		"Visible before\nContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"\nVisible after Target=\"../media/hidden.png\"",
		`PartName="/word/document.xml"`,
	}), "\n")
	for _, want := range []string{"Visible document text", "C:\\Reports\\Q1 is visible user text", "Visible before", "Visible after"} {
		if !strings.Contains(got, want) {
			t.Fatalf("uniqueStrings dropped visible text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"word/media/image1.png", "file:///C:/Users/me/hidden.docx", "rId7", "ContentType", "application/vnd.openxmlformats", "Target=", "../media/hidden.png", "PartName", "/word/document.xml"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("uniqueStrings kept hidden reference %q in %q", hidden, got)
		}
	}
}

func TestCorruptOLEFallbackTextUsesVisibleCleaning(t *testing.T) {
	data := []byte(strings.Join([]string{
		"Visible corrupt wrapper words Content-Location: word/media/image1.png",
		`ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`,
		`TargetMode="External"`,
		`r:embed="rId88"`,
		"Another visible corrupt line Target=\"../media/hidden.jpg\"",
	}, "\x00"))
	text := strings.Join(extractCorruptOLEText(data), "\n")
	for _, want := range []string{"Visible corrupt wrapper words", "Another visible corrupt line"} {
		if !strings.Contains(text, want) {
			t.Fatalf("corrupt OLE fallback dropped visible text %q in %q", want, text)
		}
	}
	for _, hidden := range []string{"Content-Location", "word/media/image1.png", "ContentType", "application/vnd.openxmlformats", "TargetMode", "External", "r:embed", "rId88", "Target=", "../media/hidden.jpg"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("corrupt OLE fallback kept hidden reference %q in %q", hidden, text)
		}
	}
}

func TestLegacyWordTOCBookmarkFieldIsStrippedFromText(t *testing.T) {
	text := cleanText(`Table of Contents TOC "__RefHeading__8_476954814"1. TERM 1
"__RefHeading__10_476954814" 1.1 Termination. 1
"_Toc70481093" INTRODUCTION 1
_Toc34542813 \h
Visible Table of Contents`)
	for _, want := range []string{"Table of Contents 1. TERM 1", "1.1 Termination. 1", "INTRODUCTION 1", "Visible Table of Contents"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible TOC text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"__RefHeading__", "_Toc", `TOC "`} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept TOC field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordSequenceAndPictureFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Figure SEQ Figure \* ARABIC 1 Spacewalk
SEQ CHAPTER \h \r 1My name is Ryan.
"c:\# astus\astus1\Tit_fav.gif" \* MERGEFORMATINET Visible favorites text.
://hk.yimg.com/i/hk/adv/compaq/powrdbyhp_blu_84x28_yahoo.gif" \* MERGEFORMATINET Yahoo visible text.
MERGEFORMATINET`)
	for _, want := range []string{"Figure 1 Spacewalk", "My name is Ryan.", "Visible favorites text.", "Yahoo visible text."} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"SEQ Figure", "SEQ CHAPTER", "MERGEFORMATINET", "Tit_fav.gif", "powrdbyhp_blu_84x28_yahoo.gif"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordIndexEntryFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Visible before XE "Internal index entry" \b \i Visible after
Heading TC "Hidden TOC entry" \f C \l "2" Display heading
Citation TA "Hidden table authority" \s "Cases" Case visible text`)
	for _, want := range []string{"Visible before Visible after", "Heading Display heading", "Citation Case visible text"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"XE", "TC", "TA", "Internal index entry", "Hidden TOC entry", "Hidden table authority", `\f`, `\l`, `\s`} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept index/TOC authority field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordPromptAndSetFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Visible before ASK ClientName "Internal prompt text" \d "Hidden Default" Visible after
Intro FILLIN "Internal fill-in prompt" \d "Hidden Fill Default" Display answer
SET InternalVar "Hidden variable value" Final visible sentence`)
	for _, want := range []string{"Visible before Visible after", "Intro Display answer", "Final visible sentence"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"ASK", "FILLIN", "SET", "ClientName", "Internal prompt text", "Hidden Default", "Internal fill-in prompt", "Hidden Fill Default", "InternalVar", "Hidden variable value"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept prompt/set field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordIfFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Region IF { MERGEFIELD State } = "CA" "Hidden California Branch" "Hidden Other Branch" California
Status IF "1" = "1" "Hidden Yes Branch" "Hidden No Branch" Approved
Tail visible`)
	for _, want := range []string{"Region California", "Status Approved", "Tail visible"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible IF field result %q in %q", want, text)
		}
	}
	for _, bad := range []string{"IF", "MERGEFIELD", "State", "Hidden California Branch", "Hidden Other Branch", "Hidden Yes Branch", "Hidden No Branch"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept IF field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordMetadataVariableFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Variable DOCVARIABLE "InternalClientCode" \* MERGEFORMAT Visible Client
Saved SAVEDATE \@ "yyyy-MM-dd" 2024-06-01
Printed PRINTDATE \@ "M/d/yyyy" 6/2/2024
Template TEMPLATE \* MERGEFORMAT Normal.dotm Display Template
User USERNAME \* MERGEFORMAT Alice Visible
Tags KEYWORDS \* MERGEFORMAT Visible Tags
Comment COMMENTS \* MERGEFORMAT Visible Comment`)
	for _, want := range []string{"Variable Visible Client", "Saved 2024-06-01", "Printed 6/2/2024", "Template Display Template", "User Alice Visible", "Tags Visible Tags", "Comment Visible Comment"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible metadata/variable field result %q in %q", want, text)
		}
	}
	for _, bad := range []string{"DOCVARIABLE", "InternalClientCode", "SAVEDATE", "PRINTDATE", "TEMPLATE", "USERNAME", "KEYWORDS", "COMMENTS", "MERGEFORMAT", `\@`, "Normal.dotm"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept metadata/variable field instruction %q in %q", bad, text)
		}
	}
}

func TestLegacyWordCitationDatabaseAndPrivateFieldsAreStrippedFromText(t *testing.T) {
	text := cleanText(`Source CITATION HiddenSourceKey \l 1033 \m HiddenLocator Visible citation text
References BIBLIOGRAPHY \l 1033 Visible bibliography text
Data DATABASE \d "C:\Users\me\hidden.odc" \s "SELECT * FROM HiddenTable" Visible table value
Layout ADVANCE \d 4 \r -2 Advanced visible text
Addin ADDIN "Hidden add-in payload" Add-in visible result
Private PRIVATE "Hidden private payload" Private visible result
The citation database remains visible as normal prose.
Private placement notes remain visible.`)
	for _, want := range []string{
		"Source Visible citation text",
		"References Visible bibliography text",
		"Data Visible table value",
		"Layout Advanced visible text",
		"Addin Add-in visible result",
		"Private Private visible result",
		"The citation database remains visible as normal prose.",
		"Private placement notes remain visible.",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible citation/database/private field result %q in %q", want, text)
		}
	}
	for _, bad := range []string{"CITATION", "HiddenSourceKey", "HiddenLocator", "BIBLIOGRAPHY", "DATABASE", "hidden.odc", "HiddenTable", "ADVANCE", `\d 4`, `\r -2`, "ADDIN", "Hidden add-in payload", "PRIVATE", "Hidden private payload"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept citation/database/private field instruction %q in %q", bad, text)
		}
	}
}

func TestMojibakePunctuationIsRepairedAndBinaryNoiseDropped(t *testing.T) {
	text := cleanText("Water quality issues remain\u9225?\n\u9225\u6dd5lobal freshwater consumption rose sixfold.\nRisk Assessment & Control Group (\u9225\u6de9AC\u9225?)\nBiochemical injury pattern \u9225\u6deaignature\u9225? vs protean\nNormal vs abnormal baseline LFT\u9225\u6a9a\nDonors should consider \u9225\u6e04ebt for water swaps\u9225? as assistance.\nTeach them how to fish \u9225\u6e22each them how to fish\u9225?\nSesi\u8d38n FT3.40\n\u6cfb\u80c1\u7b11\u88a8\u7b11\u8927\u7b11\u8927\u5e90\u8927\u7b11\u888c\u80c1\u9225\u2014\u659d\u03c0\u6d4b\u20ac\u2014\u659d\u03c0\u6b8b\u5a26\u6b8b\u5a26\n銆併€傦紝锛庛兓锛氾紱锛燂紒銈涖倻銉姐兙銈濄倿銆呫兗鈥欌€濓級銆曪冀锝濄€夈€嬨€嶃€忋€懧扳€扳€测€斥剝锟狅紖銇併亙銇呫亣銇夈仯銈冦倕銈囥値銈°偅銈ャ偋銈┿儍銉ｃ儱銉с儺銉点兌!%),.:;?]}")
	for _, want := range []string{"Water quality issues remain'", "\"Global freshwater", `("RAC")`, `"Signature" vs protean`, "LFT's", "\"debt for water swaps", "\"teach them how to fish\"", "Sesi\u00f3n FT3.40"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing repaired text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"\u9225?", "\u9225\u6dd5", "\u9225\u6de9", "\u9225\u6dea", "\u9225\u6a9a", "\u9225\u6e04", "\u9225\u6e22", "Sesi\u8d38n", "\u6cfb\u80c1\u7b11", "銆併€"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept mojibake/binary noise %q in %q", bad, text)
		}
	}
	for _, noise := range []string{
		"я0;я[я0",
		"0000еяя$([\\{bябя",
		"、。，．・：；？！゛゜ヽヾゝゞ々ー’”）〕］｝〉》」』】°‰′″℃￠％ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ!%),.:;?]}｡｣､･ｧｨｩｪｫｬｭｮｯｰﾞﾟ",
	} {
		if cleanText(noise) != "" {
			t.Fatalf("expected mojibake control line to be dropped: %q -> %q", noise, cleanText(noise))
		}
	}
}

func TestMojibakeContractionsAreRepaired(t *testing.T) {
	text := cleanText("It\u2019c consistent with our values.\nThat\u2019c a visible sentence.\nI don\u2019t plan to change this.")
	for _, want := range []string{"It\u2019s consistent", "That\u2019s a visible sentence", "I don\u2019t plan"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing repaired contraction %q in %q", want, text)
		}
	}
	for _, bad := range []string{"It\u2019c", "That\u2019c"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept mojibake contraction %q in %q", bad, text)
		}
	}
}

func TestExtractOOXMLSVGAndTIFFImages(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Vector and TIFF</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/vector.svg", testSVG())
	addZipBytes(t, zw, "word/media/scan.tif", testTIFF())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "svg-tiff.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Vector and TIFF") {
		t.Fatalf("missing text: %q", res.Text)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two images, got %d", len(res.Images))
	}
	if _, err := os.Stat(filepath.Join(outDir, "vector.svg")); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "scan.tif")); err != nil {
		t.Fatal(err)
	}
}

func TestOOXMLCompressedSVGImagesAreDecompressed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Compressed SVG</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/vector.svgz", gzipBytes(t, testSVG()))
	addZipBytes(t, zw, "word/media/bad.svgz", gzipBytes(t, []byte(`<html><body>not svg</body></html>`)))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "compressed-svg.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one decompressed SVG image, got %#v", res.Images)
	}
	if res.Images[0].Name != "vector.svg" || res.Images[0].Ext != ".svg" || !validImageData(".svg", res.Images[0].Data) {
		t.Fatalf("expected decompressed SVG, got %#v", res.Images[0])
	}
	b, err := os.ReadFile(filepath.Join(outDir, "vector.svg"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".svg", b) {
		t.Fatal("written decompressed SVG is invalid")
	}
}

func TestOOXMLCompressedImagesAreSniffedFromBin(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Compressed bin images</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/vector-emf.bin", gzipBytes(t, testEMF()))
	addZipBytes(t, zw, "word/media/vector-svg.bin", gzipBytes(t, testSVG()))
	addZipBytes(t, zw, "word/media/vector-wmf.bin", gzipBytes(t, testPlaceableWMF()))
	addZipBytes(t, zw, "word/media/bad-vector.bin", gzipBytes(t, []byte(`<html><body>not svg</body></html>`)))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "compressed-bin-images.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 3 {
		t.Fatalf("expected three decompressed images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
	}{
		{name: "vector-emf.emf", ext: ".emf"},
		{name: "vector-svg.svg", ext: ".svg"},
		{name: "vector-wmf.wmf", ext: ".wmf"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || !validImageData(w.ext, img.Data) {
			t.Fatalf("expected decompressed %s image %s, got %#v", w.ext, w.name, img)
		}
		b, err := os.ReadFile(filepath.Join(outDir, w.name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(w.ext, b) {
			t.Fatalf("written decompressed .bin image is invalid: %s", w.name)
		}
	}
}

func TestInvalidOOXMLImageIsDropped(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Hello Office</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/bad.png", []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1a, '\n', 0, 0, 0, 0, 'I', 'E', 'N', 'D', 0xae, 0x42, 0x60, 0x82})
	addZipBytes(t, zw, "word/media/bad.svg", []byte(`<html><body>not svg</body></html>`))
	addZipBytes(t, zw, "word/media/bad.tif", []byte{'I', 'I', 42, 0, 8, 0, 0, 0, 0, 0})
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "bad-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 0 {
		t.Fatalf("expected invalid image to be dropped, got %d image(s)", len(res.Images))
	}
}

func TestCorruptOOXMLImageCRCDoesNotFailExtraction(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible text survives bad image CRC</w:t></w:r></w:p></w:body></w:document>`)
	hdr := &zip.FileHeader{Name: "word/media/image1.png", Method: zip.Store}
	w, err := zw.CreateHeader(hdr)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Write(testPNG()); err != nil {
		t.Fatal(err)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	corrupt := corruptStoredZipEntryData(t, buf.Bytes(), "word/media/image1.png")
	dir := t.TempDir()
	file := filepath.Join(dir, "corrupt-image.docx")
	if err := os.WriteFile(file, corrupt, 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible text survives bad image CRC") {
		t.Fatalf("expected visible document text to be preserved, got %q", res.Text)
	}
	if len(res.Images) != 0 {
		t.Fatalf("expected corrupt image to be skipped, got %d image(s)", len(res.Images))
	}
}

func TestExtractReturnsOnlyNormalizedValidImagesWithoutImageDir(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible image document</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/mislabelled.jpg", testPNG())
	addZipBytes(t, zw, "word/media/broken.png", []byte("not a png"))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "normalized-images.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one normalized valid image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "mislabelled.png" || img.Ext != ".png" || !validImageData(".png", img.Data) {
		t.Fatalf("expected returned image to be normalized and valid, got %#v", img)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![mislabelled](images/mislabelled.png)") || strings.Contains(md, "broken.png") || strings.Contains(md, "mislabelled.jpg") {
		t.Fatalf("markdown should reference only normalized returned image:\n%s", md)
	}
}

func TestOOXMLThumbnailImagesAreMetadataOptIn(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible document text</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "docProps/thumbnail.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "thumbnail.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}

	plain, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(plain.Images) != 0 {
		t.Fatalf("expected default extraction to skip package thumbnail, got %#v", plain.Images)
	}

	withMetadata, err := Extract(file, Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(withMetadata.Images) != 1 || withMetadata.Images[0].Name != "thumbnail.png" {
		t.Fatalf("expected metadata extraction to include package thumbnail, got %#v", withMetadata.Images)
	}
}

func TestOOXMLImageExtensionIsSniffed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Sniff image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "sniff-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Ext != ".png" || res.Images[0].Name != "image1.png" {
		t.Fatalf("expected sniffed png image, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "image1.png")); err != nil {
		t.Fatal(err)
	}
}

func TestOOXMLImageExtensionCaseIsNormalized(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p"><p:cSld><p:spTree><p:sp><p:txBody><a:p xmlns:a="urn:a"><a:r><a:t>Uppercase image extension</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZipBytes(t, zw, "ppt/media/image1.PNG", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "uppercase-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "image1.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected normalized png image name, got %#v", res.Images)
	}
	b, err := os.ReadFile(filepath.Join(outDir, "image1.png"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".png", b) {
		t.Fatalf("written normalized image is invalid")
	}
}

func TestOOXMLMediaPathCaseIsMatched(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Mixed case media path</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "Word/Media/Image1.PNG", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-media-case.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Image1.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected mixed-case media image to be extracted with normalized extension, got %#v", res.Images)
	}
	b, err := os.ReadFile(filepath.Join(outDir, "Image1.png"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".png", b) {
		t.Fatalf("written mixed-case media image is invalid")
	}
}

func TestVisibleOOXMLMediaPartAllowedMatchesCanonicalKeys(t *testing.T) {
	visible := map[string]bool{"word/media/visible image.png": true}
	for _, name := range []string{
		"word/media/visible image.png",
		"word/media/visible%20image.png",
		"Word/Media/Visible Image.PNG",
		"Word/Media/Visible%20Image.PNG",
		`./Word\Media\Visible Image.PNG`,
	} {
		if !visibleOOXMLMediaPartAllowed(visible, name) {
			t.Fatalf("expected visible media key to match %q", name)
		}
	}
	if visibleOOXMLMediaPartAllowed(visible, "word/media/hidden.png") {
		t.Fatal("hidden media should not match visible key")
	}
}

func TestOOXMLRelationshipEncodedMediaTargetMatchesEncodedPartName(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:r="urn:r"><w:body><w:p><w:r><w:drawing><a:blip xmlns:a="urn:a" r:embed="rIdImage"/></w:drawing></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Target="media/visible%20image.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible%20image.png", testPNG())
	addZipBytes(t, zw, "word/media/hidden%20image.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "encoded-media-target.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible image.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected only encoded visible image to be extracted, got %#v", res.Images)
	}
	b, err := os.ReadFile(filepath.Join(outDir, "visible image.png"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".png", b) {
		t.Fatal("written encoded-target image is invalid")
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden image.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden encoded image should not be written, stat err=%v", err)
	}
}

func TestOOXMLJFIFImageIsAcceptedAndTrimmed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>JFIF image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/photo.jfif", append(testJPEG(), []byte("trailing package bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "jfif-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Ext != ".jfif" || res.Images[0].Name != "photo.jfif" {
		t.Fatalf("expected JFIF image to be extracted, got %#v", res.Images)
	}
	if !bytes.Equal(res.Images[0].Data, testJPEG()) || !validImageData(".jfif", res.Images[0].Data) {
		t.Fatalf("expected JFIF data to be valid and trimmed, got len=%d", len(res.Images[0].Data))
	}
	b, err := os.ReadFile(filepath.Join(outDir, "photo.jfif"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".jfif", b) {
		t.Fatal("written JFIF image is invalid")
	}
}

func TestOOXMLJPEImageIsAcceptedAndTrimmed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>JPE image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/photo.jpe", append(testJPEG(), []byte("trailing package bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "jpe-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Ext != ".jpe" || res.Images[0].Name != "photo.jpe" {
		t.Fatalf("expected JPE image to be extracted, got %#v", res.Images)
	}
	if !bytes.Equal(res.Images[0].Data, testJPEG()) || !validImageData(".jpe", res.Images[0].Data) {
		t.Fatalf("expected JPE data to be valid and trimmed, got len=%d", len(res.Images[0].Data))
	}
	b, err := os.ReadFile(filepath.Join(outDir, "photo.jpe"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".jpe", b) {
		t.Fatal("written JPE image is invalid")
	}
}

func TestOOXMLWebPImageIsSniffedAndTrimmed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>WebP image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(testWebP(), []byte("trailing package bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "webp-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Ext != ".webp" || res.Images[0].Name != "image1.webp" {
		t.Fatalf("expected sniffed webp image, got %#v", res.Images)
	}
	if !bytes.Equal(res.Images[0].Data, testWebP()) {
		t.Fatalf("expected WebP data to be trimmed to %d bytes, got %d", len(testWebP()), len(res.Images[0].Data))
	}
	if _, err := os.Stat(filepath.Join(outDir, "image1.webp")); err != nil {
		t.Fatal(err)
	}
}

func TestOOXMLICOImageIsSniffedAndTrimmed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>ICO image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(testICO(), []byte("trailing package bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "ico-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Ext != ".ico" || res.Images[0].Name != "image1.ico" {
		t.Fatalf("expected sniffed ico image, got %#v", res.Images)
	}
	if !bytes.Equal(res.Images[0].Data, testICO()) {
		t.Fatalf("expected ICO data to be trimmed to %d bytes, got %d", len(testICO()), len(res.Images[0].Data))
	}
	if _, err := os.Stat(filepath.Join(outDir, "image1.ico")); err != nil {
		t.Fatal(err)
	}
}

func TestOOXMLCURImageIsSniffedAndTrimmed(t *testing.T) {
	cur := testCUR()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>CUR image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/cursor.bin", append(cur, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/pointer.cur", append(cur, []byte("trailing package bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "cur-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two CUR images, got %#v", res.Images)
	}
	for _, img := range res.Images {
		if img.Ext != ".cur" || !bytes.Equal(img.Data, cur) || !validImageData(".cur", img.Data) {
			t.Fatalf("expected trimmed valid CUR image, got %#v len=%d", img, len(img.Data))
		}
	}
	for _, name := range []string{"cursor.cur", "pointer.cur"} {
		b, err := os.ReadFile(filepath.Join(outDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(".cur", b) {
			t.Fatalf("written CUR image is invalid: %s", name)
		}
	}
}

func TestOOXMLISOBMFFImagesAreSniffedAndTrimmed(t *testing.T) {
	avif := testISOBMFF("avif")
	heic := testISOBMFF("heic")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>ISO image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(avif, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/photo.heic", append(heic, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/movie.mp4", testISOBMFF("mp42"))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "iso-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two ISO-BMFF images, got %#v", res.Images)
	}
	if res.Images[0].Name != "image1.avif" || res.Images[0].Ext != ".avif" || !bytes.Equal(res.Images[0].Data, avif) {
		t.Fatalf("expected sniffed trimmed AVIF, got %#v", res.Images[0])
	}
	if res.Images[1].Name != "photo.heic" || res.Images[1].Ext != ".heic" || !bytes.Equal(res.Images[1].Data, heic) {
		t.Fatalf("expected trimmed HEIC with original extension, got %#v", res.Images[1])
	}
	for _, name := range []string{"image1.avif", "photo.heic"} {
		b, err := os.ReadFile(filepath.Join(outDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("written ISO-BMFF image is invalid: %s", name)
		}
	}
}

func TestOOXMLJPEG2000ImagesAreSniffedAndTrimmed(t *testing.T) {
	jp2 := testJP2("jp2 ")
	j2k := testJ2K()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>JPEG 2000 image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(jp2, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/scan.j2k", append(j2k, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/tile.j2c", append(j2k, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/codestream.jpc", append(j2k, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/bad.jp2", testJP2("mp42"))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "jpeg2000-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 4 {
		t.Fatalf("expected four JPEG 2000 images, got %#v", res.Images)
	}
	byName := map[string]Image{}
	for _, img := range res.Images {
		byName[img.Name] = img
	}
	for _, want := range []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "image1.jp2", ext: ".jp2", data: jp2},
		{name: "scan.j2k", ext: ".j2k", data: j2k},
		{name: "tile.j2c", ext: ".j2c", data: j2k},
		{name: "codestream.jpc", ext: ".jpc", data: j2k},
	} {
		img, ok := byName[want.name]
		if !ok || img.Ext != want.ext || !bytes.Equal(img.Data, want.data) {
			t.Fatalf("expected trimmed JPEG 2000 image %+v, got %#v", want, res.Images)
		}
	}
	for _, name := range []string{"image1.jp2", "scan.j2k", "tile.j2c", "codestream.jpc"} {
		b, err := os.ReadFile(filepath.Join(outDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("written JPEG 2000 image is invalid: %s", name)
		}
	}
}

func TestOOXMLJPEGXRImagesAreSniffedAndTrimmed(t *testing.T) {
	wdp := testJPEGXR()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>JPEG XR image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/photo.wdp", append(wdp, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/image1.bin", append(wdp, []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/bad.jxr", []byte{'I', 'I', 0xbc, 0x01, 0, 0, 0, 0})
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "jpegxr-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two JPEG XR images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
	}{
		{name: "image1.jxr", ext: ".jxr"},
		{name: "photo.wdp", ext: ".wdp"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || !bytes.Equal(img.Data, wdp) || !validImageData(w.ext, img.Data) {
			t.Fatalf("expected trimmed JPEG XR image %s, got %#v", w.name, img)
		}
		b, err := os.ReadFile(filepath.Join(outDir, w.name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(w.ext, b) {
			t.Fatalf("written JPEG XR image is invalid: %s", w.name)
		}
	}
}

func TestOOXMLCompressedMetafileImagesAreDecompressed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Compressed metafiles</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.emz", gzipBytes(t, testEMF()))
	addZipBytes(t, zw, "word/media/image2.wmz", gzipBytes(t, testPlaceableWMF()))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "compressed-metafiles.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two decompressed metafiles, got %#v", res.Images)
	}
	if res.Images[0].Name != "image1.emf" || res.Images[0].Ext != ".emf" || !validImageData(".emf", res.Images[0].Data) {
		t.Fatalf("expected decompressed EMF, got %#v", res.Images[0])
	}
	if res.Images[1].Name != "image2.wmf" || res.Images[1].Ext != ".wmf" || !validImageData(".wmf", res.Images[1].Data) {
		t.Fatalf("expected decompressed WMF, got %#v", res.Images[1])
	}
	for _, name := range []string{"image1.emf", "image2.wmf"} {
		b, err := os.ReadFile(filepath.Join(outDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("written decompressed metafile is invalid: %s", name)
		}
	}
}

func TestLegacyCompressedMetafileStreamsAreDecompressed(t *testing.T) {
	streams := []oleStream{
		{Name: "legacy-image.emz", Path: "ObjectPool/legacy-image.emz", Data: gzipBytes(t, testEMF())},
		{Name: "Contents", Path: "ObjectPool/nested/legacy-vector.wmz", Data: gzipBytes(t, testPlaceableWMF())},
		{Name: "not-an-image", Path: "ObjectPool/not-an-image", Data: gzipBytes(t, testEMF())},
	}
	images := compressedMetafileStreamImages(streams, 3)
	if len(images) != 2 {
		t.Fatalf("expected two decompressed legacy metafiles, got %#v", images)
	}
	if images[0].Name != "legacy-image.emf" || images[0].Ext != ".emf" || !validImageData(".emf", images[0].Data) {
		t.Fatalf("expected decompressed legacy EMF, got %#v", images[0])
	}
	if images[1].Name != "legacy-vector.wmf" || images[1].Ext != ".wmf" || !validImageData(".wmf", images[1].Data) {
		t.Fatalf("expected decompressed legacy WMF, got %#v", images[1])
	}
}

func TestLegacyCompressedMetafileStreamsUseSniffedExtensionWhenMislabelled(t *testing.T) {
	streams := []oleStream{
		{Name: "wmf-inside-emz.emz", Path: "ObjectPool/wmf-inside-emz.emz", Data: gzipBytes(t, testPlaceableWMF())},
		{Name: "emf-inside-wmz.wmz", Path: "ObjectPool/emf-inside-wmz.wmz", Data: gzipBytes(t, testEMF())},
	}
	images := compressedImageStreamImages(streams, 4)
	if len(images) != 2 {
		t.Fatalf("expected two sniffed mislabelled compressed legacy metafiles, got %#v", images)
	}
	if images[0].Name != "wmf-inside-emz.wmf" || images[0].Ext != ".wmf" || !validImageData(".wmf", images[0].Data) {
		t.Fatalf("expected mislabelled .emz WMF to use sniffed extension, got %#v", images[0])
	}
	if images[1].Name != "emf-inside-wmz.emf" || images[1].Ext != ".emf" || !validImageData(".emf", images[1].Data) {
		t.Fatalf("expected mislabelled .wmz EMF to use sniffed extension, got %#v", images[1])
	}
}

func TestLegacyCompressedSVGStreamsAreDecompressed(t *testing.T) {
	streams := []oleStream{
		{Name: "legacy-vector.svgz", Path: "ObjectPool/legacy-vector.svgz", Data: gzipBytes(t, testSVG())},
		{Name: "Contents", Path: "ObjectPool/nested/path-vector.svgz", Data: gzipBytes(t, testSVG())},
		{Name: "not-an-image", Path: "ObjectPool/not-an-image", Data: gzipBytes(t, testSVG())},
		{Name: "bad-vector.svgz", Path: "ObjectPool/bad-vector.svgz", Data: gzipBytes(t, []byte(`<html><body>not svg</body></html>`))},
	}
	images := compressedImageStreamImages(streams, 2)
	if len(images) != 2 {
		t.Fatalf("expected two decompressed legacy SVG images, got %#v", images)
	}
	if images[0].Name != "legacy-vector.svg" || images[0].Ext != ".svg" || !validImageData(".svg", images[0].Data) {
		t.Fatalf("expected decompressed legacy SVG by stream name, got %#v", images[0])
	}
	if images[1].Name != "path-vector.svg" || images[1].Ext != ".svg" || !validImageData(".svg", images[1].Data) {
		t.Fatalf("expected decompressed legacy SVG by stream path, got %#v", images[1])
	}
}

func TestLegacyCompressedImageStreamsAreSniffedWithoutImageExtension(t *testing.T) {
	streams := []oleStream{
		{Name: "Contents", Path: "ObjectPool/Contents", Data: gzipBytes(t, testEMF())},
		{Name: "BinaryBlob.bin", Path: "ObjectPool/BinaryBlob.bin", Data: gzipBytes(t, testPlaceableWMF())},
		{Name: "VectorPayload.bin", Path: "ObjectPool/VectorPayload.bin", Data: gzipBytes(t, testSVG())},
		{Name: "GzipText", Path: "ObjectPool/GzipText", Data: gzipBytes(t, []byte("not an image"))},
	}
	images := compressedImageStreamImages(streams, 5)
	if len(images) != 3 {
		t.Fatalf("expected three sniffed compressed legacy images, got %#v", images)
	}
	for i, want := range []struct {
		name string
		ext  string
	}{
		{name: "Contents.emf", ext: ".emf"},
		{name: "BinaryBlob.wmf", ext: ".wmf"},
		{name: "VectorPayload.svg", ext: ".svg"},
	} {
		if images[i].Name != want.name || images[i].Ext != want.ext || !validImageData(want.ext, images[i].Data) {
			t.Fatalf("expected sniffed compressed legacy image %+v, got %#v", want, images[i])
		}
	}
}

func TestSVGImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	svg := []byte(`<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="120" height="20"><text>SVGIMAGESECRET</text></svg>`)
	withJunk := append(append([]byte(nil), svg...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".svg", withJunk)
	if !ok {
		t.Fatal("expected SVG with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, svg) {
		t.Fatalf("expected SVG to be trimmed to %d bytes, got %d", len(svg), len(normalized))
	}

	legacyData := append([]byte("Visible legacy text before "), withJunk...)
	legacyData = append(legacyData, []byte(" Visible legacy text after")...)
	carved := carveImages(legacyData)
	if len(carved) != 1 || carved[0].Ext != ".svg" || !bytes.Equal(carved[0].Data, svg) {
		t.Fatalf("expected one trimmed carved SVG, got %#v", carved)
	}
	masked := maskEmbeddedImagesForText(legacyData)
	for _, bad := range [][]byte{[]byte("<svg"), []byte("SVGIMAGESECRET"), []byte("</svg>")} {
		if bytes.Contains(masked, bad) {
			t.Fatalf("expected embedded SVG image bytes %q to be masked", bad)
		}
	}
	text := strings.Join(extractBinaryStrings(legacyData), "\n")
	if !strings.Contains(text, "Visible legacy text before") || !strings.Contains(text, "Visible legacy text after") {
		t.Fatalf("expected visible text around SVG to remain, got %q", text)
	}
	if strings.Contains(text, "SVGIMAGESECRET") || strings.Contains(text, "<svg") {
		t.Fatalf("extracted SVG image XML/text as document text: %q", text)
	}
}

func TestUnsafeSVGImagesAreRejected(t *testing.T) {
	for _, svg := range [][]byte{
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"><rect width="1" height="1"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><a href="javascript:alert(1)"><text>x</text></a></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><a href="java&#x0A;script:alert(1)"><text>x</text></a></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><a xlink:href="java&#x09;script:alert(1)"><text>x</text></a></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><foreignObject><body xmlns="http://www.w3.org/1999/xhtml">x</body></foreignObject></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><rect style="background-image:url(javascript:alert(1))"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><rect style="background-image:url(java&#x0D;script:alert(1))"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><rect style="width:expression(alert(1))"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><image href="data:image/svg+xml,&lt;svg onload='alert(1)'/&gt;"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><image href="data:application/xhtml+xml,&lt;html&gt;&lt;script&gt;alert(1)&lt;/script&gt;&lt;/html&gt;"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><image href="file:///C:/Users/me/hidden.png"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><image href="word/media/hidden.png"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.test/remote.png"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><image href="//example.test/remote.png"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><rect style="fill:url(file:///C:/Users/me/hidden.png)"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><rect style="fill:url(word/media/hidden.png)"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><rect style="fill:url(https://example.test/remote.png)"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><style>@import url(https://example.test/remote.css); rect{fill:red}</style><rect width="1" height="1"/></svg>`),
		[]byte(`<svg xmlns="http://www.w3.org/2000/svg"><style>rect{fill:url(file:///C:/Users/me/hidden.png)}</style><rect width="1" height="1"/></svg>`),
	} {
		if _, ok := normalizeImageData(".svg", svg); ok {
			t.Fatalf("unsafe SVG should be rejected: %s", svg)
		}
		if carved := carveImages(append([]byte("prefix"), svg...)); len(carved) != 0 {
			t.Fatalf("unsafe SVG should not be carved, got %#v", carved)
		}
	}

	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible body</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/unsafe.svg", []byte(`<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>`))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unsafe-svg.docx")
	outDir := filepath.Join(dir, "images")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 0 {
		t.Fatalf("unsafe OOXML SVG should not be extracted, got %#v", res.Images)
	}
	entries, err := os.ReadDir(outDir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 0 {
		t.Fatalf("unsafe OOXML SVG should not be written, got %d files", len(entries))
	}
}

func TestSafeSVGLinksAndStylesAreAllowed(t *testing.T) {
	svg := []byte(`<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><defs><linearGradient id="g"/></defs><style>rect{fill:url(#g)}</style><a href="https://example.test/media/info.png"><rect width="1" height="1" fill="url(#g)" style="stroke:url(#g)"/></a></svg>`)
	normalized, ok := normalizeImageData(".svg", svg)
	if !ok {
		t.Fatal("safe SVG should be accepted")
	}
	if !bytes.Contains(normalized, []byte(`https://example.test/media/info.png`)) || !bytes.Contains(normalized, []byte(`url(#g)`)) || !bytes.Contains(normalized, []byte(`<style>rect{fill:url(#g)}</style>`)) {
		t.Fatalf("safe SVG normalization lost expected content: %s", normalized)
	}
}

func TestLegacyNamedSVGStreamsAreExtracted(t *testing.T) {
	svg := testSVG()
	streams := []oleStream{
		{Name: "legacy-vector.svg", Path: "ObjectPool/legacy-vector.svg", Data: append(append([]byte(nil), svg...), []byte("trailing ole bytes")...)},
		{Name: "Contents", Path: "ObjectPool/nested/path-vector.svg", Data: svg},
		{Name: "not-an-image", Path: "ObjectPool/not-an-image", Data: svg},
		{Name: "bad.svg", Path: "ObjectPool/bad.svg", Data: []byte(`<html><body>not svg</body></html>`)},
	}
	images := legacyNamedImageStreamImages(streams, 2)
	if len(images) != 2 {
		t.Fatalf("expected two legacy SVG stream images, got %#v", images)
	}
	if images[0].Name != "legacy-vector.svg" || images[0].Ext != ".svg" || !bytes.Equal(images[0].Data, svg) {
		t.Fatalf("expected named SVG stream to be extracted, got %#v", images[0])
	}
	if images[1].Name != "path-vector.svg" || images[1].Ext != ".svg" || !bytes.Equal(images[1].Data, svg) {
		t.Fatalf("expected path-named SVG stream to be extracted, got %#v", images[1])
	}
}

func TestLegacyNamedRasterAndDIBStreamsKeepNames(t *testing.T) {
	png := testPNG()
	png2 := testPNGWithPrivateChunk([]byte("mislabelled legacy ole png"))
	jpeg := testJPEG()
	gif := testGIF()
	svg := testSVG()
	tga := testTGA()
	dib := testDIB()
	bmp, ok := dibToBMP(dib)
	if !ok {
		t.Fatal("test DIB did not convert to BMP")
	}
	images := imagesFromOLEStreams(nil, []oleStream{
		{Name: "photo.png", Path: "ObjectPool/photo.png", Data: append(append([]byte(nil), png...), []byte("trailing ole bytes")...)},
		{Name: "mislabelled.jpg", Path: "ObjectPool/mislabelled.jpg", Data: append(append([]byte(nil), png2...), []byte("trailing ole bytes")...)},
		{Name: "wrong-ext.png", Path: "ObjectPool/wrong-ext.png", Data: append(append([]byte(nil), jpeg...), []byte("trailing ole bytes")...)},
		{Name: "Contents", Path: "ObjectPool/nested/picture.jpeg", Data: append(append([]byte(nil), jpeg...), []byte("trailing ole bytes")...)},
		{Name: "preview.dib", Path: "ObjectPool/preview.dib", Data: dib},
		{Name: "preview.bin", Path: "ObjectPool/preview.bin", Data: append(append([]byte(nil), gif...), []byte("trailing ole bytes")...)},
		{Name: "Contents", Path: "ObjectPool/contents", Data: append(append([]byte(nil), svg...), []byte("trailing ole bytes")...)},
		{Name: "texture.tga", Path: "ObjectPool/texture.tga", Data: append(append([]byte(nil), tga...), []byte("trailing ole bytes")...)},
		{Name: "texture.bin", Path: "ObjectPool/texture.bin", Data: append(append([]byte(nil), tga...), []byte("trailing ole bytes")...)},
		{Name: "not-image.bin", Path: "ObjectPool/not-image.bin", Data: []byte("not an image")},
	})
	if len(images) != 7 {
		t.Fatalf("expected seven named legacy images without generic duplicates, got %#v", images)
	}
	want := []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "photo.png", ext: ".png", data: png},
		{name: "mislabelled.png", ext: ".png", data: png2},
		{name: "wrong-ext.jpg", ext: ".jpg", data: jpeg},
		{name: "preview.bmp", ext: ".bmp", data: bmp},
		{name: "preview.gif", ext: ".gif", data: gif},
		{name: "contents.svg", ext: ".svg", data: svg},
		{name: "texture.tga", ext: ".tga", data: tga},
	}
	byName := map[string]Image{}
	for _, img := range images {
		byName[img.Name] = img
	}
	for _, w := range want {
		img, ok := byName[w.name]
		if !ok || img.Ext != w.ext || !bytes.Equal(img.Data, w.data) || !validImageData(w.ext, img.Data) {
			t.Fatalf("expected named legacy image %+v, got %#v from %#v", w, img, images)
		}
	}
	for _, wrong := range []string{"mislabelled.jpg", "wrong-ext.png"} {
		if _, ok := byName[wrong]; ok {
			t.Fatalf("legacy named image kept wrong extension %q in %#v", wrong, images)
		}
	}
}

func TestLegacyNamedImageStreamsKeepNamesWithWrapperBytes(t *testing.T) {
	png := testPNGWithPrivateChunk([]byte("wrapped legacy stream png"))
	jpeg := testJPEG()
	images := imagesFromOLEStreams(nil, []oleStream{
		{Name: "wrapped.png", Path: "ObjectPool/wrapped.png", Data: append(append([]byte("OlePres000\x00preview wrapper"), png...), []byte("trailing ole bytes")...)},
		{Name: "Contents", Path: "ObjectPool/nested/wrapped photo.jpg", Data: append(append([]byte("CompObj\x00OlePres000"), jpeg...), []byte("trailing ole bytes")...)},
	})
	if len(images) != 2 {
		t.Fatalf("expected two named wrapped legacy images without generic duplicates, got %#v", images)
	}
	want := []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "wrapped.png", ext: ".png", data: png},
		{name: "wrapped photo.jpg", ext: ".jpg", data: jpeg},
	}
	byName := map[string]Image{}
	for _, img := range images {
		byName[img.Name] = img
	}
	for _, w := range want {
		img, ok := byName[w.name]
		if !ok || img.Ext != w.ext || !bytes.Equal(img.Data, w.data) || !validImageData(w.ext, img.Data) {
			t.Fatalf("expected wrapped named legacy image %+v, got %#v from %#v", w, img, images)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, want := range []string{"![wrapped](images/wrapped.png)", "![wrapped photo](images/wrapped%20photo.jpg)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing wrapped named legacy image %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "legacy-image") || strings.Contains(md, "OlePres") || strings.Contains(md, "CompObj") {
		t.Fatalf("markdown kept generic or wrapper legacy image content:\n%s", md)
	}
}

func TestLegacyNamedImageStreamsHaveUniqueNames(t *testing.T) {
	png1 := testPNG()
	png2 := testPNGWithPrivateChunk([]byte("second named stream variant"))
	images := imagesFromOLEStreams(nil, []oleStream{
		{Name: "picture.png", Path: "ObjectPool/One/picture.png", Data: png1},
		{Name: "picture.png", Path: "ObjectPool/Two/picture.png", Data: png2},
	})
	if len(images) != 2 {
		t.Fatalf("expected two named legacy stream images, got %#v", images)
	}
	if images[0].Name != "picture.png" || images[1].Name != "picture-2.png" {
		t.Fatalf("expected unique legacy image names, got %#v", images)
	}
	for _, img := range images {
		if img.Ext != ".png" || !validImageData(".png", img.Data) {
			t.Fatalf("expected valid PNG stream image, got %#v", img)
		}
	}
}

func TestLegacyNamedImageStreamsUseBackslashPathBasename(t *testing.T) {
	png := testPNG()
	namedPNG := testPNGWithPrivateChunk([]byte("named backslash stream"))
	emf := testEMF()
	namedEMF := testEMFWithPayload([]byte("named backslash vector"))
	images := imagesFromOLEStreams(nil, []oleStream{
		{Name: "Contents", Path: `ObjectPool\nested\visible picture.png`, Data: png},
		{Name: "Contents", Path: `ObjectPool\nested\legacy vector.emz`, Data: gzipBytes(t, emf)},
		{Name: `ObjectPool\named\name picture.png`, Path: "ObjectPool/ignored", Data: namedPNG},
		{Name: `ObjectPool\named\name vector.emz`, Path: "ObjectPool/ignored", Data: gzipBytes(t, namedEMF)},
	})
	if len(images) != 4 {
		t.Fatalf("expected four legacy images from backslash paths, got %#v", images)
	}
	want := []struct {
		name string
		ext  string
	}{
		{name: "visible picture.png", ext: ".png"},
		{name: "name picture.png", ext: ".png"},
		{name: "legacy vector.emf", ext: ".emf"},
		{name: "name vector.emf", ext: ".emf"},
	}
	byName := map[string]Image{}
	for _, img := range images {
		byName[img.Name] = img
	}
	for _, w := range want {
		img, ok := byName[w.name]
		if !ok || img.Ext != w.ext || !validImageData(w.ext, img.Data) {
			t.Fatalf("expected backslash path image %+v, got %#v from %#v", w, img, images)
		}
		if strings.Contains(img.Name, `\`) || strings.Contains(img.Name, "ObjectPool") || strings.Contains(img.Name, "nested") || strings.Contains(img.Name, "named") {
			t.Fatalf("legacy image name kept OLE path components: %#v", img)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, want := range []string{"![visible picture](images/visible%20picture.png)", "![legacy vector](images/legacy%20vector.emf)", "![name picture](images/name%20picture.png)", "![name vector](images/name%20vector.emf)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing backslash-path legacy image %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{`ObjectPool`, `nested`, `named`, `\`} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept legacy OLE path component %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyNamedImageStreamsDeduplicateAliasExtensionsByContent(t *testing.T) {
	jpeg := testJPEG()
	tiff := testTIFF()
	images := imagesFromOLEStreams(nil, []oleStream{
		{Name: "photo.jpg", Path: "ObjectPool/One/photo.jpg", Data: jpeg},
		{Name: "photo.jpe", Path: "ObjectPool/Two/photo.jpe", Data: jpeg},
		{Name: "photo.jfif", Path: "ObjectPool/Three/photo.jfif", Data: jpeg},
		{Name: "scan.tif", Path: "ObjectPool/One/scan.tif", Data: tiff},
		{Name: "scan.tiff", Path: "ObjectPool/Two/scan.tiff", Data: tiff},
	})
	if len(images) != 2 {
		t.Fatalf("expected alias extensions with identical content to deduplicate, got %#v", images)
	}
	want := []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "photo.jpg", ext: ".jpg", data: jpeg},
		{name: "scan.tif", ext: ".tif", data: tiff},
	}
	for i, w := range want {
		if images[i].Name != w.name || images[i].Ext != w.ext || !bytes.Equal(images[i].Data, w.data) || !validImageData(w.ext, images[i].Data) {
			t.Fatalf("expected deduplicated legacy image %+v, got %#v", w, images[i])
		}
	}
}

func TestOOXMLEPSImagesAreSniffedAndTrimmed(t *testing.T) {
	eps := testEPS()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>EPS image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(append([]byte(nil), eps...), []byte("trailing ole bytes")...))
	addZipBytes(t, zw, "word/media/image2.ps", append(append([]byte(nil), eps...), []byte("trailing ole bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "eps.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two EPS/PostScript images, got %#v", res.Images)
	}
	if res.Images[0].Name != "image1.eps" || res.Images[0].Ext != ".eps" || !bytes.Equal(res.Images[0].Data, eps) {
		t.Fatalf("expected sniffed trimmed EPS, got %#v", res.Images[0])
	}
	if res.Images[1].Name != "image2.ps" || res.Images[1].Ext != ".ps" || !bytes.Equal(res.Images[1].Data, eps) {
		t.Fatalf("expected extension-preserved PostScript, got %#v", res.Images[1])
	}
	for _, name := range []string{"image1.eps", "image2.ps"} {
		b, err := os.ReadFile(filepath.Join(outDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("written EPS image is invalid: %s len=%d", name, len(b))
		}
	}
}

func TestEPSImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	eps := testEPS()
	withJunk := append(append([]byte(nil), eps...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".eps", withJunk)
	if !ok {
		t.Fatal("expected EPS with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, eps) {
		t.Fatalf("expected EPS to be trimmed to %d bytes, got %d", len(eps), len(normalized))
	}

	carved := carveImages(append([]byte("pref"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".eps" || !bytes.Equal(carved[0].Data, eps) {
		t.Fatalf("expected one trimmed carved EPS, got %#v", carved)
	}
	masked := maskEmbeddedImagesForText(append([]byte("prefix"), withJunk...))
	for _, bad := range [][]byte{[]byte("%!PS-Adobe-"), []byte("%%BoundingBox:"), []byte("showpage")} {
		if bytes.Contains(masked, bad) {
			t.Fatalf("expected embedded EPS image bytes %q to be masked", bad)
		}
	}
	bad := append([]byte(nil), eps...)
	bad = bytes.Replace(bad, []byte("%%BoundingBox: 0 0 10 10"), []byte("%%BoundingBox: 10 10 0 0"), 1)
	if validImageData(".eps", bad) {
		t.Fatal("expected EPS with invalid bounding box to be rejected")
	}
	for _, tc := range []struct {
		name string
		data []byte
	}{
		{name: "run operator", data: testEPSWithBody("(C:/Users/me/hidden.ps) run\n")},
		{name: "file operator", data: testEPSWithBody("(C:/Users/me/hidden.bin) (r) file\n")},
		{name: "deletefile operator", data: testEPSWithBody("(C:/Users/me/hidden.bin) deletefile\n")},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if validImageData(".eps", tc.data) {
				t.Fatalf("expected EPS with unsafe file operator to be rejected")
			}
			if carved := carveImages(append([]byte("prefix"), tc.data...)); len(carved) != 0 {
				t.Fatalf("expected unsafe EPS not to be carved, got %#v", carved)
			}
		})
	}
	commentOnly := testEPSWithBody("% (C:/Users/me/hidden.ps) run\n0 0 moveto\n")
	if !validImageData(".eps", commentOnly) {
		t.Fatal("expected EPS comment mentioning run to remain valid")
	}
}

func TestLegacyNamedEPSStreamsAreExtracted(t *testing.T) {
	eps := testEPS()
	streams := []oleStream{
		{Name: "legacy-vector.eps", Path: "ObjectPool/legacy-vector.eps", Data: append(append([]byte(nil), eps...), []byte("trailing ole bytes")...)},
		{Name: "Contents", Path: "ObjectPool/nested/path-vector.ps", Data: eps},
		{Name: "not-an-image", Path: "ObjectPool/not-an-image", Data: eps},
		{Name: "bad.eps", Path: "ObjectPool/bad.eps", Data: []byte("not eps")},
	}
	images := legacyNamedImageStreamImages(streams, 2)
	if len(images) != 2 {
		t.Fatalf("expected two legacy EPS stream images, got %#v", images)
	}
	if images[0].Name != "legacy-vector.eps" || images[0].Ext != ".eps" || !bytes.Equal(images[0].Data, eps) {
		t.Fatalf("expected named EPS stream to be extracted, got %#v", images[0])
	}
	if images[1].Name != "path-vector.ps" || images[1].Ext != ".ps" || !bytes.Equal(images[1].Data, eps) {
		t.Fatalf("expected path-named PostScript stream to be extracted, got %#v", images[1])
	}
}

func TestWMFImagesAreValidatedTrimmedAndCarved(t *testing.T) {
	wmf := testPlaceableWMF()
	withJunk := append(append([]byte(nil), wmf...), 0xde, 0xad, 0xbe, 0xef)
	normalized, ok := normalizeImageData(".wmf", withJunk)
	if !ok {
		t.Fatal("expected valid WMF with trailing junk to be recognized")
	}
	if len(normalized) != len(wmf) {
		t.Fatalf("expected WMF to be trimmed to %d bytes, got %d", len(wmf), len(normalized))
	}

	bad := append([]byte(nil), wmf...)
	bad[20] ^= 0xff
	if validWMFData(bad) {
		t.Fatal("expected bad placeable WMF checksum to be rejected")
	}
	badSmallMaxRecord := append([]byte(nil), wmf...)
	binary.LittleEndian.PutUint32(badSmallMaxRecord[22+12:], 2)
	if validWMFData(badSmallMaxRecord) {
		t.Fatal("expected WMF with too-small max record to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badSmallMaxRecord...)); len(carved) != 0 {
		t.Fatalf("expected WMF with too-small max record not to be carved, got %#v", carved)
	}
	badLargeMaxRecord := append([]byte(nil), wmf...)
	binary.LittleEndian.PutUint32(badLargeMaxRecord[22+12:], 1000)
	if validWMFData(badLargeMaxRecord) {
		t.Fatal("expected WMF with max record beyond file size to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badLargeMaxRecord...)); len(carved) != 0 {
		t.Fatalf("expected WMF with oversized max record not to be carved, got %#v", carved)
	}
	badZeroRecord := testPlaceableWMFWithRecords([]wmfTestRecord{{words: 0, function: 0x0201}}, 3)
	if validWMFData(badZeroRecord) {
		t.Fatal("expected WMF with zero-length record to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badZeroRecord...)); len(carved) != 0 {
		t.Fatalf("expected WMF with zero-length record not to be carved, got %#v", carved)
	}
	badOverrunRecord := testPlaceableWMFWithRecords([]wmfTestRecord{{words: 10, function: 0x0201}}, 10)
	if validWMFData(badOverrunRecord) {
		t.Fatal("expected WMF with overrun record to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badOverrunRecord...)); len(carved) != 0 {
		t.Fatalf("expected WMF with overrun record not to be carved, got %#v", carved)
	}

	carved := carveImages(append(append([]byte("prefix"), wmf...), []byte("suffix")...))
	if len(carved) != 1 || carved[0].Ext != ".wmf" || !bytes.Equal(carved[0].Data, wmf) {
		t.Fatalf("expected one carved WMF, got %#v", carved)
	}
}

func TestDIBImagesAreWrappedAsBMPAndCarved(t *testing.T) {
	dib := testDIB()
	bmp, ok := normalizeImageData(".dib", dib)
	if !ok {
		t.Fatal("expected DIB to be wrapped as BMP")
	}
	if !bytes.Equal(bmp[:2], []byte("BM")) || !validImageData(".bmp", bmp) {
		t.Fatalf("wrapped DIB is not a valid BMP: len=%d", len(bmp))
	}

	carved := carveImages(append(append([]byte("prefix\x00\x00"), dib...), []byte("suffix")...))
	if len(carved) != 1 || carved[0].Ext != ".bmp" || !validImageData(".bmp", carved[0].Data) {
		t.Fatalf("expected one carved BMP from DIB, got %#v", carved)
	}
}

func TestBitmapCoreHeaderDIBAndBMPAreAccepted(t *testing.T) {
	dib := testCoreDIB(1, 1)
	bmp, ok := normalizeImageData(".dib", dib)
	if !ok {
		t.Fatal("expected BITMAPCOREHEADER DIB to be wrapped as BMP")
	}
	if !bytes.Equal(bmp[:2], []byte("BM")) || !validImageData(".bmp", bmp) {
		t.Fatalf("wrapped BITMAPCOREHEADER DIB is not a valid BMP: len=%d", len(bmp))
	}
	if size, ok := bmpDeclaredSize(bmp); !ok || size != len(bmp) {
		t.Fatalf("expected BITMAPCOREHEADER BMP declared size %d, got %d ok=%v", len(bmp), size, ok)
	}
}

func TestBitfieldsDIBAndBMPAreAccepted(t *testing.T) {
	tests := []struct {
		name string
		dib  []byte
	}{
		{name: "rgb565", dib: testBitfieldsDIB(2, 1, 16, 3, []uint32{0xf800, 0x07e0, 0x001f})},
		{name: "rgba8888", dib: testBitfieldsDIB(1, 1, 32, 6, []uint32{0x00ff0000, 0x0000ff00, 0x000000ff, 0xff000000})},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			bmp, ok := normalizeImageData(".dib", tc.dib)
			if !ok {
				t.Fatal("expected bitfields DIB to be wrapped as BMP")
			}
			if !bytes.Equal(bmp[:2], []byte("BM")) || !validImageData(".bmp", bmp) {
				t.Fatalf("wrapped bitfields DIB is not a valid BMP: len=%d", len(bmp))
			}
			if size, ok := bmpDeclaredSize(bmp); !ok || size != len(bmp) {
				t.Fatalf("expected bitfields BMP declared size %d, got %d ok=%v", len(bmp), size, ok)
			}
			carved := carveImages(append(append([]byte("prefix\x00\x00"), tc.dib...), []byte("suffix")...))
			if len(carved) != 1 || carved[0].Ext != ".bmp" || !validImageData(".bmp", carved[0].Data) {
				t.Fatalf("expected one carved BMP from bitfields DIB, got %#v", carved)
			}
		})
	}
}

func TestInvalidBMPLayoutIsRejected(t *testing.T) {
	dib := testDIB()
	bmp, ok := dibToBMP(dib)
	if !ok || !validImageData(".bmp", bmp) {
		t.Fatal("expected test BMP to be valid")
	}

	badOffset := append([]byte(nil), bmp...)
	binary.LittleEndian.PutUint32(badOffset[10:], 14)
	if validImageData(".bmp", badOffset) {
		t.Fatal("expected BMP with pixel data inside DIB header to be rejected")
	}
	if _, ok := normalizeImageData(".bmp", badOffset); ok {
		t.Fatal("expected invalid-offset BMP normalization to fail")
	}

	truncated := append([]byte(nil), bmp...)
	binary.LittleEndian.PutUint32(truncated[2:], uint32(len(truncated)-1))
	if validImageData(".bmp", truncated) {
		t.Fatal("expected BMP missing declared pixel bytes to be rejected")
	}

	badReserved := append([]byte(nil), bmp...)
	binary.LittleEndian.PutUint16(badReserved[6:], 1)
	if validImageData(".bmp", badReserved) {
		t.Fatal("expected BMP with non-zero reserved header fields to be rejected")
	}
}

func TestInvalidBMPBitfieldsLayoutIsRejected(t *testing.T) {
	tests := []struct {
		name string
		data func() []byte
	}{
		{
			name: "truncated masks",
			data: func() []byte {
				dib := testBitfieldsDIB(2, 1, 16, 3, []uint32{0xf800, 0x07e0, 0x001f})
				return dib[:48]
			},
		},
		{
			name: "overlapping masks",
			data: func() []byte {
				dib := testBitfieldsDIB(2, 1, 16, 3, []uint32{0xf800, 0x07e0, 0x001f})
				binary.LittleEndian.PutUint32(dib[44:], 0xf800)
				return dib
			},
		},
		{
			name: "mask outside bit depth",
			data: func() []byte {
				dib := testBitfieldsDIB(2, 1, 16, 3, []uint32{0xf800, 0x07e0, 0x001f})
				binary.LittleEndian.PutUint32(dib[40:], 0x00010000)
				return dib
			},
		},
		{
			name: "unsupported bit depth",
			data: func() []byte {
				dib := testBitfieldsDIB(2, 1, 16, 3, []uint32{0xf800, 0x07e0, 0x001f})
				binary.LittleEndian.PutUint16(dib[14:], 24)
				return dib
			},
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			dib := tc.data()
			if validImageData(".dib", dib) {
				t.Fatal("expected invalid bitfields DIB to be rejected")
			}
			if bmp, ok := dibToBMP(dib); ok || bmp != nil {
				t.Fatalf("expected invalid bitfields DIB not to wrap, got len=%d", len(bmp))
			}
		})
	}
}

func TestCarveImagesSkipsSizedImagesInsideExistingImage(t *testing.T) {
	bmp, ok := dibToBMP(testDIB())
	if !ok {
		t.Fatal("test DIB did not convert to BMP")
	}
	png := testPNGWithPrivateChunk(bmp)
	carved := carveImages(append([]byte("prefix"), append(png, []byte("suffix")...)...))
	if len(carved) != 1 {
		t.Fatalf("expected only the containing PNG to be carved, got %#v", carved)
	}
	if carved[0].Ext != ".png" || !bytes.Equal(carved[0].Data, png) || !validImageData(".png", carved[0].Data) {
		t.Fatalf("expected valid containing PNG, got %#v", carved[0])
	}
}

func TestImageByteRangesSkipsSizedImagesInsideExistingImage(t *testing.T) {
	secretBMP := []byte("PNG-PRIVATE-BMP-SECRET")
	dib := testDIBWithPayload(secretBMP)
	bmp, ok := dibToBMP(dib)
	if !ok {
		t.Fatal("test DIB did not convert to BMP")
	}
	png := testPNGWithPrivateChunk(bmp)
	data := append([]byte("VISIBLE-TEXT "), append(png, []byte(" TAIL-TEXT")...)...)
	ranges := imageByteRanges(data)
	if len(ranges) != 1 {
		t.Fatalf("expected only outer PNG image range, got %#v", ranges)
	}
	if got := data[ranges[0].start:ranges[0].end]; !bytes.Equal(got, png) {
		t.Fatalf("expected range to cover outer PNG only, got len=%d want=%d", len(got), len(png))
	}
	text := strings.Join(extractBinaryStrings(data), "\n")
	if !strings.Contains(text, "VISIBLE-TEXT") || !strings.Contains(text, "TAIL-TEXT") {
		t.Fatalf("expected surrounding text after image masking, got %q", text)
	}
	if strings.Contains(text, string(secretBMP)) {
		t.Fatalf("extracted text from nested BMP payload: %q", text)
	}
}

func TestOOXMLDIBImagesAreWrappedAsBMP(t *testing.T) {
	dib := testDIB()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>DIB image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.dib", dib)
	addZipBytes(t, zw, "word/media/image2.bin", dib)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "dib-images.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two DIB images wrapped as BMP, got %#v", res.Images)
	}
	for i, img := range res.Images {
		wantName := fmt.Sprintf("image%d.bmp", i+1)
		if img.Name != wantName || img.Ext != ".bmp" || !validImageData(".bmp", img.Data) {
			t.Fatalf("expected wrapped BMP %s, got %#v", wantName, img)
		}
		if !bytes.Equal(img.Data[:2], []byte("BM")) {
			t.Fatalf("expected BMP file header for %s", img.Name)
		}
		b, err := os.ReadFile(filepath.Join(outDir, wantName))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(".bmp", b) {
			t.Fatalf("written wrapped DIB image is invalid: %s", wantName)
		}
	}
}

func TestOOXMLMislabelledImageUsesSniffedExtension(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Mislabelled image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/mislabelled.jpg", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mislabelled-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one image, got %#v", res.Images)
	}
	if res.Images[0].Name != "mislabelled.png" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected mislabelled PNG to be written with .png extension, got %#v", res.Images[0])
	}
	written := filepath.Join(outDir, "mislabelled.png")
	b, err := os.ReadFile(written)
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".png", b) {
		t.Fatalf("written sniffed PNG is invalid: %s", written)
	}
	if _, err := os.Stat(filepath.Join(outDir, "mislabelled.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("mislabelled .jpg should not be written, stat err=%v", err)
	}
	if !strings.Contains(res.Markdown("images"), "](images/mislabelled.png)") {
		t.Fatalf("markdown should reference sniffed extension:\n%s", res.Markdown("images"))
	}
}

func TestExtractReturnedImageNamesMatchWrittenSanitizedFilenames(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Unsafe image name</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/bad:name*diagram.jpg", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unsafe-image-name.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one image, got %#v", res.Images)
	}
	if res.Images[0].Name != "bad_name_diagram.png" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected returned image metadata to match sanitized sniffed file, got %#v", res.Images[0])
	}
	written := filepath.Join(outDir, res.Images[0].Name)
	b, err := os.ReadFile(written)
	if err != nil {
		t.Fatalf("returned image name does not point to written file %s: %v", res.Images[0].Name, err)
	}
	if !validImageData(filepath.Ext(res.Images[0].Name), b) {
		t.Fatalf("written sanitized image is invalid: %s", res.Images[0].Name)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "](images/bad_name_diagram.png)") {
		t.Fatalf("markdown should reference returned written image name:\n%s", md)
	}
	for _, hidden := range []string{"bad:name", "*.jpg", ".jpg"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Images[0].Name, hidden) {
			t.Fatalf("output kept unsafe/mislabelled image name fragment %q: name=%q markdown=\n%s", hidden, res.Images[0].Name, md)
		}
	}
}

func TestSizedImagesAreMaskedBeforeTextExtraction(t *testing.T) {
	hiddenBMP := []byte("BMPIMAGESECRET")
	hiddenDIB := []byte("DIBIMAGESECRET")
	hiddenEMF := []byte("EMFIMAGESECRET")
	hiddenWMF := []byte("WMFIMAGESECRET")

	dibWithBMPSecret := testDIBWithPayload(hiddenBMP)
	bmpWithSecret, ok := dibToBMP(dibWithBMPSecret)
	if !ok || !validImageData(".bmp", bmpWithSecret) {
		t.Fatal("expected test BMP to be valid")
	}
	dibWithSecret := testDIBWithPayload(hiddenDIB)
	if !validImageData(".dib", dibWithSecret) {
		t.Fatal("expected test DIB to be valid")
	}
	emfWithSecret := testEMFWithPayload(hiddenEMF)
	if !validImageData(".emf", emfWithSecret) {
		t.Fatal("expected test EMF to be valid")
	}
	wmfWithSecret := testWMFWithPayload(hiddenWMF)
	if !validImageData(".wmf", wmfWithSecret) {
		t.Fatal("expected test WMF to be valid")
	}

	data := append([]byte("VISIBLE-TEXT "), bmpWithSecret...)
	appendAligned := func(img []byte) {
		for len(data)%4 != 0 {
			data = append(data, ' ')
		}
		data = append(data, img...)
	}
	appendAligned(dibWithSecret)
	appendAligned(emfWithSecret)
	data = append(data, ' ')
	data = append(data, wmfWithSecret...)
	masked := maskEmbeddedImagesForText(data)
	if !bytes.Contains(masked, []byte("VISIBLE-TEXT")) {
		t.Fatal("expected non-image visible text to remain")
	}
	for _, hidden := range [][]byte{hiddenBMP, hiddenDIB, hiddenEMF, hiddenWMF} {
		if bytes.Contains(masked, hidden) {
			t.Fatalf("expected image payload text %q to be masked", hidden)
		}
	}
	text := strings.Join(extractBinaryStrings(data), "\n")
	if !strings.Contains(text, "VISIBLE-TEXT") {
		t.Fatalf("expected visible text after binary extraction, got %q", text)
	}
	for _, hidden := range []string{string(hiddenBMP), string(hiddenDIB), string(hiddenEMF), string(hiddenWMF)} {
		if strings.Contains(text, hidden) {
			t.Fatalf("extracted image payload text %q from %.400q", hidden, text)
		}
	}
}

func TestEMFImagesAreValidatedTrimmedAndCarved(t *testing.T) {
	emf := testEMF()
	withJunk := append(append([]byte(nil), emf...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".emf", withJunk)
	if !ok {
		t.Fatal("expected EMF with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, emf) {
		t.Fatalf("expected EMF to be trimmed to %d bytes, got %d", len(emf), len(normalized))
	}
	carved := carveImages(append([]byte("pref"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".emf" || !bytes.Equal(carved[0].Data, emf) {
		t.Fatalf("expected one trimmed carved EMF, got %#v", carved)
	}

	badSize := append([]byte(nil), emf...)
	binary.LittleEndian.PutUint32(badSize[4:], uint32(len(emf)-1))
	if validImageData(".emf", badSize) {
		t.Fatal("expected EMF with unaligned declared size to be rejected")
	}
	if carved := carveImages(append([]byte("pref"), badSize...)); len(carved) != 0 {
		t.Fatalf("expected EMF with unaligned declared size not to be carved, got %#v", carved)
	}

	multiRecord := testEMFWithRecordPayload([]byte("EMF-SECOND-RECORD"))
	withJunk = append(append([]byte(nil), multiRecord...), []byte("trailing ole bytes")...)
	normalized, ok = normalizeImageData(".emf", withJunk)
	if !ok {
		t.Fatal("expected multi-record EMF with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, multiRecord) {
		t.Fatalf("expected multi-record EMF to be trimmed to %d bytes, got %d", len(multiRecord), len(normalized))
	}
	carved = carveImages(append([]byte("pref"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".emf" || !bytes.Equal(carved[0].Data, multiRecord) {
		t.Fatalf("expected one full multi-record EMF, got %#v", carved)
	}

	badDeclaredSize := append([]byte(nil), multiRecord...)
	binary.LittleEndian.PutUint32(badDeclaredSize[48:], uint32(len(multiRecord)-2))
	if validImageData(".emf", badDeclaredSize) {
		t.Fatal("expected EMF with unaligned nBytes to be rejected")
	}

	badRecordSize := append([]byte(nil), multiRecord...)
	headerSize := int(binary.LittleEndian.Uint32(badRecordSize[4:]))
	binary.LittleEndian.PutUint32(badRecordSize[headerSize+4:], 0)
	if validImageData(".emf", badRecordSize) {
		t.Fatal("expected EMF with zero-sized appended record to be rejected")
	}
	missingEOF := append([]byte(nil), multiRecord[:len(multiRecord)-20]...)
	binary.LittleEndian.PutUint32(missingEOF[48:], uint32(len(missingEOF)))
	binary.LittleEndian.PutUint32(missingEOF[52:], 2)
	if validImageData(".emf", missingEOF) {
		t.Fatal("expected EMF without EOF record to be rejected")
	}
	if carved := carveImages(append([]byte("pref"), missingEOF...)); len(carved) != 0 {
		t.Fatalf("expected EMF without EOF record not to be carved, got %#v", carved)
	}
}

func TestRasterImagesAreValidatedTrimmedAndCarved(t *testing.T) {
	cases := []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "png", ext: ".png", data: testPNG()},
		{name: "jpeg", ext: ".jpg", data: testJPEG()},
		{name: "gif", ext: ".gif", data: testGIF()},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			withJunk := append(append([]byte(nil), tc.data...), []byte("trailing ole bytes")...)
			normalized, ok := normalizeImageData(tc.ext, withJunk)
			if !ok {
				t.Fatalf("expected %s with trailing bytes to be recognized", tc.ext)
			}
			if !bytes.Equal(normalized, tc.data) {
				t.Fatalf("expected %s to be trimmed to %d bytes, got %d", tc.ext, len(tc.data), len(normalized))
			}

			carved := carveImages(append([]byte("prefix"), withJunk...))
			if len(carved) != 1 || carved[0].Ext != tc.ext || !bytes.Equal(carved[0].Data, tc.data) {
				t.Fatalf("expected one trimmed carved %s, got %#v", tc.ext, carved)
			}
		})
	}

	badLogicalGIF := append([]byte(nil), testGIF()...)
	binary.LittleEndian.PutUint16(badLogicalGIF[6:], 0)
	if validImageData(".gif", badLogicalGIF) {
		t.Fatal("expected GIF with zero logical screen width to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badLogicalGIF...)); len(carved) != 0 {
		t.Fatalf("expected GIF with zero logical screen width not to be carved, got %#v", carved)
	}

	badFrameGIF := append([]byte(nil), testGIF()...)
	if off := bytes.IndexByte(badFrameGIF, 0x2c); off >= 0 && off+10 <= len(badFrameGIF) {
		binary.LittleEndian.PutUint16(badFrameGIF[off+5:], 0)
	} else {
		t.Fatal("test GIF missing image descriptor")
	}
	if validImageData(".gif", badFrameGIF) {
		t.Fatal("expected GIF with zero image width to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badFrameGIF...)); len(carved) != 0 {
		t.Fatalf("expected GIF with zero image width not to be carved, got %#v", carved)
	}
}

func TestTIFFImagesAreValidatedTrimmedAndCarved(t *testing.T) {
	tiff := testTIFF()
	withJunk := append(append([]byte(nil), tiff...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".tif", withJunk)
	if !ok {
		t.Fatal("expected TIFF with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, tiff) {
		t.Fatalf("expected TIFF to be trimmed to %d bytes, got %d", len(tiff), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".tif" || !bytes.Equal(carved[0].Data, tiff) {
		t.Fatalf("expected one trimmed carved TIFF, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved TIFF is not valid: len=%d", len(carved[0].Data))
	}

	bigTIFF := testBigTIFF()
	bigWithJunk := append(append([]byte(nil), bigTIFF...), []byte("trailing ole bytes")...)
	normalizedBig, ok := normalizeImageData(".tif", bigWithJunk)
	if !ok {
		t.Fatal("expected BigTIFF with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalizedBig, bigTIFF) {
		t.Fatalf("expected BigTIFF to be trimmed to %d bytes, got %d", len(bigTIFF), len(normalizedBig))
	}
	bigCarved := carveImages(append([]byte("prefix"), bigWithJunk...))
	if len(bigCarved) != 1 || bigCarved[0].Ext != ".tif" || !bytes.Equal(bigCarved[0].Data, bigTIFF) {
		t.Fatalf("expected one trimmed carved BigTIFF, got %#v", bigCarved)
	}
	if !validImageData(bigCarved[0].Ext, bigCarved[0].Data) {
		t.Fatalf("carved BigTIFF is not valid: len=%d", len(bigCarved[0].Data))
	}
	maskedBig := maskEmbeddedImagesForText(append([]byte("prefix"), bigWithJunk...))
	if bytes.Contains(maskedBig, []byte{'I', 'I', 43, 0, 8, 0, 0, 0}) || bytes.Contains(maskedBig, []byte{0x7f}) {
		t.Fatal("expected BigTIFF payload to be masked before text extraction")
	}

	for _, tc := range []struct {
		name string
		data []byte
	}{
		{name: "missing strip byte counts", data: testTIFFWithStripTags(true, false, 1, 1)},
		{name: "missing strip offsets", data: testTIFFWithStripTags(false, true, 1, 1)},
		{name: "mismatched strip arrays", data: testTIFFWithStripTags(true, true, 2, 1)},
		{name: "zero strip offset", data: testTIFFWithEditedEntry(273, func(entry []byte) {
			binary.LittleEndian.PutUint32(entry[8:], 0)
		})},
		{name: "zero strip byte count", data: testTIFFWithEditedEntry(279, func(entry []byte) {
			binary.LittleEndian.PutUint32(entry[8:], 0)
		})},
		{name: "zero tile offset", data: testTIFFWithTileLayout(0, 1)},
		{name: "zero tile byte count", data: testTIFFWithTileLayout(110, 0)},
		{name: "missing jpeg byte count", data: testTIFFWithJPEGInterchangeTags(true, false)},
		{name: "missing jpeg offset", data: testTIFFWithJPEGInterchangeTags(false, true)},
		{name: "self-referential next ifd", data: testTIFFWithSelfReferentialNextIFD()},
		{name: "bigtiff zero strip byte count", data: testBigTIFFWithEditedEntry(279, func(entry []byte) {
			binary.LittleEndian.PutUint64(entry[12:], 0)
		})},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if validImageData(".tif", tc.data) {
				t.Fatalf("expected invalid TIFF strip layout to be rejected")
			}
			if carved := carveImages(append([]byte("prefix"), tc.data...)); len(carved) != 0 {
				t.Fatalf("expected invalid TIFF strip layout not to be carved, got %#v", carved)
			}
		})
	}
}

func TestWebPImagesAreValidatedTrimmedAndCarved(t *testing.T) {
	webp := testWebP()
	withJunk := append(append([]byte(nil), webp...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".webp", withJunk)
	if !ok {
		t.Fatal("expected WebP with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, webp) {
		t.Fatalf("expected WebP to be trimmed to %d bytes, got %d", len(webp), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".webp" || !bytes.Equal(carved[0].Data, webp) {
		t.Fatalf("expected one trimmed carved WebP, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved WebP is not valid: len=%d", len(carved[0].Data))
	}

	animated := testAnimatedWebP()
	if normalized, ok := normalizeImageData(".webp", append(animated, []byte("trailing ole bytes")...)); !ok || !bytes.Equal(normalized, animated) {
		t.Fatal("expected animated WebP with ANMF image chunk to be recognized and trimmed")
	}
}

func TestInvalidWebPChunksAreRejected(t *testing.T) {
	for _, tc := range []struct {
		name string
		data []byte
	}{
		{name: "wrong VP8X payload length five", data: testInvalidWebP("VP8X", []byte{0, 0, 0, 0, 0})},
		{name: "wrong VP8X payload length four", data: testInvalidWebP("VP8X", []byte{0, 0, 0, 0})},
		{name: "VP8X reserved flag bit", data: testInvalidWebP("VP8X", []byte{0x01, 0, 0, 0, 0, 0, 0, 0, 0, 0})},
		{name: "VP8X oversized canvas", data: testInvalidWebP("VP8X", []byte{0, 0, 0, 0, 0xff, 0xff, 0xff, 0, 0, 0})},
		{name: "ANMF without VP8X animation flag", data: testWebPWithChunks(testWebPVP8XChunk(), testWebPChunk("ANMF", testWebPANMFPayload()))},
		{name: "ANMF without VP8X chunk", data: testWebPWithChunks(testWebPChunk("ANMF", testWebPANMFPayload()))},
		{name: "ANMF oversized frame", data: testInvalidWebP("ANMF", append([]byte{0, 0, 0, 0, 0, 0, 0xff, 0xff, 0xff, 0, 0, 0, 0, 0, 0, 0}, testWebPChunk("VP8L", []byte{0x2f, 0, 0, 0, 0})...))},
		{name: "ANMF reserved flag bit", data: testInvalidWebP("ANMF", append([]byte{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x04}, testWebPChunk("VP8L", []byte{0x2f, 0, 0, 0, 0})...))},
		{name: "bad VP8L signature", data: testInvalidWebP("VP8L", []byte{0, 0, 0, 0, 0})},
		{name: "bad VP8L version bits", data: testInvalidWebP("VP8L", []byte{0x2f, 0, 0, 0, 0xe0})},
		{name: "bad keyframe VP8 signature", data: testInvalidWebP("VP8 ", []byte{0, 0, 0, 0, 0, 0, 0, 0, 0, 0})},
		{name: "non-keyframe VP8 chunk", data: testInvalidWebP("VP8 ", []byte{1, 0, 0, 0, 0, 0, 0, 0, 0, 0})},
		{name: "hidden VP8 keyframe", data: testInvalidWebP("VP8 ", []byte{0xe0, 0, 0, 0x9d, 0x01, 0x2a, 1, 0, 1, 0})},
		{name: "bad VP8 version", data: testInvalidWebP("VP8 ", []byte{0xfe, 0, 0, 0x9d, 0x01, 0x2a, 1, 0, 1, 0})},
		{name: "short VP8 first partition", data: testInvalidWebP("VP8 ", []byte{0xd0, 0, 0, 0x9d, 0x01, 0x2a, 1, 0, 1, 0})},
		{name: "overrun VP8 first partition", data: testInvalidWebP("VP8 ", []byte{0x10, 0x7d, 0, 0x9d, 0x01, 0x2a, 1, 0, 1, 0})},
		{name: "zero width VP8 keyframe", data: testInvalidWebP("VP8 ", []byte{0, 0, 0, 0x9d, 0x01, 0x2a, 0, 0, 1, 0})},
		{name: "extended metadata without image chunk", data: testWebPMetadataOnly()},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if _, ok := normalizeImageData(".webp", tc.data); ok {
				t.Fatalf("invalid WebP should be rejected")
			}
			if validImageData(".webp", tc.data) {
				t.Fatalf("invalid WebP should not be valid")
			}
			if carved := carveImages(append([]byte("prefix"), tc.data...)); len(carved) != 0 {
				t.Fatalf("invalid WebP should not be carved, got %#v", carved)
			}
		})
	}
}

func TestICOImagesAreValidatedTrimmedAndCarved(t *testing.T) {
	ico := testICO()
	withJunk := append(append([]byte(nil), ico...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".ico", withJunk)
	if !ok {
		t.Fatal("expected ICO with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, ico) {
		t.Fatalf("expected ICO to be trimmed to %d bytes, got %d", len(ico), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".ico" || !bytes.Equal(carved[0].Data, ico) {
		t.Fatalf("expected one trimmed carved ICO, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved ICO is not valid: len=%d", len(carved[0].Data))
	}
}

func TestCURImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	cur := testCUR()
	withJunk := append(append([]byte(nil), cur...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".cur", withJunk)
	if !ok {
		t.Fatal("expected CUR with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, cur) {
		t.Fatalf("expected CUR to be trimmed to %d bytes, got %d", len(cur), len(normalized))
	}

	wrapped := append([]byte("prefix"), withJunk...)
	carved := carveImages(wrapped)
	if len(carved) != 1 || carved[0].Ext != ".cur" || !bytes.Equal(carved[0].Data, cur) {
		t.Fatalf("expected one trimmed carved CUR, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved CUR is not valid: len=%d", len(carved[0].Data))
	}
	ranges := imageByteRanges(wrapped)
	if len(ranges) != 1 || ranges[0].start != len("prefix") || ranges[0].end != len("prefix")+len(cur) {
		t.Fatalf("expected CUR byte range to mask image payload, got %#v", ranges)
	}
}

func TestIconDirectoryRejectsMismatchedPayloadDimensions(t *testing.T) {
	tests := []struct {
		name string
		ext  string
		data []byte
	}{
		{name: "ico", ext: ".ico", data: testICO()},
		{name: "cur", ext: ".cur", data: testCUR()},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			bad := append([]byte(nil), tc.data...)
			bad[6] = 2
			if _, ok := normalizeImageData(tc.ext, bad); ok {
				t.Fatalf("expected %s with mismatched directory dimensions to be rejected", tc.ext)
			}
			if validImageData(tc.ext, bad) {
				t.Fatalf("expected %s with mismatched directory dimensions to be invalid", tc.ext)
			}
			for _, img := range carveImages(append([]byte("prefix"), bad...)) {
				if img.Ext == tc.ext {
					t.Fatalf("expected mismatched %s wrapper not to be carved as itself, got %#v", tc.ext, img)
				}
				if !validImageData(img.Ext, img.Data) {
					t.Fatalf("fallback carved image is invalid: %#v", img)
				}
			}
			badReserved := append([]byte(nil), tc.data...)
			badReserved[9] = 1
			if _, ok := normalizeImageData(tc.ext, badReserved); ok {
				t.Fatalf("expected %s with non-zero directory reserved byte to be rejected", tc.ext)
			}
			if validImageData(tc.ext, badReserved) {
				t.Fatalf("expected %s with non-zero directory reserved byte to be invalid", tc.ext)
			}
			for _, img := range carveImages(append([]byte("prefix"), badReserved...)) {
				if img.Ext == tc.ext {
					t.Fatalf("expected reserved-byte %s wrapper not to be carved as itself, got %#v", tc.ext, img)
				}
				if !validImageData(img.Ext, img.Data) {
					t.Fatalf("fallback carved image is invalid: %#v", img)
				}
			}
		})
	}
}

func TestIconDirectoryRejectsOverlappingPayloads(t *testing.T) {
	for _, tc := range []struct {
		name string
		ext  string
		kind uint16
	}{
		{name: "ico", ext: ".ico", kind: 1},
		{name: "cur", ext: ".cur", kind: 2},
	} {
		t.Run(tc.name, func(t *testing.T) {
			payload := testPNG()
			dirEnd := 6 + 2*16
			b := make([]byte, dirEnd+len(payload))
			binary.LittleEndian.PutUint16(b[2:], tc.kind)
			binary.LittleEndian.PutUint16(b[4:], 2)
			for i := 0; i < 2; i++ {
				entry := b[6+i*16:]
				entry[0], entry[1] = 1, 1
				binary.LittleEndian.PutUint16(entry[4:], 1)
				binary.LittleEndian.PutUint16(entry[6:], 32)
				binary.LittleEndian.PutUint32(entry[8:], uint32(len(payload)))
				binary.LittleEndian.PutUint32(entry[12:], uint32(dirEnd))
			}
			copy(b[dirEnd:], payload)
			if _, ok := normalizeImageData(tc.ext, b); ok {
				t.Fatalf("expected %s with overlapping payloads to be rejected", tc.ext)
			}
			if validImageData(tc.ext, b) {
				t.Fatalf("expected %s with overlapping payloads to be invalid", tc.ext)
			}
			for _, img := range carveImages(append([]byte("prefix"), b...)) {
				if img.Ext == tc.ext {
					t.Fatalf("expected overlapping %s wrapper not to be carved as itself, got %#v", tc.ext, img)
				}
			}
		})
	}
}

func TestICOAcceptsDIBPayloadWithMaskHeight(t *testing.T) {
	width := 1
	height := 1
	dibHeightIncludingMask := 2
	stride := ((width*24 + 31) / 32) * 4
	dib := make([]byte, 40+stride*dibHeightIncludingMask)
	binary.LittleEndian.PutUint32(dib[0:], 40)
	binary.LittleEndian.PutUint32(dib[4:], uint32(width))
	binary.LittleEndian.PutUint32(dib[8:], uint32(dibHeightIncludingMask))
	binary.LittleEndian.PutUint16(dib[12:], 1)
	binary.LittleEndian.PutUint16(dib[14:], 24)
	binary.LittleEndian.PutUint32(dib[20:], uint32(stride*dibHeightIncludingMask))
	dib[40] = 0xff

	ico := testIconDirectoryWithPayload(1, width, height, dib)
	if _, ok := normalizeImageData(".ico", ico); !ok {
		t.Fatal("expected ICO DIB payload with doubled mask height to be accepted")
	}
}

func TestICOAcceptsBitmapCoreHeaderDIBPayload(t *testing.T) {
	ico := testIconDirectoryWithPayload(1, 1, 1, testCoreDIB(1, 2))
	normalized, ok := normalizeImageData(".ico", ico)
	if !ok {
		t.Fatal("expected ICO with BITMAPCOREHEADER DIB payload to be accepted")
	}
	if !bytes.Equal(normalized, ico) || !validImageData(".ico", normalized) {
		t.Fatalf("expected valid normalized ICO with BITMAPCOREHEADER payload, len=%d", len(normalized))
	}
}

func TestOOXMLPCXImagesAreSniffedAndTrimmed(t *testing.T) {
	pcx := testPCX()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>PCX image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(append([]byte(nil), pcx...), []byte("trailing ole bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "pcx.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one PCX image, got %#v", res.Images)
	}
	if res.Images[0].Name != "image1.pcx" || res.Images[0].Ext != ".pcx" || !bytes.Equal(res.Images[0].Data, pcx) {
		t.Fatalf("expected sniffed trimmed PCX, got %#v", res.Images[0])
	}
	b, err := os.ReadFile(filepath.Join(outDir, "image1.pcx"))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(".pcx", b) {
		t.Fatalf("written PCX image is invalid: len=%d", len(b))
	}
}

func TestPCXImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	pcx := testPCX()
	withJunk := append(append([]byte(nil), pcx...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".pcx", withJunk)
	if !ok {
		t.Fatal("expected PCX with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, pcx) {
		t.Fatalf("expected PCX to be trimmed to %d bytes, got %d", len(pcx), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".pcx" || !bytes.Equal(carved[0].Data, pcx) {
		t.Fatalf("expected one trimmed carved PCX, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved PCX is not valid: len=%d", len(carved[0].Data))
	}
	masked := maskEmbeddedImagesForText(append([]byte("prefix"), withJunk...))
	if bytes.Contains(masked, pcx[128:]) {
		t.Fatal("expected embedded PCX image bytes to be masked")
	}
	bad := append([]byte(nil), pcx...)
	bad[2] = 0
	if validImageData(".pcx", bad) {
		t.Fatal("expected PCX with unsupported encoding to be rejected")
	}
	badReserved := append([]byte(nil), pcx...)
	badReserved[64] = 1
	if validImageData(".pcx", badReserved) {
		t.Fatal("expected PCX with non-zero reserved header byte to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badReserved...)); len(carved) != 0 {
		t.Fatalf("expected PCX with non-zero reserved byte not to be carved, got %#v", carved)
	}
	badPaletteInfo := append([]byte(nil), pcx...)
	binary.LittleEndian.PutUint16(badPaletteInfo[68:], 3)
	if _, ok := normalizeImageData(".pcx", badPaletteInfo); ok {
		t.Fatal("expected PCX with unsupported palette info to be rejected")
	}
	if validImageData(".pcx", badPaletteInfo) {
		t.Fatal("expected PCX with unsupported palette info to be invalid")
	}
	if carved := carveImages(append([]byte("prefix"), badPaletteInfo...)); len(carved) != 0 {
		t.Fatalf("expected PCX with unsupported palette info not to be carved, got %#v", carved)
	}
}

func TestOOXMLTGAImagesAreAcceptedTrimmedAndSniffed(t *testing.T) {
	tga := testTGA()
	sniffedTGA := testTGARLE()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>TGA image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.tga", append(append([]byte(nil), tga...), []byte("trailing package bytes")...))
	addZipBytes(t, zw, "word/media/sniffed.bin", append(append([]byte(nil), sniffedTGA...), []byte("trailing package bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "tga.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected extension-backed and sniffed TGA images, got %#v", res.Images)
	}
	if res.Images[0].Name != "image1.tga" || res.Images[0].Ext != ".tga" || !bytes.Equal(res.Images[0].Data, tga) {
		t.Fatalf("expected trimmed TGA, got %#v", res.Images[0])
	}
	if res.Images[1].Name != "sniffed.tga" || res.Images[1].Ext != ".tga" || !bytes.Equal(res.Images[1].Data, sniffedTGA) {
		t.Fatalf("expected sniffed trimmed TGA, got %#v", res.Images[1])
	}
	for _, want := range []struct {
		name string
		data []byte
	}{
		{name: "image1.tga", data: tga},
		{name: "sniffed.tga", data: sniffedTGA},
	} {
		b, err := os.ReadFile(filepath.Join(outDir, want.name))
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(b, want.data) || !validImageData(".tga", b) {
			t.Fatalf("written TGA image %s is invalid or untrimmed: len=%d", want.name, len(b))
		}
	}
}

func TestTGAImagesAreValidatedAndTrimmed(t *testing.T) {
	tga := testTGA()
	withJunk := append(append([]byte(nil), tga...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".tga", withJunk)
	if !ok {
		t.Fatal("expected TGA with trailing bytes to be recognized by extension")
	}
	if !bytes.Equal(normalized, tga) {
		t.Fatalf("expected TGA to be trimmed to %d bytes, got %d", len(tga), len(normalized))
	}
	if imageExt(tga) != ".tga" {
		t.Fatalf("expected TGA to be sniffed without extension, got %s", imageExt(tga))
	}
	rle := testTGARLE()
	if normalized, ok := normalizeImageData(".tga", append(append([]byte(nil), rle...), []byte("tail")...)); !ok || !bytes.Equal(normalized, rle) {
		t.Fatalf("expected RLE TGA to be trimmed and valid, ok=%v len=%d", ok, len(normalized))
	}
	bad := append([]byte(nil), tga...)
	bad[2] = 0
	if validImageData(".tga", bad) {
		t.Fatal("expected unsupported TGA image type to be rejected")
	}
}

func TestLegacyExtensionlessTGAStreamsAreExtracted(t *testing.T) {
	tga := testTGA()
	rle := testTGARLE()
	streams := []oleStream{
		{Name: "Contents", Path: "ObjectPool/Contents", Data: append(append([]byte(nil), tga...), []byte("trailing ole bytes")...)},
		{Name: "Preview.bin", Path: "ObjectPool/Preview.bin", Data: append(append([]byte(nil), rle...), []byte("trailing ole bytes")...)},
	}
	images := legacyNamedImageStreamImages(streams, 3)
	if len(images) != 2 {
		t.Fatalf("expected two extensionless legacy TGA stream images, got %#v", images)
	}
	for i, want := range []struct {
		name string
		data []byte
	}{
		{name: "Contents.tga", data: tga},
		{name: "Preview.tga", data: rle},
	} {
		if images[i].Name != want.name || images[i].Ext != ".tga" || !bytes.Equal(images[i].Data, want.data) || !validImageData(".tga", images[i].Data) {
			t.Fatalf("expected extensionless TGA stream %+v, got %#v", want, images[i])
		}
	}
}

func TestInvalidTGAHeadersAreRejected(t *testing.T) {
	tests := []struct {
		name string
		edit func([]byte) []byte
	}{
		{name: "bad color map type", edit: func(b []byte) []byte { b[1] = 2; return b }},
		{name: "truecolor with color map first index", edit: func(b []byte) []byte { binary.LittleEndian.PutUint16(b[3:], 1); return b }},
		{name: "truecolor with color map length", edit: func(b []byte) []byte { binary.LittleEndian.PutUint16(b[5:], 1); return b }},
		{name: "truecolor with color map entry bits", edit: func(b []byte) []byte { b[7] = 24; return b }},
		{name: "color map index overflow", edit: func(b []byte) []byte {
			b[1] = 1
			b[2] = 1
			binary.LittleEndian.PutUint16(b[3:], 65535)
			binary.LittleEndian.PutUint16(b[5:], 2)
			b[7] = 24
			b[16] = 8
			return append(b[:18], append([]byte{0, 0, 0, 0, 0, 0}, b[18:]...)...)
		}},
		{name: "interleaved descriptor", edit: func(b []byte) []byte { b[17] = 0x40; return b }},
		{name: "rle packet overruns pixels", edit: func(b []byte) []byte { b[18] = 0x82; return b }},
		{name: "rle packet missing payload", edit: func(b []byte) []byte { return b[:19] }},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			bad := testTGARLE()
			if !strings.Contains(tc.name, "rle") {
				bad = testTGA()
			}
			bad = tc.edit(bad)
			if _, ok := normalizeImageData(".tga", bad); ok {
				t.Fatalf("expected invalid TGA to be rejected")
			}
			if validImageData(".tga", bad) {
				t.Fatalf("expected invalid TGA not to validate")
			}
		})
	}
}

func TestOOXMLPICTImagesAreSniffedAndTrimmed(t *testing.T) {
	pict := testPICT(false)
	wrapped := testPICT(true)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>PICT image</w:t></w:r></w:p></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.bin", append(append([]byte(nil), pict...), []byte("trailing ole bytes")...))
	addZipBytes(t, zw, "word/media/image2.pct", append(append([]byte(nil), wrapped...), []byte("trailing ole bytes")...))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "pict.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two PICT images, got %#v", res.Images)
	}
	if res.Images[0].Name != "image1.pict" || res.Images[0].Ext != ".pict" || !bytes.Equal(res.Images[0].Data, pict) {
		t.Fatalf("expected sniffed trimmed PICT, got %#v", res.Images[0])
	}
	if res.Images[1].Name != "image2.pct" || res.Images[1].Ext != ".pct" || !bytes.Equal(res.Images[1].Data, wrapped) {
		t.Fatalf("expected extension-preserved wrapped PICT, got %#v", res.Images[1])
	}
	for _, name := range []string{"image1.pict", "image2.pct"} {
		b, err := os.ReadFile(filepath.Join(outDir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("written PICT image is invalid: %s len=%d", name, len(b))
		}
	}
}

func TestPICTImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	pict := testPICT(true)
	withJunk := append(append([]byte(nil), pict...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".pict", withJunk)
	if !ok {
		t.Fatal("expected PICT with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, pict) {
		t.Fatalf("expected PICT to be trimmed to %d bytes, got %d", len(pict), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".pict" || !bytes.Equal(carved[0].Data, pict) {
		t.Fatalf("expected one trimmed carved PICT, got %#v", carved)
	}
	masked := maskEmbeddedImagesForText(append([]byte("prefix"), withJunk...))
	if bytes.Contains(masked, []byte("AB")) || bytes.Contains(masked, []byte{0x00, 0x11, 0x02, 0xff}) {
		t.Fatal("expected embedded PICT image bytes to be masked")
	}
	bad := append([]byte(nil), pict...)
	bad[len(bad)-1] = 0
	if validImageData(".pict", bad) {
		t.Fatal("expected PICT without end opcode to be rejected")
	}
}

func TestPICTImagesRejectVersionOpcodeFoundInBIFFData(t *testing.T) {
	// This is shaped like a plausible frame and has the PICT v2 version opcode,
	// but lacks the required extended PICT header.  It occurs in a text-only
	// legacy XLS workbook and must not become a phantom image.
	bad := make([]byte, 40)
	binary.BigEndian.PutUint16(bad[6:], 888)
	binary.BigEndian.PutUint16(bad[8:], 645)
	copy(bad[10:], []byte{0x00, 0x11, 0x02, 0xff, 0x0c, 0x00, 0xff, 0xff})
	bad[len(bad)-2] = 0x00
	bad[len(bad)-1] = 0xff
	if validPICTData(bad) || len(carveImages(bad)) != 0 {
		t.Fatal("invalid PICT-like BIFF data was accepted as an image")
	}
}

func TestLegacyNamedPICTStreamsAreExtracted(t *testing.T) {
	pict := testPICT(true)
	streams := []oleStream{
		{Name: "legacy-image.pct", Path: "ObjectPool/legacy-image.pct", Data: append(append([]byte(nil), pict...), []byte("trailing ole bytes")...)},
		{Name: "Contents", Path: "ObjectPool/nested/path-vector.pict", Data: pict},
		{Name: "not-an-image", Path: "ObjectPool/not-an-image", Data: pict},
		{Name: "bad.pict", Path: "ObjectPool/bad.pict", Data: []byte("not pict")},
	}
	images := legacyNamedImageStreamImages(streams, 1)
	if len(images) != 2 {
		t.Fatalf("expected two legacy PICT stream images, got %#v", images)
	}
	if images[0].Name != "legacy-image.pct" || images[0].Ext != ".pct" || !bytes.Equal(images[0].Data, pict) {
		t.Fatalf("expected named PCT stream to be extracted, got %#v", images[0])
	}
	if images[1].Name != "path-vector.pict" || images[1].Ext != ".pict" || !bytes.Equal(images[1].Data, pict) {
		t.Fatalf("expected path-named PICT stream to be extracted, got %#v", images[1])
	}
}

func TestISOBMFFImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	heif := testISOBMFF("mif1")
	withJunk := append(append([]byte(nil), heif...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".heif", withJunk)
	if !ok {
		t.Fatal("expected HEIF with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, heif) {
		t.Fatalf("expected HEIF to be trimmed to %d bytes, got %d", len(heif), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".heif" || !bytes.Equal(carved[0].Data, heif) {
		t.Fatalf("expected one trimmed carved HEIF, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved HEIF is not valid: len=%d", len(carved[0].Data))
	}
	masked := maskEmbeddedImagesForText(append([]byte("prefix"), withJunk...))
	if bytes.Contains(masked, []byte("ftyp")) || bytes.Contains(masked, []byte("mif1")) {
		t.Fatalf("expected embedded ISO-BMFF image bytes to be masked")
	}
	if validImageData(".avif", heif) {
		t.Fatal("expected HEIF payload to be rejected as AVIF")
	}
	if validImageData(".heic", testISOBMFF("mp42")) {
		t.Fatal("expected non-image ISO-BMFF payload to be rejected")
	}
	withoutMeta := testISOBMFFWithoutMeta("heic")
	if validImageData(".heic", withoutMeta) {
		t.Fatal("expected ISO-BMFF image without meta box to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), withoutMeta...)); len(carved) != 0 {
		t.Fatalf("expected ISO-BMFF without meta box not to be carved, got %#v", carved)
	}
	emptyMeta := testISOBMFFWithEmptyMeta("heic")
	if validImageData(".heic", emptyMeta) {
		t.Fatal("expected ISO-BMFF image with empty meta box to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), emptyMeta...)); len(carved) != 0 {
		t.Fatalf("expected ISO-BMFF with empty meta box not to be carved, got %#v", carved)
	}
}

func TestJPEG2000ImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	jpx := testJP2("jpx ")
	withJunk := append(append([]byte(nil), jpx...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".jpx", withJunk)
	if !ok {
		t.Fatal("expected JPX with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, jpx) {
		t.Fatalf("expected JPX to be trimmed to %d bytes, got %d", len(jpx), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".jpx" || !bytes.Equal(carved[0].Data, jpx) {
		t.Fatalf("expected one trimmed carved JPX, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved JPX is not valid: len=%d", len(carved[0].Data))
	}
	masked := maskEmbeddedImagesForText(append([]byte("prefix"), withJunk...))
	if bytes.Contains(masked, []byte("jP  ")) || bytes.Contains(masked, []byte("jp2c")) {
		t.Fatal("expected embedded JPEG 2000 bytes to be masked")
	}
	if validImageData(".jp2", jpx) {
		t.Fatal("expected JPX payload to be rejected as JP2")
	}
	if validImageData(".jp2", testJP2("mp42")) {
		t.Fatal("expected non-JPEG-2000 brand to be rejected")
	}
	for _, tc := range []struct {
		name string
		data []byte
	}{
		{name: "missing image header", data: testJP2WithBoxes("jp2 ", makeJP2Box("jp2h", nil), makeJP2Box("jp2c", testJ2K()))},
		{name: "zero image header width", data: testJP2WithBoxes("jp2 ", testJP2HeaderBox(0, 1, 1), makeJP2Box("jp2c", testJ2K()))},
		{name: "bad codestream payload", data: testJP2WithBoxes("jp2 ", testJP2HeaderBox(1, 1, 1), makeJP2Box("jp2c", []byte{0xff, 0x4f, 0, 0, 0, 0}))},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if validImageData(".jp2", tc.data) {
				t.Fatalf("expected malformed JPEG 2000 container to be rejected")
			}
			for _, img := range carveImages(append([]byte("prefix"), tc.data...)) {
				if img.Ext == ".jp2" {
					t.Fatalf("expected malformed JPEG 2000 container not to be carved as JP2, got %#v", img)
				}
				if !validImageData(img.Ext, img.Data) {
					t.Fatalf("fallback carved image is invalid: %#v", img)
				}
			}
		})
	}
	badSIZLength := testJ2K()
	binary.BigEndian.PutUint16(badSIZLength[4:], 42)
	badSIZLength = append(append([]byte(nil), badSIZLength[:45]...), append([]byte{0}, badSIZLength[45:]...)...)
	if validImageData(".j2k", badSIZLength) {
		t.Fatal("expected JPEG 2000 codestream with oversized SIZ segment to be rejected")
	}
	if carved := carveImages(append([]byte("prefix"), badSIZLength...)); len(carved) != 0 {
		t.Fatalf("expected malformed JPEG 2000 codestream not to be carved, got %#v", carved)
	}
	zeroSampling := testJ2K()
	zeroSampling[43] = 0
	if validImageData(".j2k", zeroSampling) {
		t.Fatal("expected JPEG 2000 codestream with zero component sampling to be rejected")
	}
	invalidPrecision := testJ2K()
	invalidPrecision[42] = 38
	if validImageData(".j2k", invalidPrecision) {
		t.Fatal("expected JPEG 2000 codestream with invalid component precision to be rejected")
	}
}

func TestJPEGXRImagesAreValidatedTrimmedMaskedAndCarved(t *testing.T) {
	jxr := testJPEGXR()
	withJunk := append(append([]byte(nil), jxr...), []byte("trailing ole bytes")...)
	normalized, ok := normalizeImageData(".jxr", withJunk)
	if !ok {
		t.Fatal("expected JPEG XR with trailing bytes to be recognized")
	}
	if !bytes.Equal(normalized, jxr) {
		t.Fatalf("expected JPEG XR to be trimmed to %d bytes, got %d", len(jxr), len(normalized))
	}

	carved := carveImages(append([]byte("prefix"), withJunk...))
	if len(carved) != 1 || carved[0].Ext != ".jxr" || !bytes.Equal(carved[0].Data, jxr) {
		t.Fatalf("expected one trimmed carved JPEG XR, got %#v", carved)
	}
	if !validImageData(carved[0].Ext, carved[0].Data) {
		t.Fatalf("carved JPEG XR is not valid: len=%d", len(carved[0].Data))
	}
	masked := maskEmbeddedImagesForText(append([]byte("prefix"), withJunk...))
	if bytes.Contains(masked, []byte{0xbc, 0x01}) || bytes.Contains(masked, []byte{0x01, 0x02, 0x03, 0x04}) {
		t.Fatal("expected embedded JPEG XR bytes to be masked")
	}
	if validImageData(".wdp", []byte{'I', 'I', 0xbc, 0x01, 0, 0, 0, 0}) {
		t.Fatal("expected malformed JPEG XR payload to be rejected")
	}
	for _, tc := range []struct {
		name string
		data []byte
	}{
		{name: "missing image offset", data: testJPEGXRWithEditedEntry(273, func(entry []byte) {
			binary.LittleEndian.PutUint16(entry[0:], 305)
		})},
		{name: "missing byte count", data: testJPEGXRWithEditedEntry(279, func(entry []byte) {
			binary.LittleEndian.PutUint16(entry[0:], 305)
		})},
		{name: "zero byte count", data: testJPEGXRWithEditedEntry(279, func(entry []byte) {
			binary.LittleEndian.PutUint32(entry[8:], 0)
		})},
		{name: "self-referential next ifd", data: testJPEGXRWithSelfReferentialNextIFD()},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if validImageData(".jxr", tc.data) {
				t.Fatalf("expected JPEG XR without valid image payload to be rejected")
			}
			if carved := carveImages(append([]byte("prefix"), tc.data...)); len(carved) != 0 {
				t.Fatalf("expected invalid JPEG XR not to be carved, got %#v", carved)
			}
		})
	}
}

func TestExtractedSampleImagesAreValid(t *testing.T) {
	samples, err := filepath.Glob(filepath.Join("testdata", "samples", "*.*"))
	if err != nil {
		t.Fatal(err)
	}
	for _, sample := range samples {
		ext := strings.ToLower(filepath.Ext(sample))
		if !map[string]bool{".doc": true, ".docx": true, ".ppt": true, ".pptx": true, ".xls": true, ".xlsx": true}[ext] {
			continue
		}
		t.Run(filepath.Base(sample), func(t *testing.T) {
			dir := t.TempDir()
			res, err := Extract(sample, Options{ImageDir: dir})
			if err != nil {
				t.Fatal(err)
			}
			for _, img := range res.Images {
				if !validImageData(img.Ext, img.Data) {
					t.Fatalf("invalid image returned: %s %s len=%d", img.Name, img.Ext, len(img.Data))
				}
			}
			entries, err := os.ReadDir(dir)
			if err != nil {
				t.Fatal(err)
			}
			if len(entries) != len(res.Images) {
				t.Fatalf("expected %d written image(s), got %d", len(res.Images), len(entries))
			}
			for _, entry := range entries {
				if entry.IsDir() {
					t.Fatalf("unexpected directory in image output: %s", entry.Name())
				}
				b, err := os.ReadFile(filepath.Join(dir, entry.Name()))
				if err != nil {
					t.Fatal(err)
				}
				ext := strings.ToLower(filepath.Ext(entry.Name()))
				if !validImageData(ext, b) {
					t.Fatalf("invalid written image: %s len=%d", entry.Name(), len(b))
				}
			}
		})
	}
}

func TestWriteImagesKeepsDuplicateNames(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "duplicate.png", Ext: ".png", Data: testPNG()},
		{Name: "duplicate.png", Ext: ".png", Data: testPNG()},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"duplicate.png", "duplicate-2.png"} {
		b, err := os.ReadFile(filepath.Join(dir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("written image is not valid: %s", name)
		}
	}
}

func TestWriteImagesSanitizesUnsafeFilenames(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: `bad:name*?<>|".png`, Ext: ".png", Data: testPNG()},
		{Name: "CON.png", Ext: ".png", Data: testPNG()},
		{Name: " trailing. ", Ext: ".png", Data: testPNG()},
		{Name: "line\nbreak.png", Ext: ".png", Data: testPNG()},
		{Name: "\u200d.png", Ext: ".png", Data: testPNG()},
		{Name: "ObjectPool/nested/path picture.png", Ext: ".png", Data: testPNG()},
		{Name: `ObjectPool\named\backslash picture.png`, Ext: ".png", Data: testPNG()},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	want := []string{
		"bad_name______.png",
		"CON_.png",
		"trailing.png",
		"line_break.png",
		"image.png",
		"path picture.png",
		"backslash picture.png",
	}
	for _, name := range want {
		b, err := os.ReadFile(filepath.Join(dir, name))
		if err != nil {
			t.Fatalf("missing sanitized image %s: %v", name, err)
		}
		if !validImageData(".png", b) {
			t.Fatalf("written sanitized image is invalid: %s", name)
		}
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != len(want) {
		t.Fatalf("expected %d sanitized images, got %d", len(want), len(entries))
	}
	md := (&Result{Images: images}).Markdown("images")
	if !strings.Contains(md, "![image](images/image.png)") {
		t.Fatalf("markdown should reference fallback image filename for extension-only sanitized name:\n%s", md)
	}
	if strings.Contains(md, "png.png") {
		t.Fatalf("markdown should not use duplicated extension-only filename:\n%s", md)
	}
	for _, want := range []string{"![path picture](images/path%20picture.png)", "![backslash picture](images/backslash%20picture.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown should use path basename %q:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"ObjectPool", "nested", "named", `\`} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept image path component %q in:\n%s", hidden, md)
		}
	}
}

func TestWriteImagesSkipsInvalidImages(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "valid.png", Ext: ".png", Data: testPNG()},
		{Name: "broken.png", Ext: ".png", Data: []byte("not a png")},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	if b, err := os.ReadFile(filepath.Join(dir, "valid.png")); err != nil || !validImageData(".png", b) {
		t.Fatalf("valid image was not written correctly: len=%d err=%v", len(b), err)
	}
	if _, err := os.Stat(filepath.Join(dir, "broken.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("invalid image should not be written, stat err=%v", err)
	}
}

func TestMarkdownSkipsInvalidImages(t *testing.T) {
	res := &Result{
		Text:               "Visible body\nBroken Alt\nVisible footer",
		StructuredMarkdown: "## Document\n\nVisible body",
		Images: []Image{
			{Name: "valid.png", Alt: "Valid Alt", Ext: ".png", Data: testPNG()},
			{Name: "broken.png", Alt: "Broken Alt", Ext: ".png", Data: []byte("not a png")},
		},
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Valid Alt](images/valid.png)") {
		t.Fatalf("markdown missing valid image:\n%s", md)
	}
	if strings.Contains(md, "broken.png") || strings.Contains(md, "![Broken Alt]") {
		t.Fatalf("markdown referenced invalid image:\n%s", md)
	}
	if !strings.Contains(md, "Broken Alt") || !strings.Contains(md, "Visible footer") {
		t.Fatalf("invalid image alt should not suppress visible text backfill:\n%s", md)
	}
}

func TestOutputImagesNormalizeSniffedExtensions(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "mislabelled.jpg", Ext: ".jpg", Data: testPNG()},
		{Name: "noext", Data: testJPEG()},
		{Name: "vector-noext", Data: testEMF()},
		{Name: "drawing-noext", Data: testPlaceableWMF()},
		{Name: "dib-noext", Data: testDIB()},
		{Name: "vector.emz", Ext: ".emz", Data: gzipBytes(t, testEMF())},
		{Name: "drawing.wmz", Ext: ".wmz", Data: gzipBytes(t, testPlaceableWMF())},
		{Name: "icon.svgz", Ext: ".svgz", Data: gzipBytes(t, testSVG())},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name string
		ext  string
	}{
		{name: "mislabelled.png", ext: ".png"},
		{name: "noext.jpg", ext: ".jpg"},
		{name: "vector-noext.emf", ext: ".emf"},
		{name: "drawing-noext.wmf", ext: ".wmf"},
		{name: "dib-noext.bmp", ext: ".bmp"},
		{name: "vector.emf", ext: ".emf"},
		{name: "drawing.wmf", ext: ".wmf"},
		{name: "icon.svg", ext: ".svg"},
	} {
		b, err := os.ReadFile(filepath.Join(dir, tc.name))
		if err != nil {
			t.Fatalf("missing normalized image %s: %v", tc.name, err)
		}
		if !validImageData(tc.ext, b) {
			t.Fatalf("normalized image is not valid: %s", tc.name)
		}
	}
	for _, wrong := range []string{"mislabelled.jpg", "noext", "vector-noext", "drawing-noext", "dib-noext", "vector.emz", "drawing.wmz", "icon.svgz"} {
		if _, err := os.Stat(filepath.Join(dir, wrong)); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("wrong image output name %s should not exist, stat err=%v", wrong, err)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, want := range []string{"![mislabelled](images/mislabelled.png)", "![noext](images/noext.jpg)", "![vector-noext](images/vector-noext.emf)", "![drawing-noext](images/drawing-noext.wmf)", "![dib-noext](images/dib-noext.bmp)", "![vector](images/vector.emf)", "![drawing](images/drawing.wmf)", "![icon](images/icon.svg)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing normalized image reference %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "mislabelled.jpg") || strings.Contains(md, "](images/noext)") || strings.Contains(md, "](images/vector-noext)") || strings.Contains(md, "](images/drawing-noext)") || strings.Contains(md, "](images/dib-noext)") || strings.Contains(md, ".emz") || strings.Contains(md, ".wmz") || strings.Contains(md, ".svgz") {
		t.Fatalf("markdown kept unnormalized image reference:\n%s", md)
	}
}

func TestOutputImagesPreserveCompatibleSniffedExtensionAliases(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "photo.jpeg", Data: testJPEG()},
		{Name: "scan.tiff", Data: testTIFF()},
		{Name: "picture.wdp", Data: testJPEGXR()},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name string
		ext  string
	}{
		{name: "photo.jpeg", ext: ".jpeg"},
		{name: "scan.tiff", ext: ".tiff"},
		{name: "picture.wdp", ext: ".wdp"},
	} {
		b, err := os.ReadFile(filepath.Join(dir, tc.name))
		if err != nil {
			t.Fatalf("missing alias-preserved image %s: %v", tc.name, err)
		}
		if !validImageData(tc.ext, b) {
			t.Fatalf("alias-preserved image is invalid: %s", tc.name)
		}
	}
	for _, wrong := range []string{"photo.jpg", "scan.tif", "picture.jxr"} {
		if _, err := os.Stat(filepath.Join(dir, wrong)); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("canonicalized alias image %s should not exist, stat err=%v", wrong, err)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, want := range []string{"![photo](images/photo.jpeg)", "![scan](images/scan.tiff)", "![picture](images/picture.wdp)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing alias-preserved image reference %q in:\n%s", want, md)
		}
	}
	for _, wrong := range []string{"images/photo.jpg)", "images/scan.tif)", "images/picture.jxr)"} {
		if strings.Contains(md, wrong) {
			t.Fatalf("markdown canonicalized compatible alias %q in:\n%s", wrong, md)
		}
	}
}

func TestMarkdownReferencesWrittenImageNames(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "duplicate.png", Ext: ".png", Data: testPNG()},
		{Name: "duplicate.png", Ext: ".png", Data: testPNG()},
		{Name: "space name.png", Ext: ".png", Data: testPNG()},
		{Name: `bad:name*?<>|".png`, Ext: ".png", Data: testPNG()},
		{Name: "hash#[100%].png", Alt: "Hash Percent Brackets", Ext: ".png", Data: testPNG()},
		{Name: "paren (draft).png", Alt: "Paren Draft", Ext: ".png", Data: testPNG()},
		{Name: "https://cdn.example.test/media/collide.jpg?token=hidden", Alt: "Remote Collide", Ext: ".jpg", Data: testPNG()},
		{Name: `C:\Users\me\Pictures\collide.png`, Alt: "Local Collide", Ext: ".png", Data: testPNG()},
		{Name: "collide.jpg", Alt: "Sniffed Collide", Ext: ".jpg", Data: testPNG()},
	}
	res := &Result{
		Text:   "Visible heading\n\nVisible body",
		Images: images,
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"Visible heading\n\nVisible body",
		"## Images",
		"![duplicate](images/duplicate.png)",
		"![duplicate-2](images/duplicate-2.png)",
		"![space name](images/space%20name.png)",
		"![bad name](images/bad_name______.png)",
		"![Hash Percent Brackets](images/hash%23%5B100%25%5D.png)",
		"![Paren Draft](images/paren%20%28draft%29.png)",
		"![Remote Collide](images/collide.png)",
		"![Local Collide](images/collide-2.png)",
		"![Sniffed Collide](images/collide-3.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing %q in:\n%s", want, md)
		}
	}
	mdURL := res.Markdown("https://cdn.example.test/assets folder")
	for _, want := range []string{
		"![space name](https://cdn.example.test/assets%20folder/space%20name.png)",
		"![Hash Percent Brackets](https://cdn.example.test/assets%20folder/hash%23%5B100%25%5D.png)",
		"![Paren Draft](https://cdn.example.test/assets%20folder/paren%20%28draft%29.png)",
	} {
		if !strings.Contains(mdURL, want) {
			t.Fatalf("markdown URL target missing %q in:\n%s", want, mdURL)
		}
	}
	if strings.Contains(mdURL, "https%3A") {
		t.Fatalf("markdown URL target escaped scheme in:\n%s", mdURL)
	}
	mdURLTrailingSlash := res.Markdown("https://cdn.example.test/assets/")
	if !strings.Contains(mdURLTrailingSlash, "![Hash Percent Brackets](https://cdn.example.test/assets/hash%23%5B100%25%5D.png)") {
		t.Fatalf("markdown URL target with trailing slash is wrong:\n%s", mdURLTrailingSlash)
	}
	mdURLSigned := res.Markdown("https://cdn.example.test/assets folder?token=a%2Fb#signed section")
	for _, want := range []string{
		"![space name](https://cdn.example.test/assets%20folder/space%20name.png?token=a%2Fb#signed%20section)",
		"![Paren Draft](https://cdn.example.test/assets%20folder/paren%20%28draft%29.png?token=a%2Fb#signed%20section)",
	} {
		if !strings.Contains(mdURLSigned, want) {
			t.Fatalf("markdown signed URL target missing %q in:\n%s", want, mdURLSigned)
		}
	}
	if strings.Contains(mdURLSigned, "%3Ftoken") || strings.Contains(mdURLSigned, "token=a%2Fb/") {
		t.Fatalf("markdown signed URL target appended image after query or escaped query separator:\n%s", mdURLSigned)
	}
	mdURLEncodedPath := res.Markdown("https://cdn.example.test/assets%2Fbucket/base%20folder?token=a%2Fb#signed%2Fsection")
	for _, want := range []string{
		"![space name](https://cdn.example.test/assets%2Fbucket/base%20folder/space%20name.png?token=a%2Fb#signed%2Fsection)",
		"![Paren Draft](https://cdn.example.test/assets%2Fbucket/base%20folder/paren%20%28draft%29.png?token=a%2Fb#signed%2Fsection)",
	} {
		if !strings.Contains(mdURLEncodedPath, want) {
			t.Fatalf("markdown encoded URL path target missing %q in:\n%s", want, mdURLEncodedPath)
		}
	}
	if strings.Contains(mdURLEncodedPath, "assets/bucket") || strings.Contains(mdURLEncodedPath, "base%2520folder") {
		t.Fatalf("markdown encoded URL path was decoded or double-encoded:\n%s", mdURLEncodedPath)
	}
	mdLocalEncodedPath := res.Markdown("images%20out")
	for _, want := range []string{
		"![space name](images%20out/space%20name.png)",
		"![Hash Percent Brackets](images%20out/hash%23%5B100%25%5D.png)",
		"![Paren Draft](images%20out/paren%20%28draft%29.png)",
	} {
		if !strings.Contains(mdLocalEncodedPath, want) {
			t.Fatalf("markdown encoded local path target missing %q in:\n%s", want, mdLocalEncodedPath)
		}
	}
	if strings.Contains(mdLocalEncodedPath, "images%2520out") {
		t.Fatalf("markdown encoded local path was double-encoded:\n%s", mdLocalEncodedPath)
	}
	mdLocalBadPercent := res.Markdown("images%zz")
	if !strings.Contains(mdLocalBadPercent, "![space name](images%25zz/space%20name.png)") {
		t.Fatalf("markdown bad-percent local path was not escaped safely:\n%s", mdLocalBadPercent)
	}
	mdWindowsBackslashPath := res.Markdown(`C:\Users\me\images`)
	if !strings.Contains(mdWindowsBackslashPath, "![space name](C:/Users/me/images/space%20name.png)") {
		t.Fatalf("markdown Windows local path should preserve drive colon and slash-normalize:\n%s", mdWindowsBackslashPath)
	}
	if strings.Contains(mdWindowsBackslashPath, "C%3A") || strings.Contains(mdWindowsBackslashPath, `C:\`) {
		t.Fatalf("markdown Windows local path was over-escaped or not slash-normalized:\n%s", mdWindowsBackslashPath)
	}
	mdWindowsSlashPath := res.Markdown("D:/shared images")
	if !strings.Contains(mdWindowsSlashPath, "![space name](D:/shared%20images/space%20name.png)") {
		t.Fatalf("markdown Windows slash path should preserve drive colon and escape spaces:\n%s", mdWindowsSlashPath)
	}
	mdWindowsUNCPath := res.Markdown(`\\server\share images`)
	if !strings.Contains(mdWindowsUNCPath, "![space name](//server/share%20images/space%20name.png)") {
		t.Fatalf("markdown Windows UNC path should be slash-normalized and escape spaces:\n%s", mdWindowsUNCPath)
	}
	if strings.Contains(mdWindowsUNCPath, `\`) {
		t.Fatalf("markdown Windows UNC path should not keep backslashes:\n%s", mdWindowsUNCPath)
	}
	imageOnly := (&Result{Images: images[:1]}).Markdown("images")
	if !strings.HasPrefix(imageOnly, "## Images\n\n![duplicate](images/duplicate.png)") {
		t.Fatalf("image-only markdown should start with an Images section, got:\n%s", imageOnly)
	}
	for _, name := range []string{"duplicate.png", "duplicate-2.png", "space name.png", "bad_name______.png", "hash#[100%].png", "paren (draft).png", "collide.png", "collide-2.png", "collide-3.png"} {
		if _, err := os.Stat(filepath.Join(dir, name)); err != nil {
			t.Fatalf("missing written image referenced by markdown %s: %v", name, err)
		}
	}
}

func TestImageFilenamesDropInvisibleFormatControls(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "weird\u200b\u202ename.png", Ext: ".png", Data: testPNG()},
		{Name: "space\u00a0name.png", Ext: ".png", Data: testPNG()},
		{Name: "\ufeffhidden.png", Ext: ".png", Data: testPNG()},
		{Name: "variant\ufe0f\u034f\U000e0100name.png", Ext: ".png", Data: testPNG()},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"weirdname.png", "space name.png", "hidden.png", "variantname.png"} {
		b, err := os.ReadFile(filepath.Join(dir, name))
		if err != nil {
			t.Fatalf("missing sanitized image %s: %v", name, err)
		}
		if !validImageData(".png", b) {
			t.Fatalf("sanitized image is not valid: %s", name)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, want := range []string{
		"![weirdname](images/weirdname.png)",
		"![space name](images/space%20name.png)",
		"![hidden](images/hidden.png)",
		"![variantname](images/variantname.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing sanitized image reference %q in:\n%s", want, md)
		}
	}
	for _, r := range []rune{'\u200b', '\u202e', '\ufeff', '\u00a0', '\ufe0f', '\u034f', '\U000e0100'} {
		if strings.ContainsRune(md, r) {
			t.Fatalf("markdown kept invisible/ambiguous filename rune %U in:\n%s", r, md)
		}
	}
}

func TestImageFilenamesUseBaseNameForInternalOfficePaths(t *testing.T) {
	dir := t.TempDir()
	images := []Image{
		{Name: "word/media/image1.png", Ext: ".png", Data: testPNG()},
		{Name: "ppt/media/image1.png", Ext: ".png", Data: testPNG()},
		{Name: `xl\media\chart image.png`, Ext: ".png", Data: testPNG()},
		{Name: "word/media/encoded%20segment.png", Ext: ".png", Data: testPNG()},
		{Name: "word%2Fmedia%2Fencoded%20image.png", Ext: ".png", Data: testPNG()},
		{Name: "ppt%252Fmedia%252Fdouble%2520encoded.png", Ext: ".png", Data: testPNG()},
		{Name: "word/media/visible.png?download=1#section", Ext: ".png", Data: testPNG()},
		{Name: `xl\media\sheet image.png#anchor`, Ext: ".png", Data: testPNG()},
		{Name: "pack://application:,,,/word/media/packed%20image.png", Ext: ".png", Data: testPNG()},
		{Name: "file:///C:/Users/me/Pictures/local%20image.png", Ext: ".png", Data: testPNG()},
		{Name: `D:\shared\media\drive image.png`, Ext: ".png", Data: testPNG()},
		{Name: `\\server\share\unc image.png`, Ext: ".png", Data: testPNG()},
		{Name: "ms-word:ofe|u|file:///C:/Users/me/Pictures/office%20link.png", Ext: ".png", Data: testPNG()},
		{Name: "https://cdn.example.test/assets/remote%20image.png?token=hidden#frag", Ext: ".png", Data: testPNG()},
		{Name: "ftp://files.example.test/scans/scan%20remote.tif", Ext: ".tif", Data: testTIFF()},
		{Name: "cid:image001.png@01DBCAFE", Ext: ".png", Data: testPNG()},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"image1.png", "image1-2.png", "chart image.png", "encoded segment.png", "encoded image.png", "double encoded.png", "visible.png", "sheet image.png", "packed image.png", "local image.png", "drive image.png", "unc image.png", "office link.png", "remote image.png", "scan remote.tif", "image001.png"} {
		b, err := os.ReadFile(filepath.Join(dir, name))
		if err != nil {
			t.Fatalf("missing basename image %s: %v", name, err)
		}
		if !validImageData(filepath.Ext(name), b) {
			t.Fatalf("basename image is not valid: %s", name)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, want := range []string{
		"![image1](images/image1.png)",
		"![image1-2](images/image1-2.png)",
		"![chart image](images/chart%20image.png)",
		"![encoded segment](images/encoded%20segment.png)",
		"![encoded image](images/encoded%20image.png)",
		"![double encoded](images/double%20encoded.png)",
		"![visible](images/visible.png)",
		"![sheet image](images/sheet%20image.png)",
		"![packed image](images/packed%20image.png)",
		"![local image](images/local%20image.png)",
		"![drive image](images/drive%20image.png)",
		"![unc image](images/unc%20image.png)",
		"![office link](images/office%20link.png)",
		"![remote image](images/remote%20image.png)",
		"![scan remote](images/scan%20remote.tif)",
		"![image001](images/image001.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing basename image reference %q in:\n%s", want, md)
		}
	}
	for _, hiddenTail := range []string{"download=1", "#anchor", "#section", "pack://", "file://", "ms-word:", "Users", "Pictures", "server", "share", "https://", "ftp://", "cdn.example.test", "files.example.test", "token=hidden", "#frag", "cid:", "01DBCAFE"} {
		if strings.Contains(md, hiddenTail) {
			t.Fatalf("markdown leaked internal image target suffix %q in:\n%s", hiddenTail, md)
		}
	}
	for _, hidden := range []string{"word", "ppt", "xl", "media", "word_media", "ppt_media", "xl_media", "word%2Fmedia", "ppt%252Fmedia", "encoded%2520segment", "application", "C%3A", "D%3A"} {
		if strings.Contains(md, hidden+"_") || strings.Contains(md, hidden+"/") || strings.Contains(md, hidden+"\\") || strings.Contains(md, hidden) {
			t.Fatalf("markdown leaked internal image path marker %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownBackfillSkipsSanitizedImageAlts(t *testing.T) {
	res := &Result{
		Text:               "Visible body\nbad name\nimage-002\nmislabelled.png\n",
		StructuredMarkdown: "## Document\n\nVisible body",
		Images: []Image{
			{Name: `bad:name*?<>|".png`, Ext: ".png", Data: testPNG()},
			{Ext: ".png", Data: testPNG()},
			{Name: "mislabelled.jpg", Ext: ".png", Data: testPNG()},
		},
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Visible body", "## Images", "![bad name](images/bad_name______.png)", "![image-002](images/image-002.png)", "![mislabelled](images/mislabelled.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Additional Text") || strings.Count(md, "bad name") != 1 || strings.Count(md, "image-002") != 2 || strings.Count(md, "mislabelled.png") != 1 {
		t.Fatalf("markdown should not backfill sanitized image alt text:\n%s", md)
	}
}

func TestMarkdownBackfillSkipsInternalOfficePartReferences(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"docProps/app.xml",
			"customXml/item1.xml",
			"[Content_Types].xml",
			"word/_rels/document.xml.rels",
			"word/media/image1.png",
			"word%2Fmedia%2Fencoded%20image.png",
			"word%252Fmedia%252Fdouble%2520encoded.png",
			"&lt;ppt%2Fmedia%2Fentity%20encoded.png&gt;",
			"media/standalone.tga",
			"../media/relative.png",
			"media%2Fencoded%20texture.tga",
			"file://server/share/hidden.docx",
		}, "\n"),
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should not backfill internal Office package references:\n%s", md)
	}
	for _, hidden := range []string{"docProps/app.xml", "customXml/item1.xml", "[Content_Types].xml", "word/_rels/document.xml.rels", "word/media/image1.png", "word%2Fmedia%2Fencoded%20image.png", "word%252Fmedia%252Fdouble%2520encoded.png", "ppt%2Fmedia%2Fentity%20encoded.png", "media/standalone.tga", "../media/relative.png", "media%2Fencoded%20texture.tga", "file://server/share/hidden.docx"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept internal Office package reference %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownBackfillSkipsMIMEHeaderMetadata(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"Visible appendix",
			`Content-Disposition: inline; filename="hidden.png"`,
			`Content-Description: "hidden preview"`,
			`Content-Base: file:///C:/Users/me/hidden/`,
			`Content-Transfer-Encoding: base64`,
			`MIME-Version: 1.0`,
		}, "\n"),
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Additional Text") || !strings.Contains(md, "Visible appendix") {
		t.Fatalf("markdown should backfill visible appendix only:\n%s", md)
	}
	for _, hidden := range []string{"Content-Disposition", "filename=", "Content-Description", "hidden preview", "Content-Base", "file:///C:/Users/me/hidden/", "Content-Transfer-Encoding", "base64", "MIME-Version"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown backfill kept MIME metadata %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownBackfillDeduplicatesRepeatedMissingText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"Repeated appendix",
			"Repeated appendix",
			"Unique appendix",
			"Repeated appendix",
		}, "\n"),
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should include missing appendix text:\n%s", md)
	}
	if strings.Count(md, "Repeated appendix") != 1 {
		t.Fatalf("markdown should backfill repeated missing text once:\n%s", md)
	}
	if strings.Count(md, "Unique appendix") != 1 {
		t.Fatalf("markdown should keep unique missing text:\n%s", md)
	}
}

func TestMarkdownBackfillSkipsOfficeXMLMetadataReferences(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`,
			`http://schemas.openxmlformats.org/drawingml/2006/main`,
			`https://schemas.microsoft.com/office/word/2010/wordml`,
			`mc:Ignorable="w14 wp14"`,
			`xsi:schemaLocation="http://schemas.openxmlformats.org/wordprocessingml/2006/main wordprocessingml.xsd"`,
			`ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`,
			`PartName="/word/document.xml"`,
			"https://example.test/visible-user-link",
		}, "\n"),
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Additional Text") || !strings.Contains(md, "https://example.test/visible-user-link") {
		t.Fatalf("markdown should still backfill ordinary visible links:\n%s", md)
	}
	for _, hidden := range []string{"xmlns:w", "schemas.openxmlformats.org", "schemas.microsoft.com/office", "mc:Ignorable", "schemaLocation", "w14 wp14", "ContentType", "application/vnd.openxmlformats", "PartName", "/word/document.xml"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown backfill kept Office XML metadata %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownBackfillKeepsVisibleLocalPaths(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"C:\\Reports\\Q1",
			"D:/shared reports/Q2",
			"file:///C:/Users/me/hidden.docx",
			"word/media/image1.png",
		}, "\n"),
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Additional Text", "C:\\Reports\\Q1", "D:/shared reports/Q2"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown backfill missing visible local path %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"file:///C:/Users/me/hidden.docx", "word/media/image1.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown backfill kept hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownBackfillIgnoresLinkTargetsWhenCheckingCoverage(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"Quarterly report",
			"Linked visible text",
			"Chart visible alt",
			"Title paren visible",
			"Download target",
			"Nested target text",
			"Title paren target",
			"Pipe | Value",
			"Q1",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"Visible body",
			"",
			"![Chart](images/Quarterly report.png)",
			"![Chart visible alt](images/chart(foo).png \"Chart title\")",
			"[Linked visible text](https://example.test/report(foo) \"Report title\")",
			"[Title paren visible](https://example.test/Title paren target \"Report (draft\")",
			"[download](https://example.test/Download target)",
			"[Nested link](https://example.test/Nested target text(foo))",
			"",
			"| Name | Value |",
			"| --- | --- |",
			"| Pipe \\| Value | Present |",
			"",
			"Q10 Forecast",
		}, "\n"),
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Additional Text") || !strings.Contains(md, "Quarterly report") || !strings.Contains(md, "Download target") || !strings.Contains(md, "Nested target text") || !strings.Contains(md, "Title paren target") || !strings.Contains(md, "Q1") {
		t.Fatalf("markdown should backfill text that appears only in link/image targets:\n%s", md)
	}
	if strings.Count(md, "Linked visible text") != 1 || strings.Count(md, "Chart visible alt") != 1 || strings.Count(md, "Title paren visible") != 1 {
		t.Fatalf("markdown should treat link/image labels with parenthesized targets as visible coverage:\n%s", md)
	}
	if strings.Count(md, "Pipe \\| Value") != 1 || strings.Contains(md, "Pipe | Value\n") {
		t.Fatalf("markdown should treat escaped visible table text as already covered:\n%s", md)
	}
}

func TestMarkdownBackfillTreatsReferenceLinkLabelsAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"Reference visible",
			"Collapsed visible",
			"Shortcut visible",
			"Reference image alt",
			"Reference target only",
			"Plain bracket visible",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"Visible body",
			"[Reference visible][report-ref]",
			"[Collapsed visible][]",
			"[Shortcut visible]",
			"![Reference image alt][image-ref]",
			"[Plain bracket visible][missing-ref]",
			"",
			"[report-ref]: https://example.test/Reference target only \"Report title\"",
			"[Collapsed visible]: https://example.test/collapsed",
			"[Shortcut visible]: https://example.test/shortcut",
			"[image-ref]: images/reference.png",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Count(md, "Reference visible") != 1 || strings.Count(md, "Collapsed visible") != 2 || strings.Count(md, "Shortcut visible") != 2 || strings.Count(md, "Reference image alt") != 1 {
		t.Fatalf("markdown should treat defined reference link labels as visible coverage:\n%s", md)
	}
	if !strings.Contains(md, "## Additional Text") || !strings.Contains(md, "Reference target only") {
		t.Fatalf("markdown should still backfill text that appears only in reference targets:\n%s", md)
	}
	if strings.Count(md, "Plain bracket visible") != 2 {
		t.Fatalf("markdown should not treat undefined reference syntax as visible label-only coverage:\n%s", md)
	}
}

func TestMarkdownBackfillIgnoresReferenceDefinitionsInsideNonVisibleBlocks(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Fenced visible",
			"Indented visible",
			"Comment visible",
			"HTML visible",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"[Fenced visible][fenced-ref]",
			"[Indented visible][indented-ref]",
			"[Comment visible][comment-ref]",
			"[HTML visible][html-ref]",
			"",
			"```",
			"[fenced-ref]: https://example.test/fenced",
			"```",
			"    [indented-ref]: https://example.test/indented",
			"<!--",
			"[comment-ref]: https://example.test/comment",
			"-->",
			"<div>",
			"[html-ref]: https://example.test/html",
			"</div>",
		}, "\n"),
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should backfill labels whose only definitions are in non-visible blocks:\n%s", md)
	}
	for _, want := range []string{"Fenced visible", "Indented visible", "Comment visible", "HTML visible"} {
		if strings.Count(md, want) != 2 {
			t.Fatalf("markdown should not treat non-visible block reference definitions as coverage for %q:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsHeadingsAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible Sheet",
			"Closed Heading",
			"Language C#",
			"# Escaped heading marker",
			"Visible body",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Visible Sheet",
			"### Closed Heading ###",
			"Language C#",
			"\\# Escaped heading marker",
			"",
			"Visible body",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat heading text as visible coverage:\n%s", md)
	}
	for _, want := range []string{"## Visible Sheet", "### Closed Heading ###", "Language C#", "\\# Escaped heading marker"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing heading content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsListItemsAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible bullet",
			"Visible numbered",
			"Visible checked task",
			"Visible unchecked task",
			"- Escaped dash stays visible",
			"+ Escaped plus stays visible",
			"1. Escaped numbered stays visible",
			"Escaped parens (visible) and bang!",
			"Dash - kept",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"- Visible bullet",
			"1. Visible numbered",
			"- [x] Visible checked task",
			"- [ ] Visible unchecked task",
			"\\- Escaped dash stays visible",
			"\\+ Escaped plus stays visible",
			"1\\. Escaped numbered stays visible",
			"Escaped parens \\(visible\\) and bang\\!",
			"Dash - kept",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat list item markers as visible coverage:\n%s", md)
	}
	for _, want := range []string{"- Visible bullet", "1. Visible numbered", "- [x] Visible checked task", "- [ ] Visible unchecked task", "\\- Escaped dash stays visible", "\\+ Escaped plus stays visible", "1\\. Escaped numbered stays visible", "Escaped parens \\(visible\\) and bang\\!", "Dash - kept"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing list content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownTextPreservesNestedListIndentation(t *testing.T) {
	got := markdownText("  - Nested bullet\n    1. Nested numbered\nPlain paragraph")
	for _, want := range []string{"  - Nested bullet", "    1. Nested numbered", "Plain paragraph"} {
		if !strings.Contains(got, want) {
			t.Fatalf("markdown text missing %q in:\n%s", want, got)
		}
	}
}

func TestMarkdownParagraphListLevelReadsDrawingMLLevel(t *testing.T) {
	dec := xml.NewDecoder(strings.NewReader(`<a:pPr xmlns:a="urn:a" lvl="2"/>`))
	tok, err := dec.Token()
	if err != nil {
		t.Fatal(err)
	}
	start, ok := tok.(xml.StartElement)
	if !ok {
		t.Fatalf("expected start element, got %T", tok)
	}
	if got := markdownParagraphListLevel(start); got != 2 {
		t.Fatalf("expected DrawingML level 2, got %d", got)
	}
}

func TestMarkdownBackfillTreatsBlockquotesAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible quote",
			"Nested quote",
			"Plain body",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"> Visible quote",
			">> Nested quote",
			"",
			"Plain body",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat blockquote markers as visible coverage:\n%s", md)
	}
	for _, want := range []string{"> Visible quote", ">> Nested quote", "Plain body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing blockquote content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsFootnoteDefinitionsAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"Footnote visible text",
			"Body with footnote marker",
			"Table footnote text",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"Visible body",
			"Body with footnote marker[^2]",
			"",
			"| Note |",
			"| --- |",
			"| Table footnote text[^3] |",
			"",
			"[^1]: Footnote visible text",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat footnote definitions and references as visible coverage:\n%s", md)
	}
	if strings.Count(md, "Footnote visible text") != 1 {
		t.Fatalf("markdown should not duplicate footnote definition text:\n%s", md)
	}
	for _, want := range []string{"Body with footnote marker[^2]", "| Table footnote text[^3] |"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing footnote-marked visible text %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsLegacyStandaloneNoteMarkersAsStructuredText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Body[footnote] keeps inline marker",
			"[footnote] Visible footnote text",
			"[comment] Visible comment text",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"Body[footnote] keeps inline marker",
			"",
			"## Footnotes and Endnotes",
			"",
			"Visible footnote text",
			"",
			"## Comments",
			"",
			"Visible comment text",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should not backfill legacy standalone note/comment markers:\n%s", md)
	}
	for _, want := range []string{"Body[footnote] keeps inline marker", "Visible footnote text", "Visible comment text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing structured legacy note content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsInlineFormattingAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Bold value",
			"Italic value",
			"Code value",
			"Struck value",
			"Task value",
			"Quoted code",
			"Cell value",
			"Sentence with bold and code values",
			"Sentence with struck and italic values",
			"snake_case literal",
			"Escaped *literal* markers",
			"Mixed cell value with bold content",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"**Bold value**",
			"*Italic value*",
			"`Code value`",
			"~~Struck value~~",
			"- [x] **Task value**",
			"> `Quoted code`",
			"Sentence with **bold** and `code` values",
			"Sentence with ~~struck~~ and *italic* values",
			"snake_case literal",
			"Escaped \\*literal\\* markers",
			"",
			"| Field | Mixed |",
			"| --- | --- |",
			"| **Cell value** | Mixed cell value with **bold** content |",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat inline formatting wrappers as visible coverage:\n%s", md)
	}
	for _, want := range []string{"**Bold value**", "*Italic value*", "`Code value`", "~~Struck value~~", "- [x] **Task value**", "> `Quoted code`", "Sentence with **bold** and `code` values", "Sentence with ~~struck~~ and *italic* values", "snake_case literal", "Escaped \\*literal\\* markers", "| **Cell value** | Mixed cell value with **bold** content |"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing formatted content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsHTMLTagsAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"HTML visible text",
			"Decoded & visible",
			"Cell HTML",
			"Cell & Entity",
			"2 < 3 > 1",
			"https://example.test/report?q=visible",
			"visible.person@example.test",
			"https://example.test/cell",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"<div>HTML visible text</div>",
			"<strong>Decoded &amp; visible</strong>",
			"2 < 3 > 1",
			"<https://example.test/report?q=visible>",
			"<visible.person@example.test>",
			"",
			"| Field | Entity | Link |",
			"| --- | --- | --- |",
			"| <em>Cell HTML</em> | Cell &amp; Entity | <https://example.test/cell> |",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat HTML tags/entities as visible coverage:\n%s", md)
	}
	for _, want := range []string{"<div>HTML visible text</div>", "<strong>Decoded &amp; visible</strong>", "<https://example.test/report?q=visible>", "<visible.person@example.test>", "| <em>Cell HTML</em> | Cell &amp; Entity | <https://example.test/cell> |", "2 < 3 > 1"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing HTML/visible content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsLiteralMarkupExamplesAsCovered(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			`<rss version="0.91">`,
			`<event id="Att1915">`,
			`<date calendar="Gregorian" era="AD" xml:lang="en">`,
			`<infinite do loop>`,
			"f",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Presentation",
			"",
			`<rss version="0.91">`,
			`<event id="Att1915">`,
			`<date calendar="Gregorian" era="AD" xml:lang="en">`,
			`<infinite do loop>`,
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should not backfill covered literal markup or lone ASCII fragments:\n%s", md)
	}
	for _, want := range []string{`<rss version="0.91">`, `<event id="Att1915">`, `<date calendar="Gregorian" era="AD" xml:lang="en">`, `<infinite do loop>`} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing literal markup example %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsHTMLBlockTextAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Block HTML visible",
			"Nested HTML visible",
			"Plain after block",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			"<div class=\"note\">",
			"Block HTML visible",
			"<span>Nested HTML visible</span>",
			"</div>",
			"",
			"Plain after block",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat visible text inside HTML blocks as coverage:\n%s", md)
	}
	for _, want := range []string{"<div class=\"note\">", "Block HTML visible", "<span>Nested HTML visible</span>", "</div>", "Plain after block"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing HTML block content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsTableCellBreaksAsVisibleText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible Sheet",
			"First line",
			"Second line",
			"Pipe | Value",
			"Literal<br>Text",
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Visible Sheet",
			"",
			"| Notes | Literal |",
			"| --- | --- |",
			"| First line<br>Second line<br>Pipe \\| Value | Literal<br>Text |",
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat table cell <br> segments as visible coverage:\n%s", md)
	}
	for _, want := range []string{"First line<br>Second line<br>Pipe \\| Value", "Literal<br>Text", "## Visible Sheet"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing expected structured content %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownBackfillTreatsHardLineBreakMarkersAsControlText(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Hard break first",
			"Hard break second",
			"Cell hard first",
			"Cell hard second",
			`C:\`,
		}, "\n"),
		StructuredMarkdown: strings.Join([]string{
			"## Document",
			"",
			`Hard break first\`,
			"Hard break second  ",
			`C:\`,
			"",
			"| Notes |",
			"| --- |",
			`| Cell hard first\<br>Cell hard second |`,
		}, "\n"),
	}
	md := res.Markdown("images")
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("markdown should treat hard line break markers as control text:\n%s", md)
	}
	for _, want := range []string{`Hard break first\`, "Hard break second  ", `C:\`, `Cell hard first\<br>Cell hard second`} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing hard-break content %q in:\n%s", want, md)
		}
	}
}

func TestCleanVisibleTextDropsInternalOfficePartReferences(t *testing.T) {
	got := cleanVisibleText(strings.Join([]string{
		"Visible body",
		"rId7",
		"(rId8)",
		"docProps/core.xml",
		"customXml/itemProps1.xml",
		"[Content_Types].xml",
		"xl/_rels/workbook.xml.rels",
		"ppt/media/image1.png",
		"ppt%2Fmedia%2Fencoded%20image.png",
		"media/standalone.tga",
		"word/media/query.png?version=1",
		"media/fragment.tga#preview",
		"media/trailing-paren.png)",
		"media/trailing-bracket.jpg]",
		"media/trailing-angle.webp>",
		"media/vector.emz",
		"media/vector.wmz",
		"../media/relative.png",
		`Target="../media/relationship.png"`,
		`Target="..%2Fmedia%2Fencoded-relative.png"/>`,
		`r:embed="rId9"`,
		`Id="rId11"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`,
		`TargetMode="External"`,
		`ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`,
		`PartName="/word/document.xml"`,
		"media%2Fencoded%20texture.tga",
		"ppt%2Fmedia%2Fencoded%20query.png%3Fcache%3D1",
		"word%252Fmedia%252Fdouble%2520encoded.png",
		`"word\media\quoted.png"`,
		`<xl\_rels\workbook.xml.rels>`,
		"([Content_Types].xml)",
		"[ppt%2Fmedia%2Fwrapped%20image.png]",
		"%5Bppt%252Fmedia%252Fwrapped%2520double.png%5D",
		"&lt;word%2Fmedia%2Fentity%20image.png&gt;",
		"&amp;lt;xl%252F_rels%252Fworkbook.xml.rels&amp;gt;",
		"&quot;docProps/core.xml&quot;",
		"{customXml/item1.xml}.",
		"“docProps/core.xml”",
		"“docProps/app.xml”",
		"‘word/media/smart.png’",
		"「ppt/media/corner.png」",
		"『xl/_rels/workbook.xml.rels』",
		"《customXml/item2.xml》",
		"（word/media/fullwidth.png）",
		"【ppt/media/bracket.png】",
		"«xl/media/photo.jpg»",
		"ppt/media/trailing.png,",
		"pack://application:,,,/word/media/packed.png",
		"opc://package/xl/_rels/workbook.xml.rels",
		"zip://container/ppt/slides/slide1.xml",
		"ms-appx:///customXml/item1.xml",
		"'docProps/core.xml';",
		"file://server/share/hidden.docx",
		`file:\C:\Users\me\hidden.docx`,
		"file:/C:/Users/me/hidden-single-slash.docx",
		"ms-word:ofe|u|file:///C:/Users/me/hidden.docx",
		"ms-excel:ofv|u|file://server/share/hidden.xlsx",
		"Visible footer",
	}, "\n"))
	if !strings.Contains(got, "Visible body") || !strings.Contains(got, "Visible footer") {
		t.Fatalf("cleanVisibleText dropped visible text: %q", got)
	}
	for _, hidden := range []string{"rId7", "rId8", "rId9", "rId11", "relationships/image", "TargetMode", "External", "ContentType", "PartName", "application/vnd.openxmlformats", "/word/document.xml", "docProps/core.xml", "customXml/itemProps1.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels", "ppt/media/image1.png", "ppt%2Fmedia%2Fencoded%20image.png", "media/standalone.tga", "word/media/query.png", "media/fragment.tga", "media/trailing-paren.png", "media/trailing-bracket.jpg", "media/trailing-angle.webp", "media/vector.emz", "media/vector.wmz", "../media/relative.png", "../media/relationship.png", "..%2Fmedia%2Fencoded-relative.png", "media%2Fencoded%20texture.tga", "ppt%2Fmedia%2Fencoded%20query.png", "word%252Fmedia%252Fdouble%2520encoded.png", `word\media\quoted.png`, `xl\_rels\workbook.xml.rels`, "ppt%2Fmedia%2Fwrapped%20image.png", "ppt%252Fmedia%252Fwrapped%2520double.png", "word%2Fmedia%2Fentity%20image.png", "xl%252F_rels%252Fworkbook.xml.rels", "customXml/item1.xml", "docProps/app.xml", "word/media/smart.png", "ppt/media/corner.png", "customXml/item2.xml", "word/media/fullwidth.png", "ppt/media/bracket.png", "xl/media/photo.jpg", "ppt/media/trailing.png", "pack://application:,,,/word/media/packed.png", "opc://package/xl/_rels/workbook.xml.rels", "zip://container/ppt/slides/slide1.xml", "ms-appx:///customXml/item1.xml", "file://server/share/hidden.docx", `file:\C:\Users\me\hidden.docx`, "file:/C:/Users/me/hidden-single-slash.docx", "ms-word:ofe|u|file:///C:/Users/me/hidden.docx", "ms-excel:ofv|u|file://server/share/hidden.xlsx"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanVisibleText kept internal Office package reference %q in %q", hidden, got)
		}
	}
	for _, visible := range []string{"Visible body", "Visible footer"} {
		if !strings.Contains(got, visible) {
			t.Fatalf("cleanVisibleText dropped visible text %q in %q", visible, got)
		}
	}
}

func TestCleanVisibleTextStripsInlineHiddenOfficeReferences(t *testing.T) {
	got := cleanVisibleText(strings.Join([]string{
		"Visible before word/media/image1.png visible after",
		"Cell value [Content_Types].xml continues",
		"Encoded media%2Fencoded%20texture.tga tail",
		"Fragment word/media/chart.png#preview tail",
		"Query media%2Fencoded%20chart.png%3Fcache%3D1 tail",
		"Paren media/standalone.png) tail",
		"Bracket media/standalone.jpg] tail",
		"Compressed media/vector.emz tail",
		"Relationship rId7 text",
		`RelationshipAttr r:embed="rId8" text`,
		`RelationshipId Id="rId11" text`,
		`RelationshipType Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" text`,
		`RelationshipMode TargetMode="External" text`,
		`ContentTypeAttr ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" text`,
		`PartNameAttr PartName="/ppt/slides/slide1.xml" text`,
		`TargetAttr Target="../media/inline-relative.png"/> text`,
		`EncodedTarget Target="..%2Fmedia%2Finline-encoded.png" text`,
		`SpacedTarget Target = "../media/spaced-relative.png" text`,
		`SpacedEmbed r:embed = "rId10" text`,
		`SpacedHref href = "file:///C:/Users/me/hidden-link.docx" text`,
		`MHTMLLocation Content-Location: word/media/mhtml-hidden.png text`,
		`MHTMLCID Content-ID: <image001.png@office> text`,
		`MHTMLType Content-Type: image/png text`,
		`MHTMLTransfer Content-Transfer-Encoding: base64 text`,
		`MHTMLDisposition Content-Disposition: inline; filename="hidden.png" text`,
		`MHTMLDescription Content-Description: "hidden preview" text`,
		`MHTMLBase Content-Base: file:///C:/Users/me/hidden/ text`,
		`MHTMLVersion MIME-Version: 1.0 text`,
		`ColonTarget Target: ../media/colon-target.png text`,
		`ColonType Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image text`,
		`ColonMode TargetMode: External text`,
		`ColonContentType ContentType: application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml text`,
		`ColonPartName PartName: /word/document.xml text`,
		`ColonEmbed r:embed: rId12 text`,
		`CSSFill fill:url(word/media/css-hidden.png) text`,
		`QuotedCSS background:url("ppt%2Fmedia%2Fcss%20hidden.png") text`,
		"Wrapped <xl/_rels/workbook.xml.rels> done",
		"Package pack://application:,,,/word/media/packed.png done",
		"OPC opc://package/xl/_rels/workbook.xml.rels done",
		`FileURI file:\C:\Users\me\hidden.docx done`,
		"SingleSlash file:/C:/Users/me/hidden-single-slash.docx done",
		"OfficeApp ms-word:ofe|u|file:///C:/Users/me/hidden.docx done",
		"WrappedInlineBefore&lt;word%2Fmedia%2Finline%20hidden.png&gt;After",
		"BracketInlineBefore[ppt%2Fmedia%2Finline%20hidden.png]After",
		"ParenInlineBefore(media%2Finline-hidden.tga)After",
		"BraceInlineBefore{word%2Fmedia%2Fbrace-hidden.png}After",
		"VML src:word/media/colon-hidden.png tail",
		`KeepHref href = "https://example.test/media/photo.png?size=large" visible`,
		`KeepCSS background:url(https://example.test/media/photo.png?size=large) visible`,
		"Keep {visible-token} wrapped prose",
		"Keep C:\\Reports\\Q1 and https://example.test/media/photo.png?size=large as visible prose",
		"Keep Target audience and Type designations as visible prose",
	}, "\n"))
	for _, want := range []string{
		"Visible before visible after",
		"Cell value continues",
		"Encoded tail",
		"Fragment tail",
		"Query tail",
		"Paren tail",
		"Bracket tail",
		"Compressed tail",
		"Relationship text",
		"RelationshipAttr text",
		"RelationshipId text",
		"RelationshipType text",
		"RelationshipMode text",
		"ContentTypeAttr text",
		"PartNameAttr text",
		"TargetAttr text",
		"EncodedTarget text",
		"SpacedTarget text",
		"SpacedEmbed text",
		"SpacedHref text",
		"MHTMLLocation text",
		"MHTMLCID text",
		"MHTMLType text",
		"MHTMLTransfer text",
		"MHTMLDisposition text",
		"MHTMLDescription text",
		"MHTMLBase text",
		"MHTMLVersion text",
		"ColonTarget text",
		"ColonType text",
		"ColonMode text",
		"ColonContentType text",
		"ColonPartName text",
		"ColonEmbed text",
		"CSSFill fill: text",
		"QuotedCSS background: text",
		"Wrapped done",
		"Package done",
		"OPC done",
		"FileURI done",
		"SingleSlash done",
		"OfficeApp done",
		"WrappedInlineBeforeAfter",
		"BracketInlineBeforeAfter",
		"ParenInlineBeforeAfter",
		"BraceInlineBeforeAfter",
		"VML tail",
		`KeepHref href = "https://example.test/media/photo.png?size=large" visible`,
		`KeepCSS background:url(https://example.test/media/photo.png?size=large) visible`,
		"Keep {visible-token} wrapped prose",
		"Keep C:\\Reports\\Q1 and https://example.test/media/photo.png?size=large as visible prose",
		"Keep Target audience and Type designations as visible prose",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("missing cleaned visible text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"word/media/image1.png", "[Content_Types].xml", "media%2Fencoded%20texture.tga", "word/media/chart.png", "media%2Fencoded%20chart.png", "media/standalone.png", "media/standalone.jpg", "media/vector.emz", "rId7", "rId8", "rId10", "rId11", "rId12", "relationships/image", "TargetMode:", "External", `ContentType="`, "ContentType:", `PartName="`, "PartName:", "application/vnd.openxmlformats", "/ppt/slides/slide1.xml", "/word/document.xml", "../media/inline-relative.png", "..%2Fmedia%2Finline-encoded.png", "../media/spaced-relative.png", "../media/colon-target.png", "file:///C:/Users/me/hidden-link.docx", "Content-Location", "Content-ID", "Content-Type", "Content-Transfer-Encoding", "Content-Disposition", "Content-Description", "Content-Base", "MIME-Version", "word/media/mhtml-hidden.png", "image001.png@office", "image/png", "base64", "filename=", "hidden preview", "file:///C:/Users/me/hidden/", "word/media/css-hidden.png", "ppt%2Fmedia%2Fcss%20hidden.png", "xl/_rels/workbook.xml.rels", "pack://application:,,,/word/media/packed.png", "opc://package/xl/_rels/workbook.xml.rels", `file:\C:\Users\me\hidden.docx`, "file:/C:/Users/me/hidden-single-slash.docx", "ms-word:ofe|u|file:///C:/Users/me/hidden.docx", "word%2Fmedia%2Finline%20hidden.png", "ppt%2Fmedia%2Finline%20hidden.png", "media%2Finline-hidden.tga", "word%2Fmedia%2Fbrace-hidden.png", "word/media/colon-hidden.png"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("kept inline hidden reference %q in %q", hidden, got)
		}
	}
}

func TestCleanVisibleTextKeepsRelationshipModeWordsInProse(t *testing.T) {
	got := cleanVisibleText("External collaboration remains visible\nInternal review remains visible\nContent type guidance remains visible\nPart name discussion remains visible\nTargetMode=\"External\"\nContentType=\"application/xml\"\nPartName=\"/word/document.xml\"")
	for _, want := range []string{"External collaboration remains visible", "Internal review remains visible", "Content type guidance remains visible", "Part name discussion remains visible"} {
		if !strings.Contains(got, want) {
			t.Fatalf("dropped visible relationship-mode prose %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"TargetMode", "ContentType", "PartName", "application/xml", "/word/document.xml"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("kept relationship/content-type metadata %q in %q", hidden, got)
		}
	}
}

func TestLongVisibleTextWithOfficeMetadataWordsIsKept(t *testing.T) {
	longVisible := "Visible long paragraph start. " +
		strings.Repeat("This prose discusses xmlns prefixes, TargetMode behavior, and schemas.openxmlformats.org references without being package metadata. ", 120) +
		"Visible long paragraph end."
	got := cleanVisibleText(strings.Join([]string{
		longVisible,
		`TargetMode="External"`,
		`ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`,
		`PartName="/word/document.xml"`,
	}, "\n"))
	for _, want := range []string{"Visible long paragraph start.", "Visible long paragraph end."} {
		if !strings.Contains(got, want) {
			t.Fatalf("dropped long visible prose %q in %q", want, got)
		}
	}
	for _, hidden := range []string{`TargetMode="External"`, `ContentType="application/vnd.openxmlformats`, `PartName="/word/document.xml"`, "application/vnd.openxmlformats", "/word/document.xml"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("kept short Office metadata %q in %q", hidden, got)
		}
	}
}

func TestLongOfficeMetadataReferencesAreStillRecognized(t *testing.T) {
	longRelationship := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" + strings.Repeat("/relationships/image", 600)
	if !looksLikeOfficeRelationshipMetadataReference(longRelationship) {
		t.Fatalf("expected long relationship metadata reference to be recognized")
	}
	longXML := `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main` + strings.Repeat("x", maxHiddenResourceMetadataReferenceBytes)
	if !looksLikeOfficeXMLMetadataReference(longXML) {
		t.Fatalf("expected long XML metadata reference to be recognized")
	}
	longVisible := "Visible " + strings.Repeat("targetmode xmlns schemas.openxmlformats.org ", 300)
	if looksLikeOfficeRelationshipMetadataReference(longVisible) {
		t.Fatalf("long visible prose was misclassified as relationship metadata")
	}
	if looksLikeOfficeXMLMetadataReference(longVisible) {
		t.Fatalf("long visible prose was misclassified as XML metadata")
	}
}

func TestCleanVisibleTextDropsOfficeXMLNamespaceMetadata(t *testing.T) {
	got := cleanVisibleText(strings.Join([]string{
		"Visible namespace discussion remains visible.",
		`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`,
		`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`,
		`mc:Ignorable="w14 wp14"`,
		`xsi:schemaLocation="http://schemas.openxmlformats.org/wordprocessingml/2006/main wordprocessingml.xsd"`,
		"Visible schema prose remains visible.",
	}, "\n"))
	for _, want := range []string{"Visible namespace discussion remains visible.", "Visible schema prose remains visible."} {
		if !strings.Contains(got, want) {
			t.Fatalf("dropped visible XML metadata prose %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"xmlns:", "schemas.openxmlformats.org", "mc:Ignorable", "schemaLocation", "w14 wp14"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("kept Office XML namespace metadata %q in %q", hidden, got)
		}
	}
}

func TestAllXMLCharDataTextDropsInternalReferences(t *testing.T) {
	xml := []byte(`<root>
		<p>Visible XML body</p>
		<p>word/media/image1.png</p>
		<p>[ppt%2Fmedia%2Fwrapped%20image.png]</p>
		<p>(rId7)</p>
		<p>Target="../media/xml-relative.png"/></p>
		<p>Target = "../media/xml-spaced-relative.png"</p>
		<p>r:embed="rId8"</p>
		<p>r:embed = "rId9"</p>
		<p>Id="rId11"</p>
		<p>Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"</p>
		<p>TargetMode="External"</p>
		<p>pack://application:,,,/word/media/packed.png</p>
		<p>opc://package/xl/_rels/workbook.xml.rels</p>
		<p>file://server/share/hidden.docx</p>
		<p>file:\C:\Users\me\hidden.docx</p>
		<p>file:/C:/Users/me/hidden-single-slash.docx</p>
		<p>ms-word:ofe|u|file:///C:/Users/me/hidden.docx</p>
		<script>Hidden Script Secret</script>
		<style>Hidden Style Secret</style>
		<x:ClientData>Hidden ClientData Secret</x:ClientData>
		<p>Visible XML footer</p>
	</root>`)
	got, err := allXMLCharDataText(xml)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible XML body", "Visible XML footer"} {
		if !strings.Contains(got, want) {
			t.Fatalf("missing visible XML text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"word/media/image1.png", "ppt%2Fmedia%2Fwrapped%20image.png", "rId7", "rId8", "rId9", "rId11", "relationships/image", "TargetMode", "External", "../media/xml-relative.png", "../media/xml-spaced-relative.png", "pack://application:,,,/word/media/packed.png", "opc://package/xl/_rels/workbook.xml.rels", "file://server/share/hidden.docx", `file:\C:\Users\me\hidden.docx`, "file:/C:/Users/me/hidden-single-slash.docx", "ms-word:ofe|u|file:///C:/Users/me/hidden.docx", "Hidden Script Secret", "Hidden Style Secret", "Hidden ClientData Secret"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("kept internal XML char data %q in %q", hidden, got)
		}
	}
}

func TestMarkdownBackfillSkipsRelationshipIDLines(t *testing.T) {
	res := &Result{
		Text:               "Visible body\nrId7\n[RID42]\n\"rId9\"\nId=\"rId11\"\nType=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"\nTargetMode=\"External\"\nVisible footer",
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible footer") {
		t.Fatalf("markdown should backfill visible text:\n%s", md)
	}
	for _, hidden := range []string{"rId7", "RID42", "rId9", "rId11", "relationships/image", "TargetMode", "External"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown backfill kept relationship ID %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownBackfillStripsInlineHiddenReferences(t *testing.T) {
	res := &Result{
		Text: strings.Join([]string{
			"Visible body",
			"Visible caption word/media/hidden.png tail",
			`Visible relation r:embed="rId9" tail`,
			`Visible target Target="../media/hidden-target.png" tail`,
			"Visible footer",
		}, "\n"),
		StructuredMarkdown: "## Document\n\nVisible body",
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible caption tail", "Visible relation tail", "Visible target tail", "Visible footer"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown backfill missing cleaned visible text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"word/media/hidden.png", `r:embed="rId9"`, "rId9", `Target="../media/hidden-target.png"`, "../media/hidden-target.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown backfill kept inline hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestMarkdownImageAltIsSingleLineAndEscaped(t *testing.T) {
	res := &Result{Images: []Image{{
		Name: "diagram.png",
		Alt:  "Revenue [draft]\nC:\\Reports\\Q1",
		Ext:  ".png",
		Data: testPNG(),
	}}}
	md := res.Markdown("images")
	want := "![Revenue \\[draft\\] C:\\\\Reports\\\\Q1](images/diagram.png)"
	if !strings.Contains(md, want) {
		t.Fatalf("markdown image alt was not normalized/escaped, want %q in:\n%s", want, md)
	}
	if strings.Contains(md, "draft]\n") || strings.Contains(md, "\nC:\\Reports") {
		t.Fatalf("markdown image alt should not contain raw line breaks:\n%s", md)
	}
}

func TestMarkdownImageAltKeepsVisibleLocalPathLabels(t *testing.T) {
	res := &Result{Images: []Image{{
		Name: "report.png",
		Alt:  "C:\\Reports\\Q1",
		Ext:  ".png",
		Data: testPNG(),
	}}}
	md := res.Markdown("images")
	if !strings.Contains(md, "![C:\\\\Reports\\\\Q1](images/report.png)") {
		t.Fatalf("markdown should keep visible local path alt label:\n%s", md)
	}
}

func TestMarkdownImageAltSkipsLocalImagePaths(t *testing.T) {
	for _, alt := range []string{
		`C:\Users\me\Pictures\hidden.jpg`,
		`D:/shared/media/hidden.png`,
		`file:///C:/Users/me/Pictures/hidden.png`,
		`C%3A%5CUsers%5Cme%5CPictures%5Chidden.jpg`,
		`D%3A/shared/media/hidden.png`,
		`file%3A%2F%2F%2FC%3A%2FUsers%2Fme%2FPictures%2Fhidden.png`,
		`&lt;C%3A%5CUsers%5Cme%5CPictures%5Chidden.jpg&gt;`,
	} {
		res := &Result{Images: []Image{{
			Name: "diagram.png",
			Alt:  alt,
			Ext:  ".png",
			Data: testPNG(),
		}}}
		md := res.Markdown("images")
		if strings.Contains(md, "hidden") || strings.Contains(md, "Pictures") || strings.Contains(md, "shared") || strings.Contains(md, "file:///") {
			t.Fatalf("markdown should not expose local image path alt %q in:\n%s", alt, md)
		}
		if !strings.Contains(md, "![diagram](images/diagram.png)") {
			t.Fatalf("markdown should fall back to filename alt for %q:\n%s", alt, md)
		}
	}
}

func TestMarkdownImageAltKeepsEncodedVisibleNonImageLabels(t *testing.T) {
	for _, alt := range []string{
		`C%3A%5CReports%5CQ1`,
		`Quarterly%20Report`,
	} {
		res := &Result{Images: []Image{{
			Name: "diagram.png",
			Alt:  alt,
			Ext:  ".png",
			Data: testPNG(),
		}}}
		md := res.Markdown("images")
		if !strings.Contains(md, "!["+alt+"](images/diagram.png)") {
			t.Fatalf("markdown should keep visible non-image alt %q in:\n%s", alt, md)
		}
	}
}

func TestMarkdownImageAltSkipsRelationshipIDs(t *testing.T) {
	res := &Result{Images: []Image{{
		Name: "diagram.png",
		Alt:  "(rId7)",
		Ext:  ".png",
		Data: testPNG(),
	}}}
	md := res.Markdown("images")
	if strings.Contains(md, "rId7") {
		t.Fatalf("markdown should not expose relationship ID image alt:\n%s", md)
	}
	if !strings.Contains(md, "![diagram](images/diagram.png)") {
		t.Fatalf("markdown should fall back to image filename alt:\n%s", md)
	}
}

func TestMarkdownImageAltSkipsOfficeRelationshipAttributes(t *testing.T) {
	for _, alt := range []string{
		`Target="../media/hidden.png"`,
		`Target="..%2Fmedia%2Fencoded-hidden.png"/>`,
		`r:embed="rId8"`,
		`Id="rId11"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`,
		`TargetMode="External"`,
		`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`,
		`mc:Ignorable="a14"`,
		`xsi:schemaLocation="http://schemas.openxmlformats.org/drawingml/2006/main drawingml.xsd"`,
		`word%2Fmedia%2Fencoded-alt.png`,
		`ppt%252Fmedia%252Fdouble-encoded-alt.png`,
		`&lt;xl%2Fmedia%2Fentity-alt.png&gt;`,
	} {
		res := &Result{Images: []Image{{
			Name: "diagram.png",
			Alt:  alt,
			Ext:  ".png",
			Data: testPNG(),
		}}}
		md := res.Markdown("images")
		for _, hidden := range []string{"Target=", "../media/hidden.png", "..%2Fmedia%2Fencoded-hidden.png", "rId8", "rId11", "relationships/image", "TargetMode", "External", "xmlns", "schemas.openxmlformats.org", "mc:Ignorable", "schemaLocation", "word%2Fmedia", "ppt%252Fmedia", "xl%2Fmedia"} {
			if strings.Contains(md, hidden) {
				t.Fatalf("markdown should not expose Office relationship attribute %q in:\n%s", hidden, md)
			}
		}
		if !strings.Contains(md, "![diagram](images/diagram.png)") {
			t.Fatalf("markdown should fall back to image filename alt for %q:\n%s", alt, md)
		}
	}
}

func TestMarkdownImageFilenameFallbackAltSkipsHiddenReferences(t *testing.T) {
	images := []Image{
		{Name: "ContentType=image_png rId77.png", Ext: ".png", Data: testPNG()},
		{Name: "rId88.png", Ext: ".png", Data: testPNG()},
		{Name: "Visible diagram ContentType=image_png rId99.png", Ext: ".png", Data: testPNG()},
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, hidden := range []string{"ContentType", "image_png", "rId77", "rId88", "rId99"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown filename fallback alt exposed hidden reference %q in:\n%s", hidden, md)
		}
	}
	for _, want := range []string{
		"![image-001](images/image-001.png)",
		"![image-002](images/image-002.png)",
		"![Visible diagram](images/Visible%20diagram.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing cleaned filename fallback image %q in:\n%s", want, md)
		}
	}
}

func TestMarkdownImageAltStripsInlineHiddenOfficeReferences(t *testing.T) {
	res := &Result{
		Text:               "Visible diagram",
		StructuredMarkdown: "## Document\n\nVisible diagram",
		Images: []Image{{
			Name: "diagram.png",
			Alt:  "Visible word/media/image1.png diagram",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if strings.Contains(md, "word/media/image1.png") {
		t.Fatalf("markdown image alt kept inline internal reference:\n%s", md)
	}
	if !strings.Contains(md, "![Visible diagram](images/diagram.png)") {
		t.Fatalf("markdown image alt was not cleaned and used for placement:\n%s", md)
	}
	if !strings.Contains(md, "Visible diagram\n![Visible diagram](images/diagram.png)") {
		t.Fatalf("markdown should preserve visible alt line before placed image:\n%s", md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("placed cleaned-alt image should not be duplicated in trailing Images section:\n%s", md)
	}
}

func TestMarkdownImageAltStripsInlineOfficeXMLMetadata(t *testing.T) {
	res := &Result{
		Text:               "Visible diagram",
		StructuredMarkdown: "## Document\n\nVisible diagram",
		Images: []Image{{
			Name: "diagram.png",
			Alt:  `Visible xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" mc:Ignorable="a14" diagram`,
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	for _, hidden := range []string{"xmlns:a", "schemas.openxmlformats.org", "mc:Ignorable", "a14"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown image alt kept inline Office XML metadata %q in:\n%s", hidden, md)
		}
	}
	if !strings.Contains(md, "![Visible diagram](images/diagram.png)") {
		t.Fatalf("markdown image alt was not cleaned and used for placement:\n%s", md)
	}
	if !strings.Contains(md, "Visible diagram\n![Visible diagram](images/diagram.png)") {
		t.Fatalf("markdown should preserve visible alt line before placed image:\n%s", md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("placed cleaned-alt image should not be duplicated in trailing Images section:\n%s", md)
	}
}

func TestMarkdownHeadingIsSingleLineAndDemoted(t *testing.T) {
	for _, tc := range []struct {
		in   string
		want string
	}{
		{in: "### Report\nQ1", want: "Report Q1"},
		{in: "  ##   Sheet   Name  ", want: "Sheet Name"},
		{in: "###", want: "Sheet"},
		{in: "xl/_rels/workbook.xml.rels", want: "Sheet"},
		{in: "media/standalone.tga", want: "Sheet"},
		{in: "Quarter word/media/image1.png Summary", want: "Quarter Summary"},
		{in: "Visible media%2Fencoded%20texture.tga Heading", want: "Visible Heading"},
		{in: `Visible xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" Heading`, want: "Visible Heading"},
		{in: `Visible mc:Ignorable="a14" Heading`, want: "Visible Heading"},
		{in: `Target="../media/heading.png"`, want: "Sheet"},
		{in: `r:embed="rId8"`, want: "Sheet"},
		{in: `xsi:schemaLocation="http://schemas.openxmlformats.org/drawingml/2006/main drawingml.xsd"`, want: "Sheet"},
		{in: "(rId7)", want: "Sheet"},
		{in: "WordPad.Document.1", want: "Sheet"},
		{in: `<b>Visible &amp; Clean</b>`, want: "Visible & Clean"},
		{in: `2 < 3 > 1`, want: "2 < 3 > 1"},
	} {
		if got := escapeMarkdownHeading(tc.in); got != tc.want {
			t.Fatalf("escapeMarkdownHeading(%q)=%q, want %q", tc.in, got, tc.want)
		}
	}
}

func TestMarkdownTableCellEscapesAndCompactsLines(t *testing.T) {
	got := escapeMarkdownTableCell(" First | Line \n\n word/media/image1.png \n C:\\Reports\\Q1 \n Second | Line ")
	want := "First \\| Line<br>C:\\\\Reports\\\\Q1<br>Second \\| Line"
	if got != want {
		t.Fatalf("escapeMarkdownTableCell got %q, want %q", got, want)
	}
	if strings.Contains(got, "<br><br>") {
		t.Fatalf("table cell should not keep blank markdown lines: %q", got)
	}
	if got := escapeMarkdownTableCell("media/standalone.tga"); got != "" {
		t.Fatalf("table cell should drop internal resource-only value, got %q", got)
	}
	if got := escapeMarkdownTableCell("Repeated\nRepeated\nDifferent\nRepeated"); got != "Repeated<br>Different<br>Repeated" {
		t.Fatalf("table cell should drop only adjacent duplicate lines, got %q", got)
	}
	for _, hidden := range []string{
		"xl/_rels/workbook.xml.rels",
		"file:///C:/Users/me/hidden.docx",
		"Target=\"../media/table.png\"",
		`Id="rId11"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`,
		`TargetMode="External"`,
		`Target: ../media/table.png`,
		`Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`,
		`ContentType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`,
		`PartName: /xl/worksheets/sheet1.xml`,
		`r:embed: rId8`,
		`xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"`,
		`mc:Ignorable="x14ac"`,
		`schemaLocation="http://schemas.openxmlformats.org/spreadsheetml/2006/main sheet.xsd"`,
		"word%2Fmedia%2Fencoded-table.png",
		"ppt%252Fmedia%252Fdouble-encoded-table.png",
		"&lt;xl%2Fmedia%2Fentity-table.png&gt;",
	} {
		if got := escapeMarkdownTableCell(hidden); got != "" {
			t.Fatalf("table cell should drop hidden resource %q, got %q", hidden, got)
		}
	}
}

func TestMarkdownTableCellDropsInternalReferences(t *testing.T) {
	for _, hidden := range []string{
		"word/media/image1.png",
		`<xl\_rels\workbook.xml.rels>`,
		"[ppt%2Fmedia%2Fwrapped%20image.png]",
		"rId7",
		"(RID42)",
		`Target="../media/table.png"`,
		`Target="..%2Fmedia%2Fencoded-table.png"/>`,
		`r:embed="rId8"`,
		`Id="rId11"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`,
		`TargetMode="External"`,
		`Target: ../media/table.png`,
		`Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`,
		`ContentType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`,
		`PartName: /xl/worksheets/sheet1.xml`,
		`r:embed: rId8`,
		`xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"`,
		`mc:Ignorable="x14ac"`,
		`schemaLocation="http://schemas.openxmlformats.org/spreadsheetml/2006/main sheet.xsd"`,
		"word%2Fmedia%2Fencoded-table.png",
		"ppt%252Fmedia%252Fdouble-encoded-table.png",
		"&lt;xl%2Fmedia%2Fentity-table.png&gt;",
	} {
		if got := cleanMarkdownTableCellValue(hidden); got != "" {
			t.Fatalf("markdown table cell kept internal reference %q as %q", hidden, got)
		}
	}
	if got := cleanMarkdownTableCellValue("Visible table text"); got != "Visible table text" {
		t.Fatalf("markdown table cell dropped visible text: %q", got)
	}
	if got := cleanMarkdownTableCellValue("Visible word/media/image1.png table text"); got != "Visible table text" {
		t.Fatalf("markdown table cell kept inline internal reference: %q", got)
	}
	if got := cleanMarkdownTableCellValue(`Visible Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" table text`); got != "Visible table text" {
		t.Fatalf("markdown table cell kept inline relationship type metadata: %q", got)
	}
	if got := cleanMarkdownTableCellValue(`Visible TargetMode="External" table text`); got != "Visible table text" {
		t.Fatalf("markdown table cell kept inline target mode metadata: %q", got)
	}
	for _, tc := range []struct {
		in   string
		want string
	}{
		{`Visible Target: ../media/table.png table text`, "Visible table text"},
		{`Visible Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image table text`, "Visible table text"},
		{`Visible ContentType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml table text`, "Visible table text"},
		{`Visible PartName: /xl/worksheets/sheet1.xml table text`, "Visible table text"},
		{`Visible r:embed: rId8 table text`, "Visible table text"},
	} {
		if got := cleanMarkdownTableCellValue(tc.in); got != tc.want {
			t.Fatalf("markdown table cell kept inline colon metadata %q as %q, want %q", tc.in, got, tc.want)
		}
	}
	if got := cleanMarkdownTableCellValue(`Before&lt;word%2Fmedia%2Fhidden-table.png&gt;After`); got != "BeforeAfter" {
		t.Fatalf("markdown table cell kept wrapped inline hidden reference: %q", got)
	}
	if got := cleanMarkdownTableCellValue(`Before{word%2Fmedia%2Fhidden-table.png}After`); got != "BeforeAfter" {
		t.Fatalf("markdown table cell kept brace-wrapped inline hidden reference: %q", got)
	}
	if got := cleanMarkdownTableCellValue(`Keep {visible-token} table text`); got != "Keep {visible-token} table text" {
		t.Fatalf("markdown table cell dropped visible brace text: %q", got)
	}
}

func TestMarkdownTableCellCleansHiddenReferencesBeforeTruncating(t *testing.T) {
	hidden := strings.Repeat("word/media/hidden.png ", maxMarkdownTableCellBytes/4)
	got := cleanMarkdownTableCellValue(hidden + "Visible tail text")
	if got != "Visible tail text" {
		t.Fatalf("markdown table cell should preserve visible text after removing hidden references, got %q", got)
	}
}

func TestMarkdownTableCellNormalizesInvisibleSpacing(t *testing.T) {
	got := cleanMarkdownTableCellValue("Visible\u00a0Cell\u200b\ufeff Text")
	if got != "Visible Cell Text" {
		t.Fatalf("markdown table cell did not normalize invisible spacing: %q", got)
	}
	escaped := escapeMarkdownTableCell("A\u202f|\u200bB\nC\u3000D")
	if escaped != "A \\|B<br>C D" {
		t.Fatalf("escaped markdown table cell did not normalize spacing, got %q", escaped)
	}
	for _, r := range []rune{'\u00a0', '\u202f', '\u200b', '\ufeff', '\u3000'} {
		if strings.ContainsRune(got, r) || strings.ContainsRune(escaped, r) {
			t.Fatalf("markdown table cell kept invisible/ambiguous rune %U in %q / %q", r, got, escaped)
		}
	}
}

func TestMarkdownVisibleLineTextFastPath(t *testing.T) {
	if got := markdownVisibleLineText("  Plain visible text  "); got != "Plain visible text" {
		t.Fatalf("plain visible text fast path got %q", got)
	}
	if got := markdownVisibleLineText("1. listed item"); got != "listed item" {
		t.Fatalf("ordered list should still be normalized, got %q", got)
	}
	if got := markdownVisibleLineText("Visible text\\"); got != "Visible text" {
		t.Fatalf("hard line break marker should still be stripped, got %q", got)
	}
}

func TestCleanTextDecodesOOXMLSoftBreakEscapes(t *testing.T) {
	got := cleanText("Line one_x000B_Line two_x000C_Page two")
	want := "Line one\nLine two\nPage two"
	if got != want {
		t.Fatalf("cleanText did not decode OOXML soft/page breaks:\n got %q\nwant %q", got, want)
	}
	cell := escapeMarkdownTableCell("Cell one_x000B_Cell two_x000C_Cell three")
	if cell != "Cell one<br>Cell two<br>Cell three" {
		t.Fatalf("markdown table cell did not preserve OOXML soft/page breaks: %q", cell)
	}
}

func TestCleanTextNormalizesRawSoftBreakControls(t *testing.T) {
	got := cleanText("Line one\vLine two\fPage two")
	want := "Line one\nLine two\nPage two"
	if got != want {
		t.Fatalf("cleanText did not normalize raw soft/page breaks:\n got %q\nwant %q", got, want)
	}
	cell := escapeMarkdownTableCell("Cell one\vCell two\fCell three")
	if cell != "Cell one<br>Cell two<br>Cell three" {
		t.Fatalf("markdown table cell did not preserve raw soft/page breaks: %q", cell)
	}
}

func TestResolveOOXMLRelationshipTargetStaysInPackageRoot(t *testing.T) {
	for _, tc := range []struct {
		source string
		target string
		want   string
	}{
		{source: "word/document.xml", target: "media/Visible%20Image.PNG", want: "word/media/Visible Image.PNG"},
		{source: "word/document.xml", target: "media/Visible%20Image.PNG?download=1#section", want: "word/media/Visible Image.PNG"},
		{source: `PPT\Slides\Slide1.XML`, target: "../Media/Visible.PNG", want: "PPT/Media/Visible.PNG"},
		{source: `PPT\Slides\Slide1.XML`, target: "../Media/Visible.PNG#rId1", want: "PPT/Media/Visible.PNG"},
		{source: "xl/worksheets/sheet1.xml", target: "../drawings/drawing1.xml", want: "xl/drawings/drawing1.xml"},
		{source: "xl/worksheets/sheet1.xml", target: "../drawings/drawing1.xml?revision=2", want: "xl/drawings/drawing1.xml"},
		{source: "xl/workbook.xml", target: "/xl/worksheets/sheet3.xml", want: "xl/worksheets/sheet3.xml"},
		{source: "word/document.xml", target: "../docProps/app.xml", want: ""},
		{source: "word/document.xml", target: "/ppt/slides/slide1.xml", want: ""},
		{source: "xl/workbook.xml", target: "../../word/document.xml", want: ""},
		{source: "ppt/slides/slide1.xml", target: "https://example.test/image.png", want: ""},
	} {
		if got := resolveOOXMLRelationshipTarget(tc.source, tc.target); got != tc.want {
			t.Fatalf("resolveOOXMLRelationshipTarget(%q,%q)=%q, want %q", tc.source, tc.target, got, tc.want)
		}
	}
}

func TestImageOutputUsesSniffedExtension(t *testing.T) {
	dir := t.TempDir()
	res := &Result{Images: []Image{{
		Name: "mislabelled.jpg",
		Ext:  ".png",
		Data: testPNG(),
	}}}
	if err := writeImages(dir, res.Images); err != nil {
		t.Fatal(err)
	}
	written := filepath.Join(dir, "mislabelled.png")
	b, err := os.ReadFile(written)
	if err != nil {
		t.Fatalf("expected image to be written with sniffed extension: %v", err)
	}
	if !validImageData(".png", b) {
		t.Fatalf("written sniffed-extension image is invalid: %s", written)
	}
	if _, err := os.Stat(filepath.Join(dir, "mislabelled.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("mislabelled image name should not be written, stat err=%v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![mislabelled](images/mislabelled.png)") || strings.Contains(md, "mislabelled.jpg") {
		t.Fatalf("markdown should reference sniffed extension only:\n%s", md)
	}
}

func TestOOXMLImageAltTextFeedsMarkdownImageAlt(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:wp="urn:wp"><w:body><w:p><w:r><w:t>Visible body</w:t></w:r></w:p><wp:docPr descr="Generated Picture Description" title="Generated Picture Title"/><wp:docPr descr="C:\Users\me\hidden.png"/></w:body></w:document>`)
	addZipBytes(t, zw, "word/media/image1.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "image-alt.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Generated Picture Description](images/image1.png)") {
		t.Fatalf("markdown did not use visible image alt text in:\n%s", md)
	}
	if strings.Contains(md, "C:\\Users\\me\\hidden.png") {
		t.Fatalf("markdown kept hidden resource alt text in:\n%s", md)
	}
	if _, err := os.Stat(filepath.Join(outDir, "image1.png")); err != nil {
		t.Fatalf("image filename changed or was not written: %v", err)
	}
}

func TestOOXMLImageAltMetadataIsCleanedInResult(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Visible OOXML image alt body</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible OOXML Picture Content-ID: &lt;hidden@office&gt; Target: ../media/hidden.png" title="Content-Type: image/png"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicture"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/visible.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "ooxml-clean-image-alt.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible.png" || res.Images[0].Alt != "Visible OOXML Picture" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected cleaned OOXML image alt: %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("expected written cleaned OOXML image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible OOXML image alt body", "![Visible OOXML Picture](images/visible.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing cleaned OOXML image content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-ID", "hidden@office", "Target:", "../media/hidden.png", "Content-Type", "image/png"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) || strings.Contains(res.Images[0].Alt, hidden) {
			t.Fatalf("kept OOXML image alt metadata %q in text=%q alt=%q markdown=\n%s", hidden, res.Text, res.Images[0].Alt, md)
		}
	}
}

func TestDOCXMarkdownPlacesImageAtVisibleAltLocation(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Visible before picture</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Inline visible picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicture"/></p:blipFill></p:pic>
<w:p><w:r><w:t>Visible after picture</w:t></w:r></w:p>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/inline.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "inline-image-position.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Alt != "Inline visible picture" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected one valid inline DOCX image, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible before picture", "![Inline visible picture](images/inline.png)", "Visible after picture"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing inline image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("markdown duplicated placed inline image in trailing image section:\n%s", md)
	}
	before := strings.Index(md, "Visible before picture")
	image := strings.Index(md, "![Inline visible picture](images/inline.png)")
	after := strings.Index(md, "Visible after picture")
	if !(before >= 0 && image > before && after > image) {
		t.Fatalf("markdown image was not placed between surrounding text:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotReplaceHeadings(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Picture\n\nVisible body",
		Images: []Image{{
			Name: "picture.png",
			Alt:  "Picture",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Picture", "Visible body", "## Images", "![Picture](images/picture.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing heading/image content %q in:\n%s", want, md)
		}
	}
	if strings.Index(md, "## Picture") > strings.Index(md, "Visible body") {
		t.Fatalf("markdown heading was not preserved before body:\n%s", md)
	}
	if strings.Count(md, "## Images") != 1 {
		t.Fatalf("markdown should keep unplaced heading-matched image in trailing image section:\n%s", md)
	}
}

func TestMarkdownImagePlacementPreservesListAndQuoteStructure(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\n- List picture\n\n> Quote picture\n\n- [x] Task picture",
		Images: []Image{
			{Name: "list.png", Alt: "List picture", Ext: ".png", Data: testPNG()},
			{Name: "quote.png", Alt: "Quote picture", Ext: ".png", Data: testPNG()},
			{Name: "task.png", Alt: "Task picture", Ext: ".png", Data: testPNG()},
		},
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"- List picture\n- ![List picture](images/list.png)",
		"> Quote picture\n> ![Quote picture](images/quote.png)",
		"- [x] Task picture\n- [x] ![Task picture](images/task.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown did not preserve structural text and image placement %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("all structurally placed images should avoid trailing Images section:\n%s", md)
	}
}

func TestMarkdownImagePlacementPreservesMatchedVisibleLine(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\nFigure 1: Visible caption\n\nVisible body",
		Images: []Image{{
			Name: "figure.png",
			Alt:  "Figure 1: Visible caption",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	want := "Figure 1: Visible caption\n![Figure 1: Visible caption](images/figure.png)"
	if !strings.Contains(md, want) {
		t.Fatalf("markdown should preserve matched visible line before placed image %q in:\n%s", want, md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("placed caption image should not be duplicated in trailing Images section:\n%s", md)
	}
	if strings.Index(md, "Figure 1: Visible caption") > strings.Index(md, "Visible body") {
		t.Fatalf("caption should remain before following body text:\n%s", md)
	}
}

func TestMarkdownImagePlacementPlacesUniqueTableCell(t *testing.T) {
	res := &Result{
		StructuredMarkdown: strings.Join([]string{
			"## Sheet",
			"",
			"| Item | Preview | Notes |",
			"| --- | --- | --- |",
			"| Widget | Cell picture | Ready |",
		}, "\n"),
		Images: []Image{{
			Name: "cell.png",
			Alt:  "Cell picture",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "| Widget | ![Cell picture](images/cell.png) | Ready |") {
		t.Fatalf("markdown did not place image in matching table cell:\n%s", md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("placed table image should not be duplicated in trailing Images section:\n%s", md)
	}
}

func TestMarkdownImagePlacementUsesCleanedAltForTableCell(t *testing.T) {
	res := &Result{
		StructuredMarkdown: strings.Join([]string{
			"## Sheet",
			"",
			"| Item | Preview | Notes |",
			"| --- | --- | --- |",
			"| Widget | Visible diagram | Ready |",
		}, "\n"),
		Images: []Image{{
			Name: "diagram.png",
			Alt:  "Visible Target: ../media/hidden.png Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image diagram",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "| Widget | ![Visible diagram](images/diagram.png) | Ready |") {
		t.Fatalf("markdown did not place image in table cell using cleaned alt:\n%s", md)
	}
	for _, hidden := range []string{"Target:", "../media/hidden.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown table image placement leaked hidden alt reference %q in:\n%s", hidden, md)
		}
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("placed cleaned-alt table image should not be duplicated in trailing Images section:\n%s", md)
	}
}

func TestMarkdownImagePlacementPlacesFormattedTableCellAndPreservesEscapes(t *testing.T) {
	res := &Result{
		StructuredMarkdown: strings.Join([]string{
			"## Sheet",
			"",
			"| Item | Preview | Notes |",
			"| --- | --- | --- |",
			"| Widget \\| Large | **Formatted picture** | A \\| B |",
		}, "\n"),
		Images: []Image{{
			Name: "formatted.png",
			Alt:  "Formatted picture",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "| Widget \\| Large | ![Formatted picture](images/formatted.png) | A \\| B |") {
		t.Fatalf("markdown did not place formatted table image while preserving escaped pipes:\n%s", md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("placed formatted table image should not be duplicated in trailing Images section:\n%s", md)
	}
}

func TestMarkdownImagePlacementPlacesMultipleImagesWithSameUniqueAlt(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\nGallery picture\n\nVisible body",
		Images: []Image{
			{Name: "gallery-a.png", Alt: "Gallery picture", Ext: ".png", Data: testPNG()},
			{Name: "gallery-b.png", Alt: "Gallery picture", Ext: ".png", Data: testPNG()},
		},
	}
	md := res.Markdown("images")
	want := strings.Join([]string{
		"Gallery picture",
		"![Gallery picture](images/gallery-a.png)",
		"![Gallery picture](images/gallery-b.png)",
	}, "\n")
	if !strings.Contains(md, want) {
		t.Fatalf("markdown should place both same-alt images after the unique visible caption %q in:\n%s", want, md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("same-alt images with a unique caption should not fall back to trailing Images section:\n%s", md)
	}
	if strings.Index(md, "![Gallery picture](images/gallery-b.png)") > strings.Index(md, "Visible body") {
		t.Fatalf("second same-alt image was not kept with its caption:\n%s", md)
	}
}

func TestMarkdownImagePlacementPlacesMultipleImagesInSameTableCell(t *testing.T) {
	res := &Result{
		StructuredMarkdown: strings.Join([]string{
			"## Sheet",
			"",
			"| Item | Preview |",
			"| --- | --- |",
			"| Widget | Shared cell picture |",
		}, "\n"),
		Images: []Image{
			{Name: "shared-a.png", Alt: "Shared cell picture", Ext: ".png", Data: testPNG()},
			{Name: "shared-b.png", Alt: "Shared cell picture", Ext: ".png", Data: testPNG()},
		},
	}
	md := res.Markdown("images")
	want := "| Widget | ![Shared cell picture](images/shared-a.png)<br>![Shared cell picture](images/shared-b.png) |"
	if !strings.Contains(md, want) {
		t.Fatalf("markdown should place both same-alt images in the matching table cell %q in:\n%s", want, md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("same-alt table cell images should not fall back to trailing Images section:\n%s", md)
	}
	if strings.Count(md, "![Shared cell picture]") != 2 {
		t.Fatalf("markdown should keep both table images exactly once:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotGuessDuplicateTableCell(t *testing.T) {
	res := &Result{
		StructuredMarkdown: strings.Join([]string{
			"## Sheet",
			"",
			"| A | B |",
			"| --- | --- |",
			"| Duplicate picture | Duplicate picture |",
		}, "\n"),
		Images: []Image{{
			Name: "duplicate.png",
			Alt:  "Duplicate picture",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "| Duplicate picture | Duplicate picture |") {
		t.Fatalf("markdown should preserve duplicate table cells:\n%s", md)
	}
	if !strings.Contains(md, "## Images") || !strings.Contains(md, "![Duplicate picture](images/duplicate.png)") {
		t.Fatalf("duplicate table image should remain in trailing Images section:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotTreatSinglePipeProseAsTable(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\nLabel | Pipe picture\n\nVisible body",
		Images: []Image{{
			Name: "pipe.png",
			Alt:  "Pipe picture",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Label | Pipe picture") {
		t.Fatalf("markdown should preserve single-pipe prose:\n%s", md)
	}
	if strings.Contains(md, "Label | ![Pipe picture]") {
		t.Fatalf("markdown treated single-pipe prose as table cell:\n%s", md)
	}
	if !strings.Contains(md, "## Images") || !strings.Contains(md, "![Pipe picture](images/pipe.png)") {
		t.Fatalf("unplaced single-pipe prose image should remain in trailing Images section:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotReplaceFencedCode(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\n```text\nInline image alt\n```\n\nVisible body",
		Images: []Image{{
			Name: "inline.png",
			Alt:  "Inline image alt",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	for _, want := range []string{"```text\nInline image alt\n```", "Visible body", "## Images", "![Inline image alt](images/inline.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing fenced-code/image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "```text\n![Inline image alt]") {
		t.Fatalf("markdown replaced fenced code content with image:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotReplaceIndentedCode(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\n    Inline image alt\n\n  - Nested list picture",
		Images: []Image{
			{Name: "inline.png", Alt: "Inline image alt", Ext: ".png", Data: testPNG()},
			{Name: "nested.png", Alt: "Nested list picture", Ext: ".png", Data: testPNG()},
		},
	}
	md := res.Markdown("images")
	for _, want := range []string{"    Inline image alt", "  - ![Nested list picture](images/nested.png)", "## Images", "![Inline image alt](images/inline.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing indented-code/list image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "    ![Inline image alt]") {
		t.Fatalf("markdown replaced indented code content with image:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotGuessDuplicateAltLocation(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\nRepeated picture label\n\nVisible body\n\nRepeated picture label",
		Images: []Image{{
			Name: "repeated.png",
			Alt:  "Repeated picture label",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	if strings.Count(md, "Repeated picture label") != 3 {
		t.Fatalf("markdown should preserve both duplicate text lines and image alt in trailing section:\n%s", md)
	}
	for _, want := range []string{"Visible body", "## Images", "![Repeated picture label](images/repeated.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing duplicate-alt content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Document\n\n![Repeated picture label]") {
		t.Fatalf("markdown guessed the first duplicate alt location for image placement:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotReplaceThematicBreaks(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\nIntro\n\n---\n\n* * *\n\nVisible body",
		Images: []Image{
			{Name: "dash.png", Alt: "---", Ext: ".png", Data: testPNG()},
			{Name: "star.png", Alt: "* * *", Ext: ".png", Data: testPNG()},
		},
	}
	md := res.Markdown("images")
	for _, want := range []string{"Intro\n\n---\n\n* * *\n\nVisible body", "## Images", "![---](images/dash.png)", "![* * *](images/star.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing thematic-break/image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "![---](images/dash.png)\n\n* * *") || strings.Contains(md, "---\n\n![* * *](images/star.png)") {
		t.Fatalf("markdown replaced a thematic break with image placement:\n%s", md)
	}
}

func TestMarkdownImagePlacementDoesNotReplaceHTMLBlocks(t *testing.T) {
	res := &Result{
		StructuredMarkdown: "## Document\n\n<!--\nHTML image alt\n-->\n\n<div>HTML image alt</div>\n\n<section>\nHTML image alt\n</section>\n\nVisible body",
		Images: []Image{{
			Name: "html.png",
			Alt:  "HTML image alt",
			Ext:  ".png",
			Data: testPNG(),
		}},
	}
	md := res.Markdown("images")
	for _, want := range []string{"<!--\nHTML image alt\n-->", "<div>HTML image alt</div>", "<section>\nHTML image alt\n</section>", "Visible body", "## Images", "![HTML image alt](images/html.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing HTML/image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "<div>![HTML image alt]") || strings.Contains(md, "<!--\n![HTML image alt]") || strings.Contains(md, "<section>\n![HTML image alt]") {
		t.Fatalf("markdown replaced HTML block content with image:\n%s", md)
	}
}

func TestOOXMLImageAltTextFollowsRelationshipTarget(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<p:pic><p:nvPicPr><p:cNvPr id="2" descr="Second picture description"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdSecond"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="First picture description"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdFirst"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSecond" Type="x" Target="media/image2.jpg"/><Relationship Id="rIdFirst" Type="x" Target="media/image1.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/image1.png", testPNG())
	addZipBytes(t, zw, "word/media/image2.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "image-alt-target.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"![First picture description](images/image1.png)",
		"![Second picture description](images/image2.jpg)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown image alt text did not follow relationship target %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "![Second picture description](images/image1.png)") ||
		strings.Contains(md, "![First picture description](images/image2.jpg)") {
		t.Fatalf("markdown image alt text was assigned by sorted image order instead of relationship target:\n%s", md)
	}
}

func TestOOXMLMediaImageNamesAreUnique(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Visible body</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="First duplicate basename image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdFirst"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" descr="Second duplicate basename image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdSecond"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFirst" Type="x" Target="media/a/same.png"/><Relationship Id="rIdSecond" Type="x" Target="media/b/same.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/a/same.png", testPNG())
	addZipBytes(t, zw, "word/media/b/same.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	imageDir := filepath.Join(dir, "images")
	file := filepath.Join(dir, "duplicate-ooxml-media-names.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: imageDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two visible OOXML images, got %#v", res.Images)
	}
	names := map[string]bool{}
	for _, img := range res.Images {
		if names[strings.ToLower(img.Name)] {
			t.Fatalf("duplicate OOXML image name %q in %#v", img.Name, res.Images)
		}
		names[strings.ToLower(img.Name)] = true
		if !validImageData(img.Ext, img.Data) {
			t.Fatalf("expected valid OOXML image data for %#v", img)
		}
	}
	if !names["same.png"] || !names["same-2.png"] {
		t.Fatalf("expected stable unique OOXML image names, got %#v", res.Images)
	}
	for _, name := range []string{"same.png", "same-2.png"} {
		if b, err := os.ReadFile(filepath.Join(imageDir, name)); err != nil || !validImageData(".png", b) {
			t.Fatalf("expected written valid image %s, err=%v len=%d", name, err, len(b))
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"![First duplicate basename image](images/same.png)",
		"![Second duplicate basename image](images/same-2.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing unique OOXML image reference %q in:\n%s", want, md)
		}
	}
}

func TestPPTXImageAltTextFollowsRelationshipTarget(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:pic><p:nvPicPr><p:cNvPr id="2" descr="Second slide picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdSecond"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="First slide picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdFirst"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSecond" Type="x" Target="../media/image2.jpg"/><Relationship Id="rIdFirst" Type="x" Target="../media/image1.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/image1.png", testPNG())
	addZipBytes(t, zw, "ppt/media/image2.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "pptx-image-alt-target.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"![First slide picture](images/image1.png)",
		"![Second slide picture](images/image2.jpg)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("PPTX markdown image alt text did not follow relationship target %q in:\n%s", want, md)
		}
	}
}

func TestXLSXImageAltTextFollowsRelationshipTarget(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="rIdDrawing"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="x" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" descr="Second sheet picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdSecond"/></xdr:blipFill></xdr:pic>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" descr="First sheet picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdFirst"/></xdr:blipFill></xdr:pic>
</xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSecond" Type="x" Target="../media/image2.jpg"/><Relationship Id="rIdFirst" Type="x" Target="../media/image1.png"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/image1.png", testPNG())
	addZipBytes(t, zw, "xl/media/image2.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "xlsx-image-alt-target.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"![First sheet picture](images/image1.png)",
		"![Second sheet picture](images/image2.jpg)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("XLSX markdown image alt text did not follow relationship target %q in:\n%s", want, md)
		}
	}
}

func TestMalformedImageAltXMLDoesNotFailExtraction(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Broken Alt"><br></b></xdr:wsDr>`)
	addZipBytes(t, zw, "xl/media/image1.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	files := map[string]*zip.File{}
	for _, f := range zr.File {
		files[f.Name] = f
	}
	images, err := extractOOXMLImages(files, "xlsx", false)
	if err != nil {
		t.Fatal(err)
	}
	if len(images) != 1 {
		t.Fatalf("expected image extraction to continue, got %#v", images)
	}
	res := &Result{Images: images}
	if !strings.Contains(res.Markdown("images"), "![image1](images/image1.png)") {
		t.Fatalf("expected fallback image markdown after malformed alt XML, got:\n%s", res.Markdown("images"))
	}
}

func TestMarkdownTextCollapsesVerticalASCIIWords(t *testing.T) {
	got := markdownText("H\ne\na\nd\ne\nr\n3\n\n\u9879\u76ee\n\u5ba1\u6279")
	if !strings.Contains(got, "Header3") {
		t.Fatalf("markdown did not collapse vertical ASCII word in %q", got)
	}
	if !strings.Contains(got, "\u9879\u76ee\n\u5ba1\u6279") {
		t.Fatalf("markdown changed non-ASCII short lines in %q", got)
	}
}

func TestPPTXMarkdownCollapsesVerticalASCIIWords(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody>
<a:p><a:r><a:t>H</a:t></a:r></a:p><a:p><a:r><a:t>e</a:t></a:r></a:p><a:p><a:r><a:t>a</a:t></a:r></a:p><a:p><a:r><a:t>d</a:t></a:r></a:p><a:p><a:r><a:t>e</a:t></a:r></a:p><a:p><a:r><a:t>r</a:t></a:r></a:p><a:p><a:r><a:t>3</a:t></a:r></a:p>
</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "vertical-text.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Header3") {
		t.Fatalf("markdown did not collapse PPTX vertical ASCII word in:\n%s", md)
	}
	if strings.Contains(md, "H\ne\na\nd\ne\nr\n3") {
		t.Fatalf("markdown kept split vertical ASCII word in:\n%s", md)
	}
}

func TestPPTXMarkdownUsesBulletAndNumberedLists(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
<p:sp><p:txBody>
<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>First bullet</a:t></a:r></a:p>
<a:p><a:pPr lvl="1"><a:buChar char="–"/></a:pPr><a:r><a:t>Nested bullet</a:t></a:r></a:p>
<a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Numbered item</a:t></a:r></a:p>
<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>Plain paragraph</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="2" hidden="1"/></p:nvSpPr><p:txBody>
<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>Hidden bullet secret</a:t></a:r></a:p>
</p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "pptx-lists.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"First bullet", "Nested bullet", "Numbered item", "Plain paragraph"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX list text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Hidden bullet secret") || strings.Contains(res.Text, "- First bullet") {
		t.Fatalf("plain text leaked hidden text or markdown markers: %q", res.Text)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"## Slide 1",
		"- First bullet",
		"- Nested bullet",
		"1. Numbered item",
		"Plain paragraph",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing PPTX list content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden bullet secret", "- Plain paragraph"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden list content or wrong bullet marker %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXMarkdownUsesVisibleSections(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body Text</w:t><w:footnoteReference w:id="2"/><w:commentReference w:id="7"/></w:r></w:p><w:del><w:p><w:r><w:t>Hidden Deleted Secret</w:t></w:r></w:p></w:del><c:chart r:id="rIdChart"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>`)
	addZip(t, zw, "word/header1.xml", `<w:hdr xmlns:w="urn:x"><w:p><w:r><w:t>Visible Header Text</w:t></w:r></w:p></w:hdr>`)
	addZip(t, zw, "word/footnotes.xml", `<w:footnotes xmlns:w="urn:x"><w:footnote w:id="2"><w:p><w:r><w:t>Visible Footnote Text</w:t></w:r></w:p></w:footnote><w:footnote w:id="3"><w:p><w:r><w:t>Unreferenced Footnote Secret</w:t></w:r></w:p></w:footnote></w:footnotes>`)
	addZip(t, zw, "word/comments.xml", `<w:comments xmlns:w="urn:x"><w:comment w:id="7"><w:p><w:r><w:t>Visible Comment Text</w:t></w:r></w:p></w:comment><w:comment w:id="8"><w:p><w:r><w:t>Unreferenced Comment Secret</w:t></w:r></w:p></w:comment></w:comments>`)
	addZip(t, zw, "word/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Visible Chart Target: ../media/docx-chart.png Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image Text</a:t></a:r></a:p></c:title></c:chartSpace>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "structured-docx.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"## Document",
		"Visible Body Text",
		"## Headers and Footers",
		"Visible Header Text",
		"## Footnotes",
		"Visible Footnote Text",
		"## Comments",
		"Visible Comment Text",
		"## Drawings",
		"Visible Chart Text",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing DOCX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Deleted Secret", "Unreferenced Footnote Secret", "Unreferenced Comment Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden or unreferenced DOCX content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXMarkdownUsesParagraphHeadingStyles(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body>
<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Visible Document Title</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Visible Chapter</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Visible Subsection</w:t></w:r></w:p>
<w:p><w:r><w:t>Visible body text</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>Hidden Styled Heading</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Styled Table Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
</w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "heading-styles.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(res.Text, "#") {
		t.Fatalf("plain text should not include markdown heading markers: %q", res.Text)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"# Visible Document Title",
		"## Document",
		"### Visible Chapter",
		"##### Visible Subsection",
		"Visible body text",
		"| Styled Table Cell |",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing styled heading/table content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Styled Heading", "### Styled Table Cell"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden content or promoted table cell heading %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXMarkdownUsesNumberingLists(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Bullet item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Nested bullet item</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Numbered item</w:t></w:r></w:p>
<w:p><w:r><w:t>Plain paragraph</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>Hidden numbered secret</w:t></w:r></w:p>
</w:body></w:document>`)
	addZip(t, zw, "word/numbering.xml", `<w:numbering xmlns:w="urn:w">
<w:abstractNum w:abstractNumId="10">
<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="20">
<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="20"/></w:num>
</w:numbering>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "numbering-lists.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Bullet item", "Nested bullet item", "Numbered item", "Plain paragraph"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible DOCX list text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Hidden numbered secret") || strings.Contains(res.Text, "- Bullet item") || strings.Contains(res.Text, "1. Numbered item") {
		t.Fatalf("plain text leaked hidden content or markdown list markers: %q", res.Text)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"## Document",
		"- Bullet item",
		"- Nested bullet item",
		"1. Numbered item",
		"Plain paragraph",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing DOCX list content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden numbered secret", "- Plain paragraph"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden list content or wrong marker %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXSystemXMLSubtreesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:x="urn:x"><w:body>
<w:p><w:r><w:t>Visible body text</w:t></w:r></w:p>
<script><text>Hidden script text</text><v>Hidden script value</v></script>
<style><text>Hidden style text</text><v>Hidden style value</v></style>
<x:ClientData><x:text>Hidden client data text</x:text><x:v>Hidden client data value</x:v></x:ClientData>
</w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "system-subtrees.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible body text") {
		t.Fatalf("missing visible DOCX body text in %q", res.Text)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible body text") {
		t.Fatalf("markdown missing visible DOCX body text in:\n%s", md)
	}
	for _, hidden := range []string{"Hidden script text", "Hidden script value", "Hidden style text", "Hidden style value", "Hidden client data text", "Hidden client data value"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("text kept system XML subtree content %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept system XML subtree content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXUnreferencedHeaderFooterContentIsNotVisible(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`)
	addZip(t, zw, "word/header1.xml", `<w:hdr xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:t>Visible Header</w:t></w:r></w:p><p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible Header Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic></w:hdr>`)
	addZip(t, zw, "word/_rels/header1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/visible-header.png"/></Relationships>`)
	addZip(t, zw, "word/header2.xml", `<w:hdr xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:t>Orphan Header Secret</w:t></w:r></w:p><p:pic><p:nvPicPr><p:cNvPr id="2" descr="Orphan Header Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdOrphan"/></p:blipFill></p:pic></w:hdr>`)
	addZip(t, zw, "word/_rels/header2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOrphan" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/orphan-header.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible-header.png", testPNG())
	addZipBytes(t, zw, "word/media/orphan-header.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "referenced-header-only.docx")
	outDir := filepath.Join(dir, "images")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Body", "Visible Header"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible DOCX text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Orphan Header Secret") {
		t.Fatalf("kept unreferenced header text in %q", res.Text)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible-header.png" || res.Images[0].Alt != "Visible Header Picture" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only visible referenced header image, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "orphan-header.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("orphan header image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Visible Body", "## Headers and Footers", "Visible Header", "![Visible Header Picture](images/visible-header.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible DOCX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Orphan Header Secret", "Orphan Header Picture", "orphan-header.jpg"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced header content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXUnreferencedNotesAndCommentsAreNotVisible(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible Body Only</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/footnotes.xml", `<w:footnotes xmlns:w="urn:x"><w:footnote w:id="2"><w:p><w:r><w:t>Unreferenced Footnote Secret</w:t></w:r></w:p></w:footnote></w:footnotes>`)
	addZip(t, zw, "word/endnotes.xml", `<w:endnotes xmlns:w="urn:x"><w:endnote w:id="3"><w:p><w:r><w:t>Unreferenced Endnote Secret</w:t></w:r></w:p></w:endnote></w:endnotes>`)
	addZip(t, zw, "word/comments.xml", `<w:comments xmlns:w="urn:x"><w:comment w:id="4"><w:p><w:r><w:t>Unreferenced Comment Secret</w:t></w:r></w:p></w:comment></w:comments>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-notes-comments.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Body Only") {
		t.Fatalf("missing visible body text in %q", res.Text)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible Body Only") {
		t.Fatalf("markdown missing visible body text in:\n%s", md)
	}
	for _, hidden := range []string{"Unreferenced Footnote Secret", "Unreferenced Endnote Secret", "Unreferenced Comment Secret", "## Footnotes", "## Endnotes", "## Comments"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("kept unreferenced DOCX note/comment content %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXNotesAndCommentsUseOnlyVisibleReferences(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body>
<w:p><w:r><w:t>Visible body</w:t><w:footnoteReference w:id="2"/><w:endnoteReference w:id="5"/><w:commentReference w:id="7"/><w:commentReference w:id="12"/></w:r></w:p>
<w:p><w:commentRangeStart w:id="10"/><w:r><w:t>Visible range-comment anchor</w:t></w:r><w:commentRangeEnd w:id="10"/></w:p>
<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden anchor</w:t><w:commentReference w:id="8"/></w:r></w:p>
<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:commentRangeStart w:id="11"/><w:t>Hidden range anchor</w:t></w:r><w:commentRangeEnd w:id="11"/></w:p>
</w:body></w:document>`)
	addZip(t, zw, "word/footnotes.xml", `<w:footnotes xmlns:w="urn:w">
<w:footnote w:id="2"><w:p><w:r><w:t>Visible Footnote Target: ../media/footnote.png Body</w:t></w:r></w:p></w:footnote>
<w:footnote w:id="3"><w:p><w:r><w:t>Unreferenced Footnote Secret</w:t></w:r></w:p></w:footnote>
</w:footnotes>`)
	addZip(t, zw, "word/endnotes.xml", `<w:endnotes xmlns:w="urn:w">
<w:endnote w:id="5"><w:p><w:r><w:t>Visible Endnote Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image Body</w:t></w:r></w:p></w:endnote>
<w:endnote w:id="6"><w:p><w:r><w:t>Unreferenced Endnote Secret</w:t></w:r></w:p></w:endnote>
</w:endnotes>`)
	addZip(t, zw, "word/comments.xml", `<w:comments xmlns:w="urn:w" xmlns:w14="urn:w14">
<w:comment w:id="7"><w:p w14:paraId="AAAA0001"><w:r><w:t>Visible Comment r:embed: rId8 Body</w:t></w:r></w:p></w:comment>
<w:comment w:id="8"><w:p><w:r><w:t>Hidden Anchor Comment Secret</w:t></w:r></w:p></w:comment>
<w:comment w:id="9"><w:p><w:r><w:t>Unreferenced Comment Secret</w:t></w:r></w:p></w:comment>
<w:comment w:id="10"><w:p><w:r><w:t>Visible Range Comment PartName: /word/document.xml Body</w:t></w:r></w:p></w:comment>
<w:comment w:id="11"><w:p><w:r><w:t>Hidden Range Comment Secret</w:t></w:r></w:p></w:comment>
<w:comment w:id="12"><w:p w14:paraId="BBBB0002"><w:r><w:t>Resolved Comment Secret</w:t></w:r></w:p></w:comment>
</w:comments>`)
	addZip(t, zw, "word/commentsExtended.xml", `<w15:commentsEx xmlns:w15="urn:w15"><w15:commentEx w15:paraId="BBBB0002" w15:done="1"/></w15:commentsEx>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "visible-note-references.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible body", "Visible range-comment anchor", "Visible Footnote Body", "Visible Endnote Body", "Visible Comment Body", "Visible Range Comment Body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible referenced content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{
		"Unreferenced Footnote Secret",
		"Unreferenced Endnote Secret",
		"Hidden Anchor Comment Secret",
		"Unreferenced Comment Secret",
		"Hidden Range Comment Secret",
		"Resolved Comment Secret",
		"Hidden anchor",
		"Hidden range anchor",
		"Target:",
		"../media/footnote.png",
		"relationships/image",
		"r:embed:",
		"rId8",
		"PartName:",
		"/word/document.xml",
	} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden/unreferenced note or comment content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Footnotes", "### Footnote 2", "Visible Footnote Body", "## Endnotes", "### Endnote 5", "Visible Endnote Body", "## Comments", "### Comment 7", "Visible Comment Body", "### Comment 10", "Visible Range Comment Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible referenced content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Unreferenced Footnote Secret", "Unreferenced Endnote Secret", "Hidden Anchor Comment Secret", "Unreferenced Comment Secret", "Hidden Range Comment Secret", "Resolved Comment Secret", "### Comment 12", "Target:", "../media/footnote.png", "relationships/image", "r:embed:", "rId8", "PartName:", "/word/document.xml"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden/unreferenced content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXMalformedNoteAndCommentIDsAreNotMarkdownHeadings(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible body</w:t><w:footnoteReference w:id="2 Target: ../media/footnote-id.png"/><w:commentReference w:id="7 rId77"/></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/footnotes.xml", `<w:footnotes xmlns:w="urn:w"><w:footnote w:id="2 Target: ../media/footnote-id.png"><w:p><w:r><w:t>Visible malformed footnote body</w:t></w:r></w:p></w:footnote></w:footnotes>`)
	addZip(t, zw, "word/comments.xml", `<w:comments xmlns:w="urn:w"><w:comment w:id="7 rId77"><w:p><w:r><w:t>Visible malformed comment body</w:t></w:r></w:p></w:comment></w:comments>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "malformed-note-comment-ids.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible body", "Visible malformed footnote body", "Visible malformed comment body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing malformed note/comment content %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Footnotes", "Visible malformed footnote body", "## Comments", "Visible malformed comment body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing malformed note/comment content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"### Footnote 2", "### Comment 7", "../media/footnote-id.png", "footnote-id.png", "rId77", "Target:"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept malformed note/comment id content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXUnreferencedRelatedTextPartsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p><c:chart r:id="rIdChart"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>`)
	addZip(t, zw, "word/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Visible Chart Text</a:t></a:r></a:p></c:title></c:chartSpace>`)
	addZip(t, zw, "word/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Internal Unreferenced Chart Secret</a:t></a:r></a:p></c:title></c:chartSpace>`)
	addZip(t, zw, "word/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Internal Unreferenced Drawing Secret"/></xdr:wsDr>`)
	addZip(t, zw, "word/diagrams/data1.xml", `<dgm:dataModel xmlns:dgm="urn:dgm" xmlns:a="urn:a"><dgm:pt><dgm:t><a:p><a:r><a:t>Internal Unreferenced SmartArt Secret</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:dataModel>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-docx-related-text.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Body Text", "Visible Chart Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible DOCX related text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced Chart Secret", "Internal Unreferenced Drawing Secret", "Internal Unreferenced SmartArt Secret", "Target:", "../media/docx-chart.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced DOCX related text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Visible Body Text", "## Drawings", "Visible Chart Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible DOCX related content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced Chart Secret", "Internal Unreferenced Drawing Secret", "Internal Unreferenced SmartArt Secret", "Target:", "../media/docx-chart.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced DOCX related content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXHiddenDrawingRelatedTextPartsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p>
<p:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdVisibleChart"/></a:graphicData></a:graphic></p:graphicFrame>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Hidden Chart Frame" hidden="1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdHiddenChart"/></a:graphicData></a:graphic></p:graphicFrame>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisibleChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/><Relationship Id="rIdHiddenChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart2.xml"/><Relationship Id="rIdUnreferencedChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart3.xml"/></Relationships>`)
	addZip(t, zw, "word/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Visible Chart Target: ../media/docx-chart.png Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image Text</a:t></a:r></a:p></c:title></c:chartSpace>`)
	addZip(t, zw, "word/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Hidden Drawing Chart Secret</a:t></a:r></a:p></c:title></c:chartSpace>`)
	addZip(t, zw, "word/charts/chart3.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Unreferenced Chart Secret</a:t></a:r></a:p></c:title></c:chartSpace>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-drawing-related-text.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Body Text", "Visible Chart Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible DOCX related text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Drawing Chart Secret", "Unreferenced Chart Secret", "Target:", "../media/docx-chart.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden DOCX related text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Visible Body Text", "## Drawings", "Visible Chart Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible DOCX related content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Drawing Chart Secret", "Unreferenced Chart Secret", "Target:", "../media/docx-chart.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden DOCX related content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXSystemFootnoteSeparatorsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/footnotes.xml", `<w:footnotes xmlns:w="urn:x">
<w:footnote w:type="separator"><w:p><w:r><w:t>System Footnote Separator Text</w:t></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator"><w:p><w:r><w:t>System Footnote Continuation Text</w:t></w:r></w:p></w:footnote>
<w:footnote><w:p><w:r><w:t>Visible Footnote Text</w:t></w:r></w:p></w:footnote>
</w:footnotes>`)
	addZip(t, zw, "word/endnotes.xml", `<w:endnotes xmlns:w="urn:x">
<w:endnote w:type="separator"><w:p><w:r><w:t>System Endnote Separator Text</w:t></w:r></w:p></w:endnote>
<w:endnote w:type="continuationNotice"><w:p><w:r><w:t>System Endnote Continuation Notice</w:t></w:r></w:p></w:endnote>
<w:endnote><w:p><w:r><w:t>Visible Endnote Text</w:t></w:r></w:p></w:endnote>
</w:endnotes>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "system-footnotes.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Body Text", "Visible Footnote Text", "Visible Endnote Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible DOCX footnote content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"System Footnote Separator Text", "System Footnote Continuation Text", "System Endnote Separator Text", "System Endnote Continuation Notice"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept system DOCX footnote content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Body Text", "Visible Footnote Text", "Visible Endnote Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible DOCX footnote content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"System Footnote Separator Text", "System Footnote Continuation Text", "System Endnote Separator Text", "System Endnote Continuation Notice"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept system DOCX footnote content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXOnlyReferencedAltChunkHTMLIsVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk.htm", `<html><head><title>Hidden Head Title</title><style>.cssHidden{display:none}.cssInvisible{visibility:hidden}.cssCollapsed{visibility:collapse}#cssMsoHidden{mso-hide:all}#hiddenOpenDialog{display:none}.hiddenClosedDetails{display:none}.cssOpacityHidden{opacity:0}.cssFontZero{font-size:0px}.cssTransparent{color:transparent}.cssClipped{position:absolute;clip:rect(0 0 0 0)}.cssClipPath{clip-path:inset(50%)}.cssOverflowHidden{overflow:hidden;width:0;height:0}.cssClipOverflowHidden{overflow:clip;width:0;height:0}.cssMaxHeightHidden{overflow-y:hidden;max-height:0}.cssContentVisibilityHidden{content-visibility:hidden}.cssTransformHidden{transform:scale(0)}.cssScaleHidden{scale:0 0}.compoundOne.compoundTwo{display:none}p.elementCompound{display:none}#compoundID.compoundRequired{display:none}.complexHidden img{display:none}</style></head><body>
		<p>Visible AltChunk HTML Text</p>
		<template><style>.templateHidesVisible{display:none}</style><p>Hidden Template HTML Secret</p></template>
		<p class="templateHidesVisible">Visible Template CSS Unapplied Text</p>
		<p>Visible HTML Target: ../media/html-target.png text</p>
		<p>Visible HTML Content-Location: word/media/mhtml-hidden.png text</p>
		<p>Visible HTML Content-Type: image/png text</p>
		<p>Visible HTML Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image text</p>
		<p style="background:url(word/media/style-attr.png)">Visible HTML styled body</p>
		<p style="opacity:0.5">Visible Semi Transparent HTML Text</p>
		<select><option>Hidden Draft HTML Choice</option><option selected>Visible Selected HTML Choice</option></select>
		<select><option>Visible Default HTML Choice</option><option>Hidden Secondary HTML Choice</option></select>
		<select><option hidden>Hidden Option HTML Choice</option><option selected style="display:none">Hidden Selected Style HTML Choice</option><option>Visible Fallback HTML Choice</option></select>
		<select><option label="Visible Label HTML Choice Target: ../media/option-label.png"></option><option value="Hidden Option Value Secret"></option></select>
		<select><optgroup label="Visible Group HTML Label rId88"><option selected>Visible Grouped HTML Choice</option></optgroup><optgroup hidden label="Hidden Group HTML Label"><option selected>Hidden Grouped HTML Choice</option></optgroup></select>
		<input value="Visible Default Input Value rId89">
		<input type="search" value="Visible Search Input Value Target: ../media/search.png">
		<input type="date" value="2026-06-28">
		<input type="button" value="Visible Button Input Value">
		<input placeholder="Visible Placeholder Input Text ContentType: image/png">
		<input type="hidden" value="Hidden Input Value">
		<input type="password" value="Hidden Password Input Value">
		<input type="checkbox" checked value="Hidden Checkbox Internal Value">
		<input type="radio" checked value="Hidden Radio Internal Value">
		<input style="display:none" value="Hidden Styled Input Value" placeholder="Hidden Styled Placeholder Input Text">
		<datalist id="internalSuggestions"><option>Hidden Datalist Option Text</option><option value="Hidden Datalist Value"></option></datalist>
		<textarea>Visible Textarea Line 1&lt;br&gt;Visible Textarea Line 2 &amp; More</textarea>
		<textarea placeholder="Visible Textarea Placeholder Target=&quot;word/media/textarea.png&quot;"></textarea>
		<textarea hidden>Hidden Textarea Value</textarea>
		<textarea hidden placeholder="Hidden Textarea Placeholder"></textarea>
		<textarea style="display:none" placeholder="Hidden Styled Textarea Placeholder">Hidden Styled Textarea Value</textarea>
		<details><summary>Visible Closed Details Summary</summary><p>Hidden Closed Details Body</p></details>
		<details open><summary>Visible Open Details Summary</summary><p>Visible Open Details Body</p></details>
		<details class="hiddenClosedDetails"><summary>Hidden CSS Closed Details Summary</summary><p>Hidden CSS Closed Details Body</p></details>
		<details><p>Hidden No Summary Details Body</p></details>
		<dialog><p>Hidden Closed Dialog Body</p></dialog>
		<dialog open><p>Visible Open Dialog Body</p></dialog>
		<dialog open id="hiddenOpenDialog"><p>Hidden CSS Open Dialog Body</p></dialog>
		<object data="word/media/object.bin"><p>Hidden Object Fallback HTML Secret</p></object>
		<iframe src="word/media/frame.htm"><p>Hidden Iframe Fallback HTML Secret</p></iframe>
		<embed src="word/media/embed.bin" alt="Hidden Embed Attr HTML Secret">
		<div style="display:none">Hidden Display None HTML Secret</div>
		<section style="visibility:hidden"><p>Hidden Visibility HTML Secret</p></section>
		<section style="visibility:collapse"><p>Hidden Collapsed HTML Secret</p></section>
		<span hidden>Hidden Boolean HTML Secret</span>
		<div aria-hidden="true">Hidden Aria HTML Secret</div>
		<p style="mso-hide:all">Hidden MSO HTML Secret</p>
		<p style="opacity: 0 !important">Hidden Opacity HTML Secret</p>
		<p style="font-size: 0">Hidden Font Zero HTML Secret</p>
		<p style="color: transparent">Hidden Transparent HTML Secret</p>
		<p style="position:absolute; clip: rect(0px, 0px, 0px, 0px)">Hidden Clipped HTML Secret</p>
		<p style="clip-path: inset(50%)">Hidden Clip Path HTML Secret</p>
		<div style="overflow:hidden;width:0;height:0">Hidden Overflow Collapsed HTML Secret</div>
		<div style="overflow:clip;width:0;height:0">Hidden Clip Overflow HTML Secret</div>
		<div style="overflow-y:hidden;max-height:0">Hidden Max Height HTML Secret</div>
		<div style="content-visibility:hidden">Hidden Content Visibility HTML Secret</div>
		<div style="transform:scale(0, 0)">Hidden Transform Scale HTML Secret</div>
		<div style="scale:0">Hidden Scale Property HTML Secret</div>
		<div style="DISPLAY : none !important">Hidden Spaced Display HTML Secret</div>
		<div style="visibility : hidden">Hidden Spaced Visibility HTML Secret</div>
		<div style="visibility : collapse">Hidden Spaced Collapsed HTML Secret</div>
		<div style="mso-hide : all">Hidden Spaced MSO HTML Secret</div>
		<div class="cssHidden">Hidden CSS Class HTML Secret</div>
		<p class="cssInvisible">Hidden CSS Visibility HTML Secret</p>
		<p class="cssCollapsed">Hidden CSS Collapsed HTML Secret</p>
		<section id="cssMsoHidden">Hidden CSS ID HTML Secret</section>
		<p class="cssOpacityHidden">Hidden CSS Opacity HTML Secret</p>
		<p class="cssFontZero">Hidden CSS Font Zero HTML Secret</p>
		<p class="cssTransparent">Hidden CSS Transparent HTML Secret</p>
		<p class="cssClipped">Hidden CSS Clipped HTML Secret</p>
		<p class="cssClipPath">Hidden CSS Clip Path HTML Secret</p>
		<p class="cssOverflowHidden">Hidden CSS Overflow HTML Secret</p>
		<p class="cssClipOverflowHidden">Hidden CSS Clip Overflow HTML Secret</p>
		<p class="cssMaxHeightHidden">Hidden CSS Max Height HTML Secret</p>
		<p class="cssContentVisibilityHidden">Hidden CSS Content Visibility HTML Secret</p>
		<p class="cssTransformHidden">Hidden CSS Transform HTML Secret</p>
		<p class="cssScaleHidden">Hidden CSS Scale HTML Secret</p>
		<p class="compoundOne">Visible Single Compound Class Text</p>
		<p class="compoundOne compoundTwo">Hidden Compound Class HTML Secret</p>
		<div class="elementCompound">Visible Different Element Compound Text</div>
		<p class="elementCompound">Hidden Element Compound HTML Secret</p>
		<p id="compoundID">Visible ID Missing Compound Class Text</p>
		<p id="compoundID" class="compoundRequired">Hidden ID Compound HTML Secret</p>
		<div class="complexHidden"><p>Visible Complex CSS Selector Text</p></div>
		<!-- Hidden Plain HTML Comment Secret word/media/comment-hidden.png -->
		<!--[if gte mso 9]><xml><w:WordDocument>Hidden Conditional HTML Secret Target: ../media/conditional.png</w:WordDocument></xml><![endif]-->
		<img src="word/media/attr-hidden.png" alt="Hidden Attr Secret">
		<p>word/media/hidden-html.png</p>
		<script>Hidden Script Secret word/media/script-hidden.png</script>
		<style>.x{background:url(word/media/style-hidden.png)} Hidden Style Secret</style>
		<w:ClientData>Hidden ClientData Secret</w:ClientData>
		<p>Visible HTML Tail</p>
	</body></html>`)
	addZip(t, zw, "word/orphan.htm", `<html><body><p>Orphan Internal HTML Secret</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "referenced-altchunk.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Body Text", "Visible AltChunk HTML Text", "Visible Template CSS Unapplied Text", "Visible HTML text", "Visible HTML styled body", "Visible Semi Transparent HTML Text", "Visible Selected HTML Choice", "Visible Default HTML Choice", "Visible Fallback HTML Choice", "Visible Label HTML Choice", "Visible Group HTML Label", "Visible Grouped HTML Choice", "Visible Default Input Value", "Visible Search Input Value", "2026-06-28", "Visible Button Input Value", "Visible Placeholder Input Text", "Visible Textarea Line 1", "Visible Textarea Line 2 & More", "Visible Textarea Placeholder", "Visible Closed Details Summary", "Visible Open Details Summary", "Visible Open Details Body", "Visible Open Dialog Body", "Visible Single Compound Class Text", "Visible Different Element Compound Text", "Visible ID Missing Compound Class Text", "Visible Complex CSS Selector Text", "Visible HTML Tail"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible DOCX altChunk content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{
		"Target:", "../media/html-target.png", "../media/option-label.png", "../media/search.png", "word/media/textarea.png", "rId88", "rId89", "ContentType:",
		"Content-Location", "word/media/mhtml-hidden.png",
		"Content-Type", "image/png", "Type:", "relationships/image", "word/media/style-attr.png",
		"word/media/attr-hidden.png", "Hidden Attr Secret", "word/media/hidden-html.png",
		"Hidden Head Title", "Hidden Script Secret", "script-hidden.png", "Hidden Style Secret",
		"Hidden Template HTML Secret", "templateHidesVisible",
		"style-hidden.png", "Hidden ClientData Secret", "Hidden Display None HTML Secret",
		"Hidden Draft HTML Choice", "Hidden Secondary HTML Choice", "Hidden Option HTML Choice", "Hidden Selected Style HTML Choice",
		"Hidden Option Value Secret",
		"Hidden Group HTML Label", "Hidden Grouped HTML Choice",
		"Hidden Input Value", "Hidden Password Input Value", "Hidden Checkbox Internal Value", "Hidden Radio Internal Value", "Hidden Styled Input Value", "Hidden Styled Placeholder Input Text",
		"Hidden Datalist Option Text", "Hidden Datalist Value",
		"Hidden Textarea Value", "Hidden Textarea Placeholder", "Hidden Styled Textarea Value", "Hidden Styled Textarea Placeholder",
		"Hidden Closed Details Body", "Hidden No Summary Details Body",
		"Hidden CSS Closed Details Summary", "Hidden CSS Closed Details Body",
		"Hidden Closed Dialog Body", "Hidden CSS Open Dialog Body",
		"Hidden Object Fallback HTML Secret", "Hidden Iframe Fallback HTML Secret", "Hidden Embed Attr HTML Secret",
		"word/media/object.bin", "word/media/frame.htm", "word/media/embed.bin",
		"Hidden Visibility HTML Secret", "Hidden Collapsed HTML Secret", "Hidden Boolean HTML Secret", "Hidden Aria HTML Secret",
		"Hidden MSO HTML Secret", "Hidden Plain HTML Comment Secret", "Hidden Conditional HTML Secret",
		"Hidden Opacity HTML Secret", "Hidden Font Zero HTML Secret", "Hidden Transparent HTML Secret",
		"Hidden Clipped HTML Secret", "Hidden Clip Path HTML Secret", "Hidden Overflow Collapsed HTML Secret", "Hidden Clip Overflow HTML Secret", "Hidden Max Height HTML Secret",
		"Hidden Content Visibility HTML Secret", "Hidden Transform Scale HTML Secret", "Hidden Scale Property HTML Secret",
		"Hidden Spaced Display HTML Secret", "Hidden Spaced Visibility HTML Secret", "Hidden Spaced Collapsed HTML Secret", "Hidden Spaced MSO HTML Secret",
		"Hidden CSS Class HTML Secret", "Hidden CSS Visibility HTML Secret", "Hidden CSS Collapsed HTML Secret", "Hidden CSS ID HTML Secret",
		"Hidden CSS Opacity HTML Secret", "Hidden CSS Font Zero HTML Secret", "Hidden CSS Transparent HTML Secret",
		"Hidden CSS Clipped HTML Secret", "Hidden CSS Clip Path HTML Secret", "Hidden CSS Overflow HTML Secret", "Hidden CSS Clip Overflow HTML Secret", "Hidden CSS Max Height HTML Secret",
		"Hidden CSS Content Visibility HTML Secret", "Hidden CSS Transform HTML Secret", "Hidden CSS Scale HTML Secret",
		"Hidden Compound Class HTML Secret", "Hidden Element Compound HTML Secret", "Hidden ID Compound HTML Secret",
		"word/media/comment-hidden.png", "../media/conditional.png", "Orphan Internal HTML Secret",
	} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden DOCX altChunk content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Body Text", "## HTML Content", "Visible AltChunk HTML Text", "Visible Template CSS Unapplied Text", "Visible HTML text", "Visible HTML styled body", "Visible Semi Transparent HTML Text", "Visible Selected HTML Choice", "Visible Default HTML Choice", "Visible Fallback HTML Choice", "Visible Label HTML Choice", "Visible Group HTML Label", "Visible Grouped HTML Choice", "Visible Default Input Value", "Visible Search Input Value", "2026-06-28", "Visible Button Input Value", "Visible Placeholder Input Text", "Visible Textarea Line 1", "Visible Textarea Line 2 & More", "Visible Textarea Placeholder", "Visible Closed Details Summary", "Visible Open Details Summary", "Visible Open Details Body", "Visible Open Dialog Body", "Visible Single Compound Class Text", "Visible Different Element Compound Text", "Visible ID Missing Compound Class Text", "Visible Complex CSS Selector Text", "Visible HTML Tail"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible DOCX altChunk content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{
		"Target:", "../media/html-target.png", "../media/option-label.png", "../media/search.png", "word/media/textarea.png", "rId88", "rId89", "ContentType:",
		"Content-Location", "word/media/mhtml-hidden.png",
		"Content-Type", "image/png", "Type:", "relationships/image", "word/media/style-attr.png",
		"word/media/attr-hidden.png", "Hidden Attr Secret", "word/media/hidden-html.png",
		"Hidden Head Title", "Hidden Script Secret", "script-hidden.png", "Hidden Style Secret",
		"Hidden Template HTML Secret", "templateHidesVisible",
		"style-hidden.png", "Hidden ClientData Secret", "Hidden Display None HTML Secret",
		"Hidden Draft HTML Choice", "Hidden Secondary HTML Choice", "Hidden Option HTML Choice", "Hidden Selected Style HTML Choice",
		"Hidden Option Value Secret",
		"Hidden Group HTML Label", "Hidden Grouped HTML Choice",
		"Hidden Input Value", "Hidden Password Input Value", "Hidden Checkbox Internal Value", "Hidden Radio Internal Value", "Hidden Styled Input Value", "Hidden Styled Placeholder Input Text",
		"Hidden Datalist Option Text", "Hidden Datalist Value",
		"Hidden Textarea Value", "Hidden Textarea Placeholder", "Hidden Styled Textarea Value", "Hidden Styled Textarea Placeholder",
		"Hidden Closed Details Body", "Hidden No Summary Details Body",
		"Hidden CSS Closed Details Summary", "Hidden CSS Closed Details Body",
		"Hidden Closed Dialog Body", "Hidden CSS Open Dialog Body",
		"Hidden Object Fallback HTML Secret", "Hidden Iframe Fallback HTML Secret", "Hidden Embed Attr HTML Secret",
		"word/media/object.bin", "word/media/frame.htm", "word/media/embed.bin",
		"Hidden Visibility HTML Secret", "Hidden Collapsed HTML Secret", "Hidden Boolean HTML Secret", "Hidden Aria HTML Secret",
		"Hidden MSO HTML Secret", "Hidden Plain HTML Comment Secret", "Hidden Conditional HTML Secret",
		"Hidden Opacity HTML Secret", "Hidden Font Zero HTML Secret", "Hidden Transparent HTML Secret",
		"Hidden Clipped HTML Secret", "Hidden Clip Path HTML Secret", "Hidden Overflow Collapsed HTML Secret", "Hidden Clip Overflow HTML Secret", "Hidden Max Height HTML Secret",
		"Hidden Content Visibility HTML Secret", "Hidden Transform Scale HTML Secret", "Hidden Scale Property HTML Secret",
		"Hidden Spaced Display HTML Secret", "Hidden Spaced Visibility HTML Secret", "Hidden Spaced Collapsed HTML Secret", "Hidden Spaced MSO HTML Secret",
		"Hidden CSS Class HTML Secret", "Hidden CSS Visibility HTML Secret", "Hidden CSS Collapsed HTML Secret", "Hidden CSS ID HTML Secret",
		"Hidden CSS Opacity HTML Secret", "Hidden CSS Font Zero HTML Secret", "Hidden CSS Transparent HTML Secret",
		"Hidden CSS Clipped HTML Secret", "Hidden CSS Clip Path HTML Secret", "Hidden CSS Overflow HTML Secret", "Hidden CSS Clip Overflow HTML Secret", "Hidden CSS Max Height HTML Secret",
		"Hidden CSS Content Visibility HTML Secret", "Hidden CSS Transform HTML Secret", "Hidden CSS Scale HTML Secret",
		"Hidden Compound Class HTML Secret", "Hidden Element Compound HTML Secret", "Hidden ID Compound HTML Secret",
		"word/media/comment-hidden.png", "../media/conditional.png", "Orphan Internal HTML Secret",
	} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden DOCX altChunk content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXAltChunkHTMLImageIsVisibleAndSniffed(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk.htm", `<html><head><style>.cssHiddenImage{display:none}#cssHiddenImage{visibility:hidden}#hiddenOpenImageDialog{display:none}.hiddenClosedImageDetails{display:none}.cssCollapseImage{visibility:collapse}.cssOverflowImage{overflow:hidden;width:0;height:0}.cssClipOverflowImage{overflow:clip;width:0;height:0}.cssOpacityImage{opacity:0}.cssClipPathImage{clip-path:inset(50%)}.cssContentVisibilityImage{content-visibility:hidden}.cssTransformImage{transform:scale(0)}.cssScaleImage{scale:0 0}.compoundImage.requiredImage{display:none}img.elementImage{display:none}#compoundImageID.requiredImage{display:none}</style></head><body><p>Visible HTML before image</p><img src="media/html-visible.jpg?version=1#preview" alt="Visible HTML Diagram Target: ../media/hidden.png"><img srcset="media/srcset-visible.jpg 1x" alt="Visible Srcset Diagram"><picture><source srcset="media/picture-source-visible.jpg 1x"><img alt="Visible Picture Source Diagram"></picture><picture style="display:none"><source srcset="media/picture-source-hidden.jpg 1x"><img alt="Hidden Picture Source Image"></picture><img src="media/compound-visible.jpg" class="compoundImage" alt="Visible Compound Single Image"><img src="media/compound-hidden.jpg" class="compoundImage requiredImage" alt="Hidden Compound Image"><div class="elementImage"><img src="media/element-visible.jpg" alt="Visible Element Mismatch Image"></div><img src="media/element-hidden.jpg" class="elementImage" alt="Hidden Element Image"><img src="media/id-compound-visible.jpg" id="compoundImageID" alt="Visible ID Missing Class Image"><img src="media/id-compound-hidden.jpg" id="compoundImageID" class="requiredImage" alt="Hidden ID Compound Image"><details><summary>Visible Details Image Summary</summary><img src="media/details-hidden.jpg" alt="Hidden Closed Details Image"></details><details open><summary>Visible Open Image Summary</summary><img src="media/details-open-visible.jpg" alt="Visible Open Details Image"></details><details class="hiddenClosedImageDetails"><summary>Hidden CSS Closed Details Image Summary</summary><img src="media/details-css-hidden.jpg" alt="Hidden CSS Closed Details Image"></details><dialog><img src="media/dialog-hidden.jpg" alt="Hidden Closed Dialog Image"></dialog><dialog open><img src="media/dialog-open-visible.jpg" alt="Visible Open Dialog Image"></dialog><dialog open id="hiddenOpenImageDialog"><img src="media/dialog-css-hidden.jpg" alt="Hidden CSS Open Dialog Image"></dialog><object data="media/object.bin"><img src="media/object-hidden.jpg" alt="Hidden Object Fallback Image"></object><iframe src="media/frame.htm"><img src="media/iframe-hidden.jpg" alt="Hidden Iframe Fallback Image"></iframe><template><img src="media/template-hidden.jpg" alt="Hidden Template Image"></template><datalist><option><img src="media/datalist-hidden.jpg" alt="Hidden Datalist Image"></option></datalist><img srcset="media/srcset-hidden.jpg 1x" alt="Hidden Srcset Image" style="display:none"><img src="media/style-hidden.jpg" alt="Hidden Style Image" style="display : none !important"><img src="media/collapse-hidden.jpg" alt="Hidden Collapse Image" style="visibility: collapse"><img src="media/overflow-hidden.jpg" alt="Hidden Overflow Image" style="overflow:hidden;width:0;height:0"><img src="media/clip-overflow-hidden.jpg" alt="Hidden Clip Overflow Image" style="overflow:clip;width:0;height:0"><img src="media/maxheight-hidden.jpg" alt="Hidden Max Height Image" style="overflow-y:hidden;max-height:0"><img src="media/content-visibility-hidden.jpg" alt="Hidden Content Visibility Image" style="content-visibility:hidden"><img src="media/transform-hidden.jpg" alt="Hidden Transform Image" style="transform:scale(0, 0)"><img src="media/scale-hidden.jpg" alt="Hidden Scale Image" style="scale:0"><img src="media/opacity-hidden.jpg" alt="Hidden Opacity Image" style="opacity:0"><img src="media/clip-hidden.jpg" alt="Hidden Clip Image" style="clip:rect(0 0 0 0)"><img src="media/css-hidden.jpg" alt="Hidden CSS Class Image" class="cssHiddenImage"><img src="media/css-id-hidden.jpg" alt="Hidden CSS ID Image" id="cssHiddenImage"><img src="media/css-collapse-hidden.jpg" alt="Hidden CSS Collapse Image" class="cssCollapseImage"><img src="media/css-overflow-hidden.jpg" alt="Hidden CSS Overflow Image" class="cssOverflowImage"><img src="media/css-clip-overflow-hidden.jpg" alt="Hidden CSS Clip Overflow Image" class="cssClipOverflowImage"><img src="media/css-opacity-hidden.jpg" alt="Hidden CSS Opacity Image" class="cssOpacityImage"><img src="media/css-clippath-hidden.jpg" alt="Hidden CSS Clip Path Image" class="cssClipPathImage"><img src="media/css-content-visibility-hidden.jpg" alt="Hidden CSS Content Visibility Image" class="cssContentVisibilityImage"><img src="media/css-transform-hidden.jpg" alt="Hidden CSS Transform Image" class="cssTransformImage"><img src="media/css-scale-hidden.jpg" alt="Hidden CSS Scale Image" class="cssScaleImage"><!-- <img src="media/comment-hidden.jpg" alt="Hidden Comment Image"> --><p>Visible HTML after image</p></body></html>`)
	addZipBytes(t, zw, "word/media/html-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/srcset-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/srcset-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/picture-source-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/picture-source-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/compound-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/compound-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/element-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/element-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/id-compound-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/id-compound-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/details-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/details-open-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/details-css-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/dialog-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/dialog-open-visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/dialog-css-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/object-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/iframe-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/template-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/datalist-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/unreferenced-hidden.png", testPNG())
	addZipBytes(t, zw, "word/media/comment-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/style-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/collapse-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/overflow-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/clip-overflow-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/maxheight-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/content-visibility-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/transform-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/scale-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/opacity-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/clip-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-id-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-collapse-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-overflow-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-clip-overflow-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-opacity-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-clippath-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-content-visibility-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-transform-hidden.jpg", testJPEG())
	addZipBytes(t, zw, "word/media/css-scale-hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 8 {
		t.Fatalf("expected only visible altChunk HTML images, got %#v", res.Images)
	}
	imagesByName := map[string]Image{}
	for _, img := range res.Images {
		imagesByName[img.Name] = img
	}
	for _, w := range []struct {
		name string
		alt  string
	}{
		{name: "html-visible.png", alt: "Visible HTML Diagram"},
		{name: "srcset-visible.png", alt: "Visible Srcset Diagram"},
		{name: "picture-source-visible.png", alt: "Visible Picture Source Diagram"},
		{name: "compound-visible.png", alt: "Visible Compound Single Image"},
		{name: "element-visible.png", alt: "Visible Element Mismatch Image"},
		{name: "id-compound-visible.png", alt: "Visible ID Missing Class Image"},
		{name: "details-open-visible.png", alt: "Visible Open Details Image"},
		{name: "dialog-open-visible.png", alt: "Visible Open Dialog Image"},
	} {
		img, ok := imagesByName[w.name]
		if !ok {
			t.Fatalf("missing visible HTML image %s in %#v", w.name, res.Images)
		}
		if img.Name != w.name || img.Ext != ".png" || img.Alt != w.alt || !validImageData(".png", img.Data) {
			t.Fatalf("unexpected visible HTML image %s: %#v", w.name, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written visible HTML image %s: %v", w.name, err)
		}
	}
	if _, err := os.Stat(filepath.Join(outDir, "unreferenced-hidden.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("unreferenced hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "comment-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("commented HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "srcset-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("srcset-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "picture-source-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("picture-source-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "template-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("template-hidden HTML media should not be written, stat err=%v", err)
	}
	for _, hiddenName := range []string{"compound-hidden.jpg", "element-hidden.jpg", "id-compound-hidden.jpg", "details-hidden.jpg", "details-css-hidden.jpg", "dialog-hidden.jpg", "dialog-css-hidden.jpg", "object-hidden.jpg", "iframe-hidden.jpg", "datalist-hidden.jpg"} {
		if _, err := os.Stat(filepath.Join(outDir, hiddenName)); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("compound-hidden HTML media %s should not be written, stat err=%v", hiddenName, err)
		}
	}
	if _, err := os.Stat(filepath.Join(outDir, "style-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("style-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "collapse-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("collapse-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "overflow-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("overflow-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "clip-overflow-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("clip-overflow-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "maxheight-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("maxheight-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "content-visibility-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("content-visibility-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "transform-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("transform-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "scale-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("scale-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "opacity-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("opacity-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "clip-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("clip-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-id-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS ID-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-collapse-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS collapse-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-overflow-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS overflow-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-clip-overflow-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS clip-overflow-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-opacity-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS opacity-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-clippath-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS clip-path-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-content-visibility-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS content-visibility-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-transform-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS transform-hidden HTML media should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "css-scale-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("CSS scale-hidden HTML media should not be written, stat err=%v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible HTML Diagram](images/html-visible.png)") {
		t.Fatalf("markdown missing visible altChunk HTML image:\n%s", md)
	}
	if !strings.Contains(md, "![Visible Srcset Diagram](images/srcset-visible.png)") {
		t.Fatalf("markdown missing visible srcset altChunk HTML image:\n%s", md)
	}
	if !strings.Contains(md, "![Visible Picture Source Diagram](images/picture-source-visible.png)") {
		t.Fatalf("markdown missing visible picture source altChunk HTML image:\n%s", md)
	}
	for _, want := range []string{"![Visible Compound Single Image](images/compound-visible.png)", "![Visible Element Mismatch Image](images/element-visible.png)", "![Visible ID Missing Class Image](images/id-compound-visible.png)", "![Visible Open Details Image](images/details-open-visible.png)", "![Visible Open Dialog Image](images/dialog-open-visible.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible compound CSS HTML image %q:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/hidden.png", "version=1", "#preview", "unreferenced-hidden.png", "comment-hidden.jpg", "Hidden Comment Image", "srcset-hidden.jpg", "Hidden Srcset Image", "picture-source-hidden.jpg", "Hidden Picture Source Image", "template-hidden.jpg", "Hidden Template Image", "datalist-hidden.jpg", "Hidden Datalist Image", "details-hidden.jpg", "Hidden Closed Details Image", "details-css-hidden.jpg", "Hidden CSS Closed Details Image", "dialog-hidden.jpg", "Hidden Closed Dialog Image", "dialog-css-hidden.jpg", "Hidden CSS Open Dialog Image", "object-hidden.jpg", "Hidden Object Fallback Image", "iframe-hidden.jpg", "Hidden Iframe Fallback Image", "compound-hidden.jpg", "Hidden Compound Image", "element-hidden.jpg", "Hidden Element Image", "id-compound-hidden.jpg", "Hidden ID Compound Image", "style-hidden.jpg", "Hidden Style Image", "collapse-hidden.jpg", "Hidden Collapse Image", "overflow-hidden.jpg", "Hidden Overflow Image", "clip-overflow-hidden.jpg", "Hidden Clip Overflow Image", "maxheight-hidden.jpg", "Hidden Max Height Image", "content-visibility-hidden.jpg", "Hidden Content Visibility Image", "transform-hidden.jpg", "Hidden Transform Image", "scale-hidden.jpg", "Hidden Scale Image", "opacity-hidden.jpg", "Hidden Opacity Image", "clip-hidden.jpg", "Hidden Clip Image", "css-hidden.jpg", "Hidden CSS Class Image", "css-id-hidden.jpg", "Hidden CSS ID Image", "css-collapse-hidden.jpg", "Hidden CSS Collapse Image", "css-overflow-hidden.jpg", "Hidden CSS Overflow Image", "css-clip-overflow-hidden.jpg", "Hidden CSS Clip Overflow Image", "css-opacity-hidden.jpg", "Hidden CSS Opacity Image", "css-clippath-hidden.jpg", "Hidden CSS Clip Path Image", "css-content-visibility-hidden.jpg", "Hidden CSS Content Visibility Image", "css-transform-hidden.jpg", "Hidden CSS Transform Image", "css-scale-hidden.jpg", "Hidden CSS Scale Image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden HTML image content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXAltChunkHTMLDataURIImageIsExtracted(t *testing.T) {
	dataURI := "data:image/jpeg;base64," + base64.StdEncoding.EncodeToString(testPNG())
	hiddenURI := "data:image/png;base64," + base64.StdEncoding.EncodeToString(testPNG())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data.htm", `<html><body><p>Visible data image before</p><img src="`+dataURI+`" alt="Visible Data Image Target: ../media/hidden.png"><object data="media/object.bin"><img src="`+hiddenURI+`" alt="Hidden Object Data Image"></object><iframe src="media/frame.htm"><img src="`+hiddenURI+`" alt="Hidden Iframe Data Image"></iframe><template><img src="`+hiddenURI+`" alt="Hidden Template Data Image"></template><img src="`+hiddenURI+`" alt="Hidden Data Image" style="display : none"><!-- <img src="`+hiddenURI+`" alt="Hidden Comment Data Image"> --><p>Visible data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one visible data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.png" || img.Ext != ".png" || img.Alt != "Visible Data Image" || !validImageData(".png", img.Data) {
		t.Fatalf("unexpected data URI image extraction: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.png")); err != nil {
		t.Fatalf("expected written data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible data image before", "Visible data image after", "![Visible Data Image](images/html-data-image-001.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/hidden.png", "Hidden Data Image", "Hidden Object Data Image", "Hidden Iframe Data Image", "Hidden Template Data Image", "Hidden Comment Data Image", "data:image"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden data URI content %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLImageAltMetadataIsCleanedInResult(t *testing.T) {
	dataURI := "data:image/png;base64," + base64.StdEncoding.EncodeToString(testPNG())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-alt-clean.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-alt-clean.htm", `<html><body><p>Visible clean alt before</p><img src="media/normal.png" alt="Visible Normal Image Content-ID: &lt;hidden@office&gt;"><img src="`+dataURI+`" title="Visible Data Title Content-Location: word/media/hidden.png"><picture><source srcset="media/picture.png 1x"><img alt="Visible Picture Image Content-Type: image/png"></picture><p>Visible clean alt after</p></body></html>`)
	addZipBytes(t, zw, "word/media/normal.png", testPNG())
	addZipBytes(t, zw, "word/media/picture.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-clean-image-alt.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 3 {
		t.Fatalf("expected three HTML images with cleaned alt text, got %#v", res.Images)
	}
	imagesByName := map[string]Image{}
	for _, img := range res.Images {
		imagesByName[img.Name] = img
	}
	for _, want := range []struct {
		name string
		alt  string
		ext  string
	}{
		{name: "normal.png", alt: "Visible Normal Image", ext: ".png"},
		{name: "html-data-image-001.png", alt: "Visible Data Title", ext: ".png"},
		{name: "picture.png", alt: "Visible Picture Image", ext: ".png"},
	} {
		img, ok := imagesByName[want.name]
		if !ok {
			t.Fatalf("missing cleaned HTML image %s in %#v", want.name, res.Images)
		}
		if img.Alt != want.alt || img.Ext != want.ext || !validImageData(want.ext, img.Data) {
			t.Fatalf("unexpected cleaned HTML image %s: %#v", want.name, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, want.name)); err != nil {
			t.Fatalf("expected written cleaned HTML image %s: %v", want.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible clean alt before", "Visible clean alt after", "![Visible Normal Image](images/normal.png)", "![Visible Data Title](images/html-data-image-001.png)", "![Visible Picture Image](images/picture.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing cleaned HTML image content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-ID", "hidden@office", "Content-Location", "word/media/hidden.png", "Content-Type", "data:image"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("kept HTML image alt metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
		for _, img := range res.Images {
			if strings.Contains(img.Alt, hidden) {
				t.Fatalf("kept HTML image alt metadata %q in image %#v", hidden, img)
			}
		}
	}
}

func TestDOCXAltChunkHTMLURLEncodedBase64DataURIImageIsExtracted(t *testing.T) {
	payload := base64.StdEncoding.EncodeToString(testPNG())
	encodedPayload := strings.ReplaceAll(url.PathEscape(payload), "%3D", "%3D\r\n")
	dataURI := "data:image/png;base64," + encodedPayload
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-encoded.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-encoded.htm", `<html><body><p>Encoded data image before</p><img src="`+dataURI+`" alt="Encoded Data Image"><p>Encoded data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-data-uri-encoded.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one encoded data URI image, got %#v", res.Images)
	}
	if res.Images[0].Name != "html-data-image-001.png" || res.Images[0].Alt != "Encoded Data Image" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected encoded data URI image: %#v", res.Images[0])
	}
	md := res.Markdown("images")
	for _, want := range []string{"Encoded data image before", "Encoded data image after", "![Encoded Data Image](images/html-data-image-001.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing encoded data URI content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "data:image") || strings.Contains(res.Text, "data:image") {
		t.Fatalf("kept encoded data URI in text=%q markdown=\n%s", res.Text, md)
	}
}

func TestDOCXAltChunkHTMLAPNGDataURIMimeAliasIsExtracted(t *testing.T) {
	dataURI := "data:image/apng;base64," + base64.StdEncoding.EncodeToString(testPNG())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-apng.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-apng.htm", `<html><body><p>APNG data image before</p><img src="`+dataURI+`" alt="Visible APNG Data"><p>APNG data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-apng-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one APNG data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.png" || img.Ext != ".png" || img.Alt != "Visible APNG Data" || !validImageData(".png", img.Data) {
		t.Fatalf("unexpected APNG data URI image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.png")); err != nil {
		t.Fatalf("expected written APNG data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"APNG data image before", "APNG data image after", "![Visible APNG Data](images/html-data-image-001.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing APNG data URI content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "data:image") || strings.Contains(res.Text, "data:image") {
		t.Fatalf("kept APNG data URI in text=%q markdown=\n%s", res.Text, md)
	}
}

func TestDOCXAltChunkHTMLParameterizedDataURIImageIsExtracted(t *testing.T) {
	dataURI := "DATA:IMAGE/JPEG;name=word/media/hidden.jpg;charset=utf-8;BASE64," + base64.StdEncoding.EncodeToString(testPNG())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-param.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-param.htm", `<html><body><p>Parameterized data image before</p><img src="`+dataURI+`" alt="Parameterized Data Image"><p>Parameterized data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-data-uri-param.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one parameterized data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.png" || img.Ext != ".png" || img.Alt != "Parameterized Data Image" || !validImageData(".png", img.Data) {
		t.Fatalf("unexpected parameterized data URI image: %#v", img)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Parameterized data image before", "Parameterized data image after", "![Parameterized Data Image](images/html-data-image-001.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing parameterized data URI content %q in:\n%s", want, md)
		}
	}
	for _, leaked := range []string{"data:image", "name=", "word/media/hidden.jpg", "hidden.jpg"} {
		if strings.Contains(md, leaked) || strings.Contains(res.Text, leaked) {
			t.Fatalf("kept parameterized data URI metadata %q in text=%q markdown=\n%s", leaked, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLSniffsGenericDataURIImage(t *testing.T) {
	dataURI := "data:application/octet-stream;base64," + base64.StdEncoding.EncodeToString(testPNG())
	hiddenURI := "data:application/octet-stream;base64," + base64.StdEncoding.EncodeToString([]byte("not an image"))
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-generic-data.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-generic-data.htm", `<html><body><p>Generic data image before</p><img src="`+dataURI+`" alt="Generic Data Image"><img src="`+hiddenURI+`" alt="Invalid Generic Data Image"><p>Generic data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-generic-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one sniffed generic data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.png" || img.Ext != ".png" || img.Alt != "Generic Data Image" || !validImageData(".png", img.Data) {
		t.Fatalf("unexpected generic data URI image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.png")); err != nil {
		t.Fatalf("expected written generic data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Generic data image before", "Generic data image after", "![Generic Data Image](images/html-data-image-001.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing generic data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:application/octet-stream", "Invalid Generic Data Image"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept generic data URI metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLIconDataURIMimeAliasIsExtracted(t *testing.T) {
	dataURI := "data:image/vnd.microsoft.icon;base64," + base64.StdEncoding.EncodeToString(testICO())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-icon.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-icon.htm", `<html><body><p>Icon data image before</p><img src="`+dataURI+`" alt="Visible Icon Data"><p>Icon data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-icon-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one icon data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.ico" || img.Ext != ".ico" || img.Alt != "Visible Icon Data" || !validImageData(".ico", img.Data) {
		t.Fatalf("unexpected icon data URI image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.ico")); err != nil {
		t.Fatalf("expected written icon data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Icon data image before", "Icon data image after", "![Visible Icon Data](images/html-data-image-001.ico)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing icon data URI content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "data:image") || strings.Contains(res.Text, "data:image") {
		t.Fatalf("kept icon data URI in text=%q markdown=\n%s", res.Text, md)
	}
}

func TestDOCXAltChunkHTMLRasterDataURIMimeAliasesAreExtracted(t *testing.T) {
	bmp, ok := dibToBMP(testDIB())
	if !ok {
		t.Fatal("failed to build test BMP")
	}
	xjpegURI := "data:image/x-jpeg;base64," + base64.StdEncoding.EncodeToString(testJPEG())
	jfifURI := "data:image/jfif;base64," + base64.StdEncoding.EncodeToString(testJPEG())
	xbmpURI := "data:image/x-bmp;base64," + base64.StdEncoding.EncodeToString(bmp)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-raster-alias.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-raster-alias.htm", `<html><body><p>Raster alias data before</p><img src="`+xjpegURI+`" alt="Visible XJPEG Data"><img src="`+jfifURI+`" alt="Visible JFIF Data"><img src="`+xbmpURI+`" alt="Visible XBMP Data"><p>Raster alias data after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-raster-alias-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 3 {
		t.Fatalf("expected three raster alias data URI images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
		alt  string
	}{
		{name: "html-data-image-001.jpg", ext: ".jpg", alt: "Visible XJPEG Data"},
		{name: "html-data-image-002.jpg", ext: ".jpg", alt: "Visible JFIF Data"},
		{name: "html-data-image-003.bmp", ext: ".bmp", alt: "Visible XBMP Data"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || img.Alt != w.alt || !validImageData(w.ext, img.Data) {
			t.Fatalf("unexpected raster alias data URI image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written raster alias data URI image %s: %v", w.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Raster alias data before", "Raster alias data after", "![Visible XJPEG Data](images/html-data-image-001.jpg)", "![Visible JFIF Data](images/html-data-image-002.jpg)", "![Visible XBMP Data](images/html-data-image-003.bmp)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing raster alias data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:image", "image/x-jpeg", "image/jfif", "image/x-bmp"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept raster alias data URI metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLTIFFDataURIMimeAliasesAreExtracted(t *testing.T) {
	tifURI := "data:image/tif;base64," + base64.StdEncoding.EncodeToString(testTIFF())
	xtifURI := "data:image/x-tif;base64," + base64.StdEncoding.EncodeToString(testTIFF())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-tiff-alias.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-tiff-alias.htm", `<html><body><p>TIFF alias data before</p><img src="`+tifURI+`" alt="Visible TIF Data"><img src="`+xtifURI+`" alt="Visible XTIF Data"><p>TIFF alias data after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-tiff-alias-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two TIFF alias data URI images, got %#v", res.Images)
	}
	want := []struct {
		name string
		alt  string
	}{
		{name: "html-data-image-001.tif", alt: "Visible TIF Data"},
		{name: "html-data-image-002.tif", alt: "Visible XTIF Data"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != ".tif" || img.Alt != w.alt || !validImageData(".tif", img.Data) {
			t.Fatalf("unexpected TIFF alias data URI image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written TIFF alias data URI image %s: %v", w.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"TIFF alias data before", "TIFF alias data after", "![Visible TIF Data](images/html-data-image-001.tif)", "![Visible XTIF Data](images/html-data-image-002.tif)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing TIFF alias data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:image", "image/tif", "image/x-tif"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept TIFF alias data URI metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLLegacyDataURIMimeAliasesAreExtracted(t *testing.T) {
	tgaURI := "data:image/x-tga;base64," + base64.StdEncoding.EncodeToString(testTGA())
	pcxURI := "data:image/x-pcx;base64," + base64.StdEncoding.EncodeToString(testPCX())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-legacy.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-legacy.htm", `<html><body><p>Legacy data image before</p><img src="`+tgaURI+`" alt="Visible TGA Data"><img src="`+pcxURI+`" alt="Visible PCX Data"><p>Legacy data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-legacy-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two legacy data URI images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
		alt  string
	}{
		{name: "html-data-image-001.tga", ext: ".tga", alt: "Visible TGA Data"},
		{name: "html-data-image-002.pcx", ext: ".pcx", alt: "Visible PCX Data"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || img.Alt != w.alt || !validImageData(w.ext, img.Data) {
			t.Fatalf("unexpected legacy data URI image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written legacy data URI image %s: %v", w.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Legacy data image before", "Legacy data image after", "![Visible TGA Data](images/html-data-image-001.tga)", "![Visible PCX Data](images/html-data-image-002.pcx)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing legacy data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:image", "image/x-tga", "image/x-pcx"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept legacy data URI metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLAVIFDataURIModernMimeIsExtracted(t *testing.T) {
	dataURI := "data:image/avif;base64," + base64.StdEncoding.EncodeToString(testISOBMFF("avif"))
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-avif.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-avif.htm", `<html><body><p>AVIF data image before</p><img src="`+dataURI+`" alt="Visible AVIF Data"><p>AVIF data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-avif-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one AVIF data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.avif" || img.Ext != ".avif" || img.Alt != "Visible AVIF Data" || !validImageData(".avif", img.Data) {
		t.Fatalf("unexpected AVIF data URI image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.avif")); err != nil {
		t.Fatalf("expected written AVIF data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"AVIF data image before", "AVIF data image after", "![Visible AVIF Data](images/html-data-image-001.avif)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing AVIF data URI content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "data:image") || strings.Contains(res.Text, "data:image") {
		t.Fatalf("kept AVIF data URI in text=%q markdown=\n%s", res.Text, md)
	}
}

func TestDOCXAltChunkHTMLHEIFSequenceDataURIMimeAliasesAreExtracted(t *testing.T) {
	heifURI := "data:image/heif-sequence;base64," + base64.StdEncoding.EncodeToString(testISOBMFF("msf1"))
	heicURI := "data:image/heic-sequence;base64," + base64.StdEncoding.EncodeToString(testISOBMFF("heic"))
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-heif-sequence.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-heif-sequence.htm", `<html><body><p>HEIF sequence data before</p><img src="`+heifURI+`" alt="Visible HEIF Sequence"><img src="`+heicURI+`" alt="Visible HEIC Sequence"><p>HEIF sequence data after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-heif-sequence-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two HEIF sequence data URI images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
		alt  string
	}{
		{name: "html-data-image-001.heif", ext: ".heif", alt: "Visible HEIF Sequence"},
		{name: "html-data-image-002.heic", ext: ".heic", alt: "Visible HEIC Sequence"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || img.Alt != w.alt || !validImageData(w.ext, img.Data) {
			t.Fatalf("unexpected HEIF sequence data URI image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written HEIF sequence data URI image %s: %v", w.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"HEIF sequence data before", "HEIF sequence data after", "![Visible HEIF Sequence](images/html-data-image-001.heif)", "![Visible HEIC Sequence](images/html-data-image-002.heic)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing HEIF sequence data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:image", "image/heif-sequence", "image/heic-sequence"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept HEIF sequence data URI metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLJPEG2000DataURIMimeAliasesAreExtracted(t *testing.T) {
	jp2URI := "data:image/jp2k;base64," + base64.StdEncoding.EncodeToString(testJP2("jp2 "))
	jpxURI := "data:image/x-jpx;base64," + base64.StdEncoding.EncodeToString(testJP2("jpx "))
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-jpeg2000-alias.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-jpeg2000-alias.htm", `<html><body><p>JPEG2000 alias data before</p><img src="`+jp2URI+`" alt="Visible JP2K Data"><img src="`+jpxURI+`" alt="Visible JPX Data"><p>JPEG2000 alias data after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-jpeg2000-alias-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two JPEG2000 alias data URI images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
		alt  string
	}{
		{name: "html-data-image-001.jp2", ext: ".jp2", alt: "Visible JP2K Data"},
		{name: "html-data-image-002.jpx", ext: ".jpx", alt: "Visible JPX Data"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || img.Alt != w.alt || !validImageData(w.ext, img.Data) {
			t.Fatalf("unexpected JPEG2000 alias data URI image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written JPEG2000 alias data URI image %s: %v", w.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"JPEG2000 alias data before", "JPEG2000 alias data after", "![Visible JP2K Data](images/html-data-image-001.jp2)", "![Visible JPX Data](images/html-data-image-002.jpx)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing JPEG2000 alias data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:image", "image/jp2k", "image/x-jpx"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept JPEG2000 alias data URI metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLURLSafeBase64DataURIImageIsExtracted(t *testing.T) {
	payload := base64.RawURLEncoding.EncodeToString(testPNG())
	dataURI := "data:image/png;base64," + payload
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-data-urlsafe.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-data-urlsafe.htm", `<html><body><p>URL-safe data image before</p><img src="`+dataURI+`" alt="URL Safe Data Image"><p>URL-safe data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-data-uri-urlsafe.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "html-data-image-001.png" || res.Images[0].Alt != "URL Safe Data Image" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected URL-safe data URI image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![URL Safe Data Image](images/html-data-image-001.png)") || strings.Contains(md, "data:image") {
		t.Fatalf("markdown missing URL-safe data URI image or leaked URI:\n%s", md)
	}
}

func TestDOCXAltChunkHTMLSrcsetDataURIImageIsExtracted(t *testing.T) {
	dataURI := "data:image/png;base64," + base64.StdEncoding.EncodeToString(testPNG())
	hiddenURI := "data:image/jpeg;base64," + base64.StdEncoding.EncodeToString(testJPEG())
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-srcset-data.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-srcset-data.htm", `<html><body><p>Srcset data image before</p><img srcset="`+dataURI+` 1x" alt="Visible Srcset Data Image"><picture><source srcset="`+dataURI+` 1x"><img alt="Visible Picture Source Data Image"></picture><picture style="display:none"><source srcset="`+hiddenURI+` 1x"><img alt="Hidden Picture Source Data Image"></picture><img srcset="`+hiddenURI+` 1x" alt="Hidden Srcset Data Image" style="display:none"><p>Srcset data image after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-srcset-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two visible srcset data URI images, got %#v", res.Images)
	}
	for i, wantAlt := range []string{"Visible Srcset Data Image", "Visible Picture Source Data Image"} {
		img := res.Images[i]
		wantName := fmt.Sprintf("html-data-image-%03d.png", i+1)
		if img.Name != wantName || img.Ext != ".png" || img.Alt != wantAlt || !validImageData(".png", img.Data) {
			t.Fatalf("unexpected srcset data URI image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, wantName)); err != nil {
			t.Fatalf("expected written srcset data URI image %s: %v", wantName, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Srcset data image before", "Srcset data image after", "![Visible Srcset Data Image](images/html-data-image-001.png)", "![Visible Picture Source Data Image](images/html-data-image-002.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing srcset data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"data:image", "Hidden Srcset Data Image", "Hidden Picture Source Data Image"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden srcset data URI content %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLPercentEncodedSVGDataURIImageIsExtracted(t *testing.T) {
	safeSVG := `<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="red"/></svg>`
	unsafeSVG := `<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>`
	dataURI := "data:image/svg+xml;charset=utf-8," + url.PathEscape(safeSVG)
	unsafeURI := "data:image/svg+xml," + url.PathEscape(unsafeSVG)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-svg.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-svg.htm", `<html><body><p>Visible SVG data before</p><img src="`+dataURI+`" alt="Visible SVG Data"><img src="`+unsafeURI+`" alt="Unsafe SVG Data"><p>Visible SVG data after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-svg-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one safe SVG data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.svg" || img.Ext != ".svg" || img.Alt != "Visible SVG Data" || !validImageData(".svg", img.Data) {
		t.Fatalf("unexpected SVG data URI image extraction: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.svg")); err != nil {
		t.Fatalf("expected written SVG data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible SVG data before", "Visible SVG data after", "![Visible SVG Data](images/html-data-image-001.svg)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing SVG data URI content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Unsafe SVG Data", "script", "data:image"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unsafe SVG data URI content %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkHTMLUTF8SVGDataURIImageIsExtracted(t *testing.T) {
	safeSVG := `<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="green"/></svg>`
	dataURI := "data:image/svg+xml;utf8," + url.PathEscape(safeSVG)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-svg-utf8.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-svg-utf8.htm", `<html><body><p>UTF8 SVG data before</p><img src="`+dataURI+`" alt="Visible UTF8 SVG Data"><p>UTF8 SVG data after</p></body></html>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-svg-utf8-data-uri.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one UTF8 SVG data URI image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "html-data-image-001.svg" || img.Ext != ".svg" || img.Alt != "Visible UTF8 SVG Data" || !validImageData(".svg", img.Data) {
		t.Fatalf("unexpected UTF8 SVG data URI image extraction: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "html-data-image-001.svg")); err != nil {
		t.Fatalf("expected written UTF8 SVG data URI image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"UTF8 SVG data before", "UTF8 SVG data after", "![Visible UTF8 SVG Data](images/html-data-image-001.svg)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing UTF8 SVG data URI content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "data:image") || strings.Contains(res.Text, "data:image") {
		t.Fatalf("kept UTF8 SVG data URI in text=%q markdown=\n%s", res.Text, md)
	}
}

func TestDOCXAltChunkHTMLCharsetTextAndImageAltAreDecoded(t *testing.T) {
	htmlBytes := []byte(`<html><head><meta http-equiv="Content-Type" content="text/html; charset=gbk"></head><body><p>GBK `)
	htmlBytes = append(htmlBytes, []byte{0xd6, 0xd0, 0xce, 0xc4, 0x20, 0xce, 0xc4, 0xb1, 0xbe}...)
	htmlBytes = append(htmlBytes, []byte(`</p><img src="media/html-gbk.jpg" alt="`)...)
	htmlBytes = append(htmlBytes, []byte{0xd6, 0xd0, 0xce, 0xc4, 0x20, 0xcd, 0xbc, 0xc6, 0xac}...)
	htmlBytes = append(htmlBytes, []byte(` Target: ../media/hidden.png"><p>Tail</p></body></html>`)...)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-gbk.htm"/></Relationships>`)
	addZip(t, zw, "word/afchunk-gbk.htm", string(htmlBytes))
	addZipBytes(t, zw, "word/media/html-gbk.jpg", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-gbk.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "GBK 中文 文本") || !strings.Contains(res.Text, "Tail") {
		t.Fatalf("missing decoded GBK HTML text in %q", res.Text)
	}
	for _, bad := range []string{"ÖÐ", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept mojibake/internal HTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "html-gbk.png" || res.Images[0].Alt != "中文 图片" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded HTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "GBK 中文 文本") || !strings.Contains(md, "![中文 图片](images/html-gbk.png)") {
		t.Fatalf("markdown missing decoded GBK HTML text/image:\n%s", md)
	}
	for _, bad := range []string{"ÖÐ", "Target:", "../media/hidden.png"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept mojibake/internal HTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkHTMLXMLEncodingTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	textWord := string([]rune{0x6587, 0x672c})
	imageWord := string([]rune{0x56fe, 0x7247})
	htmlBytes := []byte(`<?xml version="1.0" encoding="gbk"?><html><body><p>XHTML `)
	htmlBytes = append(htmlBytes, []byte{0xd6, 0xd0, 0xce, 0xc4, 0x20, 0xce, 0xc4, 0xb1, 0xbe}...)
	htmlBytes = append(htmlBytes, []byte(`</p><img src="media/xhtml-gbk.jpg" alt="`)...)
	htmlBytes = append(htmlBytes, []byte{0xd6, 0xd0, 0xce, 0xc4, 0x20, 0xcd, 0xbc, 0xc6, 0xac}...)
	htmlBytes = append(htmlBytes, []byte(` Target: ../media/hidden.png"><p>Tail</p></body></html>`)...)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-xhtml-gbk.htm"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-xhtml-gbk.htm", htmlBytes)
	addZipBytes(t, zw, "word/media/xhtml-gbk.jpg", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-xhtml-gbk.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	wantText := "XHTML " + zh + " " + textWord
	wantAlt := zh + " " + imageWord
	if !strings.Contains(res.Text, wantText) || !strings.Contains(res.Text, "Tail") {
		t.Fatalf("missing XML-declared GBK HTML text %q in %q", wantText, res.Text)
	}
	for _, bad := range []string{"脰脨", "涓枃", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept mojibake/internal XML-declared HTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "xhtml-gbk.png" || res.Images[0].Alt != wantAlt || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected XML-declared decoded HTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, wantText) || !strings.Contains(md, "!["+wantAlt+"](images/xhtml-gbk.png)") {
		t.Fatalf("markdown missing XML-declared decoded GBK HTML text/image:\n%s", md)
	}
	for _, bad := range []string{"脰脨", "涓枃", "Target:", "../media/hidden.png"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept mojibake/internal XML-declared HTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkHTMLUTF16BOMTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	htmlText := `<html><body><p>UTF16 ` + zh + ` text</p><img src="media/html-utf16.jpg" alt="UTF16 ` + zh + ` image Target: ../media/hidden.png"><p>Tail</p></body></html>`
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-utf16.htm"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-utf16.htm", utf16LEBOMBytes(htmlText))
	addZipBytes(t, zw, "word/media/html-utf16.jpg", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-utf16.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "UTF16 "+zh+" text") || !strings.Contains(res.Text, "Tail") {
		t.Fatalf("missing decoded UTF-16 HTML text in %q", res.Text)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept UTF-16/control/internal HTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "html-utf16.png" || res.Images[0].Alt != "UTF16 "+zh+" image" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded UTF-16 HTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "UTF16 "+zh+" text") || !strings.Contains(md, "![UTF16 "+zh+" image](images/html-utf16.png)") {
		t.Fatalf("markdown missing decoded UTF-16 HTML text/image:\n%s", md)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept UTF-16/control/internal HTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkHTMLUTF16NoBOMTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	htmlText := `<html><body><p>No BOM ` + zh + ` text</p><img src="media/html-utf16-nobom.jpg" alt="No BOM ` + zh + ` image Target: ../media/hidden.png"></body></html>`
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdHtml"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-utf16-nobom.htm"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-utf16-nobom.htm", utf16LEBytes(htmlText))
	addZipBytes(t, zw, "word/media/html-utf16-nobom.jpg", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-html-utf16-nobom.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "No BOM "+zh+" text") {
		t.Fatalf("missing decoded no-BOM UTF-16 HTML text in %q", res.Text)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept no-BOM UTF-16/control/internal HTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "html-utf16-nobom.png" || res.Images[0].Alt != "No BOM "+zh+" image" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded no-BOM UTF-16 HTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "No BOM "+zh+" text") || !strings.Contains(md, "![No BOM "+zh+" image](images/html-utf16-nobom.png)") {
		t.Fatalf("markdown missing decoded no-BOM UTF-16 HTML text/image:\n%s", md)
	}
}

func TestDOCXAltChunkMHTMLTextAndImagesAreVisible(t *testing.T) {
	visiblePNG := testPNG()
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_000_0001"`,
		``,
		`------=_NextPart_000_0001`,
		`Content-Type: text/html; charset="utf-8"`,
		`Content-Location: file:///C:/Users/me/hidden-source.htm`,
		``,
		`<html><body><p>Visible MHTML before image</p><img src="word/media/mhtml-visible.jpg?download=1#frag" alt="Visible MHTML Diagram Target: ../media/hidden.png"><img srcset="word/media/mhtml-srcset-visible.png 1x" alt="Visible MHTML Srcset Diagram"><picture><source srcset="word/media/mhtml-picture-source-visible.png 1x"><img alt="Visible MHTML Picture Source Diagram"></picture><picture style="display:none"><source srcset="word/media/mhtml-picture-source-hidden.png 1x"><img alt="Hidden MHTML Picture Source Image"></picture><template><img src="word/media/mhtml-template-hidden.png" alt="Hidden MHTML Template Image"></template><img srcset="word/media/mhtml-srcset-hidden.png 1x" alt="Hidden MHTML Srcset Image" style="display:none"><img src="word/media/mhtml-hidden-style.png" alt="Hidden MHTML Style Image" style="display:none"><div hidden><img src="word/media/mhtml-hidden-parent.png" alt="Hidden MHTML Parent Image"></div><!-- <img src="word/media/mhtml-hidden-comment.png" alt="Hidden MHTML Comment Image"> --><p>Visible MHTML after image</p><script>Hidden MHTML Script</script></body></html>`,
		`------=_NextPart_000_0001`,
		`Content-Type: image/jpeg`,
		`Content-Transfer-Encoding: "base64"; x=1`,
		`Content-Location: word/media/mhtml-visible.jpg?download=1#frag`,
		``,
		base64.StdEncoding.EncodeToString(visiblePNG),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-srcset-visible.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-picture-source-visible.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-picture-source-hidden.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-template-hidden.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-srcset-hidden.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-hidden-style.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-hidden-parent.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-hidden-comment.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/mhtml-unreferenced.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0001--`,
		``,
	}, "\r\n")
	orphan := strings.ReplaceAll(mhtml, "Visible MHTML", "Orphan MHTML")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk.mht", mhtml)
	addZip(t, zw, "word/orphan.mht", orphan)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "referenced-altchunk-mhtml.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Body Text", "Visible MHTML before image", "Visible MHTML after image"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible MHTML altChunk text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"MIME-Version", "Content-Type", "Content-Location", "file:///C:/Users/me/hidden-source.htm", "Hidden MHTML Script", "Orphan MHTML"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden MHTML content %q in %q", hidden, res.Text)
		}
	}
	if len(res.Images) != 3 {
		t.Fatalf("expected visible MHTML images, got %#v", res.Images)
	}
	if res.Images[0].Name != "mhtml-visible.png" || res.Images[0].Ext != ".png" || res.Images[0].Alt != "Visible MHTML Diagram" || !bytes.Equal(res.Images[0].Data, visiblePNG) {
		t.Fatalf("unexpected MHTML image extraction: %#v", res.Images[0])
	}
	if res.Images[1].Name != "mhtml-srcset-visible.png" || res.Images[1].Ext != ".png" || res.Images[1].Alt != "Visible MHTML Srcset Diagram" || !validImageData(".png", res.Images[1].Data) {
		t.Fatalf("unexpected MHTML srcset image extraction: %#v", res.Images[1])
	}
	if res.Images[2].Name != "mhtml-picture-source-visible.png" || res.Images[2].Ext != ".png" || res.Images[2].Alt != "Visible MHTML Picture Source Diagram" || !validImageData(".png", res.Images[2].Data) {
		t.Fatalf("unexpected MHTML picture source image extraction: %#v", res.Images[2])
	}
	if _, err := os.Stat(filepath.Join(outDir, "mhtml-visible.png")); err != nil {
		t.Fatalf("expected written MHTML image: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "mhtml-srcset-visible.png")); err != nil {
		t.Fatalf("expected written MHTML srcset image: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "mhtml-picture-source-visible.png")); err != nil {
		t.Fatalf("expected written MHTML picture source image: %v", err)
	}
	for _, hiddenName := range []string{"mhtml-picture-source-hidden.png", "mhtml-template-hidden.png", "mhtml-srcset-hidden.png", "mhtml-hidden-style.png", "mhtml-hidden-parent.png", "mhtml-hidden-comment.png", "mhtml-unreferenced.png"} {
		if _, err := os.Stat(filepath.Join(outDir, hiddenName)); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("hidden/unreferenced MHTML image %s should not be written, stat err=%v", hiddenName, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## HTML Content", "Visible MHTML before image", "Visible MHTML after image", "![Visible MHTML Diagram](images/mhtml-visible.png)", "![Visible MHTML Srcset Diagram](images/mhtml-srcset-visible.png)", "![Visible MHTML Picture Source Diagram](images/mhtml-picture-source-visible.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible MHTML content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/hidden.png", "download=1", "#frag", "Content-Location", "Orphan MHTML", "Hidden MHTML Picture Source Image", "Hidden MHTML Template Image", "Hidden MHTML Srcset Image", "Hidden MHTML Style Image", "Hidden MHTML Parent Image", "Hidden MHTML Comment Image", "mhtml-picture-source-hidden.png", "mhtml-template-hidden.png", "mhtml-srcset-hidden.png", "mhtml-hidden-style.png", "mhtml-hidden-parent.png", "mhtml-hidden-comment.png", "mhtml-unreferenced.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden MHTML content %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXAltChunkMHTMLImageMimeAliasIsExtracted(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_000_0002"`,
		``,
		`------=_NextPart_000_0002`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>MHTML alias before</p><img src="word/media/alias-image.jpg" alt="Alias MIME Image"><p>MHTML alias after</p></body></html>`,
		`------=_NextPart_000_0002`,
		`Content-Type: image/x-png; name="alias-image.jpg"`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: "word/media/alias-image.jpg"`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_000_0002--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-alias.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-alias.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-image-mime-alias.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one MHTML MIME alias image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "alias-image.png" || img.Ext != ".png" || img.Alt != "Alias MIME Image" || !validImageData(".png", img.Data) {
		t.Fatalf("unexpected MHTML MIME alias image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "alias-image.png")); err != nil {
		t.Fatalf("expected written MHTML MIME alias image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"MHTML alias before", "MHTML alias after", "![Alias MIME Image](images/alias-image.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing MHTML MIME alias content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Type", "image/x-png", `"word/media/alias-image.jpg"`, "alias-image.jpg)"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept MHTML MIME alias metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLRawBase64ImageIsExtracted(t *testing.T) {
	rawBase64 := strings.TrimRight(base64.StdEncoding.EncodeToString(testPNG()), "=")
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="raw-base64-boundary"`,
		``,
		`--raw-base64-boundary`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>MHTML raw payload before</p><img src="word/media/raw-base64.png" alt="Raw Payload Image Content-Transfer-Encoding: base64"><p>MHTML raw payload after</p></body></html>`,
		`--raw-base64-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/raw-base64.png`,
		``,
		rawBase64,
		`--raw-base64-boundary--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-raw-base64.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-raw-base64.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-raw-base64.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "raw-base64.png" || res.Images[0].Alt != "Raw Payload Image" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected raw-base64 MHTML image: %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "raw-base64.png")); err != nil {
		t.Fatalf("expected written raw-base64 MHTML image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"MHTML raw payload before", "MHTML raw payload after", "![Raw Payload Image](images/raw-base64.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing raw-base64 MHTML content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Transfer-Encoding", rawBase64[:16]} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) || strings.Contains(res.Images[0].Alt, hidden) {
			t.Fatalf("kept raw-base64 MHTML metadata %q in text=%q alt=%q markdown=\n%s", hidden, res.Text, res.Images[0].Alt, md)
		}
	}
}

func TestDOCXAltChunkMHTMLAngleBracketContentLocationImageIsExtracted(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="angle-location-boundary"`,
		``,
		`--angle-location-boundary`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>MHTML angle before</p><img src="word/media/angle-location.png" alt="Angle Location Image Target: ../media/hidden.png"><p>MHTML angle after</p></body></html>`,
		`--angle-location-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: <word/media/angle-location.png>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`--angle-location-boundary--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-angle-location.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-angle-location.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-angle-location.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "angle-location.png" || res.Images[0].Alt != "Angle Location Image" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected angle-bracket MHTML image: %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "angle-location.png")); err != nil {
		t.Fatalf("expected written angle-bracket MHTML image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"MHTML angle before", "MHTML angle after", "![Angle Location Image](images/angle-location.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing angle-bracket MHTML content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Location", "<word/media/angle-location.png>", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("kept angle-bracket MHTML metadata/internal text %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLSniffsGenericImageMime(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_GENERIC_IMAGE_MIME"`,
		``,
		`------=_NextPart_GENERIC_IMAGE_MIME`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>MHTML generic before</p><img src="word/media/generic-image.bin" alt="Generic MIME Image"><img src="word/media/invalid-image.bin" alt="Invalid MIME Image"><p>MHTML generic after</p></body></html>`,
		`------=_NextPart_GENERIC_IMAGE_MIME`,
		`Content-Type: application/octet-stream; name="generic-image.bin"`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/generic-image.bin`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_GENERIC_IMAGE_MIME`,
		`Content-Type: application/octet-stream; name="invalid-image.bin"`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/invalid-image.bin`,
		``,
		base64.StdEncoding.EncodeToString([]byte("not an image")),
		`------=_NextPart_GENERIC_IMAGE_MIME--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-generic-image-mime.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-generic-image-mime.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-generic-image-mime.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one sniffed MHTML generic MIME image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "generic-image.png" || img.Ext != ".png" || img.Alt != "Generic MIME Image" || !validImageData(".png", img.Data) {
		t.Fatalf("unexpected MHTML generic MIME image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "generic-image.png")); err != nil {
		t.Fatalf("expected written MHTML generic MIME image: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "invalid-image.bin")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("invalid generic MIME image should not be written, stat err=%v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"MHTML generic before", "MHTML generic after", "![Generic MIME Image](images/generic-image.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing MHTML generic MIME content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Type", "application/octet-stream", "generic-image.bin)", "Invalid MIME Image", "invalid-image.bin"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept MHTML generic MIME metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLModernImageMimeIsExtracted(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_000_0003"`,
		``,
		`------=_NextPart_000_0003`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>MHTML JP2 before</p><img src="word/media/modern-image.bin" alt="Visible JP2 Image"><p>MHTML JP2 after</p></body></html>`,
		`------=_NextPart_000_0003`,
		`Content-Type: image/jp2; name="modern-image.bin"`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/modern-image.bin`,
		``,
		base64.StdEncoding.EncodeToString(testJP2("jp2 ")),
		`------=_NextPart_000_0003--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-modern.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-modern.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-modern-image-mime.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected one MHTML modern MIME image, got %#v", res.Images)
	}
	img := res.Images[0]
	if img.Name != "modern-image.jp2" || img.Ext != ".jp2" || img.Alt != "Visible JP2 Image" || !validImageData(".jp2", img.Data) {
		t.Fatalf("unexpected MHTML modern MIME image: %#v", img)
	}
	if _, err := os.Stat(filepath.Join(outDir, "modern-image.jp2")); err != nil {
		t.Fatalf("expected written MHTML modern MIME image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"MHTML JP2 before", "MHTML JP2 after", "![Visible JP2 Image](images/modern-image.jp2)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing MHTML modern MIME content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Type", "image/jp2", "modern-image.bin)"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept MHTML modern MIME metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLLegacyImageMimeAliasesAreExtracted(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_LEGACY_IMAGE_MIME"`,
		``,
		`------=_NextPart_LEGACY_IMAGE_MIME`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>MHTML legacy before</p><img src="word/media/legacy-tga.bin" alt="Visible MHTML TGA"><img src="word/media/legacy-pcx.bin" alt="Visible MHTML PCX"><p>MHTML legacy after</p></body></html>`,
		`------=_NextPart_LEGACY_IMAGE_MIME`,
		`Content-Type: image/x-tga; name="legacy-tga.bin"`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/legacy-tga.bin`,
		``,
		base64.StdEncoding.EncodeToString(testTGA()),
		`------=_NextPart_LEGACY_IMAGE_MIME`,
		`Content-Type: image/x-pcx; name="legacy-pcx.bin"`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/legacy-pcx.bin`,
		``,
		base64.StdEncoding.EncodeToString(testPCX()),
		`------=_NextPart_LEGACY_IMAGE_MIME--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-legacy-image-mime.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-legacy-image-mime.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-legacy-image-mime.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two MHTML legacy MIME images, got %#v", res.Images)
	}
	want := []struct {
		name string
		ext  string
		alt  string
	}{
		{name: "legacy-tga.tga", ext: ".tga", alt: "Visible MHTML TGA"},
		{name: "legacy-pcx.pcx", ext: ".pcx", alt: "Visible MHTML PCX"},
	}
	for i, w := range want {
		img := res.Images[i]
		if img.Name != w.name || img.Ext != w.ext || img.Alt != w.alt || !validImageData(w.ext, img.Data) {
			t.Fatalf("unexpected MHTML legacy MIME image %d: %#v", i, img)
		}
		if _, err := os.Stat(filepath.Join(outDir, w.name)); err != nil {
			t.Fatalf("expected written MHTML legacy MIME image %s: %v", w.name, err)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"MHTML legacy before", "MHTML legacy after", "![Visible MHTML TGA](images/legacy-tga.tga)", "![Visible MHTML PCX](images/legacy-pcx.pcx)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing MHTML legacy MIME content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Type", "image/x-tga", "image/x-pcx", "legacy-tga.bin)", "legacy-pcx.bin)"} {
		if strings.Contains(md, hidden) || strings.Contains(res.Text, hidden) {
			t.Fatalf("kept MHTML legacy MIME metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLCharsetAndCIDImageAltAreDecoded(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_GBK"`,
		``,
		`------=_NextPart_GBK`,
		`Content-Type: text/html; charset=gbk`,
		`Content-Transfer-Encoding: quoted-printable`,
		``,
		`<html><body><p>GBK =D6=D0=CE=C4 =CE=C4=B1=BE</p><img src=3D"CID:IMAGE001.PNG@OFFICE" alt=3D"=D6=D0=CE=C4 =CD=BC=C6=AC Target: ../media/hidden.png"></body></html>`,
		`------=_NextPart_GBK`,
		`Content-Type: image/jpeg`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <image001.png@office>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_GBK--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-gbk.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-gbk.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-gbk.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "GBK 中文 文本") {
		t.Fatalf("missing decoded GBK MHTML text in %q", res.Text)
	}
	for _, bad := range []string{"=D6=D0", "ÖÐ", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept mojibake/internal MHTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "image001.png" || res.Images[0].Alt != "中文 图片" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded MHTML CID image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "GBK 中文 文本") || !strings.Contains(md, "![中文 图片](images/image001.png)") {
		t.Fatalf("markdown missing decoded MHTML GBK text/image:\n%s", md)
	}
	for _, bad := range []string{"=D6=D0", "ÖÐ", "Target:", "../media/hidden.png"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept mojibake/internal MHTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkMHTMLSingleQuotedHeaderParamsAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	textWord := string([]rune{0x6587, 0x672c})
	imageWord := string([]rune{0x56fe, 0x7247})
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary='single-quoted-boundary'`,
		``,
		`--single-quoted-boundary`,
		`Content-Type: text/html; charset='gbk'`,
		`Content-Transfer-Encoding: quoted-printable`,
		``,
		`<html><body><p>Single Quote =D6=D0=CE=C4 =CE=C4=B1=BE</p><img src=3D"cid:single-quoted@office" alt=3D"=D6=D0=CE=C4 =CD=BC=C6=AC Target: ../media/hidden.png"></body></html>`,
		`--single-quoted-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <single-quoted@office>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`--single-quoted-boundary--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-single-quoted.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-single-quoted.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-single-quoted.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	wantText := "Single Quote " + zh + " " + textWord
	wantAlt := zh + " " + imageWord
	if !strings.Contains(res.Text, wantText) {
		t.Fatalf("missing single-quoted MHTML charset text %q in %q", wantText, res.Text)
	}
	for _, bad := range []string{"=D6=D0", "脰脨", "Target:", "../media/hidden.png", "boundary='", "charset='gbk'"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept single-quoted MHTML mojibake/internal text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "single-quoted.png" || res.Images[0].Alt != wantAlt || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected single-quoted MHTML CID image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, wantText) || !strings.Contains(md, "!["+wantAlt+"](images/single-quoted.png)") {
		t.Fatalf("markdown missing single-quoted MHTML text/image:\n%s", md)
	}
	for _, bad := range []string{"=D6=D0", "脰脨", "Target:", "../media/hidden.png", "boundary='", "charset='gbk'"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept single-quoted MHTML mojibake/internal text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkMHTMLQuotedContentIDIsMatchedAndCleaned(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="quoted-cid-boundary"`,
		``,
		`--quoted-cid-boundary`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>Visible quoted CID before</p><img src="cid:quoted-cid@office" alt="Visible Quoted CID Image Content-ID: &lt;hidden@office&gt; Target: ../media/hidden.png"><p>Visible quoted CID after</p></body></html>`,
		`--quoted-cid-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: '<quoted-cid@office>'`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`--quoted-cid-boundary--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-quoted-cid.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-quoted-cid.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-quoted-cid.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "quoted-cid.png" || res.Images[0].Alt != "Visible Quoted CID Image" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected quoted CID MHTML image: %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "quoted-cid.png")); err != nil {
		t.Fatalf("missing written quoted CID image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible quoted CID before", "Visible quoted CID after", "![Visible Quoted CID Image](images/quoted-cid.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing quoted CID content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-ID", "<hidden@office>", "'<quoted-cid@office>'", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("kept quoted CID internal metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLFoldedContentLocationIsMatched(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="folded-location-boundary"`,
		``,
		`--folded-location-boundary`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>Visible folded location before</p><img src="word/media/folded-location.png" alt="Visible Folded Location Image Content-Location: word/media/hidden.png"><p>Visible folded location after</p></body></html>`,
		`--folded-location-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-Location: word/media/`,
		` folded-location.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`--folded-location-boundary--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-folded-location.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-folded-location.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-folded-location.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "folded-location.png" || res.Images[0].Alt != "Visible Folded Location Image" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected folded MHTML location image: %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "folded-location.png")); err != nil {
		t.Fatalf("missing written folded location image: %v", err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible folded location before", "Visible folded location after", "![Visible Folded Location Image](images/folded-location.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing folded location content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Location", "word/media/ folded-location.png", "word/media/hidden.png"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("kept folded location internal metadata %q in text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestDOCXAltChunkMHTMLQuotedBoundaryWithSemicolonIsParsed(t *testing.T) {
	mhtml := strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="semi;colon-boundary"; type="text/html"`,
		``,
		`--semi;colon-boundary`,
		`Content-Type: text/html; charset="utf-8"`,
		``,
		`<html><body><p>Visible semicolon boundary text</p><img src="cid:semicolon-boundary@office" alt="Visible Semicolon Boundary Image Target: ../media/hidden.png"></body></html>`,
		`--semi;colon-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <semicolon-boundary@office>`,
		`Content-Location: word/media/semicolon-boundary.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`--semi;colon-boundary--`,
		``,
	}, "\r\n")
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-semicolon-boundary.mht"/></Relationships>`)
	addZip(t, zw, "word/afchunk-semicolon-boundary.mht", mhtml)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-semicolon-boundary.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible semicolon boundary text") {
		t.Fatalf("missing semicolon-boundary MHTML text in %q", res.Text)
	}
	for _, hidden := range []string{"boundary=", "type=\"text/html\"", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept semicolon-boundary MHTML metadata/internal text %q in %q", hidden, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "semicolon-boundary.png" || res.Images[0].Alt != "Visible Semicolon Boundary Image" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected semicolon-boundary MHTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible semicolon boundary text") || !strings.Contains(md, "![Visible Semicolon Boundary Image](images/semicolon-boundary.png)") {
		t.Fatalf("markdown missing semicolon-boundary MHTML text/image:\n%s", md)
	}
	for _, hidden := range []string{"boundary=", "type=\"text/html\"", "Target:", "../media/hidden.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept semicolon-boundary MHTML metadata/internal text %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXAltChunkMHTMLHTMLMetaCharsetAndCIDImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	textWord := string([]rune{0x6587, 0x672c})
	imageWord := string([]rune{0x56fe, 0x7247})
	var htmlPart []byte
	htmlPart = append(htmlPart, []byte(`<html><head><meta http-equiv="Content-Type" content="text/html; charset=gbk"></head><body><p>MHTML Meta `)...)
	htmlPart = append(htmlPart, []byte{0xd6, 0xd0, 0xce, 0xc4, 0x20, 0xce, 0xc4, 0xb1, 0xbe}...)
	htmlPart = append(htmlPart, []byte(`</p><img src="cid:meta-gbk@office" alt="`)...)
	htmlPart = append(htmlPart, []byte{0xd6, 0xd0, 0xce, 0xc4, 0x20, 0xcd, 0xbc, 0xc6, 0xac}...)
	htmlPart = append(htmlPart, []byte(` Target: ../media/hidden.png"></body></html>`)...)
	var mhtml bytes.Buffer
	mhtml.WriteString(strings.Join([]string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="meta-boundary"`,
		``,
		`--meta-boundary`,
		`Content-Type: text/html`,
		`Content-Location: file:///C:/Users/me/hidden-source.htm`,
		``,
	}, "\r\n"))
	mhtml.WriteString("\r\n")
	mhtml.Write(htmlPart)
	mhtml.WriteString(strings.Join([]string{
		``,
		`--meta-boundary`,
		`Content-Type: image/png`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <meta-gbk@office>`,
		`Content-Location: word/media/meta-gbk.png`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`--meta-boundary--`,
		``,
	}, "\r\n"))
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-meta-gbk.mht"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-meta-gbk.mht", mhtml.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-meta-gbk.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	wantText := "MHTML Meta " + zh + " " + textWord
	wantAlt := zh + " " + imageWord
	if !strings.Contains(res.Text, wantText) {
		t.Fatalf("missing HTML-meta decoded GBK MHTML text %q in %q", wantText, res.Text)
	}
	for _, bad := range []string{"脰脨", "涓枃", "Target:", "../media/hidden.png", "Content-Location"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept mojibake/internal HTML-meta MHTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "meta-gbk.png" || res.Images[0].Alt != wantAlt || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("unexpected HTML-meta decoded MHTML CID image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, wantText) || !strings.Contains(md, "!["+wantAlt+"](images/meta-gbk.png)") {
		t.Fatalf("markdown missing HTML-meta decoded MHTML GBK text/image:\n%s", md)
	}
	for _, bad := range []string{"脰脨", "涓枃", "Target:", "../media/hidden.png", "Content-Location"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept mojibake/internal HTML-meta MHTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkMHTMLUTF16BOMTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	htmlText := `<html><body><p>MHTML UTF16 ` + zh + ` text</p><img src="cid:utf16-image@office" alt="MHTML UTF16 ` + zh + ` image Target: ../media/hidden.png"></body></html>`
	var mhtml bytes.Buffer
	for _, line := range []string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_UTF16"`,
		``,
		`------=_NextPart_UTF16`,
		`Content-Type: text/html`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	mhtml.Write(utf16BEBOMBytes(htmlText))
	mhtml.WriteString("\r\n")
	for _, line := range []string{
		`------=_NextPart_UTF16`,
		`Content-Type: image/jpeg`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <utf16-image@office>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_UTF16--`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-utf16.mht"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-utf16.mht", mhtml.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-utf16.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "MHTML UTF16 "+zh+" text") {
		t.Fatalf("missing decoded UTF-16 MHTML text in %q", res.Text)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept UTF-16/control/internal MHTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "utf16-image.png" || res.Images[0].Alt != "MHTML UTF16 "+zh+" image" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded UTF-16 MHTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "MHTML UTF16 "+zh+" text") || !strings.Contains(md, "![MHTML UTF16 "+zh+" image](images/utf16-image.png)") {
		t.Fatalf("markdown missing decoded UTF-16 MHTML text/image:\n%s", md)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept UTF-16/control/internal MHTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkMHTMLUTF16NoBOMTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	htmlText := `<html><body><p>MHTML No BOM ` + zh + ` text</p><img src="cid:utf16-nobom@office" alt="MHTML No BOM ` + zh + ` image Target: ../media/hidden.png"></body></html>`
	var mhtml bytes.Buffer
	for _, line := range []string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_UTF16_NOBOM"`,
		``,
		`------=_NextPart_UTF16_NOBOM`,
		`Content-Type: text/html`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	mhtml.Write(utf16BEBytes(htmlText))
	mhtml.WriteString("\r\n")
	for _, line := range []string{
		`------=_NextPart_UTF16_NOBOM`,
		`Content-Type: image/jpeg`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <utf16-nobom@office>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_UTF16_NOBOM--`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-utf16-nobom.mht"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-utf16-nobom.mht", mhtml.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-utf16-nobom.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "MHTML No BOM "+zh+" text") {
		t.Fatalf("missing decoded no-BOM UTF-16 MHTML text in %q", res.Text)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept no-BOM UTF-16/control/internal MHTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "utf16-nobom.png" || res.Images[0].Alt != "MHTML No BOM "+zh+" image" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded no-BOM UTF-16 MHTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "MHTML No BOM "+zh+" text") || !strings.Contains(md, "![MHTML No BOM "+zh+" image](images/utf16-nobom.png)") {
		t.Fatalf("markdown missing decoded no-BOM UTF-16 MHTML text/image:\n%s", md)
	}
}

func TestDOCXAltChunkMHTMLUTF16BECharsetTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	htmlText := `<html><body><p>MHTML UTF16BE charset ` + zh + ` text</p><img src="cid:utf16be-charset@office" alt="MHTML UTF16BE charset ` + zh + ` image Target: ../media/hidden.png"></body></html>`
	var mhtml bytes.Buffer
	for _, line := range []string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_UTF16BE_CHARSET"`,
		``,
		`------=_NextPart_UTF16BE_CHARSET`,
		`Content-Type: text/html; charset=utf-16be`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	mhtml.Write(utf16BEBytes(htmlText))
	mhtml.WriteString("\r\n")
	for _, line := range []string{
		`------=_NextPart_UTF16BE_CHARSET`,
		`Content-Type: image/jpeg`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <utf16be-charset@office>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_UTF16BE_CHARSET--`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-utf16be-charset.mht"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-utf16be-charset.mht", mhtml.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-utf16be-charset.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "MHTML UTF16BE charset "+zh+" text") {
		t.Fatalf("missing decoded UTF-16BE charset MHTML text in %q", res.Text)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png", "utf-16be"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept UTF-16BE charset/control/internal MHTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "utf16be-charset.png" || res.Images[0].Alt != "MHTML UTF16BE charset "+zh+" image" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded UTF-16BE charset MHTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "MHTML UTF16BE charset "+zh+" text") || !strings.Contains(md, "![MHTML UTF16BE charset "+zh+" image](images/utf16be-charset.png)") {
		t.Fatalf("markdown missing decoded UTF-16BE charset MHTML text/image:\n%s", md)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png", "utf-16be"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept UTF-16BE charset/control/internal MHTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXAltChunkMHTMLUTF16BECharsetAliasTextAndImageAltAreDecoded(t *testing.T) {
	zh := string([]rune{0x4e2d, 0x6587})
	htmlText := `<html><body><p>MHTML UTF16BE alias ` + zh + ` text</p><img src="cid:utf16be-alias@office" alt="MHTML UTF16BE alias ` + zh + ` image Target: ../media/hidden.png"></body></html>`
	var mhtml bytes.Buffer
	for _, line := range []string{
		`MIME-Version: 1.0`,
		`Content-Type: multipart/related; boundary="----=_NextPart_UTF16BE_ALIAS"`,
		``,
		`------=_NextPart_UTF16BE_ALIAS`,
		`Content-Type: text/html; charset=utf16be`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	mhtml.Write(utf16BEBytes(htmlText))
	mhtml.WriteString("\r\n")
	for _, line := range []string{
		`------=_NextPart_UTF16BE_ALIAS`,
		`Content-Type: image/jpeg`,
		`Content-Transfer-Encoding: base64`,
		`Content-ID: <utf16be-alias@office>`,
		``,
		base64.StdEncoding.EncodeToString(testPNG()),
		`------=_NextPart_UTF16BE_ALIAS--`,
		``,
	} {
		mhtml.WriteString(line + "\r\n")
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rIdMHTML"/></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMHTML" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk-utf16be-alias.mht"/></Relationships>`)
	addZipBytes(t, zw, "word/afchunk-utf16be-alias.mht", mhtml.Bytes())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "altchunk-mhtml-utf16be-alias.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "MHTML UTF16BE alias "+zh+" text") {
		t.Fatalf("missing decoded UTF-16BE alias MHTML text in %q", res.Text)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png", "charset=utf16be"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept UTF-16BE alias/control/internal MHTML text %q in %q", bad, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "utf16be-alias.png" || res.Images[0].Alt != "MHTML UTF16BE alias "+zh+" image" || res.Images[0].Ext != ".png" {
		t.Fatalf("unexpected decoded UTF-16BE alias MHTML image: %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "MHTML UTF16BE alias "+zh+" text") || !strings.Contains(md, "![MHTML UTF16BE alias "+zh+" image](images/utf16be-alias.png)") {
		t.Fatalf("markdown missing decoded UTF-16BE alias MHTML text/image:\n%s", md)
	}
	for _, bad := range []string{"\x00", "Target:", "../media/hidden.png", "charset=utf16be"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept UTF-16BE alias/control/internal MHTML text %q in:\n%s", bad, md)
		}
	}
}

func TestDOCXGlossaryIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Visible Body Text</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "word/glossary/document.xml", `<w:glossaryDocument xmlns:w="urn:x" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:docParts><w:docPart><w:docPartBody><w:p><w:r><w:t>Hidden Glossary Building Block Secret</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts><p:pic><p:nvPicPr><p:cNvPr id="1" descr="Hidden Glossary Image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdGlossaryImage"/></p:blipFill></p:pic></w:glossaryDocument>`)
	addZip(t, zw, "word/glossary/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdGlossaryImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/glossary-hidden.png"/></Relationships>`)
	addZip(t, zw, "word/glossary/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:title><a:p><a:r><a:t>Hidden Glossary Chart Secret</a:t></a:r></a:p></c:title></c:chartSpace>`)
	addZipBytes(t, zw, "word/media/glossary-hidden.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "glossary.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Body Text") {
		t.Fatalf("missing visible DOCX body in %q", res.Text)
	}
	for _, hidden := range []string{"Hidden Glossary Building Block Secret", "Hidden Glossary Chart Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden DOCX glossary content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible Body Text") {
		t.Fatalf("markdown missing visible DOCX body in:\n%s", md)
	}
	for _, hidden := range []string{"## Glossary", "Hidden Glossary Building Block Secret", "Hidden Glossary Chart Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden DOCX glossary content %q in:\n%s", hidden, md)
		}
	}
	if len(res.Images) != 0 {
		t.Fatalf("expected no DOCX glossary images, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "glossary-hidden.png")); !os.IsNotExist(err) {
		t.Fatalf("glossary image was written or stat failed unexpectedly: %v", err)
	}
}

func TestDOCXMarkdownConvertsTables(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body>
<w:p><w:r><w:t>Before Table</w:t></w:r></w:p>
<w:tbl>
<w:tr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Alice</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Ready</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Carol</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Visible Target: ../media/table.png table text</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Bob</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden Cell Secret</w:t></w:r><w:r><w:t>Done</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>After Table</w:t></w:r></w:p>
</w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "docx-table.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"Before Table",
		"| Name | Status |",
		"| --- | --- |",
		"| Alice | Ready |",
		"| Carol | Visible table text |",
		"| Bob | Done |",
		"After Table",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing DOCX table content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Cell Secret", "Target:", "../media/table.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden DOCX table text %q in:\n%s", hidden, md)
		}
	}
}

func TestDOCXMarkdownIncludesVisibleFormFieldResults(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body>
<w:p><w:r><w:t>Done </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:checked/></w:checkBox></w:ffData></w:fldChar><w:t> Status </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:result w:val="1"/><w:listEntry w:val="Draft"/><w:listEntry w:val="Approved Target: ../media/status.png"/></w:ddList></w:ffData></w:fldChar><w:t> Owner </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:default w:val="Alice Example rId33"/></w:textInput></w:ffData></w:fldChar></w:r></w:p>
<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden form </w:t><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:checked/></w:checkBox><w:ddList><w:result w:val="1"/><w:listEntry w:val="Visible"/><w:listEntry w:val="Hidden Choice"/></w:ddList><w:textInput><w:default w:val="Hidden Default"/></w:textInput></w:ffData></w:fldChar></w:r></w:p>
<w:tbl>
<w:tr><w:tc><w:p><w:r><w:t>Task</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Flag</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Choice</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox/></w:ffData></w:fldChar></w:r></w:p></w:tc><w:tc><w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList><w:listEntry w:val="Low ContentType: image/png"/><w:listEntry w:val="High"/></w:ddList></w:ffData></w:fldChar></w:r></w:p></w:tc><w:tc><w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:default w:val="Cell Note Target=&quot;media/cell.png&quot;"/></w:textInput></w:ffData></w:fldChar></w:r></w:p></w:tc></w:tr>
</w:tbl>
</w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "form-markdown.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"Done \u2612 Status Approved Owner Alice Example",
		"| Task | Flag | Choice | Note |",
		"| Alpha | \u2610 | Low | Cell Note |",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible form field result %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Draft", "Target:", "../media/status.png", "rId33", "ContentType:", "image/png", "media/cell.png", "Hidden form", "Hidden Choice", "Hidden Default"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden or unselected form field text %q in:\n%s", hidden, md)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("visible form field results should be structured, not backfilled as additional text:\n%s", md)
	}
}

func TestDOCXMarkdownTruncatesLargeTableCells(t *testing.T) {
	large := "Visible large DOCX table cell " + strings.Repeat("A", maxMarkdownTableCellBytes*4)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>`+large+`</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "large-docx-table-cell.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible large DOCX table cell") {
		t.Fatalf("markdown missing large DOCX table cell prefix in:\n%s", md)
	}
	if strings.Contains(md, strings.Repeat("A", maxMarkdownTableCellBytes*2)) {
		t.Fatalf("markdown kept oversized DOCX table cell, length=%d", len(md))
	}
}

func TestDOCXLongTableCellSamplesDoNotNeedAdditionalText(t *testing.T) {
	tests := []struct {
		name string
		want string
	}{
		{name: "60329.docx", want: "monocyte mRNA expression for pro-inflammatory cytokine and chemokine markers"},
		{name: "bib-chernigovka.netdo.ru_download_docs_17459.docx", want: "Сергея Павловича Королева"},
		{name: "deep-table-cell.docx", want: "Nested level 360"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			md := res.Markdown("images")
			if !strings.Contains(md, tc.want) {
				t.Fatalf("markdown missing long visible DOCX table text %q in %.1000q", tc.want, md)
			}
			if strings.Contains(md, "## Additional Text") {
				t.Fatalf("long DOCX table text should be structured instead of backfilled:\n%s", md)
			}
		})
	}
}

func TestDOCXDeepNumberedParagraphsDoNotNeedAdditionalText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "bug65649.docx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"В случае, когда функционал виртуальных разделов не применяется столбцы 11 и 12 Таблицы программирования ИС не заполняются.",
		"В случае, когда функционал виртуальных разделов не применяется, столбцы 11 и 12 Таблицы программирования ИС не заполняются.",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible deep numbered DOCX paragraph %q", want)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("deep numbered DOCX paragraphs should be structured instead of backfilled:\n%s", md)
	}
}

func TestDOCXEscapedTableTextDoesNotNeedAdditionalText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "bug59058.docx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	want := `167 Q1 1.00 \0.04 Q2 1.09 (0.48, 2.47) Q3 0.70 (0.27, 1.78) Q4 0.24** (0.07, 0.79)`
	if !strings.Contains(md, "167 Q1 1.00") || !strings.Contains(md, "Q4 6.34 (1.85, 21.73)") {
		t.Fatalf("markdown missing visible DOCX table statistics near %q", want)
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("escaped DOCX table text should be structured instead of backfilled:\n%s", md)
	}
}

func TestMarkdownIncludesEmbeddedOOXMLText(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Embedded DOCX Visible Text</w:t></w:r></w:p></w:body></w:document>`)
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Outer DOCX Visible Text</w:t></w:r></w:p><w:object r:id="rIdEmbedded"/></w:body></w:document>`)
	addZip(t, outerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/embedded1.docx"/></Relationships>`)
	addZipBytes(t, outerZip, "word/embeddings/embedded1.docx", inner.Bytes())
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "embedded-docx.docx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Outer DOCX Visible Text", "## Embedded Content", "Embedded DOCX Visible Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing embedded OOXML content %q in:\n%s", want, md)
		}
	}
}

func TestEmbeddedOOXMLImageNamesAreUnique(t *testing.T) {
	inner1 := testDocxPackage(t, "First embedded OOXML text", testPNG())
	inner2 := testDocxPackage(t, "Second embedded OOXML text", testPNG())

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Outer DOCX Visible Text</w:t></w:r></w:p><w:object r:id="rIdEmbedded1"/><w:object r:id="rIdEmbedded2"/></w:body></w:document>`)
	addZip(t, outerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/first/package.docx"/><Relationship Id="rIdEmbedded2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings\second\package.docx"/></Relationships>`)
	addZipBytes(t, outerZip, "word/embeddings/first/package.docx", inner1)
	addZipBytes(t, outerZip, "word/embeddings/second/package.docx", inner2)
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "embedded-ooxml-duplicate-images.docx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two embedded OOXML images, got %#v", res.Images)
	}
	names := map[string]bool{}
	for _, img := range res.Images {
		if img.Name == "" {
			t.Fatalf("expected embedded OOXML image name to be populated: %#v", img)
		}
		if names[strings.ToLower(img.Name)] {
			t.Fatalf("duplicate embedded OOXML image name %q in %#v", img.Name, res.Images)
		}
		names[strings.ToLower(img.Name)] = true
		if !validImageData(img.Ext, img.Data) {
			t.Fatalf("expected valid embedded OOXML image data for %#v", img)
		}
	}
	if !names["package.docx-image1.png"] || !names["package.docx-image1-2.png"] {
		t.Fatalf("expected stable unique embedded OOXML image names, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"First embedded OOXML text",
		"Second embedded OOXML text",
		"](images/package.docx-image1.png)",
		"](images/package.docx-image1-2.png)",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing embedded OOXML content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"embeddings", "second", `\`} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept embedded OOXML path component %q in:\n%s", hidden, md)
		}
	}
}

func TestEmbeddedOOXMLImageIsPlacedInEmbeddedMarkdown(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Embedded before image</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Embedded visible picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicture"/></p:blipFill></p:pic>
<w:p><w:r><w:t>Embedded after image</w:t></w:r></w:p>
</w:body></w:document>`)
	addZip(t, innerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline.png"/></Relationships>`)
	addZipBytes(t, innerZip, "word/media/inline.png", testPNG())
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Outer visible text</w:t></w:r></w:p><w:object r:id="rIdEmbedded"/></w:body></w:document>`)
	addZip(t, outerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/embedded1.docx"/></Relationships>`)
	addZipBytes(t, outerZip, "word/embeddings/embedded1.docx", inner.Bytes())
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "embedded-docx-image-placement.docx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "embedded1.docx-inline.png" || res.Images[0].Alt != "Embedded visible picture" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected renamed embedded OOXML image with alt, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Embedded Content", "Embedded before image", "Embedded visible picture\n![Embedded visible picture](images/embedded1.docx-inline.png)", "Embedded after image"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing placed embedded image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "](images/inline.png)") || strings.Contains(md, "## Images") {
		t.Fatalf("markdown kept stale or duplicate embedded image reference:\n%s", md)
	}
	if _, err := os.Stat(filepath.Join(dir, "images", "embedded1.docx-inline.png")); err != nil {
		t.Fatalf("renamed embedded image was not written: %v", err)
	}
}

func TestMarkdownPreservesEmbeddedOOXMLStructuredTables(t *testing.T) {
	inner := testXlsxPackage(t, "Embedded Header", "Embedded Value")

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "word/document.xml", `<w:document xmlns:w="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Outer DOCX Visible Text</w:t></w:r></w:p><w:object r:id="rIdEmbedded"/></w:body></w:document>`)
	addZip(t, outerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/embedded1.xlsx"/></Relationships>`)
	addZipBytes(t, outerZip, "word/embeddings/embedded1.xlsx", inner)
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "embedded-xlsx-table.docx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"## Document",
		"Outer DOCX Visible Text",
		"## Embedded Content",
		"### Sheet1",
		"| Embedded Header | Embedded Value |",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing embedded structured content %q in:\n%s", want, md)
		}
	}
}

func TestDOCXUnreferencedEmbeddedOOXMLIsNotVisibleText(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Unreferenced Embedded Secret</w:t></w:r></w:p></w:body></w:document>`)
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Outer DOCX Visible Text</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, outerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnused" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/embedded1.docx"/></Relationships>`)
	addZipBytes(t, outerZip, "word/embeddings/embedded1.docx", inner.Bytes())
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-embedded-docx.docx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Outer DOCX Visible Text") {
		t.Fatalf("missing outer visible text: %q", res.Text)
	}
	if strings.Contains(res.Text, "Unreferenced Embedded Secret") {
		t.Fatalf("unreferenced embedded text leaked into plain text: %q", res.Text)
	}
	md := res.Markdown("images")
	if strings.Contains(md, "Unreferenced Embedded Secret") || strings.Contains(md, "## Embedded Content") {
		t.Fatalf("unreferenced embedded text leaked into markdown:\n%s", md)
	}
}

func TestPPTXHiddenSlideEmbeddedOOXMLIsNotVisibleText(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, innerZip, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, innerZip, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Slide Embedded Secret</t></is></c></row></sheetData></worksheet>`)
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p"><p:cSld><p:spTree><p:sp><p:txBody><a:p xmlns:a="urn:a"><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, outerZip, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" show="0"><p:cSld><p:spTree><p:oleObj r:id="rIdEmbedded"/></p:spTree></p:cSld></p:sld>`)
	addZip(t, outerZip, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embedded1.xlsx"/></Relationships>`)
	addZipBytes(t, outerZip, "ppt/embeddings/embedded1.xlsx", inner.Bytes())
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-slide-embedded.pptx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Slide Text") {
		t.Fatalf("missing visible slide text: %q", res.Text)
	}
	if strings.Contains(res.Text, "Hidden Slide Embedded Secret") || strings.Contains(res.Markdown("images"), "Hidden Slide Embedded Secret") {
		t.Fatalf("hidden slide embedded text leaked: text=%q markdown=\n%s", res.Text, res.Markdown("images"))
	}
}

func TestXLSXHiddenSheetEmbeddedOOXMLIsNotVisibleText(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p"><p:cSld><p:spTree><p:sp><p:txBody><a:p xmlns:a="urn:a"><a:r><a:t>Hidden Sheet Embedded Secret</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/><sheet name="Hidden" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, outerZip, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, outerZip, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData></worksheet>`)
	addZip(t, outerZip, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Hidden Sheet Cell</t></is></c></row></sheetData><oleObjects><oleObject r:id="rIdEmbedded"/></oleObjects></worksheet>`)
	addZip(t, outerZip, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embedded1.pptx"/></Relationships>`)
	addZipBytes(t, outerZip, "xl/embeddings/embedded1.pptx", inner.Bytes())
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheet-embedded.xlsx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Sheet Text") {
		t.Fatalf("missing visible sheet text: %q", res.Text)
	}
	if strings.Contains(res.Text, "Hidden Sheet Embedded Secret") || strings.Contains(res.Markdown("images"), "Hidden Sheet Embedded Secret") {
		t.Fatalf("hidden sheet embedded text leaked: text=%q markdown=\n%s", res.Text, res.Markdown("images"))
	}
}

func TestXLSXUnreferencedSheetEmbeddedOOXMLIsNotVisibleText(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "word/document.xml", `<w:document xmlns:w="urn:x"><w:body><w:p><w:r><w:t>Orphan Sheet Embedded Secret</w:t></w:r></w:p></w:body></w:document>`)
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	var outer bytes.Buffer
	outerZip := zip.NewWriter(&outer)
	addZip(t, outerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, outerZip, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, outerZip, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, outerZip, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData></worksheet>`)
	addZip(t, outerZip, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Orphan Sheet Cell</t></is></c></row></sheetData><oleObjects><oleObject r:id="rIdEmbedded"/></oleObjects></worksheet>`)
	addZip(t, outerZip, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmbedded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embedded1.docx"/></Relationships>`)
	addZipBytes(t, outerZip, "xl/embeddings/embedded1.docx", inner.Bytes())
	if err := outerZip.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "orphan-sheet-embedded.xlsx")
	if err := os.WriteFile(file, outer.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Sheet Text") {
		t.Fatalf("missing visible sheet text: %q", res.Text)
	}
	md := res.Markdown("images")
	for _, hidden := range []string{"Orphan Sheet Cell", "Orphan Sheet Embedded Secret", "## Embedded Content"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("orphan sheet embedded content leaked %q: text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestWriteImagesTruncatesLongFilenames(t *testing.T) {
	dir := t.TempDir()
	longASCII := strings.Repeat("very-long-image-name-", 20) + ".png"
	longUnicode := strings.Repeat("图片文件名", 40) + ".png"
	images := []Image{
		{Name: longASCII, Ext: ".png", Data: testPNG()},
		{Name: longASCII, Ext: ".png", Data: testPNG()},
		{Name: longUnicode, Ext: ".png", Data: testPNG()},
	}
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	names := imageOutputFilenames(images)
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != len(images) {
		t.Fatalf("expected %d written images, got %d", len(images), len(entries))
	}
	seen := map[string]bool{}
	for _, entry := range entries {
		name := entry.Name()
		if len(name) > maxImageFilenameBytes {
			t.Fatalf("filename was not truncated: len=%d name=%q", len(name), name)
		}
		if !utf8.ValidString(name) {
			t.Fatalf("filename is not valid UTF-8: %q", name)
		}
		if filepath.Ext(name) != ".png" {
			t.Fatalf("expected png extension after truncation, got %q", name)
		}
		if seen[strings.ToLower(name)] {
			t.Fatalf("duplicate truncated filename %q", name)
		}
		seen[strings.ToLower(name)] = true
		b, err := os.ReadFile(filepath.Join(dir, name))
		if err != nil {
			t.Fatal(err)
		}
		if !validImageData(".png", b) {
			t.Fatalf("written truncated image is invalid: %s", name)
		}
	}
	if len(names) != len(images) {
		t.Fatalf("expected %d generated image names, got %d", len(images), len(names))
	}
	if !strings.HasSuffix(names[1], "-2.png") || strings.Contains(names[1], ".png-2.png") {
		t.Fatalf("duplicate truncated name should keep suffix before extension, got %q", names[1])
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, name := range names {
		target := "images/" + escapeMarkdownPath(filepath.ToSlash(name))
		if !strings.Contains(md, "]("+target+")") {
			t.Fatalf("markdown missing truncated image target %q in:\n%s", target, md)
		}
	}
}

func TestLegacyOLEImagesDeduplicateByContent(t *testing.T) {
	png := testPNG()
	jpeg := testJPEG()
	bmp, ok := dibToBMP(testDIB())
	if !ok {
		t.Fatal("test DIB did not convert to BMP")
	}
	images := imagesFromOLEStreams(nil, []oleStream{
		{Name: "first", Data: append(append([]byte("before"), png...), []byte("after")...)},
		{Name: "second", Data: append(append([]byte("again"), png...), []byte("tail")...)},
		{Name: "third", Data: append(append([]byte("other"), jpeg...), []byte("tail")...)},
		{Name: "fourth", Data: append(append([]byte("small"), bmp...), []byte("tail")...)},
		{Name: "fifth", Data: append(append([]byte("small again"), bmp...), []byte("tail")...)},
		{Name: "sixth", Data: append(append([]byte("small third"), bmp...), []byte("tail")...)},
	})
	if len(images) != 4 {
		t.Fatalf("expected normal duplicates plus one deduplicated repeated small image, got %#v", images)
	}
	if images[0].Ext != ".png" || !bytes.Equal(images[0].Data, png) {
		t.Fatalf("expected first unique PNG, got %#v", images[0])
	}
	if images[1].Ext != ".png" || !bytes.Equal(images[1].Data, png) {
		t.Fatalf("expected repeated normal-size PNG to be preserved, got %#v", images[1])
	}
	if images[2].Ext != ".jpg" || !bytes.Equal(images[2].Data, jpeg) {
		t.Fatalf("expected JPEG, got %#v", images[2])
	}
	if images[3].Ext != ".bmp" || !bytes.Equal(images[3].Data, bmp) {
		t.Fatalf("expected one deduplicated small BMP, got %#v", images[3])
	}
	if images[0].Name != "legacy-image-001.png" || images[1].Name != "legacy-image-002.png" || images[2].Name != "legacy-image-003.jpg" || images[3].Name != "legacy-image-004.bmp" {
		t.Fatalf("expected compact generated names, got %#v", images)
	}
}

func TestWrittenSampleImagesAreValid(t *testing.T) {
	samples := []string{
		"testPictures.doc",
		"pictures.ppt",
		"picture.xlsx",
		"VariousPictures.docx",
	}
	for _, sample := range samples {
		t.Run(sample, func(t *testing.T) {
			dir := t.TempDir()
			res, err := Extract(filepath.Join("testdata", "samples", sample), Options{ImageDir: dir})
			if err != nil {
				t.Fatal(err)
			}
			entries, err := os.ReadDir(dir)
			if err != nil {
				t.Fatal(err)
			}
			if len(entries) != len(res.Images) {
				t.Fatalf("expected %d written image(s), got %d", len(res.Images), len(entries))
			}
			written := map[string]bool{}
			for _, entry := range entries {
				if entry.IsDir() {
					t.Fatalf("unexpected directory in image output: %s", entry.Name())
				}
				written[entry.Name()] = true
				b, err := os.ReadFile(filepath.Join(dir, entry.Name()))
				if err != nil {
					t.Fatal(err)
				}
				ext := strings.ToLower(filepath.Ext(entry.Name()))
				if !validImageData(ext, b) {
					t.Fatalf("written image is not valid: %s len=%d", entry.Name(), len(b))
				}
			}
			markdownTargets := localMarkdownImageTargets(res.Markdown("images"), "images")
			if len(res.Images) > 0 && len(markdownTargets) == 0 {
				t.Fatalf("expected markdown to reference written image(s), got:\n%s", res.Markdown("images"))
			}
			for _, name := range markdownTargets {
				if !written[name] {
					t.Fatalf("markdown references image %q that was not written; written=%v\nmarkdown:\n%s", name, written, res.Markdown("images"))
				}
				b, err := os.ReadFile(filepath.Join(dir, name))
				if err != nil {
					t.Fatalf("markdown references missing written image %q: %v", name, err)
				}
				ext := strings.ToLower(filepath.Ext(name))
				if !validImageData(ext, b) {
					t.Fatalf("markdown references invalid written image: %s len=%d", name, len(b))
				}
			}
		})
	}
}

func localMarkdownImageTargets(markdown, imageBase string) []string {
	prefix := strings.TrimRight(strings.ReplaceAll(imageBase, "\\", "/"), "/") + "/"
	var targets []string
	for _, line := range strings.Split(markdown, "\n") {
		for start := strings.Index(line, "]("); start >= 0; start = strings.Index(line, "](") {
			line = line[start+2:]
			end := strings.IndexByte(line, ')')
			if end < 0 {
				break
			}
			target := line[:end]
			line = line[end+1:]
			if !strings.HasPrefix(target, prefix) {
				continue
			}
			name := strings.TrimPrefix(target, prefix)
			if name == "" || strings.Contains(name, "/") || strings.Contains(name, "\\") {
				continue
			}
			if unescaped, err := url.PathUnescape(name); err == nil {
				name = unescaped
			}
			targets = append(targets, name)
		}
	}
	return targets
}

func assertCleanExtractedText(t *testing.T, text string) {
	t.Helper()
	if strings.ContainsRune(text, '\uFFFD') {
		t.Fatalf("text contains replacement rune: %q", text)
	}
	for _, r := range text {
		if r == '\n' || r == '\t' {
			continue
		}
		if r < 0x20 || r == 0x7f {
			t.Fatalf("text contains control rune %U in %q", r, text)
		}
	}
}

func TestUTF16NoiseFiltersMisalignedASCII(t *testing.T) {
	if !looksLikeMisalignedUTF16([]uint16{0x5400, 0x6900, 0x6d00, 0x6500, 0x7300}) {
		t.Fatal("expected byte-swapped ASCII-looking UTF-16 to be rejected")
	}
	if looksLikeMisalignedUTF16(utf16.Encode([]rune("Times"))) {
		t.Fatal("expected normal UTF-16 ASCII to be accepted")
	}
	if looksLikeMisalignedUTF16(utf16.Encode([]rune("Résumé"))) {
		t.Fatal("expected normal UTF-16 non-ASCII text to be accepted")
	}
}

func TestTextQualityFiltersBinaryNoise(t *testing.T) {
	data := append([]byte("Valid legacy text\x00"), []byte("}~}{~}~}{~}~}{\x00")...)
	data = append(data, []byte("http://poi.apache.org/\x00")...)
	data = append(data, utf16LEBytes("Résumé Office")...)
	data = append(data, []byte("\x00PowerPoint.Slide.80\x00#ppt_w\x00___PPT10\x00<xml><control/></xml>\x00%PDF-1.4\x00AcroExch.Document.11\x00WordPad.Document.1\x00MediaPlayer.MediaPlayer.1\x00ShockwaveFlash.ShockwaveFlash.9\x00YYYYYYYYYYYY\x00j++j++j++\x00")...)
	data = append(data, []byte("{00020906-0000-0000-C000-000000000046}\x00CLSID={00020906-0000-0000-C000-000000000046}\x00ClassID:00020820-0000-0000-C000-000000000046\x00Visible ticket 00020906-0000-0000-C000-000000000046 remains visible\x00")...)
	data = append(data, 0, 0)
	parts := extractBinaryStrings(data)
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "Valid legacy text") {
		t.Fatalf("missing valid text in %q", text)
	}
	if !strings.Contains(text, "http://poi.apache.org/") {
		t.Fatalf("missing url text in %q", text)
	}
	if !strings.Contains(text, "Résumé Office") {
		t.Fatalf("missing unicode text in %q", text)
	}
	if strings.Contains(text, "}~}{") {
		t.Fatalf("kept binary-looking noise in %q", text)
	}
	if strings.Contains(text, "#ppt_w") || strings.Contains(text, "PowerPoint.Slide.80") || strings.Contains(text, "___PPT10") || strings.Contains(text, "<xml>") ||
		strings.Contains(text, "%PDF") || strings.Contains(text, "AcroExch") || strings.Contains(text, "WordPad.Document") ||
		strings.Contains(text, "MediaPlayer.MediaPlayer") || strings.Contains(text, "ShockwaveFlash.ShockwaveFlash") ||
		strings.Contains(text, "YYYYYYYY") || strings.Contains(text, "j++j") ||
		strings.Contains(text, "{00020906-0000-0000-C000-000000000046}") ||
		strings.Contains(text, "CLSID={00020906-0000-0000-C000-000000000046}") ||
		strings.Contains(text, "ClassID:00020820-0000-0000-C000-000000000046") {
		t.Fatalf("kept control-looking fragment in %q", text)
	}
	hiddenData := []byte("Visible fallback text\x00word/media/image1.png\x00xl/_rels/workbook.xml.rels\x00ppt%2Fmedia%2Fencoded%20image.png\x00")
	hiddenText := strings.Join(extractBinaryStrings(hiddenData), "\n")
	if !strings.Contains(hiddenText, "Visible fallback text") {
		t.Fatalf("missing visible fallback text in %q", hiddenText)
	}
	for _, hidden := range []string{"word/media/image1.png", "xl/_rels/workbook.xml.rels", "ppt%2Fmedia%2Fencoded%20image.png"} {
		if strings.Contains(hiddenText, hidden) {
			t.Fatalf("kept hidden Office resource reference %q in %q", hidden, hiddenText)
		}
	}
	wrapperData := []byte("CompObj\x00ObjInfo\x00Ole10Native\x00OlePres000\x00Object pooling notes are visible.\x00")
	wrapperText := strings.Join(extractBinaryStrings(wrapperData), "\n")
	for _, hidden := range []string{"CompObj", "ObjInfo", "Ole10Native", "OlePres000"} {
		if strings.Contains(wrapperText, hidden) {
			t.Fatalf("kept OLE wrapper stream name %q in %q", hidden, wrapperText)
		}
	}
	if !strings.Contains(wrapperText, "Object pooling notes are visible.") {
		t.Fatalf("dropped visible object-pooling prose in %q", wrapperText)
	}
	ooxmlMarkupData := []byte("Relationships\x00ContentType\x00PartName\x00TargetMode\x00Customer relationships are visible.\x00")
	ooxmlMarkupText := strings.Join(extractBinaryStrings(ooxmlMarkupData), "\n")
	for _, hidden := range []string{"Relationships", "ContentType", "PartName", "TargetMode"} {
		if strings.Contains(ooxmlMarkupText, hidden) {
			t.Fatalf("kept OOXML markup name %q in %q", hidden, ooxmlMarkupText)
		}
	}
	if !strings.Contains(ooxmlMarkupText, "Customer relationships are visible.") {
		t.Fatalf("dropped visible relationship prose in %q", ooxmlMarkupText)
	}
	ooxmlAttributeData := []byte(strings.Join([]string{
		`Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden.png" TargetMode="External"`,
		`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`,
		`mc:Ignorable="w14 wp14"`,
		`xsi:schemaLocation="http://schemas.openxmlformats.org/wordprocessingml/2006/main wordprocessingml.xsd"`,
		`ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"`,
		`PartName="/word/document.xml"`,
		`Visible relationship guidance remains user text.`,
		`Visible schema discussion remains user text.`,
		`Visible before xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" after`,
		`Visible before Target="../media/hidden.png" after`,
	}, "\x00"))
	ooxmlAttributeText := strings.Join(extractBinaryStrings(ooxmlAttributeData), "\n")
	for _, want := range []string{
		"Visible relationship guidance remains user text.",
		"Visible schema discussion remains user text.",
		"Visible before after",
	} {
		if !strings.Contains(ooxmlAttributeText, want) {
			t.Fatalf("fallback text missing visible OOXML attribute fragment %q in %q", want, ooxmlAttributeText)
		}
	}
	for _, hidden := range []string{"Relationship Id", "rId7", "relationships/image", "../media/hidden.png", "TargetMode", "External", "xmlns:w", "schemas.openxmlformats.org", "mc:Ignorable", "w14 wp14", "schemaLocation", "wordprocessingml.xsd", "ContentType", "application/vnd.openxmlformats", "PartName", "/word/document.xml"} {
		if strings.Contains(ooxmlAttributeText, hidden) {
			t.Fatalf("fallback text kept OOXML attribute metadata %q in %q", hidden, ooxmlAttributeText)
		}
	}
	inlineHiddenData := []byte(strings.Join([]string{
		`Visible before Target="../media/inline.png" visible after`,
		`Visible type ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" after`,
		`Visible part PartName="/word/document.xml" after`,
		`Visible relationship r:embed="rId5" after`,
		`Visible CSS background:url(word/media/fallback-css.png) after`,
		`Visible VML src:word/media/fallback-colon.png after`,
		`Visible MHTML Content-Location: word/media/fallback-mhtml.png after`,
		`Visible CID Content-ID: <image002.png@office> after`,
		`Visible MIME Content-Type: image/png after`,
		`Keep C:\Reports\Q1 and https://example.test/media/photo.png visible`,
		`Keep CSS background:url(https://example.test/media/photo.png) visible`,
	}, "\x00"))
	inlineHiddenText := strings.Join(extractBinaryStrings(inlineHiddenData), "\n")
	for _, want := range []string{
		"Visible before visible after",
		"Visible type after",
		"Visible part after",
		"Visible relationship after",
		"Visible CSS background: after",
		"Visible VML after",
		"Visible MHTML after",
		"Visible CID after",
		"Visible MIME after",
		`Keep C:\Reports\Q1 and https://example.test/media/photo.png visible`,
		`Keep CSS background:url(https://example.test/media/photo.png) visible`,
	} {
		if !strings.Contains(inlineHiddenText, want) {
			t.Fatalf("fallback text missing cleaned visible fragment %q in %q", want, inlineHiddenText)
		}
	}
	for _, hidden := range []string{"Target=", "../media/inline.png", "ContentType=", "application/vnd.openxmlformats", "PartName=", "/word/document.xml", "r:embed", "rId5", "word/media/fallback-css.png", "word/media/fallback-colon.png", "Content-Location", "word/media/fallback-mhtml.png", "Content-ID", "image002.png@office", "Content-Type", "image/png"} {
		if strings.Contains(inlineHiddenText, hidden) {
			t.Fatalf("fallback text kept inline hidden Office reference %q in %q", hidden, inlineHiddenText)
		}
	}
	embeddedZipData := []byte("Visible legacy wrapper before\x00Ole10Native\x00C:\\Users\\me\\Desktop\\hidden-embedded.docx\x00hidden-embedded.docx\x00")
	embeddedZipData = append(embeddedZipData, testDocxPackage(t, "Hidden embedded OOXML text", nil)...)
	embeddedZipData = append(embeddedZipData, []byte("\x00Visible legacy wrapper after\x00")...)
	embeddedZipText := strings.Join(extractBinaryStrings(embeddedZipData), "\n")
	for _, want := range []string{"Visible legacy wrapper before", "Visible legacy wrapper after"} {
		if !strings.Contains(embeddedZipText, want) {
			t.Fatalf("fallback text missing visible wrapper text %q in %q", want, embeddedZipText)
		}
	}
	for _, hidden := range []string{"Hidden embedded OOXML text", "Ole10Native", "C:\\Users\\me\\Desktop\\hidden-embedded.docx", "hidden-embedded.docx", "[Content_Types].xml", "word/document.xml", "word/_rels", "application/vnd.openxmlformats", "PK"} {
		if strings.Contains(embeddedZipText, hidden) {
			t.Fatalf("fallback text kept embedded OOXML ZIP content %q in %q", hidden, embeddedZipText)
		}
	}
}

func TestOLEIdentifierFragmentsAreFilteredConservatively(t *testing.T) {
	for _, hidden := range []string{
		"{00020906-0000-0000-C000-000000000046}",
		"CLSID={00020906-0000-0000-C000-000000000046}",
		"ClassID:00020820-0000-0000-C000-000000000046",
	} {
		if !looksLikeBinaryControlFragment(hidden) {
			t.Fatalf("expected OLE identifier fragment to be filtered: %q", hidden)
		}
		if got := cleanText(hidden); got != "" {
			t.Fatalf("cleanText should drop OLE identifier fragment %q, got %q", hidden, got)
		}
	}
	visible := "Visible ticket 00020906-0000-0000-C000-000000000046 remains visible"
	if !looksLikeTextFragment(visible) {
		t.Fatalf("visible GUID prose should still look textual")
	}
	if looksLikeBinaryControlFragment(visible) {
		t.Fatalf("visible GUID prose should not be a control fragment")
	}
	if looksLikeHiddenResourceReference(visible) {
		t.Fatalf("visible GUID prose should not be a hidden resource reference")
	}
	if got := cleanText(visible); got != visible {
		t.Fatalf("cleanText should preserve visible GUID prose, got %q", got)
	}
	text := strings.Join(extractBinaryStrings([]byte(
		"{00020906-0000-0000-C000-000000000046}\x00"+
			"CLSID={00020906-0000-0000-C000-000000000046}\x00"+
			"ClassID:00020820-0000-0000-C000-000000000046\x00"+
			visible+"\x00",
	)), "\n")
	for _, hidden := range []string{
		"{00020906-0000-0000-C000-000000000046}",
		"CLSID={00020906-0000-0000-C000-000000000046}",
		"ClassID:00020820-0000-0000-C000-000000000046",
	} {
		if strings.Contains(text, hidden) {
			t.Fatalf("fallback text kept OLE identifier %q in %q", hidden, text)
		}
	}
	if !strings.Contains(text, visible) {
		t.Fatalf("fallback text dropped visible GUID prose in %q", text)
	}
}

func TestXLSXRepeatedHugeSharedStringIsDeduped(t *testing.T) {
	large := "Visible repeated cell " + strings.Repeat("A", maxRepeatedTextPartBytes)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x"><sheets><sheet name="Sheet1"/></sheets></workbook>`)
	addZip(t, zw, "xl/sharedStrings.xml", `<sst xmlns="urn:x"><si><t>`+large+`</t></si></sst>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="s"><v>0</v></c><c t="s"><v>0</v></c><c t="s"><v>0</v></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "huge-shared.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if count := strings.Count(res.Text, large); count != 1 {
		t.Fatalf("expected huge shared string once, got %d occurrences and len %d", count, len(res.Text))
	}
}

func TestLegacyPPTTextAtomsPreferOriginalText(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptRecord(0x0fa8, []byte("Slide title from bytes")))
	ppt.Write(pptRecord(0x0fa0, utf16LEBytes("正文来自 TextCharsAtom")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("This is a test comment")))
	ppt.Write(pptRecord(0x03e8, []byte("#ppt_w PowerPoint.Slide.80 ___PPT10 <xml><control/></xml>")))
	parts := extractLegacyText("sample.ppt", nil, []oleStream{
		{Name: "PowerPoint Document", Data: ppt.Bytes()},
		{Name: "Pictures", Data: []byte("Picture stream control text should not appear")},
	})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "Slide title from bytes") {
		t.Fatalf("missing TextBytesAtom text in %q", text)
	}
	if !strings.Contains(text, "正文来自 TextCharsAtom") {
		t.Fatalf("missing TextCharsAtom text in %q", text)
	}
	if !strings.Contains(text, "This is a test comment") {
		t.Fatalf("missing CString comment text in %q", text)
	}
	if strings.Contains(text, "Picture stream control") || strings.Contains(text, "#ppt_w") || strings.Contains(text, "PowerPoint.Slide.80") || strings.Contains(text, "___PPT10") {
		t.Fatalf("kept non-slide control text in %q", text)
	}
}

func TestLegacyPPTCStringDoesNotMisdecodeCompressedASCII(t *testing.T) {
	if text, ok := decodePPTCString([]byte("Microsoft PowerPoint")); !ok || text != "Microsoft PowerPoint" {
		t.Fatalf("expected compressed ASCII CString, got %q ok=%v", text, ok)
	}
}

func TestLegacyPPTEmbeddedObjectLabelsAreFiltered(t *testing.T) {
	var ppt bytes.Buffer
	for _, label := range []string{
		"Equation",
		"Chart",
		"Microsoft Graph Chart",
		"Microsoft Graph 97 Chart",
		"Microsoft Graph 2000 Chart",
		"MS Org Chart",
		"MS Organization Chart 2.0",
		"Microsoft Equation 3.0",
		"Package",
		"Package Object",
		"Packager Shell Object",
		"Microsoft Word Document",
		"Microsoft Office Word Document",
		"Microsoft Excel Worksheet",
		"Microsoft Office Excel Worksheet",
		"Microsoft PowerPoint Presentation",
		"Microsoft Office PowerPoint Presentation",
		"Adobe Acrobat Document",
		"Acrobat Document",
		"PDF Document",
		"Microsoft Visio Drawing",
		"Document",
		"Worksheet",
		"Word.Document.8",
		"Excel.Sheet.8",
		"Excel.Chart.8",
		"PowerPoint.Slide.8",
		"PowerPoint.Show.8",
		"AcroExch.Document.7",
		"Visio.Drawing.11",
		"Forms.TextBox.1",
		"Forms.CheckBox.1",
		"htmlfile",
	} {
		ppt.Write(pptRecord(0x0fa8, []byte(label)))
	}
	ppt.Write(pptRecord(0x0fa8, []byte("Visible slide text")))
	ppt.Write(pptRecord(0x0fa0, utf16LEBytes("Chart revenue by quarter.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Equation (2) gives the force.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("MS Org Chart migration notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Microsoft Equation 3.0 compatibility is visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Package migration notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("PDF document review notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Microsoft Visio Drawing notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Word document migration notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Excel chart review notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("PowerPoint slide review notes are visible.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Forms textbox behavior is documented.")))
	ppt.Write(pptRecord(0x0fba, []byte("Document the worksheet review notes.")))

	parts := extractLegacyText("sample.ppt", nil, []oleStream{
		{Name: "PowerPoint Document", Data: ppt.Bytes()},
	})
	text := strings.Join(parts, "\n")
	for _, want := range []string{
		"Visible slide text",
		"Chart revenue by quarter.",
		"Equation (2) gives the force.",
		"MS Org Chart migration notes are visible.",
		"Microsoft Equation 3.0 compatibility is visible.",
		"Package migration notes are visible.",
		"PDF document review notes are visible.",
		"Microsoft Visio Drawing notes are visible.",
		"Word document migration notes are visible.",
		"Excel chart review notes are visible.",
		"PowerPoint slide review notes are visible.",
		"Forms textbox behavior is documented.",
		"Document the worksheet review notes.",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible text %q in %q", want, text)
		}
	}
	for _, bad := range []string{
		"Equation\n",
		"Chart\n",
		"Microsoft Graph Chart",
		"Microsoft Graph 97 Chart",
		"Microsoft Graph 2000 Chart",
		"MS Org Chart\n",
		"MS Organization Chart 2.0",
		"Microsoft Equation 3.0\n",
		"Package\n",
		"Package Object",
		"Packager Shell Object",
		"Microsoft Word Document",
		"Microsoft Office Word Document",
		"Microsoft Excel Worksheet",
		"Microsoft Office Excel Worksheet",
		"Microsoft PowerPoint Presentation",
		"Microsoft Office PowerPoint Presentation",
		"Adobe Acrobat Document",
		"Acrobat Document",
		"PDF Document\n",
		"Microsoft Visio Drawing\n",
		"Document\n",
		"Worksheet\n",
		"Word.Document.8",
		"Excel.Sheet.8",
		"Excel.Chart.8",
		"PowerPoint.Slide.8",
		"PowerPoint.Show.8",
		"AcroExch.Document.7",
		"Visio.Drawing.11",
		"Forms.TextBox.1",
		"Forms.CheckBox.1",
		"htmlfile\n",
	} {
		if strings.Contains(text+"\n", bad) {
			t.Fatalf("kept PPT object label %q in %q", bad, text)
		}
	}
}

func TestLegacyPPTMarkdownUsesSlideSections(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptContainerRecord(0x03ee, bytes.Join([][]byte{
		pptRecord(0x0fa8, []byte("First slide title")),
		pptRecord(0x0fa0, utf16LEBytes("First slide body")),
		pptRecord(0x0fa8, []byte("Chart")),
	}, nil)))
	ppt.Write(pptContainerRecord(0x03ee, bytes.Join([][]byte{
		pptRecord(0x0fa8, []byte("Second slide title")),
		pptRecord(0x0fba, utf16LEBytes("Second slide speaker note")),
		pptRecord(0x03e8, []byte("#ppt_w PowerPoint.Slide.80 ___PPT10 <xml><control/></xml>")),
	}, nil)))
	md := extractLegacyMarkdown("sample.ppt", nil, []oleStream{{Name: "PowerPoint Document", Data: ppt.Bytes()}})
	for _, want := range []string{"## Slide 1", "First slide title", "First slide body", "## Slide 2", "Second slide title", "## Notes and Comments", "Second slide speaker note"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy PPT markdown missing %q in:\n%s", want, md)
		}
	}
	if strings.Index(md, "## Notes and Comments") < strings.Index(md, "## Slide 2") {
		t.Fatalf("legacy PPT notes/comments should follow slide sections:\n%s", md)
	}
	slide2 := md
	if start := strings.Index(md, "## Slide 2"); start >= 0 {
		slide2 = md[start:]
	}
	if end := strings.Index(slide2, "## Notes and Comments"); end >= 0 {
		slide2 = slide2[:end]
	}
	if strings.Contains(slide2, "Second slide speaker note") {
		t.Fatalf("legacy PPT speaker note should be grouped outside slide body:\n%s", md)
	}
	for _, bad := range []string{"Chart\n", "#ppt_w", "PowerPoint.Slide.80", "___PPT10", "## Additional Text"} {
		if strings.Contains(md+"\n", bad) {
			t.Fatalf("legacy PPT markdown kept control/object text %q in:\n%s", bad, md)
		}
	}
}

func TestLegacyPPTMarkdownGroupsTopLevelCStringsAsNotesAndComments(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptRecord(0x0fa8, []byte("Visible deck title")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Visible top-level comment")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes(`Visible note Target="../media/hidden.png" text`)))
	ppt.Write(pptRecord(0x0fba, []byte(`PartName="/ppt/slides/slide1.xml"`)))
	md := extractLegacyMarkdown("sample.ppt", nil, []oleStream{{Name: "PowerPoint Document", Data: ppt.Bytes()}})
	for _, want := range []string{"## Presentation", "Visible deck title", "## Notes and Comments", "Visible top-level comment", "Visible note text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy PPT markdown missing grouped CString content %q in:\n%s", want, md)
		}
	}
	if strings.Index(md, "## Notes and Comments") < strings.Index(md, "## Presentation") {
		t.Fatalf("legacy PPT notes/comments should follow presentation text:\n%s", md)
	}
	presentation := md
	if end := strings.Index(presentation, "## Notes and Comments"); end >= 0 {
		presentation = presentation[:end]
	}
	if strings.Contains(presentation, "Visible top-level comment") {
		t.Fatalf("legacy PPT comment should be grouped outside presentation body:\n%s", md)
	}
	for _, hidden := range []string{"Target=", "../media/hidden.png", "PartName", "/ppt/slides/slide1.xml"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy PPT markdown kept hidden note/comment reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyPPTFiltersDesignThemeLabels(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptRecord(0x0fa8, []byte("Visible slide title")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Default Design")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Office Theme")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Office theme migration notes are visible.")))
	streams := []oleStream{{Name: "PowerPoint Document", Data: ppt.Bytes()}}
	text := strings.Join(extractLegacyText("sample.ppt", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.ppt", nil, streams)
	for _, want := range []string{"Visible slide title", "Office theme migration notes are visible."} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy PPT text missing visible content %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy PPT markdown missing visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Default Design", "Office Theme\n"} {
		if strings.Contains(text+"\n", hidden) {
			t.Fatalf("legacy PPT text kept design/theme label %q in %q", hidden, text)
		}
		if strings.Contains(md+"\n", hidden) {
			t.Fatalf("legacy PPT markdown kept design/theme label %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyPPTWithCommentsSampleDropsDesignThemeLabel(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "WithComments.ppt"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Test Slide", "With a comment on it"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("WithComments.ppt text missing visible content %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("WithComments.ppt markdown missing visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Default Design", "Office Theme"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("WithComments.ppt text kept design/theme label %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("WithComments.ppt markdown kept design/theme label %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyPPTJapaneseSampleDropsMasterPlaceholdersAndTheme(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "54880_chinese.ppt"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Single byte", "複数の文字", "カタカナ", "ﾊﾝｶｸ", "表十ソ", "𠮟"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("54880_chinese.ppt text missing visible content %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("54880_chinese.ppt markdown missing visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"マスタ タイトルの書式設定", "マスタ テキストの書式設定", "Office テーマ"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("54880_chinese.ppt text kept master/theme label %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("54880_chinese.ppt markdown kept master/theme label %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyPPTPlaceholderPromptsAreNotVisibleText(t *testing.T) {
	var ppt bytes.Buffer
	for _, prompt := range []string{
		"Click to add title",
		"Click to add subtitle",
		"Click to add text",
		"Click to add content",
		"Click to add vertical title",
		"Click to add vertical text",
		"Click to add chart",
		"Click to add table",
		"Click to add diagram",
		"Click to add organization chart",
		"Click to add media clip",
		"Click to add clip art",
		"Click to add picture",
		"Click to add object",
		"Click to add date",
		"Click to add footer",
		"Click to add slide number",
		"Click icon to add chart",
		"Click icon to add table",
		"Click icon to add SmartArt graphic",
		"Click icon to add picture",
		"Click icon to add media clip",
		"Double click to add chart",
		"Double click to add table",
		"Double click to add diagram",
		"Double click to add organization chart",
		"Double click to add clip art",
		"Double click to add picture",
		"Double click to add media clip",
		"Double click to add object",
		"Click to edit Master title style",
		"Click to edit Master text styles",
		"Click to edit Master body text styles",
		"Click to edit Master footer style",
	} {
		ppt.Write(pptRecord(0x0fa8, []byte(prompt)))
	}
	ppt.Write(pptRecord(0x0fa8, []byte("Visible slide title")))
	ppt.Write(pptRecord(0x0fa0, utf16LEBytes("Visible slide body")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Users may literally write Click to add title in training notes.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Training notes mention Click to add chart without being a template prompt.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Training notes mention Double click to add chart without being a template prompt.")))
	ppt.Write(pptRecord(0x0fba, utf16LEBytes("Users may write Click icon to add picture in instructions.")))

	streams := []oleStream{{Name: "PowerPoint Document", Data: ppt.Bytes()}}
	text := strings.Join(extractLegacyText("sample.ppt", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.ppt", nil, streams)
	for _, want := range []string{
		"Visible slide title",
		"Visible slide body",
		"Users may literally write Click to add title in training notes.",
		"Training notes mention Click to add chart without being a template prompt.",
		"Training notes mention Double click to add chart without being a template prompt.",
		"Users may write Click icon to add picture in instructions.",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy PPT text missing visible content %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy PPT markdown missing visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{
		"Click to add title",
		"Click to add subtitle",
		"Click to add text",
		"Click to add content",
		"Click to add vertical title",
		"Click to add vertical text",
		"Click to add chart",
		"Click to add table",
		"Click to add diagram",
		"Click to add organization chart",
		"Click to add media clip",
		"Click to add clip art",
		"Click to add picture",
		"Click to add object",
		"Click to add date",
		"Click to add footer",
		"Click to add slide number",
		"Click icon to add chart",
		"Click icon to add table",
		"Click icon to add SmartArt graphic",
		"Click icon to add picture",
		"Click icon to add media clip",
		"Double click to add chart",
		"Double click to add table",
		"Double click to add diagram",
		"Double click to add organization chart",
		"Double click to add clip art",
		"Double click to add picture",
		"Double click to add media clip",
		"Double click to add object",
		"Click to edit Master title style",
		"Click to edit Master text styles",
		"Click to edit Master body text styles",
		"Click to edit Master footer style",
	} {
		if textContainsWholeLine(text, hidden) {
			t.Fatalf("legacy PPT text kept placeholder prompt %q in %q", hidden, text)
		}
		if textContainsWholeLine(md, hidden) {
			t.Fatalf("legacy PPT markdown kept placeholder prompt %q in:\n%s", hidden, md)
		}
	}
}

func textContainsWholeLine(text, line string) bool {
	for _, candidate := range strings.Split(text, "\n") {
		if strings.TrimSpace(candidate) == line {
			return true
		}
	}
	return false
}

func TestLegacyPPTStripsInlineHiddenResourceReferences(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptContainerRecord(0x03ee, bytes.Join([][]byte{
		pptRecord(0x0fa8, []byte("Visible slide title Content-Location: ppt/media/image1.png")),
		pptRecord(0x0fa0, utf16LEBytes("Visible slide body Target=\"../media/hidden.jpg\"")),
		pptRecord(0x0fba, []byte("Speaker note keeps words Content-Type: image/png url(ppt/media/background.png)")),
		pptRecord(0x0fa8, []byte("Visible chart label Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image cleaned")),
		pptRecord(0x0fa0, utf16LEBytes("Visible part marker PartName=\"/ppt/slides/slide1.xml\" cleaned")),
		pptRecord(0x0fba, utf16LEBytes("Visible embed marker r:embed=\"rId7\" cleaned")),
		pptRecord(0x0fa8, []byte(`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`)),
		pptRecord(0x0fa0, utf16LEBytes(`ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"`)),
		pptRecord(0x0fba, []byte(`PartName="/ppt/slides/slide1.xml"`)),
		pptRecord(0x0fba, utf16LEBytes(`TargetMode="External"`)),
		pptRecord(0x0fba, utf16LEBytes(`r:embed="rId42"`)),
	}, nil)))

	streams := []oleStream{{Name: "PowerPoint Document", Data: ppt.Bytes()}}
	text := strings.Join(extractLegacyText("sample.ppt", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.ppt", nil, streams)
	for _, want := range []string{"Visible slide title", "Visible slide body", "Speaker note keeps words", "Visible chart label cleaned", "Visible part marker cleaned", "Visible embed marker cleaned"} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy PPT text missing visible content %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy PPT markdown missing visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Location", "ppt/media/image1.png", "Target=", "../media/hidden.jpg", "Content-Type", "image/png", "url(", "ppt/media/background.png", "Type:", "relationships/image", "ContentType", "presentationml.slide", "PartName", "/ppt/slides/slide1.xml", "TargetMode", "External", "r:embed", "rId7", "rId42"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy PPT text kept hidden reference %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy PPT markdown kept hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyPPTFallbackStripsInlineHiddenResourceReferences(t *testing.T) {
	fallback := []byte("Visible fallback words Content-Location: ppt/media/image1.png\x00Another visible fallback line Target=\"../media/hidden.jpg\"\x00ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"\x00TargetMode=\"External\"\x00r:embed=\"rId77\"\x00")
	streams := []oleStream{{Name: "PP40", Data: fallback}}
	text := strings.Join(extractLegacyText("sample.ppt", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.ppt", nil, streams)
	for _, want := range []string{"Visible fallback words", "Another visible fallback line"} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy PPT fallback text missing visible content %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy PPT fallback markdown missing visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Content-Location", "ppt/media/image1.png", "Target=", "../media/hidden.jpg", "ContentType", "presentationml.slide", "TargetMode", "External", "r:embed", "rId77"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy PPT fallback text kept hidden reference %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy PPT fallback markdown kept hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyPPTDropsLeadingRecordControlLine(t *testing.T) {
	var ppt bytes.Buffer
	ppt.Write(pptRecord(0x0fa8, []byte("0\vVisible slide title")))
	text := joinText(pptRecordText(ppt.Bytes()))
	if strings.HasPrefix(text, "0\n") || strings.HasPrefix(text, "0 ") || text == "0" {
		t.Fatalf("kept leading PPT record control line in %q", text)
	}
	if !strings.HasPrefix(text, "Visible slide title") {
		t.Fatalf("missing visible PPT text, got %q", text)
	}
}

func TestLegacyPPTSampleDropsMojibakeControlLines(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "53446.ppt"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.HasPrefix(res.Text, "Enron Corp. - Business Units") {
		t.Fatalf("expected visible text prefix, got %.120q", res.Text)
	}
	for _, bad := range []string{"я0;я", "0000ея", "、。，．・", "MS Org Chart", "Microsoft Graph 97 Chart"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept PPT mojibake/control text %q in %.400q", bad, res.Text)
		}
	}
	if !strings.Contains(res.Text, "Risk Assessment & Control Group") {
		t.Fatalf("missing visible PPT text in %.400q", res.Text)
	}
}

func TestDocx223624SampleDropsCyrillicMojibakeControlLines(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "web-samples", "samples", "docx", "223624.docx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "VHA Comprehensive Emergency Management Program") {
		t.Fatalf("missing visible DOCX text in %.400q", res.Text)
	}
	bad := "00яяы0яяяя›0њ0э0ю00ћ00ь0 я0=я]я 00"
	if strings.Contains(res.Text, bad) {
		t.Fatalf("kept DOCX Cyrillic mojibake control text %q in %.400q", bad, res.Text)
	}
}

func TestLegacyPPT95SampleDoesNotEmitFalseUTF16(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "PPT95.ppt"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(res.Text, "捩潲潳") || strings.Contains(res.Text, "开釀") {
		t.Fatalf("kept false UTF-16 text in %.400q", res.Text)
	}
	if strings.Contains(res.Text, "YYYYYYY") || strings.Contains(res.Text, "j++j") {
		t.Fatalf("kept low-information binary text in %.400q", res.Text)
	}
	for _, bad := range []string{"\u8100\u8165\u8167\u8169\u816b\u816d\u816f\u8171\u8173\u8175\u8177\u8179\u818f\u2490", "\u5700\u6168\u2074", "\u4d00\u6e61\u2079", "\u4a00\u7375\u2074"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept PPT95 byte-aligned mojibake %q in %.400q", bad, res.Text)
		}
	}
	if !strings.Contains(res.Text, "XML, HTML and All That") {
		t.Fatalf("missing expected PPT text in %.400q", res.Text)
	}
}

func TestLegacyPPT40UsesPP40TextWithoutBinaryPrefix(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "pp40only.ppt"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "COMMUNITY ORGANISATIONS") || !strings.Contains(res.Text, "Is there any benefit for us if we register for GST?") {
		t.Fatalf("missing PP40 body text in %.400q", res.Text)
	}
	for _, noise := range []string{"3333", ";09;", "37y9;9"} {
		if strings.Contains(res.Text, noise) {
			t.Fatalf("kept PP40 binary prefix %q in %.400q", noise, res.Text)
		}
	}
}

func TestLegacyDOCUsesWordDocumentTextRange(t *testing.T) {
	body := []byte("Actual Word body text")
	word := make([]byte, 96+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5ec)
	binary.LittleEndian.PutUint32(word[0x18:], 96)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(96+len(body)))
	copy(word[96:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{
		{Name: "WordDocument", Data: word},
		{Name: "1Table", Data: []byte("StyleSheet control text should not appear")},
	})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "Actual Word body text") {
		t.Fatalf("missing WordDocument text range in %q", text)
	}
	if strings.Contains(text, "StyleSheet control") {
		t.Fatalf("kept table stream control text in %q", text)
	}
}

func TestLegacyDOCMarkdownUsesDocumentSection(t *testing.T) {
	body := []byte("Actual Word body text")
	word := make([]byte, 96+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5ec)
	binary.LittleEndian.PutUint32(word[0x18:], 96)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(96+len(body)))
	copy(word[96:], body)
	md := extractLegacyMarkdown("sample.doc", nil, []oleStream{
		{Name: "WordDocument", Data: word},
		{Name: "1Table", Data: []byte("StyleSheet control text should not appear")},
	})
	if !strings.Contains(md, "## Document") || !strings.Contains(md, "Actual Word body text") {
		t.Fatalf("legacy DOC markdown missing document text:\n%s", md)
	}
	if strings.Contains(md, "StyleSheet control") || strings.Contains(md, "## Additional Text") {
		t.Fatalf("legacy DOC markdown kept control/fallback text:\n%s", md)
	}
}

func TestLegacyDOCMarkdownDoesNotBackfillStructuredTableText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "table-merges.doc"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Document") || !strings.Contains(md, "ABCDEFGHIJK") {
		t.Fatalf("legacy DOC markdown missing visible table text:\n%s", md)
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("legacy DOC markdown duplicated already structured table text:\n%s", md)
	}
}

func TestLegacyDOCTextStripsInlineHiddenOfficeReferences(t *testing.T) {
	body := []byte(strings.Join([]string{
		"Visible before word/media/image1.png visible after",
		`Visible target Target: ../media/inline.png after`,
		`Visible type ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" after`,
		"Keep Target audience as visible prose",
	}, "\n"))
	word := make([]byte, 96+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5ec)
	binary.LittleEndian.PutUint32(word[0x18:], 96)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(96+len(body)))
	copy(word[96:], body)
	streams := []oleStream{{Name: "WordDocument", Data: word}}
	text := strings.Join(extractLegacyText("sample.doc", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.doc", nil, streams)
	for _, want := range []string{"Visible before visible after", "Visible target after", "Visible type after", "Keep Target audience as visible prose"} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy DOC text missing cleaned visible text %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy DOC markdown missing cleaned visible text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"word/media/image1.png", "Target:", "../media/inline.png", "ContentType", "application/vnd.openxmlformats"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy DOC text kept hidden reference %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy DOC markdown kept hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyFallbackMarkdownUsesFormatHeadings(t *testing.T) {
	for _, tc := range []struct {
		name    string
		heading string
	}{
		{name: "fallback.doc", heading: "## Document"},
		{name: "fallback.ppt", heading: "## Presentation"},
		{name: "fallback.xls", heading: "## Workbook"},
	} {
		md := legacyFallbackMarkdown(tc.name, []string{"Visible fallback text", "", "Visible fallback text"})
		if !strings.Contains(md, tc.heading) || !strings.Contains(md, "Visible fallback text") {
			t.Fatalf("fallback markdown for %s missing heading/text:\n%s", tc.name, md)
		}
		if strings.Count(md, "Visible fallback text") != 1 {
			t.Fatalf("fallback markdown should deduplicate text for %s:\n%s", tc.name, md)
		}
	}
	if md := legacyFallbackMarkdown("fallback.bin", []string{"Visible fallback text"}); md != "" {
		t.Fatalf("unexpected fallback markdown for unknown extension:\n%s", md)
	}
}

func TestLegacyDOCUsesOldWordTextRange(t *testing.T) {
	body := []byte("Old Word visible body text")
	tail := []byte("Default Paragraph Font")
	word := make([]byte, 128+len(body)+len(tail))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	copy(word[128+len(body):], tail)
	parts := extractLegacyText("sample.doc", nil, []oleStream{
		{Name: "CompObj", Data: []byte("Microsoft Word 6.0 Document\x00MSWordDoc\x00Word.Document.6")},
		{Name: "WordDocument", Data: word},
	})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "Old Word visible body text") {
		t.Fatalf("missing old Word text range in %q", text)
	}
	if strings.Contains(text, "MSWordDoc") || strings.Contains(text, "Default Paragraph Font") {
		t.Fatalf("kept old Word control text in %q", text)
	}
}

func TestLegacyDOCOldWordDecodesWindows1252(t *testing.T) {
	body := []byte("My son\x92s r\xe9sum\xe9 is visible")
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "son\u2019s r\u00e9sum\u00e9") {
		t.Fatalf("missing decoded Windows-1252 text in %q", text)
	}
	if strings.Contains(text, "son\ns") || strings.ContainsRune(text, '\uFFFD') {
		t.Fatalf("kept split or replacement text in %q", text)
	}
}

func TestLegacyDOCOldWordDecodesWindows1251(t *testing.T) {
	body := []byte{0xce, 0xcf, 0xc5, 0xd0, 0xc0, 0xd6, 0xc8, 0xdf, ' ', 0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2}
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "ОПЕРАЦИЯ Привет") {
		t.Fatalf("missing decoded Windows-1251 text in %q", text)
	}
	if strings.Contains(text, "ÎÏÅ") {
		t.Fatalf("kept Windows-1251 mojibake in %q", text)
	}
	if strings.Contains(text, "쿞") || strings.Contains(text, "탅") {
		t.Fatalf("kept false UTF-16 text in %q", text)
	}
}

func TestLegacyDOCOldWordDecodesGBK(t *testing.T) {
	body := []byte{0xd6, 0xd0, 0xce, 0xc4, ' ', 0xbf, 0xc9, 0xbc, 0xfb, 0xce, 0xc4, 0xb1, 0xbe}
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "中文 可见文本") {
		t.Fatalf("missing decoded GBK text in %q", text)
	}
	for _, bad := range []string{"ÖÐÎÄ", "¿É¼û", "袩褉"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept GBK mojibake %q in %q", bad, text)
		}
	}
}

func TestLegacyDOCOldWordDecodesShiftJIS(t *testing.T) {
	body := []byte{0x93, 0xfa, 0x96, 0x7b, 0x8c, 0xea, ' ', 0x83, 0x65, 0x83, 0x4c, 0x83, 0x58, 0x83, 0x67}
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "\u65e5\u672c\u8a9e \u30c6\u30ad\u30b9\u30c8") {
		t.Fatalf("missing decoded Shift-JIS text in %q", text)
	}
	for _, bad := range []string{"“ú–{Œê", "擔杮岅", "\u0423\u042a"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept Shift-JIS mojibake %q in %q", bad, text)
		}
	}
}

func TestLegacyDOCOldWordDecodesEUCKR(t *testing.T) {
	body := []byte{0xc7, 0xd1, 0xb1, 0xdb, ' ', 0xc5, 0xd8, 0xbd, 0xba, 0xc6, 0xae}
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "\ud55c\uae00 \ud14d\uc2a4\ud2b8") {
		t.Fatalf("missing decoded EUC-KR text in %q", text)
	}
	for _, bad := range []string{"ÇÑ±Û", "ÅØ½ºÆ®", "\u0425\u041d"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept EUC-KR mojibake %q in %q", bad, text)
		}
	}
}

func TestLegacyDOCOldWordDecodesBig5(t *testing.T) {
	body := []byte{0xc1, 0x63, 0xc5, 0xe9, 0xa4, 0xa4, 0xa4, 0xe5, ' ', 0xb4, 0xfa, 0xb8, 0xd5, 0xa4, 0xe5, 0xa6, 0x72}
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "\u7e41\u9ad4\u4e2d\u6587 \u6e2c\u8a66\u6587\u5b57") {
		t.Fatalf("missing decoded Big5 text in %q", text)
	}
	for _, bad := range []string{"羸砰いゅ", "\uFFFD"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept Big5 mojibake %q in %q", bad, text)
		}
	}
}

func TestLegacyDOCOldWordKeepsReadableFootnoteAndCommentMarkers(t *testing.T) {
	body := []byte("Old body\x02 after footnote and old note\x05 after comment")
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	want := "Old body[footnote] after footnote and old note[comment] after comment"
	if !strings.Contains(text, want) {
		t.Fatalf("legacy old DOC text missing readable note markers %q in %q", want, text)
	}
	for _, hidden := range []string{"\x02", "\x05"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy old DOC text kept raw note marker %q in %q", hidden, text)
		}
	}
}

func TestLegacyDOCOldWordNormalizesLayoutControlChars(t *testing.T) {
	body := []byte("Old title\x0bnext line\x0cnew page\x07table cell")
	word := make([]byte, 128+len(body))
	binary.LittleEndian.PutUint16(word, 0xa5dc)
	binary.LittleEndian.PutUint32(word[0x18:], 128)
	binary.LittleEndian.PutUint32(word[0x1c:], uint32(128+len(body)))
	copy(word[128:], body)
	parts := extractLegacyText("sample.doc", nil, []oleStream{{Name: "WordDocument", Data: word}})
	text := strings.Join(parts, "\n")
	for _, want := range []string{"Old title", "next line", "new page", "table cell"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible old Word layout text %q in %q", want, text)
		}
	}
	for _, hidden := range []string{"\x07", "\x0b", "\x0c"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("kept raw Word layout control %q in %q", hidden, text)
		}
	}
}

func TestCompressedSingleByteTextDecodesLegacyEncodings(t *testing.T) {
	cyrillic := compressedUnicodeBytesToString([]byte{0xc2, 0xf0, 0xe5, 0xec, 0xff, ' ', 0xf0, 0xe0, 0xe1, 0xee, 0xf2, 0xfb})
	if cyrillic != "\u0412\u0440\u0435\u043c\u044f \u0440\u0430\u0431\u043e\u0442\u044b" {
		t.Fatalf("got %q", cyrillic)
	}
	utf8Text := compressedUnicodeBytesToString([]byte("Zeit\xc3\xbcberschneidung"))
	if utf8Text != "Zeit\u00fcberschneidung" {
		t.Fatalf("got %q", utf8Text)
	}
}

func TestLegacyDOCUsesPieceTableInsteadOfEmbeddedBinary(t *testing.T) {
	word, table := testWordPieceTableDocument("First Word paragraph", "第二段 Word 文本")
	copy(word[128:], []byte{0xff, 0xd8, 0xff, 0xe0})
	copy(word[134:], []byte("JFIF\x00binary jpeg payload should not appear"))
	parts := extractLegacyText("sample.doc", nil, []oleStream{
		{Name: "WordDocument", Data: word},
		{Name: "1Table", Data: table},
	})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "First Word paragraph") || !strings.Contains(text, "第二段 Word 文本") {
		t.Fatalf("missing piece table text in %q", text)
	}
	if strings.Contains(text, "JFIF") || strings.Contains(text, "binary jpeg payload") {
		t.Fatalf("kept embedded binary payload in %q", text)
	}
}

func TestLegacyDOCPieceTableStripsInlineHiddenOfficeReferences(t *testing.T) {
	word, table := testWordPieceTableDocument(
		`First target Target: ../media/piece.png Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image paragraph`,
		`Second type ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" PartName="/word/document.xml" r:embed="rId7" paragraph`,
	)
	streams := []oleStream{{Name: "WordDocument", Data: word}, {Name: "1Table", Data: table}}
	text := strings.Join(extractLegacyText("sample.doc", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.doc", nil, streams)
	for _, want := range []string{"First target paragraph", "Second type paragraph"} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy DOC piece-table text missing cleaned visible text %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy DOC piece-table markdown missing cleaned visible text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/piece.png", "Type:", "relationships/image", "ContentType", "application/vnd.openxmlformats", "PartName", "/word/document.xml", "r:embed", "rId7"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy DOC piece-table text kept hidden reference %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy DOC piece-table markdown kept hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyDOCPieceTableStripsFieldControlInstructions(t *testing.T) {
	word, table := testWordPieceTableDocument(
		"Before \x13 HYPERLINK \"https://example.test/internal\" \\h \x14 Visible Link \x15 After",
		"Unicode \x13 INCLUDEPICTURE \"word/media/hidden.png\" \\* MERGEFORMATINET \x14 Visible Caption \x15 Tail",
	)
	streams := []oleStream{{Name: "WordDocument", Data: word}, {Name: "1Table", Data: table}}
	text := strings.Join(extractLegacyText("sample.doc", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.doc", nil, streams)
	for _, want := range []string{"Before Visible Link After", "Unicode Visible Caption Tail"} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy DOC piece-table text missing field result %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy DOC piece-table markdown missing field result %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"HYPERLINK", "INCLUDEPICTURE", "example.test/internal", "word/media/hidden.png", "MERGEFORMATINET", "\x13", "\x14", "\x15"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy DOC piece-table text kept field/control content %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy DOC piece-table markdown kept field/control content %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyDOCPieceTableKeepsReadableFootnoteAndCommentMarkers(t *testing.T) {
	word, table := testWordPieceTableDocument(
		"Body\x02 after footnote and note\x05 after comment",
		"Unicode body\x02 footnote marker and unicode note\x05 comment marker",
	)
	streams := []oleStream{{Name: "WordDocument", Data: word}, {Name: "1Table", Data: table}}
	text := strings.Join(extractLegacyText("sample.doc", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.doc", nil, streams)
	for _, want := range []string{
		"Body[footnote] after footnote and note[comment] after comment",
		"Unicode body[footnote] footnote marker and unicode note[comment] comment marker",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy DOC piece-table text missing readable note marker %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy DOC piece-table markdown missing readable note marker %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"\x02", "\x05"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy DOC piece-table text kept raw note marker %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy DOC piece-table markdown kept raw note marker %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyDOCMarkdownStructuresStandaloneNotesAndComments(t *testing.T) {
	md := legacyWordMarkdown([]string{
		"[comment] Body start[footnote] keeps inline anchors",
		"[footnote] Visible footnote body",
		"[comment] Visible comment body",
		"[footnote] Visible endnote body",
	})
	for _, want := range []string{
		"## Document",
		"[comment] Body start[footnote] keeps inline anchors",
		"## Footnotes and Endnotes",
		"Visible footnote body",
		"Visible endnote body",
		"## Comments",
		"Visible comment body",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy DOC markdown missing structured note/comment content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "\x02") || strings.Contains(md, "\x05") {
		t.Fatalf("legacy DOC markdown kept raw note/comment controls:\n%s", md)
	}
	if strings.Index(md, "## Footnotes and Endnotes") < strings.Index(md, "## Document") ||
		strings.Index(md, "## Comments") < strings.Index(md, "## Footnotes and Endnotes") {
		t.Fatalf("legacy DOC markdown sections are not ordered document, notes, comments:\n%s", md)
	}
}

func TestLegacyDOCSamplesStructureVisibleNotesAndComments(t *testing.T) {
	tests := []struct {
		name string
		want []string
	}{
		{
			name: "footnote.doc",
			want: []string{"## Document", "Test text[footnote]", "## Footnotes and Endnotes", "TestFootnote", "TestEndnote", "## Comments", "TestComment"},
		},
		{
			name: "endingnote.doc",
			want: []string{"## Document", "Text, text[footnote] , text.", "## Footnotes and Endnotes", "Ending note text"},
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			md := res.Markdown("images")
			for _, want := range tc.want {
				if !strings.Contains(md, want) {
					t.Fatalf("%s markdown missing visible note/comment content %q in:\n%s", tc.name, want, md)
				}
			}
			if strings.Contains(md, "## Additional Text") {
				t.Fatalf("%s markdown should not duplicate structured notes/comments through backfill:\n%s", tc.name, md)
			}
			for _, hidden := range []string{"\x02", "\x05", "word/media/", "word/_rels/", "[Content_Types].xml"} {
				if strings.Contains(md, hidden) {
					t.Fatalf("%s markdown kept control/internal content %q in:\n%s", tc.name, hidden, md)
				}
			}
		})
	}
}

func TestLegacyDOCPieceTableNormalizesLayoutControlChars(t *testing.T) {
	word, table := testWordPieceTableDocument(
		"ASCII title\x0bASCII line\x0cASCII page\x07ASCII cell",
		"Unicode title\x0bUnicode line\x0cUnicode page\x07Unicode cell",
	)
	streams := []oleStream{{Name: "WordDocument", Data: word}, {Name: "1Table", Data: table}}
	text := strings.Join(extractLegacyText("sample.doc", nil, streams), "\n")
	md := extractLegacyMarkdown("sample.doc", nil, streams)
	for _, want := range []string{"ASCII title", "ASCII line", "ASCII page", "ASCII cell", "Unicode title", "Unicode line", "Unicode page", "Unicode cell"} {
		if !strings.Contains(text, want) {
			t.Fatalf("legacy DOC piece-table text missing layout text %q in %q", want, text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy DOC piece-table markdown missing layout text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"\x07", "\x0b", "\x0c"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("legacy DOC piece-table text kept raw layout control %q in %q", hidden, text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy DOC piece-table markdown kept raw layout control %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyDOCSampleDoesNotEmitEmbeddedJPEGBytes(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "multimedia.doc"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(res.Text, "JFIF") || strings.Contains(res.Text, "CDEFGHIJSTUVWXYZcdefghijstuvwxyz") {
		t.Fatalf("kept embedded JPEG bytes at start of extracted text: %.200q", res.Text)
	}
}

func TestLegacyDOCWord6SampleText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "57843.doc"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Dear Governor Bush") || !strings.Contains(res.Text, "Harold Koplow") {
		t.Fatalf("missing Word 6 body text in %.400q", res.Text)
	}
	if strings.Contains(res.Text, "MSWordDoc") || strings.Contains(res.Text, "Word.Document.6") {
		t.Fatalf("kept OLE class/control text in %.400q", res.Text)
	}
	if strings.Contains(res.Text, "son\ns idea") || !strings.Contains(res.Text, "son\u2019s idea") {
		t.Fatalf("did not preserve Windows-1252 apostrophe in %.400q", res.Text)
	}
}

func TestLegacyDOCCyrillicSampleText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "Bug50955.doc"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "ОПЕРАЦИЯ") || !strings.Contains(res.Text, "привет") {
		t.Fatalf("missing Cyrillic body text in %.400q", res.Text)
	}
	if strings.Contains(res.Text, "ÎÏÅÐÀÖÈß") || strings.Contains(res.Text, "ÏÐÈÂÅÒ") {
		t.Fatalf("kept Cyrillic mojibake in %.400q", res.Text)
	}
}

func TestLegacyDOCWord95CyrillicDoesNotEmitMojibake(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "word95err.doc"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "\u041a\u043e\u043c\u043f\u0430\u043d\u0438\u044f \u0410\u0441\u043a\u0430\u0442") {
		t.Fatalf("missing decoded Cyrillic text in %.400q", res.Text)
	}
	for _, want := range []string{"\u041e\u0431\u044a\u0435\u043c", "\u043c\u043b", "\u0440\u0443\u0431"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing decoded short Cyrillic text %q in %.400q", want, res.Text)
		}
	}
	for _, bad := range []string{"\u00c2\u00f0\u00e5\u00ec\u00ff", "\uff8e\u78c5\u890c", "\u00ec\u00eb", "\u00f0\u00f3\u00e1"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept Windows-1251 mojibake %q in %.400q", bad, res.Text)
		}
	}
}

func TestEmbeddedNonOfficeOLEIsNotExtractedAsText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "58325_lt.xlsx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, noise := range []string{"%PDF-1.4", "AcroExch.Document.DC", "Acrobat Document"} {
		if strings.Contains(res.Text, noise) {
			t.Fatalf("kept embedded non-Office OLE text %q in %.400q", noise, res.Text)
		}
	}
}

func TestLegacyFallbackSkipsOLEObjectWrapperStreams(t *testing.T) {
	parts := extractLegacyText("unknown.bin", nil, []oleStream{
		{Name: "\x01CompObj", Path: "\x01CompObj", Data: []byte("AcroExch.Document.DC Acrobat Document")},
		{Name: "\x03ObjInfo", Path: "ObjectPool/_123/\x03ObjInfo", Data: []byte("Package Object internal flags")},
		{Name: "\x01Ole10Native", Path: "ObjectPool/_123/\x01Ole10Native", Data: []byte(`C:\Users\me\hidden.pdf native wrapper label`)},
		{Name: "\x02OlePres000", Path: "\x02OlePres000", Data: []byte("preview wrapper bitmap label")},
		{Name: "Contents", Path: "Package/Contents", Data: append([]byte("%PDF-1.4 binary object "), testPNG()...)},
		{Name: "VisibleText", Path: "VisibleText", Data: []byte("Visible fallback paragraph")},
	})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "Visible fallback paragraph") {
		t.Fatalf("missing visible fallback text in %q", text)
	}
	for _, bad := range []string{"AcroExch", "Acrobat Document", "Package Object", "hidden.pdf", "native wrapper", "preview wrapper", "%PDF", "binary object"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept OLE wrapper text %q in %q", bad, text)
		}
	}
}

func TestLegacyFallbackSkipsInternalVBAAndEncryptionStreams(t *testing.T) {
	parts := extractLegacyText("unknown.bin", nil, []oleStream{
		{Name: "dir", Path: "VBA/dir", Data: []byte("Attribute VB_Name = \"HiddenMacro\"")},
		{Name: "Module1", Path: "VBA/Module1", Data: []byte("Sub HiddenMacro()\nMsgBox \"Hidden macro code\"\nEnd Sub")},
		{Name: "PROJECT", Path: "Macros/PROJECT", Data: []byte("Document=ThisDocument/&H00000000")},
		{Name: "EncryptionInfo", Path: "EncryptionInfo", Data: []byte("Microsoft.Container.DataSpaces StrongEncryptionData")},
		{Name: "DataSpaceInfo", Path: "DataSpaces/DataSpaceInfo/StrongEncryptionData", Data: []byte("StrongEncryptionData TransformInfo")},
		{Name: "VisibleText", Path: "VisibleText", Data: []byte("Visible fallback paragraph")},
	})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, "Visible fallback paragraph") {
		t.Fatalf("missing visible fallback text in %q", text)
	}
	for _, bad := range []string{"HiddenMacro", "Hidden macro code", "Attribute VB_Name", "Document=ThisDocument", "Microsoft.Container.DataSpaces", "StrongEncryptionData", "TransformInfo"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept internal legacy stream text %q in %q", bad, text)
		}
	}
}

func TestEmbeddedNonOfficeOLEImagesAreRecoveredWithoutText(t *testing.T) {
	png := testPNG()
	streams := []oleStream{
		{
			Name: "Contents",
			Path: "Package/Contents",
			Data: append(append([]byte("AcroExch.Document.DC binary object text "), png...), []byte(" trailing object bytes")...),
		},
		{
			Name: "vector.svgz",
			Path: "Package/vector.svgz",
			Data: gzipBytes(t, testSVG()),
		},
		{
			Name: "WordDocument",
			Path: "NestedOffice/WordDocument",
			Data: png,
		},
	}
	if images := nonOfficeOLEImagesFromStreams("ppt/embeddings/officeObject.bin", nil, streams, 0); len(images) != 0 {
		t.Fatalf("expected Office-like OLE streams to be ignored by non-Office image recovery, got %#v", images)
	}

	streams = streams[:2]
	images := nonOfficeOLEImagesFromStreams(`ppt\embeddings\oleObject1.bin`, nil, streams, 0)
	if len(images) != 2 {
		t.Fatalf("expected two non-Office OLE images, got %#v", images)
	}
	if images[0].Name != "oleObject1.bin-Contents.png" || images[0].Ext != ".png" || !bytes.Equal(images[0].Data, png) {
		t.Fatalf("expected carved embedded PNG, got %#v", images[0])
	}
	if images[1].Name != "oleObject1.bin-vector.svg" || images[1].Ext != ".svg" || !validImageData(".svg", images[1].Data) {
		t.Fatalf("expected decompressed embedded SVG, got %#v", images[1])
	}
	for _, img := range images {
		text := strings.Join(extractBinaryStrings(img.Data), "\n")
		if strings.Contains(text, "AcroExch") || strings.Contains(text, "binary object text") {
			t.Fatalf("embedded object control text leaked through image data: %q", text)
		}
	}
	md := (&Result{Images: images}).Markdown("images")
	if strings.Contains(md, `ppt`) || strings.Contains(md, `embeddings`) || strings.Contains(md, `\`) {
		t.Fatalf("markdown kept embedded OLE object path components:\n%s", md)
	}
}

func TestEmbeddedOle10NativeImageIsRecoveredWithoutWrapperMetadata(t *testing.T) {
	png := testPNG()
	streams := []oleStream{
		{
			Name: "\x01Ole10Native",
			Path: "ObjectPool/_123/\x01Ole10Native",
			Data: append(append([]byte("C:\\Users\\me\\Pictures\\hidden-preview.png\x00Ole10Native wrapper label\x00"), png...), []byte("\x00trailing native bytes")...),
		},
		{
			Name: "\x03ObjInfo",
			Path: "ObjectPool/_123/\x03ObjInfo",
			Data: []byte("Package Object internal flags"),
		},
	}
	images := nonOfficeOLEImagesFromStreams(`word\embeddings\oleObject2.bin`, nil, streams, 0)
	if len(images) != 1 {
		t.Fatalf("expected one image recovered from Ole10Native wrapper, got %#v", images)
	}
	if images[0].Ext != ".png" || !validImageData(".png", images[0].Data) {
		t.Fatalf("expected valid PNG from Ole10Native wrapper, got %#v", images[0])
	}
	dir := t.TempDir()
	if err := writeImages(dir, images); err != nil {
		t.Fatal(err)
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 {
		t.Fatalf("expected one written Ole10Native image, got %#v", entries)
	}
	writtenName := entries[0].Name()
	written, err := os.ReadFile(filepath.Join(dir, writtenName))
	if err != nil {
		t.Fatal(err)
	}
	if !validImageData(filepath.Ext(writtenName), written) {
		t.Fatalf("written Ole10Native image is invalid: %s len=%d", writtenName, len(written))
	}
	md := (&Result{Images: images}).Markdown("images")
	for _, hidden := range []string{"Ole10Native", "ObjInfo", "ObjectPool", "hidden-preview", "Users", "Pictures", "wrapper label", "Package Object"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown leaked Ole10Native wrapper metadata %q in:\n%s", hidden, md)
		}
		if strings.Contains(writtenName, hidden) {
			t.Fatalf("written Ole10Native image filename leaked wrapper metadata %q: %s", hidden, writtenName)
		}
	}
	for _, img := range images {
		text := strings.Join(extractBinaryStrings(img.Data), "\n")
		for _, hidden := range []string{"Ole10Native", "wrapper label", "hidden-preview", "Package Object"} {
			if strings.Contains(text, hidden) {
				t.Fatalf("image data leaked Ole10Native wrapper metadata %q in %q", hidden, text)
			}
		}
	}
}

func TestOLEPackageStreamContainingOOXMLIsExtractedAsOfficeDocument(t *testing.T) {
	docx := testDocxPackage(t, "Visible packaged OOXML text", testPNG())
	streams := []oleStream{
		{Name: "Package", Path: "ObjectPool/_123/Package", Data: docx},
		{Name: "\x01Ole10Native", Path: "ObjectPool/_123/\x01Ole10Native", Data: append([]byte("native wrapper label"), docx...)},
	}
	res := extractOfficePackagesFromOLEStreams("embedded-object.bin", streams, 1, Options{})
	if res == nil {
		t.Fatal("expected embedded OOXML package to be extracted")
	}
	if !strings.Contains(res.Text, "Visible packaged OOXML text") {
		t.Fatalf("missing embedded OOXML text in %.400q", res.Text)
	}
	if strings.Contains(res.Text, "native wrapper label") {
		t.Fatalf("kept Ole10Native wrapper text in %.400q", res.Text)
	}
	if len(res.Images) != 1 || res.Images[0].Ext != ".png" || !validImageData(res.Images[0].Ext, res.Images[0].Data) {
		t.Fatalf("expected one valid embedded image, got %#v", res.Images)
	}
	if !strings.Contains(res.StructuredMarkdown, "Visible packaged OOXML text") {
		t.Fatalf("missing embedded OOXML markdown in %.400q", res.StructuredMarkdown)
	}
}

func TestOLEPackageStreamEmbeddedImageNamesAreUnique(t *testing.T) {
	docx1 := testDocxPackage(t, "First embedded package text", testPNG())
	docx2 := testDocxPackage(t, "Second embedded package text", testPNG())
	res := extractOfficePackagesFromOLEStreams("embedded-object.bin", []oleStream{
		{Name: "Package", Path: "ObjectPool/_123/Package", Data: docx1},
		{Name: "Package", Path: `ObjectPool\_456\Package`, Data: docx2},
	}, 1, Options{})
	if res == nil {
		t.Fatal("expected embedded OOXML packages to be extracted")
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected two embedded images, got %#v", res.Images)
	}
	names := map[string]bool{}
	for _, img := range res.Images {
		if img.Name == "" {
			t.Fatalf("expected embedded image name to be populated: %#v", img)
		}
		if names[strings.ToLower(img.Name)] {
			t.Fatalf("duplicate embedded image name %q in %#v", img.Name, res.Images)
		}
		names[strings.ToLower(img.Name)] = true
		if !validImageData(img.Ext, img.Data) {
			t.Fatalf("expected valid embedded image data for %#v", img)
		}
	}
	if !names["package.docx-image1.png"] || !names["package.docx-image1-2.png"] {
		t.Fatalf("expected stable unique embedded image names, got %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "](images/Package.docx-image1.png)") ||
		!strings.Contains(md, "](images/Package.docx-image1-2.png)") {
		t.Fatalf("markdown did not reference unique embedded image names:\n%s", md)
	}
	if strings.Contains(md, "](images/image1.png)") {
		t.Fatalf("markdown kept stale embedded package image reference:\n%s", md)
	}
	for _, hidden := range []string{"ObjectPool", "_456", `\`} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept embedded OLE package path component %q in:\n%s", hidden, md)
		}
	}
}

func TestOLEPackageStreamEmbeddedImageIsPlacedInMarkdown(t *testing.T) {
	var inner bytes.Buffer
	innerZip := zip.NewWriter(&inner)
	addZip(t, innerZip, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, innerZip, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>OLE package before image</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="OLE package visible picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicture"/></p:blipFill></p:pic>
<w:p><w:r><w:t>OLE package after image</w:t></w:r></w:p>
</w:body></w:document>`)
	addZip(t, innerZip, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline.png"/></Relationships>`)
	addZipBytes(t, innerZip, "word/media/inline.png", testPNG())
	if err := innerZip.Close(); err != nil {
		t.Fatal(err)
	}

	res := extractOfficePackagesFromOLEStreams("embedded-object.bin", []oleStream{
		{Name: "Package", Path: "ObjectPool/_123/Package", Data: inner.Bytes()},
	}, 1, Options{})
	if res == nil {
		t.Fatal("expected embedded OOXML package to be extracted")
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Package.docx-inline.png" || res.Images[0].Alt != "OLE package visible picture" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected renamed valid OLE package image with alt, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"OLE package before image",
		"OLE package visible picture\n![OLE package visible picture](images/Package.docx-inline.png)",
		"OLE package after image",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing placed OLE package image content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "](images/inline.png)") || strings.Contains(md, "## Images") || strings.Contains(md, "ObjectPool") {
		t.Fatalf("markdown kept stale, duplicate, or internal OLE package image reference:\n%s", md)
	}
}

func TestOLEPackageStreamTrimsOOXMLPayloadWrapperBytes(t *testing.T) {
	docx := testDocxPackage(t, "Visible wrapped OOXML text", testPNG())
	wrapped := append([]byte("native prefix"), docx...)
	wrapped = append(wrapped, bytes.Repeat([]byte(" native suffix "), 6000)...)
	payloads := embeddedOfficePackagePayloadsFromStreams([]oleStream{
		{Name: "\x01Ole10Native", Path: "ObjectPool/_123/\x01Ole10Native", Data: wrapped},
	})
	if len(payloads) != 1 {
		t.Fatalf("expected one trimmed OOXML payload, got %#v", payloads)
	}
	if !bytes.Equal(payloads[0].data, docx) {
		t.Fatalf("expected wrapper bytes to be trimmed: got %d bytes, want %d", len(payloads[0].data), len(docx))
	}
	res := extractOfficePackagesFromOLEStreams("embedded-object.bin", []oleStream{
		{Name: "\x01Ole10Native", Path: "ObjectPool/_123/\x01Ole10Native", Data: wrapped},
	}, 1, Options{})
	if res == nil || !strings.Contains(res.Text, "Visible wrapped OOXML text") {
		t.Fatalf("missing wrapped OOXML text in %#v", res)
	}
	if strings.Contains(res.Text, "native prefix") || strings.Contains(res.Text, "native suffix") {
		t.Fatalf("kept wrapper text in %.400q", res.Text)
	}
	if len(res.Images) != 1 || !validImageData(res.Images[0].Ext, res.Images[0].Data) {
		t.Fatalf("expected one valid image from trimmed OOXML package, got %#v", res.Images)
	}
}

func TestNonOfficeOLEImageRecoverySkipsOfficePackageStreams(t *testing.T) {
	docx := testDocxPackage(t, "Visible packaged OOXML text", testPNG())
	images := nonOfficeOLEImagesFromStreams("ppt/embeddings/oleObject1.bin", nil, []oleStream{
		{Name: "Package", Path: "ObjectPool/_123/Package", Data: docx},
	}, 0)
	if len(images) != 0 {
		t.Fatalf("expected Office package stream to bypass non-Office image recovery, got %#v", images)
	}
}

func TestLegacyXLSUsesBoundSheetInsteadOfBinaryFallback(t *testing.T) {
	var workbook bytes.Buffer
	name := "Visible Sheet"
	rec := make([]byte, 8+len(name))
	rec[6] = byte(len(name))
	rec[7] = 0
	copy(rec[8:], name)
	writeBIFFRecord(&workbook, 0x0085, rec)
	workbook.WriteString("%PDF-1.4\nAcroExch.Document.11\nbinary embedded object should not appear")
	parts := extractLegacyText("sample.xls", nil, []oleStream{{Name: "Workbook", Data: workbook.Bytes()}})
	text := strings.Join(parts, "\n")
	if !strings.Contains(text, name) {
		t.Fatalf("missing bound sheet text in %q", text)
	}
	if strings.Contains(text, "%PDF") || strings.Contains(text, "AcroExch") || strings.Contains(text, "binary embedded object") {
		t.Fatalf("kept embedded object text in %q", text)
	}
}

func TestLegacyXLSMarkdownUsesBIFFCellGrid(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	writeBIFFRecord(&workbook, 0x0015, testXLUnicodeString("Workbook global footer should not appear"))
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	hiddenName := "Hidden Sheet"
	hiddenBound := make([]byte, 8+len(hiddenName))
	hiddenBound[4] = 1
	hiddenBound[6] = byte(len(hiddenName))
	hiddenBound[7] = 0
	copy(hiddenBound[8:], hiddenName)
	writeBIFFRecord(&workbook, 0x0085, hiddenBound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 15)
	binary.LittleEndian.PutUint32(sst[4:], 15)
	sst = append(sst, testXLUnicodeString("Name")...)
	sst = append(sst, testXLUnicodeString("Score")...)
	sst = append(sst, testXLUnicodeString("Alice")...)
	sst = append(sst, testXLUnicodeString("Bob")...)
	sst = append(sst, testXLUnicodeString("Hidden internal value")...)
	sst = append(sst, testXLUnicodeString("Hidden row secret")...)
	sst = append(sst, testXLUnicodeString("Hidden column secret")...)
	sst = append(sst, testXLUnicodeString("Ready")...)
	sst = append(sst, testXLUnicodeString("Totals")...)
	sst = append(sst, testXLUnicodeString("Errors")...)
	sst = append(sst, testXLUnicodeString("Formula")...)
	sst = append(sst, testXLUnicodeString("Rich")...)
	sst = append(sst, testXLUnicodeString("FormulaBool")...)
	sst = append(sst, testXLUnicodeString("FormulaErr")...)
	sst = append(sst, testXLUnicodeString("FormulaText")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	hiddenCol := make([]byte, 12)
	binary.LittleEndian.PutUint16(hiddenCol[0:], 2)
	binary.LittleEndian.PutUint16(hiddenCol[2:], 2)
	binary.LittleEndian.PutUint16(hiddenCol[8:], 0x0001)
	writeBIFFRecord(&workbook, 0x007d, hiddenCol)
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 1, 1))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 2, 6))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(1, 0, 2))
	writeBIFFRecord(&workbook, 0x0203, testNumberRecord(1, 1, 93.5))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(2, 0, 3))
	writeBIFFRecord(&workbook, 0x027e, testRKRecord(2, 1, 87))
	hiddenRow := make([]byte, 16)
	binary.LittleEndian.PutUint16(hiddenRow[0:], 3)
	binary.LittleEndian.PutUint16(hiddenRow[12:], 0x0020)
	writeBIFFRecord(&workbook, 0x0208, hiddenRow)
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(3, 0, 5))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(4, 0, 7))
	writeBIFFRecord(&workbook, 0x0205, testBoolErrRecord(4, 1, 1, false))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(5, 0, 8))
	writeBIFFRecord(&workbook, 0x00bd, testMulRKRecord(5, 1, []int32{12, 34}))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(6, 0, 9))
	writeBIFFRecord(&workbook, 0x0205, testBoolErrRecord(6, 1, 0x07, true))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(7, 0, 10))
	writeBIFFRecord(&workbook, 0x0006, testFormulaNumberRecord(7, 1, 42.75))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(8, 0, 12))
	writeBIFFRecord(&workbook, 0x0006, testFormulaSpecialRecord(8, 1, 0x01, 1))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(9, 0, 13))
	writeBIFFRecord(&workbook, 0x0006, testFormulaSpecialRecord(9, 1, 0x02, 0x2a))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(10, 0, 14))
	writeBIFFRecord(&workbook, 0x0006, testFormulaSpecialRecord(10, 1, 0x00, 0))
	writeBIFFRecord(&workbook, 0x0207, testXLUnicodeString("Cached formula text"))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(11, 0, 11))
	writeBIFFRecord(&workbook, 0x00d6, testRStringRecord(11, 1, "Formatted visible text"))
	writeBIFFRecord(&workbook, 0x0015, testXLUnicodeString("Visible footer note"))
	writeBIFFRecord(&workbook, 0x000a, nil)
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 4))
	writeBIFFRecord(&workbook, 0x000a, nil)

	dir := t.TempDir()
	file := filepath.Join(dir, "legacy-grid.xls")
	if err := os.WriteFile(file, workbook.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"## Visible Sheet",
		"| Name | Score |",
		"| --- | --- |",
		"| Alice | 93.5 |",
		"| Bob | 87 |",
		"| Ready | TRUE |",
		"| Totals | 12 |",
		"| Errors | #DIV/0! |",
		"| Formula | 42.75 |",
		"| FormulaBool | TRUE |",
		"| FormulaErr | #N/A |",
		"| FormulaText | Cached formula text |",
		"| Rich | Formatted visible text |",
		"### Headers and Footers",
		"Visible footer note",
		"## Workbook Sheets",
		"- Visible Sheet",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing %q in:\n%s", want, md)
		}
	}
	for _, want := range []string{"Name", "Score", "Alice", "93.5", "Bob", "87", "Ready", "TRUE", "Totals", "12", "Errors", "#DIV/0!", "Formula", "42.75", "FormulaBool", "FormulaErr", "#N/A", "FormulaText", "Cached formula text", "Rich", "Formatted visible text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("legacy XLS text missing %q in:\n%s", want, res.Text)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("legacy XLS markdown should structure footer text instead of appending it as additional text:\n%s", md)
	}
	for _, bad := range []string{"\x00", "BIFF", "Hidden Sheet", "Hidden internal value", "Hidden row secret", "Hidden column secret", "| Totals | 12 | 34 |", "Workbook global footer should not appear"} {
		if strings.Contains(md, bad) {
			t.Fatalf("legacy XLS markdown kept internal/control text %q in:\n%s", bad, md)
		}
		if strings.Contains(res.Text, bad) {
			t.Fatalf("legacy XLS text kept internal/control text %q in:\n%s", bad, res.Text)
		}
	}
}

func TestLegacyXLSSampleSheetNamesAreStructured(t *testing.T) {
	tests := []struct {
		name string
		want []string
	}{
		{name: "45538_classic_Header.xls", want: []string{"## Workbook Text", "Top Six Functions - Intern", "Top Five Industries - Intern"}},
		{name: "12843-1.xls", want: []string{"## Workbook Sheets", "SDH-line", "ADSL-表", "MARI-圖"}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			md := res.Markdown("images")
			for _, want := range tc.want {
				if !strings.Contains(md, want) {
					t.Fatalf("legacy XLS markdown missing sheet name %q in %.1200q", want, md)
				}
			}
			if strings.Contains(md, "## Additional Text") {
				t.Fatalf("legacy XLS sheet names should be structured instead of backfilled:\n%s", md)
			}
		})
	}
}

func TestLegacyXLSDoesNotTreatPrinterSettingsAsSheetNames(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "42844.xls"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, bad := range []string{"HPJOBACCT_JOBACNT", "HP LaserJet", "EXCEL.EXE", "InputBin", "FORMSOURCE", "PrnStat_SID"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("legacy XLS text kept printer/internal setting %q in %.800q", bad, res.Text)
		}
		if strings.Contains(md, bad) {
			t.Fatalf("legacy XLS markdown kept printer/internal setting %q in %.800q", bad, md)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("legacy XLS printer/internal settings should not be backfilled:\n%s", md)
	}
}

func TestLegacyXLSTextStripsInlineHiddenOfficeReferences(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 4)
	binary.LittleEndian.PutUint32(sst[4:], 4)
	sst = append(sst, testXLUnicodeString("Visible before word/media/image1.png visible after")...)
	sst = append(sst, testXLUnicodeString(`Visible type ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" after`)...)
	sst = append(sst, testXLUnicodeString(`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`)...)
	sst = append(sst, testXLUnicodeString(`ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"`)...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(1, 0, 1))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(2, 0, 2))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(3, 0, 3))
	writeBIFFRecord(&workbook, 0x0204, testBIFFLabelRecord(2, 0, `Visible target Target="../media/inline.png" after`))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("inline-hidden.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible before visible after", "Visible type after", "Visible target after"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("legacy XLS text missing cleaned visible text %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing cleaned visible text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"word/media/image1.png", "ContentType", "application/vnd.openxmlformats", `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"`, `Target="../media/inline.png"`, "../media/inline.png"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("legacy XLS text kept hidden reference %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy XLS markdown kept hidden reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyXLSIgnoresOrphanFormulaStringRecords(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 2)
	binary.LittleEndian.PutUint32(sst[4:], 2)
	sst = append(sst, testXLUnicodeString("Visible Cell")...)
	sst = append(sst, testXLUnicodeString("FormulaText")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x0207, testXLUnicodeString("Orphan internal STRING record"))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(1, 0, 1))
	writeBIFFRecord(&workbook, 0x0006, testFormulaStringRecord(1, 1))
	writeBIFFRecord(&workbook, 0x0207, testXLUnicodeString("Visible formula cache"))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("orphan-string.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Cell", "FormulaText", "Visible formula cache"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("legacy XLS text missing %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing %q in:\n%s", want, md)
		}
	}
	if strings.Contains(res.Text, "Orphan internal STRING record") || strings.Contains(md, "Orphan internal STRING record") {
		t.Fatalf("orphan BIFF STRING record leaked: text=%q markdown=\n%s", res.Text, md)
	}
}

func TestLegacyXLSRichExtendedSSTDoesNotLeakFormattingBytes(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 2)
	binary.LittleEndian.PutUint32(sst[4:], 2)
	sst = append(sst, testXLUnicodeRichExtString("Visible rich cell", []byte("Target: ../media/hidden.png Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))...)
	sst = append(sst, testXLUnicodeString("Following clean cell")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 1, 1))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("rich-ext-sst.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible rich cell", "Following clean cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("legacy XLS text missing rich/ext SST content %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing rich/ext SST content %q in:\n%s", want, md)
		}
	}
	if !strings.Contains(md, "| Visible rich cell | Following clean cell |") {
		t.Fatalf("legacy XLS markdown missing rich/ext SST table row in:\n%s", md)
	}
	for _, hidden := range []string{"Target:", "../media/hidden.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("legacy XLS text kept rich/ext hidden bytes %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy XLS markdown kept rich/ext hidden bytes %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyXLSMarkdownStripsHiddenReferencesFromSheetNames(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	for _, name := range []string{
		"word/media/sheet-hidden.png",
		"Visible target Target: ../media/sheet.png",
	} {
		bound := make([]byte, 8+len(name))
		bound[6] = byte(len(name))
		bound[7] = 0
		copy(bound[8:], name)
		writeBIFFRecord(&workbook, 0x0085, bound)
	}
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 2)
	binary.LittleEndian.PutUint32(sst[4:], 2)
	sst = append(sst, testXLUnicodeString("First visible cell")...)
	sst = append(sst, testXLUnicodeString("Second visible cell")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x000a, nil)
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 1))
	writeBIFFRecord(&workbook, 0x000a, nil)

	md := extractLegacyMarkdown("sample.xls", workbook.Bytes(), nil)
	for _, want := range []string{"## Sheet 1", "First visible cell", "## Visible target", "Second visible cell"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing cleaned sheet content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"word/media/sheet-hidden.png", "Target:", "../media/sheet.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy XLS markdown kept hidden sheet-name reference %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyXLSLateHiddenRowDoesNotLeakText(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 3)
	binary.LittleEndian.PutUint32(sst[4:], 3)
	sst = append(sst, testXLUnicodeString("Visible Cell")...)
	sst = append(sst, testXLUnicodeString("Late Hidden Secret")...)
	sst = append(sst, testXLUnicodeString("Visible Tail")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(1, 0, 1))
	writeBIFFRecord(&workbook, 0x0006, testFormulaStringRecord(1, 1))
	hiddenRow := make([]byte, 16)
	binary.LittleEndian.PutUint16(hiddenRow[0:], 1)
	binary.LittleEndian.PutUint16(hiddenRow[12:], 0x0020)
	writeBIFFRecord(&workbook, 0x0208, hiddenRow)
	writeBIFFRecord(&workbook, 0x0207, testXLUnicodeString("Late Hidden Formula String"))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(2, 0, 2))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("late-hidden.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Cell", "Visible Tail"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible BIFF text %q in %q", want, res.Text)
		}
		if !strings.Contains(res.Markdown("images"), want) {
			t.Fatalf("missing visible BIFF markdown %q in:\n%s", want, res.Markdown("images"))
		}
	}
	for _, hidden := range []string{"Late Hidden Secret", "Late Hidden Formula String"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(res.Markdown("images"), hidden) {
			t.Fatalf("late hidden row leaked %q: text=%q markdown=\n%s", hidden, res.Text, res.Markdown("images"))
		}
	}
}

func TestLegacyXLSLateHiddenColumnDoesNotLeakTextOrComments(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 4)
	binary.LittleEndian.PutUint32(sst[4:], 4)
	sst = append(sst, testXLUnicodeString("Visible Cell")...)
	sst = append(sst, testXLUnicodeString("Late Hidden Column Cell")...)
	sst = append(sst, testXLUnicodeString("Visible Tail")...)
	sst = append(sst, testXLUnicodeString("Late Hidden MulRK Neighbor")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 1, 1))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(1, 0, 2))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(1, 1, 3))
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 1))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord("Late hidden column comment"))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue("Late hidden column comment"))
	writeBIFFRecord(&workbook, 0x0006, testFormulaStringRecord(2, 1))
	hiddenCol := make([]byte, 12)
	binary.LittleEndian.PutUint16(hiddenCol[0:], 1)
	binary.LittleEndian.PutUint16(hiddenCol[2:], 1)
	binary.LittleEndian.PutUint16(hiddenCol[8:], 0x0001)
	writeBIFFRecord(&workbook, 0x007d, hiddenCol)
	writeBIFFRecord(&workbook, 0x0207, testXLUnicodeString("Late hidden formula string"))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("late-hidden-column.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Cell", "Visible Tail"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible BIFF text %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("missing visible BIFF markdown %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Late Hidden Column Cell", "Late Hidden MulRK Neighbor", "Late hidden column comment", "Late hidden formula string", "#### B1"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("late hidden column leaked into text %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("late hidden column leaked into markdown %q in:\n%s", hidden, md)
		}
	}
}

func TestLegacyXLSCommentsAreVisibleAndHiddenCellsAreFiltered(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	hiddenName := "Hidden Sheet"
	hiddenBound := make([]byte, 8+len(hiddenName))
	hiddenBound[4] = 1
	hiddenBound[6] = byte(len(hiddenName))
	hiddenBound[7] = 0
	copy(hiddenBound[8:], hiddenName)
	writeBIFFRecord(&workbook, 0x0085, hiddenBound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 1)
	binary.LittleEndian.PutUint32(sst[4:], 1)
	sst = append(sst, testXLUnicodeString("Visible Cell")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	hiddenCol := make([]byte, 12)
	binary.LittleEndian.PutUint16(hiddenCol[0:], 1)
	binary.LittleEndian.PutUint16(hiddenCol[2:], 1)
	binary.LittleEndian.PutUint16(hiddenCol[8:], 0x0001)
	writeBIFFRecord(&workbook, 0x007d, hiddenCol)
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 0))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord("Visible legacy comment"))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue("Visible legacy comment"))
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 2))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecordWithLen(len([]rune("分段批注正文"))))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOUnicodeContinue("分段批注"))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOUnicodeContinue("正文"))
	mixedComment := "Visible mixed comment\n=Hidden!A1\nHidden!B2\nSUM(A1:A2)\nword/media/comment.png"
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 3))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord(mixedComment))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue(mixedComment))
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 1))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord("Hidden column comment secret"))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue("Hidden column comment secret"))
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(1, 0))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord("Hidden row comment secret"))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue("Hidden row comment secret"))
	hiddenRow := make([]byte, 16)
	binary.LittleEndian.PutUint16(hiddenRow[0:], 1)
	binary.LittleEndian.PutUint16(hiddenRow[12:], 0x0020)
	writeBIFFRecord(&workbook, 0x0208, hiddenRow)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 0))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord("Hidden sheet comment secret"))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue("Hidden sheet comment secret"))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("comments.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible legacy comment") {
		t.Fatalf("missing visible legacy XLS comment in text: %q", res.Text)
	}
	if !strings.Contains(res.Text, "分段批注正文") {
		t.Fatalf("missing split Unicode legacy XLS comment in text: %q", res.Text)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "### Comments", "#### A1", "Visible legacy comment", "#### C1", "分段批注正文"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing comment content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "#### B1", "#### A2", "Hidden column comment secret", "Hidden row comment secret", "Hidden sheet comment secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("legacy XLS text kept hidden comment content %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("legacy XLS markdown kept hidden comment content %q in:\n%s", hidden, md)
		}
	}
	if !strings.Contains(res.Text, "Visible mixed comment") || !strings.Contains(md, "#### D1") || !strings.Contains(md, "Visible mixed comment") {
		t.Fatalf("missing cleaned mixed legacy XLS comment: text=%q markdown=\n%s", res.Text, md)
	}
	for _, hidden := range []string{"=Hidden!A1", "Hidden!B2", "SUM(A1:A2)", "word/media/comment.png"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("legacy XLS mixed comment kept hidden/formula line %q: text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestLegacyXLSMarkdownKeepsCommentOnlySheet(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Comment Only"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	writeBIFFRecord(&workbook, 0x000a, nil)

	comment := `Visible comment only Target: ../media/comment.png text`
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x001c, testBIFFNoteRecord(0, 0))
	writeBIFFRecord(&workbook, 0x01b6, testBIFFTXORecord(comment))
	writeBIFFRecord(&workbook, 0x003c, testBIFFTXOContinue(comment))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("comment-only.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible comment only text") {
		t.Fatalf("missing visible comment-only legacy XLS text: %q", res.Text)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Comment Only", "### Comments", "#### A1", "Visible comment only text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing comment-only content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/comment.png"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(md, hidden) {
			t.Fatalf("legacy XLS comment-only output kept hidden reference %q: text=%q markdown=\n%s", hidden, res.Text, md)
		}
	}
}

func TestLegacyXLSOldBIFFDoesNotEmitUnicodeMojibake(t *testing.T) {
	cases := map[string]string{
		"testEXCEL_4.xls":  "Examination Coverage",
		"testEXCEL_5.xls":  "Sample Excel Worksheet - Numbers and their Squares",
		"testEXCEL_95.xls": "Sample Excel Worksheet - Numbers and their Squares",
	}
	for sample, want := range cases {
		t.Run(sample, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", sample), Options{})
			if err != nil {
				t.Fatal(err)
			}
			if !strings.Contains(res.Text, want) {
				t.Fatalf("missing visible worksheet text in %.400q", res.Text)
			}
			for _, noise := range []string{"\u81a7", "\u92f9", "\u5577"} {
				if strings.Contains(res.Text, noise) {
					t.Fatalf("kept old BIFF binary mojibake %q in %.400q", noise, res.Text)
				}
			}
		})
	}
}

func TestLegacyXLSBIFF8CompressedUnicodeKeepsLatin1Characters(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "TestUnicode.xls"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	want := "Der große Überschriften Test"
	if !strings.Contains(res.Text, want) {
		t.Fatalf("missing BIFF8 compressed Unicode Latin-1 text %q in %q", want, res.Text)
	}
	if !strings.Contains(md, want) {
		t.Fatalf("markdown missing BIFF8 compressed Unicode Latin-1 text %q in:\n%s", want, md)
	}
	for _, bad := range []string{"gro遝", "躡erschriften"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("legacy XLS text kept BIFF8 compressed Unicode mojibake %q in %q", bad, res.Text)
		}
		if strings.Contains(md, bad) {
			t.Fatalf("legacy XLS markdown kept BIFF8 compressed Unicode mojibake %q in:\n%s", bad, md)
		}
	}
	for _, want := range []string{"## Tabelle1", "## Tabelle2", "## Tabelle3"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing visible sheet section %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("legacy XLS markdown should structure empty visible sheets instead of backfilling them:\n%s", md)
	}
}

func TestLegacyXLSMarkdownKeepsEmptyVisibleSheetsAndSkipsHiddenEmptySheets(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	visibleName := "Visible Empty"
	visibleBound := make([]byte, 8+len(visibleName))
	visibleBound[6] = byte(len(visibleName))
	visibleBound[7] = 0
	copy(visibleBound[8:], visibleName)
	writeBIFFRecord(&workbook, 0x0085, visibleBound)
	hiddenName := "Hidden Empty"
	hiddenBound := make([]byte, 8+len(hiddenName))
	hiddenBound[4] = 1
	hiddenBound[6] = byte(len(hiddenName))
	hiddenBound[7] = 0
	copy(hiddenBound[8:], hiddenName)
	writeBIFFRecord(&workbook, 0x0085, hiddenBound)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x000a, nil)
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("empty-sheets.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Visible Empty") {
		t.Fatalf("legacy XLS markdown missing empty visible sheet section:\n%s", md)
	}
	if strings.Contains(md, "Hidden Empty") {
		t.Fatalf("legacy XLS markdown kept hidden empty sheet:\n%s", md)
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("legacy XLS markdown should not backfill empty sheet names:\n%s", md)
	}
}

func TestLegacyXLSWorkbookTextBackfillsVisibleRowsPastMarkdownLimit(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 2)
	binary.LittleEndian.PutUint32(sst[4:], 2)
	sst = append(sst, testXLUnicodeString("Visible Row")...)
	sst = append(sst, testXLUnicodeString("Overflow Row")...)
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(0, 0, 0))
	writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(uint16(maxMarkdownTableRows), 0, 1))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("row-overflow.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Row", "## Workbook Text", "Overflow Row"} {
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing %q in:\n%s", want, md)
		}
	}
}

func TestLegacyXLSUTF8JapaneseCodePageMojibakeIsRepaired(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "12561-1.xls"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"作業", "画面"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing repaired Japanese text %q in %.400q", want, res.Text)
		}
	}
	for _, bad := range []string{"浣滄キ", "銉", "鐢婚潰", "瀹屼簡"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept UTF-8/codepage mojibake %q in %.400q", bad, res.Text)
		}
	}
}

func TestLegacyXLSInvalidFloatDisplayValuesAreDropped(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x0203, testNumberRecord(0, 0, 42.5))
	writeBIFFRecord(&workbook, 0x0203, testNumberRecord(0, 1, math.NaN()))
	writeBIFFRecord(&workbook, 0x0203, testNumberRecord(0, 2, math.Inf(1)))
	writeBIFFRecord(&workbook, 0x0006, testFormulaNumberRecord(1, 0, math.NaN()))
	writeBIFFRecord(&workbook, 0x0006, testFormulaNumberRecord(1, 1, math.Inf(-1)))
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("invalid-floats.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(res.Text, "42.5") || !strings.Contains(md, "42.5") {
		t.Fatalf("missing visible finite BIFF number: text=%q markdown=\n%s", res.Text, md)
	}
	for _, bad := range []string{"NaN", "+Inf", "-Inf"} {
		if strings.Contains(res.Text, bad) || strings.Contains(md, bad) {
			t.Fatalf("kept invalid BIFF float display value %q: text=%q markdown=\n%s", bad, res.Text, md)
		}
	}
}

func TestBIFFFormulaLikeTextIsFiltered(t *testing.T) {
	var out []string
	addBIFFText(&out, "SUM(A1:A10)")
	addBIFFText(&out, "AVERAGE(IF(A1:A3>0,1,0))")
	addBIFFText(&out, "Sheet1!$A$1:$B$2")
	addBIFFText(&out, "'Q1 Sales'!A1:B2")
	addBIFFText(&out, "Method 1: =SUM(A1:A10)")
	addBIFFText(&out, "Alert! A1 remains visible")
	if got := strings.Join(out, "\n"); got != "Method 1: =SUM(A1:A10)\nAlert! A1 remains visible" {
		t.Fatalf("unexpected BIFF text after formula filtering: %q", got)
	}
}

func TestLegacyXLSFormulaLikeCellTextIsFilteredFromOutput(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	name := "Visible Sheet"
	bound := make([]byte, 8+len(name))
	bound[6] = byte(len(name))
	bound[7] = 0
	copy(bound[8:], name)
	writeBIFFRecord(&workbook, 0x0085, bound)
	sst := make([]byte, 8)
	binary.LittleEndian.PutUint32(sst[0:], 8)
	binary.LittleEndian.PutUint32(sst[4:], 8)
	for _, value := range []string{
		"Visible Label",
		"SUM(A1:A10)",
		"=AVERAGE(IF(A1:A3>0,1,0))",
		"Sheet1!$A$1:$B$2",
		"'Q1 Sales'!A1:B2",
		"[Book1.xls]Sheet1!$C$3",
		"Method 1: =SUM(A1:A10)",
		"Alert! A1 remains visible",
	} {
		sst = append(sst, testXLUnicodeString(value)...)
	}
	writeBIFFRecord(&workbook, 0x00fc, sst)
	writeBIFFRecord(&workbook, 0x000a, nil)

	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	for i := 0; i < 8; i++ {
		writeBIFFRecord(&workbook, 0x00fd, testLabelSSTRecord(uint16(i), 0, uint32(i)))
	}
	writeBIFFRecord(&workbook, 0x000a, nil)

	res, err := extractLegacy("formula-like-text.xls", workbook.Bytes(), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Label", "Method 1: =SUM(A1:A10)", "Alert! A1 remains visible"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("legacy XLS text missing visible formula-like prose %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("legacy XLS markdown missing visible formula-like prose %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"SUM(A1:A10)", "=AVERAGE(IF(A1:A3>0,1,0))", "Sheet1!$A$1:$B$2", "'Q1 Sales'!A1:B2", "[Book1.xls]Sheet1!$C$3"} {
		if textContainsWholeLine(res.Text, hidden) {
			t.Fatalf("legacy XLS text kept formula/range expression %q in %q", hidden, res.Text)
		}
		if textContainsWholeLine(md, hidden) || strings.Contains(md, "| "+hidden+" |") {
			t.Fatalf("legacy XLS markdown kept formula/range expression %q in:\n%s", hidden, md)
		}
	}
}

func TestBIFFHeaderFooterControlCodesAreNotVisibleText(t *testing.T) {
	var workbook bytes.Buffer
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x05, 0x00})
	writeBIFFRecord(&workbook, 0x0015, testXLUnicodeString("Workbook global footer should not appear"))
	writeBIFFRecord(&workbook, 0x000a, nil)
	writeBIFFRecord(&workbook, 0x0809, []byte{0x00, 0x06, 0x10, 0x00})
	writeBIFFRecord(&workbook, 0x0014, testXLUnicodeString(`&L&"Arial,Bold"&14&OLeft Header Target: ../media/header.png&CPage &P of &N long &R&HRight && Header rId77`))
	writeBIFFRecord(&workbook, 0x0015, testXLUnicodeString(`&KFF0000&BRed Footer ContentType: image/png&G &OOutline Text &HShadow Text PartName=/xl/media/footer.png &K01+000Theme Color &K03-123Tinted Color &[Page] of &[Pages] &[Picture] &[Path]&[File]&[Tab]`))
	text := strings.Join(biffText(workbook.Bytes()), "\n")
	for _, want := range []string{"Left Header", "Page of", "long", "Right & Header", "Red Footer", "Outline Text", "Shadow Text", "Theme Color", "Tinted Color", "of"} {
		if !strings.Contains(text, want) {
			t.Fatalf("missing visible BIFF header/footer text %q in %q", want, text)
		}
	}
	for _, bad := range []string{"&L", "&C", "&R", "&P", "&N", "&G", "&O", "&H", "&[Page]", "&[Pages]", "&[Picture]", "&[Path]", "&[File]", "&[Tab]", "&KFF0000", "&K01+000", "&K03-123", "+000", "-123", "&B", "Arial,Bold", "&14", "Target:", "../media/header.png", "rId77", "ContentType:", "image/png", "PartName=", "/xl/media/footer.png"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept BIFF header/footer control code %q in %q", bad, text)
		}
	}
	if strings.Contains(text, "Workbook global footer should not appear") {
		t.Fatalf("kept workbook-global BIFF footer as visible text in %q", text)
	}
	md := extractLegacyMarkdown("header-footer.xls", workbook.Bytes(), nil)
	for _, want := range []string{"### Headers and Footers", "Left Header", "Page of", "long", "Right & Header", "Red Footer", "Outline Text", "Shadow Text", "Theme Color", "Tinted Color"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible BIFF header/footer text %q in:\n%s", want, md)
		}
	}
	for _, bad := range []string{"&L", "&C", "&R", "&P", "&N", "&G", "&O", "&H", "&[Page]", "&[Pages]", "&[Picture]", "&[Path]", "&[File]", "&[Tab]", "&KFF0000", "&K01+000", "&K03-123", "+000", "-123", "&B", "Arial,Bold", "&14", "Target:", "../media/header.png", "rId77", "ContentType:", "image/png", "PartName=", "/xl/media/footer.png", "Workbook global footer should not appear"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept BIFF header/footer hidden/control content %q in:\n%s", bad, md)
		}
	}
}

func TestUnicodeBinaryNoiseFilterKeepsRealCJKText(t *testing.T) {
	if !looksLikeTextFragment("\U0001D74A") {
		t.Fatalf("classified mathematical Greek symbol as non-text")
	}
	for _, good := range []string{
		"\u753b\u9762\u306b\u5408\u308f\u305b\u308b (4:3)",
		"\u8907\u6570\u306e\u6587\u5b57",
		"\u4e2d\u6587\u6587\u672c\u63d0\u53d6\u6d4b\u8bd5",
		"\u6587\u6863\u4e2d\u7684\u811a\u6ce8\u548c\u6279\u6ce8\u5e94\u8be5\u4fdd\u7559",
		"\u4e2d\u00e9\u6587",
		"\u4e2dA\u6587",
		"\U0001D74A\U0001D74B\U0001D74C\U0001D74D\U0001D74E",
	} {
		if looksLikeUnicodeBinaryNoise(good) || looksLikeBinaryControlFragment(good) {
			t.Fatalf("classified real text as binary noise: %q", good)
		}
	}
	for _, bad := range []string{
		"\u14c0\u5b00\u172d",
		"\u00e0\u8000\u81a7",
		"\u5b00\ud61b\u9300\u0e02\u3220\u0900",
		"\u846a\u887c\u91c0\u5425\u20ac\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9\u81a9",
		"\u844a\u887b\u81b0Calibri1",
		"\u8880\u91d4\u20ac\u8641\u5f40\u20ac",
		"\u704f\u822c\ue0c4\u72fb\u71b0\u91dc\u8a5f\u844a",
		"\u93c7\u5a17\u639a\u5e0e",
		"\u8989\u8989\u9689",
		"\u4980\u00c6\uc000",
		"\ud180\u00ef\u1800",
		"\ud380H\u2600",
		"\u0180\u0187\u2200",
		"\u9f80\u01e5\uc000a\uc140\u01e5\u3c00",
		"\u0100\u4100\u8003\u90f5\u8003\u3007",
		"\u8980\u00b7\u5800",
		"\u48fb\u560b\u0100",
		"\u92ef\u4f77\u59ac\u92f7\u4f77\u62cb\u92ff\u4f77\u6b4e\u9307\u4f77\ue566",
		"\u60c6\u20ac\u87f9\u4e60",
		"\u9287\u641e\u7220\u61c8\u9287\u611f\u641e\u7273",
		"\u91c4\u20ac\u4f33\u638b\ue625\u60c6\u6378",
		"\u4181\u4281\u4381\u4481\u4581\u4681\u4781",
		"\ua400\u0481\u0240",
		"\ub2e1\u00bd\u7bf7\u9cff",
		"\u0302\u0311\u0311\u0311\u0311\u0311\u0311\u0311",
		"\u9c31\u6300\u63ce",
		"\uce00\uce1c\u6300\u3131\u9c00\u31ff\uc240",
		"\u0c8c\u07caB",
		"\u00ff\ucf00\u00ff",
		"\u35c5j\u324a\u4242B\u105c",
		"\uA5DB-\u4031\u0409",
		"\u0446\u0446\u0446\u0446\u0446\u0446\u0446",
		"\u043A\u0432\u0426\u041E\u0426\u0455\u0426\u0455\u00AE\u0455\u0426\u0408\u0432\u2014\u0426\u0432\u040A",
		"\u5900 \u7a00 \u7c00 \u7e00 \u8000",
		"\u1A12\u5400\u6D69\u7365\u4E20\u7775\u5220\u6D6F\u6E61\u0900",
		"\u5674\u6666\u3066\u2400\u4444@\u5514R\u0300\u5255",
		"\ucffa\ubfc4\uaeb4\u929d\u888d\u7277\u8888r",
		"\ud7e9\uc0d2\ua9bb\u92a4\u7b8d\u6476\u4d5f\u36511",
		"\ucc99\u00ff\u99ff\u00cc\u99cc\u00ff\ucc7f\u6633\u00ff\ucc33\u00cc\ucc99",
		"\u3399f\u3333\u33333\u105c",
	} {
		if !looksLikeUnicodeBinaryNoise(bad) || !looksLikeBinaryControlFragment(bad) {
			t.Fatalf("did not classify binary-looking text as noise: %q", bad)
		}
	}
}

func TestCleanVisibleTextKeepsShortCJKAndKoreanLines(t *testing.T) {
	in := "\u9879\u76ee\u540d\u79f0\n\u5ba2\u6237\u540d\u79f0\n\u5ba1\u6279\u72b6\u6001\n\ud55c\uae00\n\ud14c\uc2a4\ud2b8\n\ubb38\uc11c"
	got := cleanVisibleText(in)
	if got != in {
		t.Fatalf("cleanVisibleText removed visible short CJK/Korean lines:\n got %q\nwant %q", got, in)
	}
}

func TestCleanVisibleTextKeepsCyrillicSheetNames(t *testing.T) {
	in := "\u041b\u0438\u0441\u04421\n\u041b\u0438\u0441\u04422\n\u041b\u0438\u0441\u04423"
	got := cleanVisibleText(in)
	if got != in {
		t.Fatalf("cleanVisibleText removed Cyrillic sheet names:\n got %q\nwant %q", got, in)
	}
}

func TestCleanTextDecodesOOXMLControlEscapes(t *testing.T) {
	got := cleanText("Frequency_x000a_of Repair\nRich Text_x000d_\nEscaped_x005F_x000D_Tail\nLiteral_x005F_Name\nEmoji_xD83D__xDE00_Text")
	want := "Frequency\nof Repair\nRich Text\nEscaped_x000D_Tail\nLiteral_Name\nEmoji😀Text"
	if got != want {
		t.Fatalf("cleanText did not normalize OOXML control escapes:\n got %q\nwant %q", got, want)
	}
	cell := escapeMarkdownTableCell("Visible_x005F_x000D_Cell")
	if cell != "Visible_x000D_Cell" {
		t.Fatalf("markdown table cell should keep escaped OOXML control literal, got %q", cell)
	}
}

func TestLegacySamplesDoNotEmitUnicodeBinaryNoise(t *testing.T) {
	cases := map[string][]string{
		"61300.xls": {"\u5b00", "Root Entry", "\u844a\u887b\u81b0Calibri1", "\u704f\u822c\ue0c4\u72fb", "\ud23f\ub7a9\uc2f1\ubfc4", "\u4800\u6165\u6964\u676e"},
		"word2.doc": {
			"\u5b00", "\ua5db-\u4031", "\u0446\u0446\u0446", "\u1a12\u5400\u6d69",
			"\u5900 \u7a00 \u7c00", "\ucfd0\u0121", "\ucffa\ubfc4", "\ud7e9\uc0d2",
		},
		"pp40only.ppt": {
			"\u81a7", "\u4980\u00c6\uc000", "\ud180\u00ef\u1800", "\ud380H\u2600", "\u0180\u0187\u2200",
			"\u9f80\u01e5\uc000a\uc140\u01e5\u3c00", "\u0100\u4100\u8003\u90f5\u8003\u3007",
			"\u8980\u00b7\u5800", "\u48fb\u560b\u0100", "\u5674\u6666\u3066",
		},
		"64130.xls":    {"\u0c8c\u07caB", "\u00ff\ucf00\u00ff", "\u35c5j\u324a"},
		"Bug47731.doc": {"\u0432\u2014\u0426", "\u1a00\u0124", "\uc7f8\uc7b6"},
		"at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx": {"\ucc99\u00ff\u99ff", "\u3399f\u3333", "\u01df\u3fd0"},
		"PPT95.ppt": {
			"\u5b00", "\u8989\u8989\u9689", "\u9289\uae7b",
			"\u92ef\u4f77\u59ac\u92f7\u4f77\u62cb", "\u60c6\u20ac\u87f9", "\u9287\u641e\u7220",
			"\u91c4\u20ac\u4f33\u638b", "\u4181\u4281\u4381", "\ua400\u0481\u0240",
			"\ub2e1\u00bd\u7bf7", "\u0302\u0311\u0311", "\u9c31\u6300\u63ce", "\uce00\uce1c\u6300",
			"\u8100\u8165\u8167\u8169\u816b\u816d\u816f\u8171\u8173\u8175\u8177\u8179\u818f\u2490",
			"\u5700\u6168\u2074", "\u4d00\u6e61\u2079", "\u4a00\u7375\u2074",
		},
		"cf5f6fde99a8b3ea5a4946c258b7abad6f30b0c5.ppt": {"Root Entry", "\u8880\u91d4\u20ac\u8641\u5f40\u20ac"},
	}
	for sample, badParts := range cases {
		t.Run(sample, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", sample), Options{})
			if err != nil {
				t.Fatal(err)
			}
			for _, bad := range badParts {
				if strings.Contains(res.Text, bad) {
					t.Fatalf("kept binary-looking Unicode fragment %q in %.400q", bad, res.Text)
				}
			}
		})
	}
}

func TestLegacyXLSPreservesCyrillicSheetNames(t *testing.T) {
	filename := filepath.Join("testdata", "samples", "56325.xls")
	res, err := Extract(filename, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "\u041b\u0438\u0441\u04421") {
		t.Fatalf("missing Cyrillic sheet name in extracted text: %q", res.Text)
	}
}

func TestXLSXHiddenDefinedNamesAreNotVisibleText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "absolute-anchor-over-empty-sheet.xlsx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, hidden := range []string{"Pal_Workbook_GUID", "SXBLYIZ9P5XS2DZ2CT6MNIAD", "_AtRisk_SimSetting"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden defined name text %q in %.400q", hidden, res.Text)
		}
	}
	if !strings.Contains(res.Text, "picture") {
		t.Fatalf("missing visible sheet name in %.400q", res.Text)
	}
}

func TestXLSXMarkdownUsesWorkbookSheetNameWithoutRelationships(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "generated-xlsx-threaded-comments.xlsx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Generated", "Generated Threaded Host", "Generated XLSX Threaded Comment Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing %q in:\n%s", want, md)
		}
	}
	for _, bad := range []string{"## sheet1", "## Additional Text"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept fallback/internal section %q in:\n%s", bad, md)
		}
	}
}

func TestXLSXMarkdownKeepsWorkbookSheetTabsWithoutRelationships(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "generated-xlsx-sheetnames.xlsx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Generated Visible Sheet Name", "## Generated Second Tab"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing workbook sheet tab %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("workbook sheet tabs should be structured instead of backfilled:\n%s", md)
	}
}

func TestXLSXMarkdownStructuresVMLDrawingText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "generated-xlsx-vml-drawing.xlsx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Generated VML Sheet", "## Drawings", "Generated XLSX VML Drawing Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing VML drawing text %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("VML drawing text should be structured instead of backfilled:\n%s", md)
	}
}

func TestXLSXMarkdownWideSheetsDoNotBackfillTruncatedColumns(t *testing.T) {
	tests := []struct {
		name string
		want []string
	}{
		{name: "49609.xlsx", want: []string{"## FAM", "## VIC", "H12"}},
		{name: "65016.xlsx", want: []string{"2014-03-26 11:14:52", "2014-03-26 11:15:43"}},
		{name: "50096.xlsx", want: []string{"## Tabelle1", "300"}},
		{name: "58896.xlsx", want: []string{"## Sheet0", "Cost of living"}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			md := res.Markdown("images")
			for _, want := range tc.want {
				if !strings.Contains(md, want) {
					t.Fatalf("sheet markdown missing %q in %.800q", want, md)
				}
			}
			if strings.Contains(md, "## Additional Text") {
				t.Fatalf("sheet markdown should not backfill truncated table rows/columns:\n%s", md)
			}
		})
	}
}

func TestXLSXMarkdownStructuresWorkbookNamesAndTables(t *testing.T) {
	tests := []struct {
		name string
		want []string
	}{
		{name: "dataValidationTableRange.xlsx", want: []string{"## Workbook Names", "## Tables", "Table_Query_from_RDS31214"}},
		{name: "StructuredRefs-lots-with-lookups.xlsx", want: []string{"## Workbook Names", "## Tables", "countylist", "Table_Query_from_RDS24"}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			md := res.Markdown("images")
			for _, want := range tc.want {
				if !strings.Contains(md, want) {
					t.Fatalf("markdown missing structured workbook/table text %q in %.800q", want, md)
				}
			}
			if strings.Contains(md, "## Additional Text") {
				t.Fatalf("workbook/table text should be structured instead of backfilled:\n%s", md)
			}
		})
	}
}

func TestMarkdownTableCellKeepsVisibleMultilineCyrillicNotes(t *testing.T) {
	got := cleanMarkdownTableCellValue("В валютах цен.\nЦены указаны на 28.02.2020")
	for _, want := range []string{"В валютах цен.", "Цены указаны на 28.02.2020"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleaned markdown cell missing visible note %q in %q", want, got)
		}
	}
}

func TestPrepareMarkdownTableCellValueSingleLineFastPath(t *testing.T) {
	if got := prepareMarkdownTableCellValue("  Alpha|Beta\\Gamma  "); got != "Alpha\\|Beta\\\\Gamma" {
		t.Fatalf("single-line markdown table cell fast path got %q", got)
	}
	if got := prepareMarkdownTableCellValue("{\\rtf1\\ansi Visible \\line Cell}"); got != "Visible<br>Cell" {
		t.Fatalf("single-line RTF markdown table cell should still render visible text, got %q", got)
	}
}

func TestBIFFCellRecordInMarkdownBounds(t *testing.T) {
	rec := make([]byte, 6)
	binary.LittleEndian.PutUint16(rec[0:], 2)
	binary.LittleEndian.PutUint16(rec[2:], 3)
	if !biffCellRecordInMarkdownBounds(rec, nil, nil) {
		t.Fatal("visible in-bounds BIFF cell record should be accepted")
	}

	hiddenRows := map[int]bool{2: true}
	if biffCellRecordInMarkdownBounds(rec, hiddenRows, nil) {
		t.Fatal("hidden row BIFF cell record should be rejected")
	}

	hiddenCols := []intRange{{min: 3, max: 3}}
	if biffCellRecordInMarkdownBounds(rec, nil, hiddenCols) {
		t.Fatal("hidden column BIFF cell record should be rejected")
	}

	binary.LittleEndian.PutUint16(rec[0:], maxMarkdownTableRows)
	if biffCellRecordInMarkdownBounds(rec, nil, nil) {
		t.Fatal("row beyond markdown table limit should be rejected")
	}

	binary.LittleEndian.PutUint16(rec[0:], 0)
	binary.LittleEndian.PutUint16(rec[2:], maxMarkdownTableCols)
	if biffCellRecordInMarkdownBounds(rec, nil, nil) {
		t.Fatal("column beyond markdown table limit should be rejected")
	}
}

func TestXLSXMarkdownDoesNotBackfillKnownCleanSamples(t *testing.T) {
	tests := []struct {
		name string
		want []string
	}{
		{name: "no_drawing_patriarch.xlsx", want: []string{"В валютах цен.", "Цены указаны на 28.02.2020"}},
		{name: "stress002.xlsx", want: []string{"## Images", "Self-Employment Calculator"}},
		{name: "SimpleNormal.xlsx", want: []string{"## Sheet Number 2", "| 1 | 2 | 3 | 4 | 5 | 6 |"}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			res, err := Extract(filepath.Join("testdata", "samples", tc.name), Options{})
			if err != nil {
				t.Fatal(err)
			}
			md := res.Markdown("images")
			for _, want := range tc.want {
				if !strings.Contains(md, want) {
					t.Fatalf("markdown missing expected visible text %q in %.800q", want, md)
				}
			}
			if strings.Contains(md, "## Additional Text") {
				t.Fatalf("%s should not need Additional Text:\n%s", tc.name, md)
			}
		})
	}
}

func TestXLSXHiddenSheetsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/><sheet name="Hidden Sheet" state="hidden" r:id="rId2"/><sheet name="Very Hidden Sheet" state="veryHidden" r:id="rId3"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="x" Target="/xl/worksheets/sheet3.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Secret</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet3.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Very Hidden Secret</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheets.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible sheet text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Very Hidden Sheet", "Hidden Secret", "Very Hidden Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden sheet text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible sheet text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Very Hidden Sheet", "Hidden Secret", "Very Hidden Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden sheet text %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXMarkdownKeepsEmptyVisibleSheetsStructured(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "picture.xlsx"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Sheet1", "## Sheet2", "## Sheet3", "Lorem", "## Drawings"} {
		if !strings.Contains(md, want) {
			t.Fatalf("picture.xlsx markdown missing structured content %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "## Additional Text") {
		t.Fatalf("picture.xlsx markdown should structure empty visible sheets instead of backfilling them:\n%s", md)
	}
}

func TestXLSXSheetNamesDropInternalResourceReferences(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet Target: ../media/sheet.png rId77 ContentType: image/png" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "sheet-name-internal-refs.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing clean sheet text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/sheet.png", "rId77", "ContentType:", "image/png"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept internal sheet-name text %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept internal sheet-name text %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXHiddenStateWhitespaceAndBooleanDefinedNamesAreNotVisible(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/><sheet name="Hidden With Spaces" state=" hidden " r:id="rId2"/><sheet name="Very Hidden With Spaces" state=" veryHidden " r:id="rId3"/></sheets><definedNames><definedName name="VisibleDefinedName">VisibleDefinedValue</definedName><definedName name="HiddenDefinedName" hidden=" true ">HiddenDefinedValue</definedName></definedNames></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="x" Target="worksheets/sheet3.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Space Secret</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet3.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Very Hidden Space Secret</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-state-whitespace.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "VisibleDefinedName", "VisibleDefinedValue"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden With Spaces", "Very Hidden With Spaces", "Hidden Space Secret", "Very Hidden Space Secret", "HiddenDefinedName", "HiddenDefinedValue"} {
		if strings.Contains(res.Text, hidden) || strings.Contains(res.Markdown("images"), hidden) {
			t.Fatalf("kept hidden XLSX text %q in text=%q markdown=\n%s", hidden, res.Text, res.Markdown("images"))
		}
	}
}

func TestXLSXHiddenSheetTablesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/><sheet name="Hidden Sheet" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Hidden Cell Secret</t></is></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdTable"/></tableParts></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/></Relationships>`)
	addZip(t, zw, "xl/tables/table1.xml", `<table xmlns="urn:x" name="VisibleTableName Target: ../media/hidden-table.png" displayName="Visible Table Display Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"><tableColumns count="1"><tableColumn id="1" name="Visible Table Column ContentType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></tableColumns></table>`)
	addZip(t, zw, "xl/tables/table2.xml", `<table xmlns="urn:x" name="HiddenTableSecret" displayName="Hidden Table Display Secret"><tableColumns count="1"><tableColumn id="1" name="Hidden Table Column Secret"/></tableColumns></table>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheet-tables.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "VisibleTableName", "Visible Table Display", "Visible Table Column"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX table content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "HiddenTableSecret", "Hidden Table Display Secret", "Hidden Table Column Secret", "Target:", "../media/hidden-table.png", "Type:", "relationships/image", "ContentType:", "application/vnd.openxmlformats"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden XLSX table content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "VisibleTableName", "Visible Table Display", "Visible Table Column"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX table content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "HiddenTableSecret", "Hidden Table Display Secret", "Hidden Table Column Secret", "Target:", "../media/hidden-table.png", "Type:", "relationships/image", "ContentType:", "application/vnd.openxmlformats"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden XLSX table content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXUnreferencedTablesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/tables/table1.xml", `<table xmlns="urn:x" name="InternalUnreferencedTableSecret" displayName="Internal Unreferenced Table Display"><tableColumns count="1"><tableColumn id="1" name="Internal Unreferenced Column Secret"/></tableColumns></table>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-table.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"InternalUnreferencedTableSecret", "Internal Unreferenced Table Display", "Internal Unreferenced Column Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced XLSX table content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"InternalUnreferencedTableSecret", "Internal Unreferenced Table Display", "Internal Unreferenced Column Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced XLSX table content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXHiddenSheetCommentsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/><sheet name="Hidden Sheet" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Cell Secret</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="../comments1.xml"/><Relationship Id="rId2" Type="x" Target="../threadedComments/threadedComment1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="../comments2.xml"/><Relationship Id="rId2" Type="x" Target="../threadedComments/threadedComment2.xml"/></Relationships>`)
	addZip(t, zw, "xl/comments1.xml", `<comments xmlns="urn:x"><commentList><comment ref="A1"><text><r><t>Visible Comment Target: ../media/comment.png Text</t></r></text></comment></commentList></comments>`)
	addZip(t, zw, "xl/comments2.xml", `<comments xmlns="urn:x"><commentList><comment ref="A1"><text><r><t>Hidden Comment Secret</t></r></text></comment></commentList></comments>`)
	addZip(t, zw, "xl/threadedComments/threadedComment1.xml", `<ThreadedComments xmlns="urn:x"><threadedComment ref="A1"><text>Visible Threaded Comment Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image Text</text></threadedComment></ThreadedComments>`)
	addZip(t, zw, "xl/threadedComments/threadedComment2.xml", `<ThreadedComments xmlns="urn:x"><threadedComment ref="A1"><text>Hidden Threaded Comment Secret</text></threadedComment></ThreadedComments>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheet-comments.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Comment Text", "Visible Threaded Comment Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX comment content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "Hidden Comment Secret", "Hidden Threaded Comment Secret", "Target:", "../media/comment.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden XLSX comment content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "## Comments", "### A1", "Visible Comment Text", "Visible Threaded Comment Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX comment content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "Hidden Comment Secret", "Hidden Threaded Comment Secret", "Target:", "../media/comment.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden XLSX comment content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXHiddenCellCommentsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><cols><col min="2" max="2" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Cell</t></is></c><c r="B1" t="inlineStr"><is><t>Hidden Column Cell</t></is></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>Hidden Row Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="x" Target="../comments1.xml"/><Relationship Id="rIdThreaded" Type="x" Target="../threadedComments/threadedComment1.xml"/></Relationships>`)
	addZip(t, zw, "xl/comments1.xml", `<comments xmlns="urn:x"><commentList><comment ref="A1"><text><r><t>Visible Cell Comment</t></r></text></comment><comment ref="$B$1"><text><r><t>Hidden Absolute Column Comment Secret</t></r></text></comment><comment ref="$A$2"><text><r><t>Hidden Absolute Row Comment Secret</t></r></text></comment></commentList></comments>`)
	addZip(t, zw, "xl/threadedComments/threadedComment1.xml", `<ThreadedComments xmlns="urn:x"><threadedComment ref="A1"><text>Visible Threaded Cell Comment</text></threadedComment><threadedComment ref="$B$1"><text>Hidden Absolute Threaded Column Secret</text></threadedComment><threadedComment ref="$A$2"><text>Hidden Absolute Threaded Row Secret</text></threadedComment></ThreadedComments>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-cell-comments.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Cell Comment", "Visible Threaded Cell Comment"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible hidden-cell XLSX comment content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Column Cell", "Hidden Row Cell", "Hidden Absolute Column Comment Secret", "Hidden Absolute Row Comment Secret", "Hidden Absolute Threaded Column Secret", "Hidden Absolute Threaded Row Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden-cell XLSX comment content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "## Comments", "### A1", "Visible Cell Comment", "Visible Threaded Cell Comment"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible hidden-cell XLSX comment content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Column Cell", "Hidden Row Cell", "### $B$1", "### $A$2", "Hidden Absolute Column Comment Secret", "Hidden Absolute Row Comment Secret", "Hidden Absolute Threaded Column Secret", "Hidden Absolute Threaded Row Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden-cell XLSX comment content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXCommentInvalidRefsAreNotMarkdownHeadings(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="x" Target="../comments1.xml"/><Relationship Id="rIdThreaded" Type="x" Target="../threadedComments/threadedComment1.xml"/></Relationships>`)
	addZip(t, zw, "xl/comments1.xml", `<comments xmlns="urn:x"><comment ref="A1 Target: ../media/comment-ref.png"><text><r><t>Visible malformed ref comment</t></r></text></comment></comments>`)
	addZip(t, zw, "xl/threadedComments/threadedComment1.xml", `<ThreadedComments xmlns="urn:x"><threadedComment ref="B2 rId77"><text>Visible malformed threaded ref</text></threadedComment></ThreadedComments>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "invalid-comment-refs.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible malformed ref comment", "Visible malformed threaded ref"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible malformed-ref comment text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Comments", "Visible malformed ref comment", "Visible malformed threaded ref"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing malformed-ref comment text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"../media/comment-ref.png", "comment-ref.png", "rId77", "Target:", "### A1", "### B2"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept invalid comment ref heading content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXUnreferencedCommentsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/comments1.xml", `<comments xmlns="urn:x"><commentList><comment ref="A1"><text><r><t>Internal Unreferenced Comment Secret</t></r></text></comment></commentList></comments>`)
	addZip(t, zw, "xl/threadedComments/threadedComment1.xml", `<ThreadedComments xmlns="urn:x"><threadedComment ref="A1"><text>Internal Unreferenced Threaded Secret</text></threadedComment></ThreadedComments>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-comments.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced Comment Secret", "Internal Unreferenced Threaded Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced XLSX comment content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced Comment Secret", "Internal Unreferenced Threaded Secret", "## Comments"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced XLSX comment content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXHiddenSheetChartTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/><sheet name="Hidden Sheet" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Hidden Cell Secret</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Visible Drawing Target: ../media/drawing.png Description"/></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing2.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Hidden Drawing Secret"/></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>`)
	addZip(t, zw, "xl/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Visible XLSX Chart Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image Title</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "xl/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Hidden XLSX Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheet-chart.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Drawing Description", "Visible XLSX Chart Title"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX chart content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "Hidden Drawing Secret", "Hidden XLSX Chart Secret", "Target:", "../media/drawing.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden XLSX chart content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "## Drawings", "Visible Drawing Description", "Visible XLSX Chart Title"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX chart content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "Hidden Drawing Secret", "Hidden XLSX Chart Secret", "Target:", "../media/drawing.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden XLSX chart content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXHiddenDrawingChartTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdVisibleChart"/></a:graphicData></a:graphic></xdr:graphicFrame>
<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Hidden Chart Frame" hidden="1"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdHiddenChart"/></a:graphicData></a:graphic></xdr:graphicFrame>
</xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisibleChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdHiddenChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/><Relationship Id="rIdUnreferencedChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/></Relationships>`)
	addZip(t, zw, "xl/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Visible XLSX Chart Text</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "xl/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Hidden Drawing Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "xl/charts/chart3.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Unreferenced Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-drawing-chart.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible XLSX Chart Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX chart text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Drawing Chart Secret", "Unreferenced Chart Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden XLSX chart text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "## Drawings", "Visible XLSX Chart Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX chart text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Drawing Chart Secret", "Unreferenced Chart Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden XLSX chart text %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXUnreferencedDrawingTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Visible Drawing Description"/></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/drawing2.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Internal Unreferenced Drawing Secret"/></xdr:wsDr>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-drawing-text.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Drawing Description"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX drawing text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Internal Unreferenced Drawing Secret") {
		t.Fatalf("kept unreferenced XLSX drawing text in %q", res.Text)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "## Drawings", "Visible Drawing Description"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX drawing text %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "Internal Unreferenced Drawing Secret") {
		t.Fatalf("markdown kept unreferenced XLSX drawing text in:\n%s", md)
	}
}

func TestXLSXAllHiddenSheetsDoNotFallbackToHiddenText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Hidden Sheet" state="hidden" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Secret</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "all-hidden-sheets.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(res.Text, "Hidden Sheet") || strings.Contains(res.Text, "Hidden Secret") {
		t.Fatalf("kept all-hidden sheet text in %q", res.Text)
	}
}

func TestXLSXHiddenRowsAndColumnsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><cols><col min="2" max="2" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Cell</t></is></c><c r="B1" t="inlineStr"><is><t>Hidden Column Secret</t></is></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>Hidden Row Secret</t></is></c><c r="C2"><f>HIDDEN_FORMULA()</f><v>Hidden Formula Result</v></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>Visible Tail</t></is></c></row></sheetData><hyperlinks><hyperlink ref="$B$1" display="Hidden Hyperlink Display" tooltip="Hidden Hyperlink Tooltip"/><hyperlink ref="$C$3" display="Visible Hyperlink Display" tooltip="Visible Hyperlink Tooltip"/></hyperlinks><dataValidations><dataValidation sqref="$A$2,$B$1" promptTitle="Hidden Prompt Title" prompt="Hidden Prompt Body"/><dataValidation sqref="$B$1 $C$3" promptTitle="Visible Prompt Title" prompt="Visible Prompt Body"/></dataValidations></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-rows-cols.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Tail", "Visible Hyperlink Display", "Visible Hyperlink Tooltip", "Visible Prompt Title", "Visible Prompt Body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Column Secret", "Hidden Row Secret", "HIDDEN_FORMULA", "Hidden Formula Result", "Hidden Hyperlink Display", "Hidden Hyperlink Tooltip", "Hidden Prompt Title", "Hidden Prompt Body"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden row/column text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "Visible Tail", "### Annotations", "Visible Hyperlink Display", "Visible Hyperlink Tooltip", "Visible Prompt Title", "Visible Prompt Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX annotation %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Column Secret", "Hidden Row Secret", "Hidden Hyperlink Display", "Hidden Hyperlink Tooltip", "Hidden Prompt Title", "Hidden Prompt Body"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden row/column annotation %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXImplicitHiddenRowsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c r="A1" t="inlineStr"><is><t>Visible First Row</t></is></c></row><row hidden="1"><c r="A2" t="inlineStr"><is><t>Implicit Hidden Row Secret</t></is></c></row><row><c r="A3" t="inlineStr"><is><t>Visible Third Row</t></is></c></row></sheetData><hyperlinks><hyperlink ref="A2" display="Implicit Hidden Hyperlink" tooltip="Implicit Hidden Tooltip"/><hyperlink ref="A3" display="Visible Hyperlink" tooltip="Visible Tooltip"/></hyperlinks><dataValidations><dataValidation sqref="A2" promptTitle="Implicit Hidden Prompt" prompt="Implicit Hidden Prompt Body"/><dataValidation sqref="A3" promptTitle="Visible Prompt" prompt="Visible Prompt Body"/></dataValidations></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "implicit-hidden-row.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible First Row", "Visible Third Row", "Visible Hyperlink", "Visible Tooltip", "Visible Prompt", "Visible Prompt Body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible implicit-row XLSX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Implicit Hidden Row Secret", "Implicit Hidden Hyperlink", "Implicit Hidden Tooltip", "Implicit Hidden Prompt", "Implicit Hidden Prompt Body"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept implicit hidden row text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible First Row", "Visible Third Row", "Visible Hyperlink", "Visible Tooltip", "Visible Prompt", "Visible Prompt Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible implicit-row XLSX text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Implicit Hidden Row Secret", "Implicit Hidden Hyperlink", "Implicit Hidden Tooltip", "Implicit Hidden Prompt", "Implicit Hidden Prompt Body"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept implicit hidden row text %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXImplicitHiddenColumnsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><cols><col min="2" max="2" hidden="1"/></cols><sheetData><row><c t="inlineStr"><is><t>Visible A1</t></is></c><c t="inlineStr"><is><t>Implicit Hidden Column Secret</t></is></c><c t="inlineStr"><is><t>Visible C1</t></is></c></row></sheetData><hyperlinks><hyperlink ref="B1" display="Implicit Hidden Column Hyperlink" tooltip="Implicit Hidden Column Tooltip"/><hyperlink ref="C1" display="Visible Hyperlink" tooltip="Visible Tooltip"/></hyperlinks><dataValidations><dataValidation sqref="B1" promptTitle="Implicit Hidden Column Prompt" prompt="Implicit Hidden Column Prompt Body"/><dataValidation sqref="C1" promptTitle="Visible Prompt" prompt="Visible Prompt Body"/></dataValidations></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "implicit-hidden-column.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible A1", "Visible C1", "Visible Hyperlink", "Visible Tooltip", "Visible Prompt", "Visible Prompt Body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible implicit-column XLSX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Implicit Hidden Column Secret", "Implicit Hidden Column Hyperlink", "Implicit Hidden Column Tooltip", "Implicit Hidden Column Prompt", "Implicit Hidden Column Prompt Body"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept implicit hidden column text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible A1", "Visible C1", "Visible Hyperlink", "Visible Tooltip", "Visible Prompt", "Visible Prompt Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible implicit-column XLSX text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Implicit Hidden Column Secret", "Implicit Hidden Column Hyperlink", "Implicit Hidden Column Tooltip", "Implicit Hidden Column Prompt", "Implicit Hidden Column Prompt Body"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept implicit hidden column text %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXZeroSizedRowsAndColumnsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><cols><col min="2" max="2" width="0" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Cell</t></is></c><c r="B1" t="inlineStr"><is><t>Zero Width Column Secret</t></is></c></row><row r="2" ht="0" customHeight="1"><c r="A2" t="inlineStr"><is><t>Zero Height Row Secret</t></is></c><c r="C2"><f>ZERO_HEIGHT_FORMULA()</f><v>Zero Height Formula Result</v></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>Visible Tail</t></is></c></row></sheetData><hyperlinks><hyperlink ref="B1" display="Zero Width Hyperlink Display" tooltip="Zero Width Hyperlink Tooltip"/><hyperlink ref="A2" display="Zero Height Hyperlink Display" tooltip="Zero Height Hyperlink Tooltip"/><hyperlink ref="C3" display="Visible Hyperlink Display" tooltip="Visible Hyperlink Tooltip"/></hyperlinks><dataValidations><dataValidation sqref="A2" promptTitle="Zero Height Prompt Title" prompt="Zero Height Prompt Body"/><dataValidation sqref="B1" promptTitle="Zero Width Prompt Title" prompt="Zero Width Prompt Body"/><dataValidation sqref="C3" promptTitle="Visible Prompt Title" prompt="Visible Prompt Body"/></dataValidations></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "zero-sized-rows-cols.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Tail", "Visible Hyperlink Display", "Visible Hyperlink Tooltip", "Visible Prompt Title", "Visible Prompt Body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible zero-sized XLSX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Zero Width Column Secret", "Zero Height Row Secret", "ZERO_HEIGHT_FORMULA", "Zero Height Formula Result", "Zero Width Hyperlink Display", "Zero Height Hyperlink Display", "Zero Width Prompt Title", "Zero Height Prompt Title"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept zero-sized row/column text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "Visible Tail", "### Annotations", "Visible Hyperlink Display", "Visible Hyperlink Tooltip", "Visible Prompt Title", "Visible Prompt Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible zero-sized XLSX text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Zero Width Column Secret", "Zero Height Row Secret", "Zero Width Hyperlink Display", "Zero Height Hyperlink Display", "Zero Width Prompt Title", "Zero Height Prompt Title"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept zero-sized row/column text %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXSimpleInlineZeroHeightRowsDoNotLeakThroughFastPath(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Fast Cell</t></is></c></row><row r="2" ht="0" customHeight="1"><c r="A2" t="inlineStr"><is><t>Fast Path Zero Height Secret</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Visible Fast Tail</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "simple-inline-zero-height.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Fast Cell", "Visible Fast Tail"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible simple-inline XLSX text %q in %q", want, res.Text)
		}
		if !strings.Contains(res.Markdown("images"), want) {
			t.Fatalf("markdown missing visible simple-inline XLSX text %q in:\n%s", want, res.Markdown("images"))
		}
	}
	if strings.Contains(res.Text, "Fast Path Zero Height Secret") {
		t.Fatalf("text leaked zero-height simple-inline row through fast path: %q", res.Text)
	}
	if strings.Contains(res.Markdown("images"), "Fast Path Zero Height Secret") {
		t.Fatalf("markdown leaked zero-height simple-inline row through fast path:\n%s", res.Markdown("images"))
	}
}

func TestXLSXRelationshipIDAttributesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData><hyperlinks><hyperlink ref="A1" display="rId23" tooltip="Visible Tooltip"/></hyperlinks><dataValidations><dataValidation sqref="A1" promptTitle="Visible Prompt Title" prompt="rId24" errorTitle="rId25" error="Visible Error Body"/></dataValidations></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "relationship-id-attrs.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Tooltip", "Visible Prompt Title", "Visible Error Body"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX attribute text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"rId23", "rId24", "rId25"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept relationship ID attribute %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Tooltip", "Visible Prompt Title", "Visible Error Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX attribute text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"rId23", "rId24", "rId25"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept relationship ID attribute %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXBooleanCellsUseVisibleDisplayText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Booleans" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Flags</t></is></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="b"><v>0</v></c></row><row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>Hidden Bool Row</t></is></c><c r="B2" t="b"><v>1</v></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "boolean-cells.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Booleans", "Flags", "TRUE", "FALSE"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible boolean cell text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Hidden Bool Row") {
		t.Fatalf("kept hidden boolean row text in %q", res.Text)
	}
}

func TestXLSXMarkdownUsesWorksheetTables(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Data" r:id="rId1"/><sheet name="Hidden Data" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/sharedStrings.xml", `<sst xmlns="urn:x"><si><t>Name</t></si><si><t>Ready</t></si><si><t>Alice</t></si><si><t>Bob</t></si><si><t>Hidden Column Secret</t></si></sst>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><cols><col min="3" max="3" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>4</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="b"><v>1</v></c><c r="C2" t="inlineStr"><is><t>Hidden Cell</t></is></c></row><row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3" t="b"><v>0</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Carol</t></is></c><c r="B4" t="inlineStr"><is><t>Visible Target: ../media/table.png table text</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Sheet Secret</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "markdown-table.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{
		"## Visible Data",
		"| Name | Ready |",
		"| --- | --- |",
		"| Alice | TRUE |",
		"| Bob | FALSE |",
		"| Carol | Visible table text |",
	} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Data", "Hidden Sheet Secret", "Hidden Column Secret", "Hidden Cell", "Target:", "../media/table.png"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden worksheet content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXMarkdownCombinesInlineRichTextRuns(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Rich Text" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>First </t></r><r><t>Second</t></r><rPh sb="0" eb="2"><t>Hidden Phonetic</t></rPh><r><t> Third</t></r></is></c><c r="B1" t="inlineStr"><is><t>Tail</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "rich-inline.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(res.Text, "First") || !strings.Contains(res.Text, "Second") || !strings.Contains(res.Text, "Third") {
		t.Fatalf("text missing rich inline runs: %q", res.Text)
	}
	if !strings.Contains(md, "| First Second Third | Tail |") {
		t.Fatalf("markdown did not combine rich inline runs in one cell:\n%s", md)
	}
	if strings.Contains(res.Text, "Hidden Phonetic") || strings.Contains(md, "Hidden Phonetic") {
		t.Fatalf("phonetic helper text leaked: text=%q markdown=\n%s", res.Text, md)
	}
}

func TestXLSXCellSystemExtensionTextIsNotVisible(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Cell</t></is><extLst><ext uri="{hidden}"><t>Hidden System Cell Text</t><v>Hidden System Cell Value</v></ext></extLst></c><c r="B1" t="inlineStr"><is><t>Visible Tail</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "cell-system-extension.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Visible Tail"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX cell text %q in %q", want, res.Text)
		}
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX cell text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden System Cell Text", "Hidden System Cell Value"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("text kept XLSX cell system extension content %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept XLSX cell system extension content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXMarkdownTruncatesLargeCells(t *testing.T) {
	large := "Visible large cell " + strings.Repeat("A", maxMarkdownTableCellBytes*4)
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Large Cells" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/sharedStrings.xml", `<sst xmlns="urn:x"><si><t>Header</t></si><si><t>`+large+`</t></si></sst>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "large-markdown-cell.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible large cell") {
		t.Fatalf("markdown missing large cell prefix in:\n%s", md)
	}
	if strings.Contains(md, strings.Repeat("A", maxMarkdownTableCellBytes*2)) {
		t.Fatalf("markdown kept an oversized cell value, length=%d", len(md))
	}
}

func TestXLSXHyperlinkTargetsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Label</t></is></c><c r="B1" t="inlineStr"><is><t>Visible Link Cell</t></is></c><c r="C1" t="inlineStr"><is><t>Visible Tail</t></is></c><c r="D1" t="inlineStr"><is><t>Visible Cleaned Link Cell</t></is></c></row></sheetData><hyperlinks><hyperlink ref="B1" display="Visible Hyperlink Display" tooltip="Visible Hyperlink Tooltip" location="'Hidden Sheet'!A1"/><hyperlink ref="C1" display="C:\Users\me\hidden.png" tooltip="/xl/media/image1.png" location="xl/worksheets/sheet2.xml"/><hyperlink ref="D1" display="Visible Target: ../media/hyperlink.png display text" tooltip="Visible Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image tooltip text"/></hyperlinks></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hyperlink-targets.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Label", "Visible Link Cell", "Visible Tail", "Visible Cleaned Link Cell", "Visible Hyperlink Display", "Visible Hyperlink Tooltip", "Visible display text", "Visible tooltip text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible hyperlink text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"'Hidden Sheet'!A1", "xl/worksheets/sheet2.xml", "C:\\Users\\me\\hidden.png", "/xl/media/image1.png", "Target:", "../media/hyperlink.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden hyperlink target/resource %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Label", "Visible Link Cell", "Visible Tail", "Visible Cleaned Link Cell", "### Annotations", "Visible Hyperlink Display", "Visible Hyperlink Tooltip", "Visible display text", "Visible tooltip text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible hyperlink text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"'Hidden Sheet'!A1", "xl/worksheets/sheet2.xml", "C:\\Users\\me\\hidden.png", "/xl/media/image1.png", "Target:", "../media/hyperlink.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden hyperlink target/resource %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXFormulaTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Label</t></is></c><c r="B1" t="str"><f>HYPERLINK(&quot;https://example.test/internal&quot;,&quot;Hidden Formula Text&quot;)</f><v>Visible Cached Result</v></c></row><row r="2"><c r="A2"><f>SUM(A1:A10)</f><v>42</v></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "formula-text.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Label", "Visible Cached Result", "42"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible worksheet value %q in %q", want, res.Text)
		}
	}
	for _, bad := range []string{"HYPERLINK", "SUM(A1:A10)", "example.test/internal", "Hidden Formula Text"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept formula text %q in %q", bad, res.Text)
		}
	}
}

func TestXLSXPhoneticRunsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/sharedStrings.xml", `<sst xmlns="urn:x"><si><r><t>東京</t></r><rPh sb="0" eb="2"><t>Shared Phonetic Secret</t></rPh><phoneticPr fontId="1"/></si></sst>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><r><t>大阪</t></r><rPh sb="0" eb="2"><t>Inline Phonetic Secret</t></rPh><phoneticPr fontId="1"/></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "phonetic-runs.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "東京", "大阪"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Shared Phonetic Secret", "Inline Phonetic Secret", "fontId"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept phonetic helper text %q in %q", hidden, res.Text)
		}
	}
}

func TestXLSXDefinedNameFormulaValuesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><definedNames><definedName name="VisibleName Target: ../media/defined.png">Visible Defined Text rId77</definedName><definedName name="FormulaName ContentType: image/png">SUM(Table1[Amount])</definedName><definedName name="RangeName">Sheet1!$A$1:$B$2</definedName><definedName name="_xlnm.Print_Area">Sheet1!$A$1:$B$2</definedName></definedNames><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "defined-name-values.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"VisibleName", "Visible Defined Text", "FormulaName", "RangeName", "Visible Cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing expected workbook text %q in %q", want, res.Text)
		}
	}
	for _, bad := range []string{"Target:", "../media/defined.png", "rId77", "ContentType:", "image/png", "SUM(Table1[Amount])", "Sheet1!$A$1:$B$2", "_xlnm.Print_Area"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept defined-name formula/reference value %q in %q", bad, res.Text)
		}
	}
}

func TestXLSXHeaderFooterControlCodesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData><headerFooter><oddHeader>&amp;L&amp;&quot;Arial,Bold&quot;&amp;14&amp;OLeft Header Target: ../media/header.png&amp;CPage &amp;P of &amp;N long &amp;[Page] of &amp;[Pages]&amp;R&amp;HRight &amp;&amp; Header rId77</oddHeader><oddFooter>&amp;KFF0000&amp;BRed Footer ContentType: image/png&amp;G &amp;[Picture] &amp;OOutline Text &amp;HShadow Text PartName=/xl/media/footer.png &amp;K01+000Theme Color &amp;K03-123Tinted Color &amp;[Path]&amp;[File]&amp;[Tab]</oddFooter></headerFooter></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "header-footer-control-codes.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "Left Header", "Page of", "long of", "Right & Header", "Red Footer", "Outline Text", "Shadow Text", "Theme Color", "Tinted Color"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible header/footer text %q in %q", want, res.Text)
		}
	}
	for _, bad := range []string{"&L", "&C", "&R", "&P", "&N", "&G", "&O", "&H", "&[Page]", "&[Pages]", "&[Picture]", "&[Path]", "&[File]", "&[Tab]", "&KFF0000", "&K01+000", "&K03-123", "+000", "-123", "&B", "Arial,Bold", "&14", "Target:", "../media/header.png", "rId77", "ContentType:", "image/png", "PartName=", "/xl/media/footer.png"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept header/footer control code %q in %q", bad, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell", "### Headers and Footers", "Left Header", "Page of", "long of", "Right & Header", "Red Footer", "Outline Text", "Shadow Text", "Theme Color", "Tinted Color"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible header/footer text %q in:\n%s", want, md)
		}
	}
	for _, bad := range []string{"&L", "&C", "&R", "&P", "&N", "&G", "&O", "&H", "&[Page]", "&[Pages]", "&[Picture]", "&[Path]", "&[File]", "&[Tab]", "&KFF0000", "&K01+000", "&K03-123", "+000", "-123", "&B", "Arial,Bold", "&14", "Target:", "../media/header.png", "rId77", "ContentType:", "image/png", "PartName=", "/xl/media/footer.png"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept header/footer control code %q in:\n%s", bad, md)
		}
	}
}

func TestXLSXSlicerInternalUniqueNamesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/></Relationships>`)
	addZip(t, zw, "xl/slicerCaches/slicerCache1.xml", `<slicerCacheDefinition xmlns="urn:x" name="VisibleSlicerCacheName" sourceName="Visible Slicer Source"><level uniqueName="[Cube].[Region]" caption="Visible Slicer Level"><item uniqueName="[Cube].[Region].&amp;[North]" caption="Visible Slicer North"/><i n="[Cube].[Region].&amp;[South]" c="1" v="Visible Slicer South"/></level></slicerCacheDefinition>`)
	addZip(t, zw, "xl/slicers/_rels/slicer1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="../slicerCaches/slicerCache1.xml"/></Relationships>`)
	addZip(t, zw, "xl/slicers/slicer1.xml", `<slicer xmlns="urn:x" name="VisibleSlicerName" caption="Visible Slicer Caption"/>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "slicer-unique-names.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "VisibleSlicerCacheName", "Visible Slicer Source", "Visible Slicer Level", "Visible Slicer North", "Visible Slicer South", "VisibleSlicerName", "Visible Slicer Caption"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible slicer text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"[Cube].[Region]", "[Cube].[Region].&[North]", "[Cube].[Region].&[South]"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept internal slicer unique name %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slicers", "VisibleSlicerCacheName", "Visible Slicer Source", "Visible Slicer Level", "Visible Slicer North", "Visible Slicer South", "VisibleSlicerName", "Visible Slicer Caption"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible slicer text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"[Cube].[Region]", "[Cube].[Region].&[North]", "[Cube].[Region].&[South]"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept internal slicer unique name %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXHiddenSheetPivotAndSlicerTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/><sheet name="Hidden Sheet" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Hidden Cell Secret</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/><Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/><Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer2.xml"/></Relationships>`)
	addZip(t, zw, "xl/pivotTables/pivotTable1.xml", `<pivotTableDefinition xmlns="urn:x" name="VisiblePivotName" dataCaption="Visible Pivot Caption"><pivotFields><pivotField name="Visible Pivot Field"/></pivotFields></pivotTableDefinition>`)
	addZip(t, zw, "xl/pivotTables/_rels/pivotTable1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`)
	addZip(t, zw, "xl/pivotCache/pivotCacheDefinition1.xml", `<pivotCacheDefinition xmlns="urn:x"><cacheFields><cacheField name="Visible Cache Field"><sharedItems><s v="Visible Shared Item"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`)
	addZip(t, zw, "xl/pivotTables/pivotTable2.xml", `<pivotTableDefinition xmlns="urn:x" name="HiddenPivotSecret" dataCaption="Hidden Pivot Caption Secret"><pivotFields><pivotField name="Hidden Pivot Field Secret"/></pivotFields></pivotTableDefinition>`)
	addZip(t, zw, "xl/pivotTables/_rels/pivotTable2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition2.xml"/></Relationships>`)
	addZip(t, zw, "xl/pivotCache/pivotCacheDefinition2.xml", `<pivotCacheDefinition xmlns="urn:x"><cacheFields><cacheField name="Hidden Cache Field Secret"><sharedItems><s v="Hidden Shared Item Secret"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`)
	addZip(t, zw, "xl/slicers/slicer1.xml", `<slicer xmlns="urn:x" name="VisibleSlicerName" caption="Visible Slicer Caption"/>`)
	addZip(t, zw, "xl/slicers/_rels/slicer1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="../slicerCaches/slicerCache1.xml"/></Relationships>`)
	addZip(t, zw, "xl/slicerCaches/slicerCache1.xml", `<slicerCacheDefinition xmlns="urn:x" name="VisibleSlicerCacheName" sourceName="Visible Slicer Source"><level caption="Visible Slicer Level"><item caption="Visible Slicer Item"/></level></slicerCacheDefinition>`)
	addZip(t, zw, "xl/slicers/slicer2.xml", `<slicer xmlns="urn:x" name="HiddenSlicerSecret" caption="Hidden Slicer Caption Secret"/>`)
	addZip(t, zw, "xl/slicers/_rels/slicer2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="../slicerCaches/slicerCache2.xml"/></Relationships>`)
	addZip(t, zw, "xl/slicerCaches/slicerCache2.xml", `<slicerCacheDefinition xmlns="urn:x" name="HiddenSlicerCacheSecret" sourceName="Hidden Slicer Source Secret"><level caption="Hidden Slicer Level Secret"><item caption="Hidden Slicer Item Secret"/></level></slicerCacheDefinition>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheet-pivot-slicer.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell", "VisiblePivotName", "Visible Pivot Field", "Visible Cache Field", "Visible Shared Item", "VisibleSlicerName", "Visible Slicer Caption", "VisibleSlicerCacheName", "Visible Slicer Source", "Visible Slicer Item"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible pivot/slicer content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Sheet", "Hidden Cell Secret", "HiddenPivotSecret", "Hidden Pivot Caption Secret", "Hidden Pivot Field Secret", "Hidden Cache Field Secret", "Hidden Shared Item Secret", "HiddenSlicerSecret", "Hidden Slicer Caption Secret", "HiddenSlicerCacheSecret", "Hidden Slicer Source Secret", "Hidden Slicer Item Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden pivot/slicer content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Pivot Tables", "VisiblePivotName", "Visible Cache Field", "## Slicers", "VisibleSlicerName", "VisibleSlicerCacheName"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible pivot/slicer content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"HiddenPivotSecret", "Hidden Cache Field Secret", "HiddenSlicerSecret", "HiddenSlicerCacheSecret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden pivot/slicer content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXUnreferencedPivotAndSlicerTextIsNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/pivotTables/pivotTable1.xml", `<pivotTableDefinition xmlns="urn:x" name="InternalUnreferencedPivotSecret" dataCaption="Internal Pivot Caption Secret"><pivotFields><pivotField name="Internal Pivot Field Secret"/></pivotFields></pivotTableDefinition>`)
	addZip(t, zw, "xl/pivotCache/pivotCacheDefinition1.xml", `<pivotCacheDefinition xmlns="urn:x"><cacheFields><cacheField name="Internal Cache Field Secret"><sharedItems><s v="Internal Shared Item Secret"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`)
	addZip(t, zw, "xl/slicers/slicer1.xml", `<slicer xmlns="urn:x" name="InternalUnreferencedSlicerSecret" caption="Internal Slicer Caption Secret"/>`)
	addZip(t, zw, "xl/slicerCaches/slicerCache1.xml", `<slicerCacheDefinition xmlns="urn:x" name="InternalSlicerCacheSecret" sourceName="Internal Slicer Source Secret"><level caption="Internal Slicer Level Secret"><item caption="Internal Slicer Item Secret"/></level></slicerCacheDefinition>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-pivot-slicer.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet", "Visible Cell"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible XLSX content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"InternalUnreferencedPivotSecret", "Internal Pivot Caption Secret", "Internal Pivot Field Secret", "Internal Cache Field Secret", "Internal Shared Item Secret", "InternalUnreferencedSlicerSecret", "Internal Slicer Caption Secret", "InternalSlicerCacheSecret", "Internal Slicer Source Secret", "Internal Slicer Item Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced pivot/slicer content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Visible Sheet", "Visible Cell"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible XLSX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"InternalUnreferencedPivotSecret", "Internal Cache Field Secret", "InternalUnreferencedSlicerSecret", "InternalSlicerCacheSecret", "## Pivot Tables", "## Slicers"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced pivot/slicer content %q in:\n%s", hidden, md)
		}
	}
}

func TestXLSXPivotTextIsStructuredMarkdown(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Cell</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>`)
	addZip(t, zw, "xl/pivotTables/pivotTable1.xml", `<pivotTableDefinition xmlns="urn:x" name="VisiblePivotName Target: ../media/hidden-pivot.png" dataCaption="Visible Data Caption Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" missingCaption="Visible Missing Caption"><pivotFields><pivotField name="Visible Pivot Field ContentType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" caption="Visible Pivot Caption"/></pivotFields><dataFields><dataField name="Visible Data Field" caption="Visible Data Field Caption"/></dataFields></pivotTableDefinition>`)
	addZip(t, zw, "xl/pivotTables/_rels/pivotTable1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`)
	addZip(t, zw, "xl/pivotCache/pivotCacheDefinition1.xml", `<pivotCacheDefinition xmlns="urn:x"><cacheFields><cacheField name="Visible Cache Field PartName: /xl/worksheets/sheet1.xml" caption="Visible Cache Caption"><sharedItems><s v="Visible Shared Item r:embed: rId8"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "pivot-markdown.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"VisiblePivotName", "Visible Data Caption", "Visible Pivot Field", "Visible Cache Field", "Visible Shared Item"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible pivot text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Pivot Tables", "VisiblePivotName", "Visible Data Caption", "Visible Pivot Field", "Visible Cache Field", "Visible Shared Item"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible pivot text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Target:", "../media/hidden-pivot.png", "Type:", "relationships/image", "ContentType:", "application/vnd.openxmlformats", "PartName:", "/xl/worksheets/sheet1.xml", "r:embed:", "rId8"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept pivot internal metadata %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept pivot internal metadata %q in:\n%s", hidden, md)
		}
	}
}

func TestOOXMLChartFormulaReferencesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp><c:chart r:id="rIdChart"/></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c"><c:chart><c:title><c:tx><c:rich><a:p xmlns:a="urn:a"><a:r><a:t>Visible Chart Title</a:t></a:r></a:p></c:rich></c:tx></c:title><c:ser><c:cat><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f><c:strCache><c:pt idx="0"><c:v>Visible Category</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f><c:numCache><c:pt idx="0"><c:v>42</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:chart></c:chartSpace>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "chart-formula-refs.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Chart Title", "Visible Category", "42"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible chart text %q in %q", want, res.Text)
		}
	}
	for _, bad := range []string{"Sheet1!$A$1:$A$2", "Sheet1!$B$1:$B$2"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept chart formula reference %q in %q", bad, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Visible Slide Text", "## Drawings", "Visible Chart Title", "Visible Category", "42"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible chart text %q in:\n%s", want, md)
		}
	}
	for _, bad := range []string{"Sheet1!$A$1:$A$2", "Sheet1!$B$1:$B$2"} {
		if strings.Contains(md, bad) {
			t.Fatalf("markdown kept chart formula reference %q in:\n%s", bad, md)
		}
	}
}

func TestOOXMLAlternateContentUsesPreferredChoiceText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:mc="urn:mc"><w:body><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>Visible Preferred Text</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>Internal Fallback Text</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "alternate-content.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Preferred Text") {
		t.Fatalf("missing preferred AlternateContent text in %q", res.Text)
	}
	if strings.Contains(res.Text, "Internal Fallback Text") {
		t.Fatalf("kept fallback AlternateContent text in %q", res.Text)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible Preferred Text") {
		t.Fatalf("markdown missing preferred AlternateContent text:\n%s", md)
	}
	if strings.Contains(md, "Internal Fallback Text") {
		t.Fatalf("markdown kept fallback AlternateContent text:\n%s", md)
	}
}

func TestDOCXAlternateContentFallbackImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:mc="urn:mc" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><mc:AlternateContent><mc:Choice Requires="wps"><p:pic><p:nvPicPr><p:cNvPr id="1" descr="Preferred Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPreferred"/></p:blipFill></p:pic></mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="2" descr="Fallback Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdFallback"/></p:blipFill></p:pic></mc:Fallback></mc:AlternateContent></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPreferred" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/preferred.png"/><Relationship Id="rIdFallback" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fallback.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/preferred.png", testPNG())
	addZipBytes(t, zw, "word/media/fallback.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "alternate-content-image.docx")
	outDir := filepath.Join(dir, "images")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "preferred.png" || res.Images[0].Alt != "Preferred Picture" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only valid preferred image, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "fallback.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("fallback AlternateContent image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Preferred Picture](images/preferred.png)") {
		t.Fatalf("markdown missing preferred image:\n%s", md)
	}
	if strings.Contains(md, "Fallback Picture") || strings.Contains(md, "fallback.jpg") {
		t.Fatalf("markdown kept fallback image content:\n%s", md)
	}
}

func TestVMLClientDataIsNotVisibleText(t *testing.T) {
	text, err := visibleVMLText([]byte(`<v:shape xmlns:v="urn:v" xmlns:x="urn:x"><v:textbox><div>Visible VML Text</div></v:textbox><x:ClientData ObjectType="Checkbox"><x:Anchor>10, 50, 24, 5</x:Anchor><x:Locked>False</x:Locked><x:FmlaMacro>[0]!CheckBox363_Click</x:FmlaMacro><x:TextVAlign>Center</x:TextVAlign><x:FmlaLink>$K$25</x:FmlaLink></x:ClientData></v:shape>`))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(text, "Visible VML Text") {
		t.Fatalf("missing visible VML text in %q", text)
	}
	for _, bad := range []string{"10, 50, 24, 5", "False", "[0]!CheckBox363_Click", "Center", "$K$25"} {
		if strings.Contains(text, bad) {
			t.Fatalf("kept VML ClientData text %q in %q", bad, text)
		}
	}
}

func TestPPTXMasterAndLayoutTextAreNotDefaultVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide1.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Notes Target: ../media/notes.png Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/comments/comment1.xml", `<p:cmLst xmlns:p="urn:p" xmlns:a="urn:a"><p:cm><p:text>Visible Comment Text</p:text></p:cm></p:cmLst>`)
	addZip(t, zw, "ppt/slideMasters/slideMaster1.xml", `<p:sldMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Internal Master Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldMaster>`)
	addZip(t, zw, "ppt/slideLayouts/slideLayout1.xml", `<p:sldLayout xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Internal Layout Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`)
	addZip(t, zw, "ppt/notesMasters/notesMaster1.xml", `<p:notesMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Internal Notes Master Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notesMaster>`)
	addZip(t, zw, "ppt/handoutMasters/handoutMaster1.xml", `<p:handoutMaster xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Internal Handout Master Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:handoutMaster>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "visible-pptx-only.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Notes Text", "Visible Comment Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Internal Master Text", "Internal Layout Text", "Internal Notes Master Text", "Internal Handout Master Text"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept PPTX internal template text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Visible Slide Text", "## Notes", "Visible Notes Text", "## Comments", "Visible Comment Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible PPTX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Internal Master Text", "Internal Layout Text", "Internal Notes Master Text", "Internal Handout Master Text"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept PPTX internal template text %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXHiddenSlidesAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hidden Slide Secret</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide3.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" show="false"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>False Hidden Slide Secret</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-slides.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Slide Text") {
		t.Fatalf("missing visible slide text in %q", res.Text)
	}
	for _, hidden := range []string{"Hidden Slide Secret", "False Hidden Slide Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden slide text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Slide 1") || !strings.Contains(md, "Visible Slide Text") {
		t.Fatalf("markdown missing visible slide section in:\n%s", md)
	}
	for _, hidden := range []string{"Hidden Slide Secret", "False Hidden Slide Secret", "## Slide 2", "## Slide 3"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden slide content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXUnreferencedSlidesAreNotVisibleContent(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/presentation.xml", `<p:presentation xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst></p:presentation>`)
	addZip(t, zw, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Referenced Slide</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible referenced picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisibleImage"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisibleImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/></Relationships>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Orphan Slide Secret</a:t></a:r></a:p></p:txBody></p:sp><c:chart r:id="rIdChart"/><p:pic><p:nvPicPr><p:cNvPr id="2" descr="Orphan picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdOrphanImage"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/><Relationship Id="rIdOrphanImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/orphan.jpg"/><Relationship Id="rIdComment" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Orphan Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/comments/comment2.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>Orphan Comment Secret</p:text></p:cm></p:cmLst>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide2.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Orphan Notes Secret</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/visible.png", testPNG())
	addZipBytes(t, zw, "ppt/media/orphan.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-slide.pptx")
	outDir := filepath.Join(dir, "images")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Referenced Slide") {
		t.Fatalf("missing referenced slide text in %q", res.Text)
	}
	for _, hidden := range []string{"Orphan Slide Secret", "Orphan Chart Secret", "Orphan Comment Secret", "Orphan Notes Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced slide content %q in %q", hidden, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible.png" || res.Images[0].Alt != "Visible referenced picture" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only referenced slide image, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "orphan.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("orphan slide image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "## Slide 1") || !strings.Contains(md, "Visible Referenced Slide") || !strings.Contains(md, "![Visible referenced picture](images/visible.png)") {
		t.Fatalf("markdown missing referenced slide content:\n%s", md)
	}
	for _, hidden := range []string{"## Slide 2", "Orphan Slide Secret", "Orphan Chart Secret", "Orphan Comment Secret", "Orphan Notes Secret", "orphan.jpg", "Orphan picture"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced slide content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXHiddenSlideRelatedTextPartsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hidden Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rIdDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/><Relationship Id="rIdDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Visible Chart Title</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Hidden Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Visible Drawing Description"/></xdr:wsDr>`)
	addZip(t, zw, "ppt/drawings/drawing2.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Hidden Drawing Secret"/></xdr:wsDr>`)
	addZip(t, zw, "ppt/diagrams/data1.xml", `<dgm:dataModel xmlns:dgm="urn:dgm" xmlns:a="urn:a"><dgm:pt><dgm:t><a:p><a:r><a:t>Visible SmartArt Text</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:dataModel>`)
	addZip(t, zw, "ppt/diagrams/data2.xml", `<dgm:dataModel xmlns:dgm="urn:dgm" xmlns:a="urn:a"><dgm:pt><dgm:t><a:p><a:r><a:t>Hidden SmartArt Secret</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:dataModel>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-slide-related-text.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Chart Title", "Visible Drawing Description", "Visible SmartArt Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX related text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Slide Text", "Hidden Chart Secret", "Hidden Drawing Secret", "Hidden SmartArt Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden PPTX related text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Visible Slide Text", "## Drawings", "Visible Chart Title", "Visible Drawing Description", "Visible SmartArt Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible PPTX related content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Slide Text", "Hidden Chart Secret", "Hidden Drawing Secret", "Hidden SmartArt Secret", "## Slide 2"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden PPTX related content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXHiddenShapeRelatedTextPartsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp>
<p:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rIdVisibleChart"/></a:graphicData></a:graphic></p:graphicFrame>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Hidden Chart Frame" hidden="1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdHiddenChart"/></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisibleChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdHiddenChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/><Relationship Id="rIdUnreferenced" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/></Relationships>`)
	addZip(t, zw, "ppt/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Visible Chart Text</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Hidden Shape Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/charts/chart3.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Unreferenced Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-shape-related-text.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Chart Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX related text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Shape Chart Secret", "Unreferenced Chart Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden PPTX related text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Visible Slide Text", "## Drawings", "Visible Chart Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible PPTX related content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Shape Chart Secret", "Unreferenced Chart Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden PPTX related content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXUnreferencedRelatedTextPartsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp><c:chart r:id="rIdChart"/></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Visible Chart Text</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/charts/chart2.xml", `<c:chartSpace xmlns:c="urn:c" xmlns:a="urn:a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Internal Unreferenced Chart Secret</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`)
	addZip(t, zw, "ppt/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Internal Unreferenced Drawing Secret"/></xdr:wsDr>`)
	addZip(t, zw, "ppt/diagrams/data1.xml", `<dgm:dataModel xmlns:dgm="urn:dgm" xmlns:a="urn:a"><dgm:pt><dgm:t><a:p><a:r><a:t>Internal Unreferenced SmartArt Secret</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:dataModel>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-pptx-related-text.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Chart Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX related text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced Chart Secret", "Internal Unreferenced Drawing Secret", "Internal Unreferenced SmartArt Secret"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept unreferenced PPTX related text %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Visible Slide Text", "## Drawings", "Visible Chart Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible PPTX related content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Internal Unreferenced Chart Secret", "Internal Unreferenced Drawing Secret", "Internal Unreferenced SmartArt Secret"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept unreferenced PPTX related content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXHiddenSlideNotesAndCommentsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hidden Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide1.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Notes Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide2.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hidden Notes Secret</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide3.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Unreferenced Notes Secret</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/comments/comment1.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>Visible Comment Type: http://schemas.openxmlformats.org/officeDocument/2006/relationships/image Text</p:text></p:cm></p:cmLst>`)
	addZip(t, zw, "ppt/comments/comment2.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>Hidden Comment Secret</p:text></p:cm></p:cmLst>`)
	addZip(t, zw, "ppt/comments/comment3.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>Unreferenced Comment Secret</p:text></p:cm></p:cmLst>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-slide-notes-comments.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Notes Text", "Visible Comment Text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Slide Text", "Hidden Notes Secret", "Unreferenced Notes Secret", "Hidden Comment Secret", "Unreferenced Comment Secret", "Target:", "../media/notes.png", "relationships/image"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden PPTX content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"Visible Slide Text", "Visible Notes Text", "Visible Comment Text"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible PPTX content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"Hidden Slide Text", "Hidden Notes Secret", "Unreferenced Notes Secret", "Hidden Comment Secret", "Unreferenced Comment Secret", "Target:", "../media/notes.png", "relationships/image"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept hidden PPTX content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXNotesMarkdownIsGroupedBySlide(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/presentation.xml", `<p:presentation xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>`)
	addZip(t, zw, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First Slide Body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second Slide Body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide1.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First Slide Notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide2.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second Slide Notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/></Relationships>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "notes-grouped-by-slide.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "First Slide Body", "## Slide 2", "Second Slide Body", "## Notes", "### Slide 1", "First Slide Notes", "### Slide 2", "Second Slide Notes"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing grouped PPTX notes content %q in:\n%s", want, md)
		}
	}
	if strings.Index(md, "### Slide 1") > strings.Index(md, "First Slide Notes") || strings.Index(md, "### Slide 2") > strings.Index(md, "Second Slide Notes") {
		t.Fatalf("markdown notes are not grouped under slide headings:\n%s", md)
	}
}

func TestPPTXCommentsMarkdownIsGroupedBySlide(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/presentation.xml", `<p:presentation xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>`)
	addZip(t, zw, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First Slide Body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second Slide Body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment2.xml"/></Relationships>`)
	addZip(t, zw, "ppt/comments/comment1.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>First Slide Comment</p:text></p:cm></p:cmLst>`)
	addZip(t, zw, "ppt/comments/comment2.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>Second Slide Comment</p:text></p:cm></p:cmLst>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "comments-grouped-by-slide.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "First Slide Body", "## Slide 2", "Second Slide Body", "## Comments", "### Slide 1", "First Slide Comment", "### Slide 2", "Second Slide Comment"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing grouped PPTX comments content %q in:\n%s", want, md)
		}
	}
	if strings.Index(md, "### Slide 1") > strings.Index(md, "First Slide Comment") || strings.Index(md, "### Slide 2") > strings.Index(md, "Second Slide Comment") {
		t.Fatalf("markdown comments are not grouped under slide headings:\n%s", md)
	}
}

func TestPPTXMarkdownCleansMalformedSlidePartTitles(t *testing.T) {
	const slideName = "ppt/slides/slide ContentType=image%2Fpng rId77.xml"
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/presentation.xml", `<p:presentation xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>`)
	addZip(t, zw, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Target="slides/slide ContentType=image%2Fpng rId77.xml"/></Relationships>`)
	addZip(t, zw, slideName, `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible malformed slide body</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide ContentType=image%2Fpng rId77.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide1.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible malformed slide notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide ContentType=image%2Fpng rId77.xml"/></Relationships>`)
	addZip(t, zw, "ppt/comments/comment1.xml", `<p:cmLst xmlns:p="urn:p"><p:cm><p:text>Visible malformed slide comment</p:text></p:cm></p:cmLst>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "malformed-slide-title.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide", "Visible malformed slide body", "## Notes", "### Slide", "Visible malformed slide notes", "## Comments", "Visible malformed slide comment"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing malformed slide visible content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"ContentType", "image/png", "image%2Fpng", "rId77"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("text kept malformed slide title internals %q in %q", hidden, res.Text)
		}
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept malformed slide title internals %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXNotesSystemPlaceholdersAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/notesSlides/notesSlide1.xml", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="1" title="Visible Notes Shape Title"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Visible Speaker Notes</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="2" title="Internal Date Placeholder Title"/><p:nvPr><p:ph type="dt"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>2026-06-25</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" descr="Internal Footer Placeholder Description"/><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Confidential Footer</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="4" title="Internal Slide Number Placeholder"/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>12</a:t></a:r></a:p></p:txBody></p:sp>
<p:pic><p:nvPicPr><p:cNvPr id="5" descr="Internal Slide Image Placeholder"/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvPicPr></p:pic>
</p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, "ppt/notesSlides/_rels/notesSlide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/></Relationships>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "notes-system-placeholders.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Slide Text", "Visible Speaker Notes", "Visible Notes Shape Title"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible PPTX notes content %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"2026-06-25", "Confidential Footer", "Internal Date Placeholder Title", "Internal Footer Placeholder Description", "Internal Slide Number Placeholder", "Internal Slide Image Placeholder"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept PPTX notes placeholder content %q in %q", hidden, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Visible Slide Text", "## Notes", "Visible Speaker Notes", "Visible Notes Shape Title"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible PPTX notes content %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"2026-06-25", "Confidential Footer", "Internal Date Placeholder Title", "Internal Footer Placeholder Description", "Internal Slide Number Placeholder", "Internal Slide Image Placeholder"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept PPTX notes placeholder content %q in:\n%s", hidden, md)
		}
	}
}

func TestPPTXHiddenSlideImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hidden Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/></Relationships>`)
	addZip(t, zw, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/visible.png", testPNG())
	addZipBytes(t, zw, "ppt/media/hidden.png", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-slide-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only visible slide image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected visible slide PNG only, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden slide image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![visible](images/visible.png)") {
		t.Fatalf("markdown missing visible image reference in:\n%s", md)
	}
	if strings.Contains(md, "hidden") {
		t.Fatalf("markdown kept hidden slide image reference in:\n%s", md)
	}
}

func TestPPTXHiddenOnlySlideImagesDoNotTriggerMediaFallback(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide Without Images</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/slide2.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" show="0"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="2" descr="Hidden Only Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden-only.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/hidden-only.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-only-slide-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 0 {
		t.Fatalf("expected hidden-only slide image to be skipped, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden-only.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden-only slide image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if strings.Contains(md, "hidden-only") || strings.Contains(md, "Hidden Only Picture") {
		t.Fatalf("markdown kept hidden-only slide image/text in:\n%s", md)
	}
}

func TestPPTXVisibleSlideTransitiveMediaIsExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/charts/chart1.xml", `<c:chartSpace xmlns:c="urn:c"><c:chart/></c:chartSpace>`)
	addZip(t, zw, "ppt/charts/_rels/chart1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart-bg.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/chart-bg.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "visible-transitive-media.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "chart-bg.png" {
		t.Fatalf("expected transitive visible chart media, got %#v", res.Images)
	}
}

func TestStrictPPTXImagesExcludeOLEPreviewAndLayoutMedia(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/presentation.xml", `<p:presentation xmlns:p="urn:p" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>`)
	addZip(t, zw, "ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="1" name="Visible"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic><p:oleObj r:id="rIdOLE"><p:pic><p:blipFill><a:blip r:embed="rIdPreview"/></p:blipFill></p:pic></p:oleObj></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/><Relationship Id="rIdPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/ole-preview.png"/><Relationship Id="rIdOLE" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/object.bin"/><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`)
	addZip(t, zw, "ppt/slideLayouts/slideLayout1.xml", `<p:sldLayout xmlns:p="urn:p"/>`)
	addZip(t, zw, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayoutImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/visible.png", testPNG())
	addZipBytes(t, zw, "ppt/media/ole-preview.png", testPNG())
	addZipBytes(t, zw, "ppt/media/layout.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	r, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	files := map[string]*zip.File{}
	for _, f := range r.File {
		files[f.Name] = f
	}
	media, ok := strictPptxVisibleMediaParts(files)
	if !ok || len(media) != 1 || !media["ppt/media/visible.png"] {
		t.Fatalf("strict PowerPoint media = %#v, ok=%v; want visible picture only", media, ok)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "strict-images.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible.png" {
		t.Fatalf("strict Extract images = %#v, want visible.png only", res.Images)
	}
	compatible, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(compatible.Images) < len(res.Images) {
		t.Fatalf("default compatibility extraction unexpectedly lost images: strict=%d compatible=%d", len(res.Images), len(compatible.Images))
	}
}

func TestStrictPPTXWebSampleMatchesPowerPointPictureCount(t *testing.T) {
	file := filepath.Join("testdata", "web-samples", "samples", "pptx", "00020011.pptx")
	res, err := Extract(file, Options{StrictOfficeImages: true})
	if err != nil {
		t.Fatal(err)
	}
	// Office COM exposes four msoPicture/msoLinkedPicture shapes for this
	// presentation. The package also contains layout artwork and an OLE preview,
	// which are intentionally excluded by StrictOfficeImages.
	if len(res.Images) != 4 {
		t.Fatalf("strict PPTX image count = %d, want 4 (PowerPoint Picture Shapes)", len(res.Images))
	}
}

func TestPPTXHiddenShapeImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:pic><p:nvPicPr><p:cNvPr id="1" name="Visible Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Hidden Picture" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/><Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/visible.png", testPNG())
	addZipBytes(t, zw, "ppt/media/hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-shape-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only visible shape image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected visible shape PNG only, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden shape image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![visible](images/visible.png)") {
		t.Fatalf("markdown missing visible shape image reference in:\n%s", md)
	}
	if strings.Contains(md, "hidden") || strings.Contains(md, "Hidden Picture") {
		t.Fatalf("markdown kept hidden shape image/text in:\n%s", md)
	}
}

func TestPPTXUnreferencedMediaImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:txBody><a:p><a:r><a:t>Visible Slide</a:t></a:r></a:p></p:txBody></p:sp>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible PPTX Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.jpg"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/visible.jpg", testPNG())
	addZipBytes(t, zw, "ppt/media/internal.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-pptx-media.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only referenced PPTX image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected mislabelled referenced PPTX PNG to be written with sniffed extension, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("mislabelled PPTX .jpg should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "internal.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("unreferenced PPTX image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible PPTX Picture](images/visible.png)") {
		t.Fatalf("markdown missing referenced PPTX image in:\n%s", md)
	}
	slidePos := strings.Index(md, "Visible Slide")
	imagePos := strings.Index(md, "![Visible PPTX Picture](images/visible.png)")
	if !(slidePos >= 0 && imagePos > slidePos) {
		t.Fatalf("markdown did not place PPTX image after visible slide text:\n%s", md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("markdown duplicated placed PPTX image in trailing image section:\n%s", md)
	}
	if strings.Contains(md, "internal") {
		t.Fatalf("markdown kept unreferenced PPTX image in:\n%s", md)
	}
}

func TestPPTXSharedHiddenAndVisibleImageIsKept(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Hidden Shared Picture" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdShared"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" descr="Visible Shared Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdShared"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/shared.png"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/shared.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "shared-shape-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "shared.png" {
		t.Fatalf("expected shared visible PPTX image to be kept, got %#v", res.Images)
	}
	if res.Images[0].Alt != "Visible Shared Picture" {
		t.Fatalf("expected visible shared PPTX image alt, got %#v", res.Images[0])
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible Shared Picture](images/shared.png)") {
		t.Fatalf("markdown missing visible shared PPTX image alt:\n%s", md)
	}
	if strings.Contains(md, "Hidden Shared Picture") {
		t.Fatalf("markdown kept hidden shared PPTX image alt:\n%s", md)
	}
}

func TestXLSXHiddenSheetImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/><sheet name="Hidden" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Hidden Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Visible Drawing"/></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing2.xml", `<xdr:wsDr xmlns:xdr="urn:xdr"><xdr:cNvPr descr="Hidden Drawing"/></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/visible.png", testPNG())
	addZipBytes(t, zw, "xl/media/hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-sheet-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only visible sheet image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected visible sheet PNG only, got %#v", res.Images[0])
	}
	if strings.Contains(res.Text, "Hidden Sheet Text") {
		t.Fatalf("kept hidden sheet text in %q", res.Text)
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden sheet image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible Drawing](images/visible.png)") {
		t.Fatalf("markdown missing visible sheet image reference in:\n%s", md)
	}
	if strings.Contains(md, "hidden") || strings.Contains(md, "Hidden") {
		t.Fatalf("markdown kept hidden sheet image/text in:\n%s", md)
	}
}

func TestXLSXHiddenDrawingImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Visible Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdVisible"/></xdr:blipFill></xdr:pic>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Hidden Picture" hidden="1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdHidden"/></xdr:blipFill></xdr:pic>
</xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/><Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/visible.png", testPNG())
	addZipBytes(t, zw, "xl/media/hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-drawing-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only visible XLSX drawing image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected visible XLSX drawing PNG only, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden XLSX drawing image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![visible](images/visible.png)") {
		t.Fatalf("markdown missing visible drawing image reference in:\n%s", md)
	}
	if strings.Contains(md, "hidden") || strings.Contains(md, "Hidden Picture") {
		t.Fatalf("markdown kept hidden drawing image/text in:\n%s", md)
	}
}

func TestXLSXHiddenOnlySheetImagesDoNotTriggerMediaFallback(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/><sheet name="Hidden" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Without Images</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Hidden Sheet Text</t></is></c></row></sheetData><drawing r:id="rIdHiddenDrawing"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHiddenDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/hiddenDrawing.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/hiddenDrawing.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" descr="Hidden Only Sheet Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdHidden"/></xdr:blipFill></xdr:pic></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/hiddenDrawing.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/hidden-only.png"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/hidden-only.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-only-sheet-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 0 {
		t.Fatalf("expected hidden-only sheet image to be skipped, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden-only.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden-only sheet image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if strings.Contains(md, "hidden-only") || strings.Contains(md, "Hidden Only Sheet Picture") || strings.Contains(md, "Hidden Sheet Text") {
		t.Fatalf("markdown kept hidden-only sheet content in:\n%s", md)
	}
}

func TestXLSXUnreferencedSheetImagesDoNotTriggerMediaFallback(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Without Images</t></is></c></row></sheetData></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Orphan Sheet Text</t></is></c></row></sheetData><drawing r:id="rIdOrphanDrawing"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOrphanDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/orphanDrawing.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/orphanDrawing.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" descr="Orphan Sheet Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdOrphan"/></xdr:blipFill></xdr:pic></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/orphanDrawing.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOrphan" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/orphan-only.png"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/orphan-only.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "orphan-only-sheet-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible Sheet Without Images") {
		t.Fatalf("missing visible sheet text in %q", res.Text)
	}
	if strings.Contains(res.Text, "Orphan Sheet Text") {
		t.Fatalf("kept unreferenced sheet text in %q", res.Text)
	}
	if len(res.Images) != 0 {
		t.Fatalf("expected orphan-only sheet image to be skipped, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "orphan-only.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("orphan-only sheet image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "Visible Sheet Without Images") {
		t.Fatalf("markdown missing visible sheet text:\n%s", md)
	}
	if strings.Contains(md, "orphan-only") || strings.Contains(md, "Orphan Sheet Picture") || strings.Contains(md, "Orphan Sheet Text") {
		t.Fatalf("markdown kept unreferenced sheet content in:\n%s", md)
	}
}

func TestXLSXUnreferencedMediaImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" descr="Visible XLSX Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdVisible"/></xdr:blipFill></xdr:pic></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/visible.png"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/visible.png", testJPEG())
	addZipBytes(t, zw, "xl/media/internal.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-xlsx-media.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only referenced XLSX image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.jpg" || res.Images[0].Ext != ".jpg" || !validImageData(".jpg", res.Images[0].Data) {
		t.Fatalf("expected mislabelled referenced XLSX JPEG to be written with sniffed extension, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.jpg")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("mislabelled XLSX .png should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "internal.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("unreferenced XLSX image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible XLSX Picture](images/visible.jpg)") {
		t.Fatalf("markdown missing referenced XLSX image in:\n%s", md)
	}
	sheetPos := strings.Index(md, "Visible Sheet Text")
	imagePos := strings.Index(md, "![Visible XLSX Picture](images/visible.jpg)")
	if !(sheetPos >= 0 && imagePos > sheetPos) {
		t.Fatalf("markdown did not place XLSX image after visible sheet text:\n%s", md)
	}
	if strings.Contains(md, "## Images") {
		t.Fatalf("markdown duplicated placed XLSX image in trailing image section:\n%s", md)
	}
	if strings.Contains(md, "internal") {
		t.Fatalf("markdown kept unreferenced XLSX image in:\n%s", md)
	}
}

func TestXLSXSharedHiddenAndVisibleImageIsKept(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" descr="Hidden Shared Picture" hidden="1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdShared"/></xdr:blipFill></xdr:pic>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" descr="Visible Shared Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdShared"/></xdr:blipFill></xdr:pic>
</xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/shared.png"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/shared.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "shared-drawing-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "shared.png" {
		t.Fatalf("expected shared visible XLSX image to be kept, got %#v", res.Images)
	}
	if res.Images[0].Alt != "Visible Shared Picture" {
		t.Fatalf("expected visible shared XLSX image alt, got %#v", res.Images[0])
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible Shared Picture](images/shared.png)") {
		t.Fatalf("markdown missing visible shared XLSX image alt:\n%s", md)
	}
	if strings.Contains(md, "Hidden Shared Picture") {
		t.Fatalf("markdown kept hidden shared XLSX image alt:\n%s", md)
	}
}

func TestXLSXDrawingPartSharedByHiddenAndVisibleSheetsIsKept(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/><sheet name="Hidden" state="hidden" r:id="rId2"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Visible Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/sheet2.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="inlineStr"><is><t>Hidden Sheet Text</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/sharedDrawing.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/sharedDrawing.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/sharedDrawing.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Shared Visible Drawing" descr="Shared Visible Drawing Description"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdShared"/></xdr:blipFill></xdr:pic></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/sharedDrawing.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/shared.png"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/shared.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "shared-sheet-drawing.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Sheet Text", "Shared Visible Drawing Description"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible shared drawing text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Hidden Sheet Text") {
		t.Fatalf("kept hidden sheet text in %q", res.Text)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "shared.png" {
		t.Fatalf("expected shared visible sheet drawing image to be kept, got %#v", res.Images)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Shared Visible Drawing Description](images/shared.png)") {
		t.Fatalf("markdown missing shared visible drawing image reference in:\n%s", md)
	}
	for _, want := range []string{"## Drawings", "Shared Visible Drawing", "Shared Visible Drawing Description"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing shared visible drawing text %q in:\n%s", want, md)
		}
	}
	if strings.Contains(md, "Hidden Sheet Text") {
		t.Fatalf("markdown kept hidden sheet text in:\n%s", md)
	}
}

func TestDOCXHiddenDrawingImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Visible Body</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" name="Visible Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Hidden Picture" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/visible.png"/><Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible.png", testPNG())
	addZipBytes(t, zw, "word/media/hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-docx-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only visible DOCX image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" {
		t.Fatalf("expected visible PNG only, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden DOCX image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![visible](images/visible.png)") {
		t.Fatalf("markdown missing visible image reference in:\n%s", md)
	}
	if strings.Contains(md, "hidden") || strings.Contains(md, "Hidden Picture") {
		t.Fatalf("markdown kept hidden DOCX image/text in:\n%s", md)
	}
}

func TestDOCXHiddenVMLShapeImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>
<w:r><w:t>Visible VML wrapper</w:t></w:r>
<w:pict><v:shape id="visible" title="Visible VML Picture"><v:imagedata r:id="rIdVisible"/></v:shape></w:pict>
<w:pict><v:shape id="hidden" title="Hidden VML Picture" style="visibility:hidden"><v:imagedata r:id="rIdHidden"/></v:shape></w:pict>
</w:p></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/visible-vml.png"/><Relationship Id="rIdHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/hidden-vml.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible-vml.png", testPNG())
	addZipBytes(t, zw, "word/media/hidden-vml.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-vml-shape-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only visible VML image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible-vml.png" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected valid visible VML PNG only, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible-vml.png")); err != nil {
		t.Fatalf("visible VML image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "hidden-vml.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden VML image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "visible-vml.png") {
		t.Fatalf("markdown missing visible VML image reference in:\n%s", md)
	}
	if strings.Contains(md, "hidden-vml") || strings.Contains(md, "Hidden VML Picture") {
		t.Fatalf("markdown kept hidden VML image content in:\n%s", md)
	}
}

func TestDOCXUnreferencedMediaImagesAreNotExtracted(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>Visible Body</w:t></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/visible.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible.jpg", testPNG())
	addZipBytes(t, zw, "word/media/internal.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "unreferenced-docx-media.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 {
		t.Fatalf("expected only referenced DOCX image, got %#v", res.Images)
	}
	if res.Images[0].Name != "visible.png" || res.Images[0].Ext != ".png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected mislabelled referenced DOCX PNG to be written with sniffed extension, got %#v", res.Images[0])
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("mislabelled DOCX .jpg should not be written, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "internal.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("unreferenced DOCX image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible Picture](images/visible.png)") {
		t.Fatalf("markdown missing referenced image in:\n%s", md)
	}
	if strings.Contains(md, "internal") {
		t.Fatalf("markdown kept unreferenced image in:\n%s", md)
	}
}

func TestDOCXSharedHiddenAndVisibleImageIsKept(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:rPr><w:vanish/></w:rPr><p:pic><p:nvPicPr><p:cNvPr id="2" descr="Hidden Shared Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdShared"/></p:blipFill></p:pic></w:r></w:p>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible Shared Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdShared"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/></Relationships>`)
	addZipBytes(t, zw, "word/media/shared.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "shared-docx-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "shared.png" {
		t.Fatalf("expected shared visible image to be kept, got %#v", res.Images)
	}
	if res.Images[0].Alt != "Visible Shared Picture" {
		t.Fatalf("expected visible shared image alt, got %#v", res.Images[0])
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible Shared Picture](images/shared.png)") {
		t.Fatalf("markdown missing visible shared image alt:\n%s", md)
	}
	if strings.Contains(md, "Hidden Shared Picture") {
		t.Fatalf("markdown kept hidden shared image alt:\n%s", md)
	}
}

func TestDOCXNonRIDImageRelationshipsAreResolved(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible custom relationship image"/></p:nvPicPr><p:blipFill><a:blip r:embed="imgVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" descr="Hidden custom relationship image" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="imgHidden"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="imgVisible" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/custom-visible.png"/><Relationship Id="imgHidden" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/custom-hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/custom-visible.png", testPNG())
	addZipBytes(t, zw, "word/media/custom-hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "custom-rel-id-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "custom-visible.png" || res.Images[0].Alt != "Visible custom relationship image" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only visible custom relationship image with alt, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "custom-visible.png")); err != nil {
		t.Fatalf("visible custom relationship image not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "custom-hidden.jpg")); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("hidden custom relationship image was written or stat failed unexpectedly: %v", err)
	}
	md := res.Markdown("images")
	if !strings.Contains(md, "![Visible custom relationship image](images/custom-visible.png)") {
		t.Fatalf("markdown missing visible custom relationship image alt/reference:\n%s", md)
	}
	if strings.Contains(md, "Hidden custom relationship image") || strings.Contains(md, "custom-hidden") {
		t.Fatalf("markdown kept hidden custom relationship image content:\n%s", md)
	}
}

func TestOOXMLStrayNestedMediaDirectoriesAreNotContentImages(t *testing.T) {
	cases := []struct {
		name     string
		bodyPart string
		bodyXML  string
		stray    string
		wantText string
	}{
		{
			name:     "stray-nested-media.docx",
			bodyPart: "word/document.xml",
			bodyXML:  `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Visible DOCX body</w:t></w:r></w:p></w:body></w:document>`,
			stray:    "custom/word/media/stray.png",
			wantText: "Visible DOCX body",
		},
		{
			name:     "stray-nested-media.pptx",
			bodyPart: "ppt/slides/slide1.xml",
			bodyXML:  `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Visible PPTX slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
			stray:    "custom/ppt/media/stray.png",
			wantText: "Visible PPTX slide",
		},
		{
			name:     "stray-nested-media.xlsx",
			bodyPart: "xl/workbook.xml",
			bodyXML:  `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible Sheet" r:id="rId1"/></sheets></workbook>`,
			stray:    "custom/xl/media/stray.png",
			wantText: "Visible Sheet",
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			var buf bytes.Buffer
			zw := zip.NewWriter(&buf)
			addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
			addZip(t, zw, tc.bodyPart, tc.bodyXML)
			if strings.HasSuffix(tc.name, ".xlsx") {
				addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
				addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row><c t="inlineStr"><is><t>Visible XLSX cell</t></is></c></row></sheetData></worksheet>`)
			}
			addZipBytes(t, zw, tc.stray, testPNG())
			if err := zw.Close(); err != nil {
				t.Fatal(err)
			}
			dir := t.TempDir()
			file := filepath.Join(dir, tc.name)
			if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
				t.Fatal(err)
			}
			res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
			if err != nil {
				t.Fatal(err)
			}
			if !strings.Contains(res.Text, tc.wantText) {
				t.Fatalf("missing visible text %q in %q", tc.wantText, res.Text)
			}
			if len(res.Images) != 0 {
				t.Fatalf("stray nested media directory was extracted as content image: %#v", res.Images)
			}
			md := res.Markdown("images")
			if strings.Contains(md, "stray.png") || strings.Contains(md, "custom/") || strings.Contains(md, "![](") {
				t.Fatalf("markdown kept stray nested media image:\n%s", md)
			}
			if entries, err := os.ReadDir(filepath.Join(dir, "images")); err == nil && len(entries) != 0 {
				t.Fatalf("stray nested media image was written to image dir: %#v", entries)
			}
		})
	}
}

func TestPPTXMalformedSlideRelsKeepsImages(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:pic><p:blipFill><a:blip r:embed="rId1"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../media/visible.png">`)
	addZipBytes(t, zw, "ppt/media/visible.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "malformed-slide-rels.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	outDir := filepath.Join(dir, "images")
	res, err := Extract(file, Options{ImageDir: outDir})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected valid visible image despite malformed slide rels, got %#v", res.Images)
	}
	if _, err := os.Stat(filepath.Join(outDir, "visible.png")); err != nil {
		t.Fatalf("visible image not written: %v", err)
	}
}

func TestDOCXMixedCasePartNamesAreStructured(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "Word/Document.XML", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Mixed Case DOCX Body</w:t></w:r></w:p></w:body></w:document>`)
	addZip(t, zw, "Word/Header1.XML", `<w:hdr xmlns:w="urn:w"><w:p><w:r><w:t>Mixed Case DOCX Header</w:t></w:r></w:p></w:hdr>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Mixed Case DOCX Body", "Mixed Case DOCX Header"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing mixed-case DOCX text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Mixed Case DOCX Body", "## Headers and Footers", "Mixed Case DOCX Header"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing mixed-case DOCX content %q in:\n%s", want, md)
		}
	}
}

func TestDOCXNonCanonicalZipPartPathsAreStructured(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, `./Word\Document.XML`, `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Backslash DOCX Body</w:t></w:r></w:p><p:pic><p:nvPicPr><p:cNvPr id="1" name="Picture 1" descr="Backslash Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill></p:pic></w:body></w:document>`)
	addZip(t, zw, `./Word\Header1.XML`, `<w:hdr xmlns:w="urn:w"><w:p><w:r><w:t>Backslash DOCX Header</w:t></w:r></w:p></w:hdr>`)
	addZip(t, zw, `./Word\_rels\Document.XML.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="x" Target="Media/Image1.PNG"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/backslash-link" TargetMode="External"/></Relationships>`)
	addZip(t, zw, `./DocProps\Core.XML`, `<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:title>Backslash Metadata Title</dc:title></cp:coreProperties>`)
	addZipBytes(t, zw, `./Word\Media\Image1.PNG`, testPNG())
	addZipBytes(t, zw, `./DocProps\Thumbnail.PNG`, testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "backslash-parts.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Backslash DOCX Body", "Backslash DOCX Header"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing non-canonical DOCX text %q in %q", want, res.Text)
		}
	}
	if strings.Contains(res.Text, "Backslash Metadata Title") || strings.Contains(res.Text, "https://example.test/backslash-link") {
		t.Fatalf("default output included non-canonical metadata in %q", res.Text)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Image1.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected valid non-canonical DOCX image, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Document", "Backslash DOCX Body", "## Headers and Footers", "Backslash DOCX Header", "![Backslash Picture](images/Image1.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing non-canonical DOCX content %q in:\n%s", want, md)
		}
	}
	meta, err := Extract(file, Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Backslash Metadata Title", "https://example.test/backslash-link"} {
		if !strings.Contains(meta.Text, want) {
			t.Fatalf("metadata output missing non-canonical value %q in %q", want, meta.Text)
		}
	}
	var foundThumbnail bool
	for _, img := range meta.Images {
		if img.Name == "Thumbnail.png" && validImageData(".png", img.Data) {
			foundThumbnail = true
			break
		}
	}
	if len(meta.Images) != 2 || !foundThumbnail {
		t.Fatalf("metadata output missing valid non-canonical thumbnail, got %#v", meta.Images)
	}
}

func TestPPTXMixedCasePartNamesAreStructured(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "PPT/Slides/Slide1.XML", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Mixed Case PPTX Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "PPT/NotesSlides/NotesSlide1.XML", `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Mixed Case PPTX Notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Mixed Case PPTX Slide", "Mixed Case PPTX Notes"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing mixed-case PPTX text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Mixed Case PPTX Slide", "## Notes", "Mixed Case PPTX Notes"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing mixed-case PPTX content %q in:\n%s", want, md)
		}
	}
}

func TestPPTXNonCanonicalZipPartPathsAreStructured(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, `./PPT\Slides\Slide1.XML`, `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:sp><p:txBody><a:p><a:r><a:t>Backslash PPTX Slide</a:t></a:r></a:p></p:txBody></p:sp>
<p:pic><p:nvPicPr><p:cNvPr id="1" name="Picture 1" descr="Backslash PPTX Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Hidden Picture" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, `./PPT\NotesSlides\NotesSlide1.XML`, `<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Backslash PPTX Notes</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
	addZip(t, zw, `./PPT\Slides\_rels\Slide1.XML.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="../Media/Visible.PNG"/><Relationship Id="rIdHidden" Type="x" Target="../Media/Hidden.JPG"/></Relationships>`)
	addZipBytes(t, zw, `./PPT\Media\Visible.PNG`, testPNG())
	addZipBytes(t, zw, `./PPT\Media\Hidden.JPG`, testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "backslash-parts.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Backslash PPTX Slide", "Backslash PPTX Notes"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing non-canonical PPTX text %q in %q", want, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected one valid non-canonical PPTX visible image, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Slide 1", "Backslash PPTX Slide", "## Notes", "Backslash PPTX Notes", "![Backslash PPTX Picture](images/Visible.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing non-canonical PPTX content %q in:\n%s", want, md)
		}
	}
}

func TestXLSXMixedCasePartNamesAreStructured(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "XL/Workbook.XML", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Mixed Case Sheet" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "XL/_rels/Workbook.XML.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="Worksheets/Sheet1.XML"/></Relationships>`)
	addZip(t, zw, "XL/SharedStrings.XML", `<sst xmlns="urn:x"><si><t>Mixed Case Shared</t></si></sst>`)
	addZip(t, zw, "XL/Worksheets/Sheet1.XML", `<worksheet xmlns="urn:x"><sheetData><row><c t="s"><v>0</v></c><c t="inlineStr"><is><t>Mixed Case Inline</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Mixed Case Sheet", "Mixed Case Shared", "Mixed Case Inline"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing mixed-case XLSX text %q in %q", want, res.Text)
		}
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Mixed Case Sheet", "Mixed Case Shared", "Mixed Case Inline"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing mixed-case XLSX content %q in:\n%s", want, md)
		}
	}
}

func TestXLSXNonCanonicalZipPartPathsAreStructured(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, `./XL\Workbook.XML`, `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Backslash Sheet" r:id="rIdSheet"/></sheets></workbook>`)
	addZip(t, zw, `./XL\_rels\Workbook.XML.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSheet" Type="x" Target="Worksheets/Sheet1.XML"/></Relationships>`)
	addZip(t, zw, `./XL\SharedStrings.XML`, `<sst xmlns="urn:x"><si><t>Backslash Shared</t></si></sst>`)
	addZip(t, zw, `./XL\Worksheets\Sheet1.XML`, `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row><c t="s"><v>0</v></c><c t="inlineStr"><is><t>Backslash Inline</t></is></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>`)
	addZip(t, zw, `./XL\Worksheets\_rels\Sheet1.XML.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="x" Target="../Drawings/Drawing1.XML"/></Relationships>`)
	addZip(t, zw, `./XL\Drawings\Drawing1.XML`, `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1" descr="Backslash XLSX Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdVisible"/></xdr:blipFill></xdr:pic>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Hidden Picture" hidden="1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdHidden"/></xdr:blipFill></xdr:pic>
</xdr:wsDr>`)
	addZip(t, zw, `./XL\Drawings\_rels\Drawing1.XML.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="../Media/Visible.PNG"/><Relationship Id="rIdHidden" Type="x" Target="../Media/Hidden.JPG"/></Relationships>`)
	addZipBytes(t, zw, `./XL\Media\Visible.PNG`, testPNG())
	addZipBytes(t, zw, `./XL\Media\Hidden.JPG`, testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "backslash-parts.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Backslash Sheet", "Backslash Shared", "Backslash Inline"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing non-canonical XLSX text %q in %q", want, res.Text)
		}
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected one valid non-canonical XLSX visible image, got %#v", res.Images)
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Backslash Sheet", "Backslash Shared", "Backslash Inline", "![Backslash XLSX Picture](images/Visible.png)"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing non-canonical XLSX content %q in:\n%s", want, md)
		}
	}
}

func TestDOCXMixedCaseHiddenImagesAreFiltered(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "Word/Document.XML", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<p:pic><p:nvPicPr><p:cNvPr id="1" name="Visible Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Hidden Picture" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "Word/_rels/Document.XML.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="Media/Visible.PNG"/><Relationship Id="rIdHidden" Type="x" Target="Media/Hidden.JPG"/></Relationships>`)
	addZipBytes(t, zw, "Word/Media/Visible.PNG", testPNG())
	addZipBytes(t, zw, "Word/Media/Hidden.JPG", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case-hidden-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only mixed-case visible DOCX image, got %#v", res.Images)
	}
}

func TestDOCXURIEncodedRelationshipImageTargetsAreResolved(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible encoded image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" hidden="1" descr="Hidden encoded image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="media/Visible%20Image.PNG?download=1#section"/><Relationship Id="rIdHidden" Type="x" Target="media/Hidden%20Image.JPG#hidden"/></Relationships>`)
	addZipBytes(t, zw, "word/media/Visible Image.PNG", testPNG())
	addZipBytes(t, zw, "word/media/Hidden Image.JPG", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "encoded-target-image.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible Image.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected URI-encoded visible DOCX image target to resolve, got %#v", res.Images)
	}
	if !strings.Contains(res.Markdown("images"), "![Visible encoded image](images/Visible%20Image.png)") {
		t.Fatalf("markdown missing URI-escaped visible DOCX image reference:\n%s", res.Markdown("images"))
	}
}

func TestPPTXMixedCaseHiddenImagesAreFiltered(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "PPT/Slides/Slide1.XML", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:pic><p:nvPicPr><p:cNvPr id="1" name="Visible Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Hidden Picture" hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "PPT/Slides/_rels/Slide1.XML.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="../Media/Visible.PNG"/><Relationship Id="rIdHidden" Type="x" Target="../Media/Hidden.JPG"/></Relationships>`)
	addZipBytes(t, zw, "PPT/Media/Visible.PNG", testPNG())
	addZipBytes(t, zw, "PPT/Media/Hidden.JPG", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case-hidden-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only mixed-case visible PPTX image, got %#v", res.Images)
	}
}

func TestPPTXURIEncodedRelationshipImageTargetsAreResolved(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
<p:pic><p:nvPicPr><p:cNvPr id="1" descr="Visible encoded PPTX image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdVisible"/></p:blipFill></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="2" hidden="1" descr="Hidden encoded PPTX image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdHidden"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>`)
	addZip(t, zw, "ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="../media/Visible%20Slide%20Image.PNG#slide"/><Relationship Id="rIdHidden" Type="x" Target="../media/Hidden%20Slide%20Image.JPG?hidden=1"/></Relationships>`)
	addZipBytes(t, zw, "ppt/media/Visible Slide Image.PNG", testPNG())
	addZipBytes(t, zw, "ppt/media/Hidden Slide Image.JPG", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "encoded-target-image.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible Slide Image.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected URI-encoded visible PPTX image target to resolve, got %#v", res.Images)
	}
	if !strings.Contains(res.Markdown("images"), "![Visible encoded PPTX image](images/Visible%20Slide%20Image.png)") {
		t.Fatalf("markdown missing URI-escaped visible PPTX image reference:\n%s", res.Markdown("images"))
	}
}

func TestXLSXMixedCaseHiddenImagesAreFiltered(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "XL/Workbook.XML", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "XL/_rels/Workbook.XML.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="Worksheets/Sheet1.XML"/></Relationships>`)
	addZip(t, zw, "XL/Worksheets/Sheet1.XML", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "XL/Worksheets/_rels/Sheet1.XML.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="../Drawings/Drawing1.XML"/></Relationships>`)
	addZip(t, zw, "XL/Drawings/Drawing1.XML", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Visible Picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdVisible"/></xdr:blipFill></xdr:pic>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Hidden Picture" hidden="1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdHidden"/></xdr:blipFill></xdr:pic>
</xdr:wsDr>`)
	addZip(t, zw, "XL/Drawings/_rels/Drawing1.XML.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="../Media/Visible.PNG"/><Relationship Id="rIdHidden" Type="x" Target="../Media/Hidden.JPG"/></Relationships>`)
	addZipBytes(t, zw, "XL/Media/Visible.PNG", testPNG())
	addZipBytes(t, zw, "XL/Media/Hidden.JPG", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "mixed-case-hidden-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected only mixed-case visible XLSX image, got %#v", res.Images)
	}
}

func TestXLSXURIEncodedRelationshipImageTargetsAreResolved(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="rIdDrawing"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="x" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" descr="Visible encoded XLSX image"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdVisible"/></xdr:blipFill></xdr:pic>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" hidden="1" descr="Hidden encoded XLSX image"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdHidden"/></xdr:blipFill></xdr:pic>
</xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVisible" Type="x" Target="../media/Visible%20Sheet%20Image.PNG?sheet=1"/><Relationship Id="rIdHidden" Type="x" Target="../media/Hidden%20Sheet%20Image.JPG#hidden"/></Relationships>`)
	addZipBytes(t, zw, "xl/media/Visible Sheet Image.PNG", testPNG())
	addZipBytes(t, zw, "xl/media/Hidden Sheet Image.JPG", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "encoded-target-image.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{ImageDir: filepath.Join(dir, "images")})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "Visible Sheet Image.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected URI-encoded visible XLSX image target to resolve, got %#v", res.Images)
	}
	if !strings.Contains(res.Markdown("images"), "![Visible encoded XLSX image](images/Visible%20Sheet%20Image.png)") {
		t.Fatalf("markdown missing URI-escaped visible XLSX image reference:\n%s", res.Markdown("images"))
	}
}

func TestXLSXMalformedDrawingRelsKeepsImages(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Visible" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="rId1"/></worksheet>`)
	addZip(t, zw, "xl/worksheets/_rels/sheet1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="../drawings/drawing1.xml"/></Relationships>`)
	addZip(t, zw, "xl/drawings/drawing1.xml", `<xdr:wsDr xmlns:xdr="urn:xdr" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:wsDr>`)
	addZip(t, zw, "xl/drawings/_rels/drawing1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../media/visible.png">`)
	addZipBytes(t, zw, "xl/media/visible.png", testPNG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "malformed-drawing-rels.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 1 || res.Images[0].Name != "visible.png" || !validImageData(".png", res.Images[0].Data) {
		t.Fatalf("expected valid visible image despite malformed drawing rels, got %#v", res.Images)
	}
}

func TestDOCXMalformedDocumentRelsKeepsImages(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w" xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><p:pic><p:blipFill><a:blip r:embed="rId1"/></p:blipFill></p:pic><p:pic><p:nvPicPr><p:cNvPr hidden="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic></w:body></w:document>`)
	addZip(t, zw, "word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/visible.png"><Relationship Id="rId2" Target="media/hidden.jpg"/></Relationships>`)
	addZipBytes(t, zw, "word/media/visible.png", testPNG())
	addZipBytes(t, zw, "word/media/hidden.jpg", testJPEG())
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "malformed-document-rels.docx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Images) != 2 {
		t.Fatalf("expected images to be retained when document rels are malformed, got %#v", res.Images)
	}
	for _, img := range res.Images {
		if !validImageData(img.Ext, img.Data) {
			t.Fatalf("extracted invalid image %q", img.Name)
		}
	}
}

func TestOOXMLHiddenDrawingObjectsAreNotVisibleText(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "ppt/slides/slide1.xml", `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="1" name="Visible Shape" descr="Visible Shape Description" title="Visible Shape Title"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Visible Shape Text</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Hidden Shape" hidden="1" descr="Hidden Shape Description" title="Hidden Shape Title"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Hidden Shape Secret</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp style="display:none !important"><p:nvSpPr><p:cNvPr id="5" name="Display Hidden Shape" descr="Display Hidden Description" title="Display Hidden Title"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Display Hidden Secret</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp style="mso-hide:all"><p:nvSpPr><p:cNvPr id="6" name="MSO Hidden Shape" descr="MSO Hidden Description" title="MSO Hidden Title"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>MSO Hidden Secret</a:t></a:r></a:p></p:txBody></p:sp>
<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Hidden Group" hidden="1" descr="Hidden Group Description" title="Hidden Group Title"/></p:nvGrpSpPr><p:sp><p:nvSpPr><p:cNvPr id="4" name="Hidden Group Child" descr="Hidden Group Child Description" title="Hidden Group Child Title"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Hidden Group Child Secret</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>
</p:spTree></p:cSld></p:sld>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "hidden-shapes.pptx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible Shape Text", "Visible Shape Description", "Visible Shape Title"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("missing visible shape text %q in %q", want, res.Text)
		}
	}
	for _, hidden := range []string{"Hidden Shape Secret", "Hidden Shape Description", "Hidden Shape Title", "Display Hidden Secret", "Display Hidden Description", "Display Hidden Title", "MSO Hidden Secret", "MSO Hidden Description", "MSO Hidden Title", "Hidden Group Description", "Hidden Group Title", "Hidden Group Child Secret", "Hidden Group Child Description", "Hidden Group Child Title"} {
		if strings.Contains(res.Text, hidden) {
			t.Fatalf("kept hidden drawing object text %q in %q", hidden, res.Text)
		}
	}
}

func TestLegacyXLSDropsEmbeddedPDFText(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "61300.xls"), Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, noise := range []string{"2 0 obj", "/Filter /FlateDecode", "endstream", "endobj", "/Type /Page", "L/f`", "%RMR", "P|Y."} {
		if strings.Contains(res.Text, noise) {
			t.Fatalf("kept embedded PDF text %q in %.400q", noise, res.Text)
		}
	}
	for _, noise := range []string{"Microsoft Excel", "Root Entry", "Calibri1", "\u844a\u887b\u81b0", "\u0200\u0400\u0107Calibri"} {
		if strings.Contains(res.Text, noise) {
			t.Fatalf("kept workbook control text %q in %.400q", noise, res.Text)
		}
	}
}

func TestPropertySetLPSTRDecodesWindows1252(t *testing.T) {
	raw := append([]byte("Int\xe9grer une vid\xe9o"), 0)
	data := make([]byte, 4+len(raw))
	binary.LittleEndian.PutUint32(data, uint32(len(raw)))
	copy(data[4:], raw)
	text, ok := propertySetLPSTR(data, 1252)
	if !ok || text != "Int\u00e9grer une vid\u00e9o" {
		t.Fatalf("got %q, ok=%v", text, ok)
	}
}

func TestPropertySetLPSTRDecodesShiftJIS(t *testing.T) {
	raw := append([]byte{0x83, 0x58, 0x83, 0x89, 0x83, 0x43, 0x83, 0x68, 0x20, 0x31}, 0)
	data := make([]byte, 4+len(raw))
	binary.LittleEndian.PutUint32(data, uint32(len(raw)))
	copy(data[4:], raw)
	text, ok := propertySetLPSTR(data, 932)
	if !ok || text != "スライド 1" {
		t.Fatalf("got %q, ok=%v", text, ok)
	}
}

func TestPropertySetLPSTRDecodesEastAsianCodePages(t *testing.T) {
	cases := []struct {
		name     string
		codePage uint16
		raw      []byte
		want     string
	}{
		{name: "gbk", codePage: 936, raw: []byte{0xd6, 0xd0, 0xce, 0xc4}, want: "\u4e2d\u6587"},
		{name: "big5", codePage: 950, raw: []byte{0xa4, 0xa4, 0xa4, 0xe5}, want: "\u4e2d\u6587"},
		{name: "korean", codePage: 949, raw: []byte{0xc7, 0xd1, 0xb1, 0xdb}, want: "\ud55c\uae00"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			raw := append(append([]byte(nil), tc.raw...), 0)
			data := make([]byte, 4+len(raw))
			binary.LittleEndian.PutUint32(data, uint32(len(raw)))
			copy(data[4:], raw)
			text, ok := propertySetLPSTR(data, tc.codePage)
			if !ok || text != tc.want {
				t.Fatalf("got %q, ok=%v", text, ok)
			}
		})
	}
}

func TestLegacyPPTShiftJISPropertiesDoNotEmitMojibake(t *testing.T) {
	res, err := Extract(filepath.Join("testdata", "samples", "54880_chinese.ppt"), Options{IncludeMetadata: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "スライド 1") || !strings.Contains(res.Text, "画面に合わせる") {
		t.Fatalf("missing decoded Shift-JIS property text in %.400q", res.Text)
	}
	for _, bad := range []string{"ƒXƒ‰ƒCƒh", "‰æ–Ê‚É"} {
		if strings.Contains(res.Text, bad) {
			t.Fatalf("kept Shift-JIS mojibake %q in %.400q", bad, res.Text)
		}
	}
}

func testPNG() []byte {
	b, err := base64.StdEncoding.DecodeString("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC")
	if err != nil {
		panic(err)
	}
	return b
}

func utf16LEBOMBytes(s string) []byte {
	out := []byte{0xff, 0xfe}
	return append(out, utf16LEBytes(s)...)
}

func utf16BEBOMBytes(s string) []byte {
	out := []byte{0xfe, 0xff}
	return append(out, utf16BEBytes(s)...)
}

func utf16BEBytes(s string) []byte {
	u := utf16.Encode([]rune(s))
	out := make([]byte, 0, len(u)*2)
	for _, v := range u {
		out = binary.BigEndian.AppendUint16(out, v)
	}
	return out
}

func testPNGWithPrivateChunk(payload []byte) []byte {
	png := testPNG()
	if len(png) < 20 {
		return png
	}
	iend := len(png) - 12
	chunkType := []byte("vpAg")
	chunk := make([]byte, 8+len(payload)+4)
	binary.BigEndian.PutUint32(chunk[0:], uint32(len(payload)))
	copy(chunk[4:], chunkType)
	copy(chunk[8:], payload)
	crcPayload := append(append([]byte(nil), chunkType...), payload...)
	binary.BigEndian.PutUint32(chunk[8+len(payload):], crc32.ChecksumIEEE(crcPayload))
	out := make([]byte, 0, len(png)+len(chunk))
	out = append(out, png[:iend]...)
	out = append(out, chunk...)
	out = append(out, png[iend:]...)
	return out
}

func testDIB() []byte {
	dib := make([]byte, 40+4)
	binary.LittleEndian.PutUint32(dib[0:], 40)
	binary.LittleEndian.PutUint32(dib[4:], 1)
	binary.LittleEndian.PutUint32(dib[8:], 1)
	binary.LittleEndian.PutUint16(dib[12:], 1)
	binary.LittleEndian.PutUint16(dib[14:], 24)
	binary.LittleEndian.PutUint32(dib[20:], 4)
	dib[40] = 0xff
	return dib
}

func testDIBWithPayload(payload []byte) []byte {
	width := 8
	height := 1
	stride := ((width*24 + 31) / 32) * 4
	dib := make([]byte, 40+stride)
	binary.LittleEndian.PutUint32(dib[0:], 40)
	binary.LittleEndian.PutUint32(dib[4:], uint32(width))
	binary.LittleEndian.PutUint32(dib[8:], uint32(height))
	binary.LittleEndian.PutUint16(dib[12:], 1)
	binary.LittleEndian.PutUint16(dib[14:], 24)
	binary.LittleEndian.PutUint32(dib[20:], uint32(stride))
	copy(dib[40:], payload)
	return dib
}

func testBitfieldsDIB(width, height int, bitCount uint16, compression uint32, masks []uint32) []byte {
	maskBytes := len(masks) * 4
	stride := ((width*int(bitCount) + 31) / 32) * 4
	dib := make([]byte, 40+maskBytes+stride*height)
	binary.LittleEndian.PutUint32(dib[0:], 40)
	binary.LittleEndian.PutUint32(dib[4:], uint32(width))
	binary.LittleEndian.PutUint32(dib[8:], uint32(height))
	binary.LittleEndian.PutUint16(dib[12:], 1)
	binary.LittleEndian.PutUint16(dib[14:], bitCount)
	binary.LittleEndian.PutUint32(dib[16:], compression)
	binary.LittleEndian.PutUint32(dib[20:], uint32(stride*height))
	for i, mask := range masks {
		binary.LittleEndian.PutUint32(dib[40+i*4:], mask)
	}
	for i := 40 + maskBytes; i < len(dib); i++ {
		dib[i] = 0xff
	}
	return dib
}

func testCoreDIB(width, height int) []byte {
	stride := ((width*24 + 31) / 32) * 4
	dib := make([]byte, 12+stride*height)
	binary.LittleEndian.PutUint32(dib[0:], 12)
	binary.LittleEndian.PutUint16(dib[4:], uint16(width))
	binary.LittleEndian.PutUint16(dib[6:], uint16(height))
	binary.LittleEndian.PutUint16(dib[8:], 1)
	binary.LittleEndian.PutUint16(dib[10:], 24)
	dib[12] = 0xff
	return dib
}

func testJPEG() []byte {
	img := image.NewRGBA(image.Rect(0, 0, 1, 1))
	img.Set(0, 0, color.RGBA{R: 0xff, A: 0xff})
	var buf bytes.Buffer
	if err := jpeg.Encode(&buf, img, &jpeg.Options{Quality: 90}); err != nil {
		panic(err)
	}
	return buf.Bytes()
}

func testGIF() []byte {
	img := image.NewPaletted(image.Rect(0, 0, 1, 1), []color.Color{
		color.Black,
		color.White,
	})
	img.SetColorIndex(0, 0, 1)
	var buf bytes.Buffer
	if err := gif.Encode(&buf, img, nil); err != nil {
		panic(err)
	}
	return buf.Bytes()
}

func testSVG() []byte {
	return []byte(`<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1" fill="red"/></svg>`)
}

func testEPS() []byte {
	return []byte("%!PS-Adobe-3.0 EPSF-3.0\n%%BoundingBox: 0 0 10 10\n/newpath { } def\n0 0 moveto\n10 10 lineto\nstroke\nshowpage\n%%EOF\n")
}

func testEPSWithBody(body string) []byte {
	return []byte("%!PS-Adobe-3.0 EPSF-3.0\n%%BoundingBox: 0 0 10 10\n" + body + "showpage\n%%EOF\n")
}

func testTIFF() []byte {
	return testTIFFWithStripTags(true, true, 1, 1)
}

func testTIFFWithEditedEntry(tag uint16, edit func([]byte)) []byte {
	b := testTIFF()
	offset := int(binary.LittleEndian.Uint32(b[4:]))
	entries := int(binary.LittleEndian.Uint16(b[offset:]))
	for i := 0; i < entries; i++ {
		pos := offset + 2 + i*12
		if binary.LittleEndian.Uint16(b[pos:]) == tag {
			edit(b[pos : pos+12])
			return b
		}
	}
	panic("missing TIFF test tag")
}

func testTIFFWithSelfReferentialNextIFD() []byte {
	b := testTIFF()
	offset := int(binary.LittleEndian.Uint32(b[4:]))
	entries := int(binary.LittleEndian.Uint16(b[offset:]))
	binary.LittleEndian.PutUint32(b[offset+2+entries*12:], uint32(offset))
	return b
}

func testTIFFWithTileLayout(tileOffset, tileByteCount uint32) []byte {
	b := testTIFF()
	testTIFFWithEditedEntryInPlace(b, 273, func(entry []byte) {
		binary.LittleEndian.PutUint16(entry[0:], 324)
		binary.LittleEndian.PutUint32(entry[8:], tileOffset)
	})
	testTIFFWithEditedEntryInPlace(b, 279, func(entry []byte) {
		binary.LittleEndian.PutUint16(entry[0:], 325)
		binary.LittleEndian.PutUint32(entry[8:], tileByteCount)
	})
	return b
}

func testTIFFWithEditedEntryInPlace(b []byte, tag uint16, edit func([]byte)) {
	offset := int(binary.LittleEndian.Uint32(b[4:]))
	entries := int(binary.LittleEndian.Uint16(b[offset:]))
	for i := 0; i < entries; i++ {
		pos := offset + 2 + i*12
		if binary.LittleEndian.Uint16(b[pos:]) == tag {
			edit(b[pos : pos+12])
			return
		}
	}
	panic("missing TIFF test tag")
}

func testBigTIFF() []byte {
	const entries = 9
	ifdOffset := 16
	dataOffset := ifdOffset + 8 + entries*20 + 8
	b := make([]byte, dataOffset+1)
	copy(b[0:], []byte("II"))
	binary.LittleEndian.PutUint16(b[2:], 43)
	binary.LittleEndian.PutUint16(b[4:], 8)
	binary.LittleEndian.PutUint64(b[8:], uint64(ifdOffset))
	binary.LittleEndian.PutUint64(b[ifdOffset:], entries)
	putBigTIFFEntry := func(index int, tag, fieldType uint16, count, value uint64) {
		pos := ifdOffset + 8 + index*20
		binary.LittleEndian.PutUint16(b[pos:], tag)
		binary.LittleEndian.PutUint16(b[pos+2:], fieldType)
		binary.LittleEndian.PutUint64(b[pos+4:], count)
		binary.LittleEndian.PutUint64(b[pos+12:], value)
	}
	putBigTIFFEntry(0, 256, 16, 1, 1)
	putBigTIFFEntry(1, 257, 16, 1, 1)
	putBigTIFFEntry(2, 258, 3, 1, 8)
	putBigTIFFEntry(3, 259, 3, 1, 1)
	putBigTIFFEntry(4, 262, 3, 1, 1)
	putBigTIFFEntry(5, 273, 16, 1, uint64(dataOffset))
	putBigTIFFEntry(6, 277, 3, 1, 1)
	putBigTIFFEntry(7, 278, 4, 1, 1)
	putBigTIFFEntry(8, 279, 16, 1, 1)
	b[dataOffset] = 0x7f
	return b
}

func testBigTIFFWithEditedEntry(tag uint16, edit func([]byte)) []byte {
	b := testBigTIFF()
	offset := int(binary.LittleEndian.Uint64(b[8:]))
	entries := int(binary.LittleEndian.Uint64(b[offset:]))
	for i := 0; i < entries; i++ {
		pos := offset + 8 + i*20
		if binary.LittleEndian.Uint16(b[pos:]) == tag {
			edit(b[pos : pos+20])
			return b
		}
	}
	panic("missing BigTIFF test tag")
}

func testTIFFWithJPEGInterchangeTags(includeOffset, includeLength bool) []byte {
	const entries = 9
	ifdOffset := 8
	dataOffset := ifdOffset + 2 + entries*12 + 4
	b := make([]byte, dataOffset+2)
	copy(b[0:], []byte("II"))
	binary.LittleEndian.PutUint16(b[2:], 42)
	binary.LittleEndian.PutUint32(b[4:], uint32(ifdOffset))
	binary.LittleEndian.PutUint16(b[ifdOffset:], entries)
	putTIFFEntry := func(index int, tag, fieldType uint16, count, value uint32) {
		pos := ifdOffset + 2 + index*12
		binary.LittleEndian.PutUint16(b[pos:], tag)
		binary.LittleEndian.PutUint16(b[pos+2:], fieldType)
		binary.LittleEndian.PutUint32(b[pos+4:], count)
		binary.LittleEndian.PutUint32(b[pos+8:], value)
	}
	putTIFFEntry(0, 256, 4, 1, 1)
	putTIFFEntry(1, 257, 4, 1, 1)
	putTIFFEntry(2, 258, 3, 1, 8)
	putTIFFEntry(3, 259, 3, 1, 7)
	putTIFFEntry(4, 262, 3, 1, 6)
	if includeOffset {
		putTIFFEntry(5, 513, 4, 1, uint32(dataOffset))
	} else {
		putTIFFEntry(5, 305, 2, 1, 0)
	}
	if includeLength {
		putTIFFEntry(6, 514, 4, 1, 2)
	} else {
		putTIFFEntry(6, 305, 2, 1, 0)
	}
	putTIFFEntry(7, 277, 3, 1, 1)
	putTIFFEntry(8, 278, 4, 1, 1)
	copy(b[dataOffset:], []byte{0xff, 0xd8})
	return b
}

func testTIFFWithStripTags(includeOffsets, includeByteCounts bool, offsetCount, byteCountCount uint32) []byte {
	const entries = 9
	ifdOffset := 8
	dataOffset := ifdOffset + 2 + entries*12 + 4
	extraData := 1
	if offsetCount > 1 || byteCountCount > 1 {
		extraData = 8
	}
	b := make([]byte, dataOffset+extraData)
	copy(b[0:], []byte("II"))
	binary.LittleEndian.PutUint16(b[2:], 42)
	binary.LittleEndian.PutUint32(b[4:], uint32(ifdOffset))
	binary.LittleEndian.PutUint16(b[ifdOffset:], entries)
	putTIFFEntry := func(index int, tag, fieldType uint16, count, value uint32) {
		pos := ifdOffset + 2 + index*12
		binary.LittleEndian.PutUint16(b[pos:], tag)
		binary.LittleEndian.PutUint16(b[pos+2:], fieldType)
		binary.LittleEndian.PutUint32(b[pos+4:], count)
		binary.LittleEndian.PutUint32(b[pos+8:], value)
	}
	putTIFFEntry(0, 256, 4, 1, 1)
	putTIFFEntry(1, 257, 4, 1, 1)
	putTIFFEntry(2, 258, 3, 1, 8)
	putTIFFEntry(3, 259, 3, 1, 1)
	putTIFFEntry(4, 262, 3, 1, 1)
	if includeOffsets {
		if offsetCount > 1 {
			putTIFFEntry(5, 273, 4, offsetCount, uint32(dataOffset))
			for i := uint32(0); i < offsetCount && int(i)*4+4 <= extraData; i++ {
				binary.LittleEndian.PutUint32(b[dataOffset+int(i)*4:], uint32(dataOffset))
			}
		} else {
			putTIFFEntry(5, 273, 4, offsetCount, uint32(dataOffset))
		}
	} else {
		putTIFFEntry(5, 270, 2, 1, 0)
	}
	putTIFFEntry(6, 277, 3, 1, 1)
	putTIFFEntry(7, 278, 4, 1, 1)
	if includeByteCounts {
		if byteCountCount > 1 {
			putTIFFEntry(8, 279, 4, byteCountCount, uint32(dataOffset))
			for i := uint32(0); i < byteCountCount && int(i)*4+4 <= extraData; i++ {
				binary.LittleEndian.PutUint32(b[dataOffset+int(i)*4:], 1)
			}
		} else {
			putTIFFEntry(8, 279, 4, byteCountCount, 1)
		}
	} else {
		putTIFFEntry(8, 305, 2, 1, 0)
	}
	return b
}

func testJPEGXR() []byte {
	const entries = 5
	ifdOffset := 8
	dataOffset := ifdOffset + 2 + entries*12 + 4
	b := make([]byte, dataOffset+4)
	copy(b[0:], []byte("II"))
	binary.LittleEndian.PutUint16(b[2:], 0x01bc)
	binary.LittleEndian.PutUint32(b[4:], uint32(ifdOffset))
	binary.LittleEndian.PutUint16(b[ifdOffset:], entries)
	putEntry := func(index int, tag, fieldType uint16, count, value uint32) {
		pos := ifdOffset + 2 + index*12
		binary.LittleEndian.PutUint16(b[pos:], tag)
		binary.LittleEndian.PutUint16(b[pos+2:], fieldType)
		binary.LittleEndian.PutUint32(b[pos+4:], count)
		binary.LittleEndian.PutUint32(b[pos+8:], value)
	}
	putEntry(0, 256, 4, 1, 1)
	putEntry(1, 257, 4, 1, 1)
	putEntry(2, 273, 4, 1, uint32(dataOffset))
	putEntry(3, 277, 3, 1, 3)
	putEntry(4, 279, 4, 1, 4)
	copy(b[dataOffset:], []byte{0x01, 0x02, 0x03, 0x04})
	return b
}

func testJPEGXRWithEditedEntry(tag uint16, edit func([]byte)) []byte {
	b := testJPEGXR()
	offset := int(binary.LittleEndian.Uint32(b[4:]))
	entries := int(binary.LittleEndian.Uint16(b[offset:]))
	for i := 0; i < entries; i++ {
		pos := offset + 2 + i*12
		if binary.LittleEndian.Uint16(b[pos:]) == tag {
			edit(b[pos : pos+12])
			return b
		}
	}
	panic("missing JPEG XR test tag")
}

func testJPEGXRWithSelfReferentialNextIFD() []byte {
	b := testJPEGXR()
	offset := int(binary.LittleEndian.Uint32(b[4:]))
	entries := int(binary.LittleEndian.Uint16(b[offset:]))
	binary.LittleEndian.PutUint32(b[offset+2+entries*12:], uint32(offset))
	return b
}

func testWebP() []byte {
	chunks := append(testWebPVP8XChunk(), testWebPChunk("VP8L", []byte{0x2f, 0, 0, 0, 0})...)
	chunks = append(chunks, testWebPChunk("EXIF", []byte("test"))...)
	return testWebPWithChunks(chunks)
}

func testWebPMetadataOnly() []byte {
	chunks := append(testWebPVP8XChunk(), testWebPChunk("EXIF", []byte("test"))...)
	return testWebPWithChunks(chunks)
}

func testAnimatedWebP() []byte {
	return testWebPWithChunks(testWebPVP8XChunkWithFlags(0x02), testWebPChunk("ANMF", testWebPANMFPayload()))
}

func testWebPANMFPayload() []byte {
	frame := make([]byte, 16)
	return append(frame, testWebPChunk("VP8L", []byte{0x2f, 0, 0, 0, 0})...)
}

func testWebPVP8XChunk() []byte {
	return testWebPVP8XChunkWithFlags(0)
}

func testWebPVP8XChunkWithFlags(flags byte) []byte {
	return testWebPChunk("VP8X", []byte{flags, 0, 0, 0, 0, 0, 0, 0, 0, 0})
}

func testWebPWithChunks(chunks ...[]byte) []byte {
	joined := bytes.Join(chunks, nil)
	b := make([]byte, 12+len(joined))
	copy(b[0:], []byte("RIFF"))
	binary.LittleEndian.PutUint32(b[4:], uint32(len(b)-8))
	copy(b[8:], []byte("WEBP"))
	copy(b[12:], joined)
	return b
}

func testWebPChunk(chunkType string, payload []byte) []byte {
	chunk := make([]byte, 8+len(payload))
	copy(chunk[0:], []byte(chunkType))
	binary.LittleEndian.PutUint32(chunk[4:], uint32(len(payload)))
	copy(chunk[8:], payload)
	if len(payload)%2 == 1 {
		chunk = append(chunk, 0)
	}
	return chunk
}

func testInvalidWebP(chunkType string, payload []byte) []byte {
	b := make([]byte, 12+8+len(payload))
	copy(b[0:], []byte("RIFF"))
	binary.LittleEndian.PutUint32(b[4:], uint32(len(b)-8))
	copy(b[8:], []byte("WEBP"))
	copy(b[12:], []byte(chunkType))
	binary.LittleEndian.PutUint32(b[16:], uint32(len(payload)))
	copy(b[20:], payload)
	return b
}

func testICO() []byte {
	return testIconDirectory(1)
}

func testCUR() []byte {
	return testIconDirectory(2)
}

func testIconDirectory(kind uint16) []byte {
	return testIconDirectoryWithPayload(kind, 1, 1, testPNG())
}

func testIconDirectoryWithPayload(kind uint16, width, height int, payload []byte) []byte {
	headerSize := 6 + 16
	b := make([]byte, headerSize+len(payload))
	binary.LittleEndian.PutUint16(b[2:], kind)
	binary.LittleEndian.PutUint16(b[4:], 1)
	b[6] = byte(width)
	b[7] = byte(height)
	b[8] = 0
	b[9] = 0
	binary.LittleEndian.PutUint16(b[10:], 1)
	binary.LittleEndian.PutUint16(b[12:], 32)
	binary.LittleEndian.PutUint32(b[14:], uint32(len(payload)))
	binary.LittleEndian.PutUint32(b[18:], uint32(headerSize))
	copy(b[headerSize:], payload)
	return b
}

func testPCX() []byte {
	b := make([]byte, 128)
	b[0] = 0x0a
	b[1] = 5
	b[2] = 1
	b[3] = 8
	binary.LittleEndian.PutUint16(b[8:], 1)
	binary.LittleEndian.PutUint16(b[10:], 1)
	binary.LittleEndian.PutUint16(b[12:], 72)
	binary.LittleEndian.PutUint16(b[14:], 72)
	b[65] = 1
	binary.LittleEndian.PutUint16(b[66:], 2)
	binary.LittleEndian.PutUint16(b[68:], 1)
	b = append(b, 1, 2, 3, 4)
	palette := make([]byte, 769)
	palette[0] = 0x0c
	palette[1] = 0xff
	palette[4] = 0xff
	return append(b, palette...)
}

func testTGA() []byte {
	b := make([]byte, 18+3)
	b[2] = 2
	binary.LittleEndian.PutUint16(b[12:], 1)
	binary.LittleEndian.PutUint16(b[14:], 1)
	b[16] = 24
	b[18], b[19], b[20] = 0x20, 0x40, 0xff
	return b
}

func testTGARLE() []byte {
	b := make([]byte, 18)
	b[2] = 10
	binary.LittleEndian.PutUint16(b[12:], 2)
	binary.LittleEndian.PutUint16(b[14:], 1)
	b[16] = 24
	return append(b, 0x81, 0x20, 0x40, 0xff)
}

func testPICT(wrapped bool) []byte {
	raw := make([]byte, 40)
	binary.BigEndian.PutUint16(raw[6:], 2)
	binary.BigEndian.PutUint16(raw[8:], 2)
	copy(raw[10:], []byte{0x00, 0x11, 0x02, 0xff, 0x0c, 0x00, 0xff, 0xfe})
	binary.BigEndian.PutUint32(raw[18:], 0x00480000)
	binary.BigEndian.PutUint32(raw[22:], 0x00480000)
	binary.BigEndian.PutUint16(raw[30:], 2)
	binary.BigEndian.PutUint16(raw[32:], 2)
	raw[38], raw[39] = 0x00, 0xff
	if !wrapped {
		return raw
	}
	return append(make([]byte, 512), raw...)
}

func testISOBMFF(brand string) []byte {
	ftyp := make([]byte, 24)
	binary.BigEndian.PutUint32(ftyp[0:], uint32(len(ftyp)))
	copy(ftyp[4:], []byte("ftyp"))
	copy(ftyp[8:], []byte(brand))
	binary.BigEndian.PutUint32(ftyp[12:], 0)
	copy(ftyp[16:], []byte(brand))
	compatible := brand
	switch brand {
	case "avif", "avis", "heic", "heix", "hevc", "hevx", "mif1", "msf1":
		compatible = "mif1"
	}
	copy(ftyp[20:], []byte(compatible))

	meta := make([]byte, 20)
	binary.BigEndian.PutUint32(meta[0:], uint32(len(meta)))
	copy(meta[4:], []byte("meta"))
	binary.BigEndian.PutUint32(meta[12:], 8)
	copy(meta[16:], []byte("pitm"))

	mdat := make([]byte, 16)
	binary.BigEndian.PutUint32(mdat[0:], uint32(len(mdat)))
	copy(mdat[4:], []byte("mdat"))
	copy(mdat[8:], []byte("payload!"))

	return append(append(ftyp, meta...), mdat...)
}

func testISOBMFFWithEmptyMeta(brand string) []byte {
	withMeta := testISOBMFF(brand)
	if len(withMeta) < 44 {
		return withMeta
	}
	ftyp := append([]byte(nil), withMeta[:24]...)
	meta := make([]byte, 12)
	binary.BigEndian.PutUint32(meta[0:], uint32(len(meta)))
	copy(meta[4:], []byte("meta"))
	mdat := append([]byte(nil), withMeta[44:]...)
	return append(append(ftyp, meta...), mdat...)
}

func testISOBMFFWithoutMeta(brand string) []byte {
	withMeta := testISOBMFF(brand)
	if len(withMeta) < 44 {
		return withMeta
	}
	ftyp := append([]byte(nil), withMeta[:24]...)
	mdat := append([]byte(nil), withMeta[44:]...)
	return append(ftyp, mdat...)
}

func testJP2(brand string) []byte {
	signature := []byte{0, 0, 0, 12, 'j', 'P', ' ', ' ', 0x0d, 0x0a, 0x87, 0x0a}
	ftyp := makeJP2Box("ftyp", append(append([]byte(brand), 0, 0, 0, 0), []byte(brand)...))

	jp2h := testJP2HeaderBox(1, 1, 1)
	jp2c := makeJP2Box("jp2c", testJ2K())

	return append(append(append(signature, ftyp...), jp2h...), jp2c...)
}

func testJP2WithBoxes(brand string, boxes ...[]byte) []byte {
	signature := []byte{0, 0, 0, 12, 'j', 'P', ' ', ' ', 0x0d, 0x0a, 0x87, 0x0a}
	ftyp := makeJP2Box("ftyp", append(append([]byte(brand), 0, 0, 0, 0), []byte(brand)...))
	out := append(append([]byte(nil), signature...), ftyp...)
	for _, box := range boxes {
		out = append(out, box...)
	}
	return out
}

func testJP2HeaderBox(width, height uint32, components uint16) []byte {
	ihdr := make([]byte, 14)
	binary.BigEndian.PutUint32(ihdr[0:], height)
	binary.BigEndian.PutUint32(ihdr[4:], width)
	binary.BigEndian.PutUint16(ihdr[8:], components)
	ihdr[10] = 7
	ihdr[11] = 7
	ihdr[12] = 0
	ihdr[13] = 0
	return makeJP2Box("jp2h", makeJP2Box("ihdr", ihdr))
}

func makeJP2Box(boxType string, payload []byte) []byte {
	box := make([]byte, 8+len(payload))
	binary.BigEndian.PutUint32(box[0:], uint32(len(box)))
	copy(box[4:], []byte(boxType))
	copy(box[8:], payload)
	return box
}

func testJ2K() []byte {
	const segLen = 41
	b := make([]byte, 4+segLen)
	b[0], b[1], b[2], b[3] = 0xff, 0x4f, 0xff, 0x51
	binary.BigEndian.PutUint16(b[4:], segLen)
	binary.BigEndian.PutUint32(b[8:], 1)
	binary.BigEndian.PutUint32(b[12:], 1)
	binary.BigEndian.PutUint32(b[24:], 1)
	binary.BigEndian.PutUint32(b[28:], 1)
	binary.BigEndian.PutUint16(b[40:], 1)
	b[42], b[43], b[44] = 7, 1, 1
	return append(b, 0xff, 0xd9)
}

func testEMF() []byte {
	b := make([]byte, 108)
	binary.LittleEndian.PutUint32(b[0:], 1)
	binary.LittleEndian.PutUint32(b[4:], uint32(len(b)))
	copy(b[40:], []byte{' ', 'E', 'M', 'F'})
	eof := make([]byte, 20)
	binary.LittleEndian.PutUint32(eof[0:], 14)
	binary.LittleEndian.PutUint32(eof[4:], uint32(len(eof)))
	b = append(b, eof...)
	binary.LittleEndian.PutUint32(b[48:], uint32(len(b)))
	binary.LittleEndian.PutUint32(b[52:], 2)
	return b
}

func testEMFWithPayload(payload []byte) []byte {
	return testEMFWithRecordPayload(payload)
}

func testEMFWithRecordPayload(payload []byte) []byte {
	b := testEMF()
	eof := append([]byte(nil), b[len(b)-20:]...)
	b = b[:len(b)-20]
	recordPayload := append([]byte(nil), payload...)
	for (8+len(recordPayload))%4 != 0 {
		recordPayload = append(recordPayload, 0)
	}
	record := make([]byte, 8+len(recordPayload))
	binary.LittleEndian.PutUint32(record[0:], 2)
	binary.LittleEndian.PutUint32(record[4:], uint32(len(record)))
	copy(record[8:], recordPayload)
	b = append(b, record...)
	b = append(b, eof...)
	binary.LittleEndian.PutUint32(b[48:], uint32(len(b)))
	binary.LittleEndian.PutUint32(b[52:], 3)
	return b
}

func testPlaceableWMF() []byte {
	standard := make([]byte, 24)
	binary.LittleEndian.PutUint16(standard[0:], 1)
	binary.LittleEndian.PutUint16(standard[2:], 9)
	binary.LittleEndian.PutUint16(standard[4:], 0x0300)
	binary.LittleEndian.PutUint32(standard[6:], 12)
	binary.LittleEndian.PutUint32(standard[12:], 3)
	binary.LittleEndian.PutUint32(standard[18:], 3)

	placeable := make([]byte, 22)
	binary.LittleEndian.PutUint32(placeable[0:], 0x9ac6cdd7)
	binary.LittleEndian.PutUint16(placeable[10:], 1)
	binary.LittleEndian.PutUint16(placeable[12:], 1)
	binary.LittleEndian.PutUint16(placeable[14:], 1440)
	var checksum uint16
	for i := 0; i < 20; i += 2 {
		checksum ^= binary.LittleEndian.Uint16(placeable[i:])
	}
	binary.LittleEndian.PutUint16(placeable[20:], checksum)
	return append(placeable, standard...)
}

type wmfTestRecord struct {
	words    uint32
	function uint16
	payload  []byte
}

func testPlaceableWMFWithRecords(records []wmfTestRecord, maxRecordWords uint32) []byte {
	standard := make([]byte, 18)
	binary.LittleEndian.PutUint16(standard[0:], 1)
	binary.LittleEndian.PutUint16(standard[2:], 9)
	binary.LittleEndian.PutUint16(standard[4:], 0x0300)
	binary.LittleEndian.PutUint32(standard[12:], maxRecordWords)
	for _, record := range records {
		rec := make([]byte, 6+len(record.payload))
		binary.LittleEndian.PutUint32(rec[0:], record.words)
		binary.LittleEndian.PutUint16(rec[4:], record.function)
		copy(rec[6:], record.payload)
		standard = append(standard, rec...)
	}
	eof := make([]byte, 6)
	binary.LittleEndian.PutUint32(eof[0:], 3)
	standard = append(standard, eof...)
	binary.LittleEndian.PutUint32(standard[6:], uint32(len(standard)/2))

	placeable := make([]byte, 22)
	binary.LittleEndian.PutUint32(placeable[0:], 0x9ac6cdd7)
	binary.LittleEndian.PutUint16(placeable[10:], 1)
	binary.LittleEndian.PutUint16(placeable[12:], 1)
	binary.LittleEndian.PutUint16(placeable[14:], 1440)
	var checksum uint16
	for i := 0; i < 20; i += 2 {
		checksum ^= binary.LittleEndian.Uint16(placeable[i:])
	}
	binary.LittleEndian.PutUint16(placeable[20:], checksum)
	return append(placeable, standard...)
}

func testWMFWithPayload(payload []byte) []byte {
	recordPayload := append([]byte(nil), payload...)
	for len(recordPayload)%2 != 0 {
		recordPayload = append(recordPayload, 0)
	}
	recordWords := uint32(3 + len(recordPayload)/2)
	return testPlaceableWMFWithRecords([]wmfTestRecord{{
		words:    recordWords,
		function: 0x0201,
		payload:  recordPayload,
	}}, recordWords)
}

func gzipBytes(t *testing.T, b []byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := gzip.NewWriter(&buf)
	if _, err := zw.Write(b); err != nil {
		t.Fatal(err)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestPropertySetTextExtractsStrings(t *testing.T) {
	data := testPropertySet(map[uint32]string{
		2:  "Generated Legacy Property Title",
		3:  "file://server/share/hidden.doc",
		4:  "(rId7)",
		5:  "word/media/image1.png",
		15: "Generated Legacy Property Company",
	})
	text := strings.Join(propertySetText(data), "\n")
	if !strings.Contains(text, "Generated Legacy Property Title") {
		t.Fatalf("missing title property in %q", text)
	}
	if !strings.Contains(text, "Generated Legacy Property Company") {
		t.Fatalf("missing company property in %q", text)
	}
	for _, hidden := range []string{"file://server/share/hidden.doc", "rId7", "word/media/image1.png"} {
		if strings.Contains(text, hidden) {
			t.Fatalf("kept internal legacy property value %q in %q", hidden, text)
		}
	}
}

func testPropertySet(values map[uint32]string) []byte {
	const sectionOffset = 48
	count := 1 + len(values)
	section := make([]byte, 8+count*8)
	binary.LittleEndian.PutUint32(section[4:], uint32(count))
	valueOffset := len(section)
	putProp := func(index int, id uint32, value []byte) {
		entry := section[8+index*8:]
		binary.LittleEndian.PutUint32(entry, id)
		binary.LittleEndian.PutUint32(entry[4:], uint32(valueOffset))
		section = append(section, value...)
		for len(section)%4 != 0 {
			section = append(section, 0)
		}
		valueOffset = len(section)
	}
	cp := make([]byte, 8)
	binary.LittleEndian.PutUint16(cp, 0x0002)
	binary.LittleEndian.PutUint16(cp[4:], 1200)
	putProp(0, 1, cp)
	i := 1
	for id, value := range values {
		putProp(i, id, testLPWSTR(value))
		i++
	}
	binary.LittleEndian.PutUint32(section, uint32(len(section)))

	out := make([]byte, sectionOffset)
	binary.LittleEndian.PutUint16(out, 0xfffe)
	binary.LittleEndian.PutUint32(out[24:], 1)
	binary.LittleEndian.PutUint32(out[44:], sectionOffset)
	out = append(out, section...)
	return out
}

func testLPWSTR(s string) []byte {
	u := utf16.Encode(append([]rune(s), 0))
	out := make([]byte, 8+len(u)*2)
	binary.LittleEndian.PutUint16(out, 0x001f)
	binary.LittleEndian.PutUint32(out[4:], uint32(len(u)))
	for i, v := range u {
		binary.LittleEndian.PutUint16(out[8+i*2:], v)
	}
	return out
}

func utf16LEBytes(s string) []byte {
	u := utf16.Encode([]rune(s))
	out := make([]byte, 0, len(u)*2)
	for _, v := range u {
		out = append(out, byte(v), byte(v>>8))
	}
	return out
}

func pptRecord(recType uint16, payload []byte) []byte {
	out := make([]byte, 8+len(payload))
	binary.LittleEndian.PutUint16(out[2:], recType)
	binary.LittleEndian.PutUint32(out[4:], uint32(len(payload)))
	copy(out[8:], payload)
	return out
}

func pptContainerRecord(recType uint16, payload []byte) []byte {
	out := pptRecord(recType, payload)
	binary.LittleEndian.PutUint16(out[0:], 0x000f)
	return out
}

func writeBIFFRecord(w *bytes.Buffer, id uint16, payload []byte) {
	var hdr [4]byte
	binary.LittleEndian.PutUint16(hdr[0:], id)
	binary.LittleEndian.PutUint16(hdr[2:], uint16(len(payload)))
	w.Write(hdr[:])
	w.Write(payload)
}

func testLabelSSTRecord(row, col uint16, idx uint32) []byte {
	out := make([]byte, 10)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	binary.LittleEndian.PutUint32(out[6:], idx)
	return out
}

func testBIFFLabelRecord(row, col uint16, s string) []byte {
	out := make([]byte, 6)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	return append(out, testXLUnicodeString(s)...)
}

func testNumberRecord(row, col uint16, value float64) []byte {
	out := make([]byte, 14)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	binary.LittleEndian.PutUint64(out[6:], math.Float64bits(value))
	return out
}

func testRKRecord(row, col uint16, value int32) []byte {
	out := make([]byte, 10)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	binary.LittleEndian.PutUint32(out[6:], uint32(value<<2)|0x02)
	return out
}

func testBoolErrRecord(row, col uint16, value byte, isErr bool) []byte {
	out := make([]byte, 8)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	out[6] = value
	if isErr {
		out[7] = 1
	}
	return out
}

func testFormulaNumberRecord(row, col uint16, value float64) []byte {
	out := make([]byte, 20)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	binary.LittleEndian.PutUint64(out[6:], math.Float64bits(value))
	return out
}

func testFormulaSpecialRecord(row, col uint16, resultType, value byte) []byte {
	out := make([]byte, 20)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	out[6] = resultType
	out[8] = value
	out[12] = 0xff
	out[13] = 0xff
	return out
}

func testFormulaStringRecord(row, col uint16) []byte {
	out := make([]byte, 20)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	out[6] = 0x00
	out[12] = 0xff
	out[13] = 0xff
	return out
}

func testRStringRecord(row, col uint16, s string) []byte {
	out := make([]byte, 0, 6+3+2+len(s)+4)
	out = append(out, make([]byte, 6)...)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	var hdr [5]byte
	binary.LittleEndian.PutUint16(hdr[0:], uint16(len(s)))
	hdr[2] = 0x08
	binary.LittleEndian.PutUint16(hdr[3:], 1)
	out = append(out, hdr[:]...)
	out = append(out, []byte(s)...)
	out = append(out, 0, 0, 0, 0)
	return out
}

func testMulRKRecord(row, firstCol uint16, values []int32) []byte {
	out := make([]byte, 4+len(values)*6+2)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], firstCol)
	for i, value := range values {
		pos := 4 + i*6
		binary.LittleEndian.PutUint32(out[pos+2:], uint32(value<<2)|0x02)
	}
	binary.LittleEndian.PutUint16(out[len(out)-2:], firstCol+uint16(len(values))-1)
	return out
}

func testBIFFNoteRecord(row, col uint16) []byte {
	out := make([]byte, 12)
	binary.LittleEndian.PutUint16(out[0:], row)
	binary.LittleEndian.PutUint16(out[2:], col)
	binary.LittleEndian.PutUint16(out[4:], 0x0002)
	binary.LittleEndian.PutUint16(out[6:], 1)
	return out
}

func testBIFFTXORecord(s string) []byte {
	return testBIFFTXORecordWithLen(len(s))
}

func testBIFFTXORecordWithLen(chars int) []byte {
	out := make([]byte, 18)
	binary.LittleEndian.PutUint16(out[10:], uint16(chars))
	binary.LittleEndian.PutUint16(out[12:], 16)
	return out
}

func testBIFFTXOContinue(s string) []byte {
	return append([]byte{0}, []byte(s)...)
}

func testBIFFTXOUnicodeContinue(s string) []byte {
	return append([]byte{1}, utf16LEBytes(s)...)
}

func testXLUnicodeString(s string) []byte {
	out := make([]byte, 3+len(s))
	binary.LittleEndian.PutUint16(out, uint16(len(s)))
	out[2] = 0
	copy(out[3:], []byte(s))
	return out
}

func testXLUnicodeRichExtString(s string, ext []byte) []byte {
	out := make([]byte, 0, 3+2+4+len(s)+4+len(ext))
	var hdr [9]byte
	binary.LittleEndian.PutUint16(hdr[0:], uint16(len(s)))
	hdr[2] = 0x0c
	binary.LittleEndian.PutUint16(hdr[3:], 1)
	binary.LittleEndian.PutUint32(hdr[5:], uint32(len(ext)))
	out = append(out, hdr[:]...)
	out = append(out, []byte(s)...)
	out = append(out, 0, 0, 0, 0)
	out = append(out, ext...)
	return out
}

func testWordPieceTableDocument(asciiText, unicodeText string) ([]byte, []byte) {
	word := make([]byte, 1024)
	binary.LittleEndian.PutUint16(word, 0xa5ec)
	binary.LittleEndian.PutUint16(word[0x0a:], 0x0200)
	asciiOffset := 512
	copy(word[asciiOffset:], []byte(asciiText))
	unicodeRaw := utf16LEBytes(unicodeText)
	unicodeOffset := asciiOffset + len(asciiText) + 32
	copy(word[unicodeOffset:], unicodeRaw)

	asciiChars := len([]rune(asciiText))
	unicodeChars := len([]rune(unicodeText))
	cp := []uint32{0, uint32(asciiChars), uint32(asciiChars + unicodeChars)}
	pieces := 2
	plc := make([]byte, (pieces+1)*4+pieces*8)
	for i, v := range cp {
		binary.LittleEndian.PutUint32(plc[i*4:], v)
	}
	pcdOff := (pieces + 1) * 4
	binary.LittleEndian.PutUint32(plc[pcdOff+2:], uint32(asciiOffset*2)|0x40000000)
	binary.LittleEndian.PutUint32(plc[pcdOff+8+2:], uint32(unicodeOffset))

	clx := make([]byte, 5+len(plc))
	clx[0] = 0x02
	binary.LittleEndian.PutUint32(clx[1:], uint32(len(plc)))
	copy(clx[5:], plc)
	table := make([]byte, 64+len(clx))
	copy(table[64:], clx)
	binary.LittleEndian.PutUint32(word[0x01a2:], 64)
	binary.LittleEndian.PutUint32(word[0x01a6:], uint32(len(clx)))
	return word, table
}

func TestCleanTextDropsInvalidUTF8(t *testing.T) {
	got := cleanText(string([]byte{'o', 'k', 0xff, ' ', 't', 'e', 'x', 't'}))
	if got != "ok text" {
		t.Fatalf("unexpected cleaned text %q", got)
	}
	if strings.ContainsRune(got, '\uFFFD') {
		t.Fatalf("cleaned text contains replacement rune: %q", got)
	}
}

func TestCleanTextExtractsVisibleRTFText(t *testing.T) {
	rtf := `{\rtf1\ansi{\fonttbl{\f0 Arial;}}{\colortbl;\red255\green0\blue0;}{\info{\title Hidden Title}}\pard Visible line\par Second \'93quoted\'94 line\par Unicode \u20320? text\par Emoji \u-10179?\u-8704? end\par Binary \bin13 HIDDENPAYLOAD after{\pict\pngblip hidden image bytes}}`
	got := cleanText(rtf)
	for _, want := range []string{"Visible line", "Second \u201cquoted\u201d line", "Unicode 你 text", "Emoji 😀 end"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF text %q in %q", want, got)
		}
	}
	if !strings.Contains(got, "Binary after") {
		t.Fatalf("cleanText missing visible RTF text after binary payload in %q", got)
	}
	for _, hidden := range []string{"rtf1", "fonttbl", "colortbl", "Hidden Title", "pict", "hidden image bytes", `\par`, `\u20320`, `\u-10179`, `\u-8704`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF control/internal text %q in %q", hidden, got)
		}
	}
	for _, hidden := range []string{"HIDDENPAYLOAD", `\bin13`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF binary payload/control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextDoesNotTreatVisibleBackslashPathAsRTF(t *testing.T) {
	text := `Visible path C:\Reports\Q1 remains text`
	if got := cleanText(text); got != text {
		t.Fatalf("cleanText changed visible backslash path: %q", got)
	}
}

func TestCleanTextPrefersRTFUnicodeAlternateDestination(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Start {\upr{ANSI fallback}{\*\ud Unicode \u20320? text}} End {\*\unknown Hidden Secret}}`
	got := cleanText(rtf)
	for _, want := range []string{"Start", "Unicode " + string(rune(20320)) + " text", "End"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF alternate text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"ANSI fallback", "Hidden Secret", "upr", "ud", "unknown", `\u20320`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept hidden RTF alternate/internal text %q in %q", hidden, got)
		}
	}
}

func TestCleanTextDecodesRTFAnsiCodePageHexText(t *testing.T) {
	rtf := `{\rtf1\ansi\ansicpg936\pard GBK \'d6\'d0\'ce\'c4 text\par Hidden {\fonttbl \'d6\'d0}}`
	got := cleanText(rtf)
	for _, want := range []string{"GBK 中文 text"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing RTF codepage text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{`\'d6`, "ÖÐ", "fonttbl"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF codepage/internal text %q in %q", hidden, got)
		}
	}
}

func TestCleanTextDecodesRTFFontCharsetHexText(t *testing.T) {
	rtf := `{\rtf1\ansi{\fonttbl{\f0\fcharset134 SimSun;}}\pard Font charset \'d6\'d0\'ce\'c4 text}`
	got := cleanText(rtf)
	if !strings.Contains(got, "Font charset 中文 text") {
		t.Fatalf("cleanText missing RTF font charset text in %q", got)
	}
	for _, hidden := range []string{"fonttbl", "fcharset", "SimSun", `\'d6`, "ÖÐ"} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF font charset/internal text %q in %q", hidden, got)
		}
	}
}

func TestCleanTextKeepsVisibleRTFHeaderFooterAndFootnote(t *testing.T) {
	rtf := `{\rtf1\ansi{\header Visible Header\par}{\footer Visible Footer\par}\pard Visible Body{\footnote Visible Footnote}\par{\info{\title Hidden Title}}}`
	got := cleanText(rtf)
	for _, want := range []string{"Visible Header", "Visible Body", "Visible Footnote", "Visible Footer"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF content %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"Hidden Title", "info", "rtf1", `\header`, `\footer`, `\footnote`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF internal/control text %q in %q", hidden, got)
		}
	}
}

func TestCleanTextDropsRTFHiddenText(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Visible before \v Hidden vanish secret\v0 Visible after {\v Hidden group secret} \v Hidden until plain \plain Visible plain tail \webhidden Hidden web secret\webhidden0 Visible web tail}`
	got := cleanText(rtf)
	for _, want := range []string{"Visible before", "Visible after", "Visible plain tail", "Visible web tail"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"Hidden vanish secret", "Hidden group secret", "Hidden until plain", "Hidden web secret", `\v`, `\webhidden`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept hidden RTF text/control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextDropsRTFDeletedRevisionText(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Visible before {\deleted Deleted revision secret} Visible middle \deleted Hidden inline deletion \deleted0 Visible after\par{\deleted Deleted paragraph secret\par}Tail visible}`
	got := cleanText(rtf)
	for _, want := range []string{"Visible before", "Visible middle", "Visible after", "Tail visible"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF revision text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"Deleted revision secret", "Hidden inline deletion", "Deleted paragraph secret", `\deleted`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept deleted RTF revision text/control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextKeepsRTFFieldResultOnly(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Before {\field{\*\fldinst HYPERLINK "https://example.test/hidden-target"}{\fldrslt Shown Link Text}} After}`
	got := cleanText(rtf)
	for _, want := range []string{"Before", "Shown Link Text", "After"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF field result %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"HYPERLINK", "https://example.test/hidden-target", "fldinst", "fldrslt", `\field`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF field instruction/control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextKeepsRTFAnnotationTextOnly(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Body text{\annotation Visible Comment Text{\*\atnauthor Hidden Author}{\*\atnid HiddenId}{\atntime Hidden Time}{\atnref Hidden Ref}{\atnparent Hidden Parent}{\atnstatus Hidden Status}{\atnicn Hidden Icon}} tail}`
	got := cleanText(rtf)
	for _, want := range []string{"Body text", "Visible Comment Text", "tail"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF annotation text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"Hidden Author", "HiddenId", "Hidden Time", "Hidden Ref", "Hidden Parent", "Hidden Status", "Hidden Icon", "atnauthor", "atnid", "atntime", "atnref", "atnparent", "atnstatus", "atnicn", `\annotation`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF annotation metadata/control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextKeepsRTFVisibleSymbolsAndTableSeparators(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Quote \lquote single\rquote  \ldblquote double\rdblquote  Dash \emdash endash\endash  List \bullet item\par Non\_breaking soft\-hyphen A\enspace B\emspace C\qmspace D\par Cell A\cell Cell B\cell\row Page Tail\page Section Tail\sect Column Tail\column End}`
	got := cleanText(rtf)
	for _, want := range []string{"Quote 'single'", `"double"`, "Dash - endash -", "List * item", "Non-breaking softhyphen A B C D", "Cell A Cell B", "Page Tail", "Section Tail", "Column Tail", "End"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF symbol/table text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{`lquote`, `rdblquote`, `emdash`, `bullet`, `\cell`, `\row`, `\page`, `\sect`, `\column`, `\_`, `\-`, `enspace`, `emspace`, `qmspace`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF visible-symbol control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextKeepsRTFAutomaticReferenceMarks(t *testing.T) {
	rtf := `{\rtf1\ansi\pard Body\chftn text\chatn comment{\footnote \chftn Footnote body}{\annotation \chatn Comment body}}`
	got := cleanText(rtf)
	for _, want := range []string{"Body[footnote] text[comment] comment", "[footnote] Footnote body", "[comment] Comment body"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF automatic mark text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{`chftn`, `chatn`, `\footnote`, `\annotation`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF automatic mark control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextKeepsRTFOptionalListText(t *testing.T) {
	rtf := `{\rtf1\ansi{\*\listtable Hidden List Table}{\*\pntext\bullet\tab}Bullet item\par{\*\listtext 1.\tab}Numbered item}`
	got := cleanText(rtf)
	for _, want := range []string{"* Bullet item", "1. Numbered item"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing visible RTF list text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"Hidden List Table", "listtable", "pntext", "listtext", `\tab`, `\bullet`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept RTF list control/internal text %q in %q", hidden, got)
		}
	}
}

func TestMarkdownOutputExtractsVisibleRTFText(t *testing.T) {
	res := &Result{Text: `{\rtf1\ansi{\fonttbl{\f0 Arial;}}\pard Visible RTF body\chftn\par \v Hidden markdown secret\v0 Second line\par {\field{\*\fldinst HYPERLINK "https://example.test/hidden"}{\fldrslt Visible markdown field}}\par {\annotation Visible markdown comment{\*\atnauthor Hidden markdown author}}\par {\deleted Deleted markdown revision}\par {\*\pntext\bullet\tab}Visible markdown list\par Quote \ldblquote visible\rdblquote \emdash tail}`}
	md := res.Markdown("images")
	for _, want := range []string{"Visible RTF body[footnote]", "Second line", "Visible markdown field", "Visible markdown comment", "* Visible markdown list", `Quote "visible" - tail`} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing visible RTF text %q in:\n%s", want, md)
		}
	}
	for _, hidden := range []string{"rtf1", "fonttbl", "pard", `\par`, "Hidden markdown secret", `\v`, "HYPERLINK", "https://example.test/hidden", "fldinst", "Hidden markdown author", "atnauthor", "Deleted markdown revision", "deleted", "ldblquote", "emdash", "chftn", "pntext"} {
		if strings.Contains(md, hidden) {
			t.Fatalf("markdown kept RTF control text %q in:\n%s", hidden, md)
		}
	}
}

func TestCleanTextNormalizesUnicodeSpaces(t *testing.T) {
	input := "Alpha\u00a0\u202fBeta\u3000Gamma\n中文\u2003空格\tTail"
	got := cleanText(input)
	want := "Alpha Beta Gamma\n中文 空格 Tail"
	if got != want {
		t.Fatalf("cleanText unicode spaces got %q, want %q", got, want)
	}
	for _, r := range []rune{'\u00a0', '\u202f', '\u3000', '\u2003'} {
		if strings.ContainsRune(got, r) {
			t.Fatalf("cleanText kept unicode space %U in %q", r, got)
		}
	}
}

func TestCleanTextDropsInvisibleVariationControls(t *testing.T) {
	input := "Report\ufe0f Text\u034f Body\U000e0100 Tail\u180b End"
	got := cleanText(input)
	want := "Report Text Body Tail End"
	if got != want {
		t.Fatalf("cleanText invisible variation controls got %q, want %q", got, want)
	}
	for _, r := range []rune{'\ufe0f', '\u034f', '\U000e0100', '\u180b'} {
		if strings.ContainsRune(got, r) {
			t.Fatalf("cleanText kept invisible variation/control rune %U in %q", r, got)
		}
	}
}

func TestCleanTextDropsEscapedInvisibleVariationControls(t *testing.T) {
	got := cleanText(`{\rtf1\ansi\pard RTF\u65039? Body\par}` + "\nOOXML_xFE0F_Text_x034F_Tail")
	for _, want := range []string{"RTF Body", "OOXMLTextTail"} {
		if !strings.Contains(got, want) {
			t.Fatalf("cleanText missing cleaned text %q in %q", want, got)
		}
	}
	for _, hidden := range []string{"\ufe0f", "\u034f", "_xFE0F_", "_x034F_", `\u65039`} {
		if strings.Contains(got, hidden) {
			t.Fatalf("cleanText kept escaped invisible control %q in %q", hidden, got)
		}
	}
}

func TestCleanTextDropsShortCyrillicMojibakeControlLineWithSpaces(t *testing.T) {
	noise := "0 000еяя$([\\{bябя"
	if got := cleanText(noise); got != "" {
		t.Fatalf("cleanText kept short Cyrillic mojibake control line: %q", got)
	}
	visible := "Проект 2024: этап 1 готов"
	if got := cleanText(visible); got != visible {
		t.Fatalf("cleanText dropped visible Cyrillic text: %q", got)
	}
}

func TestCleanTextDropsDocxCyrillicMojibakeControlLine(t *testing.T) {
	noise := "00яяы0яяяя›0њ0э0ю00ћ00ь0 я0=я]я 00"
	if got := cleanText(noise); got != "" {
		t.Fatalf("cleanText kept DOCX Cyrillic mojibake control line: %q", got)
	}
}

func TestCleanTextKeepsMarkdownTableRowsWithCyrillicText(t *testing.T) {
	row := "| 12 | 110 лет со дня рождения советского ученого, конструктора Сергея Павловича Королева |"
	for name, got := range map[string]string{
		"cleanText":                cleanText(row),
		"cleanVisibleText":         cleanVisibleText(row),
		"cleanMarkdownVisibleText": cleanMarkdownVisibleText(row),
	} {
		if !strings.Contains(got, "Сергея Павловича Королева") {
			t.Fatalf("%s dropped visible Cyrillic markdown table row: %q", name, got)
		}
	}
	noise := "00яяы0яяяя›0њ0э0ю00ћ00ь0 я0=я]я 00"
	if got := cleanText(noise); got != "" {
		t.Fatalf("cleanText kept non-table Cyrillic mojibake control line: %q", got)
	}
}

func TestCleanTextKeepsMarkdownListRowsWithCyrillicText(t *testing.T) {
	line := "1. В случае, когда функционал виртуальных разделов не применяется столбцы 11 и 12 Таблицы программирования ИС не заполняются."
	for name, got := range map[string]string{
		"cleanText":                cleanText(line),
		"cleanVisibleText":         cleanVisibleText(line),
		"cleanMarkdownVisibleText": cleanMarkdownVisibleText(markdownListIndentPrefix(5) + line),
	} {
		if !strings.Contains(got, "виртуальных разделов не применяется") {
			t.Fatalf("%s dropped visible Cyrillic markdown list row: %q", name, got)
		}
	}
	noise := "1. 00яяы0яяяя›0њ0э0ю00ћ00ь0 я0=я]я 00"
	if got := cleanText(noise); got != "" {
		t.Fatalf("cleanText kept listed Cyrillic mojibake control line: %q", got)
	}
}

func TestStripInlineHiddenOfficeReferencesFastGuard(t *testing.T) {
	visible := "Visible ratio: 3/4 and score=5 (draft) remain visible."
	if got := stripInlineHiddenOfficeReferences(visible); got != visible {
		t.Fatalf("fast guard changed visible prose: %q", got)
	}
	hidden := `Visible before Target="../media/inline.png" and Content-Type: image/png after`
	got := stripInlineHiddenOfficeReferences(hidden)
	if strings.Contains(got, `Target="../media/inline.png"`) || strings.Contains(got, "Content-Type: image/png") {
		t.Fatalf("fast guard kept hidden office references: %q", got)
	}
	if !strings.Contains(got, "Visible before") || !strings.Contains(got, "after") {
		t.Fatalf("fast guard removed visible context: %q", got)
	}
}

func TestMarkdownOutputNormalizesUnicodeSpaces(t *testing.T) {
	res := &Result{
		Text: "Visible\u00a0Body\nTable\u202fCell",
		StructuredMarkdown: strings.Join([]string{
			"## Sheet\u3000Name",
			"",
			"| A |",
			"|---|",
			"| Table\u202fCell |",
		}, "\n"),
	}
	md := res.Markdown("images")
	for _, want := range []string{"## Sheet Name", "| Table Cell |", "Visible Body"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing normalized text %q in:\n%s", want, md)
		}
	}
	for _, r := range []rune{'\u00a0', '\u202f', '\u3000'} {
		if strings.ContainsRune(md, r) {
			t.Fatalf("markdown kept unicode space %U in:\n%s", r, md)
		}
	}
}

func TestMarkdownOutputDropsInvisibleFormatControls(t *testing.T) {
	res := &Result{
		Text: "Visible\u200bBody\ufe0f",
		StructuredMarkdown: strings.Join([]string{
			"## \ufeff报告\u202e Title",
			"",
			"    Code\u200b Block\u034f",
			"",
			"  - Nested\ufeff Item 😀",
			"",
			"Visible\u200bBody",
		}, "\n"),
	}
	md := res.Markdown("images")
	for _, want := range []string{"## 报告 Title", "    Code Block", "  - Nested Item 😀", "VisibleBody"} {
		if !strings.Contains(md, want) {
			t.Fatalf("markdown missing cleaned visible text %q in:\n%s", want, md)
		}
	}
	for _, r := range []rune{'\u200b', '\ufeff', '\u202e', '\ue000', '\ufe0f', '\u034f'} {
		if strings.ContainsRune(md, r) {
			t.Fatalf("markdown kept invisible format/control rune %U in:\n%s", r, md)
		}
	}
}

func addZip(t *testing.T, zw *zip.Writer, name, text string) {
	t.Helper()
	addZipBytes(t, zw, name, []byte(text))
}

func TestExtractXLSXMalformedSharedStringFastPathFallsBackToDecoder(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/sharedStrings.xml", `<sst xmlns="urn:x"><si><t>Visible shared text</t></si></sst>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "shared-fastpath-fallback.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}

	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(res.Text, "Visible shared text") {
		t.Fatalf("expected fallback decoder text, got %q", res.Text)
	}
}

func TestExtractXLSXMalformedSimpleInlineFastPathFallsBackToDecoder(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Visible inline text</t></is></c><c r="B1" t="inlineStr"><is><t>Second visible text</t></is></c></row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	file := filepath.Join(dir, "inline-fastpath-fallback.xlsx")
	if err := os.WriteFile(file, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}

	res, err := Extract(file, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"Visible inline text", "Second visible text"} {
		if !strings.Contains(res.Text, want) {
			t.Fatalf("expected fallback decoder text %q in %q", want, res.Text)
		}
	}
}

func corruptStoredZipEntryData(t *testing.T, data []byte, name string) []byte {
	t.Helper()
	out := append([]byte(nil), data...)
	zr, err := zip.NewReader(bytes.NewReader(out), int64(len(out)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name != name {
			continue
		}
		if f.Method != zip.Store {
			t.Fatalf("expected stored zip entry for %s, got method %d", name, f.Method)
		}
		off, err := f.DataOffset()
		if err != nil {
			t.Fatal(err)
		}
		if off < 0 || off >= int64(len(out)) {
			t.Fatalf("invalid data offset %d for %s", off, name)
		}
		out[off] ^= 0x01
		return out
	}
	t.Fatalf("zip entry %s not found", name)
	return nil
}

func testDocxPackage(t *testing.T, text string, png []byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "word/document.xml", `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>`+text+`</w:t></w:r></w:p></w:body></w:document>`)
	if png != nil {
		addZipBytes(t, zw, "word/media/image1.png", png)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func testXlsxPackage(t *testing.T, values ...string) []byte {
	t.Helper()
	var cells strings.Builder
	for i, value := range values {
		col := string(rune('A' + i))
		cells.WriteString(`<c r="` + col + `1" t="inlineStr"><is><t>`)
		cells.WriteString(value)
		cells.WriteString(`</t></is></c>`)
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	addZip(t, zw, "[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`)
	addZip(t, zw, "xl/workbook.xml", `<workbook xmlns="urn:x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>`)
	addZip(t, zw, "xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`)
	addZip(t, zw, "xl/worksheets/sheet1.xml", `<worksheet xmlns="urn:x"><sheetData><row r="1">`+cells.String()+`</row></sheetData></worksheet>`)
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func addZipBytes(t *testing.T, zw *zip.Writer, name string, data []byte) {
	t.Helper()
	w, err := zw.Create(name)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Write(data); err != nil {
		t.Fatal(err)
	}
}
