package officeread

import (
	"archive/zip"
	"bytes"
	"path/filepath"
	"strings"
	"testing"
)

var benchmarkExtractResult *Result
var benchmarkStringResult string
var benchmarkBoolResult bool

func benchmarkExtractSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		res, err := Extract(samplePath, Options{})
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkExtractResult = res
	}
}

func benchmarkXLSMarkdownBackfillSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	res, err := Extract(samplePath, Options{})
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		benchmarkStringResult = missingMarkdownTextXLS(res.StructuredMarkdown, res.Text)
	}
}

func benchmarkXLSMarkdownExactSetSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	res, err := Extract(samplePath, Options{})
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		benchmarkStringResult = ""
		for line := range markdownBackfillExactSet(res.StructuredMarkdown) {
			benchmarkStringResult = line
			break
		}
	}
}

func benchmarkXLSMarkdownCoverageContainmentSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	res, err := Extract(samplePath, Options{})
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		coverage, containment := markdownBackfillBuildCoverageContainment(res.StructuredMarkdown)
		benchmarkStringResult = containment.visibleJoined
		for line := range coverage {
			benchmarkStringResult = line
			break
		}
	}
}

func benchmarkXLSVisibleContainmentReplaySampleFile(b *testing.B, samplePath string) {
	b.Helper()
	res, err := Extract(samplePath, Options{})
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	_, containment := markdownBackfillBuildCoverageContainment(res.StructuredMarkdown)
	queries := benchmarkXLSContainmentReplayQueries(res.Text, nil, true)
	if len(queries) == 0 {
		b.Skipf("%s: no visible containment queries collected", samplePath)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		hits := 0
		for _, q := range queries {
			if containment.visibleLineContainsLine(q) {
				hits++
			}
		}
		benchmarkStringResult = queries[hits%len(queries)]
	}
}

func benchmarkXLSTableContainmentReplaySampleFile(b *testing.B, samplePath string) {
	b.Helper()
	res, err := Extract(samplePath, Options{})
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	_, containment := markdownBackfillBuildCoverageContainment(res.StructuredMarkdown)
	containment.shortTableExactBeforeMinLen = true
	queries := benchmarkXLSContainmentReplayQueries(res.Text, nil, true)
	if len(queries) == 0 {
		b.Skipf("%s: no table containment queries collected", samplePath)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		hits := 0
		for _, q := range queries {
			if containment.tableTextContainsLine(q) {
				hits++
			}
		}
		benchmarkStringResult = queries[hits%len(queries)]
	}
}

func benchmarkXLSContainmentReplayQueries(text string, images []Image, escapedTableOnlyWhenPipe bool) []string {
	lines := markdownBackfillSourceLines(markdownBackfillRawLines(text))
	if len(lines) == 0 {
		return nil
	}
	imageAlts := markdownImageAltSet(images)
	candidateLineCache := map[string]string{}
	candidateLine := func(raw string) string {
		if line, ok := candidateLineCache[raw]; ok {
			return line
		}
		line := markdownBackfillCandidateLine(raw)
		candidateLineCache[raw] = line
		return line
	}
	queries := make([]string, 0, len(lines))
	seen := map[string]struct{}{}
	for _, raw := range lines {
		line := candidateLine(raw)
		if line == "" || !markdownBackfillNormalizedLineAllowed(line, imageAlts) {
			continue
		}
		variants := []string{line}
		if visibleLine := markdownBackfillVisibleText(line); visibleLine != "" && visibleLine != line {
			variants = append(variants, visibleLine)
		}
		if markdownLine := markdownVisibleLineText(line); markdownLine != "" {
			duplicate := false
			for _, variant := range variants {
				if variant == markdownLine {
					duplicate = true
					break
				}
			}
			if !duplicate {
				variants = append(variants, markdownLine)
			}
		}
		if !escapedTableOnlyWhenPipe || strings.IndexByte(line, '|') >= 0 {
			if escapedVisibleLine := markdownBackfillVisibleText(escapeMarkdownTableCell(line)); escapedVisibleLine != "" {
				duplicate := false
				for _, variant := range variants {
					if variant == escapedVisibleLine {
						duplicate = true
						break
					}
				}
				if !duplicate {
					variants = append(variants, escapedVisibleLine)
				}
			}
		}
		for _, variant := range variants {
			if _, ok := seen[variant]; ok {
				continue
			}
			seen[variant] = struct{}{}
			queries = append(queries, variant)
		}
	}
	return queries
}

func benchmarkXLSXSimpleInlineSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	zr, err := zip.OpenReader(samplePath)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	defer zr.Close()
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	sheets, err := workbookVisibleSheets(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	var worksheetBytes []byte
	var markdownData xlsxWorksheetMarkdownData
	for _, sheet := range sheets {
		f := files[sheet.Path]
		if f == nil {
			continue
		}
		bb, err := readZipFile(f)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		if simpleInlineWorksheetCandidate(bb) {
			worksheetBytes = bb
			break
		}
	}
	if len(worksheetBytes) == 0 {
		b.Skipf("%s: no simple-inline worksheet found", samplePath)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var out strings.Builder
		markdownData = xlsxWorksheetMarkdownData{}
		err := appendSimpleInlineWorksheetTextPrepared(&out, worksheetBytes, &markdownData)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkStringResult = out.String()
	}
}

func benchmarkXLSXSimpleInlineTextOnlySampleFile(b *testing.B, samplePath string) {
	b.Helper()
	zr, err := zip.OpenReader(samplePath)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	defer zr.Close()
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	sheets, err := workbookVisibleSheets(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	var worksheetBytes []byte
	for _, sheet := range sheets {
		f := files[sheet.Path]
		if f == nil {
			continue
		}
		bb, err := readZipFile(f)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		if simpleInlineWorksheetCandidate(bb) {
			worksheetBytes = bb
			break
		}
	}
	if len(worksheetBytes) == 0 {
		b.Skipf("%s: no simple-inline worksheet found", samplePath)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var out strings.Builder
		err := appendSimpleInlineWorksheetTextPrepared(&out, worksheetBytes, nil)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkStringResult = out.String()
	}
}

func benchmarkXLSXWorksheetTextSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	zr, err := zip.OpenReader(samplePath)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	defer zr.Close()
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	sheets, err := workbookVisibleSheets(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	var worksheetBytes []byte
	var shared []string
	for _, sheet := range sheets {
		f := files[sheet.Path]
		if f == nil {
			continue
		}
		bb, err := readZipFile(f)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		worksheetBytes = bb
		break
	}
	if len(worksheetBytes) == 0 {
		b.Fatalf("%s: no visible worksheet found", samplePath)
	}
	shared, err = readSharedStrings(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var out strings.Builder
		md := xlsxWorksheetMarkdownData{}
		if err := appendWorksheetText(&out, worksheetBytes, shared, nil, &md); err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkStringResult = out.String()
	}
}

func benchmarkXLSXWorksheetTextNoMarkdownSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	zr, err := zip.OpenReader(samplePath)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	defer zr.Close()
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	sheets, err := workbookVisibleSheets(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	var worksheetBytes []byte
	var shared []string
	for _, sheet := range sheets {
		f := files[sheet.Path]
		if f == nil {
			continue
		}
		bb, err := readZipFile(f)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		worksheetBytes = bb
		break
	}
	if len(worksheetBytes) == 0 {
		b.Fatalf("%s: no visible worksheet found", samplePath)
	}
	shared, err = readSharedStrings(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var out strings.Builder
		if err := appendWorksheetText(&out, worksheetBytes, shared, nil, nil); err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkStringResult = out.String()
	}
}

func benchmarkXLSXSharedStringWorksheetSample(b *testing.B, samplePath string) ([]byte, []string) {
	b.Helper()
	zr, err := zip.OpenReader(samplePath)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	defer zr.Close()
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	sheets, err := workbookVisibleSheets(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	var worksheetBytes []byte
	for _, sheet := range sheets {
		f := files[sheet.Path]
		if f == nil {
			continue
		}
		bb, err := readZipFile(f)
		if err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		worksheetBytes = bb
		break
	}
	if len(worksheetBytes) == 0 {
		b.Fatalf("%s: no visible worksheet found", samplePath)
	}
	shared, err := readSharedStrings(files)
	if err != nil {
		b.Fatalf("%s: %v", samplePath, err)
	}
	if !sharedStringWorksheetCandidate(worksheetBytes) {
		b.Skipf("%s: no shared-string fast-path worksheet found", samplePath)
	}
	return worksheetBytes, shared
}

func benchmarkXLSXSharedStringCandidateSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	worksheetBytes, _ := benchmarkXLSXSharedStringWorksheetSample(b, samplePath)
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		benchmarkBoolResult = sharedStringWorksheetCandidate(worksheetBytes)
	}
}

func benchmarkXLSXSharedStringPreparedSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	worksheetBytes, shared := benchmarkXLSXSharedStringWorksheetSample(b, samplePath)
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var out strings.Builder
		md := xlsxWorksheetMarkdownData{}
		if err := appendSharedStringWorksheetTextPrepared(&out, worksheetBytes, shared, &md); err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkStringResult = out.String()
	}
}

func benchmarkXLSXSharedStringPreparedNoMarkdownSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	worksheetBytes, shared := benchmarkXLSXSharedStringWorksheetSample(b, samplePath)
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var out strings.Builder
		if err := appendSharedStringWorksheetTextPrepared(&out, worksheetBytes, shared, nil); err != nil {
			b.Fatalf("%s: %v", samplePath, err)
		}
		benchmarkStringResult = out.String()
	}
}

func collectXLSXSharedStringMarkdownValuesForBenchmark(b *testing.B, samplePath string) []string {
	b.Helper()
	worksheetBytes, shared := benchmarkXLSXSharedStringWorksheetSample(b, samplePath)
	var values []string
	nextRow := 1
	for pos := 0; ; {
		i := bytes.Index(worksheetBytes[pos:], []byte("<row"))
		if i < 0 {
			break
		}
		i += pos
		tagEnd := bytes.IndexByte(worksheetBytes[i:], '>')
		if tagEnd < 0 {
			b.Fatalf("%s: malformed row tag", samplePath)
		}
		tagEnd += i
		tag := worksheetBytes[i : tagEnd+1]
		if !xmlStartTagNameIs(tag, "row") {
			pos = tagEnd + 1
			continue
		}
		rowEndRel := bytes.Index(worksheetBytes[tagEnd+1:], []byte("</row>"))
		if rowEndRel < 0 {
			b.Fatalf("%s: malformed row end", samplePath)
		}
		rowEnd := tagEnd + 1 + rowEndRel
		rowIndex := nextRow
		if rowRef := xmlAttrBytes(tag, "r"); len(rowRef) > 0 {
			if n, ok := atoi(string(rowRef)); ok && n > 0 {
				rowIndex = n
			}
		}
		nextRow = rowIndex + 1
		if rowIndex > maxMarkdownTableRows {
			break
		}
		rowData := worksheetBytes[tagEnd+1 : rowEnd]
		nextCol := 1
		for cellPos := 0; ; {
			cellStart := bytes.Index(rowData[cellPos:], []byte("<c"))
			if cellStart < 0 {
				break
			}
			cellStart += cellPos
			cellTagEnd := bytes.IndexByte(rowData[cellStart:], '>')
			if cellTagEnd < 0 {
				b.Fatalf("%s: malformed cell tag", samplePath)
			}
			cellTagEnd += cellStart
			cellTag := rowData[cellStart : cellTagEnd+1]
			if !xmlStartTagNameIs(cellTag, "c") {
				cellPos = cellTagEnd + 1
				continue
			}
			cellRef := xmlAttrBytes(cellTag, "r")
			cellCol := nextCol
			if col, _, ok := cellRefIndexes(string(cellRef)); ok {
				cellCol = col
			}
			if cellCol < 1 {
				cellCol = 1
			}
			nextCol = cellCol + 1
			selfClosing := len(cellTag) >= 2 && cellTag[len(cellTag)-2] == '/'
			cellEnd := cellTagEnd
			if !selfClosing {
				cellEndRel := bytes.Index(rowData[cellTagEnd+1:], []byte("</c>"))
				if cellEndRel < 0 {
					b.Fatalf("%s: malformed cell end", samplePath)
				}
				cellEnd = cellTagEnd + 1 + cellEndRel
			}
			rawValue := ""
			if !selfClosing {
				if value, ok := worksheetCellVText(rowData[cellTagEnd+1 : cellEnd]); ok {
					rawValue = value
				}
			}
			value := strings.TrimSpace(rawValue)
			if bytes.Equal(xmlAttrBytes(cellTag, "t"), []byte("s")) {
				if idx, ok := atoi(value); ok && idx >= 0 && idx < len(shared) {
					value = shared[idx]
				}
			}
			if cellCol <= maxMarkdownTableCols && value != "" {
				values = append(values, value)
			}
			if selfClosing {
				cellPos = cellTagEnd + 1
			} else {
				cellPos = cellEnd + len("</c>")
			}
		}
		pos = rowEnd + len("</row>")
	}
	if len(values) == 0 {
		b.Skipf("%s: no shared-string markdown values collected", samplePath)
	}
	return values
}

func benchmarkXLSXSharedStringMarkdownCleanSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	values := collectXLSXSharedStringMarkdownValuesForBenchmark(b, samplePath)
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for _, value := range values {
			benchmarkStringResult = cleanMarkdownTableCellValue(value)
		}
	}
}

func benchmarkXLSXSharedStringMarkdownPrepareSampleFile(b *testing.B, samplePath string) {
	b.Helper()
	values := collectXLSXSharedStringMarkdownValuesForBenchmark(b, samplePath)
	cleaned := make([]string, 0, len(values))
	for _, value := range values {
		if cleanedValue := cleanMarkdownTableCellValue(value); cleanedValue != "" {
			cleaned = append(cleaned, cleanedValue)
		}
	}
	if len(cleaned) == 0 {
		b.Skipf("%s: no cleaned shared-string markdown values collected", samplePath)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for _, value := range cleaned {
			benchmarkStringResult = prepareMarkdownTableCellValue(value)
		}
	}
}

func BenchmarkExtractXLSHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xls")
	for _, name := range []string{
		"006087.xls",
		"008055.xls",
		"016161.xls",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkExtractSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSMarkdownBackfillHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xls")
	for _, name := range []string{
		"006087.xls",
		"008055.xls",
		"016161.xls",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSMarkdownBackfillSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSMarkdownExactSetHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xls")
	for _, name := range []string{
		"006087.xls",
		"008055.xls",
		"016161.xls",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSMarkdownExactSetSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSMarkdownCoverageContainmentHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xls")
	for _, name := range []string{
		"006087.xls",
		"008055.xls",
		"016161.xls",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSMarkdownCoverageContainmentSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSVisibleContainmentReplayHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xls")
	for _, name := range []string{
		"006087.xls",
		"008055.xls",
		"016161.xls",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSVisibleContainmentReplaySampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSTableContainmentReplayHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xls")
	for _, name := range []string{
		"006087.xls",
		"008055.xls",
		"016161.xls",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSTableContainmentReplaySampleFile(b, samplePath)
		})
	}
}

func BenchmarkExtractXLSXHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xlsx")
	for _, name := range []string{
		"apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx",
		"00012389.xlsx",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkExtractSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSXSimpleInlineHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xlsx")
	for _, name := range []string{
		"apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx",
		"00012389.xlsx",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSXSimpleInlineSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSXWorksheetTextHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xlsx")
	for _, name := range []string{
		"apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx",
		"00012389.xlsx",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSXWorksheetTextSampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSXSimpleInlineTextOnlyHotspots(b *testing.B) {
	sampleDir := filepath.Join("testdata", "web-samples", "samples", "xlsx")
	for _, name := range []string{
		"apache__tika__tika-parsers_tika-parsers-standard_tika-parsers-standard-modules_tika-parser-microsoft-module_src_test_resources_test-documents_testRecordSizeExceeded.xlsx",
		"00012389.xlsx",
	} {
		samplePath := filepath.Join(sampleDir, name)
		b.Run(name, func(b *testing.B) {
			benchmarkXLSXSimpleInlineTextOnlySampleFile(b, samplePath)
		})
	}
}

func BenchmarkXLSXWorksheetTextNoMarkdown00012389(b *testing.B) {
	samplePath := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00012389.xlsx")
	benchmarkXLSXWorksheetTextNoMarkdownSampleFile(b, samplePath)
}

func BenchmarkXLSXSharedStringCandidate00012389(b *testing.B) {
	samplePath := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00012389.xlsx")
	benchmarkXLSXSharedStringCandidateSampleFile(b, samplePath)
}

func BenchmarkXLSXSharedStringPrepared00012389(b *testing.B) {
	samplePath := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00012389.xlsx")
	benchmarkXLSXSharedStringPreparedSampleFile(b, samplePath)
}

func BenchmarkXLSXSharedStringPreparedNoMarkdown00012389(b *testing.B) {
	samplePath := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00012389.xlsx")
	benchmarkXLSXSharedStringPreparedNoMarkdownSampleFile(b, samplePath)
}

func BenchmarkXLSXSharedStringMarkdownClean00012389(b *testing.B) {
	samplePath := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00012389.xlsx")
	benchmarkXLSXSharedStringMarkdownCleanSampleFile(b, samplePath)
}

func BenchmarkXLSXSharedStringMarkdownPrepare00012389(b *testing.B) {
	samplePath := filepath.Join("testdata", "web-samples", "samples", "xlsx", "00012389.xlsx")
	benchmarkXLSXSharedStringMarkdownPrepareSampleFile(b, samplePath)
}
