package officeread

import (
	"encoding/binary"
	"os"
	"path/filepath"
	"testing"
)

func TestPPTActivePersistOffsetsUsesCurrentUserDirectory(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000008.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	active := pptActivePersistOffsets(streams)
	if len(active) != 32 {
		t.Fatalf("active persist entries = %d, want 32", len(active))
	}
	if active[1] != 0 || active[40] != 24152 || active[81] != 108406 {
		t.Fatalf("unexpected active persist mapping: %#v", active)
	}
}

func TestPPTActivePersistRecordTypesDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	doc, ok := findLegacyStream(streams, "PowerPoint Document")
	if !ok {
		t.Fatal("missing PowerPoint Document")
	}
	types := map[uint16]int{}
	for _, offset := range pptActivePersistOffsets(streams) {
		off := int(offset)
		if off+8 <= len(doc.Data) {
			types[binary.LittleEndian.Uint16(doc.Data[off+2:])]++
		}
	}
	t.Logf("active persist record types: %#v", types)
}

func TestPPTDocumentSlidePersistDiagnostic(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("testdata", "web-samples", "samples", "ppt", "000133.ppt"))
	if err != nil {
		t.Fatal(err)
	}
	streams, err := readOLEStreams(data)
	if err != nil {
		t.Fatal(err)
	}
	doc, ok := findLegacyStream(streams, "PowerPoint Document")
	if !ok {
		t.Fatal("missing PowerPoint Document")
	}
	active := pptActivePersistOffsets(streams)
	off := int(active[1])
	if off+8 > len(doc.Data) {
		t.Fatal("invalid document persist offset")
	}
	var walk func([]byte, int)
	walk = func(data []byte, depth int) {
		for pos := 0; pos+8 <= len(data); {
			options := binary.LittleEndian.Uint16(data[pos:])
			typ := binary.LittleEndian.Uint16(data[pos+2:])
			size := int(binary.LittleEndian.Uint32(data[pos+4:]))
			pos += 8
			if size < 0 || size > len(data)-pos {
				return
			}
			payload := data[pos : pos+size]
			if typ == 0x03f3 && len(payload) >= 4 {
				id := binary.LittleEndian.Uint32(payload)
				if target, ok := active[id]; ok {
					targetOff := int(target)
					if targetOff+8 <= len(doc.Data) {
						t.Logf("SlidePersistAtom depth=%d persistIDRef=%d -> type=%#x offset=%d", depth, id, binary.LittleEndian.Uint16(doc.Data[targetOff+2:]), target)
					}
				}
			}
			if options&0x000f == 0x000f && depth < 8 {
				walk(payload, depth+1)
			}
			pos += size
		}
	}
	size := int(binary.LittleEndian.Uint32(doc.Data[off+4:]))
	walk(doc.Data[off+8:off+8+size], 0)
}
