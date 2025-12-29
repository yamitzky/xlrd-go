package xlrd

import (
	"bytes"
	"strings"
	"testing"
)

func TestHexCharDump(t *testing.T) {
	var buf bytes.Buffer
	data := []byte("abc\x00e\x01")
	HexCharDump(data, 0, 6, 0, &buf, false)
	s := buf.String()
	
	if !strings.Contains(s, "61 62 63 00 65 01") {
		t.Errorf("HexCharDump output should contain '61 62 63 00 65 01', got: %s", s)
	}
	if !strings.Contains(s, "abc~e?") {
		t.Errorf("HexCharDump output should contain 'abc~e?', got: %s", s)
	}
}
