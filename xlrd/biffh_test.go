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

func TestBiffTextFromNum(t *testing.T) {
	tests := []struct {
		input    int
		expected string
	}{
		{0, "(not BIFF)"},
		{20, "2.0"},
		{21, "2.1"},
		{30, "3"},
		{40, "4S"},
		{45, "4W"},
		{50, "5"},
		{70, "7"},
		{80, "8"},
		{85, "8X"},
		{99, "Unknown(99)"},
	}

	for _, test := range tests {
		result := BiffTextFromNum(test.input)
		if result != test.expected {
			t.Errorf("BiffTextFromNum(%d) = %s, expected %s", test.input, result, test.expected)
		}
	}
}

func TestErrorTextFromCode(t *testing.T) {
	tests := []struct {
		code     byte
		expected string
	}{
		{0x00, "#NULL!"},
		{0x07, "#DIV/0!"},
		{0x0F, "#VALUE!"},
		{0x17, "#REF!"},
		{0x1D, "#NAME?"},
		{0x24, "#NUM!"},
		{0x2A, "#N/A"},
	}

	for _, test := range tests {
		result := ErrorTextFromCode[test.code]
		if result != test.expected {
			t.Errorf("ErrorTextFromCode[0x%02x] = %s, expected %s", test.code, result, test.expected)
		}
	}
}

func TestIsCellOpcode(t *testing.T) {
	tests := []struct {
		code     int
		expected bool
	}{
		{XL_BOOLERR, true},
		{XL_FORMULA, true},
		{XL_LABELSST, true},
		{XL_NUMBER, true},
		{XL_RK, true},
		{XL_BOF, false},
		{XL_EOF, false},
		{0xFFFF, false},
	}

	for _, test := range tests {
		result := IsCellOpcode(test.code)
		if result != test.expected {
			t.Errorf("IsCellOpcode(%d) = %v, expected %v", test.code, result, test.expected)
		}
	}
}

func TestUnpackString(t *testing.T) {
	// Test Latin-1 encoding
	data := []byte{0x03, 0x00, 'a', 'b', 'c'} // length 3, "abc"
	result := unpack_string(data, 0, "latin_1", 2)
	expected := "abc"
	if result != expected {
		t.Errorf("unpack_string() = %s, expected %s", result, expected)
	}

	// Test UTF-16 encoding - length 2 means 4 bytes of data
	data2 := []byte{0x02, 0x00, 0x61, 0x00, 0x62, 0x00} // length 2, "ab" in UTF-16
	result2 := unpack_string(data2, 0, "utf_16_le", 2)
	expected2 := "ab"
	if result2 != expected2 {
		t.Errorf("unpack_string() = %s, expected %s", result2, expected2)
	}
}

func TestUnpackUnicode(t *testing.T) {
	// Test compressed (Latin-1) string
	data := []byte{0x03, 0x00, 0x00, 'a', 'b', 'c'} // length 3, options 0, "abc"
	result := unpack_unicode(data, 0, 2)
	expected := "abc"
	if result != expected {
		t.Errorf("unpack_unicode() = %s, expected %s", result, expected)
	}
}

func TestBaseObjectDump(t *testing.T) {
	var buf bytes.Buffer
	obj := &BaseObject{}
	obj.Dump(&buf, "Test Header", "Test Footer", 0)

	output := buf.String()
	if !strings.Contains(output, "Test Header") {
		t.Errorf("Dump output should contain header")
	}
	if !strings.Contains(output, "Test Footer") {
		t.Errorf("Dump output should contain footer")
	}
}

func TestCellRange(t *testing.T) {
	r := CellRange{FirstRow: 1, LastRow: 5, FirstCol: 2, LastCol: 8}
	if r.FirstRow != 1 || r.LastRow != 5 || r.FirstCol != 2 || r.LastCol != 8 {
		t.Errorf("CellRange values incorrect")
	}
}
