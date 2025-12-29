package xlrd

import (
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func fromSample(filename string) string {
	// Get the directory of this test file
	_, testFile, _, _ := runtime.Caller(0)
	testDir := filepath.Dir(testFile)
	// Go up to project root (xlrd -> project root)
	projectRoot := filepath.Join(testDir, "..")
	return filepath.Join(projectRoot, "testdata", "samples", filename)
}

func TestXlsx(t *testing.T) {
	format, err := InspectFormat(fromSample("sample.xlsx"), nil)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "xlsx" {
		t.Errorf("InspectFormat(sample.xlsx) = %s, want xlsx", format)
	}
}

func TestXlsb(t *testing.T) {
	format, err := InspectFormat(fromSample("sample.xlsb"), nil)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "xlsb" {
		t.Errorf("InspectFormat(sample.xlsb) = %s, want xlsb", format)
	}
}

func TestOds(t *testing.T) {
	format, err := InspectFormat(fromSample("sample.ods"), nil)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "ods" {
		t.Errorf("InspectFormat(sample.ods) = %s, want ods", format)
	}
}

func TestZip(t *testing.T) {
	format, err := InspectFormat(fromSample("sample.zip"), nil)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "zip" {
		t.Errorf("InspectFormat(sample.zip) = %s, want zip", format)
	}
}

func TestXls(t *testing.T) {
	format, err := InspectFormat(fromSample("namesdemo.xls"), nil)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "xls" {
		t.Errorf("InspectFormat(namesdemo.xls) = %s, want xls", format)
	}
}

func TestContent(t *testing.T) {
	content, err := os.ReadFile(fromSample("sample.xlsx"))
	if err != nil {
		t.Fatalf("Failed to read sample.xlsx: %v", err)
	}
	format, err := InspectFormat("", content)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "xlsx" {
		t.Errorf("InspectFormat(content=sample.xlsx) = %s, want xlsx", format)
	}
}

func TestUnknown(t *testing.T) {
	format, err := InspectFormat(fromSample("sample.txt"), nil)
	if err != nil {
		t.Fatalf("InspectFormat error: %v", err)
	}
	if format != "" {
		t.Errorf("InspectFormat(sample.txt) = %s, want empty string", format)
	}
}
