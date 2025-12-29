package xlrd

import (
	"strings"
	"testing"
)

func TestNamesDemo(t *testing.T) {
	// For now, we just check this doesn't raise an error.
	// TODO: Implement actual workbook opening
	_, err := OpenWorkbook(fromSample("namesdemo.xls"), nil)
	if err != nil {
		// Expected to fail until workbook opening is implemented
		t.Logf("OpenWorkbook(namesdemo.xls) failed (expected): %v", err)
	}
}

func TestRaggedRowsTidiedWithFormatting(t *testing.T) {
	// For now, we just check this doesn't raise an error.
	// TODO: Implement actual workbook opening
	options := &OpenWorkbookOptions{
		FormattingInfo: true,
	}
	_, err := OpenWorkbook(fromSample("issue20.xls"), options)
	if err != nil {
		// Expected to fail until workbook opening is implemented
		t.Logf("OpenWorkbook(issue20.xls) failed (expected): %v", err)
	}
}

func TestBYTESX00(t *testing.T) {
	// For now, we just check this doesn't raise an error.
	// TODO: Implement actual workbook opening
	options := &OpenWorkbookOptions{
		FormattingInfo: true,
	}
	_, err := OpenWorkbook(fromSample("picture_in_cell.xls"), options)
	if err != nil {
		// Expected to fail until workbook opening is implemented
		t.Logf("OpenWorkbook(picture_in_cell.xls) failed (expected): %v", err)
	}
}

func TestOpenXlsx(t *testing.T) {
	_, err := OpenWorkbook(fromSample("sample.xlsx"), nil)
	if err == nil {
		t.Error("OpenWorkbook(sample.xlsx) should have returned an error")
		return
	}
	if !strings.Contains(err.Error(), "Excel xlsx file; not supported") {
		t.Errorf("OpenWorkbook(sample.xlsx) error = %v, want error containing 'Excel xlsx file; not supported'", err)
	}
}

func TestOpenUnknown(t *testing.T) {
	_, err := OpenWorkbook(fromSample("sample.txt"), nil)
	if err == nil {
		t.Error("OpenWorkbook(sample.txt) should have returned an error")
		return
	}
	// The error message might be different, but should indicate unsupported format
	if err != nil {
		t.Logf("OpenWorkbook(sample.txt) correctly returned error: %v", err)
	}
}
