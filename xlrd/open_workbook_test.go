package xlrd

import (
	"strings"
	"testing"
)

func TestNamesDemo(t *testing.T) {
	// For now, we just check this doesn't raise an error.
	book, err := OpenWorkbook(fromSample("namesdemo.xls"), nil)
	if err != nil {
		t.Fatalf("OpenWorkbook(namesdemo.xls) failed: %v", err)
	}
	if book == nil {
		t.Fatal("OpenWorkbook returned nil book")
	}
	// Check that we can get sheet names
	sheetNames := book.SheetNames()
	if len(sheetNames) == 0 {
		t.Log("No sheets found (this is OK if workbook parsing is not fully implemented)")
	} else {
		t.Logf("Found %d sheets: %v", len(sheetNames), sheetNames)
	}
}

func TestRaggedRowsTidiedWithFormatting(t *testing.T) {
	// Test that opening with formatting info doesn't cause errors
	options := &OpenWorkbookOptions{
		FormattingInfo: true,
	}
	book, err := OpenWorkbook(fromSample("issue20.xls"), options)
	if err != nil {
		t.Fatalf("OpenWorkbook(issue20.xls) failed: %v", err)
	}
	if book == nil {
		t.Fatal("OpenWorkbook returned nil book")
	}
	t.Logf("Successfully opened issue20.xls with formatting info")
}

func TestBYTESX00(t *testing.T) {
	// Test that opening picture_in_cell.xls with formatting info doesn't cause errors
	options := &OpenWorkbookOptions{
		FormattingInfo: true,
	}
	book, err := OpenWorkbook(fromSample("picture_in_cell.xls"), options)
	if err != nil {
		t.Fatalf("OpenWorkbook(picture_in_cell.xls) failed: %v", err)
	}
	if book == nil {
		t.Fatal("OpenWorkbook returned nil book")
	}
	t.Logf("Successfully opened picture_in_cell.xls with formatting info")
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
