package xlrd

import (
	"os"
	"testing"
)

func TestMissingRecordsDefaultFormat(t *testing.T) {
	book, err := OpenWorkbook(fromSample("biff4_no_format_no_window2.xls"), &OpenWorkbookOptions{Verbosity: 2, Logfile: os.Stdout})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	cell := sheet.Cell(0, 0)
	if cell.CType != XL_CELL_TEXT {
		t.Errorf("cell.CType = %d, want %d", cell.CType, XL_CELL_TEXT)
	}
}

func TestMissingRecordsDefaultWindow2Options(t *testing.T) {
	book, err := OpenWorkbook(fromSample("biff4_no_format_no_window2.xls"), &OpenWorkbookOptions{Verbosity: 2, Logfile: os.Stdout})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Test default window2 options
	if sheet.CachedPageBreakPreviewMagFactor != 0 {
		t.Errorf("sheet.CachedPageBreakPreviewMagFactor = %d, want 0", sheet.CachedPageBreakPreviewMagFactor)
	}
	if sheet.CachedNormalViewMagFactor != 0 {
		t.Errorf("sheet.CachedNormalViewMagFactor = %d, want 0", sheet.CachedNormalViewMagFactor)
	}
}
