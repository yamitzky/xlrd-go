package xlrd

import (
	"math"
	"testing"
)

func TestEmptyCell(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByName("TRAVERSALCHAINAGE")
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	cell := sheet.Cell(0, 0)
	if cell.CType != XL_CELL_EMPTY {
		t.Errorf("cell.CType = %d, want %d", cell.CType, XL_CELL_EMPTY)
	}
	if cell.Value != "" {
		t.Errorf("cell.Value = %v, want empty string", cell.Value)
	}
	if cell.XFIndex <= 0 {
		t.Errorf("cell.XFIndex = %d, want > 0", cell.XFIndex)
	}
}

func TestStringCell(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByName("PROFILEDEF")
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	cell := sheet.Cell(0, 0)
	if cell.CType != XL_CELL_TEXT {
		t.Errorf("cell.CType = %d, want %d", cell.CType, XL_CELL_TEXT)
	}
	if cell.Value != "PROFIL" {
		t.Errorf("cell.Value = %v, want 'PROFIL'", cell.Value)
	}
	if cell.XFIndex <= 0 {
		t.Errorf("cell.XFIndex = %d, want > 0", cell.XFIndex)
	}
}

func TestNumberCell(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByName("PROFILEDEF")
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	cell := sheet.Cell(1, 1)
	if cell.CType != XL_CELL_NUMBER {
		t.Errorf("cell.CType = %d, want %d", cell.CType, XL_CELL_NUMBER)
	}
	if cell.Value != 100.0 {
		t.Errorf("cell.Value = %v, want 100.0", cell.Value)
	}
	if cell.XFIndex != 21 {
		t.Errorf("cell.XFIndex = %d, want 21", cell.XFIndex)
	}
}

func TestCalculatedCell(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByName("PROFILELEVELS")
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	cell := sheet.Cell(1, 3)
	if cell.CType != XL_CELL_NUMBER {
		t.Errorf("cell.CType = %d, want %d", cell.CType, XL_CELL_NUMBER)
	}
	if !almostEqual(cell.Value.(float64), 265.131, 0.001) {
		t.Errorf("cell.Value = %v, want approximately 265.131", cell.Value)
	}
	if cell.XFIndex != 29 {
		t.Errorf("cell.XFIndex = %d, want 29", cell.XFIndex)
	}
}

func TestMergedCells(t *testing.T) {
	book, err := OpenWorkbook(fromSample("xf_class.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByName("table2")
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}

	// Check merged cells
	if len(sheet.MergedCells) == 0 {
		t.Error("Expected merged cells, but found none")
		return
	}

	rowLo, rowHi, colLo, colHi := sheet.MergedCells[0][0], sheet.MergedCells[0][1], sheet.MergedCells[0][2], sheet.MergedCells[0][3]
	if rowLo != 3 || rowHi != 7 || colLo != 2 || colHi != 5 {
		t.Errorf("Merged cell range = (%d,%d,%d,%d), want (3,7,2,5)", rowLo, rowHi, colLo, colHi)
	}

	// Check the value of the merged cell
	cell := sheet.Cell(rowLo, colLo)
	if cell.Value != "MERGED" {
		t.Errorf("Merged cell value = %v, want 'MERGED'", cell.Value)
	}
}

// Helper function to compare floats with tolerance
func almostEqual(a, b, tolerance float64) bool {
	return math.Abs(a-b) <= tolerance
}
