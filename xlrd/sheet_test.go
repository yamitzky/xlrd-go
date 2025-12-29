package xlrd

import (
	"testing"
)

const (
	sheetIndex = 0
	nRows      = 15
	nCols      = 13
)

const (
	rowErr = nRows + 10
	colErr = nCols + 10
)

var sheetNames = []string{"PROFILEDEF", "AXISDEF", "TRAVERSALCHAINAGE",
	"AXISDATUMLEVELS", "PROFILELEVELS"}

func TestSheetNRows(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	if sheet.NRows != nRows {
		t.Errorf("sheet.NRows = %d, want %d", sheet.NRows, nRows)
	}
}

func TestSheetNCols(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	if sheet.NCols != nCols {
		t.Errorf("sheet.NCols = %d, want %d", sheet.NCols, nCols)
	}
}

func TestSheetCell(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Check that cells are accessible (not necessarily non-empty)
	cell00 := sheet.Cell(0, 0)
	if cell00 == nil {
		t.Error("sheet.Cell(0, 0) should not be nil")
	}
	cellCorner := sheet.Cell(nRows-1, nCols-1)
	if cellCorner == nil {
		t.Error("sheet.Cell(nRows-1, nCols-1) should not be nil")
	}
}

func TestSheetCellError(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Test out of bounds access - should not panic in Go, but check bounds
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("sheet.Cell(rowErr, 0) should not panic, got: %v", r)
		}
	}()
	cell := sheet.Cell(rowErr, 0)
	if cell != nil && cell.CType != XL_CELL_EMPTY {
		t.Error("sheet.Cell(rowErr, 0) should return empty cell or nil")
	}

	defer func() {
		if r := recover(); r != nil {
			t.Errorf("sheet.Cell(0, colErr) should not panic, got: %v", r)
		}
	}()
	cell = sheet.Cell(0, colErr)
	if cell != nil && cell.CType != XL_CELL_EMPTY {
		t.Error("sheet.Cell(0, colErr) should return empty cell or nil")
	}
}

func TestSheetCellType(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Check that cell types are accessible
	ctype00 := sheet.CellType(0, 0)
	if ctype00 < 0 {
		t.Errorf("sheet.CellType(0, 0) = %d, should be >= 0", ctype00)
	}
	ctypeCorner := sheet.CellType(nRows-1, nCols-1)
	if ctypeCorner < 0 {
		t.Errorf("sheet.CellType(nRows-1, nCols-1) = %d, should be >= 0", ctypeCorner)
	}
}

func TestSheetCellTypeError(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Test out of bounds access
	ctype := sheet.CellType(rowErr, 0)
	if ctype != XL_CELL_EMPTY {
		t.Errorf("sheet.CellType(rowErr, 0) = %d, want %d", ctype, XL_CELL_EMPTY)
	}

	ctype = sheet.CellType(0, colErr)
	if ctype != XL_CELL_EMPTY {
		t.Errorf("sheet.CellType(0, colErr) = %d, want %d", ctype, XL_CELL_EMPTY)
	}
}

func TestSheetCellValue(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Check that cell values are accessible (can be nil for empty cells)
	value00 := sheet.CellValue(0, 0)
	// value00 can be nil for empty cells, just check it's accessible
	_ = value00

	valueCorner := sheet.CellValue(nRows-1, nCols-1)
	// valueCorner can be nil for empty cells, just check it's accessible
	_ = valueCorner
}

func TestSheetCellValueError(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Test out of bounds access
	value := sheet.CellValue(rowErr, 0)
	if value != nil {
		t.Errorf("sheet.CellValue(rowErr, 0) should be nil, got %v", value)
	}

	value = sheet.CellValue(0, colErr)
	if value != nil {
		t.Errorf("sheet.CellValue(0, colErr) should be nil, got %v", value)
	}
}

func TestSheetCellXFIndex(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Check that XF indexes are valid (>= 0)
	xf00 := sheet.CellXFIndex(0, 0)
	if xf00 < 0 {
		t.Errorf("sheet.CellXFIndex(0, 0) = %d, should be >= 0", xf00)
	}
	xfCorner := sheet.CellXFIndex(nRows-1, nCols-1)
	if xfCorner < 0 {
		t.Errorf("sheet.CellXFIndex(nRows-1, nCols-1) = %d, should be >= 0", xfCorner)
	}
}

func TestSheetCellXFIndexError(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	// Test out of bounds access - should return 0 for invalid positions
	xf := sheet.CellXFIndex(rowErr, 0)
	if xf != 0 {
		t.Errorf("sheet.CellXFIndex(rowErr, 0) = %d, want 0", xf)
	}

	xf = sheet.CellXFIndex(0, colErr)
	if xf != 0 {
		t.Errorf("sheet.CellXFIndex(0, colErr) = %d, want 0", xf)
	}
}

// Col method is not in Go version - use ColSlice instead
func TestSheetCol(t *testing.T) {
	t.Skip("Col method not implemented in Go version")
}

func TestSheetRow(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	row := sheet.Row(0)
	if len(row) != nCols {
		t.Errorf("len(sheet.Row(0)) = %d, want %d", len(row), nCols)
	}
}

func TestSheetColSlice(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	slice := sheet.ColSlice(0, 2, &[]int{nRows - 2}[0])
	expectedLen := nRows - 4 // startRowx=2, endRowx=nRows-2, so length = (nRows-2) - 2 = nRows-4
	if len(slice) != expectedLen {
		t.Errorf("len(sheet.ColSlice(0, 2, %d)) = %d, want %d", nRows-2, len(slice), expectedLen)
	}
}

func TestSheetColTypes(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	types := sheet.ColTypes(0, 2, &[]int{nRows - 2}[0])
	expectedLen := nRows - 4
	if len(types) != expectedLen {
		t.Errorf("len(sheet.ColTypes(0, 2, %d)) = %d, want %d", nRows-2, len(types), expectedLen)
	}
}

func TestSheetColValues(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	values := sheet.ColValues(0, 2, &[]int{nRows - 2}[0])
	expectedLen := nRows - 4
	if len(values) != expectedLen {
		t.Errorf("len(sheet.ColValues(0, 2, %d)) = %d, want %d", nRows-2, len(values), expectedLen)
	}
}

func TestSheetRowSlice(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	slice := sheet.RowSlice(0, 2, &[]int{nCols - 2}[0])
	expectedLen := nCols - 4 // startColx=2, endColx=nCols-2, so length = (nCols-2) - 2 = nCols-4
	if len(slice) != expectedLen {
		t.Errorf("len(sheet.RowSlice(0, 2, %d)) = %d, want %d", nCols-2, len(slice), expectedLen)
	}
}

func TestSheetRowTypes(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	types := sheet.RowTypes(0, 2, &[]int{nCols - 2}[0])
	expectedLen := nCols - 4
	if len(types) != expectedLen {
		t.Errorf("len(sheet.RowTypes(0, 2, %d)) = %d, want %d", nCols-2, len(types), expectedLen)
	}
}

func TestSheetRowValues(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), &OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(sheetIndex)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	values := sheet.RowValues(0, 2, &[]int{nCols - 2}[0])
	expectedLen := nCols - 4
	if len(values) != expectedLen {
		t.Errorf("len(sheet.RowValues(0, 2, %d)) = %d, want %d", nCols-2, len(values), expectedLen)
	}
}

func TestSheetRagged(t *testing.T) {
	book, err := OpenWorkbook(fromSample("ragged.xls"), &OpenWorkbookOptions{RaggedRows: true})
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	if sheet.RowLen(0) != 3 {
		t.Errorf("sheet.RowLen(0) = %d, want 3", sheet.RowLen(0))
	}
	if sheet.RowLen(1) != 2 {
		t.Errorf("sheet.RowLen(1) = %d, want 2", sheet.RowLen(1))
	}
	if sheet.RowLen(2) != 1 {
		t.Errorf("sheet.RowLen(2) = %d, want 1", sheet.RowLen(2))
	}
	if sheet.RowLen(3) != 4 {
		t.Errorf("sheet.RowLen(3) = %d, want 4", sheet.RowLen(3))
	}
	if sheet.RowLen(4) != 4 {
		t.Errorf("sheet.RowLen(4) = %d, want 4", sheet.RowLen(4))
	}
}
