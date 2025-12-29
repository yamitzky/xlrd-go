package xlrd

import (
	"testing"
)

func TestWorkbookOpenWorkbook(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	if book == nil {
		t.Fatal("OpenWorkbook returned nil")
	}
}

func TestWorkbookNSheets(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	if book.NSheets != 5 {
		t.Errorf("book.NSheets = %d, want 5", book.NSheets)
	}
}

func TestWorkbookSheetByName(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	expectedNames := []string{"PROFILEDEF", "AXISDEF", "TRAVERSALCHAINAGE",
		"AXISDATUMLEVELS", "PROFILELEVELS"}
	for _, name := range expectedNames {
		sheet, err := book.SheetByName(name)
		if err != nil {
			t.Errorf("book.SheetByName(%q) error = %v", name, err)
			continue
		}
		if sheet == nil {
			t.Errorf("book.SheetByName(%q) returned nil", name)
			continue
		}
		if sheet.Name != name {
			t.Errorf("sheet.Name = %s, want %s", sheet.Name, name)
		}
	}
}

func TestWorkbookSheetByIndex(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	expectedNames := []string{"PROFILEDEF", "AXISDEF", "TRAVERSALCHAINAGE",
		"AXISDATUMLEVELS", "PROFILELEVELS"}
	for index := 0; index < 5; index++ {
		sheet, err := book.SheetByIndex(index)
		if err != nil {
			t.Errorf("book.SheetByIndex(%d) error = %v", index, err)
			continue
		}
		if sheet == nil {
			t.Errorf("book.SheetByIndex(%d) returned nil", index)
			continue
		}
		if index < len(expectedNames) && sheet.Name != expectedNames[index] {
			t.Errorf("sheet.Name = %s, want %s", sheet.Name, expectedNames[index])
		}
	}
}

func TestWorkbookSheets(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	expectedNames := []string{"PROFILEDEF", "AXISDEF", "TRAVERSALCHAINAGE",
		"AXISDATUMLEVELS", "PROFILELEVELS"}
	sheets := book.Sheets()
	if len(sheets) != len(expectedNames) {
		t.Errorf("book.Sheets() length = %d, want %d", len(sheets), len(expectedNames))
		return
	}
	for index, sheet := range sheets {
		if sheet == nil {
			t.Errorf("book.Sheets()[%d] is nil", index)
			continue
		}
		if index < len(expectedNames) && sheet.Name != expectedNames[index] {
			t.Errorf("sheet.Name = %s, want %s", sheet.Name, expectedNames[index])
		}
	}
}

func TestWorkbookSheetNames(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	expectedNames := []string{"PROFILEDEF", "AXISDEF", "TRAVERSALCHAINAGE",
		"AXISDATUMLEVELS", "PROFILELEVELS"}
	sheetNames := book.SheetNames()
	if len(sheetNames) != len(expectedNames) {
		t.Errorf("book.SheetNames() length = %d, want %d", len(sheetNames), len(expectedNames))
		return
	}
	for i, name := range expectedNames {
		if i >= len(sheetNames) {
			t.Errorf("Missing sheet name at index %d", i)
			continue
		}
		if sheetNames[i] != name {
			t.Errorf("book.SheetNames()[%d] = %s, want %s", i, sheetNames[i], name)
		}
	}
}

func TestWorkbookGetByIndex(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.Get(0)
	if err != nil {
		t.Fatalf("book.Get(0) error = %v", err)
	}
	if sheet == nil {
		t.Fatal("book.Get(0) returned nil")
	}
	if sheet.Name != "PROFILEDEF" {
		t.Errorf("sheet.Name = %s, want PROFILEDEF", sheet.Name)
	}
}

func TestWorkbookGetByName(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.Get("PROFILEDEF")
	if err != nil {
		t.Fatalf("book.Get(\"PROFILEDEF\") error = %v", err)
	}
	if sheet == nil {
		t.Fatal("book.Get(\"PROFILEDEF\") returned nil")
	}
	if sheet.Name != "PROFILEDEF" {
		t.Errorf("sheet.Name = %s, want PROFILEDEF", sheet.Name)
	}
}

func TestWorkbookCellAccess(t *testing.T) {
	book, err := OpenWorkbook(fromSample("profiles.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.Get(0)
	if err != nil {
		t.Fatalf("Failed to get sheet 0: %v", err)
	}

	// Test that cells are not empty
	cell00 := sheet.Cell(0, 0)
	if cell00 == nil {
		t.Error("sheet.Cell(0, 0) returned nil")
	}
	if cell00.CType == XL_CELL_EMPTY {
		t.Error("sheet.Cell(0, 0) is empty, expected non-empty")
	}

	// Check actual cell values
	if cell00.CType != XL_CELL_TEXT {
		t.Errorf("sheet.Cell(0, 0) type = %d, want %d", cell00.CType, XL_CELL_TEXT)
	}
	if cell00.Value != "PROFIL" {
		t.Errorf("sheet.Cell(0, 0) value = %v, want %q", cell00.Value, "PROFIL")
	}

	// Test corner cell (this might be empty, so just check it doesn't panic)
	cellCorner := sheet.Cell(14, 12) // NROWS-1, NCOLS-1 from Python test
	if cellCorner == nil {
		t.Error("sheet.Cell(14, 12) returned nil")
	}
}
