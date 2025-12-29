package xlrd

import (
	"testing"
)

func TestIgnoreWorkbookCorruption(t *testing.T) {
	// TODO: Implement workbook opening
	// Test that corrupted workbook raises error by default
	// _, err := OpenWorkbook(fromSample("corrupted_error.xls"), nil)
	// if err == nil {
	// 	t.Error("OpenWorkbook(corrupted_error.xls) should have returned an error")
	// 	return
	// }
	// if !strings.Contains(err.Error(), "Workbook corruption") {
	// 	t.Logf("OpenWorkbook(corrupted_error.xls) error = %v (expected error)", err)
	// }
	//
	// // Test that corrupted workbook can be opened with ignore_workbook_corruption=true
	// options := &OpenWorkbookOptions{
	// 	IgnoreWorkbookCorruption: true,
	// }
	// _, err = OpenWorkbook(fromSample("corrupted_error.xls"), options)
	// if err != nil {
	// 	t.Errorf("OpenWorkbook(corrupted_error.xls) with ignore_workbook_corruption=true should succeed, got error: %v", err)
	// }
	t.Log("TestIgnoreWorkbookCorruption: TODO - implement workbook opening with ignore_workbook_corruption option")
}
