package xlrd

import (
	"fmt"
	"strings"
	"testing"
	"unicode"
)

// ascii returns a string representation similar to Python's ascii() function.
// It escapes non-ASCII characters and handles different types appropriately.
func ascii(obj interface{}) string {
	if obj == nil {
		return "None"
	}

	switch v := obj.(type) {
	case string:
		return escapeNonASCII(v)
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		return fmt.Sprintf("%d", v)
	case float32, float64:
		// Format float similar to Python's str() - ensure .0 for whole numbers
		var f float64
		if val, ok := v.(float64); ok {
			f = val
		} else if val, ok := v.(float32); ok {
			f = float64(val)
		} else {
			return fmt.Sprintf("%v", v)
		}

		// Check if it's a whole number
		if f == float64(int64(f)) {
			return fmt.Sprintf("%.1f", f)
		}
		return fmt.Sprintf("%g", f)
	case bool:
		if v {
			return "True"
		}
		return "False"
	default:
		// For other types, convert to string first, then escape
		str := fmt.Sprintf("%v", v)
		return escapeNonASCII(str)
	}
}

// escapeNonASCII escapes non-ASCII characters in a string, similar to Python's ascii().
func escapeNonASCII(s string) string {
	var result strings.Builder
	result.WriteByte('\'')
	for _, r := range s {
		if r < 128 && unicode.IsPrint(r) && r != '\'' && r != '\\' {
			result.WriteRune(r)
		} else if r == '\'' {
			result.WriteString("\\'")
		} else if r == '\\' {
			result.WriteString("\\\\")
		} else if r < 128 {
			// Non-printable ASCII
			switch r {
			case '\n':
				result.WriteString("\\n")
			case '\r':
				result.WriteString("\\r")
			case '\t':
				result.WriteString("\\t")
			default:
				result.WriteString(fmt.Sprintf("\\x%02x", r))
			}
		} else {
			// Unicode character
			result.WriteString(fmt.Sprintf("\\u%04x", r))
		}
	}
	result.WriteByte('\'')
	return result.String()
}

func TestFormulas(t *testing.T) {
	book, err := OpenWorkbook(fromSample("formula_test_sjmachin.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}

	// Test cell B2
	values := sheet.ColValues(1, 0, nil)
	if len(values) < 2 {
		t.Fatal("Not enough values in column")
	}
	// Expected: '\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'
	expectedB2 := "'\\u041c\\u041e\\u0421\\u041a\\u0412\\u0410 \\u041c\\u043e\\u0441\\u043a\\u0432\\u0430'"
	if ascii(values[1]) != expectedB2 {
		t.Errorf("Cell B2: got %q, want %q", ascii(values[1]), expectedB2)
	}

	// Test cell B3
	if len(values) < 3 {
		t.Fatal("Not enough values in column")
	}
	// Expected: 0.14285714285714285
	expectedB3 := "0.14285714285714285"
	if ascii(values[2]) != expectedB3 {
		t.Errorf("Cell B3: got %q, want %q", ascii(values[2]), expectedB3)
	}

	// Test cell B4
	if len(values) < 4 {
		t.Fatal("Not enough values in column")
	}
	// Expected: "'ABCDEF'"
	expectedB4 := "'ABCDEF'"
	if ascii(values[3]) != expectedB4 {
		t.Errorf("Cell B4: got %q, want %q", ascii(values[3]), expectedB4)
	}

	// Test cell B5
	if len(values) < 5 {
		t.Fatal("Not enough values in column")
	}
	// Expected: "''"
	expectedB5 := "''"
	if ascii(values[4]) != expectedB5 {
		t.Errorf("Cell B5: got %q, want %q", ascii(values[4]), expectedB5)
	}

	// Test cell B6
	if len(values) < 6 {
		t.Fatal("Not enough values in column")
	}
	// Expected: '1'
	expectedB6 := "1"
	if ascii(values[5]) != expectedB6 {
		t.Errorf("Cell B6: got %q, want %q", ascii(values[5]), expectedB6)
	}

	// Test cell B7
	if len(values) < 7 {
		t.Fatal("Not enough values in column")
	}
	// Expected: '7'
	expectedB7 := "7"
	if ascii(values[6]) != expectedB7 {
		t.Errorf("Cell B7: got %q, want %q", ascii(values[6]), expectedB7)
	}

	// Test cell B8
	if len(values) < 8 {
		t.Fatal("Not enough values in column")
	}
	// Expected: '\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'
	expectedB8 := "'\\u041c\\u041e\\u0421\\u041a\\u0412\\u0410 \\u041c\\u043e\\u0441\\u043a\\u0432\\u0430'"
	if ascii(values[7]) != expectedB8 {
		t.Errorf("Cell B8: got %q, want %q", ascii(values[7]), expectedB8)
	}
}

func TestNameFormulas(t *testing.T) {
	book, err := OpenWorkbook(fromSample("formula_test_names.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}

	// Test various formula results
	values := sheet.ColValues(1, 0, nil)

	// Test unaryop: -7.0
	expectedUnaryop := "-7.0"
	if ascii(values[1]) != expectedUnaryop {
		t.Errorf("unaryop: got %q, want %q", ascii(values[1]), expectedUnaryop)
	}

	// Test attrsum: 4.0
	expectedAttrsum := "4.0"
	if ascii(values[2]) != expectedAttrsum {
		t.Errorf("attrsum: got %q, want %q", ascii(values[2]), expectedAttrsum)
	}

	// Test func: 6.0
	expectedFunc := "6.0"
	if ascii(values[3]) != expectedFunc {
		t.Errorf("func: got %q, want %q", ascii(values[3]), expectedFunc)
	}

	// Test func_var_args: 3.0
	expectedFuncVarArgs := "3.0"
	if ascii(values[4]) != expectedFuncVarArgs {
		t.Errorf("func_var_args: got %q, want %q", ascii(values[4]), expectedFuncVarArgs)
	}

	// Test if: 'b'
	expectedIf := "'b'"
	if ascii(values[5]) != expectedIf {
		t.Errorf("if: got %q, want %q", ascii(values[5]), expectedIf)
	}

	// Test choose: 'C'
	expectedChoose := "'C'"
	if ascii(values[6]) != expectedChoose {
		t.Errorf("choose: got %q, want %q", ascii(values[6]), expectedChoose)
	}
}

func TestEvaluateNameFormulaWithInvalidOperand(t *testing.T) {
	book, err := OpenWorkbook(fromSample("invalid_formula.xls"), nil)
	if err != nil {
		t.Fatalf("Failed to open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("Failed to get sheet: %v", err)
	}
	cell := sheet.Cell(0, 0)
	if cell.CType != XL_CELL_ERROR {
		t.Errorf("cell.CType = %d, want %d", cell.CType, XL_CELL_ERROR)
	}
	if cellValue, ok := cell.Value.(int); ok {
		if _, exists := ErrorTextFromCode[byte(cellValue)]; !exists {
			t.Errorf("cell.Value %d is not a valid error code", cellValue)
		}
	} else {
		t.Errorf("cell.Value is not an int, got %T", cell.Value)
	}
}
