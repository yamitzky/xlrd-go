package xlrd

import (
	"fmt"
	"io"
)

// XLRDError represents an error that occurred while reading an Excel file.
type XLRDError struct {
	Message string
}

func (e *XLRDError) Error() string {
	return e.Message
}

// NewXLRDError creates a new XLRDError with the given message.
func NewXLRDError(format string, args ...interface{}) *XLRDError {
	return &XLRDError{Message: fmt.Sprintf(format, args...)}
}

// Cell types
const (
	XL_CELL_EMPTY   = 0
	XL_CELL_TEXT    = 1
	XL_CELL_NUMBER  = 2
	XL_CELL_DATE    = 3
	XL_CELL_BOOLEAN = 4
	XL_CELL_ERROR   = 5
	XL_CELL_BLANK   = 6 // for use in debugging, gathering stats, etc
)

// Format types
const (
	FUN = 0 // unknown
	FDT = 1 // date
	FNU = 2 // number
	FGE = 3 // general
	FTX = 4 // text
)

const (
	DATEFORMAT   = FDT
	NUMBERFORMAT = FNU
)

// BIFF version constants
const (
	BIFF_FIRST_UNICODE = 80
)

var biffTextFromNum = map[int]string{
	0:  "(not BIFF)",
	20: "2.0",
	21: "2.1",
	30: "3",
	40: "4S",
	45: "4W",
	50: "5",
	70: "7",
	80: "8",
	85: "8X",
}

// BiffTextFromNum returns a text representation of a BIFF version number.
func BiffTextFromNum(num int) string {
	if text, ok := biffTextFromNum[num]; ok {
		return text
	}
	return fmt.Sprintf("Unknown(%d)", num)
}

// ErrorTextFromCode returns a text representation of an Excel error code.
var ErrorTextFromCode = map[byte]string{
	0x00: "#NULL!", // Intersection of two cell ranges is empty
	0x07: "#DIV/0!", // Division by zero
	0x0F: "#VALUE!", // Wrong type of operand
	0x17: "#REF!",   // Illegal or deleted cell reference
	0x1D: "#NAME?",  // Wrong function or range name
	0x24: "#NUM!",   // Value range overflow
	0x2A: "#N/A",    // Argument or function not available
}

// BIFF record type constants
const (
	XL_WORKBOOK_GLOBALS     = 0x5
	XL_WORKBOOK_GLOBALS_4W  = 0x100
	XL_WORKSHEET            = 0x10
	XL_BOUNDSHEET_WORKSHEET = 0x00
	XL_BOUNDSHEET_CHART     = 0x02
	XL_BOUNDSHEET_VB_MODULE = 0x06
	XL_ARRAY                = 0x0221
	XL_ARRAY2               = 0x0021
	XL_BLANK                = 0x0201
	XL_BLANK_B2             = 0x01
	XL_BOF                  = 0x809
	XL_BOOLERR              = 0x205
	XL_BOOLERR_B2           = 0x5
	XL_BOUNDSHEET           = 0x85
	XL_BUILTINFMTCOUNT       = 0x56
	XL_CF                   = 0x01B1
	XL_CODEPAGE             = 0x42
	XL_COLINFO              = 0x7D
	XL_COLUMNDEFAULT        = 0x20 // BIFF2 only
	XL_COLWIDTH             = 0x24 // BIFF2 only
	XL_CONDFMT              = 0x01B0
	XL_CONTINUE             = 0x3c
	XL_COUNTRY              = 0x8C
	XL_DATEMODE             = 0x22
	XL_DEFAULTROWHEIGHT     = 0x0225
	XL_DEFCOLWIDTH          = 0x55
	XL_DIMENSION            = 0x200
	XL_DIMENSION2           = 0x0
	XL_EFONT                = 0x45
	XL_EOF                  = 0x0a
	XL_EXTERNNAME           = 0x23
	XL_EXTERNSHEET          = 0x17
	XL_EXTSST               = 0xff
	XL_FEAT11               = 0x872
	XL_FILEPASS             = 0x2f
	XL_FONT                 = 0x31
	XL_FONT_B3B4            = 0x231
	XL_FORMAT                = 0x41e
	XL_FORMAT2               = 0x1E // BIFF2, BIFF3
	XL_FORMULA               = 0x6
	XL_FORMULA3              = 0x206
	XL_FORMULA4              = 0x406
	XL_GCW                   = 0xab
	XL_HLINK                 = 0x01B8
	XL_QUICKTIP              = 0x0800
	XL_HORIZONTALPAGEBREAKS  = 0x1b
	XL_INDEX                 = 0x20b
	XL_INTEGER               = 0x2 // BIFF2 only
	XL_IXFE                  = 0x44 // BIFF2 only
	XL_LABEL                 = 0x204
	XL_LABEL_B2              = 0x04
	XL_LABELRANGES           = 0x15f
	XL_LABELSST              = 0xfd
	XL_LEFTMARGIN            = 0x26
	XL_TOPMARGIN             = 0x28
	XL_RIGHTMARGIN           = 0x27
	XL_BOTTOMMARGIN          = 0x29
	XL_HEADER                = 0x14
	XL_FOOTER                = 0x15
	XL_HCENTER               = 0x83
	XL_VCENTER               = 0x84
	XL_MERGEDCELLS           = 0xE5
	XL_MSO_DRAWING           = 0x00EC
	XL_MSO_DRAWING_GROUP     = 0x00EB
	XL_MSO_DRAWING_SELECTION = 0x00ED
	XL_MULRK                 = 0xbd
	XL_MULBLANK              = 0xbe
	XL_NAME                  = 0x18
	XL_NOTE                  = 0x1c
	XL_NUMBER                = 0x203
	XL_NUMBER_B2             = 0x3
	XL_OBJ                   = 0x5D
	XL_PAGESETUP             = 0xA1
	XL_PALETTE               = 0x92
	XL_PANE                  = 0x41
	XL_PRINTGRIDLINES        = 0x2B
	XL_PRINTHEADERS          = 0x2A
	XL_RK                    = 0x27e
	XL_ROW                   = 0x208
	XL_ROW_B2                = 0x08
	XL_RSTRING               = 0xd6
	XL_SCL                   = 0x00A0
	XL_SHEETHDR              = 0x8F // BIFF4W only
	XL_SHEETPR               = 0x81
	XL_SHEETSOFFSET          = 0x8E // BIFF4W only
	XL_SHRFMLA               = 0x04bc
	XL_SST                   = 0xfc
	XL_STANDARDWIDTH         = 0x99
	XL_STRING                = 0x207
	XL_STRING_B2             = 0x7
	XL_STYLE                 = 0x293
	XL_SUPBOOK               = 0x1AE // aka EXTERNALBOOK in OOo docs
	XL_TABLEOP               = 0x236
	XL_TABLEOP2              = 0x37
	XL_TABLEOP_B2            = 0x36
	XL_TXO                   = 0x1b6
	XL_UNCALCED              = 0x5e
	XL_UNKNOWN               = 0xffff
	XL_VERTICALPAGEBREAKS    = 0x1a
	XL_WINDOW2               = 0x023E
	XL_WINDOW2_B2            = 0x003E
	XL_WRITEACCESS           = 0x5C
	XL_WSBOOL                = XL_SHEETPR
	XL_XF                    = 0xe0
	XL_XF2                   = 0x0043 // BIFF2 version of XF record
	XL_XF3                   = 0x0243 // BIFF3 version of XF record
	XL_XF4                   = 0x0443 // BIFF4 version of XF record
)

var boflen = map[int]int{
	0x0809: 8,
	0x0409: 6,
	0x0209: 6,
	0x0009: 4,
}

var bofcodes = []int{0x0809, 0x0409, 0x0209, 0x0009}

var XL_FORMULA_OPCODES = []int{0x0006, 0x0406, 0x0206}

var cellOpcodeSet = map[int]bool{
	XL_BOOLERR:  true,
	XL_FORMULA:  true,
	XL_FORMULA3: true,
	XL_FORMULA4: true,
	XL_LABEL:    true,
	XL_LABELSST: true,
	XL_MULRK:    true,
	XL_NUMBER:   true,
	XL_RK:       true,
	XL_RSTRING:  true,
}

// IsCellOpcode checks if the given code is a cell opcode.
func IsCellOpcode(c int) bool {
	return cellOpcodeSet[c]
}

// SupportedVersions lists all supported BIFF versions.
var SupportedVersions = []int{80, 70, 50, 45, 40, 30, 21, 20}

// EncodingFromCodepage maps codepage numbers to encoding names.
var EncodingFromCodepage = map[int]string{
	1200:  "utf_16_le",
	10000: "mac_roman",
	10006: "mac_greek",     // guess
	10007: "mac_cyrillic",  // guess
	10029: "mac_latin2",    // guess
	10079: "mac_iceland",   // guess
	10081: "mac_turkish",   // guess
	32768: "mac_roman",
	32769: "cp1252",
}

// BaseObject is a base type for most objects in the package.
// It provides a common dump method for debugging.
type BaseObject struct {
	// This is a placeholder for now
}

// Dump writes debugging information about the object to the given writer.
func (b *BaseObject) Dump(w io.Writer, header, footer string, indent int) {
	// Empty implementation for now
	if header != "" {
		fmt.Fprintf(w, "%s\n", header)
	}
	if footer != "" {
		fmt.Fprintf(w, "%s\n", footer)
	}
}
