// Package xlrd provides functionality for reading Excel files
package xlrd

import (
	"encoding/binary"
	"fmt"
	"os"
	"reflect"
	"strings"
)


// assert checks a condition and panics if false
func assert(condition bool) {
	if !condition {
		panic("assertion failed")
	}
}

// unpack unpacks binary data according to format string
func unpack(format string, data []byte) (interface{}, error) {
	if strings.HasPrefix(format, "<") {
		// Little-endian
		switch format[1:] {
		case "B":
			if len(data) < 1 {
				return nil, fmt.Errorf("not enough data")
			}
			return uint8(data[0]), nil
		case "H":
			if len(data) < 2 {
				return nil, fmt.Errorf("not enough data")
			}
			return binary.LittleEndian.Uint16(data), nil
		case "h":
			if len(data) < 2 {
				return nil, fmt.Errorf("not enough data")
			}
			return int16(binary.LittleEndian.Uint16(data)), nil
		case "BH":
			if len(data) < 3 {
				return nil, fmt.Errorf("not enough data")
			}
			return []interface{}{uint8(data[0]), binary.LittleEndian.Uint16(data[1:])}, nil
		case "HH":
			if len(data) < 4 {
				return nil, fmt.Errorf("not enough data")
			}
			return []interface{}{binary.LittleEndian.Uint16(data[:2]), binary.LittleEndian.Uint16(data[2:])}, nil
		case "hxxxxxxxxhh":
			if len(data) < 15 {
				return nil, fmt.Errorf("not enough data")
			}
			return []interface{}{
				int16(binary.LittleEndian.Uint16(data[:2])),
				int16(binary.LittleEndian.Uint16(data[10:12])),
				int16(binary.LittleEndian.Uint16(data[12:14])),
			}, nil
		case "d":
			if len(data) < 8 {
				return nil, fmt.Errorf("not enough data")
			}
			return binary.LittleEndian.Uint64(data), nil
		case "x2H":
			if len(data) < 4 {
				return nil, fmt.Errorf("not enough data")
			}
			return []interface{}{binary.LittleEndian.Uint16(data[2:4])}, nil
		case "xHB":
			if len(data) < 4 {
				return nil, fmt.Errorf("not enough data")
			}
			return []interface{}{binary.LittleEndian.Uint16(data[2:4]), uint8(data[3])}, nil
		default:
			return nil, fmt.Errorf("unsupported format: %s", format)
		}
	}
	return nil, fmt.Errorf("unsupported endianness: %s", format)
}

// copyOperand creates a deep copy of an Operand
func copyOperand(op *Operand) *Operand {
	if op == nil {
		return nil
	}
	newOp := &Operand{
		kind:  op.kind,
		value: op.value, // shallow copy for now
		text:  op.text,
		rpn:   make([]interface{}, len(op.rpn)),
		_rank: op._rank,
	}
	copy(newOp.rpn, op.rpn)
	return newOp
}

// evaluateNameFormula evaluates a named formula recursively
func evaluateNameFormula(bk *Book, tgtobj *Name, tgtnamex int, blah int, level int) {
	// TODO: Implement name formula evaluation
	// For now, just mark as evaluated to avoid infinite recursion
	tgtobj.Evaluated = true
}

// errorTextFromCode returns error text from error code
var errorTextFromCode = map[int]string{
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
}

// getExternsheetLocalRangeB57 gets external sheet local range for BIFF <= 7
func getExternsheetLocalRangeB57(bk *Book, rawExtshtx, rawShx1, rawShx2 int, blah int) (int, int) {
	if blah != 0 {
		fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range_b57(%d, %d, %d) -> ???\n",
			rawExtshtx, rawShx1, rawShx2)
	}
	if rawExtshtx >= 0 {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range_b57(raw_extshtx=%d) -> external\n", rawExtshtx)
		}
		return -4, -4 // external reference
	}
	refx := -rawExtshtx - 1
	if refx < 0 || refx >= len(bk.externsheetInfo) {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "!!! get_externsheet_local_range_b57: refx=%d, not in range(%d)\n",
				refx, len(bk.externsheetInfo))
		}
		return -101, -101
	}
	info := bk.externsheetInfo[refx]
	refRecordx := info[0]
	refFirstSheetx := info[1]
	refLastSheetx := info[2]
	if bk.supbookLocalsInx != nil && refRecordx != *bk.supbookLocalsInx {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range_b57(refx=%d) -> external\n", refx)
		}
		return -4, -4 // external reference
	}
	nsheets := len(bk.allSheetsMap)
	if !(0 <= refFirstSheetx && refFirstSheetx <= refLastSheetx && refLastSheetx < nsheets) {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range_b57(refx=%d) -> stuffed up\n", refx)
		}
		return -102, -102
	}
	xlrdSheetx1 := bk.allSheetsMap[refFirstSheetx]
	xlrdSheetx2 := bk.allSheetsMap[refLastSheetx]
	if !(0 <= xlrdSheetx1 && xlrdSheetx1 <= xlrdSheetx2) {
		return -3, -3 // internal reference, but to a macro sheet
	}
	return xlrdSheetx1, xlrdSheetx2
}

// rangename3d generates a range name for 3D reference
func rangename3d(bk *Book, ref3d *Ref3D) string {
	if ref3d.shtxlo == ref3d.shtxhi-1 {
		shname := bk.SheetNames()[ref3d.shtxlo]
		return fmt.Sprintf("%s!%s", shname, cellrange(ref3d.rlo, ref3d.clo, ref3d.rhi-1, ref3d.chi-1))
	}
	shname1 := bk.SheetNames()[ref3d.shtxlo]
	shname2 := bk.SheetNames()[ref3d.shtxhi-1]
	return fmt.Sprintf("%s:%s!%s", shname1, shname2, cellrange(ref3d.rlo, ref3d.clo, ref3d.rhi-1, ref3d.chi-1))
}

// rangename3drel generates a relative range name for 3D reference
func rangename3drel(bk *Book, ref3d *Ref3D, r1c1 int) string {
	if ref3d.shtxlo == ref3d.shtxhi-1 {
		shname := bk.SheetNames()[ref3d.shtxlo]
		return fmt.Sprintf("%s!%s", shname, cellrange_r1c1(ref3d.rlo, ref3d.clo, ref3d.rhi-1, ref3d.chi-1, ref3d.relflags))
	}
	shname1 := bk.SheetNames()[ref3d.shtxlo]
	shname2 := bk.SheetNames()[ref3d.shtxhi-1]
	return fmt.Sprintf("%s:%s!%s", shname1, shname2, cellrange_r1c1(ref3d.rlo, ref3d.clo, ref3d.rhi-1, ref3d.chi-1, ref3d.relflags))
}

// cellrange generates a cell range string
func cellrange(rlo, clo, rhi, chi int) string {
	return fmt.Sprintf("%s%d:%s%d", colname(clo), rlo+1, colname(chi), rhi+1)
}

// cellrange_r1c1 generates a cell range string in R1C1 notation
func cellrange_r1c1(rlo, clo, rhi, chi int, relflags []int) string {
	rlo_rel, rhi_rel := relflags[2], relflags[3]
	clo_rel, chi_rel := relflags[4], relflags[5]
	return fmt.Sprintf("%s%s:%s%s",
		cellnameabs(rlo, rlo_rel, 1),
		colnameabs(clo, clo_rel, 1),
		cellnameabs(rhi, rhi_rel, 1),
		colnameabs(chi, chi_rel, 1))
}

// cellnameabs generates absolute cell name
func cellnameabs(row, row_rel int, r1c1 int) string {
	if r1c1 != 0 {
		if row_rel != 0 {
			return fmt.Sprintf("R[%d]", row)
		}
		return fmt.Sprintf("R%d", row+1)
	}
	return fmt.Sprintf("%d", row+1)
}

// colnameabs generates absolute column name
func colnameabs(col, col_rel int, r1c1 int) string {
	if r1c1 != 0 {
		if col_rel != 0 {
			return fmt.Sprintf("C[%d]", col)
		}
		return fmt.Sprintf("C%d", col+1)
	}
	return colname(col)
}

// colname generates column name from column index
func colname(col int) string {
	letters := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	result := ""
	for col >= 0 {
		result = string(letters[col%26]) + result
		col = col/26 - 1
	}
	return result
}

// hexCharDump dumps hex and character representation of data
func hexCharDump(data []byte, ofs, dlen int, fout *os.File) {
	endpos := min(ofs+dlen, len(data))
	pos := ofs
	for pos < endpos {
		endsub := min(pos+16, endpos)
		substr := data[pos:endsub]
		lensub := endsub - pos
		if lensub <= 0 {
			break
		}

		hexd := ""
		chard := ""
		for _, c := range substr {
			hexd += fmt.Sprintf("%02x ", c)
			if c == 0 {
				chard += "~"
			} else if c >= 32 && c <= 126 {
				chard += string(c)
			} else {
				chard += "?"
			}
		}

		// Pad hexd to 48 characters
		for len(hexd) < 48 {
			hexd += " "
		}

		fmt.Fprintf(fout, "%5d:     %-48s %s\n", pos-ofs, hexd, chard)
		pos = endsub
	}
}

// unpackStringUpdatePos unpacks string and updates position
func unpackStringUpdatePos(data []byte, pos int, encoding string, lenlen int) (string, int) {
	var strlen int
	if lenlen == 1 {
		strlen = int(data[pos])
		pos++
	} else {
		strlen = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2
	}
	strbytes := data[pos : pos+strlen]
	pos += strlen
	// For simplicity, assume UTF-8 encoding
	return string(strbytes), pos
}

// unpackUnicodeUpdatePos unpacks unicode string and updates position
func unpackUnicodeUpdatePos(data []byte, pos int, lenlen int) (string, int) {
	var strlen int
	if lenlen == 1 {
		strlen = int(data[pos])
		pos++
	} else {
		strlen = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2
	}
	// Unicode strings in Excel are UTF-16LE
	strbytes := data[pos : pos+strlen*2]
	pos += strlen * 2

	// Convert UTF-16LE to string
	utf16 := make([]uint16, strlen)
	for i := 0; i < strlen; i++ {
		utf16[i] = binary.LittleEndian.Uint16(strbytes[i*2 : i*2+2])
	}

	runes := make([]rune, 0, strlen)
	for _, r := range utf16 {
		runes = append(runes, rune(r))
	}
	return string(runes), pos
}

// min returns the minimum of two integers
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// Formula type constants
const (
	FmlaTypeCell    = 1
	FmlaTypeShared  = 2
	FmlaTypeArray   = 4
	FmlaTypeCondFmt = 8
	FmlaTypeDataVal = 16
	FmlaTypeName    = 32
	AllFmlaTypes    = 63
)

// List separator
const listsep = ","

// Stack levels
const (
	LeafRank        = 90
	FuncRank        = 90
	StackAlarmLevel = 5
	StackPanicLevel = 10
)

// Error opcodes
var errorOpcodes = map[int]bool{
	0x07: true,
	0x08: true,
	0x0A: true,
	0x0B: true,
	0x1C: true,
	0x1D: true,
	0x2F: true,
}

// tAttrNames maps subopcodes for tAttr
var tAttrNames = map[int]string{
	0x00: "Skip??", // seen in SAMPLES.XLS which shipped with Excel 5.0
	0x01: "Volatile",
	0x02: "If",
	0x04: "Choose",
	0x08: "Skip",
	0x10: "Sum",
	0x20: "Assign",
	0x40: "Space",
	0x41: "SpaceVolatile",
}

// Box functions for range operations
var tRangeFuncs = []func(int, int) int{
	func(a, b int) int { return min(a, b) },
	func(a, b int) int { return max(a, b) },
	func(a, b int) int { return min(a, b) },
	func(a, b int) int { return max(a, b) },
	func(a, b int) int { return min(a, b) },
	func(a, b int) int { return max(a, b) },
}
var tIsectFuncs = []func(int, int) int{
	func(a, b int) int { return max(a, b) },
	func(a, b int) int { return min(a, b) },
	func(a, b int) int { return max(a, b) },
	func(a, b int) int { return min(a, b) },
	func(a, b int) int { return max(a, b) },
	func(a, b int) int { return min(a, b) },
}

// Token not allowed
var tokenNotAllowed = map[int]bool{
	0x01: true, 0x02: true, 0x03: true, 0x10: true, 0x11: true, 0x12: true, 0x13: true,
	0x14: true, 0x15: true, 0x16: true, 0x17: true, 0x18: true, 0x19: true, 0x1A: true,
	0x1B: true, 0x1C: true, 0x1D: true, 0x1E: true, 0x1F: true,
}

// Operator kind dictionary
var OkindDict = map[int]int{
	0x01: 1, 0x02: 1, 0x03: 1, 0x04: 1, 0x05: 1, 0x06: 1, 0x07: 1, 0x08: 1,
	0x09: 1, 0x0A: 1, 0x0B: 1, 0x0C: 1, 0x0D: 1, 0x0E: 1, 0x0F: 1, 0x10: 1,
	0x11: 1, 0x12: 1, 0x13: 1, 0x14: 1, 0x15: 1, 0x16: 1, 0x17: 1, 0x18: 1,
	0x19: 1, 0x1A: 1, 0x1B: 1, 0x1C: 1, 0x1D: 1, 0x1E: 1, 0x1F: 1,
}

// Size tables
var sztab0 = []int{-2, 4, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 8, 4, 2, 2, 3, 9, 8, 2, 3, 8, 4, 7, 5, 5, 5, 2, 4, 7, 4, 7, 2, 2, -2, -2, -2, -2, -2, -2, -2, -2, 3, -2, -2, -2, -2, -2, -2, -2}
var sztab1 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 11, 5, 2, 2, 3, 9, 9, 2, 3, 11, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, 3, -2, -2, -2, -2, -2, -2, -2}
var sztab2 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 11, 5, 2, 2, 3, 9, 9, 3, 4, 11, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2}
var sztab3 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, -2, -2, 2, 2, 3, 9, 9, 3, 4, 15, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, 25, 18, 21, 18, 21, -2, -2}
var sztab4 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -1, -1, -2, -2, 2, 2, 3, 9, 9, 3, 4, 5, 5, 9, 7, 7, 7, 3, 5, 9, 5, 9, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, 7, 7, 11, 7, 11, -2, -2}

// Size dictionary
var szdict = map[int][]int{
	20: sztab0,
	21: sztab0,
	30: sztab1,
	40: sztab2,
	45: sztab2,
	50: sztab3,
	70: sztab3,
	80: sztab4,
}

// Operation names
var onames = []string{
	"Unk00", "Exp", "Tbl", "Add", "Sub", "Mul", "Div", "Power", "Concat", "LT", "LE", "EQ", "GE", "GT", "NE",
	"Isect", "List", "Range", "Uplus", "Uminus", "Percent", "Paren", "MissArg", "Str", "Extended", "Attr",
	"Sheet", "EndSheet", "Err", "Bool", "Int", "Num", "Array", "Func", "FuncVar", "Name", "Ref", "Area",
	"MemArea", "MemErr", "MemNoMem", "MemFunc", "RefErr", "AreaErr", "RefN", "AreaN", "MemAreaN", "MemNoMemN",
	"", "", "", "", "", "", "", "", "FuncCE", "NameX", "Ref3d", "Area3d", "RefErr3d", "AreaErr3d", "", "",
}

// Formula type description map
var fmlaTypeDescrMap = map[int]string{
	1:  "CELL",
	2:  "SHARED",
	4:  "ARRAY",
	8:  "COND-FMT",
	16: "DATA-VAL",
	32: "NAME",
}

// Function definition structure
type funcDef struct {
	name      string
	minArgs   int
	maxArgs   int
	flags     int
	knownArgs int
	retType   string
	args      string
}

// Function definitions
var funcDefs = map[int]funcDef{
	0: {"COUNT", 0, 30, 0x04, 1, "V", "R"},
	1: {"IF", 2, 3, 0x04, 3, "V", "VRR"},
	2: {"ISNA", 1, 1, 0x02, 1, "V", "V"},
	3: {"ISERROR", 1, 1, 0x02, 1, "V", "V"},
	4: {"SUM", 0, 30, 0x04, 1, "V", "R"},
	5: {"AVERAGE", 1, 30, 0x04, 1, "V", "R"},
	6: {"MIN", 1, 30, 0x04, 1, "V", "R"},
	7: {"MAX", 1, 30, 0x04, 1, "V", "R"},
	8: {"ROW", 0, 1, 0x04, 1, "V", "R"},
}

// Arithmetic argument dictionary
var arithArgdict = map[int]interface{}{
	oNUM:  nop,
	oSTRG: func(x interface{}) interface{} { return x.(float64) },
}

// Comparison argument dictionary
var cmpArgdict = map[int]interface{}{
	oNUM:  nop,
	oSTRG: nop,
}

// String argument dictionary
var strgArgdict = map[int]interface{}{
	oNUM:  num2strg,
	oSTRG: nop,
}

// Binary operation rules structure
type binopRule struct {
	argdict    map[int]interface{}
	resultType int
	op         interface{}
	priority   int
	symbol     string
}

// Binary operation rules
var binopRules = map[int]binopRule{
	tAdd:    {arithArgdict, oNUM, opr.add, 30, "+"},
	tSub:    {arithArgdict, oNUM, opr.sub, 30, "-"},
	tMul:    {arithArgdict, oNUM, opr.mul, 40, "*"},
	tDiv:    {arithArgdict, oNUM, opr.truediv, 40, "/"},
	tPower:  {arithArgdict, oNUM, oprPow, 50, "^"},
	tConcat: {strgArgdict, oSTRG, opr.add, 20, "&"},
	tLT:     {cmpArgdict, oBOOL, oprLt, 10, "<"},
	tLE:     {cmpArgdict, oBOOL, oprLe, 10, "<="},
	tEQ:     {cmpArgdict, oBOOL, oprEq, 10, "="},
	tGE:     {cmpArgdict, oBOOL, oprGe, 10, ">="},
	tGT:     {cmpArgdict, oBOOL, oprGt, 10, ">"},
	tNE:     {cmpArgdict, oBOOL, oprNe, 10, "<>"},
}

// Unary operation rules structure
type unopRule struct {
	op       interface{}
	priority int
	prefix   string
	suffix   string
}

// Unary operation rules
var unopRules = map[int]unopRule{
	0x13: {func(x interface{}) interface{} { return -x.(float64) }, 70, "-", ""},        // unary minus
	0x12: {nop, 70, "+", ""},                                                            // unary plus
	0x14: {func(x interface{}) interface{} { return x.(float64) / 100.0 }, 60, "", "%"}, // percent
}

// Operand types
const (
	oUNK  = 0
	oSTRG = 1
	oNUM  = 2
	oBOOL = 3
	oERR  = 4
	oMSNG = 5 // tMissArg
	oREF  = -1
	oREL  = -2
)

// Exported operand type aliases (match Python __all__ names).
const (
	OUNK  = oUNK
	OSTRG = oSTRG
	ONUM  = oNUM
	OBOOL = oBOOL
	OERR  = oERR
	OREF  = oREF
	OREL  = oREL
)

// Formula type constants (match Python __all__ names).
const (
	FMLA_TYPE_CELL     = 1
	FMLA_TYPE_SHARED   = 2
	FMLA_TYPE_ARRAY    = 4
	FMLA_TYPE_COND_FMT = 8
	FMLA_TYPE_DATA_VAL = 16
	FMLA_TYPE_NAME     = 32
)

// Token types
const (
	tAdd    = 0x03
	tSub    = 0x04
	tMul    = 0x05
	tDiv    = 0x06
	tPower  = 0x07
	tConcat = 0x08
	tLT     = 0x09
	tLE     = 0x0A
	tEQ     = 0x0B
	tGE     = 0x0C
	tGT     = 0x0D
	tNE     = 0x0E
)

// Operator functions (equivalent to Python's operator module)
type oprType struct{}

var opr oprType

func (o oprType) add(a, b interface{}) interface{} {
	switch a.(type) {
	case float64:
		switch b.(type) {
		case float64:
			return a.(float64) + b.(float64)
		case string:
			return fmt.Sprintf("%v%v", a, b)
		}
	case string:
		return fmt.Sprintf("%v%v", a, b)
	}
	return nil
}

func (o oprType) sub(a, b interface{}) interface{} {
	if af, aok := a.(float64); aok {
		if bf, bok := b.(float64); bok {
			return af - bf
		}
	}
	return nil
}

func (o oprType) mul(a, b interface{}) interface{} {
	if af, aok := a.(float64); aok {
		if bf, bok := b.(float64); bok {
			return af * bf
		}
	}
	return nil
}

func (o oprType) truediv(a, b interface{}) interface{} {
	if af, aok := a.(float64); aok {
		if bf, bok := b.(float64); bok && bf != 0 {
			return af / bf
		}
	}
	return nil
}

// Comparison operators
func oprLt(a, b interface{}) interface{} { return a.(float64) < b.(float64) }
func oprLe(a, b interface{}) interface{} { return a.(float64) <= b.(float64) }
func oprEq(a, b interface{}) interface{} { return a.(float64) == b.(float64) }
func oprGe(a, b interface{}) interface{} { return a.(float64) >= b.(float64) }
func oprGt(a, b interface{}) interface{} { return a.(float64) > b.(float64) }
func oprNe(a, b interface{}) interface{} { return a.(float64) != b.(float64) }

// Power operator
func oprPow(a, b interface{}) interface{} {
	// Simple power implementation
	result := 1.0
	base := a.(float64)
	exp := int(b.(float64))
	for i := 0; i < exp; i++ {
		result *= base
	}
	return result
}

// ===== TOP LEVEL FUNCTIONS =====


// adjustCellAddrBiff8 function

// adjustCellAddrBiffLe7 function

// getCellAddr function

// getCellRangeAddr function

// getExternsheetLocalRange function

// getExternsheetLocalRangeB57 function

// nop function
func nop(x interface{}) interface{} {
	return x
}

// num2strg function - Attempt to emulate Excel's default conversion from number to string
func num2strg(num interface{}) string {
	s := fmt.Sprintf("%v", num)
	if strings.HasSuffix(s, ".0") {
		s = s[:len(s)-2]
	}
	return s
}

// rownamerel function
func rownamerel(rowx int, rowxrel int, browx *int, r1c1 int) string {
	// if no base rowx is provided, we have to return r1c1
	if browx == nil {
		r1c1 = 1
	}
	if rowxrel == 0 {
		if r1c1 != 0 {
			return fmt.Sprintf("R%d", rowx+1)
		}
		return fmt.Sprintf("$%d", rowx+1)
	}
	if r1c1 != 0 {
		if rowx != 0 {
			return fmt.Sprintf("R[%d]", rowx)
		}
		return "R"
	}
	return fmt.Sprintf("%d", (*browx+rowx)%65536+1)
}

// colnamerel function
func colnamerel(colx int, colxrel int, bcolx *int, r1c1 int) string {
	// if no base colx is provided, we have to return r1c1
	if bcolx == nil {
		r1c1 = 1
	}
	if colxrel == 0 {
		if r1c1 != 0 {
			return fmt.Sprintf("C%d", colx+1)
		}
		return "$" + colname(colx)
	}
	if r1c1 != 0 {
		if colx != 0 {
			return fmt.Sprintf("C[%d]", colx)
		}
		return "C"
	}
	return colname((*bcolx + colx) % 256)
}

// Cellname function - Utility function: (5, 7) => 'H6'
func Cellname(rowx int, colx int) string {
	return fmt.Sprintf("%s%d", colname(colx), rowx+1)
}

// Cellnameabs function - Utility function: (5, 7) => '$H$6'
func Cellnameabs(rowx int, colx int, r1c1 int) string {
	if r1c1 != 0 {
		return fmt.Sprintf("R%dC%d", rowx+1, colx+1)
	}
	return fmt.Sprintf("$%s$%d", colname(colx), rowx+1)
}

// cellnamerel function
func cellnamerel(rowx int, colx int, rowxrel int, colxrel int, browx *int, bcolx *int, r1c1 int) string {
	if rowxrel == 0 && colxrel == 0 {
		return cellnameabs(rowx, colx, r1c1)
	}
	if (rowxrel != 0 && browx == nil) || (colxrel != 0 && bcolx == nil) {
		// must flip the whole cell into R1C1 mode
		r1c1 = 1
	}
	c := colnamerel(colx, colxrel, bcolx, r1c1)
	r := rownamerel(rowx, rowxrel, browx, r1c1)
	if r1c1 != 0 {
		return r + c
	}
	return c + r
}

// Colname function - Utility function: 7 => 'H', 27 => 'AB'

// rangename2d function
func rangename2d(rlo int, rhi int, clo int, chi int, r1c1 int) string {
	if r1c1 != 0 {
		return ""
	}
	if rhi == rlo+1 && chi == clo+1 {
		return cellnameabs(rlo, clo, r1c1)
	}
	return fmt.Sprintf("%s:%s", cellnameabs(rlo, clo, r1c1), cellnameabs(rhi-1, chi-1, r1c1))
}

// rangename2drel function
func rangename2drel(rloRhiCloChi []int, rlorelRhirelClorelChirel []int, browx *int, bcolx *int, r1c1 int) string {
	rlo, rhi, clo, chi := rloRhiCloChi[0], rloRhiCloChi[1], rloRhiCloChi[2], rloRhiCloChi[3]
	rlorel, rhirel, clorel, chirel := rlorelRhirelClorelChirel[0], rlorelRhirelClorelChirel[1], rlorelRhirelClorelChirel[2], rlorelRhirelClorelChirel[3]
	if (rlorel != 0 || rhirel != 0) && browx == nil {
		r1c1 = 1
	}
	if (clorel != 0 || chirel != 0) && bcolx == nil {
		r1c1 = 1
	}
	return fmt.Sprintf("%s:%s",
		cellnamerel(rlo, clo, rlorel, clorel, browx, bcolx, r1c1),
		cellnamerel(rhi-1, chi-1, rhirel, chirel, browx, bcolx, r1c1),
	)
}

// Rangename3d function
func Rangename3d(book interface{}, ref3d interface{}) string {
	r3d := ref3d.(*Ref3D)
	coords := r3d.coords
	return fmt.Sprintf("%s!%s",
		sheetrange(book, coords[0], coords[1]),
		rangename2d(coords[2], coords[3], coords[4], coords[5], 0))
}

// Rangename3drel function
func Rangename3drel(book interface{}, ref3d interface{}, browx *int, bcolx *int, r1c1 int) string {
	r3d := ref3d.(*Ref3D)
	coords := r3d.coords
	relflags := r3d.relflags
	shdesc := sheetrangerel(book, coords[:2], relflags[:2])
	rngdesc := rangename2drel(coords[2:6], relflags[2:6], browx, bcolx, r1c1)
	if shdesc == "" {
		return rngdesc
	}
	return fmt.Sprintf("%s!%s", shdesc, rngdesc)
}

// quotedsheetname function
func quotedsheetname(shnames []string, shx int) string {
	var shname string
	if shx >= 0 {
		shname = shnames[shx]
	} else {
		switch shx {
		case -1:
			shname = "?internal; any sheet?"
		case -2:
			shname = "internal; deleted sheet"
		case -3:
			shname = "internal; macro sheet"
		case -4:
			shname = "<<external>>"
		default:
			shname = fmt.Sprintf("?error %d?", shx)
		}
	}
	if strings.Contains(shname, "'") {
		return "'" + strings.ReplaceAll(shname, "'", "''") + "'"
	}
	if strings.Contains(shname, " ") {
		return "'" + shname + "'"
	}
	return shname
}

// sheetrange function
func sheetrange(book interface{}, slo int, shi int) string {
	bk := book.(*Book)
	shnames := bk.sheetNames
	shdesc := quotedsheetname(shnames, slo)
	if slo != shi-1 {
		shdesc += ":" + quotedsheetname(shnames, shi-1)
	}
	return shdesc
}

// sheetrangerel function
func sheetrangerel(book interface{}, srange interface{}, srangerel interface{}) string {
	sr := srange.([]int)
	srr := srangerel.([]int)
	slo, shi := sr[0], sr[1]
	slorel, shirel := srr[0], srr[1]
	if slorel == 0 && shirel == 0 {
		return sheetrange(book, slo, shi)
	}
	// assert (slo == 0 == shi-1) and slorel and shirel
	if !(slo == 0 && shi-1 == 0 && slorel != 0 && shirel != 0) {
		panic("Assertion failed in sheetrangerel")
	}
	return ""
}

// ===== CLASSES/STRUCTS =====

// FormulaError represents a formula error
type FormulaError struct {
	message string
}

func (e *FormulaError) Error() string {
	return e.message
}

// Operand represents an operand in a formula
type Operand struct {
	kind  int
	value interface{}
	text  string
	rpn   []interface{}
	_rank int
}

// Rank returns the rank of the operand
func (op *Operand) Rank() int {
	return op._rank
}

// SetRank sets the rank of the operand
func (op *Operand) SetRank(rank int) {
	op._rank = rank
}

// Ref3D represents a 3D reference
type Ref3D struct {
	shtxlo   int
	shtxhi   int
	rlo      int
	rhi      int
	clo      int
	chi      int
	coords   []int // [slo, shi, rlo, rhi, clo, chi]
	relflags []int // [slorel, shirel, rlorel, rhirel, clorel, chirel]
}

// NewRef3D creates a new Ref3D
func NewRef3D(coords_relflags ...int) *Ref3D {
	r := &Ref3D{}
	if len(coords_relflags) >= 6 {
		r.coords = coords_relflags[:6]
		r.shtxlo = r.coords[0]
		r.shtxhi = r.coords[1]
		r.rlo = r.coords[2]
		r.rhi = r.coords[3]
		r.clo = r.coords[4]
		r.chi = r.coords[5]
	}
	if len(coords_relflags) >= 12 {
		r.relflags = coords_relflags[6:12]
	} else {
		r.relflags = []int{0, 0, 0, 0, 0, 0}
	}
	return r
}

// dump_formula dumps formula data for debugging
func DumpFormula(bk *Book, data []byte, fmlalen int, bv int, reldelta int, blah int, isname int) {
	if blah != 0 {
		fmt.Fprintf(bk.logfile, "dump_formula %d %d %d\n", fmlalen, bv, len(data))
		HexCharDump(data, 0, fmlalen, bv, bk.logfile, true)
	}
	// Note: this function currently supports BIFF >= 80, but may need updating for older versions
	sztab := szdict[bv]
	pos := 0
	stack := []interface{}{}

	anyRel := 0
	anyErr := 0

	for pos >= 0 && pos < fmlalen {
		op := int(data[pos])
		opcode := op & 0x1f
		optype := (op & 0x60) >> 5
		var opx int
		if optype != 0 {
			opx = opcode + 32
		} else {
			opx = opcode
		}
		oname := onames[opx]

		sz := sztab[opx]
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "Pos:%d Op:0x%02x Name:t%s Sz:%d opcode:%02xh optype:%02xh\n",
				pos, op, oname, sz, opcode, optype)
		}
		if optype == 0 {
			if opcode >= 0x01 && opcode <= 0x02 { // tExp, tTbl
				// reference to a shared formula or table record
				rowx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				colx := int(binary.LittleEndian.Uint16(data[pos+3 : pos+5]))
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "  %d, %d\n", rowx, colx)
				}
			} else if opcode == 0x10 { // tList
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tList pre %v\n", stack)
				}
				if len(stack) < 2 {
					fmt.Fprintf(bk.logfile, "tList: insufficient stack items\n")
					return
				}
				bop := stack[len(stack)-1]
				aop := stack[len(stack)-2]
				stack = stack[:len(stack)-2]
				// Concatenate lists
				result := append(aop.([]interface{}), bop.([]interface{})...)
				spush(stack, result)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tlist post %v\n", stack)
				}
			} else if opcode == 0x11 { // tRange
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRange pre %v\n", stack)
				}
				if len(stack) < 2 {
					fmt.Fprintf(bk.logfile, "tRange: insufficient stack items\n")
					return
				}
				bop := stack[len(stack)-1]
				aop := stack[len(stack)-2]
				stack = stack[:len(stack)-2]
				if len(aop.([]interface{})) != 1 || len(bop.([]interface{})) != 1 {
					fmt.Fprintf(bk.logfile, "tRange: invalid operands\n")
					return
				}
				result := doBoxFuncs(tRangeFuncs, aop.([]interface{})[0].(*Ref3D), bop.([]interface{})[0].(*Ref3D))
				spush(stack, result)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRange post %v\n", stack)
				}
			} else if opcode == 0x0F { // tIsect
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tIsect pre %v\n", stack)
				}
				if len(stack) < 2 {
					fmt.Fprintf(bk.logfile, "tIsect: insufficient stack items\n")
					return
				}
				bop := stack[len(stack)-1]
				aop := stack[len(stack)-2]
				stack = stack[:len(stack)-2]
				if len(aop.([]interface{})) != 1 || len(bop.([]interface{})) != 1 {
					fmt.Fprintf(bk.logfile, "tIsect: invalid operands\n")
					return
				}
				result := doBoxFuncs(tIsectFuncs, aop.([]interface{})[0].(*Ref3D), bop.([]interface{})[0].(*Ref3D))
				spush(stack, result)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tIsect post %v\n", stack)
				}
			} else if opcode == 0x19 { // tAttr
				subop := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				nc := int(data[pos+3])
				subname := tAttrNames[subop]
				if subname == "" {
					subname = "??Unknown??"
				}
				if subop == 0x04 { // Choose
					sz = nc*2 + 6
				} else {
					sz = 4
				}
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "   subop=%02xh subname=t%s sz=%d nc=%02xh\n", subop, subname, sz, nc)
				}
			} else if opcode == 0x17 { // tStr
				var strg string
				var newpos int
				if bv <= 70 {
					nc := int(data[pos+1])
					strg = string(data[pos+2 : pos+2+nc]) // left in 8-bit encoding
					sz = nc + 2
				} else {
					strg, newpos = unpack_unicode_update_pos(data, pos+1, 1, -1)
					sz = newpos - pos
				}
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "   sz=%d strg=%q\n", sz, strg)
				}
			} else {
				if sz <= 0 {
					fmt.Fprintf(bk.logfile, "**** Dud size; exiting ****\n")
					return
				}
			}
			pos += sz
			continue
		}

		// Handle different opcode types
		if opcode == 0x00 { // tArray
			// No special processing needed for tArray
		} else if opcode == 0x01 { // tFunc
			nb := 1
			if bv >= 40 {
				nb = 2
			}
			funcx := 0
			if nb == 1 {
				funcx = int(data[pos+1])
			} else {
				funcx = int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
			}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   FuncID=%d\n", funcx)
			}
		} else if opcode == 0x02 { // tFuncVar
			nb := 1
			if bv >= 40 {
				nb = 2
			}
			nargs := int(data[pos+1])
			funcx := 0
			if nb == 1 {
				funcx = int(data[pos+2])
			} else {
				funcx = int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
			}
			prompt := nargs >> 7
			nargs &= 0x7F
			macro := funcx >> 15
			funcx &= 0x7FFF
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   FuncID=%d nargs=%d macro=%d prompt=%d\n", funcx, nargs, macro, prompt)
			}
		} else if opcode == 0x03 { // tName
			namex := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   namex=%d\n", namex)
			}
		} else if opcode == 0x04 { // tRef
			rowx, colx, rowRel, colRel := getCellAddr(data, pos+1, bv, reldelta, nil, nil)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   (%d, %d, %d, %d)\n", rowx, colx, rowRel, colRel)
			}
		} else if opcode == 0x05 { // tArea
			res1, res2 := getCellRangeAddr(data, pos+1, bv, reldelta, nil, nil)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v %v\n", res1, res2)
			}
		} else if opcode == 0x09 { // tMemFunc
			nb := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %d bytes of cell ref formula\n", nb)
			}
		} else if opcode == 0x0C { // tRefN
			rowx, colx, rowRel, colRel := getCellAddr(data, pos+1, bv, 1, nil, nil) // note *ALL* tRefN usage has signed offset for relative addresses
			anyRel = 1
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   (%d, %d, %d, %d)\n", rowx, colx, rowRel, colRel)
			}
		} else if opcode == 0x0D { // tAreaN
			res1, res2 := getCellRangeAddr(data, pos+1, bv, 1, nil, nil) // note *ALL* tAreaN usage has signed offset for relative addresses
			anyRel = 1
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v %v\n", res1, res2)
			}
		} else if opcode == 0x1A { // tRef3d
			refx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
			rowx, colx, rowRel, colRel := getCellAddr(data, pos+3, bv, reldelta, nil, nil)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   refx=%d (%d, %d, %d, %d)\n", refx, rowx, colx, rowRel, colRel)
			}
			if rowRel != 0 || colRel != 0 {
				anyRel = 1
			}
			shx1, shx2 := getExternsheetLocalRange(bk, refx, blah)
			if shx1 < -1 {
				anyErr = 1
			}
			coords := []int{shx1, shx2 + 1, rowx, rowx + 1, colx, colx + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			if optype == 1 {
				spush(stack, coords)
			}
		} else if opcode == 0x1B { // tArea3d
			refx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
			res1, res2 := getCellRangeAddr(data, pos+3, bv, reldelta, nil, nil)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   refx=%d %v %v\n", refx, res1, res2)
			}
			rowx1, colx1, rowRel1, colRel1 := res1[0], res1[1], res1[2], res1[3]
			rowx2, colx2, rowRel2, colRel2 := res2[0], res2[1], res2[2], res2[3]
			if rowRel1 != 0 || colRel1 != 0 || rowRel2 != 0 || colRel2 != 0 {
				anyRel = 1
			}
			shx1, shx2 := getExternsheetLocalRange(bk, refx, blah)
			if shx1 < -1 {
				anyErr = 1
			}
			coords := []int{shx1, shx2 + 1, rowx1, rowx2 + 1, colx1, colx2 + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			if optype == 1 {
				spush(stack, coords)
			}
		} else if opcode == 0x19 { // tNameX
			refx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
			namex := int(binary.LittleEndian.Uint16(data[pos+3 : pos+5]))
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   refx=%d namex=%d\n", refx, namex)
			}
		} else if _, exists := errorOpcodes[opcode]; exists {
			anyErr = 1
		} else {
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "FORMULA: /// Not handled yet: t%s\n", oname)
			}
			anyErr = 1
		}

		if sz <= 0 {
			fmt.Fprintf(bk.logfile, "**** Dud size; exiting ****\n")
			return
		}
		pos += sz
	}

	if blah != 0 {
		fmt.Fprintf(bk.logfile, "End of formula. any_rel=%d any_err=%d stack=%v\n",
			boolToInt(anyRel != 0), anyErr, stack)
		if len(stack) >= 2 {
			fmt.Fprintf(bk.logfile, "*** Stack has unprocessed args\n")
		}
	}
}

// spush appends an item to the stack
func spush(stack []interface{}, item interface{}) []interface{} {
	return append(stack, item)
}

// boolToInt converts boolean to int
func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

// do_box_funcs performs box functions on two Ref3D objects
func doBoxFuncs(boxFuncs []func(int, int) int, boxa, boxb *Ref3D) []int {
	result := make([]int, 6)
	for i, fn := range boxFuncs {
		result[i] = fn(boxa.coords[i], boxb.coords[i])
	}
	return result
}

// adjust_cell_addr_biff8 adjusts cell address for BIFF8
func adjustCellAddrBiff8(rowval, colval, reldelta int, browx, bcolx interface{}) (int, int, int, int) {
	rowRel := (colval >> 15) & 1
	colRel := (colval >> 14) & 1
	rowx := rowval
	colx := colval & 0xff

	if reldelta != 0 {
		if rowRel != 0 && rowx >= 32768 {
			rowx -= 65536
		}
		if colRel != 0 && colx >= 128 {
			colx -= 256
		}
	} else {
		if browxInt, ok := browx.(int); ok && rowRel != 0 {
			rowx -= browxInt
		}
		if bcolxInt, ok := bcolx.(int); ok && colRel != 0 {
			colx -= bcolxInt
		}
	}
	return rowx, colx, rowRel, colRel
}

// adjust_cell_addr_biff_le7 adjusts cell address for BIFF <= 7
func adjustCellAddrBiffLe7(rowval, colval, reldelta int, browx, bcolx interface{}) (int, int, int, int) {
	rowRel := (rowval >> 15) & 1
	colRel := (rowval >> 14) & 1
	rowx := rowval & 0x3fff
	colx := colval

	if reldelta != 0 {
		if rowRel != 0 && rowx >= 8192 {
			rowx -= 16384
		}
		if colRel != 0 && colx >= 128 {
			colx -= 256
		}
	} else {
		if browxInt, ok := browx.(int); ok && rowRel != 0 {
			rowx -= browxInt
		}
		if bcolxInt, ok := bcolx.(int); ok && colRel != 0 {
			colx -= bcolxInt
		}
	}
	return rowx, colx, rowRel, colRel
}

// get_cell_addr gets cell address from binary data
func getCellAddr(data []byte, pos, bv, reldelta int, browx, bcolx interface{}) (int, int, int, int) {
	if bv >= 80 {
		rowval := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		colval := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		return adjustCellAddrBiff8(rowval, colval, reldelta, browx, bcolx)
	} else {
		rowval := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		colval := int(data[pos+2])
		return adjustCellAddrBiffLe7(rowval, colval, reldelta, browx, bcolx)
	}
}

// get_cell_range_addr gets cell range address from binary data
func getCellRangeAddr(data []byte, pos, bv, reldelta int, browx, bcolx interface{}) ([4]int, [4]int) {
	if bv >= 80 {
		row1val := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		row2val := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		col1val := int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))
		col2val := int(binary.LittleEndian.Uint16(data[pos+6 : pos+8]))
		r1x, c1x, r1rel, c1rel := adjustCellAddrBiff8(row1val, col1val, reldelta, browx, bcolx)
		r2x, c2x, r2rel, c2rel := adjustCellAddrBiff8(row2val, col2val, reldelta, browx, bcolx)
		return [4]int{r1x, c1x, r1rel, c1rel}, [4]int{r2x, c2x, r2rel, c2rel}
	} else {
		row1val := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		row2val := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		col1val := int(data[pos+4])
		col2val := int(data[pos+5])
		r1x, c1x, r1rel, c1rel := adjustCellAddrBiffLe7(row1val, col1val, reldelta, browx, bcolx)
		r2x, c2x, r2rel, c2rel := adjustCellAddrBiffLe7(row2val, col2val, reldelta, browx, bcolx)
		return [4]int{r1x, c1x, r1rel, c1rel}, [4]int{r2x, c2x, r2rel, c2rel}
	}
}

// get_externsheet_local_range gets external sheet local range
func getExternsheetLocalRange(bk *Book, refx, blah int) (int, int) {
	if refx < 0 || refx >= len(bk.externsheetInfo) {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "!!! get_externsheet_local_range: refx=%d, not in range(%d)\n",
				refx, len(bk.externsheetInfo))
		}
		return -101, -101
	}

	info := bk.externsheetInfo[refx]
	refRecordx := info[0]
	refFirstSheetx := info[1]
	refLastSheetx := info[2]

	if bk.supbookAddinsInx != nil && refRecordx == *bk.supbookAddinsInx {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range(refx=%d) -> addins %v\n", refx, info)
		}
		if refFirstSheetx != 0xFFFE || refLastSheetx != 0xFFFE {
			return -101, -101 // error
		}
		return -5, -5
	}

	if bk.supbookLocalsInx != nil && refRecordx != *bk.supbookLocalsInx {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range(refx=%d) -> external %v\n", refx, info)
		}
		return -4, -4 // external reference
	}

	if refFirstSheetx == 0xFFFE && refLastSheetx == 0xFFFE {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range(refx=%d) -> unspecified sheet %v\n", refx, info)
		}
		return -1, -1 // internal reference, any sheet
	}

	if refFirstSheetx == 0xFFFF && refLastSheetx == 0xFFFF {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range(refx=%d) -> deleted sheet(s)\n", refx)
		}
		return -2, -2 // internal reference, deleted sheet(s)
	}

	nsheets := len(bk.allSheetsMap)
	if !(0 <= refFirstSheetx && refFirstSheetx <= refLastSheetx && refLastSheetx < nsheets) {
		if blah != 0 {
			fmt.Fprintf(bk.logfile, "/// get_externsheet_local_range(refx=%d) -> %v\n", refx, info)
			fmt.Fprintf(bk.logfile, "--- first/last sheet not in range(%d)\n", nsheets)
		}
		return -102, -102 // stuffed up somewhere
	}

	xlrdSheetx1 := bk.allSheetsMap[refFirstSheetx]
	xlrdSheetx2 := bk.allSheetsMap[refLastSheetx]
	if !(0 <= xlrdSheetx1 && xlrdSheetx1 <= xlrdSheetx2) {
		return -3, -3 // internal reference, but to a macro sheet
	}
	return xlrdSheetx1, xlrdSheetx2
}

// evaluateNameFormula evaluates a named formula
func EvaluateNameFormula(bk *Book, nobj *Name, namex int, blah int, level int) {
	if level > StackAlarmLevel {
		blah = 1
	}
	data := nobj.RawFormula
	fmlalen := nobj.BasicFormulaLen
	bv := bk.BiffVersion
	reldelta := 1 // All defined name formulas use "Method B" [OOo docs]

	if blah != 0 {
		fmt.Fprintf(bk.logfile, "::: evaluate_name_formula %d %q %d %d %v level=%d\n",
			namex, nobj.Name, fmlalen, bv, data, level)
		hexCharDump(data, 0, fmlalen, bk.logfile.(*os.File))
	}

	if level > StackPanicLevel {
		panic(fmt.Sprintf("Excessive indirect references in NAME formula"))
	}

	sztab := szdict[bv]
	pos := 0
	stack := []interface{}{}
	anyRel := 0
	anyErr := 0
	anyExternal := 0
	unkOpnd := &Operand{kind: oUNK, value: nil}
	errorOpnd := &Operand{kind: oERR, value: nil}
	spush := func(item interface{}) {
		stack = append(stack, item)
	}

	doBinop := func(opcd int, stk []interface{}) {
		if len(stk) < 2 {
			return
		}
		bop := stk[len(stk)-1].(*Operand)
		aop := stk[len(stk)-2].(*Operand)
		rule := binopRules[opcd]
		argdict := rule.argdict
		resultType := rule.resultType
		fn := rule.op.(func(interface{}, interface{}) interface{})
		rank := rule.priority
		sym := rule.symbol

		otext := ""
		if aop._rank < rank {
			otext += "("
		}
		otext += aop.text
		if aop._rank < rank {
			otext += ")"
		}
		otext += sym
		if bop._rank < rank {
			otext += "("
		}
		otext += bop.text
		if bop._rank < rank {
			otext += ")"
		}

		resop := &Operand{kind: resultType, value: nil, _rank: rank, text: otext}

		if bop.value == nil || aop.value == nil {
			stk = stk[:len(stk)-2]
			stk = append(stk, resop)
			return
		}

		var bconv, aconv func(interface{}) interface{}
		var ok bool
		if bconv, ok = argdict[bop.kind].(func(interface{}) interface{}); !ok {
			stk = stk[:len(stk)-2]
			stk = append(stk, resop)
			return
		}
		if aconv, ok = argdict[aop.kind].(func(interface{}) interface{}); !ok {
			stk = stk[:len(stk)-2]
			stk = append(stk, resop)
			return
		}

		bval := bconv(bop.value)
		aval := aconv(aop.value)
		result := fn(aval, bval)
		if resultType == oBOOL {
			if result.(bool) {
				result = 1
			} else {
				result = 0
			}
		}
		resop.value = result
		stk = stk[:len(stk)-2]
		stk = append(stk, resop)
	}

	doUnaryop := func(opcode int, resultType int, stk []interface{}) {
		if len(stk) < 1 {
			return
		}
		aop := stk[len(stk)-1].(*Operand)
		rule := unopRules[opcode]
		fn := rule.op.(func(interface{}) interface{})
		rank := rule.priority
		sym1 := rule.prefix
		sym2 := rule.suffix

		otext := sym1
		if aop._rank < rank {
			otext += "("
		}
		otext += aop.text
		if aop._rank < rank {
			otext += ")"
		}
		otext += sym2

		val := aop.value
		if val != nil {
			val = fn(val)
		}
		res := &Operand{kind: resultType, value: val, _rank: rank, text: otext}
		stk = stk[:len(stk)-1]
		stk = append(stk, res)
	}

	notInNameFormula := func(opArg int, onameArg string) {
		msg := fmt.Sprintf("ERROR *** Token 0x%02x (%s) found in NAME formula", opArg, onameArg)
		panic(msg)
	}

	if fmlalen == 0 {
		stack = []interface{}{unkOpnd}
	}

	for pos < fmlalen {
		op := int(data[pos])
		opcode := op & 0x1f
		optype := (op & 0x60) >> 5
		opx := opcode
		if optype != 0 {
			opx = opcode + 32
		}
		oname := onames[opx]
		sz := sztab[opx]

		if blah != 0 {
			fmt.Fprintf(bk.logfile, "Pos:%d Op:0x%02x Name:%s Sz:%d opcode:%02xh optype:%02xh\n",
				pos, op, oname, sz, opcode, optype)
			fmt.Fprintf(bk.logfile, "Stack = %v\n", stack)
		}

		if sz == -2 {
			msg := fmt.Sprintf(`ERROR *** Unexpected token 0x%02x ("%s"); biff_version=%d`, op, oname, bv)
			panic(msg)
		}

		if optype == 0 {
			if 0x00 <= opcode && opcode <= 0x02 { // unk_opnd, tExp, tTbl
				notInNameFormula(op, oname)
			} else if 0x03 <= opcode && opcode <= 0x0E { // Add, Sub, Mul, Div, Power, tConcat, tLT, ..., tNE
				doBinop(opcode, stack)
			} else if opcode == 0x0F { // tIsect
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tIsect pre %v\n", stack)
				}
				if len(stack) < 2 {
					continue
				}
				bop := stack[len(stack)-1].(*Operand)
				aop := stack[len(stack)-2].(*Operand)
				sym := " "
				rank := 80
				otext := ""
				if aop._rank < rank {
					otext += "("
				}
				otext += aop.text
				if aop._rank < rank {
					otext += ")"
				}
				otext += sym
				if bop._rank < rank {
					otext += "("
				}
				otext += bop.text
				if bop._rank < rank {
					otext += ")"
				}

				res := &Operand{kind: oREF}
				res.text = otext
				if bop.kind == oERR || aop.kind == oERR {
					res.kind = oERR
				} else if bop.kind == oUNK || aop.kind == oUNK {
					// This can happen with undefined (go search in the current sheet) labels.
					// For example =Bob Sales
					// Each label gets a NAME record with an empty formula (!)
					// Evaluation of the tName token classifies it as oUNK
					// res.kind = oREF
				} else if bop.kind == oREF && aop.kind == oREF {
					if aop.value != nil && bop.value != nil {
						assert(len(aop.value.([]*Ref3D)) == 1)
						assert(len(bop.value.([]*Ref3D)) == 1)
						coords := doBoxFuncs(tIsectFuncs, aop.value.([]*Ref3D)[0], bop.value.([]*Ref3D)[0])
						res.value = []*Ref3D{NewRef3D(coords...)}
					}
				} else if bop.kind == oREL && aop.kind == oREL {
					res.kind = oREL
					if aop.value != nil && bop.value != nil {
						assert(len(aop.value.([]*Ref3D)) == 1)
						assert(len(bop.value.([]*Ref3D)) == 1)
						coords := doBoxFuncs(tIsectFuncs, aop.value.([]*Ref3D)[0], bop.value.([]*Ref3D)[0])
						relfa := aop.value.([]*Ref3D)[0].relflags
						relfb := bop.value.([]*Ref3D)[0].relflags
						if relfa != nil && relfb != nil && reflect.DeepEqual(relfa, relfb) {
							relflags := make([]int, len(relfa))
							copy(relflags, relfa)
							res.value = []*Ref3D{NewRef3D(append(coords, relflags...)...)}
						}
					}
				}
				stack = stack[:len(stack)-2]
				spush(res)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tIsect post %v\n", stack)
				}
			} else if opcode == 0x10 { // tList
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tList pre %v\n", stack)
				}
				if len(stack) < 2 {
					continue
				}
				bop := stack[len(stack)-1].(*Operand)
				aop := stack[len(stack)-2].(*Operand)
				sym := ","
				rank := 80
				otext := ""
				if aop._rank < rank {
					otext += "("
				}
				otext += aop.text
				if aop._rank < rank {
					otext += ")"
				}
				otext += sym
				if bop._rank < rank {
					otext += "("
				}
				otext += bop.text
				if bop._rank < rank {
					otext += ")"
				}

				res := &Operand{kind: oREF, value: nil, _rank: rank, text: otext}
				if bop.kind == oERR || aop.kind == oERR {
					res.kind = oERR
				} else if (bop.kind == oREF || bop.kind == oREL) && (aop.kind == oREF || aop.kind == oREL) {
					res.kind = oREF
					if aop.kind == oREL || bop.kind == oREL {
						res.kind = oREL
					}
					if aop.value != nil && bop.value != nil {
						aopVal := aop.value.([]*Ref3D)
						bopVal := bop.value.([]*Ref3D)
						assert(len(aopVal) >= 1)
						assert(len(bopVal) == 1)
						res.value = append(aopVal, bopVal...)
					}
				}
				stack = stack[:len(stack)-2]
				spush(res)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tList post %v\n", stack)
				}
			} else if opcode == 0x11 { // tRange
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRange pre %v\n", stack)
				}
				if len(stack) < 2 {
					continue
				}
				bop := stack[len(stack)-1].(*Operand)
				aop := stack[len(stack)-2].(*Operand)
				sym := ":"
				rank := 80
				otext := ""
				if aop._rank < rank {
					otext += "("
				}
				otext += aop.text
				if aop._rank < rank {
					otext += ")"
				}
				otext += sym
				if bop._rank < rank {
					otext += "("
				}
				otext += bop.text
				if bop._rank < rank {
					otext += ")"
				}

				res := &Operand{kind: oREF, value: nil, _rank: rank, text: otext}
				if bop.kind == oERR || aop.kind == oERR {
					res.kind = oERR
				} else if bop.kind == oREF && aop.kind == oREF {
					if aop.value != nil && bop.value != nil {
						assert(len(aop.value.([]*Ref3D)) == 1)
						assert(len(bop.value.([]*Ref3D)) == 1)
						coords := doBoxFuncs(tRangeFuncs, aop.value.([]*Ref3D)[0], bop.value.([]*Ref3D)[0])
						res.value = []*Ref3D{NewRef3D(coords...)}
					}
				} else if bop.kind == oREL && aop.kind == oREL {
					res.kind = oREL
					if aop.value != nil && bop.value != nil {
						assert(len(aop.value.([]*Ref3D)) == 1)
						assert(len(bop.value.([]*Ref3D)) == 1)
						coords := doBoxFuncs(tRangeFuncs, aop.value.([]*Ref3D)[0], bop.value.([]*Ref3D)[0])
						relfa := aop.value.([]*Ref3D)[0].relflags
						relfb := bop.value.([]*Ref3D)[0].relflags
						if relfa != nil && relfb != nil && reflect.DeepEqual(relfa, relfb) {
							relflags := make([]int, len(relfa))
							copy(relflags, relfa)
							res.value = []*Ref3D{NewRef3D(append(coords, relflags...)...)}
						}
					}
				}
				stack = stack[:len(stack)-2]
				spush(res)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRange post %v\n", stack)
				}
			} else if 0x12 <= opcode && opcode <= 0x14 { // tUplus, tUminus, tPercent
				doUnaryop(opcode, oNUM, stack)
			} else if opcode == 0x15 { // tParen
				// source cosmetics
			} else if opcode == 0x16 { // tMissArg
				spush(&Operand{kind: oMSNG, value: nil, _rank: LeafRank, text: ""})
			} else if opcode == 0x17 { // tStr
				var strg string
				var newpos int
				if bv <= 70 {
					strg, newpos = unpackStringUpdatePos(data, pos+1, bk.Encoding, 1)
				} else {
					strg, newpos = unpackUnicodeUpdatePos(data, pos+1, 1)
				}
				sz = newpos - pos
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "   sz=%d strg=%q\n", sz, strg)
				}
				text := "\"" + strings.ReplaceAll(strg, "\"", "\"\"") + "\""
				spush(&Operand{kind: oSTRG, value: strg, _rank: LeafRank, text: text})
			} else if opcode == 0x18 { // tExtended
				// new with BIFF 8
				panic("tExtended token not implemented")
			} else if opcode == 0x19 { // tAttr
				result, err := unpack("<BH", data[pos+1:pos+4])
				if err != nil {
					panic(err)
				}
				values := result.([]interface{})
				subop := values[0].(uint8)
				nc := values[1].(uint16)
				subname := tAttrNames[int(subop)]
				if subop == 0x04 { // Choose
					sz = int(nc)*2 + 6
				} else if subop == 0x10 { // Sum (single arg)
					sz = 4
					if blah != 0 {
						fmt.Fprintf(bk.logfile, "tAttrSum %v\n", stack)
					}
					if len(stack) >= 1 {
						aop := stack[len(stack)-1].(*Operand)
						otext := "SUM(" + aop.text + ")"
						stack[len(stack)-1] = &Operand{kind: oNUM, value: nil, _rank: FuncRank, text: otext}
					}
				} else {
					sz = 4
				}
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "   subop=%02xh subname=%s sz=%d nc=%02xh\n", subop, subname, sz, nc)
				}
			} else if 0x1A <= opcode && opcode <= 0x1B { // tSheet, tEndSheet
				assert(bv < 50)
				panic("tSheet & tEndsheet tokens not implemented")
			} else if 0x1C <= opcode && opcode <= 0x1F { // tErr, tBool, tInt, tNum
				inx := opcode - 0x1C
				nb := []int{1, 1, 2, 8}[inx]
				kind := []int{oERR, oBOOL, oNUM, oNUM}[inx]
				value, _ := unpack([]string{"<B", "<B", "<H", "<d"}[inx], data[pos+1:pos+1+nb])
				var text string
				if inx == 2 { // tInt
					value = float64(value.(int16))
					text = fmt.Sprintf("%v", value)
				} else if inx == 3 { // tNum
					text = fmt.Sprintf("%v", value)
				} else if inx == 1 { // tBool
					if value.(uint8) != 0 {
						text = "TRUE"
					} else {
						text = "FALSE"
					}
				} else {
					text = "\"" + errorTextFromCode[int(value.(uint8))] + "\""
				}
				spush(&Operand{kind: kind, value: value, _rank: LeafRank, text: text})
			} else {
				panic(fmt.Sprintf("Unhandled opcode: 0x%02x", opcode))
			}
			if sz <= 0 {
				panic(fmt.Sprintf("Size not set for opcode 0x%02x", opcode))
			}
			pos += sz
			continue
		}

		if opcode == 0x00 { // tArray
			spush(unkOpnd)
		} else if opcode == 0x01 { // tFunc
			nb := 1
			if bv >= 40 {
				nb = 2
			}
			funcx, _ := unpack(fmt.Sprintf("<%sH", []string{"", "B"}[nb-1]), data[pos+1:pos+1+nb])
			funcAttrs, ok := funcDefs[funcx.(int)]
			if !ok {
				fmt.Fprintf(bk.logfile, "*** formula/tFunc unknown FuncID:%d\n", funcx)
				spush(unkOpnd)
			} else {
				funcName := funcAttrs.name
				nargs := funcAttrs.minArgs
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "    FuncID=%d name=%s nargs=%d\n", funcx, funcName, nargs)
				}
				assert(len(stack) >= nargs)
				var otext string
				if nargs > 0 {
					argtext := make([]string, nargs)
					for i := 0; i < nargs; i++ {
						argtext[i] = stack[len(stack)-nargs+i].(*Operand).text
					}
					otext = funcName + "(" + strings.Join(argtext, listsep) + ")"
					for i := 0; i < nargs; i++ {
						stack = stack[:len(stack)-1]
					}
				} else {
					otext = funcName + "()"
				}
				res := &Operand{kind: oUNK, value: nil, _rank: FuncRank, text: otext}
				spush(res)
			}
		} else if opcode == 0x02 { // tFuncVar
			nb := 1
			if bv >= 40 {
				nb = 2
			}
			nargs_funcx, _ := unpack(fmt.Sprintf("<B%sH", []string{"", "B"}[nb-1]), data[pos+1:pos+2+nb])
			nargs := nargs_funcx.([]interface{})[0].(uint8)
			funcx := nargs_funcx.([]interface{})[1].(uint16)
			prompt := nargs >> 7
			nargs &= 0x7F
			macro := funcx >> 15
			funcx &= 0x7FFF
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   FuncID=%d nargs=%d macro=%d prompt=%d\n", funcx, nargs, macro, prompt)
			}
			funcAttrs, ok := funcDefs[int(funcx)]
			if !ok {
				fmt.Fprintf(bk.logfile, "*** formula/tFuncVar unknown FuncID:%d\n", funcx)
				spush(unkOpnd)
			} else {
				funcName := funcAttrs.name
				minargs := funcAttrs.minArgs
				maxargs := funcAttrs.maxArgs
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "    name: %s, min~max args: %d~%d\n", funcName, minargs, maxargs)
				}
				assert(minargs <= int(nargs) && int(nargs) <= maxargs)
				assert(len(stack) >= int(nargs))
				assert(len(stack) >= int(nargs))
				argtext := make([]string, nargs)
				for i := 0; i < int(nargs); i++ {
					argtext[i] = stack[len(stack)-int(nargs)+i].(*Operand).text
				}
				otext := funcName + "(" + strings.Join(argtext, listsep) + ")"
				res := &Operand{kind: oUNK, value: nil, _rank: FuncRank, text: otext}
				if funcx == 1 { // IF
					testarg := stack[len(stack)-int(nargs)].(*Operand)
					if testarg.kind != oNUM && testarg.kind != oBOOL {
						if blah != 0 && testarg.kind != oUNK {
							fmt.Fprintf(bk.logfile, "IF testarg kind?\n")
						}
					} else if testarg.value != 0 && testarg.value != 1 {
						if blah != 0 && testarg.value != nil {
							fmt.Fprintf(bk.logfile, "IF testarg value?\n")
						}
					} else {
						if int(nargs) == 2 && testarg.value == 0 {
							// IF(FALSE, tv) => FALSE
							res.kind = oBOOL
							res.value = 0
						} else {
							respos := len(stack) - int(nargs) + 2 - int(testarg.value.(int))
							chosen := stack[respos].(*Operand)
							if chosen.kind == oMSNG {
								res.kind = oNUM
								res.value = 0
							} else {
								res.kind = chosen.kind
								res.value = chosen.value
							}
							if blah != 0 {
								fmt.Fprintf(bk.logfile, "$$$$$$ IF => constant\n")
							}
						}
					}
				} else if funcx == 100 { // CHOOSE
					testarg := stack[len(stack)-int(nargs)].(*Operand)
					if testarg.kind == oNUM {
						if 1 <= testarg.value.(int) && testarg.value.(int) < int(nargs) {
							chosen := stack[len(stack)-int(nargs)+testarg.value.(int)].(*Operand)
							if chosen.kind == oMSNG {
								res.kind = oNUM
								res.value = 0
							} else {
								res.kind = chosen.kind
								res.value = chosen.value
							}
						}
					}
				}
				for i := 0; i < int(nargs); i++ {
					stack = stack[:len(stack)-1]
				}
				spush(res)
			}
		} else if opcode == 0x03 { // tName
			tgtnamex_raw, _ := unpack("<H", data[pos+1:pos+3])
			tgtnamex := int(tgtnamex_raw.(uint16)) - 1
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   tgtnamex=%d\n", tgtnamex)
			}
			tgtobj := bk.NameObjList[tgtnamex]
			if !tgtobj.Evaluated {
				// recursive
				evaluateNameFormula(bk, tgtobj, tgtnamex, blah, level+1)
			}
			var res *Operand
			if tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0 {
				if blah != 0 {
					tgtobj.Dump(bk.logfile, "!!! tgtobj has problems!!!", "-----------       --------", 0)
				}
				res = &Operand{kind: oUNK, value: nil}
				anyErr = boolToInt(anyErr != 0 || tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0)
				anyRel = boolToInt(anyRel != 0 || tgtobj.AnyRel != 0)
			} else {
				assert(len(tgtobj.Stack) == 1)
				res = copyOperand(tgtobj.Stack[0])
			}
			res._rank = LeafRank
			if tgtobj.Scope == -1 {
				res.text = tgtobj.Name
			} else {
				res.text = bk.SheetNames()[tgtobj.Scope] + "!" + tgtobj.Name
			}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "    tName: setting text to %q\n", res.text)
			}
			spush(res)
		} else if opcode == 0x04 { // tRef
			// not_in_name_formula(op, oname)
			rowx, colx, rowRel, colRel := getCellAddr(data, pos+1, bv, reldelta, nil, nil)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "  (%d, %d, %d, %d)\n", rowx, colx, rowRel, colRel)
			}
			shx1 := 0
			shx2 := 0 // N.B. relative to the CURRENT SHEET
			anyRel = 1
			coords := []int{shx1, shx2 + 1, rowx, rowx + 1, colx, colx + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			if optype == 1 {
				relflags := []int{1, 1, rowRel, rowRel, colRel, colRel}
				resOp = &Operand{kind: oREL, value: []*Ref3D{NewRef3D(append(coords, relflags...)...)}}
			}
			spush(resOp)
		} else if opcode == 0x05 { // tArea
			// not_in_name_formula(op, oname)
			res1, res2 := getCellRangeAddr(data, pos+1, bv, reldelta, nil, nil)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "  %v %v\n", res1, res2)
			}
			rowx1, colx1, rowRel1, colRel1 := res1[0], res1[1], res1[2], res1[3]
			rowx2, colx2, rowRel2, colRel2 := res2[0], res2[1], res2[2], res2[3]
			shx1 := 0
			shx2 := 0 // N.B. relative to the CURRENT SHEET
			anyRel = 1
			coords := []int{shx1, shx2 + 1, rowx1, rowx2 + 1, colx1, colx2 + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			if optype == 1 {
				relflags := []int{1, 1, rowRel1, rowRel2, colRel1, colRel2}
				resOp = &Operand{kind: oREL, value: []*Ref3D{NewRef3D(append(coords, relflags...)...)}}
			}
			spush(resOp)
		} else if opcode == 0x06 { // tMemArea
			notInNameFormula(op, oname)
		} else if opcode == 0x09 { // tMemFunc
			nb, _ := unpack("<H", data[pos+1:pos+3])
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "  %d bytes of cell ref formula\n", nb)
			}
			// no effect on stack
		} else if opcode == 0x0C { // tRefN
			notInNameFormula(op, oname)
		} else if opcode == 0x0D { // tAreaN
			notInNameFormula(op, oname)
		} else if opcode == 0x1A { // tRef3d
			var refx int
			var rowx, colx, rowRel, colRel int
			var shx1, shx2 int
			if bv >= 80 {
				rowx, colx, rowRel, colRel = getCellAddr(data, pos+3, bv, reldelta, nil, nil)
				refx_raw, _ := unpack("<H", data[pos+1:pos+3])
				refx = int(refx_raw.(uint16))
				shx1, shx2 = getExternsheetLocalRange(bk, refx, blah)
			} else {
				rowx, colx, rowRel, colRel = getCellAddr(data, pos+15, bv, reldelta, nil, nil)
				result, _ := unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
				values := result.([]interface{})
				rawExtshtx := values[0].(int16)
				rawShx1 := values[1].(int16)
				rawShx2 := values[2].(int16)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRef3d %d %d %d\n", rawExtshtx, rawShx1, rawShx2)
				}
				shx1, shx2 = getExternsheetLocalRangeB57(bk, int(rawExtshtx), int(rawShx1), int(rawShx2), blah)
			}
			isRel := rowRel != 0 || colRel != 0
			anyRel = boolToInt(anyRel != 0 || isRel)
			coords := []int{shx1, shx2 + 1, rowx, rowx + 1, colx, colx + 1}
			anyErr = boolToInt(anyErr != 0 || shx1 < -1)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			var ref3d *Ref3D
			if isRel {
				relflags := []int{0, 0, rowRel, rowRel, colRel, colRel}
				ref3d = NewRef3D(append(coords, relflags...)...)
				resOp.kind = oREL
				resOp.text = rangename3drel(bk, ref3d, 1)
			} else {
				ref3d = NewRef3D(coords...)
				resOp.kind = oREF
				resOp.text = rangename3d(bk, ref3d)
			}
			resOp._rank = LeafRank
			if optype == 1 {
				resOp.value = []*Ref3D{ref3d}
			}
			spush(resOp)
		} else if opcode == 0x1B { // tArea3d
			var refx int
			var res1, res2 [4]int
			var shx1, shx2 int
			if bv >= 80 {
				res1, res2 = getCellRangeAddr(data, pos+3, bv, reldelta, nil, nil)
				refx_raw, _ := unpack("<H", data[pos+1:pos+3])
				refx = int(refx_raw.(uint16))
				shx1, shx2 = getExternsheetLocalRange(bk, refx, blah)
			} else {
				res1, res2 = getCellRangeAddr(data, pos+15, bv, reldelta, nil, nil)
				result, _ := unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
				values := result.([]interface{})
				rawExtshtx := values[0].(int16)
				rawShx1 := values[1].(int16)
				rawShx2 := values[2].(int16)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tArea3d %d %d %d\n", rawExtshtx, rawShx1, rawShx2)
				}
				shx1, shx2 = getExternsheetLocalRangeB57(bk, int(rawExtshtx), int(rawShx1), int(rawShx2), blah)
			}
			anyErr = boolToInt(anyErr != 0 || shx1 < -1)
			rowx1, colx1, rowRel1, colRel1 := res1[0], res1[1], res1[2], res1[3]
			rowx2, colx2, rowRel2, colRel2 := res2[0], res2[1], res2[2], res2[3]
			isRel := rowRel1 != 0 || colRel1 != 0 || rowRel2 != 0 || colRel2 != 0
			anyRel = boolToInt(anyRel != 0 || isRel)
			coords := []int{shx1, shx2 + 1, rowx1, rowx2 + 1, colx1, colx2 + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			var ref3d *Ref3D
			if isRel {
				relflags := []int{0, 0, rowRel1, rowRel2, colRel1, colRel2}
				ref3d = NewRef3D(append(coords, relflags...)...)
				resOp.kind = oREL
				resOp.text = rangename3drel(bk, ref3d, 1)
			} else {
				ref3d = NewRef3D(coords...)
				resOp.kind = oREF
				resOp.text = rangename3d(bk, ref3d)
			}
			resOp._rank = LeafRank
			if optype == 1 {
				resOp.value = []*Ref3D{ref3d}
			}
			spush(resOp)
		} else if opcode == 0x19 { // tNameX
			dodgy := 0
			res := &Operand{kind: oUNK, value: nil}
			var refx, tgtnamex, origrefx int
			if bv >= 80 {
				result, _ := unpack("<HH", data[pos+1:pos+5])
				values := result.([]interface{})
				refx = int(values[0].(uint16))
				tgtnamex = int(values[1].(uint16)) - 1
				origrefx = refx
			} else {
				result, _ := unpack("<hxxxxxxxxH", data[pos+1:pos+13])
				values := result.([]interface{})
				refx := int(values[0].(int16))
				tgtnamex = int(values[1].(uint16)) - 1
				origrefx = refx
				if refx > 0 {
					refx -= 1
				} else if refx < 0 {
					refx = -refx - 1
				} else {
					dodgy = 1
				}
			}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   origrefx=%d refx=%d tgtnamex=%d dodgy=%d\n", origrefx, refx, tgtnamex, dodgy)
			}
			if tgtnamex == namex {
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "!!!! Self-referential !!!!\n")
				}
				dodgy = boolToInt(anyErr != 0 || true)
			}
			if dodgy == 0 {
				var shx1, _ int
				if bv >= 80 {
					shx1, _ = getExternsheetLocalRange(bk, refx, blah)
				} else if origrefx > 0 {
					shx1, _ = -4, -4 // external ref
				} else {
					exty := bk.externsheetTypeB57[refx]
					if exty == 4 { // non-specific sheet in own doc't
						shx1, _ = -1, -1 // internal, any sheet
					} else {
						shx1, _ = -666, -666
					}
				}
				if dodgy != 0 || shx1 < -1 {
					otext := fmt.Sprintf("<<Name #%d in external(?) file #%d>>", tgtnamex, origrefx)
					res = &Operand{kind: oUNK, value: nil, _rank: LeafRank, text: otext}
				} else {
					tgtobj := bk.NameObjList[tgtnamex]
					if !tgtobj.Evaluated {
						// recursive
						evaluateNameFormula(bk, tgtobj, tgtnamex, blah, level+1)
					}
					if tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0 {
						if blah != 0 {
							tgtobj.Dump(bk.logfile, "!!! bad tgtobj !!!", "------------------", 0)
						}
						res = &Operand{kind: oUNK, value: nil}
						anyErr = boolToInt(anyErr != 0 || tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0)
						anyRel = boolToInt(anyRel != 0 || tgtobj.AnyRel != 0)
					} else {
						assert(len(tgtobj.Stack) == 1)
						res = copyOperand(tgtobj.Stack[0])
					}
					res._rank = LeafRank
					if tgtobj.Scope == -1 {
						res.text = tgtobj.Name
					} else {
						res.text = bk.SheetNames()[tgtobj.Scope] + "!" + tgtobj.Name
					}
					if blah != 0 {
						fmt.Fprintf(bk.logfile, "    tNameX: setting text to %q\n", res.text)
					}
				}
			}
			spush(res)
		} else if _, ok := errorOpcodes[opcode]; ok {
			anyErr = 1
			spush(errorOpnd)
		} else {
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "FORMULA: /// Not handled yet: t%s\n", oname)
			}
			anyErr = 1
		}
		if sz <= 0 {
			panic("Fatal: token size is not positive")
		}
		pos += sz
	}

	if blah != 0 {
		fmt.Fprintf(bk.logfile, "End of formula. level=%d any_rel=%d any_err=%d stack=%v\n",
			level, anyRel, anyErr, stack)
		if len(stack) >= 2 {
			fmt.Fprintf(bk.logfile, "*** Stack has unprocessed args\n")
		}
		fmt.Fprintf(bk.logfile, "\n")
	}
	nobj.Stack = make([]*Operand, len(stack))
	for i, op := range stack {
		nobj.Stack[i] = op.(*Operand)
	}
	if len(stack) != 1 {
		nobj.Result = nil
	} else {
		nobj.Result = stack[0]
	}
	nobj.AnyRel = anyRel
	nobj.AnyErr = anyErr
	nobj.AnyExternal = anyExternal
	nobj.Evaluated = true
}

// decompileFormula decompiles a formula
func DecompileFormula(bk *Book, fmla []byte, fmlalen int, fmlatype int, browx, bcolx interface{}, blah int, level int, r1c1 int) string {
	if level > StackAlarmLevel {
		blah = 1
	}
	reldelta := 0
	if fmlatype&(FmlaTypeShared|FmlaTypeName|FmlaTypeCondFmt|FmlaTypeDataVal) != 0 {
		reldelta = 1
	}
	data := fmla
	bv := bk.BiffVersion
	if blah != 0 {
		fmt.Fprintf(bk.logfile, "::: decompile_formula len=%d fmlatype=%d browx=%v bcolx=%v reldelta=%d r1c1=%d level=%d\n",
			fmlalen, fmlatype, browx, bcolx, reldelta, r1c1, level)
		hexCharDump(data, 0, fmlalen, bk.logfile.(*os.File))
	}
	if level > StackPanicLevel {
		panic(fmt.Sprintf("Excessive indirect references in formula"))
	}
	sztab := szdict[bv]
	pos := 0
	stack := []interface{}{}
	anyRel := 0
	anyErr := 0
	unkOpnd := &Operand{kind: oUNK, value: nil}
	errorOpnd := &Operand{kind: oERR, value: nil}
	spush := func(item interface{}) {
		stack = append(stack, item)
	}

	doBinop := func(opcd int, stk []interface{}) {
		if len(stk) < 2 {
			return
		}
		bop := stk[len(stk)-1].(*Operand)
		aop := stk[len(stk)-2].(*Operand)
		rule := binopRules[opcd]
		_ = rule.argdict
		resultType := rule.resultType
		rank := rule.priority
		sym := rule.symbol

		otext := ""
		if aop._rank < rank {
			otext += "("
		}
		otext += aop.text
		if aop._rank < rank {
			otext += ")"
		}
		otext += sym
		if bop._rank < rank {
			otext += "("
		}
		otext += bop.text
		if bop._rank < rank {
			otext += ")"
		}

		resop := &Operand{kind: resultType, value: nil, _rank: rank, text: otext}
		stk = stk[:len(stk)-2]
		stk = append(stk, resop)
	}

	doUnaryop := func(opcode int, resultType int, stk []interface{}) {
		if len(stk) < 1 {
			return
		}
		aop := stk[len(stk)-1].(*Operand)
		rule := unopRules[opcode]
		rank := rule.priority
		sym1 := rule.prefix
		sym2 := rule.suffix

		otext := sym1
		if aop._rank < rank {
			otext += "("
		}
		otext += aop.text
		if aop._rank < rank {
			otext += ")"
		}
		otext += sym2

		res := &Operand{kind: resultType, value: nil, _rank: rank, text: otext}
		stk = stk[:len(stk)-1]
		stk = append(stk, res)
	}

	unexpectedOpcode := func(opArg int, onameArg string) {
		msg := fmt.Sprintf("ERROR *** Unexpected token 0x%02x (%s) found in formula type %s",
			opArg, onameArg, fmlaTypeDescrMap[fmlatype])
		fmt.Fprintf(bk.logfile, "%s\n", msg)
		// Note: Python version raises FormulaError but we just print and continue
	}

	if fmlalen == 0 {
		stack = []interface{}{unkOpnd}
	}

	for pos < fmlalen {
		op := int(data[pos])
		opcode := op & 0x1f
		optype := (op & 0x60) >> 5
		opx := opcode
		if optype != 0 {
			opx = opcode + 32
		}
		oname := onames[opx]
		sz := sztab[opx]

		if blah != 0 {
			fmt.Fprintf(bk.logfile, "Pos:%d Op:0x%02x opname:t%s Sz:%d opcode:%02xh optype:%02xh\n",
				pos, op, oname, sz, opcode, optype)
			fmt.Fprintf(bk.logfile, "Stack = %v\n", stack)
		}

		if sz == -2 {
			msg := fmt.Sprintf(`ERROR *** Unexpected token 0x%02x ("%s"); biff_version=%d`, op, oname, bv)
			panic(msg)
		}

		tokenMask := tokenNotAllowed[opx]
		if tokenMask {
			unexpectedOpcode(op, oname)
		}

		if optype == 0 {
			if opcode <= 0x01 { // tExp
				var fmtStr string
				if bv >= 30 {
					fmtStr = "<x2H"
				} else {
					fmtStr = "<xHB"
				}
				assert(pos == 0 && fmlalen == sz && len(stack) == 0)
				result, _ := unpack(fmtStr, data)
				values := result.([]interface{})
				rowx := int(values[0].(uint16))
				colx := int(values[1].(uint16))
				text := fmt.Sprintf("SHARED FMLA at rowx=%d colx=%d", rowx, colx)
				spush(&Operand{kind: oUNK, value: nil, _rank: LeafRank, text: text})
				if fmlatype&(FmlaTypeCell|FmlaTypeArray) == 0 {
					unexpectedOpcode(op, oname)
				}
			} else if 0x03 <= opcode && opcode <= 0x0E { // Add, Sub, Mul, Div, Power, tConcat, tLT, ..., tNE
				doBinop(opcode, stack)
			} else if opcode == 0x0F { // tIsect
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tIsect pre %v\n", stack)
				}
				if len(stack) < 2 {
					continue
				}
				bop := stack[len(stack)-1].(*Operand)
				aop := stack[len(stack)-2].(*Operand)
				sym := " "
				rank := 80
				otext := ""
				if aop._rank < rank {
					otext += "("
				}
				otext += aop.text
				if aop._rank < rank {
					otext += ")"
				}
				otext += sym
				if bop._rank < rank {
					otext += "("
				}
				otext += bop.text
				if bop._rank < rank {
					otext += ")"
				}

				res := &Operand{kind: oREF}
				res.text = otext
				if bop.kind == oERR || aop.kind == oERR {
					res.kind = oERR
				} else if bop.kind == oUNK || aop.kind == oUNK {
					// This can happen with undefined labels
				} else if bop.kind == oREF && aop.kind == oREF {
					// pass
				} else if bop.kind == oREL && aop.kind == oREL {
					res.kind = oREL
				}
				stack = stack[:len(stack)-2]
				spush(res)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tIsect post %v\n", stack)
				}
			} else if opcode == 0x10 { // tList
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tList pre %v\n", stack)
				}
				if len(stack) < 2 {
					continue
				}
				bop := stack[len(stack)-1].(*Operand)
				aop := stack[len(stack)-2].(*Operand)
				sym := ","
				rank := 80
				otext := ""
				if aop._rank < rank {
					otext += "("
				}
				otext += aop.text
				if aop._rank < rank {
					otext += ")"
				}
				otext += sym
				if bop._rank < rank {
					otext += "("
				}
				otext += bop.text
				if bop._rank < rank {
					otext += ")"
				}

				res := &Operand{kind: oREF, value: nil, _rank: rank, text: otext}
				if bop.kind == oERR || aop.kind == oERR {
					res.kind = oERR
				} else if (bop.kind == oREF || bop.kind == oREL) && (aop.kind == oREF || aop.kind == oREL) {
					res.kind = oREF
					if aop.kind == oREL || bop.kind == oREL {
						res.kind = oREL
					}
				}
				stack = stack[:len(stack)-2]
				spush(res)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tList post %v\n", stack)
				}
			} else if opcode == 0x11 { // tRange
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRange pre %v\n", stack)
				}
				if len(stack) < 2 {
					continue
				}
				bop := stack[len(stack)-1].(*Operand)
				aop := stack[len(stack)-2].(*Operand)
				sym := ":"
				rank := 80
				otext := ""
				if aop._rank < rank {
					otext += "("
				}
				otext += aop.text
				if aop._rank < rank {
					otext += ")"
				}
				otext += sym
				if bop._rank < rank {
					otext += "("
				}
				otext += bop.text
				if bop._rank < rank {
					otext += ")"
				}

				res := &Operand{kind: oREF, value: nil, _rank: rank, text: otext}
				if bop.kind == oERR || aop.kind == oERR {
					res.kind = oERR
				} else if bop.kind == oREF && aop.kind == oREF {
					// pass
				}
				stack = stack[:len(stack)-2]
				spush(res)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRange post %v\n", stack)
				}
			} else if 0x12 <= opcode && opcode <= 0x14 { // tUplus, tUminus, tPercent
				doUnaryop(opcode, oNUM, stack)
			} else if opcode == 0x15 { // tParen
				// source cosmetics
			} else if opcode == 0x16 { // tMissArg
				spush(&Operand{kind: oMSNG, value: nil, _rank: LeafRank, text: ""})
			} else if opcode == 0x17 { // tStr
				var strg string
				var newpos int
				if bv <= 70 {
					strg, newpos = unpackStringUpdatePos(data, pos+1, bk.Encoding, 1)
				} else {
					strg, newpos = unpackUnicodeUpdatePos(data, pos+1, 1)
				}
				sz = newpos - pos
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "   sz=%d strg=%q\n", sz, strg)
				}
				text := "\"" + strings.ReplaceAll(strg, "\"", "\"\"") + "\""
				spush(&Operand{kind: oSTRG, value: nil, _rank: LeafRank, text: text})
			} else if opcode == 0x18 { // tExtended
				// new with BIFF 8
				assert(bv >= 80)
				panic("tExtended token not implemented")
			} else if opcode == 0x19 { // tAttr
				result, err := unpack("<BH", data[pos+1:pos+4])
				if err != nil {
					panic(err)
				}
				values := result.([]interface{})
				subop := values[0].(uint8)
				nc := values[1].(uint16)
				subname := tAttrNames[int(subop)]
				if subop == 0x04 { // Choose
					sz = int(nc)*2 + 6
				} else if subop == 0x10 { // Sum (single arg)
					sz = 4
					if blah != 0 {
						fmt.Fprintf(bk.logfile, "tAttrSum %v\n", stack)
					}
					if len(stack) >= 1 {
						aop := stack[len(stack)-1].(*Operand)
						otext := "SUM(" + aop.text + ")"
						stack[len(stack)-1] = &Operand{kind: oNUM, value: nil, _rank: FuncRank, text: otext}
					}
				} else {
					sz = 4
				}
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "   subop=%02xh subname=t%s sz=%d nc=%02xh\n", subop, subname, sz, nc)
				}
			} else if 0x1A <= opcode && opcode <= 0x1B { // tSheet, tEndSheet
				assert(bv < 50)
				panic("tSheet & tEndsheet tokens not implemented")
			} else if 0x1C <= opcode && opcode <= 0x1F { // tErr, tBool, tInt, tNum
				inx := opcode - 0x1C
				nb := []int{1, 1, 2, 8}[inx]
				kind := []int{oERR, oBOOL, oNUM, oNUM}[inx]
				value, _ := unpack([]string{"<B", "<B", "<H", "<d"}[inx], data[pos+1:pos+1+nb])
				var text string
				if inx == 2 { // tInt
					value = float64(value.(int16))
					text = fmt.Sprintf("%v", value)
				} else if inx == 3 { // tNum
					text = fmt.Sprintf("%v", value)
				} else if inx == 1 { // tBool
					if value.(uint8) != 0 {
						text = "TRUE"
					} else {
						text = "FALSE"
					}
				} else {
					text = "\"" + errorTextFromCode[int(value.(uint8))] + "\""
				}
				spush(&Operand{kind: kind, value: value, _rank: LeafRank, text: text})
			} else {
				panic(fmt.Sprintf("Unhandled opcode: 0x%02x", opcode))
			}
			if sz <= 0 {
				panic(fmt.Sprintf("Size not set for opcode 0x%02x", opcode))
			}
			pos += sz
			continue
		}

		// optype != 0 (variable-length tokens)
		if opcode == 0x00 { // tArray
			spush(unkOpnd)
		} else if opcode == 0x01 { // tFunc
			nb := 1
			if bv >= 40 {
				nb = 2
			}
			fmtStr := "<"
			if nb == 2 {
				fmtStr += "B"
			}
			fmtStr += "H"
			funcx, _ := unpack(fmtStr, data[pos+1:pos+1+nb])
			funcAttrs, ok := funcDefs[funcx.(int)]
			if !ok {
				fmt.Fprintf(bk.logfile, "*** formula/tFunc unknown FuncID:%d\n", funcx)
				spush(unkOpnd)
			} else {
				funcName := funcAttrs.name
				nargs := funcAttrs.minArgs
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "    FuncID=%d name=%s nargs=%d\n", funcx, funcName, nargs)
				}
				assert(len(stack) >= nargs)
				var otext string
				if nargs > 0 {
					argtext := make([]string, nargs)
					for i := 0; i < nargs; i++ {
						argtext[i] = stack[len(stack)-nargs+i].(*Operand).text
					}
					otext = funcName + "(" + strings.Join(argtext, listsep) + ")"
					for i := 0; i < nargs; i++ {
						stack = stack[:len(stack)-1]
					}
				} else {
					otext = funcName + "()"
				}
				res := &Operand{kind: oUNK, value: nil, _rank: FuncRank, text: otext}
				spush(res)
			}
		} else if opcode == 0x02 { // tFuncVar
			nb := 1
			if bv >= 40 {
				nb = 2
			}
			fmtStr := "<B"
			if nb == 2 {
				fmtStr += "B"
			}
			fmtStr += "H"
			nargs_funcx, _ := unpack(fmtStr, data[pos+1:pos+2+nb])
			nargs := nargs_funcx.([]interface{})[0].(uint8)
			funcx := nargs_funcx.([]interface{})[1].(uint16)
			prompt := nargs >> 7
			nargs &= 0x7F
			macro := funcx >> 15
			funcx &= 0x7FFF
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   FuncID=%d nargs=%d macro=%d prompt=%d\n", funcx, nargs, macro, prompt)
			}
			funcAttrs, ok := funcDefs[int(funcx)]
			if !ok {
				fmt.Fprintf(bk.logfile, "*** formula/tFuncVar unknown FuncID:%d\n", funcx)
				spush(unkOpnd)
			} else {
				funcName := funcAttrs.name
				minargs := funcAttrs.minArgs
				maxargs := funcAttrs.maxArgs
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "    name: %s, min~max args: %d~%d\n", funcName, minargs, maxargs)
				}
				assert(minargs <= int(nargs) && int(nargs) <= maxargs)
				assert(len(stack) >= int(nargs))
				assert(len(stack) >= int(nargs))
				argtext := make([]string, nargs)
				for i := 0; i < int(nargs); i++ {
					argtext[i] = stack[len(stack)-int(nargs)+i].(*Operand).text
				}
				otext := funcName + "(" + strings.Join(argtext, listsep) + ")"
				res := &Operand{kind: oUNK, value: nil, _rank: FuncRank, text: otext}
				if funcx == 1 { // IF
					testarg := stack[len(stack)-int(nargs)].(*Operand)
					if testarg.kind != oNUM && testarg.kind != oBOOL {
						if blah != 0 && testarg.kind != oUNK {
							fmt.Fprintf(bk.logfile, "IF testarg kind?\n")
						}
					} else if testarg.value != 0 && testarg.value != 1 {
						if blah != 0 && testarg.value != nil {
							fmt.Fprintf(bk.logfile, "IF testarg value?\n")
						}
					} else {
						if int(nargs) == 2 && testarg.value == 0 {
							// IF(FALSE, tv) => FALSE
							res.kind = oBOOL
							res.value = 0
						} else {
							respos := len(stack) - int(nargs) + 2 - int(testarg.value.(int))
							chosen := stack[respos].(*Operand)
							if chosen.kind == oMSNG {
								res.kind = oNUM
								res.value = 0
							} else {
								res.kind = chosen.kind
								res.value = chosen.value
							}
							if blah != 0 {
								fmt.Fprintf(bk.logfile, "$$$$$$ IF => constant\n")
							}
						}
					}
				} else if funcx == 100 { // CHOOSE
					testarg := stack[len(stack)-int(nargs)].(*Operand)
					if testarg.kind == oNUM {
						if 1 <= testarg.value.(int) && testarg.value.(int) < int(nargs) {
							chosen := stack[len(stack)-int(nargs)+testarg.value.(int)].(*Operand)
							if chosen.kind == oMSNG {
								res.kind = oNUM
								res.value = 0
							} else {
								res.kind = chosen.kind
								res.value = chosen.value
							}
						}
					}
				}
				for i := 0; i < int(nargs); i++ {
					stack = stack[:len(stack)-1]
				}
				spush(res)
			}
		} else if opcode == 0x03 { // tName
			tgtnamex_raw, _ := unpack("<H", data[pos+1:pos+3])
			tgtnamex := int(tgtnamex_raw.(uint16)) - 1
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   tgtnamex=%d\n", tgtnamex)
			}
			tgtobj := bk.NameObjList[tgtnamex]
			if !tgtobj.Evaluated {
				// recursive
				evaluateNameFormula(bk, tgtobj, tgtnamex, blah, level+1)
			}
			var res *Operand
			if tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0 {
				if blah != 0 {
					tgtobj.Dump(bk.logfile, "!!! tgtobj has problems!!!", "-----------       --------", 0)
				}
				res = &Operand{kind: oUNK, value: nil}
				anyErr = boolToInt(anyErr != 0 || tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0)
				anyRel = boolToInt(anyRel != 0 || tgtobj.AnyRel != 0)
			} else {
				assert(len(tgtobj.Stack) == 1)
				res = copyOperand(tgtobj.Stack[0])
			}
			res._rank = LeafRank
			if tgtobj.Scope == -1 {
				res.text = tgtobj.Name
			} else {
				res.text = bk.SheetNames()[tgtobj.Scope] + "!" + tgtobj.Name
			}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "    tName: setting text to %q\n", res.text)
			}
			spush(res)
		} else if opcode == 0x04 { // tRef
			rowx, colx, rowRel, colRel := getCellAddr(data, pos+1, bv, reldelta, browx, bcolx)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "  (%d, %d, %d, %d)\n", rowx, colx, rowRel, colRel)
			}
			shx1 := 0
			shx2 := 0 // N.B. relative to the CURRENT SHEET #######
			anyRel = boolToInt(anyRel != 0 || rowRel != 0 || colRel != 0)
			coords := []int{shx1, shx2 + 1, rowx, rowx + 1, colx, colx + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			var ref3d *Ref3D
			if optype == 1 {
				relflags := []int{1, 1, rowRel, rowRel, colRel, colRel}
				ref3d = NewRef3D(append(coords, relflags...)...)
				resOp.kind = oREL
				resOp.text = rangename3drel(bk, ref3d, r1c1)
			} else {
				ref3d = NewRef3D(coords...)
				resOp.kind = oREF
				resOp.text = rangename3d(bk, ref3d)
			}
			resOp._rank = LeafRank
			if optype == 1 {
				resOp.value = []*Ref3D{ref3d}
			}
			spush(resOp)
		} else if opcode == 0x05 { // tArea
			res1, res2 := getCellRangeAddr(data, pos+1, bv, reldelta, browx, bcolx)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "  %v %v\n", res1, res2)
			}
			rowx1, colx1, rowRel1, colRel1 := res1[0], res1[1], res1[2], res1[3]
			rowx2, colx2, rowRel2, colRel2 := res2[0], res2[1], res2[2], res2[3]
			shx1 := 0
			shx2 := 0 // N.B. relative to the CURRENT SHEET #######
			anyRel = boolToInt(anyRel != 0 || rowRel1 != 0 || colRel1 != 0 || rowRel2 != 0 || colRel2 != 0)
			coords := []int{shx1, shx2 + 1, rowx1, rowx2 + 1, colx1, colx2 + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			var ref3d *Ref3D
			if optype == 1 {
				relflags := []int{1, 1, rowRel1, rowRel2, colRel1, colRel2}
				ref3d = NewRef3D(append(coords, relflags...)...)
				resOp.kind = oREL
				resOp.text = rangename3drel(bk, ref3d, r1c1)
			} else {
				ref3d = NewRef3D(coords...)
				resOp.kind = oREF
				resOp.text = rangename3d(bk, ref3d)
			}
			resOp._rank = LeafRank
			if optype == 1 {
				resOp.value = []*Ref3D{ref3d}
			}
			spush(resOp)
		} else if opcode == 0x06 { // tMemArea
			panic("tMemArea not implemented")
		} else if opcode == 0x09 { // tMemFunc
			nb, _ := unpack("<H", data[pos+1:pos+3])
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "  %d bytes of cell ref formula\n", nb)
			}
			// no effect on stack
		} else if opcode == 0x0C { // tRefN
			panic("tRefN not implemented")
		} else if opcode == 0x0D { // tAreaN
			panic("tAreaN not implemented")
		} else if opcode == 0x1A { // tRef3d
			var refx int
			var shx1, shx2 int
			var rowx, colx, rowRel, colRel int
			if bv >= 80 {
				rowx, colx, rowRel, colRel = getCellAddr(data, pos+3, bv, reldelta, browx, bcolx)
				refx_raw, _ := unpack("<H", data[pos+1:pos+3])
				refx = int(refx_raw.(uint16))
				shx1, shx2 = getExternsheetLocalRange(bk, refx, blah)
			} else {
				rowx, colx, rowRel, colRel = getCellAddr(data, pos+15, bv, reldelta, browx, bcolx)
				result, _ := unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
				values := result.([]interface{})
				rawExtshtx := values[0].(int16)
				rawShx1 := values[1].(int16)
				rawShx2 := values[2].(int16)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tRef3d %d %d %d\n", rawExtshtx, rawShx1, rawShx2)
				}
				shx1, shx2 = getExternsheetLocalRangeB57(bk, int(rawExtshtx), int(rawShx1), int(rawShx2), blah)
			}
			isRel := rowRel != 0 || colRel != 0
			anyRel = boolToInt(anyRel != 0 || isRel)
			coords := []int{shx1, shx2 + 1, rowx, rowx + 1, colx, colx + 1}
			anyErr = boolToInt(anyErr != 0 || shx1 < -1)
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			var ref3d *Ref3D
			if isRel {
				relflags := []int{0, 0, rowRel, rowRel, colRel, colRel}
				ref3d = NewRef3D(append(coords, relflags...)...)
				resOp.kind = oREL
				resOp.text = rangename3drel(bk, ref3d, r1c1)
			} else {
				ref3d = NewRef3D(coords...)
				resOp.kind = oREF
				resOp.text = rangename3d(bk, ref3d)
			}
			resOp._rank = LeafRank
			if optype == 1 {
				resOp.value = []*Ref3D{ref3d}
			}
			spush(resOp)
		} else if opcode == 0x1B { // tArea3d
			var res1, res2 [4]int
			var refx int
			var shx1, shx2 int
			if bv >= 80 {
				res1, res2 = getCellRangeAddr(data, pos+3, bv, reldelta, browx, bcolx)
				refx_raw, _ := unpack("<H", data[pos+1:pos+3])
				refx = int(refx_raw.(uint16))
				shx1, shx2 = getExternsheetLocalRange(bk, refx, blah)
			} else {
				res1, res2 = getCellRangeAddr(data, pos+15, bv, reldelta, browx, bcolx)
				result, _ := unpack("<hxxxxxxxxhh", data[pos+1:pos+15])
				values := result.([]interface{})
				rawExtshtx := values[0].(int16)
				rawShx1 := values[1].(int16)
				rawShx2 := values[2].(int16)
				if blah != 0 {
					fmt.Fprintf(bk.logfile, "tArea3d %d %d %d\n", rawExtshtx, rawShx1, rawShx2)
				}
				shx1, shx2 = getExternsheetLocalRangeB57(bk, int(rawExtshtx), int(rawShx1), int(rawShx2), blah)
			}
			anyErr = boolToInt(anyErr != 0 || shx1 < -1)
			rowx1, colx1, rowRel1, colRel1 := res1[0], res1[1], res1[2], res1[3]
			rowx2, colx2, rowRel2, colRel2 := res2[0], res2[1], res2[2], res2[3]
			isRel := rowRel1 != 0 || colRel1 != 0 || rowRel2 != 0 || colRel2 != 0
			anyRel = boolToInt(anyRel != 0 || isRel)
			coords := []int{shx1, shx2 + 1, rowx1, rowx2 + 1, colx1, colx2 + 1}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   %v\n", coords)
			}
			resOp := &Operand{kind: oUNK, value: nil}
			var ref3d *Ref3D
			if isRel {
				relflags := []int{0, 0, rowRel1, rowRel2, colRel1, colRel2}
				ref3d = NewRef3D(append(coords, relflags...)...)
				resOp.kind = oREL
				resOp.text = rangename3drel(bk, ref3d, r1c1)
			} else {
				ref3d = NewRef3D(coords...)
				resOp.kind = oREF
				resOp.text = rangename3d(bk, ref3d)
			}
			resOp._rank = LeafRank
			if optype == 1 {
				resOp.value = []*Ref3D{ref3d}
			}
			spush(resOp)
		} else if opcode == 0x19 { // tNameX
			dodgy := 0
			res := &Operand{kind: oUNK, value: nil}
			var refx, tgtnamex, origrefx int
			if bv >= 80 {
				result, _ := unpack("<HH", data[pos+1:pos+5])
				values := result.([]interface{})
				refx = int(values[0].(uint16))
				tgtnamex = int(values[1].(uint16)) - 1
				origrefx = refx
			} else {
				result, _ := unpack("<hxxxxxxxxH", data[pos+1:pos+13])
				values := result.([]interface{})
				refx := int(values[0].(int16))
				tgtnamex = int(values[1].(uint16)) - 1
				origrefx = refx
				if refx > 0 {
					refx -= 1
				} else if refx < 0 {
					refx = -refx - 1
				} else {
					dodgy = 1
				}
			}
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "   origrefx=%d refx=%d tgtnamex=%d dodgy=%d\n", origrefx, refx, tgtnamex, dodgy)
			}
			if tgtnamex == -1 { // special marker for current formula name
				dodgy = boolToInt(anyErr != 0 || true)
			}
			if dodgy == 0 {
				var shx1, _ int
				if bv >= 80 {
					shx1, _ = getExternsheetLocalRange(bk, refx, blah)
				} else if origrefx > 0 {
					shx1, _ = -4, -4 // external ref
				} else {
					exty := bk.externsheetTypeB57[refx]
					if exty == 4 { // non-specific sheet in own doc't
						shx1, _ = -1, -1 // internal, any sheet
					} else {
						shx1, _ = -666, -666
					}
				}
				if dodgy != 0 || shx1 < -1 {
					otext := fmt.Sprintf("<<Name #%d in external(?) file #%d>>", tgtnamex, origrefx)
					res = &Operand{kind: oUNK, value: nil, _rank: LeafRank, text: otext}
				} else {
					tgtobj := bk.NameObjList[tgtnamex]
					if !tgtobj.Evaluated {
						// recursive
						evaluateNameFormula(bk, tgtobj, tgtnamex, blah, level+1)
					}
					if tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0 {
						if blah != 0 {
							tgtobj.Dump(bk.logfile, "!!! bad tgtobj !!!", "------------------", 0)
						}
						res = &Operand{kind: oUNK, value: nil}
						anyErr = boolToInt(anyErr != 0 || tgtobj.Macro != 0 || tgtobj.Binary != 0 || tgtobj.AnyErr != 0)
						anyRel = boolToInt(anyRel != 0 || tgtobj.AnyRel != 0)
					} else {
						assert(len(tgtobj.Stack) == 1)
						res = copyOperand(tgtobj.Stack[0])
					}
					res._rank = LeafRank
					if tgtobj.Scope == -1 {
						res.text = tgtobj.Name
					} else {
						res.text = bk.SheetNames()[tgtobj.Scope] + "!" + tgtobj.Name
					}
					if blah != 0 {
						fmt.Fprintf(bk.logfile, "    tNameX: setting text to %q\n", res.text)
					}
				}
			}
			spush(res)
		} else if _, ok := errorOpcodes[opcode]; ok {
			anyErr = 1
			spush(errorOpnd)
		} else {
			if blah != 0 {
				fmt.Fprintf(bk.logfile, "FORMULA: /// Not handled yet: t%s\n", oname)
			}
			anyErr = 1
		}
		if sz <= 0 {
			panic("Fatal: token size is not positive")
		}
		pos += sz
	}

	if blah != 0 {
		fmt.Fprintf(bk.logfile, "End of formula. level=%d any_rel=%d any_err=%d stack=%v\n",
			level, anyRel, anyErr, stack)
		if len(stack) >= 2 {
			fmt.Fprintf(bk.logfile, "*** Stack has unprocessed args\n")
		}
		fmt.Fprintf(bk.logfile, "\n")
	}

	if len(stack) == 1 {
		result := stack[0].(*Operand)
		return result.text
	}
	return "<<Stack underflow>>"
}
