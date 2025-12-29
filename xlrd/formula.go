package xlrd

import (
	"encoding/binary"
	"fmt"
	"math"
	"strings"
)

// FormulaError represents an error in formula parsing.
type FormulaError struct {
	Message string
}

func (e *FormulaError) Error() string {
	return e.Message
}

// Formula type constants
const (
	FMLA_TYPE_CELL     = 1
	FMLA_TYPE_SHARED   = 2
	FMLA_TYPE_ARRAY    = 4
	FMLA_TYPE_COND_FMT = 8
	FMLA_TYPE_DATA_VAL = 16
	FMLA_TYPE_NAME     = 32
	ALL_FMLA_TYPES     = 63
)

// Operand type constants
const (
	oBOOL = 3
	oERR  = 4
	oNUM  = 2
	oREF  = -1
	oREL  = -2
	oSTRG = 1
	oUNK  = 0
	oMSNG = 5 // tMissArg
)

// OkindDict maps operand types to their string representations.
var OkindDict = map[int]string{
	-2: "oREL",
	-1: "oREF",
	0:  "oUNK",
	1:  "oSTRG",
	2:  "oNUM",
	3:  "oBOOL",
	4:  "oERR",
	5:  "oMSNG",
}

// FmlaTypeDescrMap maps formula types to their string descriptions.
var FmlaTypeDescrMap = map[int]string{
	1:  "CELL",
	2:  "SHARED",
	4:  "ARRAY",
	8:  "COND-FMT",
	16: "DATA-VAL",
	32: "NAME",
}

// TokenNotAllowed is a function that returns forbidden formula types for a token.
var TokenNotAllowed = map[int]int{
	0x01: ALL_FMLA_TYPES - FMLA_TYPE_CELL,                            // tExp
	0x02: ALL_FMLA_TYPES - FMLA_TYPE_CELL,                            // tTbl
	0x0F: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tIsect
	0x10: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tUnion/List
	0x11: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tRange
	0x20: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tArray
	0x23: FMLA_TYPE_SHARED,                                           // tName
	0x39: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tNameX
	0x3A: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tRef3d
	0x3B: FMLA_TYPE_SHARED + FMLA_TYPE_COND_FMT + FMLA_TYPE_DATA_VAL, // tArea3d
	0x2C: FMLA_TYPE_CELL + FMLA_TYPE_ARRAY,                           // tRefN
	0x2D: FMLA_TYPE_CELL + FMLA_TYPE_ARRAY,                           // tAreaN
}

// Formula evaluation constants
const (
	LEAF_RANK         = 90
	FUNC_RANK         = 90
	STACK_ALARM_LEVEL = 5
	STACK_PANIC_LEVEL = 10
)

// Operand represents an operand used in evaluating formulas.
// The kind field determines how the value is represented:
//   - oBOOL (3): integer where 0 => False, 1 => True
//   - oERR (4): None or an int error code
//   - oMSNG (5): placeholder for missing function argument, value is nil
//   - oNUM (2): a float value
//   - oREF (-1): nil or a non-empty slice of absolute Ref3D instances
//   - oREL (-2): nil or a non-empty slice of relative Ref3D instances
//   - oSTRG (1): a string value
//   - oUNK (0): unknown or ambiguous kind, value is nil
type Operand struct {
	// Value represents the actual value of the operand.
	// nil means the value is a variable (depends on cell data), not a constant.
	Value interface{}

	// Kind indicates the type of operand (oUNK means unknown/ambiguous).
	Kind int

	// Rank is an internal operator precedence value used in reconstructing formula text.
	Rank int

	// Text is the reconstituted text of the original formula.
	Text string
}

// NewOperand creates a new Operand with the specified parameters.
func NewOperand(akind int, avalue interface{}, arank int, atext string) *Operand {
	if atext == "" {
		atext = "?"
	}
	return &Operand{
		Kind:  akind,
		Value: avalue,
		Rank:  arank,
		Text:  atext,
	}
}

// String returns a string representation of the Operand.
func (o *Operand) String() string {
	kindText := OkindDict[o.Kind]
	if kindText == "" {
		kindText = "?Unknown kind?"
	}
	return fmt.Sprintf("Operand(kind=%s, value=%v, text=%s)", kindText, o.Value, o.Text)
}

// Ref3D represents an absolute or relative 3-dimensional reference to a box of one or more cells.
// The coords field is a slice of the form: [shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi]
// where 0 <= thingxlo <= thingx < thingxhi.
//
// The relflags field is a slice of 6 flags indicating whether the corresponding
// (sheet|row|col)(lo|hi) is relative (1) or absolute (0).
//
// Individual coordinates are also available as separate fields for convenience.
type Ref3D struct {
	// Coords contains [shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi]
	Coords [6]int

	// RelFlags contains 6 flags for relative/absolute components
	RelFlags [6]int

	// Individual coordinate fields for convenience
	ShtXLo, ShtXHi int
	RowXLo, RowXHi int
	ColXLo, ColXHi int
}

// NewRef3D creates a new Ref3D from a slice of integers.
// The slice should contain at least 6 elements for coords, and optionally 6 more for relflags.
func NewRef3D(atuple []int) *Ref3D {
	ref3d := &Ref3D{}

	// Set coords (first 6 elements)
	copy(ref3d.Coords[:], atuple[:6])

	// Set relflags (next 6 elements, or default to absolute)
	if len(atuple) >= 12 {
		copy(ref3d.RelFlags[:], atuple[6:12])
	} else {
		// Default to absolute (all zeros)
		ref3d.RelFlags = [6]int{0, 0, 0, 0, 0, 0}
	}

	// Set individual fields for convenience
	ref3d.ShtXLo = ref3d.Coords[0]
	ref3d.ShtXHi = ref3d.Coords[1]
	ref3d.RowXLo = ref3d.Coords[2]
	ref3d.RowXHi = ref3d.Coords[3]
	ref3d.ColXLo = ref3d.Coords[4]
	ref3d.ColXHi = ref3d.Coords[5]

	return ref3d
}

// String returns a string representation of the Ref3D.
func (r *Ref3D) String() string {
	if r.RelFlags == [6]int{0, 0, 0, 0, 0, 0} {
		return fmt.Sprintf("Ref3D(coords=%v)", r.Coords)
	}
	return fmt.Sprintf("Ref3D(coords=%v, relflags=%v)", r.Coords, r.RelFlags)
}

// RowNameRel returns a relative row name.
// If no base rowx is provided, returns R1C1 format.
func RowNameRel(rowx, rowxrel int, browx *int, r1c1 bool) string {
	// If no base rowx is provided, we have to return r1c1
	if browx == nil {
		r1c1 = true
	}
	if rowxrel == 0 {
		if r1c1 {
			return fmt.Sprintf("R%d", rowx+1)
		}
		return fmt.Sprintf("$%d", rowx+1)
	}
	if r1c1 {
		if rowx != 0 {
			return fmt.Sprintf("R[%d]", rowx)
		}
		return "R"
	}
	return fmt.Sprintf("%d", (*browx+rowx)%65536+1)
}

// ColNameRel returns a relative column name.
// If no base colx is provided, returns R1C1 format.
func ColNameRel(colx, colxrel int, bcolx *int, r1c1 bool) string {
	// If no base colx is provided, we have to return r1c1
	if bcolx == nil {
		r1c1 = true
	}
	if colxrel == 0 {
		if r1c1 {
			return fmt.Sprintf("C%d", colx+1)
		}
		return "$" + colname(colx)
	}
	if r1c1 {
		if colx != 0 {
			return fmt.Sprintf("C[%d]", colx)
		}
		return "C"
	}
	return colname((*bcolx + colx) % 256)
}

// CellNameRel returns a relative cell name.
func CellNameRel(rowx, colx, rowxrel, colxrel int, browx, bcolx *int, r1c1 bool) string {
	if rowxrel == 0 && colxrel == 0 {
		return CellNameAbs(rowx, colx, r1c1)
	}
	if (rowxrel != 0 && browx == nil) || (colxrel != 0 && bcolx == nil) {
		// must flip the whole cell into R1C1 mode
		r1c1 = true
	}
	c := ColNameRel(colx, colxrel, bcolx, r1c1)
	r := RowNameRel(rowx, rowxrel, browx, r1c1)
	if r1c1 {
		return r + c
	}
	return c + r
}

// RangeName2DRel returns a relative 2D range name.
func RangeName2DRel(rloRhiCloChi []int, relFlags []int, browx, bcolx *int, r1c1 bool) string {
	rlo, rhi, clo, chi := rloRhiCloChi[0], rloRhiCloChi[1], rloRhiCloChi[2], rloRhiCloChi[3]
	rlorel, rhirel, clorel, chirel := relFlags[0], relFlags[1], relFlags[2], relFlags[3]

	topleft := CellNameRel(rlo, clo, rlorel, clorel, browx, bcolx, r1c1)
	botright := CellNameRel(rhi-1, chi-1, rhirel, chirel, browx, bcolx, r1c1)
	return fmt.Sprintf("%s:%s", topleft, botright)
}

// Binary operation token constants
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

// Nop returns the input value unchanged.
func Nop(x interface{}) interface{} {
	return x
}

// OprPow returns x raised to the power of y.
func OprPow(x, y float64) float64 {
	return math.Pow(x, y)
}

// OprLt returns true if x < y.
func OprLt(x, y interface{}) bool {
	switch xv := x.(type) {
	case float64:
		switch yv := y.(type) {
		case float64:
			return xv < yv
		case string:
			return false // In Excel, numbers are "less than" strings
		}
	case string:
		switch yv := y.(type) {
		case float64:
			return true // In Excel, strings are "greater than" numbers
		case string:
			return xv < yv
		}
	}
	return false
}

// OprLe returns true if x <= y.
func OprLe(x, y interface{}) bool {
	switch xv := x.(type) {
	case float64:
		switch yv := y.(type) {
		case float64:
			return xv <= yv
		case string:
			return false
		}
	case string:
		switch yv := y.(type) {
		case float64:
			return true
		case string:
			return xv <= yv
		}
	}
	return false
}

// OprEq returns true if x == y.
func OprEq(x, y interface{}) bool {
	switch xv := x.(type) {
	case float64:
		switch yv := y.(type) {
		case float64:
			return xv == yv
		case string:
			return false
		}
	case string:
		switch yv := y.(type) {
		case float64:
			return false
		case string:
			return xv == yv
		}
	}
	return false
}

// OprGe returns true if x >= y.
func OprGe(x, y interface{}) bool {
	switch xv := x.(type) {
	case float64:
		switch yv := y.(type) {
		case float64:
			return xv >= yv
		case string:
			return true
		}
	case string:
		switch yv := y.(type) {
		case float64:
			return false
		case string:
			return xv >= yv
		}
	}
	return false
}

// OprGt returns true if x > y.
func OprGt(x, y interface{}) bool {
	switch xv := x.(type) {
	case float64:
		switch yv := y.(type) {
		case float64:
			return xv > yv
		case string:
			return true
		}
	case string:
		switch yv := y.(type) {
		case float64:
			return false
		case string:
			return xv > yv
		}
	}
	return false
}

// OprNe returns true if x != y.
func OprNe(x, y interface{}) bool {
	return !OprEq(x, y)
}

// Num2Str converts a number to string, emulating Excel's default conversion.
func Num2Str(num float64) string {
	s := fmt.Sprintf("%g", num)
	if strings.HasSuffix(s, ".0") {
		s = s[:len(s)-2]
	}
	return s
}

// nop returns the input value unchanged.
func nop(x interface{}) interface{} {
	return x
}

// AdjustCellAddrBiff8 adjusts cell address for BIFF8 format.
func AdjustCellAddrBiff8(rowval, colval int, reldelta bool, browx, bcolx *int) (int, int, int, int) {
	rowRel := (colval >> 15) & 1
	colRel := (colval >> 14) & 1
	rowx := rowval
	colx := colval & 0xff

	if reldelta {
		if rowRel != 0 && rowx >= 32768 {
			rowx -= 65536
		}
		if colRel != 0 && colx >= 128 {
			colx -= 256
		}
	} else {
		if rowRel != 0 && browx != nil {
			rowx -= *browx
		}
		if colRel != 0 && bcolx != nil {
			colx -= *bcolx
		}
	}
	return rowx, colx, rowRel, colRel
}

// AdjustCellAddrBiffLe7 adjusts cell address for BIFF <= 7 format.
func AdjustCellAddrBiffLe7(rowval, colval int, reldelta bool, browx, bcolx *int) (int, int, int, int) {
	rowRel := (rowval >> 15) & 1
	colRel := (rowval >> 14) & 1
	rowx := rowval & 0x3fff
	colx := colval

	if reldelta {
		if rowRel != 0 && rowx >= 8192 {
			rowx -= 16384
		}
		if colRel != 0 && colx >= 128 {
			colx -= 256
		}
	} else {
		if rowRel != 0 && browx != nil {
			rowx -= *browx
		}
		if colRel != 0 && bcolx != nil {
			colx -= *bcolx
		}
	}
	return rowx, colx, rowRel, colRel
}

// GetCellAddr extracts cell address from binary data.
func GetCellAddr(data []byte, pos int, bv int, reldelta bool, browx, bcolx *int) (int, int, int, int) {
	if bv >= 80 {
		rowval := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		colval := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		return AdjustCellAddrBiff8(rowval, colval, reldelta, browx, bcolx)
	} else {
		rowval := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		colval := int(data[pos+2])
		return AdjustCellAddrBiffLe7(rowval, colval, reldelta, browx, bcolx)
	}
}

// GetCellRangeAddr extracts cell range address from binary data.
func GetCellRangeAddr(data []byte, pos int, bv int, reldelta bool, browx, bcolx *int) ([4]int, [4]int) {
	if bv >= 80 {
		row1val := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		row2val := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		col1val := int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))
		col2val := int(binary.LittleEndian.Uint16(data[pos+6 : pos+8]))

		r1, c1, rr1, cr1 := AdjustCellAddrBiff8(row1val, col1val, reldelta, browx, bcolx)
		r2, c2, rr2, cr2 := AdjustCellAddrBiff8(row2val, col2val, reldelta, browx, bcolx)
		return [4]int{r1, c1, rr1, cr1}, [4]int{r2, c2, rr2, cr2}
	} else {
		row1val := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		row2val := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		col1val := int(data[pos+4])
		col2val := int(data[pos+5])

		r1, c1, rr1, cr1 := AdjustCellAddrBiffLe7(row1val, col1val, reldelta, browx, bcolx)
		r2, c2, rr2, cr2 := AdjustCellAddrBiffLe7(row2val, col2val, reldelta, browx, bcolx)
		return [4]int{r1, c1, rr1, cr1}, [4]int{r2, c2, rr2, cr2}
	}
}

// QuotedSheetName returns a quoted sheet name if necessary.
func QuotedSheetName(shnames []string, shx int) string {
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

// SheetRange returns a sheet range description.
func SheetRange(book *Book, slo, shi int) string {
	shnames := book.SheetNames()
	shdesc := QuotedSheetName(shnames, slo)
	if slo != shi-1 {
		shdesc += ":" + QuotedSheetName(shnames, shi-1)
	}
	return shdesc
}

// SheetRangeRel returns a relative sheet range description.
func SheetRangeRel(book *Book, srange []int, srangerel []int) string {
	slo, shi := srange[0], srange[1]
	slorel, shirel := srangerel[0], srangerel[1]

	if slorel == 0 && shirel == 0 {
		return SheetRange(book, slo, shi)
	}
	// Current sheet relative reference
	if slo == 0 && shi == 1 && slorel != 0 && shirel != 0 {
		return ""
	}
	return ""
}

// RangeName3D returns a 3D range name.
func RangeName3D(book *Book, ref3d *Ref3D) string {
	return fmt.Sprintf("%s!%s",
		SheetRange(book, ref3d.ShtXLo, ref3d.ShtXHi),
		RangeName2D(ref3d.RowXLo, ref3d.RowXHi, ref3d.ColXLo, ref3d.ColXHi, false))
}

// RangeName3DRel returns a relative 3D range name.
func RangeName3DRel(book *Book, ref3d *Ref3D, browx, bcolx *int, r1c1 bool) string {
	shdesc := SheetRangeRel(book, []int{ref3d.ShtXLo, ref3d.ShtXHi}, []int{ref3d.RelFlags[0], ref3d.RelFlags[1]})
	rngdesc := RangeName2DRel([]int{ref3d.RowXLo, ref3d.RowXHi, ref3d.ColXLo, ref3d.ColXHi},
		[]int{ref3d.RelFlags[2], ref3d.RelFlags[3], ref3d.RelFlags[4], ref3d.RelFlags[5]},
		browx, bcolx, r1c1)

	if shdesc == "" {
		return rngdesc
	}
	return fmt.Sprintf("%s!%s", shdesc, rngdesc)
}

// DecompileFormula decompiles a formula into a human-readable string.
func DecompileFormula(bk *Book, fmla []byte, fmlalen int, fmlatype int, browx, bcolx *int, blah, level, r1c1 int) (string, error) {
	if level > STACK_ALARM_LEVEL {
		blah = 1
	}

	reldelta := 0
	if fmlatype == FMLA_TYPE_SHARED || fmlatype == FMLA_TYPE_NAME || fmlatype == FMLA_TYPE_COND_FMT || fmlatype == FMLA_TYPE_DATA_VAL {
		reldelta = 1
	}

	data := fmla
	bv := bk.BiffVersion

	if blah != 0 && bk.logfile != nil {
		fmt.Fprintf(bk.logfile, "::: decompile_formula len=%d fmlatype=%d browx=%v bcolx=%v reldelta=%d r1c1=%d level=%d\n",
			fmlalen, fmlatype, browx, bcolx, reldelta, r1c1, level)
		// TODO: Implement hex dump
	}

	if level > STACK_PANIC_LEVEL {
		return "#RECURSION!", fmt.Errorf("excessive indirect references in formula")
	}

	sztab, exists := szdict[bv]
	if !exists {
		return "#UNSUPPORTED_BIFF!", fmt.Errorf("unsupported BIFF version %d", bv)
	}

	pos := 0
	stack := []*Operand{}
	unkOpnd := NewOperand(oUNK, nil, LEAF_RANK, "?")
	errorOpnd := NewOperand(oERR, nil, LEAF_RANK, "#ERROR!")

	_ = errorOpnd // Mark as used

	doBinop := func(opcd int, stk []*Operand) {
		if len(stk) < 2 {
			return
		}
		bop := stk[len(stk)-1]
		stk = stk[:len(stk)-1]
		aop := stk[len(stk)-1]
		stk = stk[:len(stk)-1]

		rule, exists := binopRules[opcd]
		if !exists {
			stk = append(stk, unkOpnd)
			return
		}

		resultKind := rule[1].(int)
		rank := rule[3].(int)
		sym := rule[4].(string)

		otext := ""
		if aop.Rank < rank {
			otext += "("
		}
		otext += aop.Text
		if aop.Rank < rank {
			otext += ")"
		}
		otext += sym
		if bop.Rank < rank {
			otext += "("
		}
		otext += bop.Text
		if bop.Rank < rank {
			otext += ")"
		}

		resop := NewOperand(resultKind, nil, rank, otext)
		stk = append(stk, resop)
	}

	doUnaryop := func(opcode, resultKind int, stk []*Operand) {
		if len(stk) < 1 {
			return
		}
		aop := stk[len(stk)-1]
		stk = stk[:len(stk)-1]

		rule, exists := unopRules[opcode]
		if !exists {
			stack = append(stack, unkOpnd)
			return
		}

		rank := rule[1].(int)
		sym1 := rule[2].(string)
		sym2 := rule[3].(string)

		otext := sym1
		if aop.Rank < rank {
			otext += "("
		}
		otext += aop.Text
		if aop.Rank < rank {
			otext += ")"
		}
		otext += sym2

		stack = append(stack, NewOperand(resultKind, nil, rank, otext))
	}

	unexpectedOpcode := func(opArg int, onameArg string) {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "ERROR *** Unexpected token 0x%02x (%s) found in formula type %s\n",
				opArg, onameArg, FmlaTypeDescrMap[fmlatype])
		}
	}

	if fmlalen == 0 {
		stack = append(stack, unkOpnd)
	}

	for pos < fmlalen {
		if pos >= len(data) {
			break
		}

		op := int(data[pos])
		opcode := op & 0x1f
		optype := (op & 0x60) >> 5
		opx := opcode
		if optype != 0 {
			opx = opcode + 32
		}

		oname := "?"
		if opx < len(onames) {
			oname = onames[opx]
		}

		sz := -1
		if opx < len(sztab) {
			sz = sztab[opx]
		}

		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "Pos:%d Op:0x%02x opname:%s Sz:%d opcode:%02xh optype:%02xh\n",
				pos, op, oname, sz, opcode, optype)
			fmt.Fprintf(bk.logfile, "Stack = %v\n", stack)
		}

		if sz == -2 {
			return "#INVALID!", fmt.Errorf("ERROR *** Unexpected token 0x%02x (\"%s\"); biff_version=%d", op, oname, bv)
		}

		// Check if token is not allowed for this formula type
		if fmlatype >= 0 && fmlatype < len(TokenNotAllowed) {
			if TokenNotAllowed[fmlatype]&(1<<uint(opx)) != 0 {
				unexpectedOpcode(op, oname)
			}
		}

		if optype == 0 {
			if opcode <= 0x01 { // tExp, tTbl
				if bv >= 30 {
					// Parse shared formula reference
					if pos+4 >= len(data) {
						return "#INVALID!", fmt.Errorf("insufficient data for shared formula reference")
					}
					rowx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
					colx := int(binary.LittleEndian.Uint16(data[pos+3 : pos+5]))
					text := fmt.Sprintf("SHARED FMLA at rowx=%d colx=%d", rowx, colx)
					stack = append(stack, NewOperand(oUNK, nil, LEAF_RANK, text))
					if fmlatype != FMLA_TYPE_CELL && fmlatype != FMLA_TYPE_ARRAY {
						unexpectedOpcode(op, oname)
					}
				} else {
					// BIFF < 30 shared formula
					if pos+3 >= len(data) {
						return "#INVALID!", fmt.Errorf("insufficient data for shared formula reference")
					}
					rowx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
					colx := int(data[pos+3])
					text := fmt.Sprintf("SHARED FMLA at rowx=%d colx=%d", rowx, colx)
					stack = append(stack, NewOperand(oUNK, nil, LEAF_RANK, text))
					if fmlatype != FMLA_TYPE_CELL && fmlatype != FMLA_TYPE_ARRAY {
						unexpectedOpcode(op, oname)
					}
				}
			} else if 0x03 <= opcode && opcode <= 0x0E { // binary operations
				doBinop(opcode, stack)
			} else if opcode == 0x0F { // tIsect
				// TODO: Implement intersection logic
				pos += sz
				continue
			} else if opcode == 0x10 { // tUnion/List
				// TODO: Implement union logic
				pos += sz
				continue
			} else if opcode == 0x11 { // tRange
				// TODO: Implement range logic
				pos += sz
				continue
			} else if opcode == 0x12 { // tUplus
				doUnaryop(0x12, oNUM, stack)
			} else if opcode == 0x13 { // tUminus
				doUnaryop(0x13, oNUM, stack)
			} else if opcode == 0x14 { // tPercent
				doUnaryop(0x14, oNUM, stack)
			} else if opcode == 0x15 { // tParen
				if len(stack) > 0 {
					aop := stack[len(stack)-1]
					stack = stack[:len(stack)-1]
					otext := "(" + aop.Text + ")"
					stack = append(stack, NewOperand(aop.Kind, aop.Value, FUNC_RANK, otext))
				}
			} else if opcode == 0x16 { // tMissArg
				stack = append(stack, NewOperand(oMSNG, nil, LEAF_RANK, ""))
			} else if opcode == 0x17 { // tStr
				// Parse string
				if pos+1 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for string")
				}
				strlen := int(data[pos+1])
				if pos+2+strlen > len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for string content")
				}
				strval := string(data[pos+2 : pos+2+strlen])
				stack = append(stack, NewOperand(oSTRG, strval, LEAF_RANK, fmt.Sprintf("\"%s\"", strval)))
			} else if opcode == 0x18 { // tExtended
				// TODO: Handle extended token
				pos += sz
				continue
			} else if opcode == 0x19 { // tAttr
				// Parse attributes
				if pos+2 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for attributes")
				}
				attr := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				if attr&0x04 != 0 { // volatile
					// TODO: Handle volatile
				}
				if attr&0x10 != 0 { // if
					// TODO: Handle if
				}
				pos += sz
				continue
			} else if opcode == 0x1A { // tSheet
				// TODO: Handle sheet reference
				pos += sz
				continue
			} else if opcode == 0x1B { // tEndSheet
				// TODO: Handle end sheet
				pos += sz
				continue
			} else if opcode == 0x1C { // tErr
				// Parse error code
				if pos+2 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for error")
				}
				errcode := int(data[pos+1])
				errtext := "#ERROR!"
				switch errcode {
				case 0x00:
					errtext = "#NULL!"
				case 0x07:
					errtext = "#DIV/0!"
				case 0x0F:
					errtext = "#VALUE!"
				case 0x17:
					errtext = "#REF!"
				case 0x1D:
					errtext = "#NAME?"
				case 0x24:
					errtext = "#NUM!"
				case 0x2A:
					errtext = "#N/A"
				}
				stack = append(stack, NewOperand(oERR, errcode, LEAF_RANK, errtext))
			} else if opcode == 0x1D { // tBool
				// Parse boolean
				if pos+2 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for boolean")
				}
				boolval := int(data[pos+1])
				text := "FALSE"
				if boolval != 0 {
					text = "TRUE"
				}
				stack = append(stack, NewOperand(oBOOL, boolval, LEAF_RANK, text))
			} else if opcode == 0x1E { // tInt
				// Parse integer
				if pos+3 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for integer")
				}
				intval := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				stack = append(stack, NewOperand(oNUM, float64(intval), LEAF_RANK, fmt.Sprintf("%d", intval)))
			} else if opcode == 0x1F { // tNum
				// Parse number
				if pos+9 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for number")
				}
				numval := math.Float64frombits(binary.LittleEndian.Uint64(data[pos+1 : pos+9]))
				stack = append(stack, NewOperand(oNUM, numval, LEAF_RANK, fmt.Sprintf("%g", numval)))
			}
		} else {
			// Handle RVA (Reference, Value, Array) tokens
			switch opx {
			case 0x20: // tArray
				// TODO: Parse array constant
				pos += sz
				continue
			case 0x21: // tFunc
				// Parse function
				if pos+2 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for function")
				}
				funcid := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				funcname := "UNKNOWN_FUNC"
				if funcid < len(funcDefs) {
					if name, ok := funcDefs[funcid][0].(string); ok {
						funcname = name
					}
				}
				// TODO: Handle function arguments
				stack = append(stack, NewOperand(oUNK, nil, FUNC_RANK, funcname+"(?)"))
			case 0x22: // tFuncVar
				// Parse variable function
				if pos+3 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for variable function")
				}
				funcid := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				nargs := int(data[pos+3])
				funcname := "UNKNOWN_FUNC"
				if funcid < len(funcDefs) {
					if name, ok := funcDefs[funcid][0].(string); ok {
						funcname = name
					}
				}
				// TODO: Handle function arguments
				stack = append(stack, NewOperand(oUNK, nil, FUNC_RANK, fmt.Sprintf("%s(?%d)", funcname, nargs)))
			case 0x23: // tName
				// Parse name reference
				if pos+4 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for name reference")
				}
				namex := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				// TODO: Resolve name
				stack = append(stack, NewOperand(oUNK, nil, LEAF_RANK, fmt.Sprintf("NAME_%d", namex)))
			case 0x24: // tRef
				// Parse cell reference
				r1, c1, rr1, cr1 := GetCellAddr(data, pos+1, bv, reldelta != 0, browx, bcolx)
				reftext := CellNameRel(r1, c1, rr1, cr1, browx, bcolx, r1c1 != 0)
				stack = append(stack, NewOperand(oREF, []int{r1, c1, rr1, cr1}, LEAF_RANK, reftext))
			case 0x25: // tArea
				// Parse area reference
				addr1, addr2 := GetCellRangeAddr(data, pos+1, bv, reldelta != 0, browx, bcolx)
				r1, c1, rr1, cr1 := addr1[0], addr1[1], addr1[2], addr1[3]
				r2, c2, rr2, cr2 := addr2[0], addr2[1], addr2[2], addr2[3]
				reftext := RangeName2DRel([]int{r1, r2, c1, c2}, []int{rr1, rr2, cr1, cr2}, browx, bcolx, r1c1 != 0)
				stack = append(stack, NewOperand(oREF, []int{r1, c1, rr1, cr1, r2, c2, rr2, cr2}, LEAF_RANK, reftext))
			case 0x26: // tMemArea
				// TODO: Parse memory area
				pos += sz
				continue
			case 0x27: // tMemErr
				// TODO: Parse memory error
				pos += sz
				continue
			case 0x28: // tMemNoMem
				// TODO: Parse no memory
				pos += sz
				continue
			case 0x29: // tMemFunc
				// TODO: Parse memory function
				pos += sz
				continue
			case 0x2A: // tRefErr
				// Parse reference error
				stack = append(stack, NewOperand(oERR, nil, LEAF_RANK, "#REF!"))
			case 0x2B: // tAreaErr
				// Parse area error
				stack = append(stack, NewOperand(oERR, nil, LEAF_RANK, "#REF!"))
			case 0x2C: // tRefN
				// Parse relative reference
				r1, c1, rr1, cr1 := GetCellAddr(data, pos+1, bv, reldelta != 0, browx, bcolx)
				reftext := CellNameRel(r1, c1, rr1, cr1, browx, bcolx, r1c1 != 0)
				stack = append(stack, NewOperand(oREL, []int{r1, c1, rr1, cr1}, LEAF_RANK, reftext))
			case 0x2D: // tAreaN
				// Parse relative area
				addr1, addr2 := GetCellRangeAddr(data, pos+1, bv, reldelta != 0, browx, bcolx)
				r1, c1, rr1, cr1 := addr1[0], addr1[1], addr1[2], addr1[3]
				r2, c2, rr2, cr2 := addr2[0], addr2[1], addr2[2], addr2[3]
				reftext := RangeName2DRel([]int{r1, r2, c1, c2}, []int{rr1, rr2, cr1, cr2}, browx, bcolx, r1c1 != 0)
				stack = append(stack, NewOperand(oREL, []int{r1, c1, rr1, cr1, r2, c2, rr2, cr2}, LEAF_RANK, reftext))
			case 0x2E: // tMemAreaN
				// TODO: Parse relative memory area
				pos += sz
				continue
			case 0x2F: // tMemNoMemN
				// TODO: Parse relative no memory
				pos += sz
				continue
			case 0x39: // tNameX
				// Parse external name
				if pos+6 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for external name")
				}
				extshtx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				namex := int(binary.LittleEndian.Uint16(data[pos+3 : pos+5]))
				// TODO: Resolve external name
				sht1, sht2 := GetExternsheetLocalRange(bk, extshtx, blah)
				if sht1 >= 0 && sht2 >= 0 {
					sheetname := bk.SheetNames()[sht1]
					stack = append(stack, NewOperand(oUNK, nil, LEAF_RANK, fmt.Sprintf("%s!NAME_%d", sheetname, namex)))
				} else {
					stack = append(stack, NewOperand(oUNK, nil, LEAF_RANK, fmt.Sprintf("EXTERN_NAME_%d_%d", extshtx, namex)))
				}
			case 0x3A: // tRef3d
				// Parse 3D reference
				if pos+6 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for 3D reference")
				}
				extshtx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				r1, c1, rr1, cr1 := GetCellAddr(data, pos+3, bv, reldelta != 0, browx, bcolx)
				sht1, sht2 := GetExternsheetLocalRange(bk, extshtx, blah)
				reftext := ""
				if sht1 >= 0 && sht2 >= 0 {
					ref3d := NewRef3D([]int{sht1, sht2, r1, r1 + 1, c1, c1 + 1, rr1, rr1, cr1, cr1})
					reftext = RangeName3DRel(bk, ref3d, browx, bcolx, r1c1 != 0)
				}
				stack = append(stack, NewOperand(oREF, []int{sht1, sht2, r1, c1, rr1, cr1}, LEAF_RANK, reftext))
			case 0x3B: // tArea3d
				// Parse 3D area
				if pos+10 >= len(data) {
					return "#INVALID!", fmt.Errorf("insufficient data for 3D area")
				}
				extshtx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
				addr1, addr2 := GetCellRangeAddr(data, pos+3, bv, reldelta != 0, browx, bcolx)
				r1, c1, rr1, cr1 := addr1[0], addr1[1], addr1[2], addr1[3]
				r2, c2, rr2, cr2 := addr2[0], addr2[1], addr2[2], addr2[3]
				sht1, sht2 := GetExternsheetLocalRange(bk, extshtx, blah)
				reftext := ""
				if sht1 >= 0 && sht2 >= 0 {
					ref3d := NewRef3D([]int{sht1, sht2, r1, r2, c1, c2, rr1, rr2, cr1, cr2})
					reftext = RangeName3DRel(bk, ref3d, browx, bcolx, r1c1 != 0)
				}
				stack = append(stack, NewOperand(oREF, []int{sht1, sht2, r1, c1, rr1, cr1, r2, c2, rr2, cr2}, LEAF_RANK, reftext))
			case 0x3C: // tRefErr3d
				// Parse 3D reference error
				stack = append(stack, NewOperand(oERR, nil, LEAF_RANK, "#REF!"))
			case 0x3D: // tAreaErr3d
				// Parse 3D area error
				stack = append(stack, NewOperand(oERR, nil, LEAF_RANK, "#REF!"))
			}
		}

		if sz <= 0 {
			sz = 1
		}
		pos += sz
	}

	if len(stack) == 0 {
		return "?", nil
	}

	result := stack[len(stack)-1]
	return result.Text, nil
}

// DumpFormula dumps a formula for debugging purposes.
func DumpFormula(bk *Book, data []byte, fmlalen int, bv int, reldelta int, blah int, isname int) {
	if blah == 0 || bk.logfile == nil {
		return
	}

	fmt.Fprintf(bk.logfile, "dump_formula %d %d %d\n", fmlalen, bv, len(data))
	// Use the existing HexCharDump function from biffh.go
	HexCharDump(data, 0, fmlalen, 0, bk.logfile, false)

	if bv < 80 {
		fmt.Fprintf(bk.logfile, "DumpFormula: BIFF version %d not supported\n", bv)
		return
	}

	sztab, exists := szdict[bv]
	if !exists {
		fmt.Fprintf(bk.logfile, "DumpFormula: unsupported BIFF version %d\n", bv)
		return
	}

	pos := 0
	for pos < fmlalen {
		if pos >= len(data) {
			break
		}

		op := int(data[pos])
		opcode := op & 0x1f
		optype := (op & 0x60) >> 5
		opx := opcode
		if optype != 0 {
			opx = opcode + 32
		}

		oname := "?"
		if opx < len(onames) {
			oname = onames[opx]
		}

		sz := -1
		if opx < len(sztab) {
			sz = sztab[opx]
		}

		fmt.Fprintf(bk.logfile, "Pos:%d Op:0x%02x Name:t%s Sz:%d opcode:%02xh optype:%02xh\n",
			pos, op, oname, sz, opcode, optype)

		if !(optype == 0) {
			if opcode >= 0x01 && opcode <= 0x02 { // tExp, tTbl
				// reference to a shared formula or table record
				if pos+4 < len(data) {
					rowx := int(binary.LittleEndian.Uint16(data[pos+1 : pos+3]))
					colx := int(binary.LittleEndian.Uint16(data[pos+3 : pos+5]))
					fmt.Fprintf(bk.logfile, "  (%d, %d)\n", rowx, colx)
				}
			}
		} else {
			// RVA tokens - could add more detailed parsing here
			// For now, just show the token
		}

		if sz <= 0 {
			sz = 1
		}
		pos += sz
	}
}

// colname returns the column name for a given column index (0-based).
// Example: colname(0) returns "A", colname(25) returns "Z", colname(26) returns "AA"
func colname(colx int) string {
	alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if colx <= 25 {
		return string(alphabet[colx])
	}
	xdiv26 := colx / 26
	xmod26 := colx % 26
	return string(alphabet[xdiv26-1]) + string(alphabet[xmod26])
}

// CellName returns the cell name for a given row and column (0-based).
// Example: CellName(0, 0) returns "A1", CellName(5, 7) returns "H6"
func CellName(rowx, colx int) string {
	return colname(colx) + fmt.Sprintf("%d", rowx+1)
}

// CellNameAbs returns the absolute cell name.
// Example: CellNameAbs(5, 7, false) returns "$H$6"
// If r1c1 is true, returns R1C1 style: "R6C8"
func CellNameAbs(rowx, colx int, r1c1 bool) string {
	if r1c1 {
		return fmt.Sprintf("R%dC%d", rowx+1, colx+1)
	}
	return fmt.Sprintf("$%s$%d", colname(colx), rowx+1)
}

// RangeName2D returns a 2D range name.
// Example: RangeName2D(5, 20, 7, 10, false) returns "$H$6:$J$20"
func RangeName2D(rlo, rhi, clo, chi int, r1c1 bool) string {
	if r1c1 {
		return fmt.Sprintf("R%dC%d:R%dC%d", rlo+1, clo+1, rhi, chi)
	}
	if rhi == rlo+1 && chi == clo+1 {
		return CellNameAbs(rlo, clo, r1c1)
	}
	return fmt.Sprintf("%s:%s", CellNameAbs(rlo, clo, r1c1), CellNameAbs(rhi-1, chi-1, r1c1))
}

// DoBoxFuncs applies functions to corresponding coordinates of two Ref3D boxes.
func DoBoxFuncs(boxFuncs []func(int, int) int, boxa, boxb *Ref3D) []int {
	result := make([]int, 6)
	for i := 0; i < 6; i++ {
		result[i] = boxFuncs[i](boxa.Coords[i], boxb.Coords[i])
	}
	return result
}

// GetExternsheetLocalRange returns the local sheet range for an EXTERNSHEET reference.
// This is used for BIFF8 and later.
func GetExternsheetLocalRange(bk *Book, refx int, blah int) (int, int) {
	if refx < 0 || refx >= len(bk.externsheetInfo) {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "!!! GetExternsheetLocalRange: refx=%d, not in range(%d)\n", refx, len(bk.externsheetInfo))
		}
		return -101, -101
	}

	info := bk.externsheetInfo[refx]
	refRecordx, refFirstSheetx, refLastSheetx := info[0], info[1], info[2]

	if bk.supbookAddinsInx != nil && refRecordx == *bk.supbookAddinsInx {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRange(refx=%d) -> addins %v\n", refx, info)
		}
		if refFirstSheetx != 0xFFFE || refLastSheetx != 0xFFFE {
			return -103, -103 // stuffed up somewhere
		}
		return -5, -5
	}

	if bk.supbookLocalsInx == nil || refRecordx != *bk.supbookLocalsInx {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRange(refx=%d) -> external %v\n", refx, info)
		}
		return -4, -4 // external reference
	}

	if refFirstSheetx == 0xFFFE && refLastSheetx == 0xFFFE {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRange(refx=%d) -> unspecified sheet %v\n", refx, info)
		}
		return -1, -1 // internal reference, any sheet
	}

	if refFirstSheetx == 0xFFFF && refLastSheetx == 0xFFFF {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRange(refx=%d) -> deleted sheet(s)\n", refx)
		}
		return -2, -2 // internal reference, deleted sheet(s)
	}

	nsheets := len(bk.sheetNames)
	if !(0 <= refFirstSheetx && refFirstSheetx <= refLastSheetx && refLastSheetx < nsheets) {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRange(refx=%d) -> %v\n", refx, info)
			fmt.Fprintf(bk.logfile, "--- first/last sheet not in range(%d)\n", nsheets)
		}
		return -102, -102 // stuffed up somewhere
	}

	// In Go, sheetNames only contains worksheets, so indices are direct
	if !(0 <= refFirstSheetx && refLastSheetx < nsheets) {
		return -3, -3 // internal reference, but to a macro sheet
	}

	return refFirstSheetx, refLastSheetx
}

// GetExternsheetLocalRangeB57 returns the local sheet range for an EXTERNSHEET reference.
// This is used for BIFF 5.7 and earlier.
func GetExternsheetLocalRangeB57(bk *Book, rawExtshtx, refFirstSheetx, refLastSheetx int, blah int) (int, int) {
	if rawExtshtx > 0 {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRangeB57(raw_extshtx=%d) -> external\n", rawExtshtx)
		}
		return -4, -4 // external reference
	}

	if refFirstSheetx == -1 && refLastSheetx == -1 {
		return -2, -2 // internal reference, deleted sheet(s)
	}

	nsheets := len(bk.sheetNames)
	if !(0 <= refFirstSheetx && refFirstSheetx <= refLastSheetx && refLastSheetx < nsheets) {
		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "/// GetExternsheetLocalRangeB57(%d, %d, %d) -> ???\n", rawExtshtx, refFirstSheetx, refLastSheetx)
			fmt.Fprintf(bk.logfile, "--- first/last sheet not in range(%d)\n", nsheets)
		}
		return -103, -103 // stuffed up somewhere
	}

	// In Go, sheetNames only contains worksheets, so indices are direct
	if !(0 <= refFirstSheetx && refLastSheetx < nsheets) {
		return -3, -3 // internal reference, but to a macro sheet
	}

	return refFirstSheetx, refLastSheetx
}

// EvaluateNameFormula evaluates a named formula.
// This is a complex function that parses formula tokens for named references.
func EvaluateNameFormula(bk *Book, nobj *Name, namex int, blah, level int) *Operand {
	if level > STACK_ALARM_LEVEL {
		blah = 1
	}

	data := nobj.RawFormula
	fmlalen := nobj.BasicFormulaLen
	bv := bk.BiffVersion
	// reldelta := 1 // All defined name formulas use "Method B" [OOo docs] - not used

	if blah != 0 && bk.logfile != nil {
		fmt.Fprintf(bk.logfile, "::: evaluate_name_formula %d %s %d %d %v level=%d\n",
			namex, nobj.Name, fmlalen, bv, data, level)
		// TODO: Implement hex dump
	}

	if level > STACK_PANIC_LEVEL {
		// TODO: Return error operand instead of panicking
		return NewOperand(oERR, nil, 0, "#ERROR!")
	}

	sztab, exists := szdict[bv]
	if !exists {
		return NewOperand(oERR, nil, 0, "#UNSUPPORTED_BIFF!")
	}

	pos := 0
	stack := []*Operand{}
	anyRel := 0
	anyErr := 0
	anyExternal := 0
	unkOpnd := NewOperand(oUNK, nil, LEAF_RANK, "?")
	errorOpnd := NewOperand(oERR, nil, LEAF_RANK, "#ERROR!")

	_ = anyRel      // Mark as used
	_ = anyErr      // Mark as used
	_ = anyExternal // Mark as used
	_ = errorOpnd   // Mark as used

	doBinop := func(opcd int, stk []*Operand) {
		if len(stk) < 2 {
			return
		}
		bop := stk[len(stk)-1]
		stk = stk[:len(stk)-1]
		aop := stk[len(stk)-1]
		stk = stk[:len(stk)-1]

		rule, exists := binopRules[opcd]
		if !exists {
			stk = append(stk, unkOpnd)
			return
		}

		argdict := rule[0].(map[int]func(interface{}) interface{})
		resultKind := rule[1].(int)
		fn := rule[2].(func(interface{}, interface{}) interface{})
		rank := rule[3].(int)
		sym := rule[4].(string)

		otext := ""
		if aop.Rank < rank {
			otext += "("
		}
		otext += aop.Text
		if aop.Rank < rank {
			otext += ")"
		}
		otext += sym
		if bop.Rank < rank {
			otext += "("
		}
		otext += bop.Text
		if bop.Rank < rank {
			otext += ")"
		}

		resop := NewOperand(resultKind, nil, rank, otext)

		bconv, bExists := argdict[bop.Kind]
		aconv, aExists := argdict[aop.Kind]
		if !bExists || !aExists {
			stack = append(stack, resop)
			return
		}

		if bop.Value == nil || aop.Value == nil {
			stack = append(stack, resop)
			return
		}

		bval := bconv(bop.Value)
		aval := aconv(aop.Value)
		result := fn(aval, bval)
		if resultKind == oBOOL {
			if result.(bool) {
				result = 1
			} else {
				result = 0
			}
		}
		resop.Value = result
		stack = append(stack, resop)
	}

	doUnaryop := func(opcode, resultKind int, stk []*Operand) {
		if len(stk) < 1 {
			return
		}
		aop := stk[len(stk)-1]
		stk = stk[:len(stk)-1]

		rule, exists := unopRules[opcode]
		if !exists {
			stack = append(stack, unkOpnd)
			return
		}

		fn := rule[0].(func(float64) float64)
		rank := rule[1].(int)
		sym1 := rule[2].(string)
		sym2 := rule[3].(string)

		otext := sym1
		if aop.Rank < rank {
			otext += "("
		}
		otext += aop.Text
		if aop.Rank < rank {
			otext += ")"
		}
		otext += sym2

		val := aop.Value
		if val != nil {
			if fval, ok := val.(float64); ok {
				val = fn(fval)
			}
		}
		stack = append(stack, NewOperand(resultKind, val, rank, otext))
	}

	if fmlalen == 0 {
		return unkOpnd
	}

	for pos < fmlalen {
		if pos >= len(data) {
			break
		}

		op := int(data[pos])
		opcode := op & 0x1f
		optype := (op & 0x60) >> 5
		opx := opcode
		if optype != 0 {
			opx = opcode + 32
		}

		oname := "?"
		if opx < len(onames) {
			oname = onames[opx]
		}

		sz := -1
		if opx < len(sztab) {
			sz = sztab[opx]
		}

		if blah != 0 && bk.logfile != nil {
			fmt.Fprintf(bk.logfile, "Pos:%d Op:0x%02x Name:%s Sz:%d opcode:%02xh optype:%02xh\n",
				pos, op, oname, sz, opcode, optype)
			fmt.Fprintf(bk.logfile, "Stack: %v\n", stack)
		}

		if sz == -2 {
			return NewOperand(oERR, nil, 0, "#INVALID_TOKEN!")
		}

		if optype == 0 {
			switch opcode {
			case 0x00: // tExp
				if bv >= 30 {
					// TODO: Handle tExp for BIFF >= 30
					pos += sz
					continue
				} else {
					// TODO: Handle tExp for BIFF < 30
					pos += sz
					continue
				}
			case 0x01: // tTbl
				// Not allowed in name formulas
				return errorOpnd
			case 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E: // binary ops
				doBinop(opcode, stack)
			case 0x0F: // tIsect
				// TODO: Implement intersection
				pos += sz
				continue
			case 0x10: // tUnion/List
				// TODO: Implement union
				pos += sz
				continue
			case 0x11: // tRange
				// TODO: Implement range
				pos += sz
				continue
			case 0x12: // tUplus
				doUnaryop(0x12, oNUM, stack)
			case 0x13: // tUminus
				doUnaryop(0x13, oNUM, stack)
			case 0x14: // tPercent
				doUnaryop(0x14, oNUM, stack)
			case 0x15: // tParen
				if len(stack) > 0 {
					aop := stack[len(stack)-1]
					stack = stack[:len(stack)-1]
					otext := "(" + aop.Text + ")"
					stack = append(stack, NewOperand(aop.Kind, aop.Value, FUNC_RANK, otext))
				}
			case 0x16: // tMissArg
				stack = append(stack, NewOperand(oMSNG, nil, LEAF_RANK, ""))
			case 0x17: // tStr
				// TODO: Parse string
				pos += sz
				continue
			case 0x18: // tExtended
				// TODO: Handle extended token
				pos += sz
				continue
			case 0x19: // tAttr
				// TODO: Handle attributes
				pos += sz
				continue
			case 0x1A: // tSheet
				// TODO: Handle sheet reference
				pos += sz
				continue
			case 0x1B: // tEndSheet
				// TODO: Handle end sheet
				pos += sz
				continue
			case 0x1C: // tErr
				// TODO: Parse error code
				pos += sz
				continue
			case 0x1D: // tBool
				// TODO: Parse boolean
				pos += sz
				continue
			case 0x1E: // tInt
				// TODO: Parse integer
				pos += sz
				continue
			case 0x1F: // tNum
				// TODO: Parse number
				pos += sz
				continue
			default:
				pos += sz
				continue
			}
		} else {
			// Handle RVA (Reference, Value, Array) tokens
			switch opx {
			case 0x20: // tArray
				// TODO: Parse array
				pos += sz
				continue
			case 0x21: // tFunc
				// TODO: Parse function
				pos += sz
				continue
			case 0x22: // tFuncVar
				// TODO: Parse variable function
				pos += sz
				continue
			case 0x23: // tName
				// TODO: Parse name reference
				pos += sz
				continue
			case 0x24: // tRef
				// TODO: Parse cell reference
				pos += sz
				continue
			case 0x25: // tArea
				// TODO: Parse area reference
				pos += sz
				continue
			case 0x26: // tMemArea
				// TODO: Parse memory area
				pos += sz
				continue
			case 0x27: // tMemErr
				// TODO: Parse memory error
				pos += sz
				continue
			case 0x28: // tMemNoMem
				// TODO: Parse no memory
				pos += sz
				continue
			case 0x29: // tMemFunc
				// TODO: Parse memory function
				pos += sz
				continue
			case 0x2A: // tRefErr
				// TODO: Parse reference error
				pos += sz
				continue
			case 0x2B: // tAreaErr
				// TODO: Parse area error
				pos += sz
				continue
			case 0x2C: // tRefN
				// TODO: Parse relative reference
				pos += sz
				continue
			case 0x2D: // tAreaN
				// TODO: Parse relative area
				pos += sz
				continue
			case 0x2E: // tMemAreaN
				// TODO: Parse relative memory area
				pos += sz
				continue
			case 0x2F: // tMemNoMemN
				// TODO: Parse relative no memory
				pos += sz
				continue
			case 0x39: // tNameX
				// TODO: Parse external name
				pos += sz
				continue
			case 0x3A: // tRef3d
				// TODO: Parse 3D reference
				pos += sz
				continue
			case 0x3B: // tArea3d
				// TODO: Parse 3D area
				pos += sz
				continue
			case 0x3C: // tRefErr3d
				// TODO: Parse 3D reference error
				pos += sz
				continue
			case 0x3D: // tAreaErr3d
				// TODO: Parse 3D area error
				pos += sz
				continue
			default:
				pos += sz
				continue
			}
		}

		if sz <= 0 {
			sz = 1 // Minimum size
		}
		pos += sz
	}

	if len(stack) == 0 {
		return unkOpnd
	}

	result := stack[len(stack)-1]
	if anyErr != 0 {
		return errorOpnd
	}

	return result
}

// Formula parsing constants and dictionaries

var (
	// Size tables for different BIFF versions
	sztab0 = []int{-2, 4, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 8, 4, 2, 2, 3, 9, 8, 2, 3, 8, 4, 7, 5, 5, 5, 2, 4, 7, 4, 7, 2, 2, -2, -2, -2, -2, -2, -2, -2, -2, 3, -2, -2, -2, -2, -2, -2, -2}
	sztab1 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 11, 5, 2, 2, 3, 9, 9, 2, 3, 11, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, 3, -2, -2, -2, -2, -2, -2, -2}
	sztab2 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, 11, 5, 2, 2, 3, 9, 9, 3, 4, 11, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2, -2}
	sztab3 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -2, -1, -2, -2, 2, 2, 3, 9, 9, 3, 4, 15, 4, 7, 7, 7, 7, 3, 4, 7, 4, 7, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, 25, 18, 21, 18, 21, -2, -2}
	sztab4 = []int{-2, 5, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1, -1, -1, -2, -2, 2, 2, 3, 9, 9, 3, 4, 5, 5, 9, 7, 7, 7, 3, 5, 9, 5, 9, 3, 3, -2, -2, -2, -2, -2, -2, -2, -2, -2, 7, 7, 11, 7, 11, -2, -2}

	// Size dictionary mapping BIFF version to size table
	szdict = map[int][]int{
		20: sztab0,
		21: sztab0,
		30: sztab1,
		40: sztab2,
		45: sztab2,
		50: sztab3,
		70: sztab3,
		80: sztab4,
	}

	// Operation names for debugging
	onames = []string{
		"Unk00", "Exp", "Tbl", "Add", "Sub", "Mul", "Div", "Power", "Concat", "LT", "LE", "EQ", "GE", "GT", "NE",
		"Isect", "List", "Range", "Uplus", "Uminus", "Percent", "Paren", "MissArg", "Str", "Extended", "Attr",
		"Sheet", "EndSheet", "Err", "Bool", "Int", "Num", "Array", "Func", "FuncVar", "Name", "Ref", "Area",
		"MemArea", "MemErr", "MemNoMem", "MemFunc", "RefErr", "AreaErr", "RefN", "AreaN", "MemAreaN", "MemNoMemN",
		"", "", "", "", "", "", "", "", "FuncCE", "NameX", "Ref3d", "Area3d", "RefErr3d", "AreaErr3d", "", "",
	}

	// Function definitions: index -> (name, min_args, max_args, flags, known_args, return_type, arg_types)
	funcDefs = map[int][7]interface{}{
		0:   {"COUNT", 0, 30, 0x04, 1, "V", "R"},
		1:   {"IF", 2, 3, 0x04, 3, "V", "VRR"},
		2:   {"ISNA", 1, 1, 0x02, 1, "V", "V"},
		3:   {"ISERROR", 1, 1, 0x02, 1, "V", "V"},
		4:   {"SUM", 0, 30, 0x04, 1, "V", "R"},
		5:   {"AVERAGE", 1, 30, 0x04, 1, "V", "R"},
		6:   {"MIN", 1, 30, 0x04, 1, "V", "R"},
		7:   {"MAX", 1, 30, 0x04, 1, "V", "R"},
		8:   {"ROW", 0, 1, 0x04, 1, "V", "R"},
		9:   {"COLUMN", 0, 1, 0x04, 1, "V", "R"},
		10:  {"NA", 0, 0, 0x02, 0, "V", ""},
		11:  {"NPV", 2, 30, 0x04, 2, "V", "VR"},
		12:  {"STDEV", 1, 30, 0x04, 1, "V", "R"},
		13:  {"DOLLAR", 1, 2, 0x04, 1, "V", "V"},
		14:  {"FIXED", 2, 3, 0x04, 3, "V", "VVV"},
		15:  {"SIN", 1, 1, 0x02, 1, "V", "V"},
		16:  {"COS", 1, 1, 0x02, 1, "V", "V"},
		17:  {"TAN", 1, 1, 0x02, 1, "V", "V"},
		18:  {"ATAN", 1, 1, 0x02, 1, "V", "V"},
		19:  {"PI", 0, 0, 0x02, 0, "V", ""},
		20:  {"SQRT", 1, 1, 0x02, 1, "V", "V"},
		21:  {"EXP", 1, 1, 0x02, 1, "V", "V"},
		22:  {"LN", 1, 1, 0x02, 1, "V", "V"},
		23:  {"LOG10", 1, 1, 0x02, 1, "V", "V"},
		24:  {"ABS", 1, 1, 0x02, 1, "V", "V"},
		25:  {"INT", 1, 1, 0x02, 1, "V", "V"},
		26:  {"SIGN", 1, 1, 0x02, 1, "V", "V"},
		27:  {"ROUND", 2, 2, 0x04, 2, "V", "VV"},
		28:  {"LOOKUP", 2, 3, 0x04, 3, "V", "VRR"},
		29:  {"INDEX", 2, 4, 0x04, 4, "V", "RVVV"},
		30:  {"REPT", 2, 2, 0x04, 2, "V", "VV"},
		31:  {"MID", 3, 3, 0x04, 3, "V", "VVV"},
		32:  {"LEN", 1, 1, 0x02, 1, "V", "V"},
		33:  {"VALUE", 1, 1, 0x02, 1, "V", "V"},
		34:  {"TRUE", 0, 0, 0x02, 0, "V", ""},
		35:  {"FALSE", 0, 0, 0x02, 0, "V", ""},
		36:  {"AND", 1, 30, 0x04, 1, "V", "R"},
		37:  {"OR", 1, 30, 0x04, 1, "V", "R"},
		38:  {"NOT", 1, 1, 0x02, 1, "V", "V"},
		39:  {"MOD", 2, 2, 0x04, 2, "V", "VV"},
		40:  {"DCOUNT", 3, 3, 0x04, 3, "V", "RRR"},
		41:  {"DSUM", 3, 3, 0x04, 3, "V", "RRR"},
		42:  {"DAVERAGE", 3, 3, 0x04, 3, "V", "RRR"},
		43:  {"DMIN", 3, 3, 0x04, 3, "V", "RRR"},
		44:  {"DMAX", 3, 3, 0x04, 3, "V", "RRR"},
		45:  {"DSTDEV", 3, 3, 0x04, 3, "V", "RRR"},
		46:  {"VAR", 1, 30, 0x04, 1, "V", "R"},
		47:  {"DVAR", 3, 3, 0x04, 3, "V", "RRR"},
		48:  {"TEXT", 2, 2, 0x04, 2, "V", "VV"},
		49:  {"LINEST", 1, 4, 0x04, 4, "A", "RRVV"},
		50:  {"TREND", 1, 4, 0x04, 4, "A", "RRVV"},
		51:  {"LOGEST", 1, 4, 0x04, 4, "A", "RRVV"},
		52:  {"GROWTH", 1, 4, 0x04, 4, "A", "RRVV"},
		57:  {"TRANSPOSE", 1, 1, 0x02, 1, "A", "A"},
		61:  {"RAND", 0, 0, 0x02, 0, "V", ""},
		62:  {"MATCH", 2, 3, 0x04, 3, "V", "VRR"},
		63:  {"DATE", 3, 3, 0x04, 3, "V", "VVV"},
		64:  {"TIME", 3, 3, 0x04, 3, "V", "VVV"},
		65:  {"DAY", 1, 1, 0x02, 1, "V", "V"},
		66:  {"MONTH", 1, 1, 0x02, 1, "V", "V"},
		67:  {"YEAR", 1, 1, 0x02, 1, "V", "V"},
		68:  {"WEEKDAY", 1, 2, 0x04, 2, "V", "VV"},
		69:  {"HOUR", 1, 1, 0x02, 1, "V", "V"},
		70:  {"MINUTE", 1, 1, 0x02, 1, "V", "V"},
		71:  {"SECOND", 1, 1, 0x02, 1, "V", "V"},
		72:  {"NOW", 0, 0, 0x02, 0, "V", ""},
		73:  {"AREAS", 1, 1, 0x02, 1, "V", "R"},
		74:  {"ROWS", 1, 1, 0x02, 1, "V", "R"},
		75:  {"COLUMNS", 1, 1, 0x02, 1, "V", "R"},
		76:  {"OFFSET", 3, 5, 0x04, 5, "R", "VRVVV"},
		77:  {"SEARCH", 2, 3, 0x04, 3, "V", "VVV"},
		78:  {"TRANSPOSE", 1, 1, 0x02, 1, "A", "A"},
		79:  {"TYPE", 1, 1, 0x02, 1, "V", "V"},
		82:  {"ATAN2", 2, 2, 0x04, 2, "V", "VV"},
		83:  {"ASIN", 1, 1, 0x02, 1, "V", "V"},
		84:  {"ACOS", 1, 1, 0x02, 1, "V", "V"},
		85:  {"CHOOSE", 2, 30, 0x04, 2, "V", "VR"},
		86:  {"HLOOKUP", 3, 4, 0x04, 4, "V", "VRRV"},
		87:  {"VLOOKUP", 3, 4, 0x04, 4, "V", "VRRV"},
		88:  {"ISREF", 1, 1, 0x02, 1, "V", "R"},
		89:  {"LOG", 1, 2, 0x04, 2, "V", "VV"},
		97:  {"CHAR", 1, 1, 0x02, 1, "V", "V"},
		98:  {"LOWER", 1, 1, 0x02, 1, "V", "V"},
		99:  {"UPPER", 1, 1, 0x02, 1, "V", "V"},
		100: {"PROPER", 1, 1, 0x02, 1, "V", "V"},
		101: {"LEFT", 1, 2, 0x04, 2, "V", "VV"},
		102: {"RIGHT", 1, 2, 0x04, 2, "V", "VV"},
		103: {"EXACT", 2, 2, 0x04, 2, "V", "VV"},
		104: {"TRIM", 1, 1, 0x02, 1, "V", "V"},
		105: {"REPLACE", 4, 4, 0x04, 4, "V", "VVVV"},
		106: {"SUBSTITUTE", 3, 4, 0x04, 4, "V", "VVVV"},
		107: {"CODE", 1, 1, 0x02, 1, "V", "V"},
		109: {"FIND", 2, 3, 0x04, 3, "V", "VVV"},
		111: {"ISERR", 1, 1, 0x02, 1, "V", "V"},
		112: {"ISTEXT", 1, 1, 0x02, 1, "V", "V"},
		113: {"ISNUMBER", 1, 1, 0x02, 1, "V", "V"},
		114: {"ISBLANK", 1, 1, 0x02, 1, "V", "V"},
		115: {"T", 1, 1, 0x02, 1, "V", "R"},
		116: {"N", 1, 1, 0x02, 1, "V", "R"},
		117: {"DATEVALUE", 1, 1, 0x02, 1, "V", "V"},
		118: {"TIMEVALUE", 1, 1, 0x02, 1, "V", "V"},
		119: {"SLN", 3, 3, 0x04, 3, "V", "VVV"},
		120: {"SYD", 4, 4, 0x04, 4, "V", "VVVV"},
		121: {"DDB", 4, 5, 0x04, 5, "V", "VVVVV"},
		124: {"INDIRECT", 1, 2, 0x04, 2, "R", "VV"},
		125: {"CALLER", 0, 0, 0x02, 0, "R", ""},
		126: {"CLEAN", 1, 1, 0x02, 1, "V", "V"},
		127: {"MDETERM", 1, 1, 0x02, 1, "V", "A"},
		128: {"MINVERSE", 1, 1, 0x02, 1, "A", "A"},
		129: {"MMULT", 2, 2, 0x04, 2, "A", "AA"},
		130: {"IPMT", 4, 6, 0x04, 6, "V", "VVVVVV"},
		131: {"PPMT", 4, 6, 0x04, 6, "V", "VVVVVV"},
		132: {"COUNTA", 0, 30, 0x04, 1, "V", "R"},
		133: {"PRODUCT", 0, 30, 0x04, 1, "V", "R"},
		134: {"FACT", 1, 1, 0x02, 1, "V", "V"},
		135: {"DPRODUCT", 3, 3, 0x04, 3, "V", "RRR"},
		136: {"ISNONTEXT", 1, 1, 0x02, 1, "V", "V"},
		137: {"STDEVP", 1, 30, 0x04, 1, "V", "R"},
		138: {"VARP", 1, 30, 0x04, 1, "V", "R"},
		139: {"DSTDEVP", 3, 3, 0x04, 3, "V", "RRR"},
		140: {"DVARP", 3, 3, 0x04, 3, "V", "RRR"},
		141: {"TRUNC", 1, 2, 0x04, 2, "V", "VV"},
		142: {"ISLOGICAL", 1, 1, 0x02, 1, "V", "V"},
		143: {"DCOUNTA", 3, 3, 0x04, 3, "V", "RRR"},
		144: {"FINDB", 2, 3, 0x04, 3, "V", "VVV"},
		145: {"SEARCHB", 2, 3, 0x04, 3, "V", "VVV"},
		146: {"REPLACEB", 4, 4, 0x04, 4, "V", "VVVV"},
		147: {"LEFTB", 1, 2, 0x04, 2, "V", "VV"},
		148: {"RIGHTB", 1, 2, 0x04, 2, "V", "VV"},
		149: {"MIDB", 3, 3, 0x04, 3, "V", "VVV"},
		150: {"LENB", 1, 1, 0x02, 1, "V", "V"},
		151: {"ROUNDUP", 2, 2, 0x04, 2, "V", "VV"},
		152: {"ROUNDDOWN", 2, 2, 0x04, 2, "V", "VV"},
		153: {"ASC", 1, 1, 0x02, 1, "V", "V"},
		154: {"DBCS", 1, 1, 0x02, 1, "V", "V"},
		155: {"RANK", 2, 3, 0x04, 3, "V", "VRR"},
		156: {"ADDRESS", 2, 5, 0x04, 5, "V", "VVVVV"},
		157: {"DAYS360", 2, 2, 0x04, 2, "V", "VV"},
		158: {"TODAY", 0, 0, 0x02, 0, "V", ""},
		159: {"VDB", 5, 7, 0x04, 7, "V", "VVVVVVV"},
		160: {"MEDIAN", 1, 30, 0x04, 1, "V", "R"},
		161: {"SUMPRODUCT", 1, 30, 0x04, 1, "V", "A"},
		162: {"SINH", 1, 1, 0x02, 1, "V", "V"},
		163: {"COSH", 1, 1, 0x02, 1, "V", "V"},
		164: {"TANH", 1, 1, 0x02, 1, "V", "V"},
		165: {"ASINH", 1, 1, 0x02, 1, "V", "V"},
		166: {"ACOSH", 1, 1, 0x02, 1, "V", "V"},
		167: {"ATANH", 1, 1, 0x02, 1, "V", "V"},
		168: {"DGET", 3, 3, 0x04, 3, "V", "RRR"},
		169: {"INFO", 1, 1, 0x02, 1, "V", "V"},
		183: {"FREQUENCY", 2, 2, 0x04, 2, "A", "RV"},
		184: {"ERROR.TYPE", 1, 1, 0x02, 1, "V", "V"},
		185: {"REGISTER.ID", 2, 3, 0x04, 3, "V", "VVV"},
		186: {"AVEDEV", 1, 30, 0x04, 1, "V", "R"},
		187: {"BETADIST", 3, 5, 0x04, 5, "V", "VVVVV"},
		188: {"GAMMALN", 1, 1, 0x02, 1, "V", "V"},
		189: {"BETAINV", 3, 5, 0x04, 5, "V", "VVVVV"},
		190: {"BINOMDIST", 4, 4, 0x04, 4, "V", "VVVV"},
		191: {"CHIDIST", 2, 2, 0x04, 2, "V", "VV"},
		192: {"CHIINV", 2, 2, 0x04, 2, "V", "VV"},
		193: {"COMBIN", 2, 2, 0x04, 2, "V", "VV"},
		194: {"CONFIDENCE", 3, 3, 0x04, 3, "V", "VVV"},
		195: {"CRITBINOM", 3, 3, 0x04, 3, "V", "VVV"},
		196: {"EVEN", 1, 1, 0x02, 1, "V", "V"},
		197: {"EXPONDIST", 3, 3, 0x04, 3, "V", "VVV"},
		198: {"FDIST", 3, 3, 0x04, 3, "V", "VV"},
		199: {"FINV", 3, 3, 0x04, 3, "V", "VV"},
		200: {"FISHER", 1, 1, 0x02, 1, "V", "V"},
		201: {"FISHERINV", 1, 1, 0x02, 1, "V", "V"},
		202: {"FLOOR", 2, 2, 0x04, 2, "V", "VV"},
		203: {"GAMMADIST", 4, 4, 0x04, 4, "V", "VVVV"},
		204: {"GAMMAINV", 3, 3, 0x04, 3, "V", "VVV"},
		205: {"CEILING", 2, 2, 0x04, 2, "V", "VV"},
		206: {"HYPGEOMDIST", 4, 4, 0x04, 4, "V", "VVVV"},
		207: {"LOGNORMDIST", 3, 3, 0x04, 3, "V", "VVV"},
		208: {"LOGINV", 3, 3, 0x04, 3, "V", "VVV"},
		209: {"NEGBINOMDIST", 3, 3, 0x04, 3, "V", "VVV"},
		210: {"NORMDIST", 4, 4, 0x04, 4, "V", "VVVV"},
		211: {"NORMSDIST", 1, 1, 0x02, 1, "V", "V"},
		212: {"NORMSINV", 1, 1, 0x02, 1, "V", "V"},
		213: {"NORMINV", 3, 3, 0x04, 3, "V", "VVV"},
		214: {"PEARSON", 2, 2, 0x04, 2, "V", "AA"},
		215: {"POISSON", 3, 3, 0x04, 3, "V", "VVV"},
		216: {"TDIST", 3, 3, 0x04, 3, "V", "VVV"},
		217: {"TINV", 2, 2, 0x04, 2, "V", "VV"},
		218: {"WEIBULL", 4, 4, 0x04, 4, "V", "VVVV"},
		219: {"SUMXMY2", 2, 2, 0x04, 2, "V", "AA"},
		220: {"SUMX2MY2", 2, 2, 0x04, 2, "V", "AA"},
		221: {"SUMX2PY2", 2, 2, 0x04, 2, "V", "AA"},
		222: {"CHITEST", 2, 2, 0x04, 2, "V", "AA"},
		223: {"CORREL", 2, 2, 0x04, 2, "V", "AA"},
		224: {"COVAR", 2, 2, 0x04, 2, "V", "AA"},
		225: {"FTEST", 2, 2, 0x04, 2, "V", "AA"},
		226: {"INTERCEPT", 2, 2, 0x04, 2, "V", "AA"},
		227: {"PEARSON", 2, 2, 0x04, 2, "V", "AA"},
		228: {"RSQ", 2, 2, 0x04, 2, "V", "AA"},
		229: {"STEYX", 2, 2, 0x04, 2, "V", "AA"},
		230: {"SLOPE", 2, 2, 0x04, 2, "V", "AA"},
		231: {"TTEST", 4, 4, 0x04, 4, "V", "AAVV"},
		232: {"PROB", 3, 4, 0x04, 4, "V", "AAVV"},
		233: {"DEVSQ", 1, 30, 0x04, 1, "V", "R"},
		234: {"GEOMEAN", 1, 30, 0x04, 1, "V", "R"},
		235: {"HARMEAN", 1, 30, 0x04, 1, "V", "R"},
		236: {"SUMSQ", 1, 30, 0x04, 1, "V", "R"},
		237: {"KURT", 1, 30, 0x04, 1, "V", "R"},
		238: {"SKEW", 1, 30, 0x04, 1, "V", "R"},
		239: {"ZTEST", 2, 3, 0x04, 3, "V", "RVV"},
		240: {"LARGE", 2, 2, 0x04, 2, "V", "RV"},
		241: {"SMALL", 2, 2, 0x04, 2, "V", "RV"},
		242: {"QUARTILE", 2, 2, 0x04, 2, "V", "RV"},
		243: {"PERCENTILE", 2, 2, 0x04, 2, "V", "RV"},
		244: {"PERCENTRANK", 2, 3, 0x04, 3, "V", "RVV"},
		245: {"MODE", 1, 30, 0x04, 1, "V", "R"},
		246: {"TRIMMEAN", 2, 2, 0x04, 2, "V", "RV"},
		247: {"TINV2", 2, 2, 0x04, 2, "V", "VV"},
		252: {"CONCATENATE", 1, 30, 0x04, 1, "V", "V"},
		253: {"POWER", 2, 2, 0x04, 2, "V", "VV"},
		254: {"RADIANS", 1, 1, 0x02, 1, "V", "V"},
		255: {"DEGREES", 1, 1, 0x02, 1, "V", "V"},
		256: {"SUBTOTAL", 2, 30, 0x04, 2, "V", "VR"},
		257: {"SUMIF", 2, 3, 0x04, 3, "V", "RRV"},
		258: {"COUNTIF", 2, 2, 0x04, 2, "V", "RV"},
		259: {"COUNTBLANK", 1, 1, 0x02, 1, "V", "R"},
		260: {"ISPMT", 4, 4, 0x04, 4, "V", "VVVV"},
		261: {"DATEDIF", 3, 3, 0x04, 3, "V", "VVV"},
		262: {"DATESTRING", 1, 1, 0x02, 1, "V", "V"},
		263: {"NUMBERSTRING", 2, 2, 0x04, 2, "V", "VV"},
		269: {"SQRTPI", 1, 1, 0x02, 1, "V", "V"},
		270: {"RAND", 0, 0, 0x02, 0, "V", ""},
		271: {"NOW", 0, 0, 0x02, 0, "V", ""},
		272: {"TODAY", 0, 0, 0x02, 0, "V", ""},
		273: {"AREAS", 1, 1, 0x02, 1, "V", "R"},
		274: {"ROWS", 1, 1, 0x02, 1, "V", "R"},
		275: {"COLUMNS", 1, 1, 0x02, 1, "V", "R"},
		276: {"OFFSET", 3, 5, 0x04, 5, "R", "VRVVV"},
		277: {"SEARCH", 2, 3, 0x04, 3, "V", "VVV"},
		278: {"TRANSPOSE", 1, 1, 0x02, 1, "A", "A"},
		279: {"TYPE", 1, 1, 0x02, 1, "V", "V"},
		285: {"CALLER", 0, 0, 0x02, 0, "R", ""},
		288: {"SERIESSUM", 4, 4, 0x04, 4, "V", "VVVA"},
		289: {"FACTDOUBLE", 1, 1, 0x02, 1, "V", "V"},
		290: {"SQRTPI", 1, 1, 0x02, 1, "V", "V"},
		291: {"RANDBETWEEN", 2, 2, 0x04, 2, "V", "VV"},
		292: {"PRODUCT", 0, 30, 0x04, 1, "V", "R"},
		293: {"FACT", 1, 1, 0x02, 1, "V", "V"},
		294: {"DPRODUCT", 3, 3, 0x04, 3, "V", "RRR"},
		295: {"ISNONTEXT", 1, 1, 0x02, 1, "V", "V"},
		296: {"STDEVP", 1, 30, 0x04, 1, "V", "R"},
		297: {"VARP", 1, 30, 0x04, 1, "V", "R"},
		298: {"DSTDEVP", 3, 3, 0x04, 3, "V", "RRR"},
		299: {"DVARP", 3, 3, 0x04, 3, "V", "RRR"},
		300: {"TRUNC", 1, 2, 0x04, 2, "V", "VV"},
		301: {"ISLOGICAL", 1, 1, 0x02, 1, "V", "V"},
		302: {"DCOUNTA", 3, 3, 0x04, 3, "V", "RRR"},
		303: {"FINDB", 2, 3, 0x04, 3, "V", "VVV"},
		304: {"SEARCHB", 2, 3, 0x04, 3, "V", "VVV"},
		305: {"REPLACEB", 4, 4, 0x04, 4, "V", "VVVV"},
		306: {"LEFTB", 1, 2, 0x04, 2, "V", "VV"},
		307: {"RIGHTB", 1, 2, 0x04, 2, "V", "VV"},
		308: {"MIDB", 3, 3, 0x04, 3, "V", "VVV"},
		309: {"LENB", 1, 1, 0x02, 1, "V", "V"},
		310: {"ROUNDUP", 2, 2, 0x04, 2, "V", "VV"},
		311: {"ROUNDDOWN", 2, 2, 0x04, 2, "V", "VV"},
		312: {"ASC", 1, 1, 0x02, 1, "V", "V"},
		313: {"DBCS", 1, 1, 0x02, 1, "V", "V"},
		314: {"RANK", 2, 3, 0x04, 3, "V", "VRR"},
		315: {"ADDRESS", 2, 5, 0x04, 5, "V", "VVVVV"},
		316: {"DAYS360", 2, 2, 0x04, 2, "V", "VV"},
		317: {"TODAY", 0, 0, 0x02, 0, "V", ""},
		318: {"VDB", 5, 7, 0x04, 7, "V", "VVVVVVV"},
		319: {"MEDIAN", 1, 30, 0x04, 1, "V", "R"},
		320: {"SUMPRODUCT", 1, 30, 0x04, 1, "V", "A"},
		321: {"SINH", 1, 1, 0x02, 1, "V", "V"},
		322: {"COSH", 1, 1, 0x02, 1, "V", "V"},
		323: {"TANH", 1, 1, 0x02, 1, "V", "V"},
		324: {"ASINH", 1, 1, 0x02, 1, "V", "V"},
		325: {"ACOSH", 1, 1, 0x02, 1, "V", "V"},
		326: {"ATANH", 1, 1, 0x02, 1, "V", "V"},
		336: {"ISPMT", 4, 6, 0x04, 6, "V", "VVVVVV"},
		337: {"DATEDIF", 3, 3, 0x04, 3, "V", "VVV"},
		338: {"DATESTRING", 1, 1, 0x02, 1, "V", "V"},
		339: {"NUMBERSTRING", 2, 2, 0x04, 2, "V", "VV"},
		342: {"SUMSQ", 1, 30, 0x04, 1, "V", "R"},
		343: {"SUMX2MY2", 2, 2, 0x04, 2, "V", "AA"},
		344: {"SUMX2PY2", 2, 2, 0x04, 2, "V", "AA"},
		345: {"SUMXMY2", 2, 2, 0x04, 2, "V", "AA"},
		346: {"FACTDOUBLE", 1, 1, 0x02, 1, "V", "V"},
		347: {"SQRTPI", 1, 1, 0x02, 1, "V", "V"},
		348: {"RANDBETWEEN", 2, 2, 0x04, 2, "V", "VV"},
		349: {"SERIESSUM", 4, 4, 0x04, 4, "V", "VVVA"},
		350: {"SUBTOTAL", 2, 30, 0x04, 2, "V", "VR"},
		351: {"SUMIF", 2, 3, 0x04, 3, "V", "RRV"},
		352: {"COUNTIF", 2, 2, 0x04, 2, "V", "RV"},
		353: {"COUNTBLANK", 1, 1, 0x02, 1, "V", "R"},
		354: {"SCENARIO_GET", 2, 2, 0x04, 2, "V", "VV"},
		355: {"ISPMT", 4, 4, 0x04, 4, "V", "VVVV"},
		356: {"DATEDIF", 3, 3, 0x04, 3, "V", "VVV"},
		357: {"DATESTRING", 1, 1, 0x02, 1, "V", "V"},
		358: {"NUMBERSTRING", 2, 2, 0x04, 2, "V", "VV"},
		359: {"ROMAN", 1, 2, 0x04, 2, "V", "VV"},
		360: {"GETPIVOTDATA", 2, 30, 0x04, 2, "V", "VR"},
		361: {"HYPERLINK", 1, 2, 0x04, 2, "V", "VV"},
		362: {"PHONETIC", 1, 1, 0x02, 1, "V", "R"},
		363: {"AVERAGEA", 1, 30, 0x04, 1, "V", "R"},
		364: {"MAXA", 1, 30, 0x04, 1, "V", "R"},
		365: {"MINA", 1, 30, 0x04, 1, "V", "R"},
		366: {"STDEVPA", 1, 30, 0x04, 1, "V", "R"},
		367: {"VARPA", 1, 30, 0x04, 1, "V", "R"},
		368: {"STDEVA", 1, 30, 0x04, 1, "V", "R"},
		369: {"VARA", 1, 30, 0x04, 1, "V", "R"},
		370: {"BAHTTEXT", 1, 1, 0x02, 1, "V", "V"},
		384: {"THAIDAYOFWEEK", 1, 1, 0x02, 1, "V", "V"},
		385: {"THAIDIGIT", 1, 1, 0x02, 1, "V", "V"},
		386: {"THAIMONTHOFYEAR", 1, 1, 0x02, 1, "V", "V"},
		387: {"THAINUMSOUND", 1, 1, 0x02, 1, "V", "V"},
		388: {"THAINUMSTRING", 1, 1, 0x02, 1, "V", "V"},
		389: {"THAISTRINGLENGTH", 1, 1, 0x02, 1, "V", "V"},
		390: {"ISTHAIDIGIT", 1, 1, 0x02, 1, "V", "V"},
		391: {"ROUNDBAHTDOWN", 1, 1, 0x02, 1, "V", "V"},
		392: {"ROUNDBAHTUP", 1, 1, 0x02, 1, "V", "V"},
		393: {"THAIYEAR", 1, 1, 0x02, 1, "V", "V"},
		394: {"RTD", 2, 30, 0x04, 2, "V", "VV"},
	}

	// Argument conversion dictionaries for binary operations
	arithArgdict = map[int]func(interface{}) interface{}{
		oNUM:  nop,
		oSTRG: func(x interface{}) interface{} { return float64(x.(float64)) }, // identity for float64
	}

	cmpArgdict = map[int]func(interface{}) interface{}{
		oNUM:  nop,
		oSTRG: nop,
	}

	strgArgdict = map[int]func(interface{}) interface{}{
		oNUM:  func(x interface{}) interface{} { return Num2Str(x.(float64)) },
		oSTRG: nop,
	}

	// Binary operation rules: token -> (argdict, result_kind, func, rank, symbol)
	binopRules = map[int][5]interface{}{
		tAdd:    {arithArgdict, oNUM, func(a, b float64) float64 { return a + b }, 30, "+"},
		tSub:    {arithArgdict, oNUM, func(a, b float64) float64 { return a - b }, 30, "-"},
		tMul:    {arithArgdict, oNUM, func(a, b float64) float64 { return a * b }, 40, "*"},
		tDiv:    {arithArgdict, oNUM, func(a, b float64) float64 { return a / b }, 40, "/"},
		tPower:  {arithArgdict, oNUM, OprPow, 50, "^"},
		tConcat: {strgArgdict, oSTRG, func(a, b string) string { return a + b }, 20, "&"},
		tLT:     {cmpArgdict, oBOOL, OprLt, 10, "<"},
		tLE:     {cmpArgdict, oBOOL, OprLe, 10, "<="},
		tEQ:     {cmpArgdict, oBOOL, OprEq, 10, "="},
		tGE:     {cmpArgdict, oBOOL, OprGe, 10, ">="},
		tGT:     {cmpArgdict, oBOOL, OprGt, 10, ">"},
		tNE:     {cmpArgdict, oBOOL, OprNe, 10, "<>"},
	}

	// Unary operation rules: token -> (func, rank, sym1, sym2)
	unopRules = map[int][4]interface{}{
		0x12: {func(x float64) float64 { return x }, 70, "+", ""},         // unary plus
		0x13: {func(x float64) float64 { return -x }, 70, "-", ""},        // unary minus
		0x14: {func(x float64) float64 { return x / 100.0 }, 60, "", "%"}, // percent
	}
)
