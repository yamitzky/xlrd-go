package xlrd

import (
	"encoding/binary"
	"fmt"
	"io"
	"reflect"
	"unicode/utf16"
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

// DEBUG constant
const DEBUG = 0

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
	0x00: "#NULL!",  // Intersection of two cell ranges is empty
	0x07: "#DIV/0!", // Division by zero
	0x0F: "#VALUE!", // Wrong type of operand
	0x17: "#REF!",   // Illegal or deleted cell reference
	0x1D: "#NAME?",  // Wrong function or range name
	0x24: "#NUM!",   // Value range overflow
	0x2A: "#N/A",    // Argument or function not available
}

// BIFF record type constants
const (
	XL_WORKBOOK_GLOBALS      = 0x5
	XL_WORKBOOK_GLOBALS_4W   = 0x100
	XL_WORKSHEET             = 0x10
	XL_BOUNDSHEET_WORKSHEET  = 0x00
	XL_BOUNDSHEET_CHART      = 0x02
	XL_BOUNDSHEET_VB_MODULE  = 0x06
	XL_ARRAY                 = 0x0221
	XL_ARRAY2                = 0x0021
	XL_BLANK                 = 0x0201
	XL_BLANK_B2              = 0x01
	XL_BOF                   = 0x809
	XL_BOOLERR               = 0x205
	XL_BOOLERR_B2            = 0x5
	XL_BOUNDSHEET            = 0x85
	XL_BUILTINFMTCOUNT       = 0x56
	XL_CF                    = 0x01B1
	XL_CODEPAGE              = 0x42
	XL_COLINFO               = 0x7D
	XL_COLUMNDEFAULT         = 0x20 // BIFF2 only
	XL_COLWIDTH              = 0x24 // BIFF2 only
	XL_CONDFMT               = 0x01B0
	XL_CONTINUE              = 0x3c
	XL_COUNTRY               = 0x8C
	XL_DATEMODE              = 0x22
	XL_DEFAULTROWHEIGHT      = 0x0225
	XL_DEFCOLWIDTH           = 0x55
	XL_DIMENSION             = 0x200
	XL_DIMENSION2            = 0x0
	XL_EFONT                 = 0x45
	XL_EOF                   = 0x0a
	XL_EXTERNNAME            = 0x23
	XL_EXTERNSHEET           = 0x17
	XL_EXTSST                = 0xff
	XL_FEAT11                = 0x872
	XL_FILEPASS              = 0x2f
	XL_FONT                  = 0x31
	XL_FONT_B3B4             = 0x231
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
	XL_INTEGER               = 0x2  // BIFF2 only
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
	10006: "mac_greek",    // guess
	10007: "mac_cyrillic", // guess
	10029: "mac_latin2",   // guess
	10079: "mac_iceland",  // guess
	10081: "mac_turkish",  // guess
	32768: "mac_roman",
	32769: "cp1252",
}

// biffRecNameDict maps BIFF record codes to their names.
var biffRecNameDict = map[int]string{
	0x0000: "DIMENSIONS_B2",
	0x0001: "BLANK_B2",
	0x0002: "INTEGER_B2_ONLY",
	0x0003: "NUMBER_B2",
	0x0004: "LABEL_B2",
	0x0005: "BOOLERR_B2",
	0x0006: "FORMULA",
	0x0007: "STRING_B2",
	0x0008: "ROW_B2",
	0x0009: "BOF_B2",
	0x000A: "EOF",
	0x000B: "INDEX_B2_ONLY",
	0x000C: "CALCCOUNT",
	0x000D: "CALCMODE",
	0x000E: "PRECISION",
	0x000F: "REFMODE",
	0x0010: "DELTA",
	0x0011: "ITERATION",
	0x0012: "PROTECT",
	0x0013: "PASSWORD",
	0x0014: "HEADER",
	0x0015: "FOOTER",
	0x0016: "EXTERNCOUNT",
	0x0017: "EXTERNSHEET",
	0x0018: "NAME_B2,5+",
	0x0019: "WINDOWPROTECT",
	0x001A: "VERTICALPAGEBREAKS",
	0x001B: "HORIZONTALPAGEBREAKS",
	0x001C: "NOTE",
	0x001D: "SELECTION",
	0x001E: "FORMAT_B2-3",
	0x001F: "BUILTINFMTCOUNT_B2",
	0x0020: "COLUMNDEFAULT_B2_ONLY",
	0x0021: "ARRAY_B2_ONLY",
	0x0022: "DATEMODE",
	0x0023: "EXTERNNAME",
	0x0024: "COLWIDTH_B2_ONLY",
	0x0025: "DEFAULTROWHEIGHT_B2_ONLY",
	0x0026: "LEFTMARGIN",
	0x0027: "RIGHTMARGIN",
	0x0028: "TOPMARGIN",
	0x0029: "BOTTOMMARGIN",
	0x002A: "PRINTHEADERS",
	0x002B: "PRINTGRIDLINES",
	0x002F: "FILEPASS",
	0x0031: "FONT",
	0x0032: "FONT2_B2_ONLY",
	0x0036: "TABLEOP_B2",
	0x0037: "TABLEOP2_B2",
	0x003C: "CONTINUE",
	0x003D: "WINDOW1",
	0x003E: "WINDOW2_B2",
	0x0040: "BACKUP",
	0x0041: "PANE",
	0x0042: "CODEPAGE",
	0x0043: "XF_B2",
	0x0044: "IXFE_B2_ONLY",
	0x0045: "EFONT_B2_ONLY",
	0x004D: "PLS",
	0x0051: "DCONREF",
	0x0055: "DEFCOLWIDTH",
	0x0056: "BUILTINFMTCOUNT_B3-4",
	0x0059: "XCT",
	0x005A: "CRN",
	0x005B: "FILESHARING",
	0x005C: "WRITEACCESS",
	0x005D: "OBJECT",
	0x005E: "UNCALCED",
	0x005F: "SAVERECALC",
	0x0063: "OBJECTPROTECT",
	0x007D: "COLINFO",
	0x007E: "RK2_mythical_?",
	0x0080: "GUTS",
	0x0081: "WSBOOL",
	0x0082: "GRIDSET",
	0x0083: "HCENTER",
	0x0084: "VCENTER",
	0x0085: "BOUNDSHEET",
	0x0086: "WRITEPROT",
	0x008C: "COUNTRY",
	0x008D: "HIDEOBJ",
	0x008E: "SHEETSOFFSET",
	0x008F: "SHEETHDR",
	0x0090: "SORT",
	0x0092: "PALETTE",
	0x0099: "STANDARDWIDTH",
	0x009B: "FILTERMODE",
	0x009C: "FNGROUPCOUNT",
	0x009D: "AUTOFILTERINFO",
	0x009E: "AUTOFILTER",
	0x00A0: "SCL",
	0x00A1: "SETUP",
	0x00AB: "GCW",
	0x00BD: "MULRK",
	0x00BE: "MULBLANK",
	0x00C1: "MMS",
	0x00D6: "RSTRING",
	0x00DA: "BOOKBOOL",
	0x00DD: "SCENPROTECT",
	0x00E0: "XF",
	0x00E1: "INTERFACEHDR",
	0x00E2: "INTERFACEEND",
	0x00E5: "MERGEDCELLS",
	0x00E9: "BITMAP",
	0x00EB: "MSO_DRAWING_GROUP",
	0x00EC: "MSO_DRAWING",
	0x00ED: "MSO_DRAWING_SELECTION",
	0x00EF: "PHONETIC",
	0x00FC: "SST",
	0x00FD: "LABELSST",
	0x00FF: "EXTSST",
	0x013D: "TABID",
	0x015F: "LABELRANGES",
	0x0160: "USESELFS",
	0x0161: "DSF",
	0x01AE: "SUPBOOK",
	0x01AF: "PROTECTIONREV4",
	0x01B0: "CONDFMT",
	0x01B1: "CF",
	0x01B2: "DVAL",
	0x01B6: "TXO",
	0x01B7: "REFRESHALL",
	0x01B8: "HLINK",
	0x01BC: "PASSWORDREV4",
	0x01BE: "DV",
	0x01C0: "XL9FILE",
	0x01C1: "RECALCID",
	0x0200: "DIMENSIONS",
	0x0201: "BLANK",
	0x0203: "NUMBER",
	0x0204: "LABEL",
	0x0205: "BOOLERR",
	0x0206: "FORMULA_B3",
	0x0207: "STRING",
	0x0208: "ROW",
	0x0209: "BOF",
	0x020B: "INDEX_B3+",
	0x0218: "NAME",
	0x0221: "ARRAY",
	0x0223: "EXTERNNAME_B3-4",
	0x0225: "DEFAULTROWHEIGHT",
	0x0231: "FONT_B3B4",
	0x0236: "TABLEOP",
	0x023E: "WINDOW2",
	0x0243: "XF_B3",
	0x027E: "RK",
	0x0293: "STYLE",
	0x0406: "FORMULA_B4",
	0x0409: "BOF",
	0x041E: "FORMAT",
	0x0443: "XF_B4",
	0x04BC: "SHRFMLA",
	0x0800: "QUICKTIP",
	0x0809: "BOF",
	0x0862: "SHEETLAYOUT",
	0x0867: "SHEETPROTECTION",
	0x0868: "RANGEPROTECTION",
}

// BaseObject is a base type for most objects in the package.
// It provides a common dump method for debugging.
type BaseObject struct {
	// This is a placeholder for now
}

// _repr_these specifies which attributes should be shown in repr output.
var _repr_these = map[string]bool{}

// Dump writes debugging information about the object to the given writer.
func (b *BaseObject) Dump(w io.Writer, header, footer string, indent int) {
	if header != "" {
		fmt.Fprintf(w, "%s\n", header)
	}

	v := reflect.ValueOf(b).Elem()
	t := v.Type()

	pad := ""
	for i := 0; i < indent; i++ {
		pad += " "
	}

	// Collect field information
	type fieldInfo struct {
		name  string
		value reflect.Value
	}
	var fields []fieldInfo

	for i := 0; i < v.NumField(); i++ {
		field := v.Field(i)
		fieldType := t.Field(i)

		// Skip unexported fields
		if !field.CanInterface() {
			continue
		}

		fields = append(fields, fieldInfo{
			name:  fieldType.Name,
			value: field,
		})
	}

	// Sort fields by name
	for i := 0; i < len(fields); i++ {
		for j := i + 1; j < len(fields); j++ {
			if fields[i].name > fields[j].name {
				fields[i], fields[j] = fields[j], fields[i]
			}
		}
	}

	for _, field := range fields {
		attr := field.name
		value := field.value

		// Check if the value has a Dump method
		if value.Kind() == reflect.Ptr && !value.IsNil() {
			dumpMethod := value.MethodByName("Dump")
			if dumpMethod.IsValid() {
				dumpMethod.Call([]reflect.Value{
					reflect.ValueOf(w),
					reflect.ValueOf(fmt.Sprintf("%s%s (%s object):", pad, attr, value.Elem().Type().Name())),
					reflect.ValueOf(""),
					reflect.ValueOf(indent + 4),
				})
				continue
			}
		} else if value.Kind() == reflect.Interface && !value.IsNil() {
			dumpMethod := value.MethodByName("Dump")
			if dumpMethod.IsValid() {
				dumpMethod.Call([]reflect.Value{
					reflect.ValueOf(w),
					reflect.ValueOf(fmt.Sprintf("%s%s (%s object):", pad, attr, value.Elem().Type().Name())),
					reflect.ValueOf(""),
					reflect.ValueOf(indent + 4),
				})
				continue
			}
		}

		// Check if it's a list or dict type
		isContainer := false
		switch value.Kind() {
		case reflect.Slice, reflect.Array, reflect.Map:
			isContainer = true
		}

		if !_repr_these[attr] && isContainer {
			typeName := value.Type().String()
			length := 0
			switch value.Kind() {
			case reflect.Slice, reflect.Array:
				length = value.Len()
			case reflect.Map:
				length = value.Len()
			}
			fmt.Fprintf(w, "%s%s: %s, len = %d\n", pad, attr, typeName, length)
		} else {
			fmt.Fprintf(w, "%s%s: %v\n", pad, attr, value.Interface())
		}
	}

	if footer != "" {
		fmt.Fprintf(w, "%s\n", footer)
	}
}

// fprintf is equivalent to Python's fprintf function
func fprintf(w io.Writer, format string, args ...interface{}) {
	fmt.Fprintf(w, format, args...)
}

// upkbits unpacks bit fields from src into tgt_obj using the manifest.
// manifest is a slice of [shift, mask, fieldName] tuples.
// This function uses reflection to set fields on the target object.
func upkbits(tgt_obj interface{}, src uint32, manifest [][3]interface{}) {
	if tgt_obj == nil {
		return
	}
	if tgt_map, ok := tgt_obj.(map[string]interface{}); ok {
		for _, item := range manifest {
			n := item[0].(int)
			mask := item[1].(uint32)
			attr := item[2].(string)
			tgt_map[attr] = (src & mask) >> uint32(n)
		}
		return
	}
	setBitsOnStruct(tgt_obj, src, manifest, false)
}

// upkbitsL is like upkbits but ensures the result is an int.
func upkbitsL(tgt_obj interface{}, src uint32, manifest [][3]interface{}) {
	if tgt_obj == nil {
		return
	}
	if tgt_map, ok := tgt_obj.(map[string]interface{}); ok {
		for _, item := range manifest {
			n := item[0].(int)
			mask := item[1].(uint32)
			attr := item[2].(string)
			tgt_map[attr] = int((src & mask) >> uint32(n))
		}
		return
	}
	setBitsOnStruct(tgt_obj, src, manifest, true)
}

func setBitsOnStruct(tgtObj interface{}, src uint32, manifest [][3]interface{}, forceInt bool) {
	v := reflect.ValueOf(tgtObj)
	if v.Kind() == reflect.Ptr {
		if v.IsNil() {
			return
		}
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return
	}
	for _, item := range manifest {
		shift := item[0].(int)
		mask := item[1].(uint32)
		attr := item[2].(string)
		val := (src & mask) >> uint32(shift)
		field := v.FieldByName(attr)
		if !field.IsValid() || !field.CanSet() {
			continue
		}
		switch field.Kind() {
		case reflect.Bool:
			field.SetBool(val != 0)
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			field.SetInt(int64(val))
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			if forceInt {
				field.SetUint(uint64(val))
			} else {
				field.SetUint(uint64(val))
			}
		}
	}
}

// HexCharDump dumps a byte slice in hex and character format.
func HexCharDump(strg []byte, ofs, dlen, base int, fout io.Writer, unnumbered bool) {
	endpos := ofs + dlen
	if endpos > len(strg) {
		endpos = len(strg)
	}
	pos := ofs
	numbered := !unnumbered
	numPrefix := ""

	for pos < endpos {
		endsub := pos + 16
		if endsub > endpos {
			endsub = endpos
		}
		substrg := strg[pos:endsub]
		lensub := len(substrg)

		if lensub <= 0 {
			fmt.Fprintf(fout, "??? hex_char_dump: ofs=%d dlen=%d base=%d -> endpos=%d pos=%d endsub=%d\n",
				ofs, dlen, base, endpos, pos, endsub)
			break
		}

		// Build hex string
		hexd := ""
		for i := 0; i < lensub; i++ {
			if i > 0 {
				hexd += " "
			}
			hexd += fmt.Sprintf("%02x", substrg[i])
		}
		// Pad to 48 characters
		for len(hexd) < 48 {
			hexd += " "
		}

		// Build character string
		chard := ""
		for i := 0; i < lensub; i++ {
			c := substrg[i]
			if c == 0 {
				chard += "~"
			} else if c < 32 || c > 126 {
				chard += "?"
			} else {
				chard += string(c)
			}
		}

		if numbered {
			numPrefix = fmt.Sprintf("%5d: ", base+pos-ofs)
		} else {
			numPrefix = ""
		}

		fmt.Fprintf(fout, "%s     %-48s %s\n", numPrefix, hexd, chard)
		pos = endsub
	}
}

// unpack_string unpacks a string from binary data.
func unpack_string(data []byte, pos int, encoding string, lenlen int) string {
	if lenlen < 1 {
		lenlen = 1
	}
	if lenlen > 2 {
		lenlen = 2
	}

	var nchars uint16
	if lenlen == 1 {
		nchars = uint16(data[pos])
		pos += 1
	} else {
		nchars = binary.LittleEndian.Uint16(data[pos : pos+2])
		pos += 2
	}

	var dataLen int
	switch encoding {
	case "utf_16_le":
		dataLen = int(nchars) * 2
	default:
		dataLen = int(nchars)
	}

	if pos+dataLen > len(data) {
		return ""
	}

	strData := data[pos : pos+dataLen]
	switch encoding {
	case "utf_16_le":
		utf16Data := make([]uint16, len(strData)/2)
		for i := 0; i < len(utf16Data); i++ {
			utf16Data[i] = binary.LittleEndian.Uint16(strData[i*2:])
		}
		runes := utf16.Decode(utf16Data)
		return string(runes)
	case "latin_1", "cp1252":
		return string(strData)
	default:
		return string(strData)
	}
}

// unpack_string_update_pos unpacks a string and updates the position.
func unpack_string_update_pos(data []byte, pos int, encoding string, lenlen int, known_len int) (string, int) {
	var nchars int
	if known_len >= 0 {
		nchars = known_len
	} else {
		if lenlen == 1 {
			nchars = int(data[pos])
		} else {
			nchars = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		}
		pos += lenlen
	}

	if pos+nchars > len(data) {
		return "", pos
	}

	newpos := pos + nchars
	str := unpack_string(data, pos-lenlen, encoding, lenlen)
	return str, newpos
}

// unpack_unicode unpacks a unicode string from binary data.
func unpack_unicode(data []byte, pos int, lenlen int) string {
	if lenlen < 2 {
		lenlen = 2
	}
	if lenlen > 2 {
		lenlen = 2
	}

	nchars := int(binary.LittleEndian.Uint16(data[pos : pos+lenlen]))
	if nchars == 0 {
		return ""
	}
	pos += lenlen

	options := data[pos]
	pos++

	// Skip rich text and phonetic info if present
	if options&0x08 != 0 { // richtext
		pos += 2
	}
	if options&0x04 != 0 { // phonetic
		pos += 4
	}

	if options&0x01 != 0 { // compressed UTF-16
		if pos+2*nchars > len(data) {
			return ""
		}
		utf16Data := make([]uint16, nchars)
		for i := 0; i < nchars; i++ {
			utf16Data[i] = binary.LittleEndian.Uint16(data[pos+i*2:])
		}
		runes := utf16.Decode(utf16Data)
		return string(runes)
	} else { // compressed (latin_1)
		if pos+nchars > len(data) {
			return ""
		}
		return string(data[pos : pos+nchars])
	}
}

// unpack_unicode_update_pos unpacks a unicode string and updates the position.
func unpack_unicode_update_pos(data []byte, pos int, lenlen int, known_len int) (string, int) {
	if known_len >= 0 {
		nchars := known_len
		if nchars == 0 && len(data) <= pos {
			return "", pos
		}
		options := data[pos]
		pos++

		phonetic := options&0x04 != 0
		richtext := options&0x08 != 0

		if richtext {
			if pos+2 > len(data) {
				return "", pos
			}
			rt := binary.LittleEndian.Uint16(data[pos:])
			pos += 2
			_ = rt // unused in Python version too
		}

		if phonetic {
			if pos+4 > len(data) {
				return "", pos
			}
			sz := binary.LittleEndian.Uint32(data[pos:])
			pos += 4
			_ = sz // unused in Python version too
		}

		var str string
		var newpos int
		if options&0x01 != 0 { // uncompressed UTF-16
			if pos+2*nchars > len(data) {
				return "", pos
			}
			utf16Data := make([]uint16, nchars)
			for i := 0; i < nchars; i++ {
				utf16Data[i] = binary.LittleEndian.Uint16(data[pos+i*2:])
			}
			runes := utf16.Decode(utf16Data)
			str = string(runes)
			newpos = pos + 2*nchars
		} else { // compressed
			if pos+nchars > len(data) {
				return "", pos
			}
			str = string(data[pos : pos+nchars])
			newpos = pos + nchars
		}

		if richtext {
			newpos += 4 * int(binary.LittleEndian.Uint16(data[pos-2:]))
		}
		if phonetic {
			newpos += int(binary.LittleEndian.Uint32(data[pos-4:]))
		}

		return str, newpos
	} else {
		nchars := int(binary.LittleEndian.Uint16(data[pos : pos+lenlen]))
		pos += lenlen
		if nchars == 0 && len(data) <= pos {
			return "", pos
		}
		options := data[pos]
		pos++

		phonetic := options&0x04 != 0
		richtext := options&0x08 != 0

		if richtext {
			rt := binary.LittleEndian.Uint16(data[pos:])
			pos += 2
			_ = rt
		}

		if phonetic {
			sz := binary.LittleEndian.Uint32(data[pos:])
			pos += 4
			_ = sz
		}

		var str string
		var newpos int
		if options&0x01 != 0 { // uncompressed UTF-16
			utf16Data := make([]uint16, nchars)
			for i := 0; i < nchars; i++ {
				utf16Data[i] = binary.LittleEndian.Uint16(data[pos+i*2:])
			}
			runes := utf16.Decode(utf16Data)
			str = string(runes)
			newpos = pos + 2*nchars
		} else { // compressed
			str = string(data[pos : pos+nchars])
			newpos = pos + nchars
		}

		if richtext {
			newpos += 4 * int(binary.LittleEndian.Uint16(data[pos-2:]))
		}
		if phonetic {
			newpos += int(binary.LittleEndian.Uint32(data[pos-4:]))
		}

		return str, newpos
	}
}

// unpack_cell_range_address_list_update_pos unpacks cell range addresses and updates position.
func unpack_cell_range_address_list_update_pos(output_list *[]CellRange, data []byte, pos int, biff_version int, addr_size int) int {
	if addr_size != 6 && addr_size != 8 {
		addr_size = 6
	}

	n := int(binary.LittleEndian.Uint16(data[pos:]))
	pos += 2

	if n > 0 {
		for i := 0; i < n; i++ {
			var ra, rb, ca, cb uint16
			if addr_size == 6 {
				ra = binary.LittleEndian.Uint16(data[pos:])
				rb = binary.LittleEndian.Uint16(data[pos+2:])
				ca = uint16(data[pos+4])
				cb = uint16(data[pos+5])
			} else { // addr_size == 8
				ra = binary.LittleEndian.Uint16(data[pos:])
				rb = binary.LittleEndian.Uint16(data[pos+2:])
				ca = binary.LittleEndian.Uint16(data[pos+4:])
				cb = binary.LittleEndian.Uint16(data[pos+6:])
			}
			*output_list = append(*output_list, CellRange{
				FirstRow: int(ra),
				LastRow:  int(rb) + 1,
				FirstCol: int(ca),
				LastCol:  int(cb) + 1,
			})
			pos += addr_size
		}
	}
	return pos
}

// CellRange represents a cell range with row and column bounds.
type CellRange struct {
	FirstRow int
	LastRow  int
	FirstCol int
	LastCol  int
}

// BiffDump dumps BIFF records from binary data.
func BiffDump(mem []byte, stream_offset, stream_len, base int, fout io.Writer, unnumbered bool) {
	pos := stream_offset
	stream_end := stream_offset + stream_len
	adj := base - stream_offset
	dummies := 0
	savpos := 0
	numbered := !unnumbered
	num_prefix := ""
	var length uint16

	for stream_end-pos >= 4 {
		rc := binary.LittleEndian.Uint16(mem[pos:])
		length = binary.LittleEndian.Uint16(mem[pos+2:])

		if rc == 0 && length == 0 {
			if pos+4 <= len(mem) && mem[pos+4:] != nil && len(mem[pos+4:]) >= stream_end-pos-4 {
				// Check if remaining data is all zeros
				all_zeros := true
				for i := pos; i < stream_end; i++ {
					if mem[i] != 0 {
						all_zeros = false
						break
					}
				}
				if all_zeros {
					dummies = stream_end - pos
					pos = stream_end
					break
				}
			}
			if dummies > 0 {
				dummies += 4
			} else {
				savpos = pos
				dummies = 4
			}
			pos += 4
		} else {
			if dummies > 0 {
				if numbered {
					num_prefix = fmt.Sprintf("%5d: ", adj+savpos)
				}
				fprintf(fout, "%s---- %d zero bytes skipped ----\n", num_prefix, dummies)
				dummies = 0
			}
			recname := biffRecNameDict[int(rc)]
			if recname == "" {
				recname = fmt.Sprintf("<UNKNOWN>")
			}
			if numbered {
				num_prefix = fmt.Sprintf("%5d: ", adj+pos)
			}
			fprintf(fout, "%s%04x %s len = %04x (%d)\n", num_prefix, rc, recname, length, length)
			pos += 4
			HexCharDump(mem, pos, int(length), adj+pos, fout, unnumbered)
			pos += int(length)
		}
	}

	if dummies > 0 {
		if numbered {
			num_prefix = fmt.Sprintf("%5d: ", adj+savpos)
		}
		fprintf(fout, "%s---- %d zero bytes skipped ----\n", num_prefix, dummies)
	}

	if pos < stream_end {
		if numbered {
			num_prefix = fmt.Sprintf("%5d: ", adj+pos)
		}
		fprintf(fout, "%s---- Misc bytes at end ----\n", num_prefix)
		HexCharDump(mem, pos, stream_end-pos, adj+pos, fout, unnumbered)
	} else if pos > stream_end {
		fprintf(fout, "Last dumped record has length (%d) that is too large\n", length)
	}
}

// BiffCountRecords counts BIFF records in binary data.
func BiffCountRecords(mem []byte, stream_offset, stream_len int, fout io.Writer) {
	pos := stream_offset
	stream_end := stream_offset + stream_len
	tally := make(map[string]int)

	for stream_end-pos >= 4 {
		rc := binary.LittleEndian.Uint16(mem[pos:])
		length := binary.LittleEndian.Uint16(mem[pos+2:])

		if rc == 0 && length == 0 {
			if pos+4 <= len(mem) && mem[pos+4:] != nil && len(mem[pos+4:]) >= stream_end-pos-4 {
				// Check if remaining data is all zeros
				all_zeros := true
				for i := pos; i < stream_end; i++ {
					if mem[i] != 0 {
						all_zeros = false
						break
					}
				}
				if all_zeros {
					break
				}
			}
			recname := "<Dummy (zero)>"
			if count, ok := tally[recname]; ok {
				tally[recname] = count + 1
			} else {
				tally[recname] = 1
			}
		} else {
			recname := biffRecNameDict[int(rc)]
			if recname == "" {
				recname = fmt.Sprintf("Unknown_0x%04X", rc)
			}
			if count, ok := tally[recname]; ok {
				tally[recname] = count + 1
			} else {
				tally[recname] = 1
			}
		}
		pos += int(length) + 4

		// Prevent infinite loops
		if length == 0 {
			break
		}
	}

	// Sort and print results
	type recordCount struct {
		name  string
		count int
	}
	var records []recordCount
	for name, count := range tally {
		records = append(records, recordCount{name, count})
	}

	// Simple sort by count descending, then by name
	for i := 0; i < len(records); i++ {
		for j := i + 1; j < len(records); j++ {
			if records[i].count < records[j].count ||
				(records[i].count == records[j].count && records[i].name > records[j].name) {
				records[i], records[j] = records[j], records[i]
			}
		}
	}

	for _, record := range records {
		fprintf(fout, "%8d %s\n", record.count, record.name)
	}
}

// hex_char_dump dumps hex and character representation of binary data
func hex_char_dump(data []byte, ofs int, dlen int, base int, fout io.Writer, unnumbered bool) {
	endpos := ofs + dlen
	if endpos > len(data) {
		endpos = len(data)
	}
	pos := ofs
	numbered := !unnumbered

	for pos < endpos {
		endsub := pos + 16
		if endsub > endpos {
			endsub = endpos
		}
		substrg := data[pos:endsub]
		lensub := len(substrg)

		if lensub <= 0 {
			fmt.Fprintf(fout, "??? hex_char_dump: ofs=%d dlen=%d base=%d -> endpos=%d pos=%d endsub=%d\n",
				ofs, dlen, base, endpos, pos, endsub)
			break
		}

		// Build hex string
		hexd := ""
		for i := 0; i < lensub; i++ {
			if i > 0 {
				hexd += " "
			}
			hexd += fmt.Sprintf("%02x", substrg[i])
		}
		// Pad to 48 characters (3 chars per byte: "xx ")
		for len(hexd) < 48 {
			hexd += " "
		}

		// Build character string
		chard := ""
		for i := 0; i < lensub; i++ {
			c := substrg[i]
			if c == 0 {
				chard += "~"
			} else if c < 32 || c > 126 {
				chard += "?"
			} else {
				chard += string(c)
			}
		}

		numPrefix := ""
		if numbered {
			numPrefix = fmt.Sprintf("%5d: ", base+pos-ofs)
		}

		fmt.Fprintf(fout, "%s     %-48s %s\n", numPrefix, hexd, chard)
		pos = endsub
	}
}
