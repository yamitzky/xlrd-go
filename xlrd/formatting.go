package xlrd

import (
	"fmt"
	"reflect"
	"regexp"
	"strings"
)

var unknownRGB = [3]int{-1, -1, -1}

var excelDefaultPaletteB5 = [][3]int{
	{0, 0, 0}, {255, 255, 255}, {255, 0, 0}, {0, 255, 0},
	{0, 0, 255}, {255, 255, 0}, {255, 0, 255}, {0, 255, 255},
	{128, 0, 0}, {0, 128, 0}, {0, 0, 128}, {128, 128, 0},
	{128, 0, 128}, {0, 128, 128}, {192, 192, 192}, {128, 128, 128},
	{153, 153, 255}, {153, 51, 102}, {255, 255, 204}, {204, 255, 255},
	{102, 0, 102}, {255, 128, 128}, {0, 102, 204}, {204, 204, 255},
	{0, 0, 128}, {255, 0, 255}, {255, 255, 0}, {0, 255, 255},
	{128, 0, 128}, {128, 0, 0}, {0, 128, 128}, {0, 0, 255},
	{0, 204, 255}, {204, 255, 255}, {204, 255, 204}, {255, 255, 153},
	{153, 204, 255}, {255, 153, 204}, {204, 153, 255}, {227, 227, 227},
	{51, 102, 255}, {51, 204, 204}, {153, 204, 0}, {255, 204, 0},
	{255, 153, 0}, {255, 102, 0}, {102, 102, 153}, {150, 150, 150},
	{0, 51, 102}, {51, 153, 102}, {0, 51, 0}, {51, 51, 0},
	{153, 51, 0}, {153, 51, 102}, {51, 51, 153}, {51, 51, 51},
}

var excelDefaultPaletteB2 = excelDefaultPaletteB5[:16]

var excelDefaultPaletteB8 = [][3]int{
	{0, 0, 0}, {255, 255, 255}, {255, 0, 0}, {0, 255, 0},
	{0, 0, 255}, {255, 255, 0}, {255, 0, 255}, {0, 255, 255},
	{128, 0, 0}, {0, 128, 0}, {0, 0, 128}, {128, 128, 0},
	{128, 0, 128}, {0, 128, 128}, {192, 192, 192}, {128, 128, 128},
	{153, 153, 255}, {153, 51, 102}, {255, 255, 204}, {204, 255, 255},
	{102, 0, 102}, {255, 128, 128}, {0, 102, 204}, {204, 204, 255},
	{0, 0, 128}, {255, 0, 255}, {255, 255, 0}, {0, 255, 255},
	{128, 0, 128}, {128, 0, 0}, {0, 128, 128}, {0, 0, 255},
	{0, 204, 255}, {204, 255, 255}, {204, 255, 204}, {255, 255, 153},
	{153, 204, 255}, {255, 153, 204}, {204, 153, 255}, {255, 204, 153},
	{51, 102, 255}, {51, 204, 204}, {153, 204, 0}, {255, 204, 0},
	{255, 153, 0}, {255, 102, 0}, {102, 102, 153}, {150, 150, 150},
	{0, 51, 102}, {51, 153, 102}, {0, 51, 0}, {51, 51, 0},
	{153, 51, 0}, {153, 51, 102}, {51, 51, 153}, {51, 51, 51},
}

var defaultPalette = map[int][][3]int{
	80: excelDefaultPaletteB8,
	70: excelDefaultPaletteB5,
	50: excelDefaultPaletteB5,
	45: excelDefaultPaletteB2,
	40: excelDefaultPaletteB2,
	30: excelDefaultPaletteB2,
	21: excelDefaultPaletteB2,
	20: excelDefaultPaletteB2,
}

var builtInStyleNames = []string{
	"Normal",
	"RowLevel_",
	"ColLevel_",
	"Comma",
	"Currency",
	"Percent",
	"Comma [0]",
	"Currency [0]",
	"Hyperlink",
	"Followed Hyperlink",
}

var stdFormatStrings = map[int]string{
	0x00: "General",
	0x01: "0",
	0x02: "0.00",
	0x03: "#,##0",
	0x04: "#,##0.00",
	0x05: "$#,##0_);($#,##0)",
	0x06: "$#,##0_);[Red]($#,##0)",
	0x07: "$#,##0.00_);($#,##0.00)",
	0x08: "$#,##0.00_);[Red]($#,##0.00)",
	0x09: "0%",
	0x0a: "0.00%",
	0x0b: "0.00E+00",
	0x0c: "# ?/?",
	0x0d: "# ??/??",
	0x0e: "m/d/yy",
	0x0f: "d-mmm-yy",
	0x10: "d-mmm",
	0x11: "mmm-yy",
	0x12: "h:mm AM/PM",
	0x13: "h:mm:ss AM/PM",
	0x14: "h:mm",
	0x15: "h:mm:ss",
	0x16: "m/d/yy h:mm",
	0x25: "#,##0_);(#,##0)",
	0x26: "#,##0_);[Red](#,##0)",
	0x27: "#,##0.00_);(#,##0.00)",
	0x28: "#,##0.00_);[Red](#,##0.00)",
	0x29: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
	0x2a: "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
	0x2b: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
	0x2c: "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
	0x2d: "mm:ss",
	0x2e: "[h]:mm:ss",
	0x2f: "mm:ss.0",
	0x30: "##0.0E+0",
	0x31: "@",
}

var stdFormatCodeTypes = func() map[int]int {
	ranges := [][3]int{
		{0, 0, FGE},
		{1, 13, FNU},
		{14, 22, FDT},
		{27, 36, FDT},
		{37, 44, FNU},
		{45, 47, FDT},
		{48, 48, FNU},
		{49, 49, FTX},
		{50, 58, FDT},
		{59, 62, FNU},
		{67, 70, FNU},
		{71, 81, FDT},
	}
	m := make(map[int]int)
	for _, r := range ranges {
		for x := r[0]; x <= r[1]; x++ {
			m[x] = r[2]
		}
	}
	return m
}()

var cellTypeFromFormatType = map[int]int{
	FNU: XL_CELL_NUMBER,
	FUN: XL_CELL_NUMBER,
	FGE: XL_CELL_NUMBER,
	FDT: XL_CELL_DATE,
	FTX: XL_CELL_NUMBER,
}

// Font represents font information.
type Font struct {
	BaseObject

	// FontIndex is the index used to refer to this font record.
	FontIndex int

	// Name is the font name.
	Name string

	// Bold indicates if the font is bold.
	Bold bool

	// Italic indicates if the font is italic.
	Italic bool

	// Underlined indicates if the font is underlined.
	Underlined bool

	// Underline indicates the underline style.
	Underline int

	// Escapement indicates the escapement style.
	Escapement int

	// StruckOut indicates if the font is struck out.
	StruckOut bool

	// ColourIndex is the colour index.
	ColourIndex int

	// Height is the font height in twips (1/20 of a point).
	Height int

	// Weight is the font weight.
	Weight int

	// Family is the font family.
	Family int

	// CharacterSet is the character set.
	CharacterSet int

	// Outline indicates outline style (Macintosh only).
	Outline bool

	// Shadow indicates shadow style (Macintosh only).
	Shadow bool
}

// Format represents number format information.
type Format struct {
	BaseObject

	// FormatKey is the format key.
	FormatKey int

	// Type is the format type (FUN, FDT, FNU, FGE, FTX).
	Type int

	// FormatString is the format string.
	FormatString string
}

// XF represents extended format information.
type XF struct {
	BaseObject

	// IsStyle is 0 for cell XF, 1 for style XF.
	IsStyle int

	// ParentStyleIndex is the parent style index.
	ParentStyleIndex int

	// FormatFlag controls format inheritance.
	FormatFlag int

	// FontFlag controls font inheritance.
	FontFlag int

	// AlignmentFlag controls alignment inheritance.
	AlignmentFlag int

	// BorderFlag controls border inheritance.
	BorderFlag int

	// BackgroundFlag controls background inheritance.
	BackgroundFlag int

	// ProtectionFlag controls protection inheritance.
	ProtectionFlag int

	// Lotus123Prefix is the Lotus 1-2-3 prefix flag (meaning unknown).
	Lotus123Prefix int

	// XFIndex is the index into the XF list.
	XFIndex int

	// FontIndex is the index into the font list.
	FontIndex int

	// FormatKey is the format key.
	FormatKey int

	// Locked indicates if the cell is locked.
	Locked bool

	// Hidden indicates if the cell is hidden.
	Hidden bool

	// Alignment is the alignment information.
	Alignment *XFAlignment

	// Border is the border information.
	Border *XFBorder

	// Background is the background information.
	Background *XFBackground

	// Protection is the protection information.
	Protection *XFProtection
}

// XFAlignment represents alignment information.
type XFAlignment struct {
	BaseObject

	// Horizontal is the horizontal alignment.
	Horizontal int

	// Vertical is the vertical alignment.
	Vertical int

	// Rotation is the rotation angle.
	Rotation int

	// TextWrapped indicates if text should wrap.
	TextWrapped bool

	// IndentLevel is the indent level.
	IndentLevel int

	// ShrinkToFit indicates if text should shrink to fit.
	ShrinkToFit bool

	// TextDirection indicates the text direction.
	TextDirection int

	// HorAlign is the horizontal alignment (Python-compatible name).
	HorAlign int

	// VertAlign is the vertical alignment (Python-compatible name).
	VertAlign int

	// WrapText indicates if text should wrap (legacy alias).
	WrapText bool
}

// XFBorder represents border information.
type XFBorder struct {
	BaseObject

	// Left is the left border style.
	Left int

	// Right is the right border style.
	Right int

	// Top is the top border style.
	Top int

	// Bottom is the bottom border style.
	Bottom int

	// DiagLineStyle is the diagonal line style.
	DiagLineStyle int

	// LeftColourIndex is the left border colour index.
	LeftColourIndex int

	// RightColourIndex is the right border colour index.
	RightColourIndex int

	// TopColourIndex is the top border colour index.
	TopColourIndex int

	// BottomColourIndex is the bottom border colour index.
	BottomColourIndex int

	// DiagColourIndex is the diagonal border colour index.
	DiagColourIndex int

	// LeftLineStyle is the left border line style.
	LeftLineStyle int

	// RightLineStyle is the right border line style.
	RightLineStyle int

	// TopLineStyle is the top border line style.
	TopLineStyle int

	// BottomLineStyle is the bottom border line style.
	BottomLineStyle int

	// DiagDown indicates a diagonal from top-left to bottom-right.
	DiagDown int

	// DiagUp indicates a diagonal from bottom-left to top-right.
	DiagUp int
}

// XFBackground represents background information.
type XFBackground struct {
	BaseObject

	// FillPattern is the fill pattern.
	FillPattern int

	// PatternColourIndex is the pattern colour index.
	PatternColourIndex int

	// BackgroundColourIndex is the background colour index.
	BackgroundColourIndex int
}

// XFProtection represents protection information.
type XFProtection struct {
	BaseObject

	// CellLocked indicates if the cell is locked.
	CellLocked bool

	// FormulaHidden indicates if the formula is hidden.
	FormulaHidden bool
}

// Equal compares two Font values.
func (f *Font) Equal(other *Font) bool {
	return reflect.DeepEqual(f, other)
}

// NotEqual compares two Font values for inequality.
func (f *Font) NotEqual(other *Font) bool {
	return !f.Equal(other)
}

// Equal compares two Format values.
func (f *Format) Equal(other *Format) bool {
	return reflect.DeepEqual(f, other)
}

// NotEqual compares two Format values for inequality.
func (f *Format) NotEqual(other *Format) bool {
	return !f.Equal(other)
}

// Equal compares two XF values.
func (x *XF) Equal(other *XF) bool {
	return reflect.DeepEqual(x, other)
}

// NotEqual compares two XF values for inequality.
func (x *XF) NotEqual(other *XF) bool {
	return !x.Equal(other)
}

// Date and number character dictionaries for format string analysis
var dateCharDict = map[rune]int{
	'y': 5, 'Y': 5, 'm': 5, 'M': 5, 'd': 5, 'D': 5, 'h': 5, 'H': 5, 's': 5, 'S': 5,
}

var skipCharDict = map[rune]bool{
	'$': true, '-': true, '+': true, '/': true, '(': true, ')': true, ':': true, ' ': true,
}

var numCharDict = map[rune]int{
	'0': 5, '#': 5, '?': 5,
}

var nonDateFormats = map[string]bool{
	"0.00E+00": true,
	"##0.0E+0": true,
	"General":  true,
	"GENERAL":  true,
	"general":  true,
	"@":        true,
}

func defaultPaletteForBiff(biffVersion int) [][3]int {
	if palette, ok := defaultPalette[biffVersion]; ok {
		return palette
	}
	return excelDefaultPaletteB8
}

func initialiseColourMap(book *Book) {
	book.ColourMap = make(map[int][3]int)
	book.ColourIndexesUsed = make(map[int]bool)
	if !book.formattingInfo {
		return
	}
	for i := 0; i < 8; i++ {
		book.ColourMap[i] = excelDefaultPaletteB8[i]
	}
	dpal := defaultPaletteForBiff(book.BiffVersion)
	for i := 0; i < len(dpal); i++ {
		book.ColourMap[i+8] = dpal[i]
	}
	ndpal := len(dpal)
	book.ColourMap[ndpal+8] = unknownRGB
	book.ColourMap[ndpal+9] = unknownRGB
	book.ColourMap[0x51] = unknownRGB
	book.ColourMap[0x7FFF] = unknownRGB
}

func fillInStandardFormats(book *Book) {
	for formatCode, formatType := range stdFormatCodeTypes {
		if _, ok := book.FormatMap[formatCode]; ok {
			continue
		}
		fmtString := stdFormatStrings[formatCode]
		format := &Format{
			FormatKey:    formatCode,
			Type:         formatType,
			FormatString: fmtString,
		}
		book.FormatMap[formatCode] = format
	}
}

func checkColourIndexesInObj(book *Book, obj interface{}, origIndex int) {
	if book.ColourIndexesUsed == nil {
		book.ColourIndexesUsed = make(map[int]bool)
	}
	v := reflect.ValueOf(obj)
	if !v.IsValid() {
		return
	}
	if v.Kind() == reflect.Ptr {
		if v.IsNil() {
			return
		}
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return
	}
	t := v.Type()
	for i := 0; i < v.NumField(); i++ {
		field := v.Field(i)
		fieldType := t.Field(i)
		fieldName := fieldType.Name
		if strings.Contains(strings.ToLower(fieldName), "colourindex") {
			if !field.IsValid() {
				continue
			}
			val := 0
			switch field.Kind() {
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				val = int(field.Int())
			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				val = int(field.Uint())
			default:
				continue
			}
			if _, ok := book.ColourMap[val]; ok {
				book.ColourIndexesUsed[val] = true
				continue
			}
			if book.verbosity > 0 {
				oname := t.Name()
				fmt.Fprintf(book.logfile, "*** xf #%d : %s.%s =  0x%04x (unknown)\n", origIndex, oname, fieldName, val)
			}
			continue
		}
		switch field.Kind() {
		case reflect.Struct:
			checkColourIndexesInObj(book, field.Addr().Interface(), origIndex)
		case reflect.Ptr:
			if !field.IsNil() && field.Elem().Kind() == reflect.Struct {
				checkColourIndexesInObj(book, field.Interface(), origIndex)
			}
		}
	}
}

// IsDateFormatString checks if a format string represents a date format.
func IsDateFormatString(book *Book, formatStr string) bool {
	// Heuristics:
	// Ignore "text" and [stuff in square brackets].
	// Handle backslashed-escaped chars properly.
	// Date formats have one or more of ymdhs (caseless) in them.
	// Numeric formats have # and 0.

	state := 0
	var s strings.Builder

	for _, c := range formatStr {
		if state == 0 {
			if c == '"' {
				state = 1
			} else if c == '\\' || c == '_' || c == '*' {
				state = 2
			} else if skipCharDict[c] {
				// skip
			} else {
				s.WriteRune(c)
			}
		} else if state == 1 {
			if c == '"' {
				state = 0
			}
		} else if state == 2 {
			// Ignore char after backslash, underscore or asterisk
			state = 0
		}
	}

	reducedFmt := s.String()
	if book.verbosity >= 4 {
		if book.logfile != nil {
			fmt.Fprintf(book.logfile, "is_date_format_string: reduced format is %s\n", reducedFmt)
		} else {
			fmt.Printf("is_date_format_string: reduced format is %s\n", reducedFmt)
		}
	}

	// Remove bracketed expressions like [h], [m], etc.
	re := regexp.MustCompile(`\[.*?\]`)
	reducedFmt = re.ReplaceAllString(reducedFmt, "")

	if nonDateFormats[reducedFmt] {
		return false
	}

	dateCount := 0
	numCount := 0
	gotSep := false
	separator := ';'
	for _, c := range reducedFmt {
		if count, ok := dateCharDict[c]; ok {
			dateCount += count
		} else if count, ok := numCharDict[c]; ok {
			numCount += count
		} else if c == separator {
			gotSep = true
		}
	}

	if dateCount > 0 && numCount == 0 {
		return true
	}
	if numCount > 0 && dateCount == 0 {
		return false
	}
	if dateCount > 0 {
		if book.verbosity >= 1 {
			fmt.Fprintf(book.logfile,
				"WARNING *** is_date_format: ambiguous d=%d n=%d fmt=%q\n",
				dateCount, numCount, formatStr)
		}
	}
	if !gotSep && dateCount == 0 && book.verbosity >= 1 {
		fmt.Fprintf(book.logfile, "WARNING *** format %q produces constant result\n", formatStr)
	}
	return dateCount > numCount
}

// NearestColourIndex finds the nearest colour index for a given RGB value.
// Uses Euclidean distance. So far used only for pre-BIFF8 WINDOW2 record.
func NearestColourIndex(colourMap map[int][3]int, rgb [3]int, debug int) int {
	bestMetric := 3 * 256 * 256
	bestColourx := 0

	for colourx, candRGB := range colourMap {
		if candRGB[0] < 0 || candRGB[1] < 0 || candRGB[2] < 0 {
			continue
		}

		metric := 0
		for i := 0; i < 3; i++ {
			diff := rgb[i] - candRGB[i]
			metric += diff * diff
		}

		if metric < bestMetric {
			bestMetric = metric
			bestColourx = colourx
			if metric == 0 {
				break
			}
		}
	}

	if debug > 0 {
		fmt.Printf("nearest_colour_index for %v is %d -> %v; best_metric is %d\n",
			rgb, bestColourx, colourMap[bestColourx], bestMetric)
	}

	return bestColourx
}
