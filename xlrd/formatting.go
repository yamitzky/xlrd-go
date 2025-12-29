package xlrd

import (
	"fmt"
	"regexp"
	"strings"
)

// Font represents font information.
type Font struct {
	BaseObject

	// Name is the font name.
	Name string

	// Bold indicates if the font is bold.
	Bold bool

	// Italic indicates if the font is italic.
	Italic bool

	// Underline indicates the underline style.
	Underline int

	// Escapement indicates the escapement style.
	Escapement int

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

	// FontIndex is the index into the font list.
	FontIndex int

	// FormatKey is the format key.
	FormatKey int

	// Locked indicates if the cell is locked.
	Locked bool

	// Hidden indicates if the cell is hidden.
	Hidden bool

	// ParentStyleIndex is the parent style index.
	ParentStyleIndex int

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

	// IndentLevel is the indent level.
	IndentLevel int

	// ShrinkToFit indicates if text should shrink to fit.
	ShrinkToFit bool

	// WrapText indicates if text should wrap.
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

	// LeftColourIndex is the left border colour index.
	LeftColourIndex int

	// RightColourIndex is the right border colour index.
	RightColourIndex int

	// TopColourIndex is the top border colour index.
	TopColourIndex int

	// BottomColourIndex is the bottom border colour index.
	BottomColourIndex int
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
	"0.00E+00":  true,
	"##0.0E+0":  true,
	"General":   true,
	"GENERAL":   true,
	"general":   true,
	"@":         true,
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
	for _, c := range reducedFmt {
		if count, ok := dateCharDict[c]; ok {
			dateCount += count
		} else if count, ok := numCharDict[c]; ok {
			numCount += count
		}
	}

	// Date format if it has date chars and no number chars
	return dateCount > 0 && numCount == 0
}

// NearestColourIndex finds the nearest colour index for a given RGB value.
// Uses Euclidean distance. So far used only for pre-BIFF8 WINDOW2 record.
func NearestColourIndex(colourMap map[int][3]int, rgb [3]int, debug int) int {
	bestMetric := 3 * 256 * 256
	bestColourx := 0

	for colourx, candRGB := range colourMap {
		if candRGB == [3]int{} { // nil equivalent
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
