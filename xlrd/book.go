package xlrd

import (
	"io"
	"os"
)

// Book represents the contents of a "workbook".
//
// You should not instantiate this type yourself. You use the Book
// object that was returned when you called OpenWorkbook.
type Book struct {
	BaseObject

	// NSheets is the number of worksheets present in the workbook file.
	// This information is available even when no sheets have yet been loaded.
	NSheets int

	// Datemode indicates which date system was in force when this file was last saved.
	// 0: 1900 system (the Excel for Windows default).
	// 1: 1904 system (the Excel for Macintosh default).
	// Defaults to 0 in case it's not specified in the file.
	Datemode int

	// BiffVersion is the version of BIFF (Binary Interchange File Format) used to create the file.
	// Latest is 8.0 (represented here as 80), introduced with Excel 97.
	// Earliest supported by this module: 2.0 (represented as 20).
	BiffVersion int

	// NameObjList contains a Name object for each NAME record in the workbook.
	NameObjList []*Name

	// Codepage is an integer denoting the character set used for strings in this file.
	// For BIFF 8 and later, this will be 1200, meaning Unicode;
	// more precisely, UTF_16_LE.
	// For earlier versions, this is used to derive the appropriate encoding.
	Codepage *int

	// Encoding is the encoding that was derived from the codepage.
	Encoding string

	// Countries is a tuple containing the telephone country code for:
	// [0]: the user-interface setting when the file was created.
	// [1]: the regional settings.
	Countries [2]int

	// UserName is what (if anything) is recorded as the name of the last user to save the file.
	UserName string

	// FontList is a list of Font class instances, each corresponding to a FONT record.
	FontList []*Font

	// XFList is a list of XF class instances, each corresponding to an XF record.
	XFList []*XF

	// FormatList is a list of Format objects, each corresponding to a FORMAT record.
	FormatList []*Format

	// FormatMap is the mapping from XF.format_key to Format object.
	FormatMap map[int]*Format

	// StyleNameMap provides access via name to the extended format information.
	StyleNameMap map[string][2]int // maps name to (built_in, xf_index)

	// ColourMap provides definitions for colour indexes.
	ColourMap map[int][3]int // maps index to (red, green, blue)

	// PaletteRecord contains RGB values if the user has changed any colours.
	PaletteRecord [][3]int

	// LoadTimeStage1 is the time in seconds to extract the XLS image as a contiguous string.
	LoadTimeStage1 float64

	// LoadTimeStage2 is the time in seconds to parse the data from the contiguous string.
	LoadTimeStage2 float64

	// Internal fields
	sheetList   []*Sheet
	sheetNames  []string
	onDemand    bool
	logfile     io.Writer
	verbosity   int
	mem         []byte
	base        int
	streamLen   int
}

// Name represents information relating to a named reference, formula, macro, etc.
//
// Note: Name information is not extracted from files older than Excel 5.0 (Book.BiffVersion < 50)
type Name struct {
	BaseObject
	Book *Book

	// Hidden: 0 = Visible; 1 = Hidden
	Hidden int

	// Func: 0 = Command macro; 1 = Function macro. Relevant only if Macro == 1
	Func int

	// VBasic: 0 = Sheet macro; 1 = VisualBasic macro. Relevant only if Macro == 1
	VBasic int

	// Macro: 0 = Standard name; 1 = Macro name
	Macro int

	// Name is the name of the object
	Name string

	// Scope is the sheet index (0-based) or -1 for global scope
	Scope int

	// Result is the result of the formula evaluation
	Result interface{}

	// Formula is the formula text
	Formula string
}

// Sheets returns a list of all sheets in the book.
// All sheets not already loaded will be loaded.
func (b *Book) Sheets() []*Sheet {
	// Empty implementation for now
	return b.sheetList
}

// SheetByIndex returns a sheet by its index.
func (b *Book) SheetByIndex(sheetx int) (*Sheet, error) {
	// Empty implementation for now
	if sheetx < 0 || sheetx >= len(b.sheetList) {
		return nil, NewXLRDError("sheet index %d out of range", sheetx)
	}
	return b.sheetList[sheetx], nil
}

// SheetByName returns a sheet by its name.
func (b *Book) SheetByName(sheetName string) (*Sheet, error) {
	// Empty implementation for now
	for i, name := range b.sheetNames {
		if name == sheetName {
			return b.SheetByIndex(i)
		}
	}
	return nil, NewXLRDError("No sheet named <%s>", sheetName)
}

// SheetNames returns a list of all sheet names.
func (b *Book) SheetNames() []string {
	return b.sheetNames
}

// OpenWorkbookOptions contains options for opening a workbook.
type OpenWorkbookOptions struct {
	// Logfile is an open file to which messages and diagnostics are written.
	Logfile io.Writer

	// Verbosity increases the volume of trace material written to the logfile.
	Verbosity int

	// UseMmap determines whether to use memory mapping.
	// Whether to use the mmap module is determined heuristically.
	// Use this arg to override the result.
	UseMmap bool

	// FileContents is the file contents as bytes.
	// If FileContents is supplied, Filename will not be used, except (possibly) in messages.
	FileContents []byte

	// EncodingOverride is used to overcome missing or bad codepage information in older-version files.
	EncodingOverride string

	// FormattingInfo: The default is false, which saves memory.
	// When true, formatting information will be read from the spreadsheet file.
	FormattingInfo bool

	// OnDemand governs whether sheets are all loaded initially or when demanded by the caller.
	OnDemand bool

	// RaggedRows: The default of false means all rows are padded out with empty cells.
	// True means that there are no empty cells at the ends of rows.
	RaggedRows bool

	// IgnoreWorkbookCorruption allows to read corrupted workbooks.
	// When false you may face CompDocError: Workbook corruption.
	// When true that error will be ignored.
	IgnoreWorkbookCorruption bool
}

// OpenWorkbook opens a spreadsheet file for data extraction.
//
// filename: The path to the spreadsheet file to be opened.
// options: Optional parameters for opening the workbook.
//
// Returns: An instance of the Book class.
func OpenWorkbook(filename string, options *OpenWorkbookOptions) (*Book, error) {
	// Empty implementation for now
	if options == nil {
		options = &OpenWorkbookOptions{
			Logfile: os.Stdout,
		}
	}

	fileFormat, err := InspectFormat(filename, nil)
	if err != nil {
		return nil, err
	}

	if fileFormat != "" && fileFormat != "xls" {
		return nil, NewXLRDError("%s; not supported", FileFormatDescriptions[fileFormat])
	}

	return OpenWorkbookXLS(filename, options)
}

// OpenWorkbookXLS opens an XLS workbook file.
func OpenWorkbookXLS(filename string, options *OpenWorkbookOptions) (*Book, error) {
	// Empty implementation for now
	bk := &Book{
		sheetList:  []*Sheet{},
		sheetNames: []string{},
	}
	return bk, nil
}

// Colname returns the column name for a given column index (0-based).
// Example: Colname(0) returns "A", Colname(25) returns "Z", Colname(26) returns "AA"
func Colname(colx int) string {
	// Empty implementation for now
	return ""
}
