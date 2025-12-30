package xlrd

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"reflect"
	"sort"
	"strings"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
)

const (
	SUPBOOK_UNK      = 0
	SUPBOOK_INTERNAL = 1
	SUPBOOK_EXTERNAL = 2
	SUPBOOK_ADDIN    = 3
	SUPBOOK_DDEOLE   = 4
	MY_EOF           = 0xF00BAAA // not a 16-bit number
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

	// ColourIndexesUsed tracks which colour indexes are referenced by formatting records.
	ColourIndexesUsed map[int]bool

	// LoadTimeStage1 is the time in seconds to extract the XLS image as a contiguous string.
	LoadTimeStage1 float64

	// LoadTimeStage2 is the time in seconds to parse the data from the contiguous string.
	LoadTimeStage2 float64

	// Internal fields
	sheetList                []*Sheet
	sheetNames               []string
	sheetAbsPosn             []int // Absolute positions of sheets in the stream
	sheetStreamLen           []int // Stream lengths of sheets
	sheetVisibility          []int
	onDemand                 bool
	logfile                  io.Writer
	verbosity                int
	mem                      []byte
	base                     int
	streamLen                int
	position                 int
	filestr                  []byte
	formattingInfo           bool
	raggedRows               bool
	encodingOverride         string
	ignoreWorkbookCorruption bool
	sharedStrings            []string
	richTextRunlistMap       map[int][][]int

	// Name mappings
	nameAndScopeMap map[string]map[int]*Name // maps (lower_case_name, scope) to Name object
	nameMap         map[string][]*Name       // maps lower_case_name to list of Name objects

	// Processing flags
	xfEpilogueDone    bool // whether XF epilogue has been processed
	resourcesReleased bool // whether resources have been released

	// BIFF format information
	builtinfmtcount int // number of built-in formats (BIFF 3, 4S, 4W)
	sheethdrCount   int // BIFF 4W only
	sheetsoffset    int // sheet offset for BIFF 4W
	xfCount         int // number of XF records seen so far
	actualFmtCount  int // number of FORMAT records seen so far

	xfIndexToXLTypeMap map[int]int

	// External reference handling
	supbookCount       int
	supbookLocalsInx   *int
	supbookAddinsInx   *int
	externsheetInfo    [][]int
	externsheetTypeB57 []int
	extnshtNameFromNum map[int]string
	extnshtCount       int
	supbookTypes       []int
	addinFuncNames     []string
	allSheetsMap       []int // maps an all_sheets index to a calc-sheets index (or -1)
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

	// Complex: 0 = Simple formula; 1 = Complex formula (array formula or user defined)
	Complex int

	// Builtin: 0 = User-defined name; 1 = Built-in name
	Builtin int

	// Funcgroup: Function group. Relevant only if macro == 1
	Funcgroup int

	// Binary: 0 = Formula definition; 1 = Binary data
	Binary int

	// NameIndex: The index of this object in book.name_obj_list
	NameIndex int

	// Name is the name of the object
	Name string

	// Scope is the sheet index (0-based) or -1 for global scope
	Scope int

	// Result is the result of the formula evaluation
	Result interface{}

	// Formula is the formula text
	Formula string

	// RawFormula contains the raw formula bytes
	RawFormula []byte

	// BasicFormulaLen is the length of the basic formula
	BasicFormulaLen int

	// Evaluated indicates if the name has been evaluated
	Evaluated bool

	// AnyErr indicates if there are any errors in the name
	AnyErr int

	// AnyRel indicates if the name contains relative references
	AnyRel int

	// AnyExternal indicates if the name refers to external references
	AnyExternal int

	// Stack contains the evaluation stack for the name
	Stack []*Operand
}

// Cell returns a single cell that the name refers to.
// This is a convenience method for names that refer to a single cell.
func (n *Name) Cell() (*Cell, error) {
	// Note: Full implementation requires formula evaluation
	// For now, return an error indicating this feature is not yet implemented
	return nil, NewXLRDError("Name.Cell() not yet implemented - requires formula evaluation")
}

// Area2D returns a rectangular area that the name refers to.
// Returns (sheet, rowxlo, rowxhi, colxlo, colxhi)
func (n *Name) Area2D(clipped bool) (*Sheet, int, int, int, int, error) {
	// Note: Full implementation requires formula evaluation
	// For now, return an error indicating this feature is not yet implemented
	return nil, 0, 0, 0, 0, NewXLRDError("Name.Area2D() not yet implemented - requires formula evaluation")
}

// Sheets returns a list of all sheets in the book.
// All sheets not already loaded will be loaded.
func (b *Book) Sheets() []*Sheet {
	for sheetx := 0; sheetx < len(b.sheetList); sheetx++ {
		if b.sheetList[sheetx] == nil {
			sheet, _ := b.getSheet(sheetx)
			b.sheetList[sheetx] = sheet
		}
	}
	return b.sheetList
}

// SheetByIndex returns a sheet by its index.
func (b *Book) SheetByIndex(sheetx int) (*Sheet, error) {
	if sheetx < 0 || sheetx >= len(b.sheetList) {
		return nil, NewXLRDError("sheet index %d out of range", sheetx)
	}

	// Sheet should already be loaded by readWorksheets
	if b.sheetList[sheetx] != nil {
		return b.sheetList[sheetx], nil
	}

	return nil, NewXLRDError("sheet %d not loaded", sheetx)
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

// Get returns a sheet by index or name.
// This implements Python-like indexing: book[0] or book["sheetname"]
func (b *Book) Get(key interface{}) (*Sheet, error) {
	switch k := key.(type) {
	case int:
		return b.SheetByIndex(k)
	case string:
		return b.SheetByName(k)
	default:
		return nil, NewXLRDError("Invalid key type for sheet access")
	}
}

// SheetLoaded returns true if the sheet is loaded, false otherwise.
func (b *Book) SheetLoaded(sheetNameOrIndex interface{}) (bool, error) {
	var sheetx int

	switch v := sheetNameOrIndex.(type) {
	case int:
		sheetx = v
	case string:
		for i, name := range b.sheetNames {
			if name == v {
				sheetx = i
				goto found
			}
		}
		return false, NewXLRDError("No sheet named <%s>", v)
	}

found:
	if sheetx < 0 || sheetx >= len(b.sheetList) {
		return false, NewXLRDError("sheet index %d out of range", sheetx)
	}
	return b.sheetList[sheetx] != nil, nil
}

// UnloadSheet unloads a sheet by name or index.
func (b *Book) UnloadSheet(sheetNameOrIndex interface{}) error {
	var sheetx int

	switch v := sheetNameOrIndex.(type) {
	case int:
		sheetx = v
	case string:
		for i, name := range b.sheetNames {
			if name == v {
				sheetx = i
				goto found
			}
		}
		return NewXLRDError("No sheet named <%s>", v)
	}

found:
	if sheetx < 0 || sheetx >= len(b.sheetList) {
		return NewXLRDError("sheet index %d out of range", sheetx)
	}
	b.sheetList[sheetx] = nil
	return nil
}

// ReleaseResources releases memory-consuming objects and possibly a memory-mapped file.
func (b *Book) ReleaseResources() {
	b.resourcesReleased = true
	// Note: In Go, memory management is automatic, so we mainly set flags
	// If there were mmap objects, they would be closed here
	b.mem = nil
	b.filestr = nil
	b.sharedStrings = nil
}

// GetBOF gets the BOF (Beginning of File) record for a given sheet type.
func (b *Book) GetBOF(sheetType int) int {
	// Empty implementation for now
	return 0
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

	// IgnoreWorkbookCorruption allows reading corrupted workbooks.
	// When false (default), you may face CompDocError: Workbook corruption.
	// When true, that exception will be ignored.
	IgnoreWorkbookCorruption bool
}

// OpenWorkbook opens a spreadsheet file for data extraction.
//
// filename: The path to the spreadsheet file to be opened.
// options: Optional parameters for opening the workbook.
//
// Returns: An instance of the Book class.
func OpenWorkbook(filename string, options *OpenWorkbookOptions) (*Book, error) {
	if options == nil {
		options = &OpenWorkbookOptions{
			Logfile: os.Stdout,
		}
	}
	if options.Logfile == nil {
		options.Logfile = os.Stdout
	}

	var fileFormat string
	var err error
	if options.FileContents != nil {
		fileFormat, err = InspectFormat("", options.FileContents)
	} else {
		fileFormat, err = InspectFormat(filename, nil)
	}
	if err != nil {
		return nil, err
	}

	// Allow unknown formats to pass through, as some ancient files that xlrd can parse
	// don't start with the expected signature (e.g., raw BIFF files)
	if fileFormat != "" && fileFormat != "xls" {
		return nil, NewXLRDError("%s; not supported", FileFormatDescriptions[fileFormat])
	}

	return OpenWorkbookXLS(filename, options)
}

// OpenWorkbookXLS opens an XLS workbook file.
func OpenWorkbookXLS(filename string, options *OpenWorkbookOptions) (*Book, error) {
	bk := &Book{
		sheetList:          []*Sheet{},
		sheetNames:         []string{},
		externsheetInfo:    [][]int{},
		externsheetTypeB57: []int{},
		extnshtNameFromNum: make(map[int]string),
		supbookTypes:       []int{},
		addinFuncNames:     []string{},
	}

	if options == nil {
		options = &OpenWorkbookOptions{
			Logfile: os.Stdout,
		}
	}
	if options.Logfile == nil {
		options.Logfile = os.Stdout
	}

	bk.logfile = options.Logfile
	bk.verbosity = options.Verbosity
	bk.onDemand = options.OnDemand
	bk.formattingInfo = options.FormattingInfo
	bk.raggedRows = options.RaggedRows
	bk.encodingOverride = options.EncodingOverride
	bk.ignoreWorkbookCorruption = options.IgnoreWorkbookCorruption

	// Read file
	var err error
	var fileContents []byte
	if options.FileContents != nil {
		fileContents = options.FileContents
	} else {
		fileContents, err = os.ReadFile(filename)
		if err != nil {
			return nil, err
		}
	}

	if len(fileContents) == 0 {
		return nil, NewXLRDError("File size is 0 bytes")
	}

	bk.filestr = fileContents
	bk.streamLen = len(fileContents)
	bk.base = 0

	// Check if it's an OLE2 compound document
	if len(fileContents) >= 8 && string(fileContents[:8]) == string(XLS_SIGNATURE) {
		// It's an OLE2 compound document
		cd, err := NewCompDoc(fileContents, options.Logfile, 0, options.IgnoreWorkbookCorruption)
		if err != nil {
			return nil, err
		}

		// Try to locate Workbook or Book stream
		var mem []byte
		var base, streamLen int
		var lastErr error
		for _, qname := range []string{"Workbook", "Book"} {
			mem, base, streamLen, err = cd.LocateNamedStream(qname)
			if err == nil && mem != nil {
				break
			}
			if err != nil {
				lastErr = err
				// Check if it's a corruption error that should not be ignored
				if compDocErr, ok := err.(*CompDocError); ok && !options.IgnoreWorkbookCorruption {
					return nil, compDocErr
				}
			}
		}

		if mem == nil {
			if lastErr != nil {
				return nil, lastErr
			}
			return nil, NewXLRDError("Can't find workbook in OLE2 compound document")
		}

		bk.mem = mem
		bk.base = base
		bk.streamLen = streamLen
	} else {
		// Not an OLE2 compound document - treat as raw BIFF file
		bk.mem = fileContents
		bk.base = 0
		bk.streamLen = len(fileContents)
	}

	bk.position = bk.base

	// Parse BIFF records to extract sheet names and other information
	err = bk.parseGlobals(options)
	if err != nil {
		return nil, err
	}

	// Read all worksheets
	err = bk.readWorksheets(options)
	if err != nil {
		return nil, err
	}

	return bk, nil
}

// parseGlobals parses the workbook globals section.
func (b *Book) parseGlobals(options *OpenWorkbookOptions) error {
	// Get BOF record
	biffVersion, err := b.getBOF(XL_WORKBOOK_GLOBALS)
	if err != nil {
		return err
	}

	if biffVersion == 0 {
		return NewXLRDError("Can't determine file's BIFF version")
	}

	// Check if version is supported
	supported := false
	for _, v := range SupportedVersions {
		if v == biffVersion {
			supported = true
			break
		}
	}
	if !supported {
		return NewXLRDError("BIFF version %s is not supported", BiffTextFromNum(biffVersion))
	}

	b.BiffVersion = biffVersion

	// Parse records based on BIFF version
	if biffVersion <= 40 {
		// BIFF 4.0 and earlier - no workbook globals, only 1 worksheet
		if options.OnDemand {
			fmt.Fprintf(b.logfile, "*** WARNING: on_demand is not supported for this Excel version.\n*** Setting on_demand to False.\n")
			options.OnDemand = false
		}
		b.fakeGlobalsGetSheet()
	} else if biffVersion == 45 {
		// BIFF 4W - worksheet(s) embedded in global stream
		b.parseGlobalsRecords(options)
		if options.OnDemand {
			fmt.Fprintf(b.logfile, "*** WARNING: on_demand is not supported for this Excel version.\n*** Setting on_demand to False.\n")
			options.OnDemand = false
		}
	} else {
		// BIFF 5 and later
		b.parseGlobalsRecords(options)
		b.sheetList = make([]*Sheet, len(b.sheetNames))
		if !options.OnDemand {
			// Load all sheets
			err = b.getSheets()
			if err != nil {
				return err
			}
		}
	}

	b.NSheets = len(b.sheetList)
	return nil
}

// getBOF gets the BOF (Beginning of File) record.
func (b *Book) getBOF(rqdStream int) (int, error) {
	if b.position+4 > len(b.mem) {
		return 0, NewXLRDError("Expected BOF record; met end of file")
	}

	opcode := int(binary.LittleEndian.Uint16(b.mem[b.position : b.position+2]))
	b.position += 2

	// Check if it's a valid BOF code
	validBOF := false
	for _, code := range bofcodes {
		if opcode == code {
			validBOF = true
			break
		}
	}
	if !validBOF {
		return 0, NewXLRDError("Expected BOF record; found 0x%04x", opcode)
	}

	if b.position+2 > len(b.mem) {
		return 0, NewXLRDError("Incomplete BOF record; met end of file")
	}

	length := int(binary.LittleEndian.Uint16(b.mem[b.position : b.position+2]))
	b.position += 2

	if length < 4 || length > 20 {
		return 0, NewXLRDError("Invalid length (%d) for BOF record type 0x%04x", length, opcode)
	}

	expectedLen, ok := boflen[opcode]
	if !ok {
		return 0, NewXLRDError("Unknown BOF record type 0x%04x", opcode)
	}

	if b.position+length > len(b.mem) {
		return 0, NewXLRDError("Incomplete BOF record; met end of file")
	}

	data := b.mem[b.position : b.position+length]
	b.position += length

	// Pad if necessary
	if length < expectedLen {
		padding := make([]byte, expectedLen-length)
		data = append(data, padding...)
	}

	version1 := opcode >> 8
	version2 := binary.LittleEndian.Uint16(data[0:2])
	streamtype := binary.LittleEndian.Uint16(data[2:4])

	var version int
	if version1 == 0x08 {
		build := binary.LittleEndian.Uint16(data[4:6])
		year := binary.LittleEndian.Uint16(data[6:8])
		if version2 == 0x0600 {
			version = 80
		} else if version2 == 0x0500 {
			if year < 1994 || (build == 2412 || build == 3218 || build == 3321) {
				version = 50
			} else {
				version = 70
			}
		} else {
			// Handle other versions
			versionMap := map[uint16]int{
				0x0000: 21,
				0x0007: 21,
			}
			if v, ok := versionMap[version2]; ok {
				version = v
			} else {
				return 0, NewXLRDError("Unknown BIFF version: 0x%04x", version2)
			}
		}
	} else if version1 == 0x04 {
		version = 40
	} else if version1 == 0x02 {
		version = 30
	} else if version1 == 0x00 {
		version = 20
	} else {
		return 0, NewXLRDError("Unknown BIFF version: 0x%02x", version1)
	}

	// Check stream type (Python xlrd logic)
	gotGlobals := streamtype == XL_WORKBOOK_GLOBALS || (version == 45 && streamtype == XL_WORKBOOK_GLOBALS_4W)
	if (rqdStream == XL_WORKBOOK_GLOBALS && gotGlobals) || streamtype == uint16(rqdStream) {
		return version, nil
	}
	if version < 50 && streamtype == XL_WORKSHEET {
		return version, nil
	}
	if version >= 50 && streamtype == 0x0100 {
		return 0, NewXLRDError("Workspace file -- no spreadsheet data")
	}
	return 0, NewXLRDError("BOF not workbook/worksheet: op=0x%04x vers=0x%04x strm=0x%04x -> BIFF%d", opcode, version2, streamtype, version)
}

// parseGlobalsRecords parses the workbook globals records.
func (b *Book) parseGlobalsRecords(options *OpenWorkbookOptions) error {
	b.initializeFormatInfo()
	b.sheetNames = make([]string, 0)
	b.sheetList = make([]*Sheet, 0)

	// Set encoding with override if provided, or derive from codepage
	b.Encoding = b.deriveEncoding()

	for b.position < len(b.mem) {
		if b.position+4 > len(b.mem) {
			break
		}

		code := int(binary.LittleEndian.Uint16(b.mem[b.position : b.position+2]))
		length := int(binary.LittleEndian.Uint16(b.mem[b.position+2 : b.position+4]))
		b.position += 4

		if b.position+length > len(b.mem) {
			break
		}

		data := b.mem[b.position : b.position+length]
		b.position += length

		switch code {
		case XL_EOF:
			if !b.xfEpilogueDone {
				b.xfEpilogue()
			}
			b.namesEpilogue()
			b.paletteEpilogue()
			if b.Encoding == "" {
				b.Encoding = b.deriveEncoding()
			}
			break
		case XL_BOUNDSHEET:
			err := b.handleBoundsheet(data)
			if err != nil {
				return err
			}
		case XL_CODEPAGE:
			b.handleCodepage(data)
		case XL_DATEMODE:
			b.handleDatemode(data)
		case XL_COUNTRY:
			b.handleCountry(data)
		case XL_WRITEACCESS:
			b.handleWriteAccess(data)
		case XL_FONT:
			err := b.handleFont(data)
			if err != nil {
				return err
			}
		case XL_EFONT:
			err := b.handleEFont(data)
			if err != nil {
				return err
			}
		case XL_FORMAT:
			err := b.handleFormat(data, XL_FORMAT)
			if err != nil {
				return err
			}
		case XL_FORMAT2:
			err := b.handleFormat(data, XL_FORMAT2)
			if err != nil {
				return err
			}
		case XL_XF:
			err := b.handleXF(data)
			if err != nil {
				return err
			}
		case XL_STYLE:
			err := b.handleStyle(data)
			if err != nil {
				return err
			}
		case XL_PALETTE:
			err := b.handlePalette(data)
			if err != nil {
				return err
			}
		case XL_NAME:
			err := b.handleName(data)
			if err != nil {
				return err
			}
		case XL_EXTERNNAME:
			err := b.handleExternname(data)
			if err != nil {
				return err
			}
		case XL_EXTERNSHEET:
			err := b.handleExternsheet(data)
			if err != nil {
				return err
			}
		case XL_SUPBOOK:
			err := b.handleSupbook(data)
			if err != nil {
				return err
			}
		case XL_SST:
			err := b.handleSST(data)
			if err != nil {
				return err
			}
		case XL_BUILTINFMTCOUNT:
			b.handleBuiltinfmtcount(data)
		case XL_FILEPASS:
			err := b.handleFilepass(data)
			if err != nil {
				return err
			}
		case XL_OBJ:
			b.handleObj(data)
		case XL_SHEETHDR:
			err := b.handleSheethdr(data)
			if err != nil {
				return err
			}
		case XL_SHEETSOFFSET:
			b.handleSheetsoffset(data)
		}

		if code == XL_EOF {
			break
		}
	}

	// Calculate stream lengths for each sheet
	b.sheetStreamLen = make([]int, len(b.sheetAbsPosn))
	for i := 0; i < len(b.sheetAbsPosn)-1; i++ {
		b.sheetStreamLen[i] = b.sheetAbsPosn[i+1] - b.sheetAbsPosn[i]
	}
	// Last sheet goes to end of stream
	if len(b.sheetAbsPosn) > 0 {
		b.sheetStreamLen[len(b.sheetAbsPosn)-1] = len(b.mem) - b.sheetAbsPosn[len(b.sheetAbsPosn)-1]
	}

	return nil
}

// handleBoundsheet handles a BOUNDSHEET record.
func (b *Book) handleBoundsheet(data []byte) error {
	bv := b.BiffVersion

	var sheetName string
	var visibility int
	var absPosn int
	var sheetType int
	var err error

	if bv == 45 {
		// BIFF 4W - only sheet name
		sheetName, err = UnpackString(data, 0, b.Encoding, 1)
		if err != nil {
			return err
		}
		visibility = 0
		sheetType = XL_BOUNDSHEET_WORKSHEET // guess, patch later
		if len(b.sheetAbsPosn) == 0 {
			// For BIFF4W, sheets are embedded in the global stream
			// _sheetsoffset would be calculated here, but for now use base
			absPosn = b.base
			// Note (a) this won't be used
			// (b) it's the position of the SHEETHDR record
			// (c) add 11 to get to the worksheet BOF record
		} else {
			absPosn = -1 // unknown
		}
	} else {
		if len(data) < 6 {
			return NewXLRDError("BOUNDSHEET record too short")
		}

		offset := int(int32(binary.LittleEndian.Uint32(data[0:4])))
		visibility = int(data[4])
		sheetType = int(data[5])
		absPosn = offset + b.base // because global BOF is always at posn 0 in the stream

		if bv < BIFF_FIRST_UNICODE {
			sheetName, err = UnpackString(data, 6, b.Encoding, 1)
		} else {
			sheetName, err = UnpackUnicode(data, 6, 1)
		}
		if err != nil {
			return err
		}
	}

	if sheetType == XL_BOUNDSHEET_WORKSHEET {
		b.sheetNames = append(b.sheetNames, sheetName)
		b.sheetList = append(b.sheetList, nil)
		b.sheetAbsPosn = append(b.sheetAbsPosn, absPosn)
		b.sheetVisibility = append(b.sheetVisibility, visibility)
	}

	return nil
}

// handleCodepage handles a CODEPAGE record.
func (b *Book) handleCodepage(data []byte) {
	if len(data) < 2 {
		return
	}
	codepage := int(binary.LittleEndian.Uint16(data[0:2]))
	b.Codepage = &codepage
}

// handleDatemode handles a DATEMODE record.
func (b *Book) handleDatemode(data []byte) {
	if len(data) < 2 {
		return
	}
	datemode := int(binary.LittleEndian.Uint16(data[0:2]))
	if datemode == 0 || datemode == 1 {
		b.Datemode = datemode
	}
}

// handleCountry handles a COUNTRY record.
func (b *Book) handleCountry(data []byte) {
	if len(data) < 4 {
		return
	}
	b.Countries[0] = int(binary.LittleEndian.Uint16(data[0:2]))
	b.Countries[1] = int(binary.LittleEndian.Uint16(data[2:4]))
}

// handleFont handles a FONT record.
func (b *Book) handleFont(data []byte) error {
	if !b.formattingInfo {
		return nil
	}
	if b.Encoding == "" {
		b.Encoding = b.deriveEncoding()
	}

	k := len(b.FontList)
	if k == 4 {
		dummy := &Font{
			Name:      "Dummy Font",
			FontIndex: k,
		}
		b.FontList = append(b.FontList, dummy)
		k++
	}

	font := &Font{FontIndex: k}
	bv := b.BiffVersion

	if bv >= 50 {
		if len(data) < 14 {
			return nil
		}
		font.Height = int(binary.LittleEndian.Uint16(data[0:2]))
		optionFlags := binary.LittleEndian.Uint16(data[2:4])
		font.ColourIndex = int(binary.LittleEndian.Uint16(data[4:6]))
		font.Weight = int(binary.LittleEndian.Uint16(data[6:8]))
		font.Escapement = int(binary.LittleEndian.Uint16(data[8:10]))
		font.Underline = int(data[10])
		font.Family = int(data[11])
		font.CharacterSet = int(data[12])

		font.Bold = (optionFlags & 0x0001) != 0
		font.Italic = (optionFlags & 0x0002) != 0
		font.Underlined = (optionFlags & 0x0004) != 0
		font.StruckOut = (optionFlags & 0x0008) != 0
		font.Outline = (optionFlags & 0x0010) != 0
		font.Shadow = (optionFlags & 0x0020) != 0

		var err error
		if bv >= 80 {
			font.Name, err = UnpackUnicode(data, 14, 1)
		} else {
			font.Name, err = UnpackString(data, 14, b.Encoding, 1)
		}
		if err != nil {
			return err
		}
	} else if bv >= 30 {
		if len(data) < 6 {
			return nil
		}
		font.Height = int(binary.LittleEndian.Uint16(data[0:2]))
		optionFlags := binary.LittleEndian.Uint16(data[2:4])
		font.ColourIndex = int(binary.LittleEndian.Uint16(data[4:6]))

		font.Bold = (optionFlags & 0x0001) != 0
		font.Italic = (optionFlags & 0x0002) != 0
		font.Underlined = (optionFlags & 0x0004) != 0
		font.StruckOut = (optionFlags & 0x0008) != 0
		font.Outline = (optionFlags & 0x0010) != 0
		font.Shadow = (optionFlags & 0x0020) != 0

		name, err := UnpackString(data, 6, b.Encoding, 1)
		if err != nil {
			return err
		}
		font.Name = name
		if font.Bold {
			font.Weight = 700
		} else {
			font.Weight = 400
		}
		font.Escapement = 0
		font.Underline = 0
		if font.Underlined {
			font.Underline = 1
		}
		font.Family = 0
		font.CharacterSet = 1
	} else {
		if len(data) < 4 {
			return nil
		}
		font.Height = int(binary.LittleEndian.Uint16(data[0:2]))
		optionFlags := binary.LittleEndian.Uint16(data[2:4])
		font.ColourIndex = 0x7FFF

		font.Bold = (optionFlags & 0x0001) != 0
		font.Italic = (optionFlags & 0x0002) != 0
		font.Underlined = (optionFlags & 0x0004) != 0
		font.StruckOut = (optionFlags & 0x0008) != 0
		font.Outline = false
		font.Shadow = false

		name, err := UnpackString(data, 4, b.Encoding, 1)
		if err != nil {
			return err
		}
		font.Name = name
		if font.Bold {
			font.Weight = 700
		} else {
			font.Weight = 400
		}
		font.Escapement = 0
		font.Underline = 0
		if font.Underlined {
			font.Underline = 1
		}
		font.Family = 0
		font.CharacterSet = 1
	}

	b.FontList = append(b.FontList, font)
	return nil
}

func (b *Book) handleEFont(data []byte) error {
	if !b.formattingInfo {
		return nil
	}
	if len(data) < 2 {
		return nil
	}
	if len(b.FontList) == 0 {
		return nil
	}
	b.FontList[len(b.FontList)-1].ColourIndex = int(binary.LittleEndian.Uint16(data[0:2]))
	return nil
}

// handleFormat handles a FORMAT record.
func (b *Book) handleFormat(data []byte, rectype int) error {
	if len(data) < 2 {
		return nil
	}
	if b.Encoding == "" {
		b.Encoding = b.deriveEncoding()
	}
	bv := b.BiffVersion
	if rectype == XL_FORMAT2 && bv > 30 {
		bv = 30
	}

	formatKey := 0
	strPos := 2
	if bv >= 50 {
		formatKey = int(binary.LittleEndian.Uint16(data[0:2]))
	} else {
		formatKey = b.actualFmtCount
		if bv <= 30 {
			strPos = 0
		}
	}
	b.actualFmtCount++

	formatString := ""
	var err error
	if bv >= 80 {
		formatString, err = UnpackUnicode(data, 2, 2)
		if err != nil {
			return err
		}
	} else {
		formatString, err = UnpackString(data, strPos, b.Encoding, 1)
		if err != nil {
			return err
		}
	}

	isDate := IsDateFormatString(b, formatString)
	formatType := FGE
	if isDate {
		formatType = FDT
	}
	if !(formatKey > 163 || bv < 50) {
		stdType, ok := stdFormatCodeTypes[formatKey]
		if ok && b.verbosity > 0 {
			isDateCode := stdType == FDT
			if formatKey > 0 && formatKey < 50 && (isDateCode != isDate) {
				fmt.Fprintf(b.logfile,
					"WARNING *** Conflict between std format key %d and its format string %q\n",
					formatKey, formatString)
			}
		}
	}

	format := &Format{
		FormatKey:    formatKey,
		Type:         formatType,
		FormatString: formatString,
	}
	b.FormatMap[formatKey] = format
	b.FormatList = append(b.FormatList, format)
	return nil
}

// handleXF handles an XF (Extended Format) record.
func (b *Book) handleXF(data []byte) error {
	if len(data) < 4 {
		return nil
	}

	xf := &XF{
		Alignment:  &XFAlignment{},
		Border:     &XFBorder{},
		Background: &XFBackground{},
		Protection: &XFProtection{},
	}
	xf.Alignment.IndentLevel = 0
	xf.Alignment.ShrinkToFit = false
	xf.Alignment.TextDirection = 0
	xf.Border.DiagUp = 0
	xf.Border.DiagDown = 0
	xf.Border.DiagColourIndex = 0
	xf.Border.DiagLineStyle = 0

	bv := b.BiffVersion
	if bv >= 50 && b.xfCount == 0 {
		fillInStandardFormats(b)
	}

	switch {
	case bv >= 80:
		if len(data) < 20 {
			return nil
		}
		xf.FontIndex = int(binary.LittleEndian.Uint16(data[0:2]))
		xf.FormatKey = int(binary.LittleEndian.Uint16(data[2:4]))
		pkdTypePar := binary.LittleEndian.Uint16(data[4:6])
		pkdAlign1 := data[6]
		xf.Alignment.Rotation = int(data[7])
		pkdAlign2 := data[8]
		pkdUsed := data[9]
		pkdBrdBkg1 := binary.LittleEndian.Uint32(data[10:14])
		pkdBrdBkg2 := binary.LittleEndian.Uint32(data[14:18])
		pkdBrdBkg3 := binary.LittleEndian.Uint16(data[18:20])

		upkbits(xf.Protection, uint32(pkdTypePar), [][3]interface{}{
			{0, uint32(0x01), "CellLocked"},
			{1, uint32(0x02), "FormulaHidden"},
		})
		upkbits(xf, uint32(pkdTypePar), [][3]interface{}{
			{2, uint32(0x0004), "IsStyle"},
			{3, uint32(0x0008), "Lotus123Prefix"},
			{4, uint32(0xFFF0), "ParentStyleIndex"},
		})
		upkbits(xf.Alignment, uint32(pkdAlign1), [][3]interface{}{
			{0, uint32(0x07), "HorAlign"},
			{3, uint32(0x08), "TextWrapped"},
			{4, uint32(0x70), "VertAlign"},
		})
		upkbits(xf.Alignment, uint32(pkdAlign2), [][3]interface{}{
			{0, uint32(0x0f), "IndentLevel"},
			{4, uint32(0x10), "ShrinkToFit"},
			{6, uint32(0xC0), "TextDirection"},
		})

		reg := pkdUsed >> 2
		xf.FormatFlag = int(reg & 1)
		reg >>= 1
		xf.FontFlag = int(reg & 1)
		reg >>= 1
		xf.AlignmentFlag = int(reg & 1)
		reg >>= 1
		xf.BorderFlag = int(reg & 1)
		reg >>= 1
		xf.BackgroundFlag = int(reg & 1)
		reg >>= 1
		xf.ProtectionFlag = int(reg & 1)

		upkbitsL(xf.Border, pkdBrdBkg1, [][3]interface{}{
			{0, uint32(0x0000000f), "LeftLineStyle"},
			{4, uint32(0x000000f0), "RightLineStyle"},
			{8, uint32(0x00000f00), "TopLineStyle"},
			{12, uint32(0x0000f000), "BottomLineStyle"},
			{16, uint32(0x007f0000), "LeftColourIndex"},
			{23, uint32(0x3f800000), "RightColourIndex"},
			{30, uint32(0x40000000), "DiagDown"},
			{31, uint32(0x80000000), "DiagUp"},
		})
		upkbits(xf.Border, pkdBrdBkg2, [][3]interface{}{
			{0, uint32(0x0000007F), "TopColourIndex"},
			{7, uint32(0x00003F80), "BottomColourIndex"},
			{14, uint32(0x001FC000), "DiagColourIndex"},
			{21, uint32(0x01E00000), "DiagLineStyle"},
		})
		upkbitsL(xf.Background, pkdBrdBkg2, [][3]interface{}{
			{26, uint32(0xFC000000), "FillPattern"},
		})
		upkbits(xf.Background, uint32(pkdBrdBkg3), [][3]interface{}{
			{0, uint32(0x007F), "PatternColourIndex"},
			{7, uint32(0x3F80), "BackgroundColourIndex"},
		})
	case bv >= 50:
		if len(data) < 16 {
			return nil
		}
		xf.FontIndex = int(binary.LittleEndian.Uint16(data[0:2]))
		xf.FormatKey = int(binary.LittleEndian.Uint16(data[2:4]))
		pkdTypePar := binary.LittleEndian.Uint16(data[4:6])
		pkdAlign1 := data[6]
		pkdOrientUsed := data[7]
		pkdBrdBkg1 := binary.LittleEndian.Uint32(data[8:12])
		pkdBrdBkg2 := binary.LittleEndian.Uint32(data[12:16])

		upkbits(xf.Protection, uint32(pkdTypePar), [][3]interface{}{
			{0, uint32(0x01), "CellLocked"},
			{1, uint32(0x02), "FormulaHidden"},
		})
		upkbits(xf, uint32(pkdTypePar), [][3]interface{}{
			{2, uint32(0x0004), "IsStyle"},
			{3, uint32(0x0008), "Lotus123Prefix"},
			{4, uint32(0xFFF0), "ParentStyleIndex"},
		})
		upkbits(xf.Alignment, uint32(pkdAlign1), [][3]interface{}{
			{0, uint32(0x07), "HorAlign"},
			{3, uint32(0x08), "TextWrapped"},
			{4, uint32(0x70), "VertAlign"},
		})

		orientation := pkdOrientUsed & 0x03
		switch orientation {
		case 1:
			xf.Alignment.Rotation = 255
		case 2:
			xf.Alignment.Rotation = 90
		case 3:
			xf.Alignment.Rotation = 180
		default:
			xf.Alignment.Rotation = 0
		}

		reg := pkdOrientUsed >> 2
		xf.FormatFlag = int(reg & 1)
		reg >>= 1
		xf.FontFlag = int(reg & 1)
		reg >>= 1
		xf.AlignmentFlag = int(reg & 1)
		reg >>= 1
		xf.BorderFlag = int(reg & 1)
		reg >>= 1
		xf.BackgroundFlag = int(reg & 1)
		reg >>= 1
		xf.ProtectionFlag = int(reg & 1)

		upkbitsL(xf.Background, pkdBrdBkg1, [][3]interface{}{
			{0, uint32(0x0000007F), "PatternColourIndex"},
			{7, uint32(0x00003F80), "BackgroundColourIndex"},
			{16, uint32(0x003F0000), "FillPattern"},
		})
		upkbitsL(xf.Border, pkdBrdBkg1, [][3]interface{}{
			{22, uint32(0x01C00000), "BottomLineStyle"},
			{25, uint32(0xFE000000), "BottomColourIndex"},
		})
		upkbits(xf.Border, pkdBrdBkg2, [][3]interface{}{
			{0, uint32(0x00000007), "TopLineStyle"},
			{3, uint32(0x00000038), "LeftLineStyle"},
			{6, uint32(0x000001C0), "RightLineStyle"},
			{9, uint32(0x0000FE00), "TopColourIndex"},
			{16, uint32(0x007F0000), "LeftColourIndex"},
			{23, uint32(0x3F800000), "RightColourIndex"},
		})
	case bv >= 40:
		if len(data) < 12 {
			return nil
		}
		xf.FontIndex = int(data[0])
		xf.FormatKey = int(data[1])
		pkdTypePar := binary.LittleEndian.Uint16(data[2:4])
		pkdAlignOrient := data[4]
		pkdUsed := data[5]
		pkdBkg34 := binary.LittleEndian.Uint16(data[6:8])
		pkdBrd34 := binary.LittleEndian.Uint32(data[8:12])

		upkbits(xf.Protection, uint32(pkdTypePar), [][3]interface{}{
			{0, uint32(0x01), "CellLocked"},
			{1, uint32(0x02), "FormulaHidden"},
		})
		upkbits(xf, uint32(pkdTypePar), [][3]interface{}{
			{2, uint32(0x0004), "IsStyle"},
			{3, uint32(0x0008), "Lotus123Prefix"},
			{4, uint32(0xFFF0), "ParentStyleIndex"},
		})
		upkbits(xf.Alignment, uint32(pkdAlignOrient), [][3]interface{}{
			{0, uint32(0x07), "HorAlign"},
			{3, uint32(0x08), "TextWrapped"},
			{4, uint32(0x30), "VertAlign"},
		})
		orientation := (pkdAlignOrient & 0xC0) >> 6
		switch orientation {
		case 1:
			xf.Alignment.Rotation = 255
		case 2:
			xf.Alignment.Rotation = 90
		case 3:
			xf.Alignment.Rotation = 180
		default:
			xf.Alignment.Rotation = 0
		}

		reg := pkdUsed >> 2
		xf.FormatFlag = int(reg & 1)
		reg >>= 1
		xf.FontFlag = int(reg & 1)
		reg >>= 1
		xf.AlignmentFlag = int(reg & 1)
		reg >>= 1
		xf.BorderFlag = int(reg & 1)
		reg >>= 1
		xf.BackgroundFlag = int(reg & 1)
		reg >>= 1
		xf.ProtectionFlag = int(reg & 1)

		upkbits(xf.Background, uint32(pkdBkg34), [][3]interface{}{
			{0, uint32(0x003F), "FillPattern"},
			{6, uint32(0x07C0), "PatternColourIndex"},
			{11, uint32(0xF800), "BackgroundColourIndex"},
		})
		upkbitsL(xf.Border, pkdBrd34, [][3]interface{}{
			{0, uint32(0x00000007), "TopLineStyle"},
			{3, uint32(0x000000F8), "TopColourIndex"},
			{8, uint32(0x00000700), "LeftLineStyle"},
			{11, uint32(0x0000F800), "LeftColourIndex"},
			{16, uint32(0x00070000), "BottomLineStyle"},
			{19, uint32(0x00F80000), "BottomColourIndex"},
			{24, uint32(0x07000000), "RightLineStyle"},
			{27, uint32(0xF8000000), "RightColourIndex"},
		})
	case bv == 30:
		if len(data) < 12 {
			return nil
		}
		xf.FontIndex = int(data[0])
		xf.FormatKey = int(data[1])
		pkdTypeProt := data[2]
		pkdUsed := data[3]
		pkdAlignPar := binary.LittleEndian.Uint16(data[4:6])
		pkdBkg34 := binary.LittleEndian.Uint16(data[6:8])
		pkdBrd34 := binary.LittleEndian.Uint32(data[8:12])

		upkbits(xf.Protection, uint32(pkdTypeProt), [][3]interface{}{
			{0, uint32(0x01), "CellLocked"},
			{1, uint32(0x02), "FormulaHidden"},
		})
		upkbits(xf, uint32(pkdTypeProt), [][3]interface{}{
			{2, uint32(0x0004), "IsStyle"},
			{3, uint32(0x0008), "Lotus123Prefix"},
		})
		upkbits(xf.Alignment, uint32(pkdAlignPar), [][3]interface{}{
			{0, uint32(0x07), "HorAlign"},
			{3, uint32(0x08), "TextWrapped"},
		})
		upkbits(xf, uint32(pkdAlignPar), [][3]interface{}{
			{4, uint32(0xFFF0), "ParentStyleIndex"},
		})

		reg := pkdUsed >> 2
		xf.FormatFlag = int(reg & 1)
		reg >>= 1
		xf.FontFlag = int(reg & 1)
		reg >>= 1
		xf.AlignmentFlag = int(reg & 1)
		reg >>= 1
		xf.BorderFlag = int(reg & 1)
		reg >>= 1
		xf.BackgroundFlag = int(reg & 1)
		reg >>= 1
		xf.ProtectionFlag = int(reg & 1)

		upkbits(xf.Background, uint32(pkdBkg34), [][3]interface{}{
			{0, uint32(0x003F), "FillPattern"},
			{6, uint32(0x07C0), "PatternColourIndex"},
			{11, uint32(0xF800), "BackgroundColourIndex"},
		})
		upkbitsL(xf.Border, pkdBrd34, [][3]interface{}{
			{0, uint32(0x00000007), "TopLineStyle"},
			{3, uint32(0x000000F8), "TopColourIndex"},
			{8, uint32(0x00000700), "LeftLineStyle"},
			{11, uint32(0x0000F800), "LeftColourIndex"},
			{16, uint32(0x00070000), "BottomLineStyle"},
			{19, uint32(0x00F80000), "BottomColourIndex"},
			{24, uint32(0x07000000), "RightLineStyle"},
			{27, uint32(0xF8000000), "RightColourIndex"},
		})
		xf.Alignment.VertAlign = 2
		xf.Alignment.Rotation = 0
	case bv == 21:
		if len(data) < 4 {
			return nil
		}
		xf.FontIndex = int(data[0])
		formatEtc := data[2]
		halignEtc := data[3]
		xf.FormatKey = int(formatEtc & 0x3F)
		upkbits(xf.Protection, uint32(formatEtc), [][3]interface{}{
			{6, uint32(0x40), "CellLocked"},
			{7, uint32(0x80), "FormulaHidden"},
		})
		upkbits(xf.Alignment, uint32(halignEtc), [][3]interface{}{
			{0, uint32(0x07), "HorAlign"},
		})
		for _, side := range []struct {
			mask uint8
			name string
		}{
			{0x08, "Left"},
			{0x10, "Right"},
			{0x20, "Top"},
			{0x40, "Bottom"},
		} {
			if halignEtc&side.mask != 0 {
				switch side.name {
				case "Left":
					xf.Border.LeftColourIndex = 8
					xf.Border.LeftLineStyle = 1
				case "Right":
					xf.Border.RightColourIndex = 8
					xf.Border.RightLineStyle = 1
				case "Top":
					xf.Border.TopColourIndex = 8
					xf.Border.TopLineStyle = 1
				case "Bottom":
					xf.Border.BottomColourIndex = 8
					xf.Border.BottomLineStyle = 1
				}
			} else {
				switch side.name {
				case "Left":
					xf.Border.LeftColourIndex = 0
					xf.Border.LeftLineStyle = 0
				case "Right":
					xf.Border.RightColourIndex = 0
					xf.Border.RightLineStyle = 0
				case "Top":
					xf.Border.TopColourIndex = 0
					xf.Border.TopLineStyle = 0
				case "Bottom":
					xf.Border.BottomColourIndex = 0
					xf.Border.BottomLineStyle = 0
				}
			}
		}
		if halignEtc&0x80 != 0 {
			xf.Background.FillPattern = 17
		} else {
			xf.Background.FillPattern = 0
		}
		xf.Background.BackgroundColourIndex = 9
		xf.Background.PatternColourIndex = 8
		xf.ParentStyleIndex = 0
		xf.Alignment.VertAlign = 2
		xf.Alignment.Rotation = 0
		xf.FormatFlag = 1
		xf.FontFlag = 1
		xf.AlignmentFlag = 1
		xf.BorderFlag = 1
		xf.BackgroundFlag = 1
		xf.ProtectionFlag = 1
	default:
		return NewXLRDError("unknown BIFF version %d in XF record", bv)
	}

	xf.Alignment.Horizontal = xf.Alignment.HorAlign
	xf.Alignment.Vertical = xf.Alignment.VertAlign
	xf.Alignment.WrapText = xf.Alignment.TextWrapped

	xf.Border.Left = xf.Border.LeftLineStyle
	xf.Border.Right = xf.Border.RightLineStyle
	xf.Border.Top = xf.Border.TopLineStyle
	xf.Border.Bottom = xf.Border.BottomLineStyle

	xf.Locked = xf.Protection.CellLocked
	xf.Hidden = xf.Protection.FormulaHidden

	xf.XFIndex = len(b.XFList)
	b.XFList = append(b.XFList, xf)
	b.xfCount++

	if b.verbosity >= 3 {
		xf.Dump(b.logfile, fmt.Sprintf("--- handle_xf: xf[%d] ---", xf.XFIndex), " ", 0)
	}

	cellType := XL_CELL_NUMBER
	if b.FormatMap == nil {
		b.FormatMap = make(map[int]*Format)
	}
	if fmtObj, ok := b.FormatMap[xf.FormatKey]; ok {
		if ty, ok := cellTypeFromFormatType[fmtObj.Type]; ok {
			cellType = ty
		}
	}
	if b.xfIndexToXLTypeMap == nil {
		b.xfIndexToXLTypeMap = make(map[int]int)
	}
	b.xfIndexToXLTypeMap[xf.XFIndex] = cellType

	if b.formattingInfo {
		if b.verbosity > 0 && xf.IsStyle != 0 && xf.ParentStyleIndex != 0x0FFF {
			fmt.Fprintf(b.logfile,
				"WARNING *** XF[%d] is a style XF but parent_style_index is 0x%04x, not 0x0fff\n",
				xf.XFIndex, xf.ParentStyleIndex)
		}
		checkColourIndexesInObj(b, xf, xf.XFIndex)
	}
	if _, ok := b.FormatMap[xf.FormatKey]; !ok {
		if b.verbosity > 0 {
			fmt.Fprintf(b.logfile,
				"WARNING *** XF[%d] unknown (raw) format key (%d, 0x%04x)\n",
				xf.XFIndex, xf.FormatKey, xf.FormatKey)
		}
		xf.FormatKey = 0
	}
	return nil
}

// handleStyle handles a STYLE record.
func (b *Book) handleStyle(data []byte) error {
	if !b.formattingInfo {
		return nil
	}
	if len(data) < 4 {
		return nil
	}
	if b.Encoding == "" {
		b.Encoding = b.deriveEncoding()
	}

	bv := b.BiffVersion
	flagAndXfx := binary.LittleEndian.Uint16(data[0:2])
	builtInID := int(data[2])
	level := int(data[3])
	xfIndex := int(flagAndXfx & 0x0fff)

	builtIn := 0
	name := ""
	if len(data) >= 4 && data[0] == 0 && data[1] == 0 && data[2] == 0 && data[3] == 0 {
		if _, ok := b.StyleNameMap["Normal"]; !ok {
			builtIn = 1
			builtInID = 0
			xfIndex = 0
			name = "Normal"
			level = 255
		}
	} else if flagAndXfx&0x8000 != 0 {
		builtIn = 1
		if builtInID >= 0 && builtInID < len(builtInStyleNames) {
			name = builtInStyleNames[builtInID]
		}
		if builtInID == 1 || builtInID == 2 {
			name = fmt.Sprintf("%s%d", name, level+1)
		}
	} else {
		builtIn = 0
		builtInID = 0
		level = 0
		var err error
		if bv >= 80 {
			name, err = UnpackUnicode(data, 2, 2)
		} else {
			name, err = UnpackString(data, 2, b.Encoding, 1)
		}
		if err != nil {
			return err
		}
		if b.verbosity >= 2 && name == "" {
			fmt.Fprintln(b.logfile, "WARNING *** A user-defined style has a zero-length name")
		}
	}

	b.StyleNameMap[name] = [2]int{builtIn, xfIndex}
	if b.verbosity >= 2 {
		fmt.Fprintf(b.logfile,
			"STYLE: built_in=%d xf_index=%d built_in_id=%d level=%d name=%q\n",
			builtIn, xfIndex, builtInID, level, name)
	}
	return nil
}

// handlePalette handles a PALETTE record.
func (b *Book) handlePalette(data []byte) error {
	if !b.formattingInfo {
		return nil
	}
	if len(data) < 2 {
		return nil
	}
	if b.ColourMap == nil {
		b.ColourMap = make(map[int][3]int)
	}

	pos := 0
	// Number of colors (2 bytes)
	numColors := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2
	expectedSize := 4*numColors + 2
	if len(data) < expectedSize || len(data) > expectedSize+4 {
		return NewXLRDError("PALETTE record: expected size %d, actual size %d", expectedSize, len(data))
	}

	expectedColors := 16
	if b.BiffVersion >= 50 {
		expectedColors = 56
	}
	if b.verbosity >= 1 && numColors != expectedColors {
		fmt.Fprintf(b.logfile,
			"NOTE *** Expected %d colours in PALETTE record, found %d\n",
			expectedColors, numColors)
	} else if b.verbosity >= 2 {
		fmt.Fprintf(b.logfile, "PALETTE record with %d colours\n", numColors)
	}
	if len(b.PaletteRecord) != 0 {
		return NewXLRDError("PALETTE record: multiple palette records found")
	}

	b.PaletteRecord = make([][3]int, 0, numColors)

	// Each color is 4 bytes: RGB + reserved
	for i := 0; i < numColors && pos+4 <= len(data); i++ {
		r := int(data[pos])
		g := int(data[pos+1])
		b_val := int(data[pos+2])
		oldRGB := b.ColourMap[8+i]
		b.PaletteRecord = append(b.PaletteRecord, [3]int{r, g, b_val})
		b.ColourMap[8+i] = [3]int{r, g, b_val}
		if b.verbosity >= 2 {
			newRGB := [3]int{r, g, b_val}
			if newRGB != oldRGB {
				fmt.Fprintf(b.logfile, "%2d: %v -> %v\n", i, oldRGB, newRGB)
			}
		}
		pos += 4
	}

	return nil
}

// handleWriteAccess handles a WRITEACCESS record.
func (b *Book) handleWriteAccess(data []byte) {
	if len(data) == 0 {
		return
	}

	// For BIFF8, it's UTF-16LE encoded
	if b.BiffVersion >= 80 {
		// Simple conversion - in practice this would need proper UTF-16 handling
		b.UserName = string(data)
	} else {
		b.UserName = string(data)
	}
}

// handleName handles a NAME record.
func (b *Book) handleName(data []byte) error {
	if len(data) < 14 {
		return nil
	}

	name := &Name{
		Book: b,
	}
	pos := 0

	// Options (2 bytes)
	options := binary.LittleEndian.Uint16(data[pos : pos+2])
	pos += 2

	if (options & 0x0001) != 0 {
		name.Hidden = 1
	}
	name.Func = int((options & 0x0002) >> 1)
	name.VBasic = int((options & 0x0004) >> 2)
	name.Macro = int((options & 0x0008) >> 3)

	// Keyboard shortcut (1 byte)
	pos++

	// Length of name (1 byte)
	nameLen := int(data[pos])
	pos++

	// Length of formula (2 bytes)
	pos += 2

	// Unused (2 bytes)
	pos += 2

	// Sheet index (2 bytes)
	name.Scope = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Name string
	if pos+nameLen <= len(data) {
		if b.BiffVersion >= 80 {
			name.Name = string(data[pos : pos+nameLen])
		} else {
			name.Name = string(data[pos : pos+nameLen])
		}
	}

	b.NameObjList = append(b.NameObjList, name)
	return nil
}

// handleBuiltinfmtcount handles a BUILTINFMTCOUNT record.
func (b *Book) handleBuiltinfmtcount(data []byte) {
	// N.B. This count appears to be utterly useless.
	if len(data) >= 2 {
		b.builtinfmtcount = int(binary.LittleEndian.Uint16(data[0:2]))
		if b.verbosity >= 2 {
			fmt.Fprintf(b.logfile, "BUILTINFMTCOUNT: %d\n", b.builtinfmtcount)
		}
	}
}

// handleFilepass handles a FILEPASS record (file encryption).
func (b *Book) handleFilepass(data []byte) error {
	if b.verbosity >= 2 {
		logf := b.logfile
		fmt.Fprintf(logf, "FILEPASS:\n")
		// Note: hex dump implementation would be needed here
		if b.BiffVersion >= 80 {
			if len(data) >= 2 {
				kind1 := binary.LittleEndian.Uint16(data[:2])
				if kind1 == 0 { // weak XOR encryption
					if len(data) >= 6 {
						key := binary.LittleEndian.Uint16(data[2:4])
						hashValue := binary.LittleEndian.Uint16(data[4:6])
						fmt.Fprintf(logf, "weak XOR: key=0x%04x hash=0x%04x\n", key, hashValue)
					}
				} else if kind1 == 1 {
					if len(data) >= 6 {
						kind2 := binary.LittleEndian.Uint16(data[4:6])
						var caption string
						if kind2 == 1 { // BIFF8 standard encryption
							caption = "BIFF8 std"
						} else if kind2 == 2 {
							caption = "BIFF8 strong"
						} else {
							caption = "** UNKNOWN ENCRYPTION METHOD **"
						}
						fmt.Fprintf(logf, "%s\n", caption)
					}
				}
			}
		}
	}
	return NewXLRDError("Workbook is encrypted")
}

// handleObj handles an OBJ record.
// Not doing much handling at all. Worrying about embedded (BOF ... EOF) substreams is done elsewhere.
func (b *Book) handleObj(data []byte) {
	// Not doing much handling at all.
	// Worrying about embedded (BOF ... EOF) substreams is done elsewhere.
	if len(data) >= 10 {
		objType := binary.LittleEndian.Uint16(data[4:6])
		objId := binary.LittleEndian.Uint32(data[6:10])
		// Debug print would go here if needed
		_ = objType
		_ = objId
	}
}

// handleSheethdr handles a SHEETHDR record (BIFF 4W special).
func (b *Book) handleSheethdr(data []byte) error {
	// This a BIFF 4W special.
	// The SHEETHDR record is followed by a (BOF ... EOF) substream containing a worksheet.
	if len(data) < 4 {
		return NewXLRDError("SHEETHDR record too short")
	}

	sheetLen := int(binary.LittleEndian.Uint32(data[:4]))
	sheetName, err := UnpackString(data, 4, b.Encoding, 1)
	if err != nil {
		return err
	}

	sheetno := b.sheethdrCount
	if sheetName != b.sheetNames[sheetno] {
		return NewXLRDError("SHEETHDR name mismatch")
	}
	b.sheethdrCount++

	BOFPosn := b.position
	// posn := BOFPosn - 4 - len(data) // Not used in Go implementation

	b.initializeFormatInfo()
	b.sheetList = append(b.sheetList, nil) // get_sheet updates _sheet_list but needs a None beforehand
	_, err = b.getSheet(sheetno, false)
	if err != nil {
		return err
	}

	b.position = BOFPosn + sheetLen
	return nil
}

// handleSheetsoffset handles a SHEETSOFFSET record.
func (b *Book) handleSheetsoffset(data []byte) {
	if len(data) >= 4 {
		posn := int(binary.LittleEndian.Uint32(data[:4]))
		b.sheetsoffset = posn
	}
}

// handleSST handles a Shared String Table record.
func (b *Book) handleSST(data []byte) error {
	if len(data) < 8 {
		return nil // Not enough data
	}

	// Number of unique strings (BIFF8 SST header)
	numStrings := int(binary.LittleEndian.Uint32(data[4:8]))

	strlist := [][]byte{data}
	for {
		code, _, cont := b.getRecordPartsConditional(XL_CONTINUE)
		if code == 0 {
			break
		}
		strlist = append(strlist, cont)
	}

	shared, richtextRuns := UnpackSSTTable(strlist, numStrings)
	b.sharedStrings = shared
	b.richTextRunlistMap = richtextRuns
	return nil
}

// handleSupbook handles a SUPBOOK record (external book references).
func (b *Book) handleSupbook(data []byte) error {
	if len(data) < 2 {
		return nil // Not enough data
	}

	b.supbookTypes = append(b.supbookTypes, SUPBOOK_UNK)

	numSheets := int(binary.LittleEndian.Uint16(data[0:2]))
	b.supbookCount++

	// Check for internal 3D references
	if len(data) >= 4 && data[2] == 0x01 && data[3] == 0x04 {
		b.supbookTypes[len(b.supbookTypes)-1] = SUPBOOK_INTERNAL
		supbookLocalsInx := b.supbookCount - 1
		b.supbookLocalsInx = &supbookLocalsInx
		return nil
	}

	// Check for add-in functions
	if len(data) >= 4 && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x01 && data[3] == 0x3A {
		b.supbookTypes[len(b.supbookTypes)-1] = SUPBOOK_ADDIN
		supbookAddinsInx := b.supbookCount - 1
		b.supbookAddinsInx = &supbookAddinsInx
		return nil
	}

	// Parse URL for external references
	_, pos, err := UnpackUnicodeUpdatePos(data, 2, 2, nil)
	if err != nil {
		return err
	}

	if numSheets == 0 {
		// DDE/OLE document
		b.supbookTypes[len(b.supbookTypes)-1] = SUPBOOK_DDEOLE
		return nil
	}

	// External book
	b.supbookTypes[len(b.supbookTypes)-1] = SUPBOOK_EXTERNAL

	// Parse sheet names (simplified - not handling all edge cases)
	for i := 0; i < numSheets && pos < len(data); i++ {
		_, newPos, err := UnpackUnicodeUpdatePos(data, pos, 2, nil)
		if err != nil {
			break
		}
		pos = newPos
	}

	return nil
}

// handleExternname handles an EXTERNNAME record (external names).
func (b *Book) handleExternname(data []byte) error {
	if len(data) < 6 {
		return nil
	}

	if b.BiffVersion >= 80 {
		pos := 6
		name, _, err := UnpackUnicodeUpdatePos(data, pos, 1, nil)
		if err != nil {
			return err
		}

		// Check if this is from an add-in supbook
		if len(b.supbookTypes) > 0 && b.supbookTypes[len(b.supbookTypes)-1] == SUPBOOK_ADDIN {
			b.addinFuncNames = append(b.addinFuncNames, name)
		}
	}

	return nil
}

// handleExternsheet handles an EXTERNSHEET record (external sheet references).
func (b *Book) handleExternsheet(data []byte) error {
	b.extnshtCount++

	if b.BiffVersion >= 80 {
		// BIFF 8.0 and later
		if len(data) < 2 {
			return nil
		}
		numRefs := int(binary.LittleEndian.Uint16(data[0:2]))
		pos := 2

		for i := 0; i < numRefs && pos+6 <= len(data); i++ {
			refRecordx := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
			refFirstSheetx := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
			refLastSheetx := int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))

			info := []int{refRecordx, refFirstSheetx, refLastSheetx}
			b.externsheetInfo = append(b.externsheetInfo, info)
			pos += 6
		}
	} else {
		// BIFF 7 and earlier
		if len(data) >= 2 {
			nc := int(data[0])
			ty := int(data[1])
			b.externsheetTypeB57 = append(b.externsheetTypeB57, ty)

			if ty == 3 && len(data) >= nc+2 {
				sheetName := string(data[2 : nc+2])
				b.extnshtNameFromNum[b.extnshtCount] = sheetName
			}
		}
	}

	return nil
}

// deriveEncoding derives the encoding from the codepage.
func (b *Book) deriveEncoding() string {
	if b.encodingOverride != "" {
		return b.encodingOverride
	}
	if b.Codepage == nil {
		if b.BiffVersion < 80 {
			return "iso-8859-1"
		}
		codepage := 1200
		b.Codepage = &codepage
		return "utf_16_le"
	}

	codepage := *b.Codepage
	if enc, ok := EncodingFromCodepage[codepage]; ok {
		return enc
	}

	if codepage >= 300 && codepage <= 1999 {
		return fmt.Sprintf("cp%d", codepage)
	}

	if b.BiffVersion >= 80 {
		codepage = 1200
		b.Codepage = &codepage
		return "utf_16_le"
	}

	return fmt.Sprintf("unknown_codepage_%d", codepage)
}

// initializeFormatInfo initializes format information structures.
func (b *Book) initializeFormatInfo() {
	b.FormatMap = make(map[int]*Format)
	b.FormatList = make([]*Format, 0)
	b.XFList = make([]*XF, 0)
	b.FontList = make([]*Font, 0)
	b.StyleNameMap = make(map[string][2]int)
	b.PaletteRecord = make([][3]int, 0)
	b.ColourIndexesUsed = make(map[int]bool)
	b.actualFmtCount = 0
	b.xfCount = 0
	b.xfEpilogueDone = false
	b.xfIndexToXLTypeMap = map[int]int{0: XL_CELL_NUMBER}
	initialiseColourMap(b)
}

// fakeGlobalsGetSheet handles BIFF 4.0 and earlier (no workbook globals).
func (b *Book) fakeGlobalsGetSheet() {
	// For BIFF 4.0 and earlier, there's only one worksheet
	b.sheetNames = []string{"Sheet1"}
	b.sheetList = []*Sheet{nil}
	b.sheetAbsPosn = []int{b.base}
	b.sheetVisibility = []int{0}
	b.NSheets = 1
}

// getSheet loads a sheet by its index.
func (b *Book) getSheet(shNumber int, updatePos ...bool) (*Sheet, error) {
	updatePosition := true
	if len(updatePos) > 0 {
		updatePosition = updatePos[0]
	}
	if shNumber < 0 || shNumber >= len(b.sheetNames) {
		return nil, NewXLRDError("sheet index %d out of range", shNumber)
	}

	// Set position to sheet's absolute position
	if shNumber >= len(b.sheetAbsPosn) {
		return nil, NewXLRDError("sheet position not found for sheet %d", shNumber)
	}

	if updatePosition {
		b.position = b.sheetAbsPosn[shNumber]
	}

	// Get BOF record for worksheet
	_, err := b.getBOF(XL_WORKSHEET)
	if err != nil {
		return nil, err
	}

	// Create sheet
	sheet := &Sheet{
		Book:               b,
		Name:               b.sheetNames[shNumber],
		ColInfoMap:         make(map[int]*ColInfo),
		RowInfoMap:         make(map[int]*RowInfo),
		ColLabelRanges:     make([][4]int, 0),
		RowLabelRanges:     make([][4]int, 0),
		MergedCells:        make([][4]int, 0),
		HyperlinkList:      make([]*Hyperlink, 0),
		HyperlinkMap:       make(map[[2]int]*Hyperlink),
		CellNoteMap:        make(map[[2]int]*Note),
		RichTextRunlistMap: make(map[[2]int][][]int),
		cellAttrToXF:       make(map[[3]byte]int),
		ixfe:               -1,
	}

	// Read sheet data
	err = sheet.read(b)
	if err != nil {
		return nil, err
	}

	return sheet, nil
}

// getSheets loads all sheets in the workbook.
func (b *Book) getSheets() error {
	for sheetNo := 0; sheetNo < len(b.sheetNames); sheetNo++ {
		sheet, err := b.getSheet(sheetNo)
		if err != nil {
			return err
		}
		b.sheetList[sheetNo] = sheet
	}
	return nil
}

// readWorksheets reads all worksheets in the workbook.
func (b *Book) readWorksheets(options *OpenWorkbookOptions) error {
	for sheetNo := 0; sheetNo < len(b.sheetNames); sheetNo++ {
		sheet, err := b.getSheet(sheetNo)
		if err != nil {
			return err
		}
		b.sheetList[sheetNo] = sheet
	}
	return nil
}

// get2bytes reads 2 bytes from the current position and advances the position.
func (b *Book) get2bytes() int {
	if b.position+2 > len(b.mem) {
		return MY_EOF
	}
	result := int(binary.LittleEndian.Uint16(b.mem[b.position : b.position+2]))
	b.position += 2
	return result
}

// getRecordPartsConditional reads a record only if it matches the required record type.
func (b *Book) getRecordPartsConditional(reqdRecord int) (int, int, []byte) {
	if b.position+4 > len(b.mem) {
		return 0, 0, nil
	}
	code := int(binary.LittleEndian.Uint16(b.mem[b.position : b.position+2]))
	length := int(binary.LittleEndian.Uint16(b.mem[b.position+2 : b.position+4]))
	if code != reqdRecord {
		return 0, 0, nil
	}
	b.position += 4
	if b.position+length > len(b.mem) {
		return code, 0, nil
	}
	data := b.mem[b.position : b.position+length]
	b.position += length
	return code, length, data
}

// read reads data from the specified position and advances the current position.
func (b *Book) read(pos, length int) []byte {
	if pos+length > len(b.mem) {
		length = len(b.mem) - pos
	}
	data := b.mem[pos : pos+length]
	b.position = pos + len(data)
	return data
}

// biff2_8_load loads BIFF data from file or file contents.
// This method handles the core file loading logic.
func (b *Book) biff2_8_load(filename string, fileContents []byte,
	logfile io.Writer, verbosity int, useMmap bool,
	encodingOverride string,
	formattingInfo bool,
	onDemand bool,
	raggedRows bool,
	ignoreWorkbookCorruption bool) error {

	b.logfile = logfile
	b.verbosity = verbosity
	b.formattingInfo = formattingInfo
	b.onDemand = onDemand
	b.raggedRows = raggedRows
	b.encodingOverride = encodingOverride
	b.ignoreWorkbookCorruption = ignoreWorkbookCorruption

	if fileContents == nil {
		// Read from file
		content, err := os.ReadFile(filename)
		if err != nil {
			return err
		}
		if len(content) == 0 {
			return NewXLRDError("File size is 0 bytes")
		}
		b.filestr = content
		b.streamLen = len(content)
	} else {
		b.filestr = fileContents
		b.streamLen = len(fileContents)
	}

	b.base = 0
	if len(b.filestr) >= 8 && string(b.filestr[:8]) == string(XLS_SIGNATURE) {
		// OLE2 compound document
		cd, err := NewCompDoc(b.filestr, logfile, 0, ignoreWorkbookCorruption)
		if err != nil {
			return err
		}

		// Try to locate Workbook or Book stream
		var mem []byte
		var base, streamLen int
		var lastErr error
		for _, qname := range []string{"Workbook", "Book"} {
			mem, base, streamLen, err = cd.LocateNamedStream(qname)
			if err == nil && mem != nil {
				break
			}
			if err != nil {
				lastErr = err
				if compDocErr, ok := err.(*CompDocError); ok && !ignoreWorkbookCorruption {
					return compDocErr
				}
			}
		}

		if mem == nil {
			if lastErr != nil {
				return lastErr
			}
			return NewXLRDError("Can't find workbook in OLE2 compound document")
		}

		b.mem = mem
		b.base = base
		b.streamLen = streamLen
	} else {
		// Not an OLE2 compound document - treat as raw BIFF file
		b.mem = b.filestr
		b.base = 0
		b.streamLen = len(b.filestr)
	}

	b.position = b.base
	return nil
}

// namesEpilogue processes NAME records and builds mapping dictionaries.
func (b *Book) namesEpilogue() {
	if b.verbosity >= 2 {
		fmt.Fprintf(b.logfile, "+++++ names_epilogue +++++\n")
	}

	// Note: Scope handling is already done during handleName
	// Additional BIFF version-specific handling could be added here if needed

	// Build mapping dictionaries
	b.nameAndScopeMap = make(map[string]map[int]*Name)
	b.nameMap = make(map[string][]*Name)

	for namex := 0; namex < len(b.NameObjList); namex++ {
		nobj := b.NameObjList[namex]
		nameLcase := strings.ToLower(nobj.Name)

		// name_and_scope_map: (name.lower(), scope) -> Name object
		if b.nameAndScopeMap[nameLcase] == nil {
			b.nameAndScopeMap[nameLcase] = make(map[int]*Name)
		}
		if _, exists := b.nameAndScopeMap[nameLcase][nobj.Scope]; exists && b.verbosity >= 1 {
			fmt.Fprintf(b.logfile, "Duplicate entry (%s, %d) in name_and_scope_map\n", nameLcase, nobj.Scope)
		}
		b.nameAndScopeMap[nameLcase][nobj.Scope] = nobj

		// name_map: name.lower() -> list of Name objects (sorted by scope)
		if b.nameMap[nameLcase] == nil {
			b.nameMap[nameLcase] = []*Name{}
		}
		// Insert in sorted order by scope
		inserted := false
		for i, existing := range b.nameMap[nameLcase] {
			if existing.Scope > nobj.Scope {
				// Insert before this element
				b.nameMap[nameLcase] = append(b.nameMap[nameLcase][:i], append([]*Name{nobj}, b.nameMap[nameLcase][i:]...)...)
				inserted = true
				break
			}
		}
		if !inserted {
			b.nameMap[nameLcase] = append(b.nameMap[nameLcase], nobj)
		}
	}

	if b.verbosity >= 2 {
		fmt.Fprintf(b.logfile, "---------- name object dump ----------\n")
		for namex := 0; namex < len(b.NameObjList); namex++ {
			nobj := b.NameObjList[namex]
			fmt.Fprintf(b.logfile, "--- name[%d]: %s ---\n", namex, nobj.Name)
		}
		fmt.Fprintf(b.logfile, "--------------------------------------\n")
	}
}

// xfEpilogue processes extended format information after all XF records are read.
func (b *Book) xfEpilogue() {
	b.xfEpilogueDone = true
	numXfs := len(b.XFList)
	if b.verbosity >= 3 {
		fmt.Fprintln(b.logfile, "xf_epilogue called ...")
	}

	checkSame := func(xf *XF, parent *XF, attr string) {
		if !reflect.DeepEqual(reflect.ValueOf(xf).Elem().FieldByName(attr).Interface(),
			reflect.ValueOf(parent).Elem().FieldByName(attr).Interface()) {
			fmt.Fprintf(b.logfile,
				"NOTE !!! XF[%d] parent[%d] %s different\n",
				xf.XFIndex, parent.XFIndex, attr)
		}
	}

	for xfx := 0; xfx < numXfs; xfx++ {
		xf := b.XFList[xfx]
		cellType := XL_CELL_TEXT
		if fmtObj, ok := b.FormatMap[xf.FormatKey]; ok {
			if ty, ok := cellTypeFromFormatType[fmtObj.Type]; ok {
				cellType = ty
			}
		}
		b.xfIndexToXLTypeMap[xf.XFIndex] = cellType

		if !b.formattingInfo {
			continue
		}
		if xf.IsStyle != 0 {
			continue
		}
		if xf.ParentStyleIndex < 0 || xf.ParentStyleIndex >= numXfs {
			if b.verbosity >= 1 {
				fmt.Fprintf(b.logfile,
					"WARNING *** XF[%d]: is_style=%d but parent_style_index=%d\n",
					xf.XFIndex, xf.IsStyle, xf.ParentStyleIndex)
			}
			xf.ParentStyleIndex = 0
		}
		if b.BiffVersion >= 30 {
			if b.verbosity >= 1 {
				if xf.ParentStyleIndex == xf.XFIndex {
					fmt.Fprintf(b.logfile,
						"NOTE !!! XF[%d]: parent_style_index is also %d\n",
						xf.XFIndex, xf.ParentStyleIndex)
				} else if b.XFList[xf.ParentStyleIndex].IsStyle == 0 {
					fmt.Fprintf(b.logfile,
						"NOTE !!! XF[%d]: parent_style_index is %d; style flag not set\n",
						xf.XFIndex, xf.ParentStyleIndex)
				}
				if xf.ParentStyleIndex > xf.XFIndex {
					fmt.Fprintf(b.logfile,
						"NOTE !!! XF[%d]: parent_style_index is %d; out of order?\n",
						xf.XFIndex, xf.ParentStyleIndex)
				}
			}
			parent := b.XFList[xf.ParentStyleIndex]
			if xf.AlignmentFlag == 0 && parent.AlignmentFlag == 0 {
				if b.verbosity >= 1 {
					checkSame(xf, parent, "Alignment")
				}
			}
			if xf.BackgroundFlag == 0 && parent.BackgroundFlag == 0 {
				if b.verbosity >= 1 {
					checkSame(xf, parent, "Background")
				}
			}
			if xf.BorderFlag == 0 && parent.BorderFlag == 0 {
				if b.verbosity >= 1 {
					checkSame(xf, parent, "Border")
				}
			}
			if xf.ProtectionFlag == 0 && parent.ProtectionFlag == 0 {
				if b.verbosity >= 1 {
					checkSame(xf, parent, "Protection")
				}
			}
			if xf.FormatFlag == 0 && parent.FormatFlag == 0 && b.verbosity >= 1 {
				if xf.FormatKey != parent.FormatKey {
					fmt.Fprintf(b.logfile,
						"NOTE !!! XF[%d] fmtk=%d, parent[%d] fmtk=%d\n",
						xf.XFIndex, xf.FormatKey, parent.XFIndex, parent.FormatKey)
				}
			}
			if xf.FontFlag == 0 && parent.FontFlag == 0 && b.verbosity >= 1 {
				if xf.FontIndex != parent.FontIndex {
					fmt.Fprintf(b.logfile,
						"NOTE !!! XF[%d] fontx=%d, parent[%d] fontx=%d\n",
						xf.XFIndex, xf.FontIndex, parent.XFIndex, parent.FontIndex)
				}
			}
		}
	}
	b.xfEpilogueDone = true
}

// paletteEpilogue processes palette information after all records are read.
func (b *Book) paletteEpilogue() {
	if !b.formattingInfo {
		return
	}
	for _, font := range b.FontList {
		if font.FontIndex == 4 {
			continue
		}
		cx := font.ColourIndex
		if cx == 0x7FFF {
			continue
		}
		if _, ok := b.ColourMap[cx]; ok {
			b.ColourIndexesUsed[cx] = true
			continue
		}
		if b.verbosity > 0 {
			fmt.Fprintf(b.logfile, "Size of colour table: %d\n", len(b.ColourMap))
			fmt.Fprintf(b.logfile,
				"*** Font #%d (%q): colour index 0x%04x is unknown\n",
				font.FontIndex, font.Name, cx)
		}
	}
	if b.verbosity >= 1 {
		used := make([]int, 0, len(b.ColourIndexesUsed))
		for k := range b.ColourIndexesUsed {
			used = append(used, k)
		}
		sort.Ints(used)
		fmt.Fprintf(b.logfile, "\nColour indexes used:\n%v\n", used)
	}
}

// getRecordParts reads the next BIFF record from the current position.
func (b *Book) getRecordParts() (int, int, []byte) {
	if b.position+4 > len(b.mem) {
		return 0, 0, nil
	}
	code := int(binary.LittleEndian.Uint16(b.mem[b.position : b.position+2]))
	length := int(binary.LittleEndian.Uint16(b.mem[b.position+2 : b.position+4]))
	b.position += 4
	if b.position+length > len(b.mem) {
		return code, 0, nil
	}
	data := b.mem[b.position : b.position+length]
	b.position += length
	return code, length, data
}

// ExpandCellAddress expands a cell address from BIFF format.
// Ref: OOo docs, "4.3.4 Cell Addresses in BIFF8"
// Returns: outrow, outcol, relrow, relcol
func ExpandCellAddress(inrow, incol int) (int, int, int, int) {
	outrow := inrow
	var relrow int
	if incol&0x8000 != 0 {
		if outrow >= 32768 {
			outrow -= 65536
		}
		relrow = 1
	} else {
		relrow = 0
	}

	outcol := incol & 0xFF
	var relcol int
	if incol&0x4000 != 0 {
		if outcol >= 128 {
			outcol -= 256
		}
		relcol = 1
	} else {
		relcol = 0
	}

	return outrow, outcol, relrow, relcol
}

// UnpackSSTTable unpacks the Shared String Table from SST record data.
// Returns list of strings and rich text run information.
func UnpackSSTTable(datatab [][]byte, nstrings int) ([]string, map[int][][]int) {
	if len(datatab) == 0 {
		return []string{}, make(map[int][][]int)
	}

	datainx := 0
	ndatas := len(datatab)
	data := datatab[0]
	datalen := len(data)
	pos := 8

	strings := make([]string, 0, nstrings)
	richtextRuns := make(map[int][][]int)

	for i := 0; i < nstrings; i++ {
		if pos+2 > datalen {
			break
		}

		nchars := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2
		if pos >= datalen {
			break
		}

		options := data[pos]
		pos++

		rtcount := 0
		phosz := 0
		if options&0x08 != 0 { // richtext
			rtcount = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
			pos += 2
		}
		if options&0x04 != 0 { // phonetic
			phosz = int(binary.LittleEndian.Uint32(data[pos : pos+4]))
			pos += 4
		}

		accstrg := ""
		charsgot := 0
		for charsgot < nchars {
			charsneed := nchars - charsgot
			charsavail := 0
			if options&0x01 != 0 {
				// Uncompressed UTF-16
				charsavail = min((datalen-pos)>>1, charsneed)
				rawstrg := data[pos : pos+2*charsavail]
				words := make([]uint16, charsavail)
				for j := 0; j < charsavail; j++ {
					words[j] = binary.LittleEndian.Uint16(rawstrg[j*2 : (j+1)*2])
				}
				accstrg += string(utf16.Decode(words))
				pos += 2 * charsavail
			} else {
				// Compressed (Latin-1)
				charsavail = min(datalen-pos, charsneed)
				rawstrg := data[pos : pos+charsavail]
				utf8Bytes, err := charmap.ISO8859_1.NewDecoder().Bytes(rawstrg)
				if err != nil {
					accstrg += string(rawstrg)
				} else {
					accstrg += string(utf8Bytes)
				}
				pos += charsavail
			}

			charsgot += charsavail
			if charsgot == nchars {
				break
			}

			datainx++
			if datainx >= ndatas {
				break
			}
			data = datatab[datainx]
			datalen = len(data)
			options = data[0]
			pos = 1
		}

		if rtcount > 0 {
			runs := make([][]int, 0, rtcount)
			for runindex := 0; runindex < rtcount; runindex++ {
				if pos == datalen {
					pos = 0
					datainx++
					if datainx >= ndatas {
						break
					}
					data = datatab[datainx]
					datalen = len(data)
				}
				run1 := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
				run2 := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
				runs = append(runs, []int{run1, run2})
				pos += 4
			}
			richtextRuns[len(strings)] = runs
		}

		pos += phosz
		if pos >= datalen {
			pos = pos - datalen
			datainx++
			if datainx < ndatas {
				data = datatab[datainx]
				datalen = len(data)
			}
		}

		strings = append(strings, accstrg)
	}

	return strings, richtextRuns
}

// Iter returns an iterator over all sheets in the book.
// This provides Python-like iteration: for sheet := range book.Iter()
func (b *Book) Iter() <-chan *Sheet {
	ch := make(chan *Sheet)
	go func() {
		defer close(ch)
		for i := 0; i < b.NSheets; i++ {
			sheet, err := b.SheetByIndex(i)
			if err == nil {
				ch <- sheet
			}
		}
	}()
	return ch
}

// Enter implements context manager enter (Python __enter__).
// Returns the book itself for use in with statements.
func (b *Book) Enter() *Book {
	return b
}

// Exit implements context manager exit (Python __exit__).
// Automatically releases resources.
func (b *Book) Exit() {
	b.ReleaseResources()
}

// Colname returns the column name for a given column index (0-based).
// Example: Colname(0) returns "A", Colname(25) returns "Z", Colname(26) returns "AA"
func Colname(colx int) string {
	if colx < 0 {
		return ""
	}

	const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	name := ""
	for {
		quot := colx / 26
		rem := colx % 26
		name = string(alphabet[rem]) + name
		if quot == 0 {
			break
		}
		colx = quot - 1
	}
	return name
}
