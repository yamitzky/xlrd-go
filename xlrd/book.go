package xlrd

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"
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
		case XL_FORMAT:
			err := b.handleFormat(data)
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
	if len(data) < 14 {
		return nil // Not enough data
	}

	font := &Font{}
	pos := 0

	// Height (2 bytes, little endian)
	font.Height = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Option flags (2 bytes)
	options := binary.LittleEndian.Uint16(data[pos : pos+2])
	pos += 2

	font.Bold = (options & 0x0001) != 0
	font.Italic = (options & 0x0002) != 0
	font.Underline = int((options & 0x000C) >> 2)
	font.Escapement = int((options & 0x0070) >> 4)

	// Colour index (2 bytes)
	font.ColourIndex = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Weight (2 bytes)
	font.Weight = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Escapement (1 byte, already handled in options)
	pos += 1

	// Underline (1 byte, already handled in options)
	pos += 1

	// Family (1 byte)
	font.Family = int(data[pos])
	pos += 1

	// Character set (1 byte)
	font.CharacterSet = int(data[pos])
	pos += 1

	// Reserved (1 byte)
	pos += 1

	// Font name length (1 byte)
	nameLen := int(data[pos])
	pos += 1

	// Font name
	if pos+nameLen <= len(data) {
		// For BIFF8, font name is UTF-16LE encoded
		if b.BiffVersion >= 80 {
			font.Name = string(data[pos : pos+nameLen])
		} else {
			font.Name = string(data[pos : pos+nameLen])
		}
	}

	b.FontList = append(b.FontList, font)
	return nil
}

// handleFormat handles a FORMAT record.
func (b *Book) handleFormat(data []byte) error {
	if len(data) < 2 {
		return nil
	}

	format := &Format{}
	pos := 0

	// Format index (2 bytes, but we use the list position)
	format.FormatKey = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Format string
	if b.BiffVersion >= 80 {
		// BIFF8: UTF-16LE encoded
		if pos < len(data) {
			formatString, err := UnpackUnicode(data, pos, 0)
			if err == nil {
				format.FormatString = formatString
			}
		}
	} else {
		// Earlier versions: byte string
		if pos < len(data) {
			strLen := int(data[pos])
			pos++
			if pos+strLen <= len(data) {
				format.FormatString = string(data[pos : pos+strLen])
			}
		}
	}

	b.FormatList = append(b.FormatList, format)
	if b.FormatMap != nil {
		b.FormatMap[format.FormatKey] = format
	}
	return nil
}

// handleXF handles an XF (Extended Format) record.
func (b *Book) handleXF(data []byte) error {
	if len(data) < 16 {
		return nil
	}

	xf := &XF{}
	pos := 0

	// Font index (2 bytes)
	xf.FontIndex = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Format key (2 bytes)
	xf.FormatKey = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	// Protection flags and other options (2 bytes)
	protection := binary.LittleEndian.Uint16(data[pos : pos+2])
	pos += 2

	xf.Locked = (protection & 0x0001) != 0
	xf.Hidden = (protection & 0x0002) != 0

	// Alignment options (1 byte)
	if pos < len(data) {
		xf.Alignment = &XFAlignment{
			Horizontal: int(data[pos] & 0x07),
			Vertical:   int((data[pos] & 0x70) >> 4),
		}
		pos++
	}

	// Fill/rotation options (1 byte)
	if pos < len(data) {
		pos++ // Skip for now
	}

	// Border options (4 bytes)
	if pos+4 <= len(data) {
		xf.Border = &XFBorder{}
		pos += 4
	}

	// Background options (2 bytes)
	if pos+2 <= len(data) {
		xf.Background = &XFBackground{}
		pos += 2
	}

	b.XFList = append(b.XFList, xf)
	return nil
}

// handleStyle handles a STYLE record.
func (b *Book) handleStyle(data []byte) error {
	if len(data) < 2 {
		return nil
	}

	// For now, we just skip style records
	// Style information is complex and not always needed
	return nil
}

// handlePalette handles a PALETTE record.
func (b *Book) handlePalette(data []byte) error {
	if len(data) < 2 {
		return nil
	}

	pos := 0
	// Number of colors (2 bytes)
	numColors := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
	pos += 2

	b.PaletteRecord = make([][3]int, 0, numColors)

	// Each color is 4 bytes: RGB + reserved
	for i := 0; i < numColors && pos+4 <= len(data); i++ {
		r := int(data[pos])
		g := int(data[pos+1])
		b_val := int(data[pos+2])
		b.PaletteRecord = append(b.PaletteRecord, [3]int{r, g, b_val})
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

	// Number of unique strings
	numStrings := int(binary.LittleEndian.Uint32(data[0:4]))

	b.sharedStrings = make([]string, 0, numStrings)

	pos := 8 // Skip header

	for i := 0; i < numStrings && pos < len(data); i++ {
		if pos+2 > len(data) {
			break
		}

		// Number of characters
		nchars := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2

		if pos >= len(data) {
			break
		}

		// Options byte
		options := data[pos]
		pos++

		// Skip richtext and phonetic info for now
		if (options & 0x08) != 0 { // richtext
			if pos+2 > len(data) {
				break
			}
			pos += 2
		}
		if (options & 0x04) != 0 { // phonetic
			if pos+4 > len(data) {
				break
			}
			pos += 4
		}

		var str string
		if (options & 0x01) != 0 { // Uncompressed UTF-16
			strLen := nchars * 2
			if pos+strLen > len(data) {
				break
			}
			utf16Bytes := data[pos : pos+strLen]
			// Convert UTF-16 LE to string
			words := make([]uint16, nchars)
			for j := 0; j < nchars; j++ {
				words[j] = binary.LittleEndian.Uint16(utf16Bytes[j*2 : (j+1)*2])
			}
			str = string(utf16.Decode(words))
			pos += strLen
		} else { // Compressed (Latin-1)
			strLen := nchars
			if pos+strLen > len(data) {
				break
			}
			// Convert Latin-1 to UTF-8
			latin1Bytes := data[pos : pos+strLen]
			utf8Bytes, err := charmap.ISO8859_1.NewDecoder().Bytes(latin1Bytes)
			if err != nil {
				str = string(latin1Bytes) // fallback
			} else {
				str = string(utf8Bytes)
			}
			pos += strLen
		}

		b.sharedStrings = append(b.sharedStrings, str)
	}

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
	b.ColourMap = make(map[int][3]int)
	b.PaletteRecord = make([][3]int, 0)
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
		Book:           b,
		Name:           b.sheetNames[shNumber],
		ColInfoMap:     make(map[int]*ColInfo),
		RowInfoMap:     make(map[int]*RowInfo),
		ColLabelRanges: make([][4]int, 0),
		RowLabelRanges: make([][4]int, 0),
		MergedCells:    make([][4]int, 0),
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
	// XF epilogue processing - currently minimal implementation
	// In full implementation, this would handle XF record post-processing
	b.xfEpilogueDone = true
}

// paletteEpilogue processes palette information after all records are read.
func (b *Book) paletteEpilogue() {
	// Palette epilogue processing - currently minimal implementation
	// In full implementation, this would finalize palette mappings
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

		// Number of characters
		nchars := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2

		if pos >= datalen {
			break
		}

		// Options byte
		options := data[pos]
		pos++

		rtcount := 0
		if options&0x08 != 0 { // richtext
			if pos+2 > datalen {
				break
			}
			rtcount = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
			pos += 2
		}

		if options&0x04 != 0 { // phonetic
			if pos+4 > datalen {
				break
			}
			pos += 4 // Skip phonetic size
		}

		var accstrg string
		charsgot := 0

		for charsgot < nchars {
			charsneed := nchars - charsgot
			charsavail := 0

			if options&0x01 != 0 {
				// Uncompressed UTF-16
				charsavail = min((datalen-pos)>>1, charsneed)
				if pos+2*charsavail > datalen {
					break
				}
				rawstrg := data[pos : pos+2*charsavail]
				// Convert UTF-16 LE to string
				words := make([]uint16, charsavail)
				for j := 0; j < charsavail; j++ {
					words[j] = binary.LittleEndian.Uint16(rawstrg[j*2 : (j+1)*2])
				}
				accstrg += string(utf16.Decode(words))
				pos += 2 * charsavail
			} else {
				// Compressed (Latin-1)
				charsavail = min(datalen-pos, charsneed)
				if pos+charsavail > datalen {
					break
				}
				rawstrg := data[pos : pos+charsavail]
				// Convert Latin-1 to UTF-8
				utf8Bytes, err := charmap.ISO8859_1.NewDecoder().Bytes(rawstrg)
				if err != nil {
					accstrg += string(rawstrg) // fallback
				} else {
					accstrg += string(utf8Bytes)
				}
				pos += charsavail
			}

			charsgot += charsavail

			if charsgot == nchars {
				break
			}

			// Move to next data block
			datainx++
			if datainx < ndatas {
				data = datatab[datainx]
				datalen = len(data)
				if datalen > 0 {
					options = data[0]
					pos = 1
				}
			} else {
				break
			}
		}

		if rtcount > 0 {
			runs := make([][]int, 0, rtcount)
			for runindex := 0; runindex < rtcount; runindex++ {
				if pos+4 > datalen {
					break
				}
				run1 := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
				run2 := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
				runs = append(runs, []int{run1, run2})
				pos += 4
			}
			richtextRuns[len(strings)] = runs
		}

		// Skip phonetic data
		if options&0x04 != 0 {
			// Skip remaining phonetic data if any
			for pos >= datalen && datainx+1 < ndatas {
				pos -= datalen
				datainx++
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
