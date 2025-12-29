package xlrd

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
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
	sheetList        []*Sheet
	sheetNames       []string
	sheetAbsPosn     []int // Absolute positions of sheets in the stream
	sheetVisibility  []int
	onDemand         bool
	logfile          io.Writer
	verbosity        int
	mem              []byte
	base              int
	streamLen         int
	position          int
	filestr           []byte
	formattingInfo    bool
	raggedRows        bool
	encodingOverride  string
	sharedStrings     []string
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
	for sheetx := 0; sheetx < len(b.sheetList); sheetx++ {
		if b.sheetList[sheetx] == nil {
			sheet, err := b.getSheet(sheetx)
			if err == nil {
				b.sheetList[sheetx] = sheet
			}
		}
	}
	return b.sheetList
}

// SheetByIndex returns a sheet by its index.
func (b *Book) SheetByIndex(sheetx int) (*Sheet, error) {
	if sheetx < 0 || sheetx >= len(b.sheetList) {
		return nil, NewXLRDError("sheet index %d out of range", sheetx)
	}
	
	// If sheet is already loaded, return it
	if b.sheetList[sheetx] != nil {
		return b.sheetList[sheetx], nil
	}
	
	// Load the sheet
	sheet, err := b.getSheet(sheetx)
	if err != nil {
		return nil, err
	}
	
	b.sheetList[sheetx] = sheet
	return sheet, nil
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
	// Empty implementation for now
	return false, nil
}

// UnloadSheet unloads a sheet by name or index.
func (b *Book) UnloadSheet(sheetNameOrIndex interface{}) error {
	// Empty implementation for now
	return nil
}

// ReleaseResources releases memory-consuming objects and possibly a memory-mapped file.
func (b *Book) ReleaseResources() {
	// Empty implementation for now
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
	if options == nil {
		options = &OpenWorkbookOptions{
			Logfile: os.Stdout,
		}
	}

	fileFormat, err := InspectFormat(filename, nil)
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
		sheetList:  []*Sheet{},
		sheetNames: []string{},
	}
	
	if options == nil {
		options = &OpenWorkbookOptions{
			Logfile: os.Stdout,
		}
	}
	
	bk.logfile = options.Logfile
	bk.verbosity = options.Verbosity
	bk.onDemand = options.OnDemand
	bk.formattingInfo = options.FormattingInfo
	bk.raggedRows = options.RaggedRows
	bk.encodingOverride = options.EncodingOverride
	
	// Read file
	fileContents, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
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
		for _, qname := range []string{"Workbook", "Book"} {
			mem, base, streamLen, err = cd.LocateNamedStream(qname)
			if err == nil && mem != nil {
				break
			}
		}
		
		if mem == nil {
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
		case XL_EXTERNSHEET:
			// TODO: handle externsheet - for now just skip
		case XL_SUPBOOK:
			// TODO: handle supbook - for now just skip
		case XL_SST:
			err := b.handleSST(data)
			if err != nil {
				return err
			}
		}

		if code == XL_EOF {
			break
		}
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
	// format.FormatKey = int(binary.LittleEndian.Uint16(data[pos : pos+2]))
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
func (b *Book) getSheet(shNumber int) (*Sheet, error) {
	if shNumber < 0 || shNumber >= len(b.sheetNames) {
		return nil, NewXLRDError("sheet index %d out of range", shNumber)
	}
	
	// Set position to sheet's absolute position
	if shNumber >= len(b.sheetAbsPosn) {
		return nil, NewXLRDError("sheet position not found for sheet %d", shNumber)
	}
	
	b.position = b.sheetAbsPosn[shNumber]

	// Get BOF record for worksheet
	_, err := b.getBOF(XL_WORKSHEET)
	if err != nil {
		return nil, err
	}
	
	// Create sheet
	sheet := &Sheet{
		Book:          b,
		Name:          b.sheetNames[shNumber],
		ColInfoMap:    make(map[int]*ColInfo),
		RowInfoMap:    make(map[int]*RowInfo),
		ColLabelRanges: make([][4]int, 0),
		RowLabelRanges: make([][4]int, 0),
		MergedCells:   make([][4]int, 0),
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
