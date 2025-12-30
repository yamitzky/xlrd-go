package xlrd

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"math"
	"strings"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
)

// Sheet contains the data for one worksheet.
//
// In the cell access functions, rowx is a row index, counting from zero,
// and colx is a column index, counting from zero.
// Negative values for row/column indexes and slice positions are supported.
//
// You don't instantiate this type yourself. You access Sheet objects via
// the Book object that was returned when you called OpenWorkbook.
type Sheet struct {
	BaseObject

	// Name is the name of the sheet.
	Name string

	// Book is a reference to the Book object to which this sheet belongs.
	Book *Book

	// NRows is the number of rows in sheet. A row index is in range(thesheet.NRows).
	NRows int

	// NCols is the nominal number of columns in sheet.
	// It is one more than the maximum column index found, ignoring trailing empty cells.
	NCols int

	// ColInfoMap is the map from a column index to a Colinfo object.
	ColInfoMap map[int]*ColInfo

	// RowInfoMap is the map from a row index to a Rowinfo object.
	RowInfoMap map[int]*RowInfo

	// ColLabelRanges is a list of address ranges of cells containing column labels.
	ColLabelRanges [][4]int

	// RowLabelRanges is a list of address ranges of cells containing row labels.
	RowLabelRanges [][4]int

	// MergedCells is a list of address ranges of cells which have been merged.
	MergedCells [][4]int

	// HyperlinkList contains HLINK records in this sheet.
	HyperlinkList []*Hyperlink

	// HyperlinkMap maps cell coordinates to hyperlinks.
	HyperlinkMap map[[2]int]*Hyperlink

	// CellNoteMap maps cell coordinates to notes/comments.
	CellNoteMap map[[2]int]*Note

	// RichTextRunlistMap maps cell coordinates to rich text run lists.
	RichTextRunlistMap map[[2]int][][]int

	// UtterMaxRows is the maximum row count supported by the BIFF version.
	UtterMaxRows int

	// cellAttrToXF maps BIFF2 cell attributes to XF indexes.
	cellAttrToXF map[[3]byte]int

	// ixfe holds the last IXFE record value for BIFF2.
	ixfe    int
	hasIXFE bool

	// Cell data storage (similar to Python's _cell_values, _cell_types, _cell_xf_indexes)
	cellValues    [][]interface{}
	cellTypes     [][]int
	cellXFIndexes [][]int

	// Sheet formatting and view info
	DefColWidth                     int
	StandardWidth                   int
	GCW                             []int
	FirstVisibleRowx                int
	FirstVisibleColx                int
	GridlineColourIndex             int
	GridlineColourRGB               [3]byte
	SclMagFactor                    int
	VertSplitPos                    int
	HorzSplitPos                    int
	HorzSplitFirstVisible           int
	VertSplitFirstVisible           int
	SplitActivePane                 int
	HasPaneRecord                   bool
	DefaultRowHeight                int
	DefaultRowHeightMismatch        int
	DefaultRowHidden                int
	DefaultAdditionalSpaceAbove     int
	DefaultAdditionalSpaceBelow     int
	HorizontalPageBreaks            [][3]int
	VerticalPageBreaks              [][3]int
	CachedPageBreakPreviewMagFactor int // default 0 (60%), from WINDOW2 record
	CachedNormalViewMagFactor       int // default 0 (100%), from WINDOW2 record
}

// Cell represents a cell in a worksheet.
type Cell struct {
	BaseObject

	// CType is the type of the cell.
	// One of: XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_DATE, XL_CELL_BOOLEAN, XL_CELL_ERROR, XL_CELL_BLANK
	CType int

	// Value is the value of the cell.
	Value interface{}

	// XFIndex is the index of the XF record for this cell.
	XFIndex int
}

// ColInfo contains information about a column.
type ColInfo struct {
	BaseObject

	// Width is the column width.
	Width int

	// Hidden indicates if the column is hidden.
	Hidden bool

	// XFIndex is the index of the XF record for this column.
	XFIndex int

	// OutlineLevel is the outline level for the column.
	OutlineLevel int

	// Collapsed indicates if the column is collapsed.
	Collapsed bool

	// Bit1Flag is an undocumented flag present in BIFF.
	Bit1Flag bool
}

// RowInfo contains information about a row.
type RowInfo struct {
	BaseObject

	// Height is the row height.
	Height int

	// Hidden indicates if the row is hidden.
	Hidden bool

	// XFIndex is the index of the XF record for this row.
	XFIndex int

	// HasDefaultHeight indicates if the row uses default height.
	HasDefaultHeight int

	// OutlineLevel is the outline level for the row.
	OutlineLevel int

	// OutlineGroupStartsEnds indicates if the row starts/ends an outline group.
	OutlineGroupStartsEnds int

	// HeightMismatch indicates if the row height differs from default.
	HeightMismatch int

	// HasDefaultXFIndex indicates if XFIndex is the default.
	HasDefaultXFIndex int

	// AdditionalSpaceAbove indicates extra space above the row.
	AdditionalSpaceAbove int

	// AdditionalSpaceBelow indicates extra space below the row.
	AdditionalSpaceBelow int
}

// Hyperlink contains the attributes of a hyperlink.
type Hyperlink struct {
	BaseObject

	FRowx int
	LRowx int
	FColx int
	LColx int

	Type      string
	URLOrPath interface{}
	Desc      string
	Target    string
	Textmark  string
	QuickTip  string
}

// Note represents a user "comment" or "note".
type Note struct {
	BaseObject

	Author          string
	ColHidden       int
	Colx            int
	RichTextRunlist [][2]int
	RowHidden       int
	Rowx            int
	Show            int
	Text            string
	ObjectID        int
}

// MSODrawing represents a drawing container from MSO records.
type MSODrawing struct {
	BaseObject

	AnchorUnk   int
	AnchorColLo int
	AnchorRowLo int
	AnchorColHi int
	AnchorRowHi int
}

// MSObj represents an OBJ record.
type MSObj struct {
	BaseObject

	Type          int
	ID            int
	Locked        int
	Printable     int
	AutoFilter    int
	ScrollbarFlag int
	Autofill      int
	Autoline      int

	ScrollbarValue int
	ScrollbarMin   int
	ScrollbarMax   int
	ScrollbarInc   int
	ScrollbarPage  int
}

// MSTxo represents a TXO record.
type MSTxo struct {
	BaseObject

	Rot             int
	ControlInfo     []byte
	IfntEmpty       int
	Fmla            []byte
	Text            string
	RichTextRunlist [][2]int

	HorzAlign  int
	VertAlign  int
	LockText   int
	JustLast   int
	SecretEdit int
}

// CellValue returns the value of the cell at the given row and column.
func (s *Sheet) CellValue(rowx, colx int) interface{} {
	if rowx < 0 || rowx >= s.NRows || colx < 0 || colx >= s.NCols {
		return nil
	}

	// Check if this cell is within a merged cell range
	for _, mergedRange := range s.MergedCells {
		rlo, rhi, clo, chi := mergedRange[0], mergedRange[1], mergedRange[2], mergedRange[3]
		if rlo <= rowx && rowx < rhi && clo <= colx && colx < chi {
			// This cell is within a merged range, return the value from the top-left cell
			if rlo >= len(s.cellValues) || s.cellValues[rlo] == nil || clo >= len(s.cellValues[rlo]) {
				return ""
			}
			value := s.cellValues[rlo][clo]
			if value == nil {
				return ""
			}
			return value
		}
	}

	if rowx >= len(s.cellValues) || s.cellValues[rowx] == nil || colx >= len(s.cellValues[rowx]) {
		return ""
	}
	value := s.cellValues[rowx][colx]
	if value == nil {
		return ""
	}
	return value
}

// RawCellValue returns the cell value without merged-cell expansion.
func (s *Sheet) RawCellValue(rowx, colx int) interface{} {
	if rowx < 0 || rowx >= s.NRows || colx < 0 || colx >= s.NCols {
		return nil
	}
	if rowx >= len(s.cellValues) || s.cellValues[rowx] == nil || colx >= len(s.cellValues[rowx]) {
		return ""
	}
	value := s.cellValues[rowx][colx]
	if value == nil {
		return ""
	}
	return value
}

// CellType returns the type of the cell at the given row and column.
func (s *Sheet) CellType(rowx, colx int) int {
	if rowx < 0 || rowx >= s.NRows || colx < 0 || colx >= s.NCols {
		return XL_CELL_EMPTY
	}

	// Check if this cell is within a merged cell range
	for _, mergedRange := range s.MergedCells {
		rlo, rhi, clo, chi := mergedRange[0], mergedRange[1], mergedRange[2], mergedRange[3]
		if rlo <= rowx && rowx < rhi && clo <= colx && colx < chi {
			// This cell is within a merged range, return the type from the top-left cell
			if rlo >= len(s.cellTypes) || s.cellTypes[rlo] == nil || clo >= len(s.cellTypes[rlo]) {
				return XL_CELL_EMPTY
			}
			return s.cellTypes[rlo][clo]
		}
	}

	if rowx >= len(s.cellTypes) || s.cellTypes[rowx] == nil || colx >= len(s.cellTypes[rowx]) {
		return XL_CELL_EMPTY
	}
	return s.cellTypes[rowx][colx]
}

// RawCellType returns the cell type without merged-cell expansion.
func (s *Sheet) RawCellType(rowx, colx int) int {
	if rowx < 0 || rowx >= s.NRows || colx < 0 || colx >= s.NCols {
		return XL_CELL_EMPTY
	}
	if rowx >= len(s.cellTypes) || s.cellTypes[rowx] == nil || colx >= len(s.cellTypes[rowx]) {
		return XL_CELL_EMPTY
	}
	return s.cellTypes[rowx][colx]
}

// Cell returns the Cell object at the given row and column.
func (s *Sheet) Cell(rowx, colx int) *Cell {
	ctype := s.CellType(rowx, colx)
	value := s.CellValue(rowx, colx)
	xfIndex := s.CellXFIndex(rowx, colx)
	return &Cell{
		CType:   ctype,
		Value:   value,
		XFIndex: xfIndex,
	}
}

// Row returns a slice of Cell objects for the given row.
func (s *Sheet) Row(rowx int) []*Cell {
	if rowx < 0 || rowx >= s.NRows {
		return []*Cell{}
	}
	cells := make([]*Cell, s.NCols)
	for colx := 0; colx < s.NCols; colx++ {
		cells[colx] = s.Cell(rowx, colx)
	}
	return cells
}

// RowLen returns the length of the row (number of non-empty cells).
func (s *Sheet) RowLen(rowx int) int {
	if rowx < 0 || rowx >= len(s.cellTypes) {
		return 0
	}
	// For ragged rows, count non-empty cells from the end
	if len(s.cellTypes[rowx]) == 0 {
		return 0
	}
	for colx := len(s.cellTypes[rowx]) - 1; colx >= 0; colx-- {
		if s.cellTypes[rowx][colx] != XL_CELL_EMPTY {
			return colx + 1
		}
	}
	return 0
}

// RowTypes returns a slice of the types of the cells in the given row.
func (s *Sheet) RowTypes(rowx, startColx int, endColx *int) []int {
	if rowx < 0 || rowx >= s.NRows {
		return []int{}
	}
	end := s.NCols
	if endColx != nil {
		end = *endColx
	}
	if startColx < 0 {
		startColx = 0
	}
	if end > s.NCols {
		end = s.NCols
	}
	if startColx >= end {
		return []int{}
	}
	types := make([]int, end-startColx)
	for i, colx := 0, startColx; colx < end; i, colx = i+1, colx+1 {
		types[i] = s.CellType(rowx, colx)
	}
	return types
}

// RowValues returns a slice of the values of the cells in the given row.
func (s *Sheet) RowValues(rowx, startColx int, endColx *int) []interface{} {
	if rowx < 0 || rowx >= s.NRows {
		return []interface{}{}
	}
	end := s.NCols
	if endColx != nil {
		end = *endColx
	}
	if startColx < 0 {
		startColx = 0
	}
	if end > s.NCols {
		end = s.NCols
	}
	if startColx >= end {
		return []interface{}{}
	}
	values := make([]interface{}, end-startColx)
	for i, colx := 0, startColx; colx < end; i, colx = i+1, colx+1 {
		values[i] = s.CellValue(rowx, colx)
	}
	return values
}

// RowSlice returns a slice of the Cell objects in the given row.
func (s *Sheet) RowSlice(rowx, startColx int, endColx *int) []*Cell {
	if rowx < 0 || rowx >= s.NRows {
		return []*Cell{}
	}
	end := s.NCols
	if endColx != nil {
		end = *endColx
	}
	if startColx < 0 {
		startColx = 0
	}
	if end > s.NCols {
		end = s.NCols
	}
	if startColx >= end {
		return []*Cell{}
	}
	cells := make([]*Cell, end-startColx)
	for i, colx := 0, startColx; colx < end; i, colx = i+1, colx+1 {
		cells[i] = s.Cell(rowx, colx)
	}
	return cells
}

// ColSlice returns a slice of the Cell objects in the given column.
func (s *Sheet) ColSlice(colx, startRowx int, endRowx *int) []*Cell {
	if colx < 0 || colx >= s.NCols {
		return []*Cell{}
	}
	end := s.NRows
	if endRowx != nil {
		end = *endRowx
	}
	if startRowx < 0 {
		startRowx = 0
	}
	if end > s.NRows {
		end = s.NRows
	}
	if startRowx >= end {
		return []*Cell{}
	}
	cells := make([]*Cell, end-startRowx)
	for i, rowx := 0, startRowx; rowx < end; i, rowx = i+1, rowx+1 {
		cells[i] = s.Cell(rowx, colx)
	}
	return cells
}

// Col is an alias for ColSlice with default parameters (startRowx=0, endRowx=nil).
func (s *Sheet) Col(colx int) []*Cell {
	return s.ColSlice(colx, 0, nil)
}

// ColValues returns a slice of the values of the cells in the given column.
func (s *Sheet) ColValues(colx, startRowx int, endRowx *int) []interface{} {
	if colx < 0 || colx >= s.NCols {
		return []interface{}{}
	}
	end := s.NRows
	if endRowx != nil {
		end = *endRowx
	}
	if startRowx < 0 {
		startRowx = 0
	}
	if end > s.NRows {
		end = s.NRows
	}
	if startRowx >= end {
		return []interface{}{}
	}
	values := make([]interface{}, end-startRowx)
	for i, rowx := 0, startRowx; rowx < end; i, rowx = i+1, rowx+1 {
		values[i] = s.CellValue(rowx, colx)
	}
	return values
}

// ColTypes returns a slice of the types of the cells in the given column.
func (s *Sheet) ColTypes(colx, startRowx int, endRowx *int) []int {
	if colx < 0 || colx >= s.NCols {
		return []int{}
	}
	end := s.NRows
	if endRowx != nil {
		end = *endRowx
	}
	if startRowx < 0 {
		startRowx = 0
	}
	if end > s.NRows {
		end = s.NRows
	}
	if startRowx >= end {
		return []int{}
	}
	types := make([]int, end-startRowx)
	for i, rowx := 0, startRowx; rowx < end; i, rowx = i+1, rowx+1 {
		types[i] = s.CellType(rowx, colx)
	}
	return types
}

// CellXFIndex returns the XF index of the cell at the given row and column.
func (s *Sheet) CellXFIndex(rowx, colx int) int {
	if rowx < 0 || rowx >= s.NRows || colx < 0 || colx >= s.NCols {
		return 0
	}

	// Check if this cell is within a merged cell range
	for _, mergedRange := range s.MergedCells {
		rlo, rhi, clo, chi := mergedRange[0], mergedRange[1], mergedRange[2], mergedRange[3]
		if rlo <= rowx && rowx < rhi && clo <= colx && colx < chi {
			// This cell is within a merged range, return the XF index from the top-left cell
			if rlo >= len(s.cellXFIndexes) || s.cellXFIndexes[rlo] == nil || clo >= len(s.cellXFIndexes[rlo]) {
				return 15 // Default XF index for empty cells
			}
			xfIndex := s.cellXFIndexes[rlo][clo]
			if xfIndex == 0 {
				return 15 // Default XF index for empty cells
			}
			return xfIndex
		}
	}

	if rowx >= len(s.cellXFIndexes) || s.cellXFIndexes[rowx] == nil || colx >= len(s.cellXFIndexes[rowx]) || s.cellXFIndexes[rowx][colx] == 0 {
		return 15 // Default XF index for empty cells
	}
	return s.cellXFIndexes[rowx][colx]
}

// RawCellXFIndex returns the XF index without merged-cell expansion.
func (s *Sheet) RawCellXFIndex(rowx, colx int) int {
	if rowx < 0 || rowx >= s.NRows || colx < 0 || colx >= s.NCols {
		return 0
	}
	if rowx >= len(s.cellXFIndexes) || s.cellXFIndexes[rowx] == nil || colx >= len(s.cellXFIndexes[rowx]) || s.cellXFIndexes[rowx][colx] == 0 {
		return 15
	}
	return s.cellXFIndexes[rowx][colx]
}

// EmptyCell returns an empty cell.
func EmptyCell() *Cell {
	return &Cell{CType: XL_CELL_EMPTY}
}

// putCell stores cell data at the specified row and column.
func (s *Sheet) putCell(rowx, colx int, ctype int, value interface{}, xfIndex int) {
	// Extend cell arrays if necessary
	for len(s.cellValues) <= rowx {
		s.cellValues = append(s.cellValues, nil)
		s.cellTypes = append(s.cellTypes, nil)
		s.cellXFIndexes = append(s.cellXFIndexes, nil)
	}
	for len(s.cellValues[rowx]) <= colx {
		s.cellValues[rowx] = append(s.cellValues[rowx], nil)
		s.cellTypes[rowx] = append(s.cellTypes[rowx], XL_CELL_EMPTY)
		s.cellXFIndexes[rowx] = append(s.cellXFIndexes[rowx], 0)
	}

	// Set cell data
	s.cellTypes[rowx][colx] = ctype
	s.cellValues[rowx][colx] = value
	s.cellXFIndexes[rowx][colx] = xfIndex

	// Update NRows and NCols
	if rowx >= s.NRows {
		s.NRows = rowx + 1
	}
	if colx >= s.NCols {
		s.NCols = colx + 1
	}
}

// read reads and parses the sheet data from the workbook.
func (s *Sheet) read(bk *Book) error {
	oldpos := bk.position
	defer func() {
		bk.position = oldpos
	}()

	if s.UtterMaxRows == 0 {
		if bk.BiffVersion >= 80 {
			s.UtterMaxRows = 65536
		} else {
			s.UtterMaxRows = 16384
		}
	}

	// Find the sheet index to get stream length
	sheetIndex := -1
	for i, name := range bk.sheetNames {
		if name == s.Name {
			sheetIndex = i
			break
		}
	}
	var maxPosition int
	if sheetIndex >= 0 && sheetIndex < len(bk.sheetAbsPosn)-1 {
		maxPosition = bk.sheetAbsPosn[sheetIndex+1]
	} else {
		maxPosition = len(bk.mem)
	}

	// Initialize cell arrays
	s.cellValues = make([][]interface{}, 0)
	s.cellTypes = make([][]int, 0)
	s.cellXFIndexes = make([][]int, 0)
	dimRows := 0
	dimCols := 0
	fmtInfo := bk.formattingInfo
	doSSTRichText := fmtInfo && bk.richTextRunlistMap != nil
	rowinfoSharing := make(map[[2]int]*RowInfo)
	rowinfoSharingB2 := make(map[[3]int]*RowInfo)
	txos := make(map[int]*MSTxo)
	savedObjID := 0
	eofFound := false

	// Parse BIFF records until EOF or end of sheet stream
	for {
		if bk.position >= maxPosition {
			break
		}
		rc, dataLen, data := bk.getRecordParts()
		if rc == XL_EOF {
			eofFound = true
			break
		}

		switch rc {
		case XL_NUMBER:
			if dataLen >= 14 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				bits := binary.LittleEndian.Uint64(data[6:14])
				value := math.Float64frombits(bits)
				if bk.verbosity >= 2 && s.Name == "PROFILEDEF" {
					fmt.Fprintf(bk.logfile, "DEBUG: %s XL_NUMBER at (%d,%d): value=%f, xf=%d\n", s.Name, rowx, colx, value, xfIndex)
				}
				s.putCell(rowx, colx, XL_CELL_NUMBER, value, xfIndex)
			}
		case XL_NUMBER_B2:
			if dataLen >= 15 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				cellAttr := data[4:7]
				xfIndex, err := s.fixedBIFF2XFIndex(cellAttr, rowx, colx, nil)
				if err != nil {
					return err
				}
				bits := binary.LittleEndian.Uint64(data[7:15])
				value := math.Float64frombits(bits)
				s.putCell(rowx, colx, XL_CELL_NUMBER, value, xfIndex)
			}
		case XL_LABELSST:
			if dataLen >= 10 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				sstIndex := int(binary.LittleEndian.Uint32(data[6:10]))
				if sstIndex < len(bk.sharedStrings) {
					value := bk.sharedStrings[sstIndex]
					s.putCell(rowx, colx, XL_CELL_TEXT, value, xfIndex)
					if doSSTRichText {
						if runlist, ok := bk.richTextRunlistMap[sstIndex]; ok {
							s.RichTextRunlistMap[[2]int{rowx, colx}] = runlist
						}
					}
				}
			}
		case XL_RK:
			if dataLen >= 10 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				rkData := data[6:10]
				rkValue := unpackRK(rkData)
				if bk.verbosity >= 2 && (s.Name == "PROFILEDEF" || rkValue == 100.0) {
					fmt.Fprintf(bk.logfile, "DEBUG: %s XL_RK at (%d,%d): rkData=%x, value=%f, xf=%d\n", s.Name, rowx, colx, rkData, rkValue, xfIndex)
				}
				s.putCell(rowx, colx, XL_CELL_NUMBER, rkValue, xfIndex)
			}
		case XL_MULRK:
			if dataLen >= 6 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				firstColx := int(binary.LittleEndian.Uint16(data[2:4]))
				lastColx := int(binary.LittleEndian.Uint16(data[dataLen-2 : dataLen]))

				pos := 4
				for colx := firstColx; colx <= lastColx && pos+6 <= dataLen-2; colx++ {
					xfIndex := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					rkData := data[pos+2 : pos+6]
					rkValue := unpackRK(rkData)
					s.putCell(rowx, colx, XL_CELL_NUMBER, rkValue, xfIndex)
					pos += 6
				}
			}
		case XL_ROW:
			if !fmtInfo {
				continue
			}
			if dataLen >= 16 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				if rowx < 0 || rowx >= s.UtterMaxRows {
					if bk.verbosity > 0 {
						fmt.Fprintf(bk.logfile,
							"*** NOTE: ROW record has row index %d; should have 0 <= rowx < %d -- record ignored!\n",
							rowx, s.UtterMaxRows)
					}
					continue
				}
				bits1 := int(binary.LittleEndian.Uint16(data[6:8]))
				bits2 := int(binary.LittleEndian.Uint32(data[8:12]))
				key := [2]int{bits1, bits2}
				r := rowinfoSharing[key]
				if r == nil {
					r = &RowInfo{}
					r.Height = bits1 & 0x7fff
					r.HasDefaultHeight = (bits1 >> 15) & 1
					r.OutlineLevel = bits2 & 7
					r.OutlineGroupStartsEnds = (bits2 >> 4) & 1
					r.Hidden = ((bits2 >> 5) & 1) != 0
					r.HeightMismatch = (bits2 >> 6) & 1
					r.HasDefaultXFIndex = (bits2 >> 7) & 1
					r.XFIndex = (bits2 >> 16) & 0xfff
					r.AdditionalSpaceAbove = (bits2 >> 28) & 1
					r.AdditionalSpaceBelow = (bits2 >> 29) & 1
					if r.HasDefaultXFIndex == 0 {
						r.XFIndex = -1
					}
					rowinfoSharing[key] = r
				}
				s.RowInfoMap[rowx] = r
			}
		case XL_DIMENSION, XL_DIMENSION2:
			if dataLen == 0 {
				break
			}
			if bk.BiffVersion < 80 {
				if dataLen >= 8 {
					dimRows = int(binary.LittleEndian.Uint16(data[2:4]))
					dimCols = int(binary.LittleEndian.Uint16(data[6:8]))
				}
			} else {
				if dataLen >= 12 {
					dimRows = int(binary.LittleEndian.Uint32(data[4:8]))
					dimCols = int(binary.LittleEndian.Uint16(data[10:12]))
				}
			}
			if bk.BiffVersion < 80 && bk.BiffVersion >= 20 && bk.XFList != nil && !bk.xfEpilogueDone {
				bk.xfEpilogue()
			}
		case XL_BOOLERR:
			if dataLen >= 8 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				value := int(data[6])
				isErr := data[7]
				if isErr != 0 {
					s.putCell(rowx, colx, XL_CELL_ERROR, value, xfIndex)
				} else {
					s.putCell(rowx, colx, XL_CELL_BOOLEAN, value, xfIndex)
				}
			}
		case XL_COLINFO:
			if !fmtInfo || dataLen < 10 {
				continue
			}
			firstColx := int(binary.LittleEndian.Uint16(data[0:2]))
			lastColx := int(binary.LittleEndian.Uint16(data[2:4]))
			width := int(binary.LittleEndian.Uint16(data[4:6]))
			xfIndex := int(binary.LittleEndian.Uint16(data[6:8]))
			flags := binary.LittleEndian.Uint16(data[8:10])
			if !(0 <= firstColx && firstColx <= lastColx && lastColx <= 256) {
				if bk.verbosity > 0 {
					fmt.Fprintf(bk.logfile,
						"*** NOTE: COLINFO record has first col index %d, last %d; should have 0 <= first <= last <= 255 -- record ignored!\n",
						firstColx, lastColx)
				}
				continue
			}
			c := &ColInfo{
				Width:   width,
				XFIndex: xfIndex,
			}
			c.Hidden = (flags & 0x0001) != 0
			c.Bit1Flag = (flags & 0x0002) != 0
			c.OutlineLevel = int((flags >> 8) & 0x0007)
			c.Collapsed = (flags & 0x1000) != 0
			for colx := firstColx; colx <= lastColx && colx <= 255; colx++ {
				s.ColInfoMap[colx] = c
			}
		case XL_DEFCOLWIDTH:
			if dataLen >= 2 {
				s.DefColWidth = int(binary.LittleEndian.Uint16(data[0:2]))
			}
		case XL_STANDARDWIDTH:
			if dataLen >= 2 {
				s.StandardWidth = int(binary.LittleEndian.Uint16(data[0:2]))
			} else if bk.verbosity > 0 {
				fmt.Fprintf(bk.logfile, "*** ERROR *** STANDARDWIDTH %d %x\n", dataLen, data)
			}
		case XL_GCW:
			if !fmtInfo || dataLen < 34 {
				continue
			}
			if data[0] != 0x20 || data[1] != 0x00 {
				continue
			}
			gcw := make([]int, 0, 256)
			for i := 0; i < 8; i++ {
				bits := binary.LittleEndian.Uint32(data[2+i*4 : 6+i*4])
				for j := 0; j < 32; j++ {
					gcw = append(gcw, int(bits&1))
					bits >>= 1
				}
			}
			s.GCW = gcw
		case XL_BLANK:
			if !fmtInfo {
				continue
			}
			if dataLen >= 6 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				s.putCell(rowx, colx, XL_CELL_BLANK, "", xfIndex)
			}
		case XL_MULBLANK:
			if !fmtInfo {
				continue
			}
			if dataLen >= 4 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				firstColx := int(binary.LittleEndian.Uint16(data[2:4]))
				pos := 4
				lastColx := int(binary.LittleEndian.Uint16(data[dataLen-2:]))
				for colx := firstColx; colx <= lastColx && pos+2 <= dataLen-2; colx++ {
					xfIndex := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					s.putCell(rowx, colx, XL_CELL_BLANK, "", xfIndex)
					pos += 2
				}
			}
		case XL_LABEL:
			if dataLen >= 6 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				// Parse string using BIFF record format
				if dataLen > 6 {
					enc := bk.Encoding
					if enc == "" {
						enc = bk.deriveEncoding()
					}
					var value string
					var err error
					if bk.BiffVersion < BIFF_FIRST_UNICODE {
						value, err = UnpackString(data, 6, enc, 2)
					} else {
						value, err = UnpackUnicode(data, 6, 2)
					}
					if err == nil {
						s.putCell(rowx, colx, XL_CELL_TEXT, value, xfIndex)
					}
				}
			}
		case XL_RSTRING:
			if dataLen >= 6 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				if bk.BiffVersion < BIFF_FIRST_UNICODE {
					strg, pos := unpack_string_update_pos(data, 6, bk.Encoding, 2, -1)
					if pos >= 0 && pos < len(data) {
						nrt := int(data[pos])
						pos++
						runlist := make([][]int, 0, nrt)
						for i := 0; i < nrt && pos+2 <= len(data); i++ {
							runlist = append(runlist, []int{int(data[pos]), int(data[pos+1])})
							pos += 2
						}
						s.putCell(rowx, colx, XL_CELL_TEXT, strg, xfIndex)
						s.RichTextRunlistMap[[2]int{rowx, colx}] = runlist
					}
				} else {
					strg, pos := unpack_unicode_update_pos(data, 6, 2, -1)
					if pos >= 0 && pos+2 <= len(data) {
						nrt := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
						pos += 2
						runlist := make([][]int, 0, nrt)
						for i := 0; i < nrt && pos+4 <= len(data); i++ {
							runlist = append(runlist, []int{int(binary.LittleEndian.Uint16(data[pos : pos+2])), int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))})
							pos += 4
						}
						s.putCell(rowx, colx, XL_CELL_TEXT, strg, xfIndex)
						s.RichTextRunlistMap[[2]int{rowx, colx}] = runlist
					}
				}
			}
		case XL_LABEL_B2:
			if dataLen >= 7 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				cellAttr := data[4:7]
				xfIndex, err := s.fixedBIFF2XFIndex(cellAttr, rowx, colx, nil)
				if err != nil {
					return err
				}
				enc := bk.Encoding
				if enc == "" {
					enc = bk.deriveEncoding()
				}
				value, err := UnpackString(data, 7, enc, 1)
				if err == nil {
					s.putCell(rowx, colx, XL_CELL_TEXT, value, xfIndex)
				}
			}
		case XL_INTEGER:
			if dataLen >= 9 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				cellAttr := data[4:7]
				xfIndex, err := s.fixedBIFF2XFIndex(cellAttr, rowx, colx, nil)
				if err != nil {
					return err
				}
				value := float64(binary.LittleEndian.Uint16(data[7:9]))
				s.putCell(rowx, colx, XL_CELL_NUMBER, value, xfIndex)
			}
		case XL_BOOLERR_B2:
			if dataLen >= 9 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				cellAttr := data[4:7]
				xfIndex, err := s.fixedBIFF2XFIndex(cellAttr, rowx, colx, nil)
				if err != nil {
					return err
				}
				value := data[7]
				isErr := data[8]
				if isErr != 0 {
					s.putCell(rowx, colx, XL_CELL_ERROR, value, xfIndex)
				} else {
					s.putCell(rowx, colx, XL_CELL_BOOLEAN, value, xfIndex)
				}
			}
		case XL_BLANK_B2:
			if dataLen >= 7 {
				if !fmtInfo {
					continue
				}
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				cellAttr := data[4:7]
				xfIndex, err := s.fixedBIFF2XFIndex(cellAttr, rowx, colx, nil)
				if err != nil {
					return err
				}
				s.putCell(rowx, colx, XL_CELL_BLANK, "", xfIndex)
			}
		case XL_ROW_B2:
			if !fmtInfo || dataLen < 11 {
				continue
			}
			rowx := int(binary.LittleEndian.Uint16(data[0:2]))
			if rowx < 0 || rowx >= s.UtterMaxRows {
				if bk.verbosity > 0 {
					fmt.Fprintf(bk.logfile,
						"*** NOTE: ROW_B2 record has row index %d; should have 0 <= rowx < %d -- record ignored!\n",
						rowx, s.UtterMaxRows)
				}
				continue
			}
			bits1 := int(binary.LittleEndian.Uint16(data[6:8]))
			bits2 := int(data[10])
			xfIndex := -1
			if bits2&1 != 0 {
				if dataLen == 18 {
					trueXfx := int(binary.LittleEndian.Uint16(data[16:18]))
					xf, err := s.fixedBIFF2XFIndex(nil, rowx, -1, &trueXfx)
					if err != nil {
						return err
					}
					xfIndex = xf
				} else if dataLen >= 16 {
					cellAttr := data[13:16]
					xf, err := s.fixedBIFF2XFIndex(cellAttr, rowx, -1, nil)
					if err != nil {
						return err
					}
					xfIndex = xf
				}
			}
			key := [3]int{bits1, bits2, xfIndex}
			r := rowinfoSharingB2[key]
			if r == nil {
				r = &RowInfo{
					Height:           bits1 & 0x7fff,
					HasDefaultHeight: (bits1 >> 15) & 1,
					XFIndex:          xfIndex,
					HasDefaultXFIndex: func() int {
						if bits2&1 != 0 {
							return 1
						}
						return 0
					}(),
				}
				rowinfoSharingB2[key] = r
			}
			s.RowInfoMap[rowx] = r
		case XL_COLWIDTH:
			if !fmtInfo || dataLen < 4 {
				continue
			}
			firstColx := int(data[0])
			lastColx := int(data[1])
			width := int(binary.LittleEndian.Uint16(data[2:4]))
			if firstColx > lastColx {
				if bk.verbosity > 0 {
					fmt.Fprintf(bk.logfile,
						"*** NOTE: COLWIDTH record has first col index %d, last %d; should have first <= last -- record ignored!\n",
						firstColx, lastColx)
				}
				continue
			}
			for colx := firstColx; colx <= lastColx; colx++ {
				c := s.ColInfoMap[colx]
				if c == nil {
					c = &ColInfo{}
					s.ColInfoMap[colx] = c
				}
				c.Width = width
			}
		case XL_COLUMNDEFAULT:
			if !fmtInfo || dataLen < 4 {
				continue
			}
			firstColx := int(binary.LittleEndian.Uint16(data[0:2]))
			lastColx := int(binary.LittleEndian.Uint16(data[2:4]))
			if !(0 <= firstColx && firstColx < lastColx && lastColx <= 256) {
				if bk.verbosity > 0 {
					fmt.Fprintf(bk.logfile,
						"*** NOTE: COLUMNDEFAULT record has first col index %d, last %d; should have 0 <= first < last <= 256\n",
						firstColx, lastColx)
				}
				if lastColx > 256 {
					lastColx = 256
				}
			}
			for colx := firstColx; colx < lastColx; colx++ {
				offset := 4 + 3*(colx-firstColx)
				if offset+3 > dataLen {
					break
				}
				cellAttr := data[offset : offset+3]
				xfIndex, err := s.fixedBIFF2XFIndex(cellAttr, -1, colx, nil)
				if err != nil {
					return err
				}
				c := s.ColInfoMap[colx]
				if c == nil {
					c = &ColInfo{}
					s.ColInfoMap[colx] = c
				}
				c.XFIndex = xfIndex
			}
		case XL_WINDOW2_B2:
			if dataLen >= 13 {
				s.FirstVisibleRowx = int(binary.LittleEndian.Uint16(data[5:7]))
				s.FirstVisibleColx = int(binary.LittleEndian.Uint16(data[7:9]))
				s.GridlineColourRGB = [3]byte{data[10], data[11], data[12]}
				s.GridlineColourIndex = NearestColourIndex(bk.ColourMap, [3]int{int(s.GridlineColourRGB[0]), int(s.GridlineColourRGB[1]), int(s.GridlineColourRGB[2])}, 0)
			}
		case XL_IXFE:
			if dataLen >= 2 {
				s.ixfe = int(binary.LittleEndian.Uint16(data[0:2]))
				s.hasIXFE = true
			}
		case XL_FORMULA, XL_FORMULA3, XL_FORMULA4:
			s.handleFormula(bk, data, dataLen)
		case XL_MERGEDCELLS:
			s.handleMergedCells(data, dataLen)
		case XL_FORMAT, XL_FORMAT2:
			if bk.BiffVersion <= 45 {
				if err := bk.handleFormat(data, rc); err != nil {
					return err
				}
			}
		case XL_FONT, XL_FONT_B3B4:
			if bk.BiffVersion <= 45 {
				if err := bk.handleFont(data); err != nil {
					return err
				}
			}
		case XL_STYLE:
			if bk.BiffVersion <= 45 {
				if !bk.xfEpilogueDone {
					bk.xfEpilogue()
				}
				if err := bk.handleStyle(data); err != nil {
					return err
				}
			}
		case XL_PALETTE:
			if bk.BiffVersion <= 45 {
				if err := bk.handlePalette(data); err != nil {
					return err
				}
			}
		case XL_BUILTINFMTCOUNT:
			if bk.BiffVersion <= 45 {
				bk.handleBuiltinfmtcount(data)
			}
		case XL_XF4, XL_XF3, XL_XF2:
			if bk.BiffVersion <= 45 {
				if err := bk.handleXF(data); err != nil {
					return err
				}
			}
		case XL_DATEMODE:
			if bk.BiffVersion <= 45 {
				bk.handleDatemode(data)
			}
		case XL_CODEPAGE:
			if bk.BiffVersion <= 45 {
				bk.handleCodepage(data)
			}
		case XL_FILEPASS:
			if bk.BiffVersion <= 45 {
				if err := bk.handleFilepass(data); err != nil {
					return err
				}
			}
		case XL_WRITEACCESS:
			if bk.BiffVersion <= 45 {
				bk.handleWriteAccess(data)
			}
		case XL_HLINK:
			if fmtInfo {
				s.handleHlink(data)
			}
		case XL_QUICKTIP:
			if fmtInfo {
				s.handleQuicktip(data)
			}
		case XL_OBJ:
			if fmtInfo {
				saved := s.handleObj(bk, data)
				if saved != nil {
					savedObjID = saved.ID
				} else {
					savedObjID = 0
				}
			}
		case XL_MSO_DRAWING:
			if fmtInfo {
				s.handleMSODrawingEtc(bk, rc, dataLen, data)
			}
		case XL_TXO:
			if fmtInfo {
				txo := s.handleTxo(bk, data)
				if txo != nil && savedObjID != 0 {
					txos[savedObjID] = txo
					savedObjID = 0
				}
			}
		case XL_NOTE:
			if fmtInfo {
				s.handleNote(bk, data, txos)
			}
		case XL_FEAT11:
			if fmtInfo {
				s.handleFeat11(bk, data)
			}
		case XL_COUNTRY:
			bk.handleCountry(data)
		case XL_LABELRANGES:
			pos := 0
			var ranges []CellRange
			pos = unpack_cell_range_address_list_update_pos(&ranges, data, pos, bk.BiffVersion, 8)
			for _, r := range ranges {
				s.RowLabelRanges = append(s.RowLabelRanges, [4]int{r.FirstRow, r.LastRow, r.FirstCol, r.LastCol})
			}
			ranges = nil
			pos = unpack_cell_range_address_list_update_pos(&ranges, data, pos, bk.BiffVersion, 8)
			for _, r := range ranges {
				s.ColLabelRanges = append(s.ColLabelRanges, [4]int{r.FirstRow, r.LastRow, r.FirstCol, r.LastCol})
			}
			_ = pos
		case XL_ARRAY:
			if bk.verbosity >= 2 {
				// Consume array header and ignore.
				if dataLen >= 14 {
					row1x := int(binary.LittleEndian.Uint16(data[0:2]))
					rownx := int(binary.LittleEndian.Uint16(data[2:4]))
					col1x := int(data[4])
					colnx := int(data[5])
					_ = row1x
					_ = rownx
					_ = col1x
					_ = colnx
				}
			}
		case XL_SHRFMLA:
			if bk.verbosity >= 2 && dataLen >= 10 {
				// Shared formula header only (no-op unless debug).
			}
		case XL_CONDFMT:
			if fmtInfo && bk.verbosity >= 1 {
				fmt.Fprintf(bk.logfile, "\n*** WARNING: Ignoring CONDFMT (conditional formatting) record in Sheet %q.\n", s.Name)
			}
		case XL_CF:
			if fmtInfo && bk.verbosity >= 1 {
				fmt.Fprintf(bk.logfile, "\n*** WARNING: Ignoring CF (conditional formatting) sub-record in Sheet %q.\n", s.Name)
			}
		case XL_DEFAULTROWHEIGHT:
			if dataLen == 4 {
				bits := int(binary.LittleEndian.Uint16(data[0:2]))
				s.DefaultRowHeight = int(binary.LittleEndian.Uint16(data[2:4]))
				s.DefaultRowHeightMismatch = bits & 1
				s.DefaultRowHidden = (bits >> 1) & 1
				s.DefaultAdditionalSpaceAbove = (bits >> 2) & 1
				s.DefaultAdditionalSpaceBelow = (bits >> 3) & 1
			} else if dataLen == 2 {
				s.DefaultRowHeight = int(binary.LittleEndian.Uint16(data[0:2]))
				if bk.verbosity > 0 {
					fmt.Fprintf(bk.logfile, "*** WARNING: DEFAULTROWHEIGHT record len is 2, should be 4; assuming BIFF2 format\n")
				}
			} else if bk.verbosity > 0 {
				fmt.Fprintf(bk.logfile, "*** WARNING: DEFAULTROWHEIGHT record len is %d, should be 4; ignoring this record\n", dataLen)
			}
		case XL_WINDOW2:
			if bk.BiffVersion >= 80 && dataLen >= 14 {
				options := binary.LittleEndian.Uint16(data[0:2])
				s.FirstVisibleRowx = int(binary.LittleEndian.Uint16(data[2:4]))
				s.FirstVisibleColx = int(binary.LittleEndian.Uint16(data[4:6]))
				s.GridlineColourIndex = int(binary.LittleEndian.Uint16(data[6:8]))
				s.CachedPageBreakPreviewMagFactor = int(binary.LittleEndian.Uint16(data[10:12]))
				s.CachedNormalViewMagFactor = int(binary.LittleEndian.Uint16(data[12:14]))
				_ = options
			} else if bk.BiffVersion >= 30 && dataLen >= 9 {
				options := binary.LittleEndian.Uint16(data[0:2])
				s.FirstVisibleRowx = int(binary.LittleEndian.Uint16(data[2:4]))
				s.FirstVisibleColx = int(binary.LittleEndian.Uint16(data[4:6]))
				s.GridlineColourRGB = [3]byte{data[6], data[7], data[8]}
				s.GridlineColourIndex = NearestColourIndex(bk.ColourMap, [3]int{int(s.GridlineColourRGB[0]), int(s.GridlineColourRGB[1]), int(s.GridlineColourRGB[2])}, 0)
				_ = options
			}
		case XL_SCL:
			if dataLen >= 4 {
				num := int(binary.LittleEndian.Uint16(data[0:2]))
				den := int(binary.LittleEndian.Uint16(data[2:4]))
				result := 0
				if den != 0 {
					result = (num * 100) / den
				}
				if result < 10 || result > 400 {
					if bk.verbosity >= 0 {
						fmt.Fprintf(bk.logfile,
							"WARNING *** SCL rcd sheet %q: should have 0.1 <= num/den <= 4; got %d/%d\n",
							s.Name, num, den)
					}
					result = 100
				}
				s.SclMagFactor = result
			}
		case XL_PANE:
			if dataLen >= 9 {
				s.VertSplitPos = int(binary.LittleEndian.Uint16(data[0:2]))
				s.HorzSplitPos = int(binary.LittleEndian.Uint16(data[2:4]))
				s.HorzSplitFirstVisible = int(binary.LittleEndian.Uint16(data[4:6]))
				s.VertSplitFirstVisible = int(binary.LittleEndian.Uint16(data[6:8]))
				s.SplitActivePane = int(data[8])
				s.HasPaneRecord = true
			}
		case XL_HORIZONTALPAGEBREAKS:
			if !fmtInfo || dataLen < 2 {
				continue
			}
			numBreaks := int(binary.LittleEndian.Uint16(data[0:2]))
			pos := 2
			if bk.BiffVersion < 80 {
				for pos+2 <= dataLen {
					br := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					s.HorizontalPageBreaks = append(s.HorizontalPageBreaks, [3]int{br, 0, 255})
					pos += 2
				}
			} else {
				for i := 0; i < numBreaks && pos+6 <= dataLen; i++ {
					br := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					fc := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
					lc := int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))
					s.HorizontalPageBreaks = append(s.HorizontalPageBreaks, [3]int{br, fc, lc})
					pos += 6
				}
			}
		case XL_VERTICALPAGEBREAKS:
			if !fmtInfo || dataLen < 2 {
				continue
			}
			numBreaks := int(binary.LittleEndian.Uint16(data[0:2]))
			pos := 2
			if bk.BiffVersion < 80 {
				for pos+2 <= dataLen {
					br := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					s.VerticalPageBreaks = append(s.VerticalPageBreaks, [3]int{br, 0, 65535})
					pos += 2
				}
			} else {
				for i := 0; i < numBreaks && pos+6 <= dataLen; i++ {
					br := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					fr := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
					lr := int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))
					s.VerticalPageBreaks = append(s.VerticalPageBreaks, [3]int{br, fr, lr})
					pos += 6
				}
			}
		case XL_EOF:
			// handled by loop condition
			break
		}
	}

	if dimRows > s.NRows {
		s.NRows = dimRows
	}
	if s.NCols == 0 && dimCols > 0 {
		s.NCols = dimCols
	}
	if !eofFound {
		return NewXLRDError("Sheet %q missing EOF record", s.Name)
	}

	return nil
}

// handleMergedCells processes XL_MERGEDCELLS records.
func (s *Sheet) handleMergedCells(data []byte, dataLen int) {
	if dataLen < 2 {
		return
	}

	numRanges := int(binary.LittleEndian.Uint16(data[0:2]))
	pos := 2

	for i := 0; i < numRanges && pos+8 <= dataLen; i++ {
		row1 := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		row2 := int(binary.LittleEndian.Uint16(data[pos+2 : pos+4]))
		col1 := int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))
		col2 := int(binary.LittleEndian.Uint16(data[pos+6 : pos+8]))

		// Excel stores merged ranges as (start_row, end_row, start_col, end_col) where end is inclusive
		// Python xlrd adds 1 to end_row and end_col to match Python slice conventions
		// We follow the same convention for compatibility
		s.MergedCells = append(s.MergedCells, [4]int{row1, row2 + 1, col1, col2 + 1})
		pos += 8
	}
}

func (s *Sheet) fixedBIFF2XFIndex(cellAttr []byte, rowx, colx int, trueXfx *int) (int, error) {
	bk := s.Book
	if bk.BiffVersion == 21 {
		if len(bk.XFList) > 0 {
			var xfx int
			if trueXfx != nil {
				xfx = *trueXfx
			} else if len(cellAttr) > 0 {
				xfx = int(cellAttr[0] & 0x3F)
			}
			if xfx == 0x3F {
				if !s.hasIXFE {
					return 0, NewXLRDError("BIFF2 cell record has XF index 63 but no preceding IXFE record")
				}
				xfx = s.ixfe
			}
			return xfx, nil
		}
		bk.BiffVersion = 20
	}

	if len(cellAttr) < 3 {
		return 0, NewXLRDError("BIFF2 cell_attr too short at (%d,%d)", rowx, colx)
	}

	xfxSlot := cellAttr[0] & 0x3F
	if xfxSlot != 0 && bk.verbosity > 0 {
		fmt.Fprintf(bk.logfile, "WARNING: BIFF2 cell_attr slot not zero at (%d,%d): %02x\n", rowx, colx, xfxSlot)
	}

	var key [3]byte
	copy(key[:], cellAttr[:3])
	if xfx, ok := s.cellAttrToXF[key]; ok {
		return xfx, nil
	}

	if len(bk.XFList) == 0 {
		for i := 0; i < 16; i++ {
			s.insertNewBIFF20XF([]byte{0x40, 0x00, 0x00}, i < 15)
		}
	}

	xfx := s.insertNewBIFF20XF(cellAttr, false)
	return xfx, nil
}

func (s *Sheet) insertNewBIFF20XF(cellAttr []byte, style bool) int {
	bk := s.Book
	xfx := len(bk.XFList)
	xf := s.fakeXFFromBIFF20CellAttr(cellAttr, style)
	xf.XFIndex = xfx
	bk.XFList = append(bk.XFList, xf)

	if bk.FormatMap == nil {
		bk.FormatMap = make(map[int]*Format)
	}
	if _, ok := bk.FormatMap[xf.FormatKey]; !ok {
		if xf.FormatKey != 0 && bk.verbosity > 0 {
			fmt.Fprintf(bk.logfile, "ERROR *** XF[%d] unknown format key (%d, 0x%04x)\n",
				xf.XFIndex, xf.FormatKey, xf.FormatKey)
		}
		format := &Format{FormatKey: xf.FormatKey, Type: FUN, FormatString: "General"}
		bk.FormatMap[xf.FormatKey] = format
		bk.FormatList = append(bk.FormatList, format)
	}

	cellType := XL_CELL_NUMBER
	if fmtObj, ok := bk.FormatMap[xf.FormatKey]; ok {
		if ty, ok := cellTypeFromFormatType[fmtObj.Type]; ok {
			cellType = ty
		}
	}
	if bk.xfIndexToXLTypeMap == nil {
		bk.xfIndexToXLTypeMap = make(map[int]int)
	}
	bk.xfIndexToXLTypeMap[xf.XFIndex] = cellType

	var key [3]byte
	copy(key[:], cellAttr[:3])
	s.cellAttrToXF[key] = xfx
	return xfx
}

func (s *Sheet) fakeXFFromBIFF20CellAttr(cellAttr []byte, style bool) *XF {
	xf := &XF{
		Alignment:  &XFAlignment{},
		Border:     &XFBorder{},
		Background: &XFBackground{},
		Protection: &XFProtection{},
	}
	if len(cellAttr) < 3 {
		return xf
	}
	protBits := cellAttr[0]
	fontAndFormat := cellAttr[1]
	halignEtc := cellAttr[2]

	xf.FormatKey = int(fontAndFormat & 0x3F)
	xf.FontIndex = int((fontAndFormat & 0xC0) >> 6)

	xf.Protection.CellLocked = (protBits & 0x40) != 0
	xf.Protection.FormulaHidden = (protBits & 0x80) != 0

	xf.Alignment.HorAlign = int(halignEtc & 0x07)
	xf.Alignment.Horizontal = xf.Alignment.HorAlign
	xf.Alignment.VertAlign = 2
	xf.Alignment.Vertical = 2
	xf.Alignment.Rotation = 0

	if halignEtc&0x08 != 0 {
		xf.Border.LeftColourIndex = 8
		xf.Border.LeftLineStyle = 1
	} else {
		xf.Border.LeftColourIndex = 0
		xf.Border.LeftLineStyle = 0
	}
	if halignEtc&0x10 != 0 {
		xf.Border.RightColourIndex = 8
		xf.Border.RightLineStyle = 1
	} else {
		xf.Border.RightColourIndex = 0
		xf.Border.RightLineStyle = 0
	}
	if halignEtc&0x20 != 0 {
		xf.Border.TopColourIndex = 8
		xf.Border.TopLineStyle = 1
	} else {
		xf.Border.TopColourIndex = 0
		xf.Border.TopLineStyle = 0
	}
	if halignEtc&0x40 != 0 {
		xf.Border.BottomColourIndex = 8
		xf.Border.BottomLineStyle = 1
	} else {
		xf.Border.BottomColourIndex = 0
		xf.Border.BottomLineStyle = 0
	}

	if halignEtc&0x80 != 0 {
		xf.Background.FillPattern = 17
	} else {
		xf.Background.FillPattern = 0
	}
	xf.Background.BackgroundColourIndex = 9
	xf.Background.PatternColourIndex = 8

	if style {
		xf.ParentStyleIndex = 0x0FFF
		xf.IsStyle = 1
	} else {
		xf.ParentStyleIndex = 0
	}

	xf.FormatFlag = 1
	xf.FontFlag = 1
	xf.AlignmentFlag = 1
	xf.BorderFlag = 1
	xf.BackgroundFlag = 1
	xf.ProtectionFlag = 1

	xf.Locked = xf.Protection.CellLocked
	xf.Hidden = xf.Protection.FormulaHidden
	xf.Border.Left = xf.Border.LeftLineStyle
	xf.Border.Right = xf.Border.RightLineStyle
	xf.Border.Top = xf.Border.TopLineStyle
	xf.Border.Bottom = xf.Border.BottomLineStyle
	xf.Alignment.WrapText = xf.Alignment.TextWrapped

	return xf
}

func decodeUTF16LE(data []byte) string {
	if len(data)%2 != 0 {
		data = data[:len(data)-1]
	}
	words := make([]uint16, len(data)/2)
	for i := 0; i < len(words); i++ {
		words[i] = binary.LittleEndian.Uint16(data[i*2 : (i+1)*2])
	}
	return string(utf16.Decode(words))
}

func getNulTerminatedUnicode(data []byte, offset int) (string, int) {
	if offset+4 > len(data) {
		return "", offset
	}
	nchars := int(binary.LittleEndian.Uint32(data[offset : offset+4]))
	offset += 4
	nb := nchars * 2
	if offset+nb > len(data) {
		return "", offset
	}
	uc := decodeUTF16LE(data[offset : offset+nb])
	if len(uc) > 0 {
		uc = uc[:len(uc)-1]
	}
	offset += nb
	return uc, offset
}

func (s *Sheet) handleHlink(data []byte) {
	if len(data) < 32 {
		return
	}
	h := &Hyperlink{}
	h.FRowx = int(binary.LittleEndian.Uint16(data[0:2]))
	h.LRowx = int(binary.LittleEndian.Uint16(data[2:4]))
	h.FColx = int(binary.LittleEndian.Uint16(data[4:6]))
	h.LColx = int(binary.LittleEndian.Uint16(data[6:8]))
	guid0 := data[8:24]
	dummy := data[24:28]
	options := int(binary.LittleEndian.Uint32(data[28:32]))

	if string(guid0) != string([]byte{0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B}) {
		return
	}
	if string(dummy) != string([]byte{0x02, 0x00, 0x00, 0x00}) {
		return
	}
	offset := 32

	if options&0x14 != 0 {
		var desc string
		desc, offset = getNulTerminatedUnicode(data, offset)
		h.Desc = desc
	}
	if options&0x80 != 0 {
		var target string
		target, offset = getNulTerminatedUnicode(data, offset)
		h.Target = target
	}

	if (options&1) != 0 && (options&0x100) == 0 {
		if offset+16 > len(data) {
			return
		}
		clsid := data[offset : offset+16]
		offset += 16
		if string(clsid) == string([]byte{0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B}) {
			h.Type = "url"
			if offset+4 > len(data) {
				return
			}
			nbytes := int(binary.LittleEndian.Uint32(data[offset : offset+4]))
			offset += 4
			if offset+nbytes > len(data) {
				return
			}
			raw := data[offset : offset+nbytes]
			ustr := decodeUTF16LE(raw)
			if idx := strings.IndexByte(ustr, 0); idx >= 0 {
				ustr = ustr[:idx]
			}
			h.URLOrPath = ustr
			endpos := len(ustr)
			trueNBytes := 2 * (endpos + 1)
			offset += trueNBytes
			extra := nbytes - trueNBytes
			if extra > 0 && offset+extra <= len(data) {
				offset += extra
			}
		} else if string(clsid) == string([]byte{0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}) {
			h.Type = "local file"
			if offset+6 > len(data) {
				return
			}
			uplevels := int(binary.LittleEndian.Uint16(data[offset : offset+2]))
			nbytes := int(binary.LittleEndian.Uint32(data[offset+2 : offset+6]))
			offset += 6
			if offset+nbytes > len(data) {
				return
			}
			shortpath := append(bytes.Repeat([]byte("..\\"), uplevels), data[offset:offset+nbytes-1]...)
			offset += nbytes
			if offset+24 > len(data) {
				return
			}
			offset += 24
			if offset+4 > len(data) {
				return
			}
			sz := int(binary.LittleEndian.Uint32(data[offset : offset+4]))
			offset += 4
			if sz != 0 {
				if offset+6 > len(data) {
					return
				}
				xl := int(binary.LittleEndian.Uint32(data[offset : offset+4]))
				offset += 4
				offset += 2
				if offset+xl > len(data) {
					return
				}
				extendedPath := decodeUTF16LE(data[offset : offset+xl])
				offset += xl
				h.URLOrPath = extendedPath
			} else {
				h.URLOrPath = shortpath
			}
		} else {
			h.Type = "unknown"
		}
	} else if options&0x163 == 0x103 {
		h.Type = "unc"
		unc, newOffset := getNulTerminatedUnicode(data, offset)
		offset = newOffset
		h.URLOrPath = unc
	} else if options&0x16B == 8 {
		h.Type = "workbook"
	} else {
		h.Type = "unknown"
	}

	if options&0x8 != 0 {
		textmark, newOffset := getNulTerminatedUnicode(data, offset)
		offset = newOffset
		h.Textmark = textmark
	}

	extra := len(data) - offset
	if extra > 0 && s.Book != nil && s.Book.verbosity > 0 {
		fmt.Fprintf(s.Book.logfile, "*** WARNING: hyperlink at R%dC%d has %d extra data bytes\n",
			h.FRowx+1, h.FColx+1, extra)
	}

	s.HyperlinkList = append(s.HyperlinkList, h)
	for rowx := h.FRowx; rowx <= h.LRowx; rowx++ {
		for colx := h.FColx; colx <= h.LColx; colx++ {
			s.HyperlinkMap[[2]int{rowx, colx}] = h
		}
	}
}

func (s *Sheet) handleQuicktip(data []byte) {
	if len(data) < 10 || len(s.HyperlinkList) == 0 {
		return
	}
	frowx := int(binary.LittleEndian.Uint16(data[2:4]))
	lrowx := int(binary.LittleEndian.Uint16(data[4:6]))
	fcolx := int(binary.LittleEndian.Uint16(data[6:8]))
	lcolx := int(binary.LittleEndian.Uint16(data[8:10]))
	h := s.HyperlinkList[len(s.HyperlinkList)-1]
	if frowx != h.FRowx || lrowx != h.LRowx || fcolx != h.FColx || lcolx != h.LColx {
		return
	}
	if len(data) >= 12 && data[len(data)-2] == 0x00 && data[len(data)-1] == 0x00 {
		h.QuickTip = decodeUTF16LE(data[10 : len(data)-2])
	}
}

func (s *Sheet) handleMSODrawingEtc(bk *Book, recid int, dataLen int, data []byte) {
	if recid != XL_MSO_DRAWING || bk.BiffVersion < 80 {
		return
	}
	if dataLen == 0 {
		return
	}
	pos := 0
	for pos+8 <= dataLen {
		tmp := binary.LittleEndian.Uint16(data[pos : pos+2])
		fbt := binary.LittleEndian.Uint16(data[pos+2 : pos+4])
		cb := binary.LittleEndian.Uint32(data[pos+4 : pos+8])
		ver := tmp & 0xF
		ndb := int(cb)
		if ver == 0xF {
			ndb = 0
		}
		if fbt == 0xF010 && pos+8+ndb <= dataLen && ndb >= 18 {
			_ = binary.LittleEndian.Uint16(data[pos+8 : pos+10])
		}
		pos += ndb + 8
	}
}

func (s *Sheet) handleObj(bk *Book, data []byte) *MSObj {
	if bk.BiffVersion < 80 || len(data) < 4 {
		return nil
	}
	o := &MSObj{}
	pos := 0
	for pos+4 <= len(data) {
		ft := binary.LittleEndian.Uint16(data[pos : pos+2])
		cb := binary.LittleEndian.Uint16(data[pos+2 : pos+4])
		if pos == 0 && !(ft == 0x15 && cb == 18) {
			if bk.verbosity > 0 {
				fmt.Fprintf(bk.logfile, "*** WARNING Ignoring antique or corrupt OBJECT record\n")
			}
			return nil
		}
		if ft == 0x15 && pos+10 <= len(data) {
			o.Type = int(binary.LittleEndian.Uint16(data[pos+4 : pos+6]))
			o.ID = int(binary.LittleEndian.Uint16(data[pos+6 : pos+8]))
			optionFlags := binary.LittleEndian.Uint16(data[pos+8 : pos+10])
			o.Locked = int(optionFlags & 0x0001)
			o.Printable = int((optionFlags >> 4) & 1)
			o.AutoFilter = int((optionFlags >> 8) & 1)
			o.ScrollbarFlag = int((optionFlags >> 9) & 1)
			o.Autofill = int((optionFlags >> 13) & 1)
			o.Autoline = int((optionFlags >> 14) & 1)
		} else if ft == 0x0C && pos+18 <= len(data) {
			o.ScrollbarValue = int(binary.LittleEndian.Uint16(data[pos+8 : pos+10]))
			o.ScrollbarMin = int(binary.LittleEndian.Uint16(data[pos+10 : pos+12]))
			o.ScrollbarMax = int(binary.LittleEndian.Uint16(data[pos+12 : pos+14]))
			o.ScrollbarInc = int(binary.LittleEndian.Uint16(data[pos+14 : pos+16]))
			o.ScrollbarPage = int(binary.LittleEndian.Uint16(data[pos+16 : pos+18]))
		} else if ft == 0x00 {
			break
		}
		pos += int(cb) + 4
	}
	return o
}

func (s *Sheet) handleNote(bk *Book, data []byte, txos map[int]*MSTxo) {
	o := &Note{}
	if bk.BiffVersion < 80 {
		if len(data) < 6 {
			return
		}
		o.Rowx = int(binary.LittleEndian.Uint16(data[0:2]))
		o.Colx = int(binary.LittleEndian.Uint16(data[2:4]))
		expectedBytes := int(binary.LittleEndian.Uint16(data[4:6]))
		nb := len(data) - 6
		pieces := [][]byte{data[6:]}
		expectedBytes -= nb
		for expectedBytes > 0 {
			rc2, data2Len, data2 := bk.getRecordParts()
			if rc2 != XL_NOTE || data2Len < 6 {
				return
			}
			nb = int(binary.LittleEndian.Uint16(data2[4:6]))
			pieces = append(pieces, data2[6:])
			expectedBytes -= nb
		}
		enc := bk.Encoding
		if enc == "" {
			enc = bk.deriveEncoding()
		}
		raw := bytes.Join(pieces, nil)
		switch enc {
		case "latin_1", "cp1252":
			utf8Bytes, err := charmap.ISO8859_1.NewDecoder().Bytes(raw)
			if err == nil {
				o.Text = string(utf8Bytes)
			} else {
				o.Text = string(raw)
			}
		default:
			o.Text = string(raw)
		}
		o.RichTextRunlist = [][2]int{{0, 0}}
		s.CellNoteMap[[2]int{o.Rowx, o.Colx}] = o
		return
	}
	if len(data) < 8 {
		return
	}
	o.Rowx = int(binary.LittleEndian.Uint16(data[0:2]))
	o.Colx = int(binary.LittleEndian.Uint16(data[2:4]))
	optionFlags := binary.LittleEndian.Uint16(data[4:6])
	o.ObjectID = int(binary.LittleEndian.Uint16(data[6:8]))
	o.Show = int((optionFlags >> 1) & 1)
	o.RowHidden = int((optionFlags >> 7) & 1)
	o.ColHidden = int((optionFlags >> 8) & 1)
	author, endpos := unpack_unicode_update_pos(data, 8, 2, -1)
	o.Author = author
	if endpos < len(data) {
		_ = endpos
	}
	if txo, ok := txos[o.ObjectID]; ok {
		o.Text = txo.Text
		o.RichTextRunlist = txo.RichTextRunlist
		s.CellNoteMap[[2]int{o.Rowx, o.Colx}] = o
	}
}

func (s *Sheet) handleTxo(bk *Book, data []byte) *MSTxo {
	if bk.BiffVersion < 80 || len(data) < 18 {
		return nil
	}
	o := &MSTxo{}
	optionFlags := binary.LittleEndian.Uint16(data[0:2])
	o.Rot = int(binary.LittleEndian.Uint16(data[2:4]))
	o.ControlInfo = data[4:10]
	cchText := int(binary.LittleEndian.Uint16(data[10:12]))
	cbRuns := int(binary.LittleEndian.Uint16(data[12:14]))
	o.IfntEmpty = int(binary.LittleEndian.Uint16(data[14:16]))
	o.Fmla = data[16:]

	o.HorzAlign = int((optionFlags >> 3) & 0x7)
	o.VertAlign = int((optionFlags >> 6) & 0x7)
	o.LockText = int((optionFlags >> 9) & 1)
	o.JustLast = int((optionFlags >> 14) & 1)
	o.SecretEdit = int((optionFlags >> 15) & 1)

	totChars := 0
	for totChars < cchText {
		rc2, data2Len, data2 := bk.getRecordParts()
		if rc2 != XL_CONTINUE || data2Len == 0 {
			break
		}
		nb := int(data2[0])
		nchars := data2Len - 1
		if nb != 0 {
			if nchars%2 != 0 {
				nchars--
			}
			nchars /= 2
		}
		utext, _ := unpack_unicode_update_pos(data2, 0, 1, nchars)
		o.Text += utext
		totChars += nchars
	}
	o.RichTextRunlist = make([][2]int, 0)
	totRuns := 0
	for totRuns < cbRuns {
		rc3, data3Len, data3 := bk.getRecordParts()
		if rc3 != XL_CONTINUE || data3Len%8 != 0 {
			break
		}
		for pos := 0; pos+8 <= data3Len; pos += 8 {
			run := [2]int{int(binary.LittleEndian.Uint16(data3[pos : pos+2])), int(binary.LittleEndian.Uint16(data3[pos+2 : pos+4]))}
			o.RichTextRunlist = append(o.RichTextRunlist, run)
			totRuns += 8
		}
	}
	for len(o.RichTextRunlist) > 0 && o.RichTextRunlist[len(o.RichTextRunlist)-1][0] == cchText {
		o.RichTextRunlist = o.RichTextRunlist[:len(o.RichTextRunlist)-1]
	}
	return o
}

func (s *Sheet) handleFeat11(bk *Book, data []byte) {
	_ = bk
	_ = data
}

// handleFormula processes XL_FORMULA, XL_FORMULA3, and XL_FORMULA4 records.
func (s *Sheet) handleFormula(bk *Book, data []byte, dataLen int) {
	if dataLen < 16 {
		return
	}

	// Parse formula record header
	rowx := int(binary.LittleEndian.Uint16(data[0:2]))
	colx := int(binary.LittleEndian.Uint16(data[2:4]))
	xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
	resultStr := data[6:14]                     // 8 bytes of cached result
	_ = binary.LittleEndian.Uint16(data[14:16]) // flags (unused for now)

	// Formula record parsed

	// Check for string result (indicated by 0xFF 0xFF in last 2 bytes of result_str)
	if len(resultStr) >= 2 && resultStr[6] == 0xFF && resultStr[7] == 0xFF {
		firstByte := resultStr[0]
		switch firstByte {
		case 0:
			// String result - need to read next STRING record
			s.handleFormulaStringResult(bk, rowx, colx, xfIndex)
		case 1:
			// Boolean result
			value := int(resultStr[2])
			s.putCell(rowx, colx, XL_CELL_BOOLEAN, value, xfIndex)
		case 2:
			// Error result
			value := int(resultStr[2])
			s.putCell(rowx, colx, XL_CELL_ERROR, value, xfIndex)
		case 3:
			// Empty string result
			s.putCell(rowx, colx, XL_CELL_TEXT, "", xfIndex)
		default:
			// Unknown special case - treat as empty
			s.putCell(rowx, colx, XL_CELL_EMPTY, nil, xfIndex)
		}
	} else {
		// Numeric result (IEEE 754 double)
		if len(resultStr) >= 8 {
			bits := binary.LittleEndian.Uint64(resultStr[0:8])
			value := math.Float64frombits(bits)
			s.putCell(rowx, colx, XL_CELL_NUMBER, value, xfIndex)
		}
	}

	// xlrd does not evaluate formulas - it only reads cached results from Excel
	// Formula evaluation is performed by Excel itself when the file was saved
}

// handleFormulaStringResult handles formulas that result in strings.
// These are followed by a STRING record containing the actual string value.
func (s *Sheet) handleFormulaStringResult(bk *Book, rowx, colx, xfIndex int) {
	// Read the next record which should be a STRING record
	rc, _, data := bk.getRecordParts()
	if rc != XL_STRING && rc != XL_STRING_B2 {
		for {
			switch rc {
			case XL_ARRAY, XL_SHRFMLA, XL_TABLEOP, XL_TABLEOP2, XL_ARRAY2, XL_TABLEOP_B2:
				rc, _, data = bk.getRecordParts()
			default:
				s.putCell(rowx, colx, XL_CELL_EMPTY, nil, xfIndex)
				return
			}
			if rc == XL_STRING || rc == XL_STRING_B2 {
				break
			}
		}
	}

	// Parse the string record using proper string record format
	strg, err := s.stringRecordContents(bk, data)
	if err != nil {
		// Error parsing string - put empty cell
		s.putCell(rowx, colx, XL_CELL_EMPTY, nil, xfIndex)
		return
	}

	s.putCell(rowx, colx, XL_CELL_TEXT, strg, xfIndex)
}

// unpackRK decodes an RK value (Real number + Key) from Excel BIFF format.
func unpackRK(rkData []byte) float64 {
	if len(rkData) != 4 {
		return 0.0
	}

	flags := rkData[0]
	if flags&2 != 0 {
		i := int32(binary.LittleEndian.Uint32(rkData))
		i >>= 2
		if flags&1 != 0 {
			return float64(i) / 100.0
		}
		return float64(i)
	}

	floatBytes := []byte{0, 0, 0, 0, rkData[0] & 0xFC, rkData[1], rkData[2], rkData[3]}
	bits := binary.LittleEndian.Uint64(floatBytes)
	result := math.Float64frombits(bits)
	if flags&1 != 0 {
		result /= 100.0
	}
	return result
}

// stringRecordContents parses a STRING record's content.
// This is different from cell strings - formula result strings have their own format.
func (s *Sheet) stringRecordContents(bk *Book, data []byte) (string, error) {
	if len(data) < 2 {
		return "", fmt.Errorf("string record too short")
	}

	bv := bk.BiffVersion
	lenlen := 1
	if bv >= 30 {
		lenlen = 2
	}

	if len(data) < lenlen {
		return "", fmt.Errorf("string record too short for length")
	}

	var nchars uint16
	if lenlen == 1 {
		nchars = uint16(data[0])
	} else {
		nchars = binary.LittleEndian.Uint16(data[0:2])
	}

	offset := lenlen

	if bv >= 80 {
		// BIFF 8+: check encoding flag
		if len(data) <= offset {
			return "", fmt.Errorf("string record too short for encoding flag")
		}
		flag := data[offset] & 1
		offset++

		if flag == 0 {
			// Latin-1
			if len(data) < offset+int(nchars) {
				return "", fmt.Errorf("string record too short for Latin-1 data")
			}
			latin1Bytes := data[offset : offset+int(nchars)]
			utf8Bytes, err := charmap.ISO8859_1.NewDecoder().Bytes(latin1Bytes)
			if err != nil {
				return "", fmt.Errorf("failed to decode Latin-1: %v", err)
			}
			return string(utf8Bytes), nil
		} else {
			// UTF-16 LE
			if len(data) < offset+int(nchars)*2 {
				return "", fmt.Errorf("string record too short for UTF-16 data")
			}
			utf16Bytes := data[offset : offset+int(nchars)*2]
			words := make([]uint16, nchars)
			for j := 0; j < int(nchars); j++ {
				words[j] = binary.LittleEndian.Uint16(utf16Bytes[j*2 : (j+1)*2])
			}
			return string(utf16.Decode(words)), nil
		}
	} else {
		// BIFF < 8: use workbook encoding
		enc := bk.Encoding
		if enc == "" {
			enc = bk.deriveEncoding()
		}
		if len(data) < offset+int(nchars) {
			return "", fmt.Errorf("string record too short for data")
		}
		bytes := data[offset : offset+int(nchars)]
		return string(bytes), nil // Assume encoding is handled elsewhere
	}
}
