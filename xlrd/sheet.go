package xlrd

import (
	"encoding/binary"
	"fmt"
	"math"
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

	// Cell data storage (similar to Python's _cell_values, _cell_types, _cell_xf_indexes)
	cellValues   [][]interface{}
	cellTypes    [][]int
	cellXFIndexes [][]int

	// Cached magnification factors from WINDOW2 record
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

	// Parse BIFF records until EOF or end of sheet stream
	for {
		if bk.position >= maxPosition {
			break
		}
		rc, dataLen, data := bk.getRecordParts()
		if rc == XL_EOF {
			break
		}


		// Debug for column B records
		switch rc {
		case XL_NUMBER, XL_NUMBER_B2:
			if dataLen >= 14 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				bits := binary.LittleEndian.Uint64(data[6:14])
				value := math.Float64frombits(bits)
				if s.Name == "PROFILEDEF" {
					fmt.Fprintf(bk.logfile, "DEBUG: %s XL_NUMBER at (%d,%d): value=%f, xf=%d\n", s.Name, rowx, colx, value, xfIndex)
				}
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
				}
			}
		case XL_RK:
			if dataLen >= 10 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				rkData := data[6:10]
				rkValue := unpackRK(rkData)
				if s.Name == "PROFILEDEF" || rkValue == 100.0 {
					fmt.Printf("DEBUG: %s XL_RK at (%d,%d): rkData=%x, value=%f, xf=%d\n", s.Name, rowx, colx, rkData, rkValue, xfIndex)
				}
				s.putCell(rowx, colx, XL_CELL_NUMBER, rkValue, xfIndex)
			}
		case XL_MULRK:
			if dataLen >= 6 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				firstColx := int(binary.LittleEndian.Uint16(data[2:4]))
				lastColx := int(binary.LittleEndian.Uint16(data[dataLen-2 : dataLen]))

				pos := 4
				for colx := firstColx; colx <= lastColx && pos+4 <= dataLen-2; colx++ {
					xfIndex := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
					rkData := data[pos+2 : pos+6]
					rkValue := unpackRK(rkData)
					s.putCell(rowx, colx, XL_CELL_NUMBER, rkValue, xfIndex)
					pos += 4
				}
			}
		case XL_LABEL:
			if dataLen >= 6 {
				rowx := int(binary.LittleEndian.Uint16(data[0:2]))
				colx := int(binary.LittleEndian.Uint16(data[2:4]))
				xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
				// Parse string using UnpackString
				if dataLen > 6 {
					value, err := UnpackString(data, 6, bk.Encoding, 1)
					if err == nil && value != "" {
						s.putCell(rowx, colx, XL_CELL_TEXT, value, xfIndex)
					}
				}
			}
		case XL_FORMULA, XL_FORMULA3, XL_FORMULA4:
			s.handleFormula(bk, data, dataLen)
		case XL_MERGEDCELLS:
			s.handleMergedCells(data, dataLen)
		}
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

// handleFormula processes XL_FORMULA, XL_FORMULA3, and XL_FORMULA4 records.
func (s *Sheet) handleFormula(bk *Book, data []byte, dataLen int) {
	if dataLen < 16 {
		return
	}

	// Parse formula record header
	rowx := int(binary.LittleEndian.Uint16(data[0:2]))
	colx := int(binary.LittleEndian.Uint16(data[2:4]))
	xfIndex := int(binary.LittleEndian.Uint16(data[4:6]))
	resultStr := data[6:14] // 8 bytes of cached result
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

	// TODO: Formula bytecode parsing would go here for actual formula evaluation
	// For now, we only use cached results as xlrd does
}

// handleFormulaStringResult handles formulas that result in strings.
// These are followed by a STRING record containing the actual string value.
func (s *Sheet) handleFormulaStringResult(bk *Book, rowx, colx, xfIndex int) {
	// Read the next record which should be a STRING record
	rc, _, data := bk.getRecordParts()
	if rc != XL_STRING && rc != XL_STRING_B2 {
		// Not a string record - put empty cell
		s.putCell(rowx, colx, XL_CELL_EMPTY, nil, xfIndex)
		return
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

	rkValue := binary.LittleEndian.Uint32(rkData)
	flags := rkValue & 3 // Lower 2 bits are flags

	if flags&2 != 0 {
		// Signed 30-bit integer
		i := int32(rkValue) >> 2 // Shift right by 2 to drop flag bits
		result := float64(i)
		if flags&1 != 0 {
			result /= 100.0
		}
		return result
	} else {
		// IEEE 754 64-bit float (30 most significant bits)
		// Reconstruct the 64-bit float from 30 bits + 2 flag bits
		// Python: b'\0\0\0\0' + BYTES_LITERAL(chr(flags & 252)) + rk_str[1:4]
		flagsByte := byte(flags)
		rkDataByte0 := rkData[0]
		msb := (flagsByte & 0xFC) | ((rkDataByte0 & 0xC0) >> 6) // Clear lower 2 bits, add 2 bits from rkData[0]
		middle := []byte{rkData[1], rkData[2], rkData[3]}

		// Create 64-bit IEEE 754 float bytes: 4 zero bytes + reconstructed bytes
		floatBytes := []byte{0, 0, 0, 0, msb, middle[0], middle[1], middle[2]}
		bits := binary.LittleEndian.Uint64(floatBytes)
		result := math.Float64frombits(bits)
		if flags&1 != 0 {
			result /= 100.0
		}
		return result
	}
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
