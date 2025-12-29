package xlrd

import (
	"encoding/binary"
	"math"
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
	if rowx >= len(s.cellValues) || colx >= len(s.cellValues[rowx]) {
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
	if rowx < 0 || rowx >= len(s.cellTypes) || colx < 0 || colx >= len(s.cellTypes[rowx]) {
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
	if rowx >= len(s.cellXFIndexes) || colx >= len(s.cellXFIndexes[rowx]) || s.cellXFIndexes[rowx][colx] == 0 {
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
	// Initialize cell arrays
	s.cellValues = make([][]interface{}, 0)
	s.cellTypes = make([][]int, 0)
	s.cellXFIndexes = make([][]int, 0)

	// Parse BIFF records
	for {
		rc, dataLen, data := bk.getRecordParts()
		if rc == XL_EOF {
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
		}
	}

	return nil
}
