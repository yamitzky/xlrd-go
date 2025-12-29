package xlrd

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
	// Empty implementation for now
	return nil
}

// CellType returns the type of the cell at the given row and column.
func (s *Sheet) CellType(rowx, colx int) int {
	// Empty implementation for now
	return XL_CELL_EMPTY
}

// Cell returns the Cell object at the given row and column.
func (s *Sheet) Cell(rowx, colx int) *Cell {
	// Empty implementation for now
	return &Cell{CType: XL_CELL_EMPTY}
}

// Row returns a slice of Cell objects for the given row.
func (s *Sheet) Row(rowx int) []*Cell {
	// Empty implementation for now
	return []*Cell{}
}

// RowLen returns the length of the row (number of non-empty cells).
func (s *Sheet) RowLen(rowx int) int {
	// Empty implementation for now
	return 0
}

// EmptyCell returns an empty cell.
func EmptyCell() *Cell {
	return &Cell{CType: XL_CELL_EMPTY}
}
