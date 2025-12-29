package xlrd

// FormulaError represents an error in formula parsing.
type FormulaError struct {
	Message string
}

func (e *FormulaError) Error() string {
	return e.Message
}

// DecompileFormula decompiles a formula into a human-readable string.
func DecompileFormula(bk *Book, fmla []byte, fmlalen int, bv int, reldelta int, blah int) (string, error) {
	// Empty implementation for now
	return "", nil
}

// DumpFormula dumps a formula for debugging purposes.
func DumpFormula(bk *Book, data []byte, fmlalen int, bv int, reldelta int, blah int, isname int) {
	// Empty implementation for now
}

// CellName returns the cell name for a given row and column (0-based).
// Example: CellName(0, 0) returns "A1"
func CellName(rowx, colx int) string {
	// Empty implementation for now
	return ""
}

// CellNameAbs returns the absolute cell name.
func CellNameAbs(rowx, colx int, r1c1 bool) string {
	// Empty implementation for now
	return ""
}

// RangeName2D returns a 2D range name.
func RangeName2D(rlo, rhi, clo, chi int, r1c1 bool) string {
	// Empty implementation for now
	return ""
}
