package xlrd

import (
	"io"
)

// CompDocError represents an error in compound document handling.
type CompDocError struct {
	Message string
}

func (e *CompDocError) Error() string {
	return e.Message
}

// CompDoc handles OLE2 compound document files.
type CompDoc struct {
	// Mem is the raw contents of the file.
	Mem []byte

	// Logfile is the file to which messages are written.
	Logfile io.Writer

	// DEBUG is the debug level.
	DEBUG int

	// IgnoreWorkbookCorruption indicates whether to ignore workbook corruption errors.
	IgnoreWorkbookCorruption bool
}

// NewCompDoc creates a new CompDoc instance.
func NewCompDoc(mem []byte, logfile io.Writer, debug int, ignoreWorkbookCorruption bool) (*CompDoc, error) {
	// Empty implementation for now
	return &CompDoc{
		Mem:                      mem,
		Logfile:                  logfile,
		DEBUG:                    debug,
		IgnoreWorkbookCorruption: ignoreWorkbookCorruption,
	}, nil
}
