package xlrd

import (
	"io"
)

// Dump dumps an XLS file's BIFF records in char & hex format for debugging.
//
// filename: The path to the file to be dumped.
// outfile: An open file, to which the dump is written.
// unnumbered: If true, omit offsets (for meaningful diffs).
func Dump(filename string, outfile io.Writer, unnumbered bool) error {
	bk, err := OpenWorkbook(filename, nil)
	if err != nil {
		return err
	}
	defer bk.ReleaseResources()

	BiffDump(bk.mem, bk.base, bk.streamLen, 0, outfile, unnumbered)
	return nil
}

// CountRecords summarises the file's BIFF records.
// It produces a sorted file of (record_name, count).
//
// filename: The path to the file to be summarised.
// outfile: An open file, to which the summary is written.
func CountRecords(filename string, outfile io.Writer) error {
	bk, err := OpenWorkbook(filename, nil)
	if err != nil {
		return err
	}
	defer bk.ReleaseResources()

	BiffCountRecords(bk.mem, bk.base, bk.streamLen, outfile)
	return nil
}
