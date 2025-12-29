package xlrd

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"sort"
	"strings"
)


// hexCharDump dumps data in hex and character format
func hexCharDump(data []byte, offset, length int, base int, outfile io.Writer, unnumbered bool) {
	endPos := min(offset+length, len(data))
	pos := offset
	numbered := !unnumbered
	numPrefix := ""

	for pos < endPos {
		endSub := min(pos+16, endPos)
		substr := data[pos:endSub]
		lenSub := endSub - pos
		if lenSub <= 0 || lenSub != len(substr) {
			fmt.Fprintf(outfile, "??? hexCharDump: offset=%d length=%d base=%d -> endPos=%d pos=%d endSub=%d substr=%q\n",
				offset, length, base, endPos, pos, endSub, substr)
			break
		}

		// Create hex representation
		hexParts := make([]string, len(substr))
		for i, c := range substr {
			hexParts[i] = fmt.Sprintf("%02x", c)
		}
		hexStr := strings.Join(hexParts, " ")

		// Create character representation
		charParts := make([]string, len(substr))
		for i, c := range substr {
			char := byte(c)
			if char == 0 {
				charParts[i] = "~"
			} else if char >= 32 && char <= 126 {
				charParts[i] = string(char)
			} else {
				charParts[i] = "?"
			}
		}
		charStr := strings.Join(charParts, "")

		if numbered {
			numPrefix = fmt.Sprintf("%5d: ", base+pos-offset)
		}

		fmt.Fprintf(outfile, "%s     %-48s %s\n", numPrefix, hexStr, charStr)
		pos = endSub
	}
}

// biffDump dumps BIFF records from memory
func biffDump(mem []byte, streamOffset, streamLen int, base int, outfile io.Writer, unnumbered bool) {
	pos := streamOffset
	streamEnd := streamOffset + streamLen
	adj := base - streamOffset
	dummies := 0
	savPos := 0
	numbered := !unnumbered
	numPrefix := ""
	var length int

	for streamEnd-pos >= 4 {
		rc := int(binary.LittleEndian.Uint16(mem[pos : pos+2]))
		length = int(binary.LittleEndian.Uint16(mem[pos+2 : pos+4]))

		if rc == 0 && length == 0 {
			remainingBytes := streamEnd - pos
			if pos+remainingBytes <= len(mem) && bytes.Equal(mem[pos:streamEnd], make([]byte, remainingBytes)) {
				dummies = remainingBytes
				savPos = pos
				pos = streamEnd
				break
			}
			if dummies != 0 {
				dummies += 4
			} else {
				savPos = pos
				dummies = 4
			}
			pos += 4
		} else {
			if dummies != 0 {
				if numbered {
					numPrefix = fmt.Sprintf("%5d: ", adj+savPos)
				}
				fmt.Fprintf(outfile, "%s---- %d zero bytes skipped ----\n", numPrefix, dummies)
				dummies = 0
			}
			recName := biffRecNameDict[rc]
			if recName == "" {
				recName = "<UNKNOWN>"
			}
			if numbered {
				numPrefix = fmt.Sprintf("%5d: ", adj+pos)
			}
			fmt.Fprintf(outfile, "%s%04x %s len = %04x (%d)\n", numPrefix, rc, recName, length, length)
			pos += 4
			hexCharDump(mem, pos, length, adj+pos, outfile, unnumbered)
			pos += length
		}
	}

	if dummies != 0 {
		if numbered {
			numPrefix = fmt.Sprintf("%5d: ", adj+savPos)
		}
		fmt.Fprintf(outfile, "%s---- %d zero bytes skipped ----\n", numPrefix, dummies)
	}

	if pos < streamEnd {
		if numbered {
			numPrefix = fmt.Sprintf("%5d: ", adj+pos)
		}
		fmt.Fprintf(outfile, "%s---- Misc bytes at end ----\n", numPrefix)
		hexCharDump(mem, pos, streamEnd-pos, adj+pos, outfile, unnumbered)
	} else if pos > streamEnd {
		fmt.Fprintf(outfile, "Last dumped record has length (%d) that is too large\n", length)
	}
}

// biffCountRecords counts BIFF records and produces a summary
func biffCountRecords(mem []byte, streamOffset, streamLen int, outfile io.Writer) {
	pos := streamOffset
	streamEnd := streamOffset + streamLen
	tally := make(map[string]int)

	for streamEnd-pos >= 4 {
		rc := int(binary.LittleEndian.Uint16(mem[pos : pos+2]))
		length := int(binary.LittleEndian.Uint16(mem[pos+2 : pos+4]))

		if rc == 0 && length == 0 {
			remainingBytes := streamEnd - pos
			if pos+remainingBytes <= len(mem) && bytes.Equal(mem[pos:streamEnd], make([]byte, remainingBytes)) {
				break
			}
			recName := "<Dummy (zero)>"
			if count, exists := tally[recName]; exists {
				tally[recName] = count + 1
			} else {
				tally[recName] = 1
			}
		} else {
			recName := biffRecNameDict[rc]
			if recName == "" {
				recName = fmt.Sprintf("Unknown_0x%04X", rc)
			}
			if count, exists := tally[recName]; exists {
				tally[recName] = count + 1
			} else {
				tally[recName] = 1
			}
		}
		pos += length + 4
	}

	// Sort by record name
	type recordCount struct {
		name  string
		count int
	}
	var sortedRecords []recordCount
	for name, count := range tally {
		sortedRecords = append(sortedRecords, recordCount{name, count})
	}
	sort.Slice(sortedRecords, func(i, j int) bool {
		return sortedRecords[i].name < sortedRecords[j].name
	})

	// Print results
	for _, rc := range sortedRecords {
		fmt.Fprintf(outfile, "%8d %s\n", rc.count, rc.name)
	}
}

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

	biffDump(bk.mem, bk.base, bk.streamLen, 0, outfile, unnumbered)
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

	biffCountRecords(bk.mem, bk.base, bk.streamLen, outfile)
	return nil
}
