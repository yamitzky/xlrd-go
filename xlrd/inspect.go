package xlrd

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"strings"
)

// FileFormatDescriptions provides descriptions of the file types that can be inspected.
var FileFormatDescriptions = map[string]string{
	"xls":  "Excel xls",
	"xlsb": "Excel 2007 xlsb file",
	"xlsx": "Excel xlsx file",
	"ods":  "Openoffice.org ODS file",
	"zip":  "Unknown ZIP file",
	"":     "Unknown file type",
}

// XLS_SIGNATURE is the magic cookie that should appear in the first 8 bytes of an XLS file.
var XLS_SIGNATURE = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// ZIP_SIGNATURE is the magic cookie for ZIP files.
var ZIP_SIGNATURE = []byte("PK\x03\x04")

// PEEK_SIZE is the maximum size needed to peek at file signatures.
const PEEK_SIZE = 8

// InspectFormat inspects the content at the supplied path or the bytes content provided
// and returns the file's type as a string, or empty string if it cannot be determined.
//
// path: A string path containing the content to inspect. ~ will be expanded.
// content: The bytes content to inspect.
//
// Returns: A string, or empty string if the format cannot be determined.
// The return value can always be looked up in FileFormatDescriptions
// to return a human-readable description of the format found.
func InspectFormat(path string, content []byte) (string, error) {
	var peek []byte
	var err error

	if content != nil {
		if len(content) < PEEK_SIZE {
			return "", nil
		}
		peek = content[:PEEK_SIZE]
	} else {
		expandedPath := path
		if strings.HasPrefix(path, "~") {
			homeDir, err := os.UserHomeDir()
			if err != nil {
				return "", err
			}
			expandedPath = strings.Replace(path, "~", homeDir, 1)
		}

		f, err := os.Open(expandedPath)
		if err != nil {
			return "", err
		}
		defer f.Close()

		peek = make([]byte, PEEK_SIZE)
		n, err := f.Read(peek)
		if err != nil && err != io.EOF {
			return "", err
		}
		peek = peek[:n]
	}

	if len(peek) < len(XLS_SIGNATURE) {
		return "", nil
	}

	if bytes.HasPrefix(peek, XLS_SIGNATURE) {
		return "xls", nil
	}

	if bytes.HasPrefix(peek, ZIP_SIGNATURE) {
		var zf *zip.Reader
		if content != nil {
			zf, err = zip.NewReader(bytes.NewReader(content), int64(len(content)))
		} else {
			expandedPath := path
			if strings.HasPrefix(path, "~") {
				homeDir, err := os.UserHomeDir()
				if err != nil {
					return "", err
				}
				expandedPath = strings.Replace(path, "~", homeDir, 1)
			}
			r, err := zip.OpenReader(expandedPath)
			if err != nil {
				return "", err
			}
			defer r.Close()
			zf = &r.Reader
		}

		if err != nil {
			return "", err
		}

		// Workaround for some third party files that use forward slashes and
		// lower case names. We map the expected name in lowercase to the
		// actual filename in the zip container.
		componentNames := make(map[string]string)
		for _, name := range zf.File {
			lowerName := strings.ToLower(strings.ReplaceAll(name.Name, "\\", "/"))
			componentNames[lowerName] = name.Name
		}

		if _, ok := componentNames["xl/workbook.xml"]; ok {
			return "xlsx", nil
		}
		if _, ok := componentNames["xl/workbook.bin"]; ok {
			return "xlsb", nil
		}
		if _, ok := componentNames["content.xml"]; ok {
			return "ods", nil
		}
		return "zip", nil
	}

	return "", nil
}
