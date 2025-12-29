package xlrd

import (
	"path/filepath"
	"runtime"
)

// fromSample returns the path to a test sample file.
func fromSample(filename string) string {
	_, testFile, _, _ := runtime.Caller(1)
	testDir := filepath.Dir(testFile)
	projectRoot := filepath.Join(testDir, "..")
	return filepath.Join(projectRoot, "testdata", "samples", filename)
}
