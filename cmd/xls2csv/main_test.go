package main

import (
	"bytes"
	"encoding/csv"
	"path/filepath"
	"strings"
	"testing"

	"github.com/yamitzky/xlrd-go/xlrd"
)

func TestRunDefault(t *testing.T) {
	sample := samplePath(t, "Formate.xls")
	out, errOut, code := runCLI([]string{sample})
	if code != 0 {
		t.Fatalf("exit code %d, stderr: %s", code, errOut)
	}
	record := firstRecord(t, out, ',')
	if len(record) < 2 {
		t.Fatalf("expected at least 2 fields, got %d", len(record))
	}
	if record[0] != "Huber" {
		t.Fatalf("field[0]=%q, want %q", record[0], "Huber")
	}

	book, err := xlrd.OpenWorkbook(sample, &xlrd.OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("sheet: %v", err)
	}
	expected, _ := formatCell(book, sheet, 0, 1, options{})
	if record[1] != expected {
		t.Fatalf("field[1]=%q, want %q", record[1], expected)
	}
}

func TestRunFloatFormat(t *testing.T) {
	out, errOut, code := runCLI([]string{"--floatformat", "%.2f", "-s", "3", samplePath(t, "Formate.xls")})
	if code != 0 {
		t.Fatalf("exit code %d, stderr: %s", code, errOut)
	}
	record := firstRecord(t, out, ',')
	if len(record) < 1 {
		t.Fatalf("expected at least 1 field, got %d", len(record))
	}
	if record[0] != "100.00" {
		t.Fatalf("field[0]=%q, want %q", record[0], "100.00")
	}
}

func TestRunDelimiterTab(t *testing.T) {
	out, errOut, code := runCLI([]string{"-d", "tab", samplePath(t, "Formate.xls")})
	if code != 0 {
		t.Fatalf("exit code %d, stderr: %s", code, errOut)
	}
	if !strings.Contains(out, "\t") {
		t.Fatalf("expected tab delimiter, got output: %q", firstLine(out))
	}
	record := firstRecord(t, out, '\t')
	if len(record) < 2 || record[0] != "Huber" {
		t.Fatalf("unexpected first record: %v", record)
	}
}

func TestRunDateFormat(t *testing.T) {
	book, err := xlrd.OpenWorkbook(samplePath(t, "Formate.xls"), &xlrd.OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(0)
	if err != nil {
		t.Fatalf("sheet: %v", err)
	}
	xfIndex := sheet.CellXFIndex(0, 1)
	if !isDateCell(book, xfIndex) {
		t.Skip("date format not detected for sample cell")
	}
	dateValue, ok := toFloat(sheet.CellValue(0, 1))
	if !ok {
		t.Fatalf("date value is not numeric")
	}
	dt, err := xlrd.XldateAsDatetime(dateValue, book.Datemode)
	if err != nil {
		t.Fatalf("xldate: %v", err)
	}
	expected := strftime(dt, "%Y/%m/%d")

	out, errOut, code := runCLI([]string{"-f", "%Y/%m/%d", samplePath(t, "Formate.xls")})
	if code != 0 {
		t.Fatalf("exit code %d, stderr: %s", code, errOut)
	}
	record := firstRecord(t, out, ',')
	if len(record) < 2 {
		t.Fatalf("expected at least 2 fields, got %d", len(record))
	}
	if record[1] != expected {
		t.Fatalf("field[1]=%q, want %q", record[1], expected)
	}
}

func TestRunMergeCells(t *testing.T) {
	book, err := xlrd.OpenWorkbook(samplePath(t, "Formate.xls"), &xlrd.OpenWorkbookOptions{FormattingInfo: true})
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	sheet, err := book.SheetByIndex(1)
	if err != nil {
		t.Fatalf("sheet: %v", err)
	}

	rowx, colx, mergedText := findMergedCell(sheet)
	if mergedText == "" {
		t.Skip("no merged cell found with non-empty value")
	}

	out, errOut, code := runCLI([]string{"-s", "2", samplePath(t, "Formate.xls")})
	if code != 0 {
		t.Fatalf("exit code %d, stderr: %s", code, errOut)
	}
	rawCell := csvCellAt(t, out, rowx, colx)
	if rawCell != "" {
		t.Fatalf("expected empty cell without -m, got %q", rawCell)
	}

	outMerged, errOutMerged, codeMerged := runCLI([]string{"-m", "-s", "2", samplePath(t, "Formate.xls")})
	if codeMerged != 0 {
		t.Fatalf("exit code %d, stderr: %s", codeMerged, errOutMerged)
	}
	mergedCell := csvCellAt(t, outMerged, rowx, colx)
	if mergedCell != mergedText {
		t.Fatalf("merged cell=%q, want %q", mergedCell, mergedText)
	}
}

func runCLI(args []string) (string, string, int) {
	var stdout, stderr bytes.Buffer
	code := run(args, strings.NewReader(""), &stdout, &stderr)
	return stdout.String(), stderr.String(), code
}

func samplePath(t *testing.T, name string) string {
	t.Helper()
	return filepath.Join("..", "..", "testdata", "samples", name)
}

func firstRecord(t *testing.T, output string, delimiter rune) []string {
	t.Helper()
	reader := csv.NewReader(strings.NewReader(output))
	reader.Comma = delimiter
	reader.FieldsPerRecord = -1
	record, err := reader.Read()
	if err != nil {
		t.Fatalf("read csv: %v", err)
	}
	return record
}

func firstLine(output string) string {
	if idx := strings.IndexByte(output, '\n'); idx >= 0 {
		return output[:idx]
	}
	return output
}

func csvCellAt(t *testing.T, output string, rowx, colx int) string {
	t.Helper()
	reader := csv.NewReader(strings.NewReader(output))
	reader.Comma = ','
	reader.FieldsPerRecord = -1
	row := 0
	for {
		record, err := reader.Read()
		if err != nil {
			t.Fatalf("read csv: %v", err)
		}
		if row == rowx {
			if colx >= len(record) {
				return ""
			}
			return record[colx]
		}
		row++
	}
}

func findMergedCell(sheet *xlrd.Sheet) (int, int, string) {
	for _, merged := range sheet.MergedCells {
		rlo, rhi, clo, chi := merged[0], merged[1], merged[2], merged[3]
		for row := rlo; row < rhi; row++ {
			for col := clo; col < chi; col++ {
				if row == rlo && col == clo {
					continue
				}
				raw := sheet.RawCellValue(row, col)
				mergedVal := sheet.CellValue(row, col)
				rawText := toString(raw)
				mergedText := toString(mergedVal)
				if rawText == "" && mergedText != "" {
					return row, col, mergedText
				}
			}
		}
	}
	return 0, 0, ""
}
