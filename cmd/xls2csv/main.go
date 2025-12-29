package main

import (
	"bufio"
	"bytes"
	"errors"
	"flag"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/yamitzky/xlrd-go/xlrd"
)

const defaultSheetDelimiter = "--------"

var version = "dev"

type quotingMode int

const (
	quotingNone quotingMode = iota
	quotingMinimal
	quotingNonNumeric
	quotingAll
)

type stringList []string

func (s *stringList) String() string {
	return strings.Join(*s, ",")
}

func (s *stringList) Set(value string) error {
	*s = append(*s, value)
	return nil
}

type options struct {
	allSheets           bool
	sheetID             int
	sheetName           string
	delimiter           rune
	lineTerminator      string
	dateFormat          string
	floatFormat         string
	outputEncoding      string
	ignoreEmpty         bool
	escape              bool
	sheetDelimiter      string
	quoting             quotingMode
	includeSheetPattern []*regexp.Regexp
	excludeSheetPattern []*regexp.Regexp
	mergeCells          bool
	ignoreWorkbookCorruption bool
}

type csvWriter struct {
	w              io.Writer
	delimiter      rune
	lineTerminator string
	quoting        quotingMode
}

type field struct {
	text      string
	isNumeric bool
}

func main() {
	os.Exit(run(os.Args[1:], os.Stdin, os.Stdout, os.Stderr))
}

func run(args []string, stdin io.Reader, stdout, stderr io.Writer) int {
	var includePatterns stringList
	var excludePatterns stringList

	fs := flag.NewFlagSet("xls2csv", flag.ContinueOnError)
	fs.SetOutput(stderr)

	showVersion := fs.Bool("v", false, "show version")
	fs.BoolVar(showVersion, "version", false, "show version")
	allSheets := fs.Bool("a", false, "export all sheets")
	fs.BoolVar(allSheets, "all", false, "export all sheets")

	outputEncoding := fs.String("c", "utf-8", "output CSV encoding")
	fs.StringVar(outputEncoding, "outputencoding", "utf-8", "output CSV encoding")

	sheetID := fs.Int("s", -1, "sheet number to convert, 0 for all")
	fs.IntVar(sheetID, "sheet", -1, "sheet number to convert, 0 for all")

	sheetName := fs.String("n", "", "sheet name to convert")
	fs.StringVar(sheetName, "sheetname", "", "sheet name to convert")

	delimiterFlag := fs.String("d", ",", "delimiter")
	fs.StringVar(delimiterFlag, "delimiter", ",", "delimiter")

	lineTerminatorFlag := fs.String("l", "", "line terminator")
	fs.StringVar(lineTerminatorFlag, "lineterminator", "", "line terminator")

	dateFormat := fs.String("f", "", "override date/time format")
	fs.StringVar(dateFormat, "dateformat", "", "override date/time format")

	floatFormat := fs.String("floatformat", "", "override float format")

	ignoreEmpty := fs.Bool("i", false, "skip empty lines")
	fs.BoolVar(ignoreEmpty, "ignoreempty", false, "skip empty lines")

	escape := fs.Bool("e", false, "escape \\r\\n\\t characters")
	fs.BoolVar(escape, "escape", false, "escape \\r\\n\\t characters")

	ignoreWorkbookCorruption := fs.Bool("ignore-workbook-corruption", false, "ignore workbook corruption")

	sheetDelimiter := fs.String("p", defaultSheetDelimiter, "sheet delimiter")
	fs.StringVar(sheetDelimiter, "sheetdelimiter", defaultSheetDelimiter, "sheet delimiter")

	quotingFlag := fs.String("q", "minimal", "field quoting")
	fs.StringVar(quotingFlag, "quoting", "minimal", "field quoting")

	hyperlinks := fs.Bool("hyperlinks", false, "include hyperlinks")

	fs.Var(&includePatterns, "I", "include sheet patterns")
	fs.Var(&includePatterns, "include_sheet_pattern", "include sheet patterns")
	fs.Var(&excludePatterns, "E", "exclude sheet patterns")
	fs.Var(&excludePatterns, "exclude_sheet_pattern", "exclude sheet patterns")

	mergeCells := fs.Bool("m", false, "merge cells")
	fs.BoolVar(mergeCells, "merge-cells", false, "merge cells")

	fs.Usage = func() {
		fmt.Fprint(stderr, usageText())
	}

	if err := fs.Parse(args); err != nil {
		if errors.Is(err, flag.ErrHelp) {
			return 0
		}
		return 2
	}

	if *showVersion {
		fmt.Fprintln(stdout, version)
		return 0
	}

	if *hyperlinks {
		fmt.Fprintln(stderr, "hyperlinks are not supported")
		return 2
	}

	rest := fs.Args()
	if len(rest) < 1 {
		fs.Usage()
		return 2
	}

	if *sheetName != "" && (*allSheets || *sheetID >= 0) {
		fmt.Fprintln(stderr, "cannot combine --sheetname with --sheet or --all")
		return 2
	}

	if strings.ToLower(*outputEncoding) != "utf-8" && strings.ToLower(*outputEncoding) != "utf8" {
		fmt.Fprintf(stderr, "unsupported output encoding: %s\n", *outputEncoding)
		return 2
	}

	delimiter, err := parseDelimiter(*delimiterFlag)
	if err != nil {
		fmt.Fprintf(stderr, "invalid delimiter: %v\n", err)
		return 2
	}

	lineTerminator := *lineTerminatorFlag
	if lineTerminator == "" {
		lineTerminator = osLineSep()
	} else {
		lineTerminator, err = parseEscapedString(lineTerminator)
		if err != nil {
			fmt.Fprintf(stderr, "invalid line terminator: %v\n", err)
			return 2
		}
	}

	sheetDelimiterValue := *sheetDelimiter
	if sheetDelimiterValue != "" {
		sheetDelimiterValue, err = parseSheetDelimiter(sheetDelimiterValue)
		if err != nil {
			fmt.Fprintf(stderr, "invalid sheet delimiter: %v\n", err)
			return 2
		}
	}

	quoting, err := parseQuoting(*quotingFlag)
	if err != nil {
		fmt.Fprintf(stderr, "invalid quoting: %v\n", err)
		return 2
	}

	includeRegex, err := compilePatterns(includePatterns)
	if err != nil {
		fmt.Fprintf(stderr, "invalid include pattern: %v\n", err)
		return 2
	}
	excludeRegex, err := compilePatterns(excludePatterns)
	if err != nil {
		fmt.Fprintf(stderr, "invalid exclude pattern: %v\n", err)
		return 2
	}

	opts := options{
		allSheets:           *allSheets || *sheetID == 0,
		sheetID:             *sheetID,
		sheetName:           *sheetName,
		delimiter:           delimiter,
		lineTerminator:      lineTerminator,
		dateFormat:          *dateFormat,
		floatFormat:         *floatFormat,
		outputEncoding:      *outputEncoding,
		ignoreEmpty:         *ignoreEmpty,
		escape:              *escape,
		sheetDelimiter:      sheetDelimiterValue,
		quoting:             quoting,
		includeSheetPattern: includeRegex,
		excludeSheetPattern: excludeRegex,
		mergeCells:          *mergeCells,
		ignoreWorkbookCorruption: *ignoreWorkbookCorruption,
	}

	inputPath := rest[0]
	outputPath := ""
	if len(rest) > 1 {
		outputPath = rest[1]
	}

	if inputPath == "-" {
		content, err := io.ReadAll(stdin)
		if err != nil {
			fmt.Fprintf(stderr, "failed to read stdin: %v\n", err)
			return 1
		}
		if err := convertFile("-", content, outputPath, opts, stdout); err != nil {
			fmt.Fprintln(stderr, err)
			return 1
		}
		return 0
	}

	info, err := os.Stat(inputPath)
	if err != nil {
		fmt.Fprintln(stderr, err)
		return 1
	}

	if info.IsDir() {
		if err := convertDir(inputPath, outputPath, opts, stdout); err != nil {
			fmt.Fprintln(stderr, err)
			return 1
		}
		return 0
	}

	if err := convertFile(inputPath, nil, outputPath, opts, stdout); err != nil {
		fmt.Fprintln(stderr, err)
		return 1
	}
	return 0
}

func usageText() string {
	return `Usage:

 xls2csv [-h] [-v] [-a] [-c OUTPUTENCODING] [-s SHEETID]
                   [-n SHEETNAME] [-d DELIMITER] [-l LINETERMINATOR]
                   [-f DATEFORMAT] [--floatformat FLOATFORMAT]
                   [-i] [-e] [-p SHEETDELIMITER]
                   [--hyperlinks]
                   [-I INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...]]
                   [-E EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...]] [-m]
                   xlsxfile [outfile]
positional arguments:

  xlsxfile              xlsx file path, use '-' to read from STDIN
  outfile               output csv file path, or directory if -s 0 is specified
optional arguments:

  -h, --help            show this help message and exit
  -v, --version         show program's version number and exit
  -a, --all             export all sheets
  -c OUTPUTENCODING, --outputencoding OUTPUTENCODING
                        encoding of output CSV **Python 3 only** (default: utf-8)
  -s SHEETID, --sheet SHEETID
                        sheet number to convert, 0 for all
  -n SHEETNAME, --sheetname SHEETNAME
                        sheet name to convert
  -d DELIMITER, --delimiter DELIMITER
                        delimiter - column delimiter in CSV, 'tab' or 'x09'
                        for a tab (default: comma ',')
  -l LINETERMINATOR, --lineterminator LINETERMINATOR
                        line terminator - line terminator in CSV, '\n' '\r\n'
                        or '\r' (default: os.linesep)
  -f DATEFORMAT, --dateformat DATEFORMAT
                        override date/time format (ex. %Y/%m/%d)
  --floatformat FLOATFORMAT
                        override float format (ex. %.15f)
  -i, --ignoreempty     skip empty lines
  -e, --escape          escape \r\n\t characters
  -p SHEETDELIMITER, --sheetdelimiter SHEETDELIMITER
                        sheet delimiter used to separate sheets, pass '' if
                        you do not need a delimiter, or 'x07' or '\\f' for form
                        feed (default: '--------')
  -q QUOTING, --quoting QUOTING
                        field quoting, 'none' 'minimal' 'nonnumeric' or 'all' (default: 'minimal')
  --hyperlinks
                        include hyperlinks
  -I INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...], --include_sheet_pattern INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...]
                        only include sheets with names matching the given pattern, only
                        affects when -a option is enabled.
  -E EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...], --exclude_sheet_pattern EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...]
                        exclude sheets with names matching the given pattern, only
                        affects when -a option is enabled.
  -m, --merge-cells     merge cells
Usage with a folder containing multiple xlsx files:

    xls2csv /path/to/input/dir /path/to/output/dir
will output each file in the input directory converted to .csv in the output directory. If omitting the output directory, it will output the converted files in the input directory.
`
}

func parseDelimiter(value string) (rune, error) {
	switch strings.ToLower(value) {
	case "tab", "x09":
		return '\t', nil
	}
	if value == "" {
		return 0, fmt.Errorf("delimiter cannot be empty")
	}
	if strings.HasPrefix(value, "x") && len(value) == 3 {
		decoded, err := strconv.ParseUint(value[1:], 16, 8)
		if err != nil {
			return 0, err
		}
		return rune(decoded), nil
	}
	r, _ := utf8DecodeRune(value)
	return r, nil
}

func parseSheetDelimiter(value string) (string, error) {
	switch value {
	case "\\f":
		return "\f", nil
	}
	if strings.HasPrefix(value, "x") && len(value) == 3 {
		decoded, err := strconv.ParseUint(value[1:], 16, 8)
		if err != nil {
			return "", err
		}
		return string([]byte{byte(decoded)}), nil
	}
	return value, nil
}

func parseEscapedString(value string) (string, error) {
	var b strings.Builder
	for i := 0; i < len(value); i++ {
		if value[i] != '\\' {
			b.WriteByte(value[i])
			continue
		}
		if i+1 >= len(value) {
			return "", fmt.Errorf("dangling escape")
		}
		i++
		switch value[i] {
		case 'n':
			b.WriteByte('\n')
		case 'r':
			b.WriteByte('\r')
		case 't':
			b.WriteByte('\t')
		case '\\':
			b.WriteByte('\\')
		default:
			return "", fmt.Errorf("unknown escape \\%c", value[i])
		}
	}
	return b.String(), nil
}

func parseQuoting(value string) (quotingMode, error) {
	switch strings.ToLower(value) {
	case "none":
		return quotingNone, nil
	case "minimal":
		return quotingMinimal, nil
	case "nonnumeric":
		return quotingNonNumeric, nil
	case "all":
		return quotingAll, nil
	default:
		return quotingMinimal, fmt.Errorf("unsupported quoting: %s", value)
	}
}

func compilePatterns(values []string) ([]*regexp.Regexp, error) {
	if len(values) == 0 {
		return nil, nil
	}
	patterns := make([]*regexp.Regexp, 0, len(values))
	for _, value := range values {
		re, err := regexp.Compile(value)
		if err != nil {
			return nil, err
		}
		patterns = append(patterns, re)
	}
	return patterns, nil
}

func osLineSep() string {
	if runtimeGOOS() == "windows" {
		return "\r\n"
	}
	return "\n"
}

func runtimeGOOS() string {
	return runtime.GOOS
}

func convertDir(inputDir, outputDir string, opts options, stdout io.Writer) error {
	if outputDir == "" {
		outputDir = inputDir
	}
	info, err := os.Stat(outputDir)
	if err != nil {
		if err := os.MkdirAll(outputDir, 0o755); err != nil {
			return err
		}
	} else if !info.IsDir() {
		return fmt.Errorf("output path is not a directory: %s", outputDir)
	}
	entries, err := os.ReadDir(inputDir)
	if err != nil {
		return err
	}
	found := false
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		inputPath := filepath.Join(inputDir, entry.Name())
		format, err := xlrd.InspectFormat(inputPath, nil)
		if err != nil {
			return err
		}
		if format != "xls" {
			continue
		}
		found = true
		outputPath := filepath.Join(outputDir, changeExt(entry.Name(), ".csv"))
		if err := convertFile(inputPath, nil, outputPath, opts, stdout); err != nil {
			return err
		}
	}
	if !found {
		return fmt.Errorf("no xls files found in %s", inputDir)
	}
	return nil
}

func convertFile(inputPath string, content []byte, outputPath string, opts options, stdout io.Writer) error {
	openOpts := &xlrd.OpenWorkbookOptions{
		FormattingInfo: true,
		FileContents:   content,
		IgnoreWorkbookCorruption: opts.ignoreWorkbookCorruption,
	}
	book, err := xlrd.OpenWorkbook(inputPath, openOpts)
	if err != nil {
		return err
	}

	sheetIndexes, err := selectSheets(book, opts)
	if err != nil {
		return err
	}

	if opts.sheetID == 0 && outputPath != "" {
		info, err := os.Stat(outputPath)
		if err != nil {
			if err := os.MkdirAll(outputPath, 0o755); err != nil {
				return err
			}
		} else if !info.IsDir() {
			return fmt.Errorf("outfile must be a directory when -s 0 is specified")
		}
	}

	if outputPath == "" {
		writer := bufio.NewWriter(stdout)
		if err := writeSheets(writer, book, sheetIndexes, opts); err != nil {
			return err
		}
		return writer.Flush()
	}

	info, err := os.Stat(outputPath)
	if err == nil && info.IsDir() {
		base := strings.TrimSuffix(filepath.Base(inputPath), filepath.Ext(inputPath))
		for _, sheetIndex := range sheetIndexes {
			sheet, err := book.SheetByIndex(sheetIndex)
			if err != nil {
				return err
			}
			filename := fmt.Sprintf("%s-%s.csv", base, sanitizeFilename(sheet.Name))
			fullPath := filepath.Join(outputPath, filename)
			if err := writeSheetToFile(fullPath, book, sheetIndex, opts); err != nil {
				return err
			}
		}
		return nil
	}

	if err := writeSheetsToFile(outputPath, book, sheetIndexes, opts); err != nil {
		return err
	}
	return nil
}

func selectSheets(book *xlrd.Book, opts options) ([]int, error) {
	if opts.sheetName != "" {
		for i, name := range book.SheetNames() {
			if name == opts.sheetName {
				return []int{i}, nil
			}
		}
		return nil, fmt.Errorf("sheet %s not found", opts.sheetName)
	}

	if opts.allSheets {
		names := book.SheetNames()
		indexes := make([]int, 0, len(names))
		for i, name := range names {
			if !matchPatterns(name, opts.includeSheetPattern, opts.excludeSheetPattern) {
				continue
			}
			indexes = append(indexes, i)
		}
		if len(indexes) == 0 {
			return nil, fmt.Errorf("no sheets matched selection")
		}
		return indexes, nil
	}

	if opts.sheetID > 0 {
		index := opts.sheetID - 1
		if index < 0 || index >= book.NSheets {
			return nil, fmt.Errorf("sheet index %d out of range", opts.sheetID)
		}
		return []int{index}, nil
	}

	if book.NSheets == 0 {
		return nil, fmt.Errorf("no sheets found")
	}
	return []int{0}, nil
}

func matchPatterns(name string, include, exclude []*regexp.Regexp) bool {
	if len(include) > 0 {
		matched := false
		for _, re := range include {
			if re.MatchString(name) {
				matched = true
				break
			}
		}
		if !matched {
			return false
		}
	}
	for _, re := range exclude {
		if re.MatchString(name) {
			return false
		}
	}
	return true
}

func writeSheetToFile(path string, book *xlrd.Book, sheetIndex int, opts options) error {
	return writeSheetsToFile(path, book, []int{sheetIndex}, opts)
}

func writeSheetsToFile(path string, book *xlrd.Book, sheetIndexes []int, opts options) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := bufio.NewWriter(file)
	if err := writeSheets(writer, book, sheetIndexes, opts); err != nil {
		return err
	}
	return writer.Flush()
}

func writeSheets(w io.Writer, book *xlrd.Book, sheetIndexes []int, opts options) error {
	cw := &csvWriter{
		w:              w,
		delimiter:      opts.delimiter,
		lineTerminator: opts.lineTerminator,
		quoting:        opts.quoting,
	}

	for i, sheetIndex := range sheetIndexes {
		if i > 0 && opts.sheetDelimiter != "" {
			if _, err := fmt.Fprint(w, opts.sheetDelimiter, opts.lineTerminator); err != nil {
				return err
			}
		}
		sheet, err := book.SheetByIndex(sheetIndex)
		if err != nil {
			return err
		}
		if err := writeSheet(cw, book, sheet, opts); err != nil {
			return err
		}
	}
	return nil
}

func writeSheet(cw *csvWriter, book *xlrd.Book, sheet *xlrd.Sheet, opts options) error {
	for rowx := 0; rowx < sheet.NRows; rowx++ {
		fields := make([]field, sheet.NCols)
		allEmpty := true
		for colx := 0; colx < sheet.NCols; colx++ {
			text, isNumeric := formatCell(book, sheet, rowx, colx, opts)
			if text != "" {
				allEmpty = false
			}
			fields[colx] = field{text: text, isNumeric: isNumeric}
		}
		if opts.ignoreEmpty && allEmpty {
			continue
		}
		if err := cw.writeRow(fields); err != nil {
			return err
		}
	}
	return nil
}

func formatCell(book *xlrd.Book, sheet *xlrd.Sheet, rowx, colx int, opts options) (string, bool) {
	var ctype int
	var value interface{}
	var xfIndex int
	if opts.mergeCells {
		ctype = sheet.CellType(rowx, colx)
		value = sheet.CellValue(rowx, colx)
		xfIndex = sheet.CellXFIndex(rowx, colx)
	} else {
		ctype = sheet.RawCellType(rowx, colx)
		value = sheet.RawCellValue(rowx, colx)
		xfIndex = sheet.RawCellXFIndex(rowx, colx)
	}

	switch ctype {
	case xlrd.XL_CELL_TEXT:
		return maybeEscape(toString(value), opts.escape), false
	case xlrd.XL_CELL_NUMBER:
		if val, ok := toFloat(value); ok && isDateCell(book, xfIndex) {
			if formatted, ok := formatDate(val, book.Datemode, opts.dateFormat); ok {
				return maybeEscape(formatted, opts.escape), false
			}
		}
		return maybeEscape(formatFloat(value, opts.floatFormat), opts.escape), true
	case xlrd.XL_CELL_BOOLEAN:
		return maybeEscape(formatBool(value), opts.escape), false
	case xlrd.XL_CELL_ERROR:
		return maybeEscape(formatError(value), opts.escape), false
	case xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK:
		return "", false
	default:
		return maybeEscape(toString(value), opts.escape), false
	}
}

func formatFloat(value interface{}, floatFormat string) string {
	val, ok := toFloat(value)
	if !ok {
		return toString(value)
	}
	if floatFormat != "" {
		return fmt.Sprintf(floatFormat, val)
	}
	return strconv.FormatFloat(val, 'g', -1, 64)
}

func formatBool(value interface{}) string {
	switch v := value.(type) {
	case bool:
		if v {
			return "TRUE"
		}
		return "FALSE"
	case int:
		if v != 0 {
			return "TRUE"
		}
		return "FALSE"
	default:
		return toString(value)
	}
}

func formatError(value interface{}) string {
	switch v := value.(type) {
	case byte:
		if text, ok := xlrd.ErrorTextFromCode[v]; ok {
			return text
		}
	case int:
		if text, ok := xlrd.ErrorTextFromCode[byte(v)]; ok {
			return text
		}
	}
	return "#ERROR"
}

func toString(value interface{}) string {
	if value == nil {
		return ""
	}
	switch v := value.(type) {
	case string:
		return v
	default:
		return fmt.Sprint(value)
	}
}

func toFloat(value interface{}) (float64, bool) {
	switch v := value.(type) {
	case float64:
		return v, true
	case float32:
		return float64(v), true
	case int:
		return float64(v), true
	case int64:
		return float64(v), true
	default:
		return 0, false
	}
}

func maybeEscape(value string, enabled bool) string {
	if !enabled || value == "" {
		return value
	}
	replacer := strings.NewReplacer("\r", "\\r", "\n", "\\n", "\t", "\\t")
	return replacer.Replace(value)
}

func (cw *csvWriter) writeRow(fields []field) error {
	var buf bytes.Buffer
	for i, field := range fields {
		if i > 0 {
			buf.WriteRune(cw.delimiter)
		}
		buf.WriteString(cw.formatField(field))
	}
	buf.WriteString(cw.lineTerminator)
	_, err := cw.w.Write(buf.Bytes())
	return err
}

func (cw *csvWriter) formatField(f field) string {
	quote := cw.needsQuote(f)
	if !quote {
		return f.text
	}
	escaped := strings.ReplaceAll(f.text, `"`, `""`)
	return `"` + escaped + `"`
}

func (cw *csvWriter) needsQuote(f field) bool {
	switch cw.quoting {
	case quotingAll:
		return true
	case quotingNonNumeric:
		return !f.isNumeric
	case quotingMinimal:
		return strings.ContainsRune(f.text, cw.delimiter) || strings.ContainsAny(f.text, "\"\r\n")
	case quotingNone:
		return false
	default:
		return false
	}
}

func isDateCell(book *xlrd.Book, xfIndex int) bool {
	if xfIndex < 0 || xfIndex >= len(book.XFList) {
		return false
	}
	formatKey := book.XFList[xfIndex].FormatKey
	if isBuiltinDateFormat(formatKey) {
		return true
	}
	if book.FormatMap == nil {
		return false
	}
	format := book.FormatMap[formatKey]
	if format == nil || format.FormatString == "" {
		return false
	}
	return xlrd.IsDateFormatString(book, format.FormatString)
}

func isBuiltinDateFormat(key int) bool {
	switch key {
	case 14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 50, 57, 58:
		return true
	default:
		return false
	}
}

func formatDate(value float64, datemode int, dateFormat string) (string, bool) {
	if math.IsNaN(value) || math.IsInf(value, 0) {
		return "", false
	}
	t, err := xlrd.XldateAsDatetime(value, datemode)
	if err != nil {
		return "", false
	}
	if dateFormat != "" {
		return strftime(t, dateFormat), true
	}
	if value < 1 {
		return t.Format("15:04:05"), true
	}
	if value-math.Floor(value) != 0 {
		return t.Format("2006-01-02 15:04:05"), true
	}
	return t.Format("2006-01-02"), true
}

func strftime(t time.Time, format string) string {
	var b strings.Builder
	for i := 0; i < len(format); i++ {
		if format[i] != '%' || i+1 >= len(format) {
			b.WriteByte(format[i])
			continue
		}
		i++
		switch format[i] {
		case '%':
			b.WriteByte('%')
		case 'Y':
			b.WriteString(t.Format("2006"))
		case 'y':
			b.WriteString(t.Format("06"))
		case 'm':
			b.WriteString(t.Format("01"))
		case 'd':
			b.WriteString(t.Format("02"))
		case 'H':
			b.WriteString(t.Format("15"))
		case 'M':
			b.WriteString(t.Format("04"))
		case 'S':
			b.WriteString(t.Format("05"))
		case 'b':
			b.WriteString(t.Format("Jan"))
		case 'B':
			b.WriteString(t.Format("January"))
		case 'a':
			b.WriteString(t.Format("Mon"))
		case 'A':
			b.WriteString(t.Format("Monday"))
		default:
			b.WriteByte('%')
			b.WriteByte(format[i])
		}
	}
	return b.String()
}

func changeExt(name, ext string) string {
	return strings.TrimSuffix(name, filepath.Ext(name)) + ext
}

func sanitizeFilename(name string) string {
	invalid := strings.NewReplacer(string(os.PathSeparator), "_", "/", "_", "\\", "_")
	clean := strings.TrimSpace(invalid.Replace(name))
	if clean == "" {
		return "sheet"
	}
	return clean
}

func utf8DecodeRune(value string) (rune, int) {
	r, size := utf8.DecodeRuneInString(value)
	if r == utf8.RuneError && size == 1 {
		return rune(value[0]), 1
	}
	return r, size
}
