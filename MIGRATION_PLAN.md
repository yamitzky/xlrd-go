# xlrd Python to Go Migration Plan

## Overview
This document outlines the plan to migrate the xlrd Python library to pure Go, maintaining all original interfaces, functionality, and test cases.

## Repository Information
- Target repository: `github.com/yamitzky/xlrd-go`
- Package name: `xlrd` (or `xlrdgo` if `xlrd` conflicts)

## Migration Principles

### DO
- Pure Go implementation (no Python embedding)
- Go-style naming conventions (PascalCase for exported, camelCase for unexported)
- Go-style package structure
- All code must compile
- Migrate all test cases
- Migrate all documentation
- Update README with Go usage examples
- Code comments and documentation in English
- Commits in English

### DON'T
- Embed Python code
- Keep Pythonic implementations
- Remove original Python interfaces or parameters
- Remove original test cases
- Use Japanese code comments

## Package Structure

```
xlrd-go/
├── go.mod
├── go.sum
├── README.md (updated)
├── LICENSE
├── cmd/
│   └── xlrd/          # CLI tool (equivalent to scripts/runxlrd.py)
├── xlrd/              # Main package
│   ├── book.go        # Book, Name, open_workbook_xls
│   ├── sheet.go       # Sheet, Cell, Colinfo, Rowinfo
│   ├── biffh.go       # BIFF constants, errors, utilities
│   ├── compdoc.go     # Compound document handling
│   ├── formatting.go  # Format, Font, XF, etc.
│   ├── formula.go     # Formula parsing and evaluation
│   ├── xldate.go      # Date/time conversion
│   └── inspect.go     # File format inspection
├── xlrd_test/         # Test package
│   ├── book_test.go
│   ├── sheet_test.go
│   ├── biffh_test.go
│   ├── compdoc_test.go
│   ├── formatting_test.go
│   ├── formula_test.go
│   ├── xldate_test.go
│   ├── inspect_test.go
│   ├── open_workbook_test.go
│   ├── cell_test.go
│   ├── formats_test.go
│   ├── formulas_test.go
│   ├── ignore_workbook_corruption_error_test.go
│   ├── missing_records_test.go
│   ├── workbook_test.go
│   ├── xldate_to_datetime_test.go
│   └── helpers.go
├── testdata/          # Test samples (from tests/samples/)
└── docs/              # Hugo documentation (migrated from docs/)

```

## Python to Go Mapping

### Core Types

| Python | Go |
|--------|-----|
| `class Book` | `type Book struct` |
| `class Sheet` | `type Sheet struct` |
| `class Cell` | `type Cell struct` |
| `class Name` | `type Name struct` |
| `class XLRDError` | `type XLRDError struct` (implements `error`) |
| `XL_CELL_*` constants | `const XL_CELL_*` |
| `open_workbook()` | `OpenWorkbook()` |
| `xldate_as_datetime()` | `XldateAsDatetime()` |
| `xldate_as_tuple()` | `XldateAsTuple()` |

### Naming Conventions

- Python snake_case → Go PascalCase (exported) or camelCase (unexported)
- Python `rowx`, `colx` → Go `rowX`, `colX` (or keep as-is for API compatibility)
- Python `nrows`, `ncols` → Go `NRows`, `NCols`
- Python `sheet_by_index()` → Go `SheetByIndex()`
- Python `cell_value()` → Go `CellValue()`

### Error Handling

- Python exceptions → Go `error` return values
- Custom exceptions → Custom error types implementing `error`

### File I/O

- Python `open()` → Go `os.Open()` or `os.ReadFile()`
- Python `mmap` → Go `mmap` package or `[]byte` slices
- Python `file_contents` parameter → Go `io.Reader` or `[]byte`

### Data Structures

- Python `dict` → Go `map`
- Python `list` → Go `slice`
- Python `tuple` → Go struct or multiple return values
- Python `array.array` → Go `[]byte` or appropriate slice type

## Implementation Phases

### Phase 1: Empty Go Implementation (Compiles)
- [x] Create `go.mod`
- [x] Create basic package structure
- [x] Define all interfaces/types (empty structs)
- [x] Define all function signatures (empty implementations)
- [x] Ensure code compiles
- **Commit**: "Initial Go implementation structure"

### Phase 2: Interface Definitions (Compiles)
- [x] Define all types with proper fields
- [x] Define all function signatures with proper parameters
- [x] Define all constants
- [x] Define all error types
- [x] Ensure code compiles
- **Commit**: "Define all interfaces and types"

### Phase 3: Test File Migration (One by One)

Each test file migration includes:
1. Convert test file to Go test
2. Implement minimal functionality to pass tests
3. Run tests to verify
4. Commit

#### Test Files to Migrate:
1. `test_xldate.py` → `xldate_test.go`
2. `test_xldate_to_datetime.py` → `xldate_to_datetime_test.go`
3. `test_biffh.py` → `biffh_test.go`
4. `test_cell.py` → `cell_test.go`
5. `test_formats.py` → `formats_test.go`
6. `test_formulas.py` → `formulas_test.go`
7. `test_inspect.py` → `inspect_test.go`
8. `test_open_workbook.py` → `open_workbook_test.go`
9. `test_sheet.py` → `sheet_test.go`
10. `test_workbook.py` → `workbook_test.go`
11. `test_ignore_workbook_corruption_error.py` → `ignore_workbook_corruption_error_test.go`
12. `test_missing_records.py` → `missing_records_test.go`

**Commit pattern**: "Implement [feature] and migrate [test_file]"

### Phase 4: Documentation Setup
- [ ] Set up Hugo for documentation
- [ ] Configure Hugo theme and structure
- **Commit**: "Set up Hugo documentation"

### Phase 5: Documentation Migration
- [ ] Migrate all `.rst` files to Hugo-compatible Markdown
- [ ] Update API documentation
- [ ] Update usage examples for Go
- **Commit**: "Migrate documentation to Hugo"

## Key Implementation Details

### Date/Time Handling
- Excel date system (1900 vs 1904)
- Julian day number calculations
- Time-only values (0.0 <= xldate < 1.0)
- Leap year bug handling (1900)

### BIFF Format Support
- BIFF versions: 2.0, 2.1, 3, 4S, 4W, 5, 7, 8, 8X
- Record parsing
- Cell type handling (EMPTY, TEXT, NUMBER, DATE, BOOLEAN, ERROR, BLANK)

### Compound Document (OLE2)
- OLE2 file structure parsing
- Sector allocation table (SAT)
- Directory tree traversal
- Stream extraction

### Formatting
- Font handling
- Number formats
- Cell formatting (XF records)
- Color palette

### Formulas
- Formula parsing (not evaluation - results only)
- Shared formulas
- Name references
- 3D references

## Testing Strategy

- Use Go's `testing` package
- Maintain all original test cases
- Use `testdata/` directory for test files
- Use table-driven tests where appropriate
- Ensure test coverage matches or exceeds Python version

## Dependencies

Potential Go dependencies:
- `github.com/stretchr/testify` - Testing utilities (optional)
- Standard library only preferred

## Notes

- Keep Python files during migration for reference
- Remove Python files only after full migration is complete
- Maintain git history where possible
- Update CI/CD configuration for Go
