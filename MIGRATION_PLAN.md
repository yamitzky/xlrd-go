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
1. `test_xldate.py` → `xldate_test.go` ✅
2. `test_xldate_to_datetime.py` → `xldate_to_datetime_test.go` ✅
3. `test_biffh.py` → `biffh_test.go` ✅
4. `test_cell.py` → `cell_test.go` ✅
5. `test_formats.py` → `formats_test.go` ✅
6. `test_formulas.py` → `formulas_test.go` ✅
7. `test_inspect.py` → `inspect_test.go` ✅
8. `test_open_workbook.py` → `open_workbook_test.go` ✅
9. `test_sheet.py` → `sheet_test.go` ✅
10. `test_workbook.py` → `workbook_test.go` ✅
11. `test_ignore_workbook_corruption_error.py` → `ignore_workbook_corruption_error_test.go` ✅
12. `test_missing_records.py` → `missing_records_test.go` ✅

**Commit pattern**: "Implement [feature] and migrate [test_file]"

✅ **Phase 3 Complete**: All test files migrated and core functionality implemented

### Phase 4: Documentation Setup
- [ ] Set up Hugo for documentation
- [ ] Configure Hugo theme and structure
- **Commit**: "Set up Hugo documentation"

### Phase 5: Documentation Migration
- [ ] Migrate all `.rst` files to Hugo-compatible Markdown
- [ ] Update API documentation
- [ ] Update usage examples for Go
- **Commit**: "Migrate documentation to Hugo"

### Phase 6: CI/CD Migration to GitHub Actions
- [ ] Create `.github/workflows/` directory
- [ ] Create CI workflow (`.github/workflows/ci.yml`) with Go test matrix
- [ ] Create release workflow (`.github/workflows/release.yml`) with GoReleaser
- [ ] Create docs workflow (`.github/workflows/docs.yml`) for Hugo deployment
- [ ] Configure Codecov integration for Go coverage
- [ ] Update repository settings for GitHub Pages
- [ ] Test all workflows
- [ ] Remove CircleCI configuration (`.circleci/`)
- **Commit**: "Migrate CI/CD from CircleCI to GitHub Actions"

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

## CI/CD Configuration Migration

### Current CircleCI Setup (Python)
- **Test Matrix**: Python 2.7, 3.6, 3.9
- **Coverage**: Coverage.py with Codecov integration
- **Documentation**: Sphinx-based docs
- **Packaging**: setuptools-based wheel and source distribution
- **Release**: Carthorse-based automated releases

### Target GitHub Actions Setup (Go)
- **Test Matrix**: Go 1.19, 1.20, 1.21 (latest stable versions)
- **Coverage**: Go native coverage with Codecov integration
- **Documentation**: Hugo-based docs
- **Packaging**: Go modules with goreleaser
- **Release**: GoReleaser-based automated releases
- **Quality Checks**: gofmt, go vet, golint

#### GitHub Actions Workflows to Create:

1. **CI Workflow** (`.github/workflows/ci.yml`):
   ```yaml
   name: CI
   on: [push, pull_request]
   jobs:
     test:
       runs-on: ubuntu-latest
       strategy:
         matrix:
           go-version: ['1.19', '1.20', '1.21']
       steps:
         - uses: actions/checkout@v4
         - name: Set up Go
           uses: actions/setup-go@v4
           with:
             go-version: ${{ matrix.go-version }}
         - name: Run tests
           run: go test ./...
         - name: Run tests with coverage
           run: go test -coverprofile=coverage.out ./...
         - name: Upload coverage to Codecov
           uses: codecov/codecov-action@v3
           with:
             file: ./coverage.out

     lint:
       runs-on: ubuntu-latest
       steps:
         - uses: actions/checkout@v4
         - name: Set up Go
           uses: actions/setup-go@v4
           with:
             go-version: '1.21'
         - name: Run gofmt
           run: gofmt -d .
         - name: Run go vet
           run: go vet ./...
         - name: Run golint
           run: golint ./...
   ```

2. **Release Workflow** (`.github/workflows/release.yml`):
   ```yaml
   name: Release
   on:
     push:
       tags: ['v*']
   jobs:
     release:
       runs-on: ubuntu-latest
       steps:
         - uses: actions/checkout@v4
         - name: Set up Go
           uses: actions/setup-go@v4
           with:
             go-version: '1.21'
         - name: Run GoReleaser
           uses: goreleaser/goreleaser-action@v5
           with:
             version: latest
             args: release --clean
           env:
             GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
   ```

3. **Docs Workflow** (`.github/workflows/docs.yml`):
   ```yaml
   name: Docs
   on:
     push:
       branches: [main]
     pull_request:
       branches: [main]
   jobs:
     docs:
       runs-on: ubuntu-latest
       steps:
         - uses: actions/checkout@v4
         - name: Set up Hugo
           uses: peaceiris/actions-hugo@v2
           with:
             hugo-version: 'latest'
         - name: Build docs
           run: hugo --source docs/
         - name: Deploy to GitHub Pages
           if: github.ref == 'refs/heads/main'
           uses: peaceiris/actions-gh-pages@v3
           with:
             github_token: ${{ secrets.GITHUB_TOKEN }}
             publish_dir: ./docs/public
   ```

### Migration Steps for CI/CD:
1. Create `.github/workflows/` directory
2. Create CI workflow for testing and linting
3. Create release workflow with GoReleaser
4. Create docs workflow for Hugo deployment
5. Configure Codecov integration for Go coverage
6. Update repository settings for GitHub Pages
7. Remove CircleCI configuration after migration

## Notes

- Keep Python files during migration for reference
- Remove Python files only after full migration is complete
- Maintain git history where possible
- Migrate CI/CD from CircleCI to GitHub Actions for Go
