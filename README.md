# xlrd-go

xlrd-go is a pure Go port of the classic xlrd library for reading legacy `.xls`
files (BIFF2-8). It focuses on extracting cell values and formatting
information from historical Excel spreadsheets.

This library only supports `.xls` files. For newer formats such as `.xlsx`,
use a different library.

Not supported (ignored safely):

- Charts, macros, pictures, and embedded objects (including embedded worksheets)
- VBA modules
- Formula evaluation beyond returning cached results
- Comments and hyperlinks
- Autofilters, advanced filters, pivot tables, conditional formatting, data validation

Password-protected files are not supported.

## Quick start

```bash
go get github.com/yamitzky/xlrd-go
```

```go
package main

import (
    "fmt"

    "github.com/yamitzky/xlrd-go/xlrd"
)

func main() {
    book, err := xlrd.OpenWorkbook("myfile.xls", nil)
    if err != nil {
        panic(err)
    }

    fmt.Printf("Worksheets: %d\n", book.NSheets)
    fmt.Printf("Sheet names: %v\n", book.SheetNames())

    sheet, err := book.SheetByIndex(0)
    if err != nil {
        panic(err)
    }

    fmt.Printf("%s %d %d\n", sheet.Name, sheet.NRows, sheet.NCols)
    fmt.Printf("Cell D30: %v\n", sheet.CellValue(29, 3))
}
```

Documentation lives under `docs/` and is built with Hugo.

## CLI: xls2csv

This repo also ships a CLI named `xls2csv` for converting legacy `.xls` files
to CSV. The interface is flavored after
https://github.com/dilshod/xlsx2csv.

```bash
go install github.com/yamitzky/xlrd-go/cmd/xls2csv@latest
```

```bash
xls2csv -s 1 input.xls output.csv
xls2csv -a input.xls
xls2csv -s 0 input.xls outdir
```

Notes:

- Only `.xls` files are supported.
- `--outputencoding` currently supports `utf-8` only.
- `--hyperlinks` is parsed but not supported yet.
