---
title: "Overview"
description: "What xlrd-go does and what it does not support."
---

`xlrd-go` is a pure Go port of the classic `xlrd` library for reading Excel
`.xls` files (BIFF2-8). It focuses on extracting cell values and formatting
information from legacy spreadsheets.

## Not supported

The following features are ignored safely and will not be extracted:

- Charts, macros, pictures, and embedded objects (including embedded worksheets)
- VBA modules
- Formula evaluation beyond returning cached results
- Comments and hyperlinks
- Autofilters, advanced filters, pivot tables, conditional formatting, and data validation

Password-protected files are not supported.

## Quick start

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

## CLI: xls2csv

`xls2csv` converts `.xls` files into CSV. The interface is flavored after
https://github.com/dilshod/xlsx2csv.

```bash
xls2csv -s 1 input.xls output.csv
xls2csv -a input.xls
xls2csv -s 0 input.xls outdir
```

Notes:

- Only `.xls` files are supported.
- `--outputencoding` currently supports `utf-8` only.
- `--hyperlinks` is parsed but not supported yet.
