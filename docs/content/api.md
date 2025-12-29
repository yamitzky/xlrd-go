---
title: "API Reference"
description: "Top-level types and functions in the Go package."
---

The Go API follows the structure of the original xlrd package, but uses Go
naming conventions. For the complete reference, run:

```bash
go doc github.com/yamitzky/xlrd-go/xlrd
```

## Package `xlrd`

Key functions:

- `OpenWorkbook(filename string, options *OpenWorkbookOptions) (*Book, error)`
- `OpenWorkbookXLS(filename string, options *OpenWorkbookOptions) (*Book, error)`
- `XldateAsTuple(xldate float64, datemode int) (year, month, day, hour, min, sec int, err error)`
- `XldateAsDatetime(xldate float64, datemode int) (time.Time, error)`

Key types:

- `Book`: workbook container, sheet access, name maps, and workbook metadata
- `Sheet`: worksheet data, rows, columns, and cell access
- `Cell`: value and type information
- `Name`: named references, formulas, and macros
- `Format`, `Font`, `XF`: formatting records
- `CompDoc`: OLE2/compound document parser

For detailed struct fields and methods, use `go doc` or browse the source.
