xlrd-go
=======

xlrd-go is a pure Go port of the classic xlrd library for reading legacy
``.xls`` files (BIFF2-8). It focuses on extracting cell values and formatting
information from historical Excel spreadsheets.

This library only supports ``.xls`` files. For newer formats such as ``.xlsx``,
use a different library.

Not supported (ignored safely):

* Charts, macros, pictures, and embedded objects (including embedded worksheets)
* VBA modules
* Formula evaluation beyond returning cached results
* Comments and hyperlinks
* Autofilters, advanced filters, pivot tables, conditional formatting, data validation

Password-protected files are not supported.

Quick start:

.. code-block:: bash

    go get github.com/yamitzky/xlrd-go

.. code-block:: go

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

Documentation lives under ``docs/`` and is built with Hugo.
