---
title: "CLI"
description: "xls2csv command line usage."
---

`xls2csv` converts legacy `.xls` files into CSV. The CLI is flavored after
https://github.com/dilshod/xlsx2csv.

## Install

```bash
go install github.com/yamitzky/xlrd-go/cmd/xls2csv@latest
```

## Usage

```text
xls2csv [-h] [-v] [-a] [-c OUTPUTENCODING] [-s SHEETID]
        [-n SHEETNAME] [-d DELIMITER] [-l LINETERMINATOR]
        [-f DATEFORMAT] [--floatformat FLOATFORMAT]
        [-i] [-e] [-p SHEETDELIMITER]
        [--hyperlinks]
        [-I INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...]]
        [-E EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...]] [-m]
        [--ignore-workbook-corruption]
        xlsfile [outfile]
```

Positional arguments:

- `xlsfile`: `.xls` file path, use `-` to read from STDIN
- `outfile`: output CSV file path, or directory if `-s 0` is specified

Options:

- `-h, --help`: show help and exit
- `-v, --version`: show version and exit
- `-a, --all`: export all sheets
- `-c, --outputencoding`: output CSV encoding (default: `utf-8`)
- `-s, --sheet`: sheet number to convert, `0` for all
- `-n, --sheetname`: sheet name to convert
- `-d, --delimiter`: CSV delimiter; `tab` or `x09` for tab (default: `,`)
- `-l, --lineterminator`: line terminator; `\n`, `\r\n`, or `\r` (default: OS)
- `-f, --dateformat`: override date/time format (ex. `%Y/%m/%d`)
- `--floatformat`: override float format (ex. `%.15f`)
- `-i, --ignoreempty`: skip empty lines
- `-e, --escape`: escape `\r\n\t` characters
- `-p, --sheetdelimiter`: delimiter between sheets (default: `--------`)
- `-q, --quoting`: quoting mode: `none`, `minimal`, `nonnumeric`, `all`
- `--hyperlinks`: include hyperlinks (not supported yet)
- `-I, --include_sheet_pattern`: include sheet names matching patterns (when `-a`)
- `-E, --exclude_sheet_pattern`: exclude sheet names matching patterns (when `-a`)
- `-m, --merge-cells`: expand merged cells to their top-left value
- `--ignore-workbook-corruption`: ignore workbook corruption checks

## Examples

```bash
xls2csv -s 1 input.xls output.csv
xls2csv -a input.xls
xls2csv -s 0 input.xls outdir
xls2csv -d tab -i input.xls
xls2csv -f "%Y/%m/%d" input.xls
```

## Directory mode

If `xlsfile` is a directory, each `.xls` file is converted to a `.csv` file in
the output directory. If `outfile` is omitted, the output is written alongside
the input files.

## Notes

- Only `.xls` files are supported.
- `--outputencoding` currently supports `utf-8` only.
- `--hyperlinks` is parsed but not supported yet.
