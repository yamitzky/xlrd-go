---
title: "Formatting"
description: "Formatting information available in XLS files."
---

## Introduction

This collection of features is intended to provide the information needed to:

- display or render spreadsheet contents (for example, on a screen or in a PDF)
- copy spreadsheet data to another file without losing the ability to render it

## The palette and color indexes

A color is represented in Excel as an `(red, green, blue)` ("RGB") tuple with
values in `range(256)`. Excel files do not provide an unlimited number of
colors. Each spreadsheet is limited to a palette of 64 colors (24 in Excel 3.0
and 4.0, 8 in Excel 2.0). Colors are referenced by an index into this palette.

Color indexes 0 to 7 represent eight fixed built-in colors: black, white, red,
green, blue, yellow, magenta, and cyan.

The remaining colors in the palette (8 to 63 in Excel 5.0 and later) can be
changed by the user. In the Excel 2003 UI, Tools -> Options -> Color presents a
palette of 7 rows of 8 colors. The last two rows are reserved for charts.

The correspondence between that grid and the assigned color indexes is not
left-to-right top-to-bottom.

Indexes 8 to 15 correspond to changeable parallels of the 8 fixed colors.
For example, index 7 is always cyan; index 15 starts as cyan but can be
changed by the user.

The default color for each index depends on the file version; tables of the
defaults are available in the source code. If the user changes one or more
colors, a `PALETTE` record appears in the XLS file and gives RGB values for
all changeable indexes.

Colors can also be used in number formats: `[CYAN]....` and `[COLOR8]....`
refer to color index 7; `[COLOR16]....` will produce cyan unless the user
changes color index 15 to something else.

In addition, there are several "magic" color indexes used by Excel:

- `0x18` (BIFF3-BIFF4), `0x40` (BIFF5-BIFF8): System window text color for
  border lines (used in `XF`, `CF`, and `WINDOW2` records)
- `0x19` (BIFF3-BIFF4), `0x41` (BIFF5-BIFF8): System window background color for
  pattern background (used in `XF` and `CF` records)
- `0x43`: System face color (dialog background color)
- `0x4D`: System window text color for chart border lines
- `0x4E`: System window background color for chart areas
- `0x4F`: Automatic color for chart border lines (seems to be always Black)
- `0x50`: System ToolTip background color (used in note objects)
- `0x51`: System ToolTip text color (used in note objects)
- `0x7FFF`: System window text color for fonts (used in `FONT` and `CF` records)

`0x7FFF` appears to be the default color index and appears often in `FONT`
records.

## Default formatting

Default formatting is applied to all empty cells (those not described by a cell
record):

- Row default information (`ROW` record, `Rowinfo` type) if available
- Otherwise, column default information (`COLINFO` record, `Colinfo` type)
- As a last resort, the worksheet/workbook default cell format (the `XF` record
  with fixed index 15, which itself defaults to the first `XF` record)

## Formatting features not included in xlrd

- Asian phonetic text ("ruby"), used for Japanese furigana
- Conditional formatting (see OOo docs: CONDFMT/CF records)
- Miscellaneous sheet-level and book-level items, such as printing layout or
  screen panes
- Modern Excel file versions do not keep most of the built-in number formats in
  the file; Excel loads formats according to locale. xlrd's emulation is
  limited to a hard-wired table for US English, so currency symbols, date order,
  separators, and decimal marks may be inappropriate.

This does not affect users who are copying XLS files, only those who are
visually rendering cells.
