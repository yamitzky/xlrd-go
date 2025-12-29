---
title: "Excel Dates"
description: "How Excel stores dates and how xlrd-go interprets them."
---

In reality, there are no such things as Excel date types. You have floating
point numbers and hope. There are several problems with Excel dates:

1. Dates are not stored as a separate data type; they are stored as floating
   point numbers and you have to rely on:

   - the number format applied to them in Excel
   - knowing which cells are supposed to have dates in them

   xlrd helps with the former by inspecting the format applied to each number
   cell. If it appears to be a date format, the cell is classified as a date
   rather than a number.

2. Excel for Windows stores dates by default as the number of days (or fraction
   thereof) since `1899-12-31T00:00:00`. Excel for Macintosh uses a default
   start date of `1904-01-01T00:00:00`.

   The date system can be changed in Excel on a per-workbook basis. This is a
   bad idea if there are already dates in the workbook.

   Which date system is in use is recorded in the workbook. When converting
   numbers from a workbook, you must use the `Book.Datemode` value from the
   workbook that the numbers came from. If you guess, you run the risk of being
   1462 days out of kilter.

   Reference:
   https://support.microsoft.com/en-us/help/180162/xl-the-1900-date-system-vs.-the-1904-date-system

3. The Excel implementation of the Windows-default 1900-based date system
   works on the incorrect premise that 1900 was a leap year. It interprets the
   number 60 as meaning `1900-02-29`, which is not a valid date.

   Consequently, any number less than 61 is ambiguous. For example, is 59 the
   result of `1900-02-28` entered directly, or is it `1900-03-01` minus 2 days?

   The OpenOffice.org Calc program "corrects" the Microsoft problem; entering
   `1900-02-27` causes the number 59 to be stored. Save as an XLS file, then
   open the file with Excel and you'll see `1900-02-28` displayed.

   Reference:
   https://support.microsoft.com/en-us/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year

4. The Macintosh-default 1904-based date system counts `1904-01-02` as day 1
   and `1904-01-01` as day zero. Thus any number such that
   `(0.0 <= number < 1.0)` is ambiguous. Is 0.625 a time of day (`15:00:00`),
   independent of the calendar, or should it be interpreted as an instant on a
   particular day (`1904-01-01T15:00:00`)?

   The functions in `xlrd/xldate` take the view that such a number is a
   calendar-independent time of day (like Go's `time.Time`/`time.Duration` in
   context) for both date systems. This is consistent with more recent
   Microsoft documentation.

5. Usage of the Excel `DATE()` function may leave strange dates in a
   spreadsheet. Quoting the help file in respect of the 1900 date system:

   If year is between 0 (zero) and 1899 (inclusive), Excel adds that value to
   1900 to calculate the year. For example, `DATE(108,1,2)` returns January 2,
   2008 (1900+108).

   This gimmick means that `DATE(1899, 12, 31)` is interpreted as
   `3799-12-31`.

For conversion helpers, see `XldateAsTuple` and `XldateAsDatetime` in the
`xlrd` package.
