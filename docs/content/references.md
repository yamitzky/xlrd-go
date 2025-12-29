---
title: "Named References"
description: "Names, constants, formulas, and macros in legacy XLS files."
---

A name is used to refer to a cell, a group of cells, a constant value, a
formula, or a macro. Usually the scope of a name is global across the whole
workbook. However it can be local to a worksheet. For example, if the sales
figures are in different cells in different sheets, the user may define the
name "Sales" in each sheet. There are built-in names, like "Print_Area" and
"Print_Titles"; these two are naturally local to a sheet.

To inspect the names with a user interface like Excel or LibreOffice Calc,
use Insert -> Name -> Define. This will show the global names plus those local
to the currently selected sheet.

A `Book` provides two maps (`Book.NameMap` and `Book.NameAndScopeMap`) and a
list (`Book.NameObjList`) which allow various ways of accessing `Name` objects.
There is one `Name` object for each `NAME` record found in the workbook.
`Name` objects have many attributes, several of which are relevant only when
`Name.Macro` is 1.

There is a convenience method `Name.Cell()` that is intended to extract the
value when the name refers to a single cell. In the Go port, this requires
formula evaluation support and is not fully implemented yet.

Note: Name information is not extracted from files older than Excel 5.0
(`Book.BiffVersion < 50`).
