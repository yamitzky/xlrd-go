---
title: "On-Demand Loading"
description: "Loading worksheets on demand to reduce memory usage."
---

This feature is governed by the `OpenWorkbookOptions.OnDemand` option and
allows saving memory and time by loading only those sheets that the caller is
interested in, and releasing sheets when no longer required.

`OnDemand = false` (default):
- `OpenWorkbook` loads global data and all sheets, releases resources no longer
  required, and returns.

`OnDemand = true` and BIFF version < 5.0:
- A warning is emitted, `OnDemand` is recorded as `false`, and the old process
  is followed.

`OnDemand = true` and BIFF version >= 5.0:
- `OpenWorkbook` loads global data and returns without releasing resources.
  At this stage, the only information available about sheets is `Book.NSheets`
  and `Book.SheetNames`.

`Book.SheetByName` and `Book.SheetByIndex` return a sheet if it is already
loaded. When on-demand loading is enabled, use `Book.Sheets` to load all
unloaded sheets when needed.

`Book.Sheets` will load all unloaded sheets.

The caller may save memory by calling `Book.UnloadSheet` when finished with a
sheet. This applies irrespective of the state of `OnDemand`.

The caller may re-load an unloaded sheet by calling `Book.SheetByName` or
`Book.SheetByIndex`, except if the required resources have been released (which
will have happened automatically when `OnDemand` is false). This is the only
case where an error will be returned.

The caller may query the state of a sheet using `Book.SheetLoaded`.

`Book.ReleaseResources` may be used to save memory and close any memory-mapped
file before proceeding to examine already-loaded sheets. Once resources are
released, no further sheets can be loaded.

When using on-demand loading, ensure that `Book.ReleaseResources` is always
called, even if an error is raised in your code. This is especially important
if the input file has been memory-mapped.
