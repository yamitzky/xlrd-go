---
title: "Development"
description: "Contributing and working on xlrd-go."
---

If you wish to contribute, fork the repository and open a pull request.

## Running tests

From the repository root:

```bash
go test ./...
```

## Formatting

```bash
gofmt -w .
```

## Building documentation

The documentation site is built with Hugo. From the repository root:

```bash
hugo -s docs
```

## Making a release

Update the changelog and tag a release. Pushing a `v*` tag triggers GoReleaser
to publish the `xls2csv` binaries and release notes to GitHub Releases.
