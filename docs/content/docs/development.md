---
title: "Development"
description: "Contributing and working on xlrd-go."
---

If you wish to contribute, fork the repository and open a pull request.

## Running tests

From the repository root:

```bash
go test ./...

## Formatting

```bash
gofmt -w .
```
```

## Building documentation

The documentation site is built with Hugo. From the repository root:

```bash
hugo -s docs
```

## Making a release

Update the changelog and tag a release. Release automation will be added in CI.
