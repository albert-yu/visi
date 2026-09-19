# visi monorepo

> [!NOTE]
> While ~99.9999% of the code is LLM-authored, the _prose_
> should not be. See the [LLM policy](#llm-policy).

[![CI](https://github.com/albert-yu/visi/actions/workflows/ci.yml/badge.svg)](https://github.com/albert-yu/visi/actions/workflows/ci.yml)

## Overview

`visi` is a CLI application for editing and evaluating Excel spreadsheets.

### Goals

- Match Excel's execution behavior 100%, including pivot tables and macros (VBA).
  - In some cases, `visi` even produces more numerically accurate results
  - Excludes some functionality that require Microsoft web services
  - `visi` aims to match the classes of errors that Excel produces, but not the error message itself (`visi` may be able to improve upon Excel here)
- Prioritize performance
  - Written in Rust to maximize potential performance capacity
  - Fast startup time
  - Handle large workflows

### Why `visi`

LLMs are good at authoring spreadsheets with existing tools
(such as `openpyxl` for Python), but in order
to _evaluate_ the formula results, you need a real
spreadsheet application. Excel is
unsuitable for headless evaluation of
spreadsheet files, especially on non-Windows
platforms (e.g. no COM automation).

`visi` aims to bridge this gap by providing an
execution/verification layer for spreadsheets.

## Installation

### Homebrew (macOS/Linux)

```bash
brew install albert-yu/tap/visi
```

### Examples

#### Inspect Workbook Structure

```bash
# Display summary of sheets, dimensions, and formula counts
visi info data.xlsx

# Output summary as JSON
visi info data.xlsx --json
```

#### Read Sheet Contents, Ranges, or Cells

```bash
# View first sheet as a formatted ASCII table in the terminal
visi read data.xlsx

# View specific sheet and range
visi read data.xlsx --sheet Sheet1 --range A1:C10

# Read a single cell result or raw formula
visi read data.xlsx --cell A1
visi read data.xlsx --cell A3 --raw

visi read data.xlsx --format csv
visi read data.xlsx --format json
```

#### Recalculate Formulas

You can take a spreadsheet authored with `openpyxl` and
perform the computation with `eval`.

```bash
# --in-place or -i will write back to the input file
visi eval data.xlsx --in-place

# Recalculate and print to stdout,
# leaving the original workbook unmodified
visi eval data.xlsx --print --format table
```

#### Update Cells & Set Formulas

```bash
visi set data.xlsx --sheet Sheet1 --cell A1 --value 100 --output updated.xlsx

# Set multiple cell values and formulas at once
visi set data.xlsx -s Sheet1 -S A1=100 -S A2=200 -S A3="=A1+A2" -S A4="=AVERAGE(A1:A2)" -i

# Cross-sheet reference
visi set data.xlsx -s Sheet2 -S B1="=Sheet1!A3 + 50" -i
```

#### Manage Worksheets

```bash
visi sheet list data.xlsx
visi sheet add data.xlsx --name "Summary" -i
visi sheet rename data.xlsx --old "Sheet1" --new "Data" -i
visi sheet delete data.xlsx --name "OldSheet" -i
```

#### Manipulate Rows and Columns

Use `--index` for a 0-based offset or `--label`
if you want to follow the UI labels.

```bash
# Insert a new row at start
visi row delete data.xlsx --sheet Sheet1 --label 1 -i
visi row insert data.xlsx --sheet Sheet1 --index 0 -i

# Delete column label 'C'
visi col delete data.xlsx --sheet Sheet1 --label C -i
visi col delete data.xlsx --sheet Sheet1 --index 2 -i
```

#### Export Sheet Data

```bash
visi export data.xlsx --sheet Sheet1 --format csv --output sheet1.csv
visi export data.xlsx --sheet Sheet1 --format json --output sheet1.json
```

---

## Development

This monorepo is structured follows:

- **[`visi-core`](visi-core/)**: embeddedable spreadsheet engine that parses and executes the formulas in the workbook 
- **[`visi`](visi/)**: Command-line application using `visi-core` which can edit and execute Excel files headlessly

`visi` aims for feature parity with Excel by using
a harness that drives a real copy of Excel via AppleScript or COM automation,
runs computations, and compares the results. Both the cell values and types
should match exactly. See [`fuzz`](./fuzz/README.md) for more details.

## LLM Policy

- LLMs may be used for source code generation
- No LLM use for writing prose, unless clearly attributed at the _beginning_ of
  the content
  - Source code comments count as prose
  - AGENTS.md should automatically enforce this

## License

Dual-licensed:

- [Apache License, Version 2.0](LICENSE-APACHE)
- [MIT License](LICENSE-MIT)
