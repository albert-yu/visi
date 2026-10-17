# visi

> [!NOTE]
> While use of LLMs for generating _code_ is encouraged, LLM-generated _prose_
> is severely restricted. See the [LLM policy](#llm-policy) for more details.

[![CI](https://github.com/albert-yu/visi/actions/workflows/ci.yml/badge.svg)](https://github.com/albert-yu/visi/actions/workflows/ci.yml)

## Overview

`visi` is a CLI application for editing and evaluating Excel spreadsheets.

### Goals

- Match Excel's execution behavior 100%, including pivot tables and macros (VBA).
  - In some cases, `visi` even produces more numerically accurate results
  - Excludes some functionality that require Microsoft web services
  - `visi` aims to match the classes of errors that Excel produces,
    but not the error message itself (`visi` may be able to improve upon Excel here)
  - [Documented discrepancies](./docs/excel-discrepancies.md)
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

## Usage

`visi` provides a deep API to read/write XLSX files.
In addition, it can evaluate (`eval`) and typecheck
(`check`) spreadsheets.

### Recalculate Formulas

You can take a spreadsheet authored with `openpyxl` and
perform the computation with `eval`.

```bash
# --in-place or -i will write back to the input file
visi eval data.xlsx --in-place

# Recalculate and print to stdout,
# leaving the original workbook unmodified
visi eval data.xlsx --print --format table
```

### Typecheck Formulas

```bash
visi set example.xlsx \
  --cell A1 --value '=SUM("not numeric")' \
  --output example.xlsx \
  --quiet

visi check example.xlsx
```

```text
Sheet1!A1: error[type]: SUM expects a number, got text
```

---

## Development

This monorepo is structured follows:

- **[`visi-core`](visi-core/)**: embeddable spreadsheet engine that
  parses and executes the formulas in the workbook
- **[`visi`](visi/)**: Command-line application using `visi-core`
  which can edit and execute Excel files headlessly
- [`fuzz`](fuzz/): differential fuzzing harness to match Excel behavior
- [`visi-python`](visi-python/): Python bindings for `visi-core` (API unstable)

`visi` aims for feature parity with Excel by using
a harness (`fuzz`) that drives a real copy of Excel via AppleScript or COM automation,
runs computations, and compares the results. Both the cell values and types
should match exactly.

## LLM Policy

- LLMs may be used for source code generation
- No LLM-generated comments in source code
- No LLM use for writing prose, unless clearly attributed at the _beginning_ of
  the content
- EXCEPTION: commit messages may be 100% LLM-authored
- AGENTS.md should automatically enforce this

## License

Dual-licensed:

- [Apache License, Version 2.0](LICENSE-APACHE)
- [MIT License](LICENSE-MIT)
