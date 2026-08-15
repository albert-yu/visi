# visi

CLI tool for authoring _and_ executing Excel workbooks. The execution
layer should match Excel's behavior 100%.

Libraries such as `openpyxl` can author Excel workbooks,
but they cannot evaluate formulas. `visi` enables headless shell-scripting
and LLM automation of Excel workflows.


## Installation

### Homebrew (macOS/Linux)

```bash
brew install albert-yu/tap/visi
```

### Building from Source

#### Requirements

- [Rust](https://www.rust-lang.org/) (2024 edition supported)

#### Build Binary

```bash
# Build dev binary
cargo build --workspace

# Build optimized release binary
cargo build --release --workspace
```
The compiled CLI executable will be located at `target/release/visi` (or `target/debug/visi`).

### Examples

#### 1. Inspect Workbook Structure

```bash
# Display summary of sheets, dimensions, and formula counts
visi info data.xlsx

# Output summary as JSON
visi info data.xlsx --json
```

#### 2. Read Sheet Contents, Ranges, or Cells

```bash
# View first sheet as a formatted ASCII table in the terminal
visi read data.xlsx

# View specific sheet and range
visi read data.xlsx --sheet Sheet1 --range A1:C10

# Read a single cell result or raw formula
visi read data.xlsx --cell A1
visi read data.xlsx --cell A3 --raw

# Output as CSV, TSV, or JSON
visi read data.xlsx --format csv
visi read data.xlsx --format json
```

#### 3. Update Cells & Set Formulas

```bash
# Set cell values and save to output file
visi set data.xlsx --sheet Sheet1 --cell A1 --value 100 --output updated.xlsx

# Set multiple cell values and formulas at once
visi set data.xlsx -s Sheet1 -S A1=100 -S A2=200 -S A3="=A1+A2" -S A4="=AVERAGE(A1:A2)" --in-place

# Cross-sheet reference
visi set data.xlsx -s Sheet2 -S B1="=Sheet1!A3 + 50" -i
```

#### 4. Recalculate Formulas

```bash
# Force recalculation of all formulas across all sheets and save in-place
visi eval data.xlsx --in-place

# Recalculate and print evaluated grid to stdout
visi eval data.xlsx --print --format table
```

#### 5. Manage Worksheets

```bash
# List sheets
visi sheet list data.xlsx

# Add a new worksheet
visi sheet add data.xlsx --name "Summary" -i

# Rename a worksheet
visi sheet rename data.xlsx --old "Sheet1" --new "Data" -i

# Delete a worksheet
visi sheet delete data.xlsx --name "OldSheet" -i
```

#### 6. Manipulate Rows and Columns

```bash
# Insert a new row at row 2 (1-based index)
visi row insert data.xlsx --sheet Sheet1 --index 2 -i

# Delete row 5
visi row delete data.xlsx --sheet Sheet1 --index 5 -i

# Insert a column at column 'B'
visi col insert data.xlsx --sheet Sheet1 --index B -i

# Delete column 'C'
visi col delete data.xlsx --sheet Sheet1 --index C -i
```

#### 7. Export Sheet Data

```bash
# Export sheet to CSV or JSON file
visi export data.xlsx --sheet Sheet1 --format csv --output sheet1.csv
visi export data.xlsx --sheet Sheet1 --format json --output sheet1.json
```
