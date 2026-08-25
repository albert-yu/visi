use usage::{Args, Cli as UsageCli, Subcommands, ValueEnum};

#[derive(UsageCli, Debug)]
#[usage(
    bin = "visi",
    version,
    about = "Read, evaluate formulas, and update Excel (.xlsx) spreadsheets",
    long_about = "visi is a developer-friendly command line utility to inspect, evaluate formulas in, and update Excel (.xlsx) files powered by the visi-core calculation engine.\n\nExamples:\n  visi info data.xlsx\n  visi read data.xlsx --sheet Sheet1 --format table\n  visi set data.xlsx --sheet Sheet1 --cell A1 --value 100 --in-place\n  visi eval data.xlsx --output calculated.xlsx"
)]
pub struct Cli {
    #[usage(subcommand)]
    pub command: Commands,

    /// Enable verbose logging output to stderr
    #[usage(short, long, global)]
    pub verbose: bool,

    /// Suppress non-essential informational messages to stderr
    #[usage(short, long, global)]
    pub quiet: bool,
}

#[derive(Subcommands, Debug)]
pub enum Commands {
    /// Display structure summary and metadata of an Excel workbook
    Info(InfoArgs),

    /// Read and display contents of a sheet, cell range, or single cell
    #[usage(alias = "view")]
    Read(ReadArgs),

    /// Update cell values or formulas in an Excel workbook
    #[usage(alias = "update")]
    Set(SetArgs),

    /// Recalculate all formulas across all sheets in the workbook
    #[usage(alias = "recalc")]
    Eval(EvalArgs),

    /// Manage worksheets in the workbook (list, add, delete, rename)
    Sheet(SheetArgs),

    /// Perform row operations (insert, delete)
    Row(RowArgs),

    /// Perform column operations (insert, delete)
    Col(ColArgs),

    /// Manage embedded charts (list, add, delete)
    Chart(ChartArgs),

    /// Manage Excel Tables (list, add, delete, rename, resize, rename-column)
    Table(TableArgs),

    /// Manage pivot tables (create, list, delete, refresh, field CRUD, filters)
    Pivot(PivotArgs),

    /// Manage VBA macro modules (list, add, remove, rename, set-source)
    Macro(MacroArgs),

    /// Manage cell styles (color, fill background, bold, italic, font) and table themes
    Style(StyleArgs),

    /// Export a worksheet to CSV, TSV, or JSON format
    Export(ExportArgs),
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum OutputFormat {
    /// ASCII table layout for human terminal view
    Table,
    /// Comma-Separated Values format
    Csv,
    /// Tab-Separated Values format
    Tsv,
    /// JSON format
    Json,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExportFormat {
    /// Comma-Separated Values format
    Csv,
    /// Tab-Separated Values format
    Tsv,
    /// JSON format
    Json,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChartTypeArg {
    Line,
    Bar,
    Column,
    Pie,
    Scatter,
    Area,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum PivotAreaArg {
    Row,
    Column,
    Value,
    Filter,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum PivotAggArg {
    Sum,
    Count,
    CountNumbers,
    Average,
    Max,
    Min,
}

#[derive(Args, Debug)]
pub struct InfoArgs {
    /// Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// Format summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct ReadArgs {
    /// Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// Target sheet name (defaults to first sheet)
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// Target cell range in A1 notation (e.g. A1:D10)
    #[usage(short, long)]
    pub range: Option<String>,

    /// Target single cell in A1 notation (e.g. A1)
    #[usage(short, long)]
    pub cell: Option<String>,

    /// Output display format [table, csv, tsv, json]
    #[usage(short, long, value_enum, default = "table")]
    pub format: OutputFormat,

    /// Recalculate and execute formulas before displaying (enabled by default)
    #[usage(long, default = "true", negate = "--no-eval")]
    pub eval: bool,

    /// Output raw cell formulas instead of calculated result values
    #[usage(long)]
    pub raw: bool,

    /// Treat the first row of selected cells as header names
    #[usage(long)]
    pub headers: bool,
}

#[derive(Args, Debug)]
pub struct SetArgs {
    /// Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// Target sheet name (defaults to first sheet)
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// Target cell coordinate in A1 notation (e.g. A1)
    #[usage(short, long)]
    pub cell: Vec<String>,

    /// Value or formula to set (e.g. 100, "Hello", "=SUM(A1:A10)")
    #[usage(long)]
    pub value: Vec<String>,

    /// Set cell assignments in CELL=VALUE format (e.g. -S A1=100 -S B1="=A1*2")
    #[usage(short = 'S', long = "set", value_name = "CELL=VALUE")]
    pub set_pairs: Vec<String>,

    /// Write updated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// Save updated workbook in-place, overwriting the input file
    #[usage(short = 'i', long)]
    pub in_place: bool,

    /// Recalculate formulas after setting values (enabled by default)
    #[usage(long, default = "true", negate = "--no-eval")]
    pub eval: bool,

    /// Text/font color in Hex format (e.g. "#FF0000") or color name ("red", "blue")
    #[usage(long = "font-color", alias = "color")]
    pub font_color: Option<String>,

    /// Background fill color in Hex format (e.g. "#00FF00") or color name ("green", "yellow")
    #[usage(long = "bg-color", alias = "bg")]
    pub bg_color: Option<String>,

    /// Enable bold text style
    #[usage(long)]
    pub bold: bool,

    /// Enable italic text style
    #[usage(long)]
    pub italic: bool,

    /// Enable underline text style
    #[usage(long)]
    pub underline: bool,

    /// Font family name (e.g. "Arial", "Calibri", "Courier New")
    #[usage(long = "font-family")]
    pub font_family: Option<String>,

    /// Font size in points (e.g. 11, 12, 14)
    #[usage(long = "font-size")]
    pub font_size: Option<f64>,
}

#[derive(Args, Debug)]
pub struct EvalArgs {
    /// Input Excel file path
    pub file: String,

    /// Write calculated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// Overwrite input file with calculated formula results
    #[usage(short = 'i', long)]
    pub in_place: bool,

    /// Print calculated sheet content(s) to stdout
    #[usage(short, long)]
    pub print: bool,

    /// Specific sheet to print if --print is specified
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// Output display format if --print is specified [table, csv, tsv, json]
    #[usage(short, long, value_enum, default = "table")]
    pub format: OutputFormat,
}

#[derive(Args, Debug)]
pub struct SheetArgs {
    #[usage(subcommand)]
    pub command: SheetSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum SheetSubcommands {
    /// List all worksheets in the workbook
    List(SheetListArgs),
    /// Add a new empty worksheet
    Add(SheetAddArgs),
    /// Delete a worksheet
    Delete(SheetDeleteArgs),
    /// Rename an existing worksheet
    Rename(SheetRenameArgs),
}

#[derive(Args, Debug)]
pub struct SheetListArgs {
    /// Input Excel file path
    pub file: String,
    /// Format output as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct SheetAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Name for the new worksheet
    #[usage(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct SheetDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the worksheet to delete
    #[usage(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct SheetRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current worksheet name
    #[usage(long)]
    pub old: String,
    /// New worksheet name
    #[usage(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct RowArgs {
    #[usage(subcommand)]
    pub command: RowSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum RowSubcommands {
    /// Insert a new row at specified 1-based index
    Insert(RowOpArgs),
    /// Delete a row at specified 1-based index
    Delete(RowOpArgs),
}

#[derive(Args, Debug)]
pub struct RowOpArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// 1-based row index (e.g. 1, 5)
    #[usage(short = 'x', long)]
    pub index: usize,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ColArgs {
    #[usage(subcommand)]
    pub command: ColSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum ColSubcommands {
    /// Insert a new column at specified index or letter
    Insert(ColOpArgs),
    /// Delete a column at specified index or letter
    Delete(ColOpArgs),
}

#[derive(Args, Debug)]
pub struct ColOpArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// Column letter or 1-based index (e.g. "B" or "2")
    #[usage(short = 'x', long)]
    pub index: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartArgs {
    #[usage(subcommand)]
    pub command: ChartSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum ChartSubcommands {
    /// List all charts in the workbook
    List(ChartListArgs),
    /// Add a new chart to a worksheet
    Add(ChartAddArgs),
    /// Edit an existing chart's properties
    Edit(ChartEditArgs),
    /// Delete a chart by ID
    Delete(ChartDeleteArgs),
}

#[derive(Args, Debug)]
pub struct ChartListArgs {
    /// Input Excel file path
    pub file: String,
    /// Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct ChartAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// Chart type [line, bar, column, pie, scatter, area]
    #[usage(short, long, value_enum)]
    pub chart_type: ChartTypeArg,
    /// Data range reference (e.g. Sheet1!A1:B10)
    #[usage(short, long)]
    pub range: String,
    /// Optional chart title
    #[usage(short, long)]
    pub title: Option<String>,
    /// Cell where the chart's top-left corner is anchored (e.g. D5); defaults to A1
    #[usage(long)]
    pub anchor: Option<String>,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartEditArgs {
    /// Input Excel file path
    pub file: String,
    /// Chart ID to edit
    #[usage(long)]
    pub id: u64,
    /// New display name for the chart
    #[usage(long)]
    pub name: Option<String>,
    /// New chart type [line, bar, column, pie, scatter, area]
    #[usage(long, value_enum)]
    pub chart_type: Option<ChartTypeArg>,
    /// New data range reference (e.g. Sheet1!A1:B10)
    #[usage(long)]
    pub range: Option<String>,
    /// Set the chart title
    #[usage(long, conflicts = "--clear-title")]
    pub title: Option<String>,
    /// Remove the chart title
    #[usage(long)]
    pub clear_title: bool,
    /// Set the X-axis label
    #[usage(long, conflicts = "--clear-xlabel")]
    pub xlabel: Option<String>,
    /// Remove the X-axis label
    #[usage(long)]
    pub clear_xlabel: bool,
    /// Set the Y-axis label
    #[usage(long, conflicts = "--clear-ylabel")]
    pub ylabel: Option<String>,
    /// Remove the Y-axis label
    #[usage(long)]
    pub clear_ylabel: bool,
    /// Show the chart legend
    #[usage(long, conflicts = "--hide-legend")]
    pub show_legend: bool,
    /// Hide the chart legend
    #[usage(long)]
    pub hide_legend: bool,
    /// Move the chart: cell for its new top-left anchor (e.g. D5)
    #[usage(long)]
    pub anchor: Option<String>,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Chart ID to delete
    #[usage(long)]
    pub id: u64,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableArgs {
    #[usage(subcommand)]
    pub command: TableSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum TableSubcommands {
    /// List all Excel Tables in the workbook
    List(TableListArgs),
    /// Define a new Excel Table over an existing cell range
    Add(TableAddArgs),
    /// Delete an Excel Table (leaves its cell contents untouched)
    Delete(TableDeleteArgs),
    /// Rename an Excel Table
    Rename(TableRenameArgs),
    /// Resize an Excel Table by moving its bottom-right corner
    Resize(TableResizeArgs),
    /// Rename one column of an Excel Table
    RenameColumn(TableRenameColumnArgs),
    /// Modify visual style theme of an Excel Table
    Style(StyleTableArgs),
}

#[derive(Args, Debug)]
pub struct TableListArgs {
    /// Input Excel file path
    pub file: String,
    /// Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct TableAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name (defaults to first sheet, or the range's own
    /// sheet prefix if given, e.g. Sheet1!A1:D10)
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// Name for the new table
    #[usage(short, long)]
    pub name: String,
    /// Cell range the table should occupy (e.g. A1:D10)
    #[usage(short, long)]
    pub range: String,
    /// Treat the range's first row as plain data, not column headers
    #[usage(long)]
    pub no_header_row: bool,
    /// Reserve the range's last row as a totals row
    #[usage(long)]
    pub totals_row: bool,
    /// Visual style theme name (e.g. "TableStyleMedium9", "TableStyleLight1")
    #[usage(long)]
    pub style: Option<String>,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the table to delete
    #[usage(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current table name
    #[usage(long)]
    pub old: String,
    /// New table name
    #[usage(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableResizeArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the table to resize
    #[usage(short, long)]
    pub name: String,
    /// New cell range for the table; its top-left corner must match the
    /// table's current top-left corner (e.g. A1:E12)
    #[usage(short, long)]
    pub range: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableRenameColumnArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the table containing the column
    #[usage(short, long)]
    pub name: String,
    /// Existing column name or 1-based column index within the table
    #[usage(short, long)]
    pub column: String,
    /// New column name
    #[usage(long)]
    pub new_name: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotArgs {
    #[usage(subcommand)]
    pub command: PivotSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum PivotSubcommands {
    /// List all pivot tables in the workbook
    List(PivotListArgs),
    /// Create a new pivot table from an Excel Table or a plain cell range
    Create(PivotCreateArgs),
    /// Delete a pivot table (leaves its source data untouched)
    Delete(PivotDeleteArgs),
    /// Rename a pivot table
    Rename(PivotRenameArgs),
    /// Recompute a pivot table's output from its current source data
    Refresh(PivotRefreshArgs),
    /// Add a field to a pivot table's Row/Column/Value/Filter area
    AddField(PivotAddFieldArgs),
    /// Remove a field from a pivot table's Row/Column/Value/Filter area
    RemoveField(PivotRemoveFieldArgs),
    /// Restrict or clear a filter field's allowed values
    Filter(PivotFilterArgs),
}

#[derive(Args, Debug)]
pub struct PivotListArgs {
    /// Input Excel file path
    pub file: String,
    /// Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct PivotCreateArgs {
    /// Input Excel file path
    pub file: String,
    /// Name for the new pivot table
    #[usage(short, long)]
    pub name: String,
    /// Name of an existing Excel Table to use as the source (mutually
    /// exclusive with --source-range)
    #[usage(long)]
    pub source_table: Option<String>,
    /// Cell range to use as the source, first row treated as headers (e.g.
    /// A1:D100; mutually exclusive with --source-table)
    #[usage(long)]
    pub source_range: Option<String>,
    /// Worksheet the source range lives on (defaults to first sheet, or the
    /// range's own sheet prefix, e.g. Sheet1!A1:D10)
    #[usage(long)]
    pub source_sheet: Option<String>,
    /// Top-left cell of the pivot table's output (e.g. A1)
    #[usage(long, default = "A1")]
    pub dest: String,
    /// Worksheet the pivot table's output is written to (defaults to first
    /// sheet)
    #[usage(long)]
    pub dest_sheet: Option<String>,
    /// Omit the grand-total row at the bottom of the output (shown by
    /// default, matching Excel)
    #[usage(long)]
    pub no_grand_totals_row: bool,
    /// Omit the grand-total column at the right of the output (shown by
    /// default, matching Excel)
    #[usage(long)]
    pub no_grand_totals_col: bool,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to delete
    #[usage(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current pivot table name
    #[usage(long)]
    pub old: String,
    /// New pivot table name
    #[usage(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRefreshArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to refresh (omit with --all to refresh every
    /// pivot table in the workbook)
    #[usage(short, long)]
    pub name: Option<String>,
    /// Refresh every pivot table in the workbook
    #[usage(long)]
    pub all: bool,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotAddFieldArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to modify
    #[usage(short, long)]
    pub name: String,
    /// Area to add the field to [row, column, value, filter]
    #[usage(short, long, value_enum)]
    pub area: PivotAreaArg,
    /// Source column name
    #[usage(short, long)]
    pub column: String,
    /// Aggregation function, only used when --area value [sum, count,
    /// count-numbers, average, max, min] (defaults to sum)
    #[usage(long, value_enum)]
    pub agg: Option<PivotAggArg>,
    /// Custom display label for a value field (defaults to e.g. "Sum of
    /// Amount")
    #[usage(long)]
    pub label: Option<String>,
    /// Disable the subtotal row Excel normally shows for this field when
    /// it isn't the innermost field in its Row/Column area (only used when
    /// --area row or --area column)
    #[usage(long)]
    pub no_subtotal: bool,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRemoveFieldArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to modify
    #[usage(short, long)]
    pub name: String,
    /// Area to remove the field from [row, column, value, filter]
    #[usage(short, long, value_enum)]
    pub area: PivotAreaArg,
    /// Source column name
    #[usage(short, long)]
    pub column: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotFilterArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to modify
    #[usage(short, long)]
    pub name: String,
    /// Filter field's source column name
    #[usage(short, long)]
    pub column: String,
    /// Comma-separated list of the only values to include (e.g.
    /// "East,West", or "-7,3,12" for a negative-number-valued column)
    #[usage(long, delimiter = ',', allow_hyphen_values)]
    pub values: Vec<String>,
    /// Remove the filter, allowing every value again
    #[usage(long)]
    pub clear: bool,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroArgs {
    #[usage(subcommand)]
    pub command: MacroSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum MacroSubcommands {
    /// List all VBA modules in the workbook
    List(MacroListArgs),
    /// Add a new VBA module
    Add(MacroAddArgs),
    /// Remove a VBA module
    Remove(MacroRemoveArgs),
    /// Rename a VBA module
    Rename(MacroRenameArgs),
    /// Replace a VBA module's source code
    SetSource(MacroSetSourceArgs),
    /// Check VBA module source for syntax errors
    Check(MacroCheckArgs),
    /// Run a VBA procedure (opt-in; see --help for what is and isn't supported)
    Run(MacroRunArgs),
}

/// Running a macro executes code the workbook's author wrote, so it is never
/// implicit: no other subcommand runs one, and `eval` in particular does not.
///
/// Given a workbook the macro runs *against* it and can read and write cells,
/// so it needs `--output` or `--in-place` like any other write command -- a
/// macro that changes the workbook with neither is an error rather than a
/// silent discard. Given a `.bas` file there is no workbook, and anything
/// reaching for one reports so rather than doing nothing quietly.
///
/// Only part of Excel's object model is implemented (`Range`, `Cells`,
/// `Worksheets`, `WorksheetFunction`, and the properties in the Phase 2 list
/// of `docs/vba-macro-support.md`). Everything else -- styles, tables,
/// pivots, `CreateObject`, `MsgBox`, file and network I/O -- raises a
/// run-time error naming the construct.
#[derive(Args, Debug)]
pub struct MacroRunArgs {
    /// Input Excel file path, or a .bas source file, or - for stdin
    pub file: String,
    /// Name of the procedure to run
    #[usage(short, long)]
    pub name: String,
    /// Module to take the procedure from (defaults to searching all modules)
    #[usage(short, long)]
    pub module: Option<String>,
    /// Argument to pass, repeatable and positional in order
    #[usage(short = 'a', long = "arg")]
    pub args: Vec<String>,
    /// Where to write the workbook the macro changed (must end in .xlsm)
    #[usage(short, long)]
    pub output: Option<String>,
    /// Write the changed workbook back over the input file
    #[usage(short, long, conflicts = "--output")]
    pub in_place: bool,
    /// Output the result as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct MacroCheckArgs {
    /// Input Excel file path, or a .bas source file, or - for stdin
    pub file: String,
    /// Check only this module (defaults to every module in the workbook)
    #[usage(short, long)]
    pub name: Option<String>,
    /// Treat the input as part of a larger project
    ///
    /// A name used with call syntax that resolves nowhere is accepted
    /// rather than reported, since a module not supplied here -- a sibling
    /// of a loose .bas file, or a referenced project -- may declare it.
    /// Everything the source's own text disproves is still reported.
    #[usage(long)]
    pub partial: bool,
    /// Output results as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum VbaModuleKindArg {
    /// A plain module with no host object binding
    Standard,
    /// A class module (not fully validated against real Excel -- see the
    /// VBA feature's known limitations)
    Class,
    /// A document module (e.g. a worksheet's code-behind); requires
    /// --sheet
    Document,
}

#[derive(Args, Debug)]
pub struct MacroListArgs {
    /// Input Excel file path
    pub file: String,
    /// Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct MacroAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Name for the new module
    #[usage(short, long)]
    pub name: String,
    /// Module kind
    #[usage(short, long, value_enum, default = "standard")]
    pub kind: VbaModuleKindArg,
    /// Sheet this document module belongs to (required for --kind document)
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// Module source code, given inline
    #[usage(long, conflicts = "--source-file")]
    pub source: Option<String>,
    /// Module source code, read from a file
    #[usage(long)]
    pub source_file: Option<String>,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroRemoveArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the module to remove
    #[usage(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current module name
    #[usage(long)]
    pub old: String,
    /// New module name
    #[usage(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroSetSourceArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the module to update
    #[usage(short, long)]
    pub name: String,
    /// New module source code, given inline
    #[usage(long, conflicts = "--source-file")]
    pub source: Option<String>,
    /// New module source code, read from a file
    #[usage(long)]
    pub source_file: Option<String>,
    /// Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ExportArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name (defaults to first sheet)
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// Export format [csv, tsv, json]
    #[usage(short, long, value_enum, default = "csv")]
    pub format: ExportFormat,
    /// Output file path (if omitted, writes to stdout)
    #[usage(short, long)]
    pub output: Option<String>,
    /// Recalculate formulas before exporting (enabled by default)
    #[usage(long, default = "true", negate = "--no-eval")]
    pub eval: bool,
}

#[derive(Args, Debug)]
pub struct StyleArgs {
    #[usage(subcommand)]
    pub command: StyleSubcommands,
}

#[derive(Subcommands, Debug)]
pub enum StyleSubcommands {
    /// Modify cell font color, background color, font styles, and family/size
    Cell(StyleCellArgs),
    /// Modify visual style theme of an Excel Table
    Table(StyleTableArgs),
}

#[derive(Args, Debug)]
pub struct StyleCellArgs {
    /// Input Excel file path
    pub file: String,

    /// Target sheet name (defaults to first sheet or sheet prefix in cell/range)
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// Target cell coordinate in A1 notation (e.g. A1)
    #[usage(short, long)]
    pub cell: Option<String>,

    /// Target range in A1 notation (e.g. A1:C10)
    #[usage(short, long)]
    pub range: Option<String>,

    /// Text/font color in Hex format (e.g. "#FF0000") or color name ("red", "blue")
    #[usage(long = "font-color", alias = "color")]
    pub font_color: Option<String>,

    /// Background fill color in Hex format (e.g. "#00FF00") or color name ("green", "yellow")
    #[usage(long = "bg-color", alias = "bg")]
    pub bg_color: Option<String>,

    /// Enable bold text style
    #[usage(long)]
    pub bold: bool,

    /// Enable italic text style
    #[usage(long)]
    pub italic: bool,

    /// Enable underline text style
    #[usage(long)]
    pub underline: bool,

    /// Font family name (e.g. "Arial", "Calibri", "Courier New")
    #[usage(long = "font-family")]
    pub font_family: Option<String>,

    /// Font size in points (e.g. 11, 12, 14)
    #[usage(long = "font-size")]
    pub font_size: Option<f64>,

    /// Write updated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct StyleTableArgs {
    /// Input Excel file path
    pub file: String,

    /// Name of the target Excel Table
    #[usage(short, long)]
    pub name: String,

    /// Visual style theme name (e.g. "TableStyleMedium9", "TableStyleLight1", "TableStyleDark11")
    #[usage(short, long)]
    pub style: String,

    /// Write updated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}
