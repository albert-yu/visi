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

    /// [LLM-generated] Enable verbose logging output to stderr
    #[usage(short, long, global)]
    pub verbose: bool,

    /// [LLM-generated] Suppress non-essential informational messages to stderr
    #[usage(short, long, global)]
    pub quiet: bool,

    /// [LLM-generated] Regional locale for date and number parsing (e.g. en-US, en-GB, de-DE, fr-FR)
    #[usage(long, global)]
    pub locale: Option<String>,
}

#[derive(Subcommands, Debug)]
pub enum Commands {
    /// [LLM-generated] Display structure summary and metadata of an Excel workbook
    Info(InfoArgs),

    /// [LLM-generated] Read and display contents of a sheet, cell range, or single cell
    #[usage(alias = "view")]
    Read(ReadArgs),

    /// [LLM-generated] Update cell values or formulas in an Excel workbook
    #[usage(alias = "update")]
    Set(SetArgs),

    /// [LLM-generated] Recalculate all formulas across all sheets in the workbook
    #[usage(alias = "recalc")]
    Eval(EvalArgs),

    /// [LLM-generated] Manage worksheets in the workbook (list, add, delete, rename)
    Sheet(SheetArgs),

    /// [LLM-generated] Perform row operations (insert, delete)
    Row(RowArgs),

    /// [LLM-generated] Perform column operations (insert, delete)
    Col(ColArgs),

    /// [LLM-generated] Manage embedded charts (list, add, delete)
    Chart(ChartArgs),

    /// [LLM-generated] Manage Excel Tables (list, add, delete, rename, resize, rename-column)
    Table(TableArgs),

    /// [LLM-generated] Manage pivot tables (create, list, delete, refresh, field CRUD, filters)
    Pivot(PivotArgs),

    /// [LLM-generated] Manage VBA macro modules (list, add, remove, rename, set-source)
    Macro(MacroArgs),

    /// [LLM-generated] Manage cell styles (color, fill background, bold, italic, font) and table themes
    Style(StyleArgs),

    /// [LLM-generated] Export a worksheet to CSV, TSV, or JSON format
    Export(ExportArgs),
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum OutputFormat {
    /// [LLM-generated] ASCII table layout for human terminal view
    Table,
    /// [LLM-generated] Comma-Separated Values format
    Csv,
    /// [LLM-generated] Tab-Separated Values format
    Tsv,
    /// [LLM-generated] JSON format
    Json,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExportFormat {
    /// [LLM-generated] Comma-Separated Values format
    Csv,
    /// [LLM-generated] Tab-Separated Values format
    Tsv,
    /// [LLM-generated] JSON format
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

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellTypeArg {
    Empty,
    Int,
    Float,
    String,
    Bool,
    DateTime,
    DateTimeIso,
    DurationIso,
    Error,
}

#[derive(Args, Debug)]
pub struct InfoArgs {
    /// [LLM-generated] Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// [LLM-generated] Format summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct ReadArgs {
    /// [LLM-generated] Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// [LLM-generated] Target sheet name (defaults to first sheet)
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// [LLM-generated] Target cell range in A1 notation (e.g. A1:D10)
    #[usage(short, long)]
    pub range: Option<String>,

    /// [LLM-generated] Target single cell in A1 notation (e.g. A1)
    #[usage(short, long)]
    pub cell: Option<String>,

    /// [LLM-generated] Output display format [table, csv, tsv, json]
    #[usage(short, long, value_enum, default = "table")]
    pub format: OutputFormat,

    /// [LLM-generated] Recalculate and execute formulas before displaying (enabled by default)
    #[usage(long, default = "true", negate = "--no-eval")]
    pub eval: bool,

    /// [LLM-generated] Output raw cell formulas instead of calculated result values
    #[usage(long)]
    pub raw: bool,

    /// [LLM-generated] Treat the first row of selected cells as header names
    #[usage(long)]
    pub headers: bool,
}

#[derive(Args, Debug)]
pub struct SetArgs {
    /// [LLM-generated] Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// [LLM-generated] Target sheet name (defaults to first sheet)
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// [LLM-generated] Target cell coordinate in A1 notation (e.g. A1)
    #[usage(short, long)]
    pub cell: Vec<String>,

    /// [LLM-generated] Value or formula to set (e.g. 100, "Hello", "=SUM(A1:A10)")
    #[usage(long)]
    pub value: Vec<String>,

    /// [LLM-generated] Set cell assignments in CELL=VALUE format (e.g. -S A1=100 -S B1="=A1*2")
    #[usage(short = 'S', long = "set", value_name = "CELL=VALUE")]
    pub set_pairs: Vec<String>,

    /// [LLM-generated] Explicit cell type [empty, int, float, string, bool, date-time, date-time-iso, duration-iso, error]
    #[usage(short = 't', long = "type", alias = "cell-type", value_enum)]
    pub cell_type: Option<CellTypeArg>,

    /// [LLM-generated] Write updated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// [LLM-generated] Save updated workbook in-place, overwriting the input file
    #[usage(short = 'i', long)]
    pub in_place: bool,

    /// [LLM-generated] Recalculate formulas after setting values (enabled by default)
    #[usage(long, default = "true", negate = "--no-eval")]
    pub eval: bool,

    /// [LLM-generated] Text/font color in Hex format (e.g. "#FF0000") or color name ("red", "blue")
    #[usage(long = "font-color", alias = "color")]
    pub font_color: Option<String>,

    /// [LLM-generated] Background fill color in Hex format (e.g. "#00FF00") or color name ("green", "yellow")
    #[usage(long = "bg-color", alias = "bg")]
    pub bg_color: Option<String>,

    /// [LLM-generated] Enable bold text style
    #[usage(long)]
    pub bold: bool,

    /// [LLM-generated] Enable italic text style
    #[usage(long)]
    pub italic: bool,

    /// [LLM-generated] Enable underline text style
    #[usage(long)]
    pub underline: bool,

    /// [LLM-generated] Font family name (e.g. "Arial", "Calibri", "Courier New")
    #[usage(long = "font-family")]
    pub font_family: Option<String>,

    /// [LLM-generated] Font size in points (e.g. 11, 12, 14)
    #[usage(long = "font-size")]
    pub font_size: Option<f64>,

    #[usage(long = "num-format", alias = "number-format")]
    pub num_format: Option<String>,
}

#[derive(Args, Debug)]
pub struct EvalArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,

    /// [LLM-generated] Write calculated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// [LLM-generated] Overwrite input file with calculated formula results
    #[usage(short = 'i', long)]
    pub in_place: bool,

    /// [LLM-generated] Print calculated sheet content(s) to stdout
    #[usage(short, long)]
    pub print: bool,

    /// [LLM-generated] Specific sheet to print if --print is specified
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// [LLM-generated] Output display format if --print is specified [table, csv, tsv, json]
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
    /// [LLM-generated] List all worksheets in the workbook
    List(SheetListArgs),
    /// [LLM-generated] Add a new empty worksheet
    Add(SheetAddArgs),
    /// [LLM-generated] Delete a worksheet
    Delete(SheetDeleteArgs),
    /// [LLM-generated] Rename an existing worksheet
    Rename(SheetRenameArgs),
}

#[derive(Args, Debug)]
pub struct SheetListArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Format output as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct SheetAddArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name for the new worksheet
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct SheetDeleteArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the worksheet to delete
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct SheetRenameArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Current worksheet name
    #[usage(long)]
    pub old: String,
    /// [LLM-generated] New worksheet name
    #[usage(long)]
    pub new: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
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
    /// [LLM-generated] Insert a new row at specified 1-based index
    Insert(RowOpArgs),
    /// [LLM-generated] Delete a row at specified 1-based index
    Delete(RowOpArgs),
}

#[derive(Args, Debug)]
pub struct RowOpArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Target worksheet name
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// [LLM-generated] 1-based row index (e.g. 1, 5)
    #[usage(short = 'x', long)]
    pub index: usize,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
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
    /// [LLM-generated] Insert a new column at specified index or letter
    Insert(ColOpArgs),
    /// [LLM-generated] Delete a column at specified index or letter
    Delete(ColOpArgs),
}

#[derive(Args, Debug)]
pub struct ColOpArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Target worksheet name
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// [LLM-generated] Column letter or 1-based index (e.g. "B" or "2")
    #[usage(short = 'x', long)]
    pub index: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
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
    /// [LLM-generated] List all charts in the workbook
    List(ChartListArgs),
    /// [LLM-generated] Add a new chart to a worksheet
    Add(ChartAddArgs),
    /// [LLM-generated] Edit an existing chart's properties
    Edit(ChartEditArgs),
    /// [LLM-generated] Delete a chart by ID
    Delete(ChartDeleteArgs),
}

#[derive(Args, Debug)]
pub struct ChartListArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct ChartAddArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Target worksheet name
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// [LLM-generated] Chart type [line, bar, column, pie, scatter, area]
    #[usage(short, long, value_enum)]
    pub chart_type: ChartTypeArg,
    /// [LLM-generated] Data range reference (e.g. Sheet1!A1:B10)
    #[usage(short, long)]
    pub range: String,
    /// [LLM-generated] Optional chart title
    #[usage(short, long)]
    pub title: Option<String>,
    /// [LLM-generated] Cell where the chart's top-left corner is anchored (e.g. D5); defaults to A1
    #[usage(long)]
    pub anchor: Option<String>,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartEditArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Chart ID to edit
    #[usage(long)]
    pub id: u64,
    /// [LLM-generated] New display name for the chart
    #[usage(long)]
    pub name: Option<String>,
    /// [LLM-generated] New chart type [line, bar, column, pie, scatter, area]
    #[usage(long, value_enum)]
    pub chart_type: Option<ChartTypeArg>,
    /// [LLM-generated] New data range reference (e.g. Sheet1!A1:B10)
    #[usage(long)]
    pub range: Option<String>,
    /// [LLM-generated] Set the chart title
    #[usage(long, conflicts = "--clear-title")]
    pub title: Option<String>,
    /// [LLM-generated] Remove the chart title
    #[usage(long)]
    pub clear_title: bool,
    /// [LLM-generated] Set the X-axis label
    #[usage(long, conflicts = "--clear-xlabel")]
    pub xlabel: Option<String>,
    /// [LLM-generated] Remove the X-axis label
    #[usage(long)]
    pub clear_xlabel: bool,
    /// [LLM-generated] Set the Y-axis label
    #[usage(long, conflicts = "--clear-ylabel")]
    pub ylabel: Option<String>,
    /// [LLM-generated] Remove the Y-axis label
    #[usage(long)]
    pub clear_ylabel: bool,
    /// [LLM-generated] Show the chart legend
    #[usage(long, conflicts = "--hide-legend")]
    pub show_legend: bool,
    /// [LLM-generated] Hide the chart legend
    #[usage(long)]
    pub hide_legend: bool,
    /// [LLM-generated] Move the chart: cell for its new top-left anchor (e.g. D5)
    #[usage(long)]
    pub anchor: Option<String>,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartDeleteArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Chart ID to delete
    #[usage(long)]
    pub id: u64,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
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
    /// [LLM-generated] List all Excel Tables in the workbook
    List(TableListArgs),
    /// [LLM-generated] Define a new Excel Table over an existing cell range
    Add(TableAddArgs),
    /// [LLM-generated] Delete an Excel Table (leaves its cell contents untouched)
    Delete(TableDeleteArgs),
    /// [LLM-generated] Rename an Excel Table
    Rename(TableRenameArgs),
    /// [LLM-generated] Resize an Excel Table by moving its bottom-right corner
    Resize(TableResizeArgs),
    /// [LLM-generated] Rename one column of an Excel Table
    RenameColumn(TableRenameColumnArgs),
    /// [LLM-generated] Modify visual style theme of an Excel Table
    Style(StyleTableArgs),
}

#[derive(Args, Debug)]
pub struct TableListArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct TableAddArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Target worksheet name (defaults to first sheet, or the range's own
    /// sheet prefix if given, e.g. Sheet1!A1:D10)
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// [LLM-generated] Name for the new table
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Cell range the table should occupy (e.g. A1:D10)
    #[usage(short, long)]
    pub range: String,
    /// [LLM-generated] Treat the range's first row as plain data, not column headers
    #[usage(long)]
    pub no_header_row: bool,
    /// [LLM-generated] Reserve the range's last row as a totals row
    #[usage(long)]
    pub totals_row: bool,
    /// [LLM-generated] Visual style theme name (e.g. "TableStyleMedium9", "TableStyleLight1")
    #[usage(long)]
    pub style: Option<String>,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableDeleteArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the table to delete
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableRenameArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Current table name
    #[usage(long)]
    pub old: String,
    /// [LLM-generated] New table name
    #[usage(long)]
    pub new: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableResizeArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the table to resize
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] New cell range for the table; its top-left corner must match the
    /// table's current top-left corner (e.g. A1:E12)
    #[usage(short, long)]
    pub range: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableRenameColumnArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the table containing the column
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Existing column name or 1-based column index within the table
    #[usage(short, long)]
    pub column: String,
    /// [LLM-generated] New column name
    #[usage(long)]
    pub new_name: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
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
    /// [LLM-generated] List all pivot tables in the workbook
    List(PivotListArgs),
    /// [LLM-generated] Create a new pivot table from an Excel Table or a plain cell range
    Create(PivotCreateArgs),
    /// [LLM-generated] Delete a pivot table (leaves its source data untouched)
    Delete(PivotDeleteArgs),
    /// [LLM-generated] Rename a pivot table
    Rename(PivotRenameArgs),
    /// [LLM-generated] Recompute a pivot table's output from its current source data
    Refresh(PivotRefreshArgs),
    /// [LLM-generated] Add a field to a pivot table's Row/Column/Value/Filter area
    AddField(PivotAddFieldArgs),
    /// [LLM-generated] Remove a field from a pivot table's Row/Column/Value/Filter area
    RemoveField(PivotRemoveFieldArgs),
    /// [LLM-generated] Restrict or clear a filter field's allowed values
    Filter(PivotFilterArgs),
}

#[derive(Args, Debug)]
pub struct PivotListArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct PivotCreateArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name for the new pivot table
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Name of an existing Excel Table to use as the source (mutually
    /// exclusive with --source-range)
    #[usage(long)]
    pub source_table: Option<String>,
    /// [LLM-generated] Cell range to use as the source, first row treated as headers (e.g.
    /// A1:D100; mutually exclusive with --source-table)
    #[usage(long)]
    pub source_range: Option<String>,
    /// [LLM-generated] Worksheet the source range lives on (defaults to first sheet, or the
    /// range's own sheet prefix, e.g. Sheet1!A1:D10)
    #[usage(long)]
    pub source_sheet: Option<String>,
    /// [LLM-generated] Top-left cell of the pivot table's output (e.g. A1)
    #[usage(long, default = "A1")]
    pub dest: String,
    /// [LLM-generated] Worksheet the pivot table's output is written to (defaults to first
    /// sheet)
    #[usage(long)]
    pub dest_sheet: Option<String>,
    /// [LLM-generated] Omit the grand-total row at the bottom of the output (shown by
    /// default, matching Excel)
    #[usage(long)]
    pub no_grand_totals_row: bool,
    /// [LLM-generated] Omit the grand-total column at the right of the output (shown by
    /// default, matching Excel)
    #[usage(long)]
    pub no_grand_totals_col: bool,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotDeleteArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the pivot table to delete
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRenameArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Current pivot table name
    #[usage(long)]
    pub old: String,
    /// [LLM-generated] New pivot table name
    #[usage(long)]
    pub new: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRefreshArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the pivot table to refresh (omit with --all to refresh every
    /// pivot table in the workbook)
    #[usage(short, long)]
    pub name: Option<String>,
    /// [LLM-generated] Refresh every pivot table in the workbook
    #[usage(long)]
    pub all: bool,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotAddFieldArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the pivot table to modify
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Area to add the field to [row, column, value, filter]
    #[usage(short, long, value_enum)]
    pub area: PivotAreaArg,
    /// [LLM-generated] Source column name
    #[usage(short, long)]
    pub column: String,
    /// [LLM-generated] Aggregation function, only used when --area value [sum, count,
    /// count-numbers, average, max, min] (defaults to sum)
    #[usage(long, value_enum)]
    pub agg: Option<PivotAggArg>,
    /// [LLM-generated] Custom display label for a value field (defaults to e.g. "Sum of
    /// Amount")
    #[usage(long)]
    pub label: Option<String>,
    /// [LLM-generated] Disable the subtotal row Excel normally shows for this field when
    /// it isn't the innermost field in its Row/Column area (only used when
    /// --area row or --area column)
    #[usage(long)]
    pub no_subtotal: bool,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRemoveFieldArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the pivot table to modify
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Area to remove the field from [row, column, value, filter]
    #[usage(short, long, value_enum)]
    pub area: PivotAreaArg,
    /// [LLM-generated] Source column name
    #[usage(short, long)]
    pub column: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotFilterArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the pivot table to modify
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Filter field's source column name
    #[usage(short, long)]
    pub column: String,
    /// [LLM-generated] Comma-separated list of the only values to include (e.g.
    /// "East,West", or "-7,3,12" for a negative-number-valued column)
    #[usage(long, delimiter = ',', allow_hyphen_values)]
    pub values: Vec<String>,
    /// [LLM-generated] Remove the filter, allowing every value again
    #[usage(long)]
    pub clear: bool,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
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
    /// [LLM-generated] List all VBA modules in the workbook
    List(MacroListArgs),
    /// [LLM-generated] Add a new VBA module
    Add(MacroAddArgs),
    /// [LLM-generated] Remove a VBA module
    Remove(MacroRemoveArgs),
    /// [LLM-generated] Rename a VBA module
    Rename(MacroRenameArgs),
    /// [LLM-generated] Replace a VBA module's source code
    SetSource(MacroSetSourceArgs),
    /// [LLM-generated] Check VBA module source for syntax errors
    Check(MacroCheckArgs),
    /// [LLM-generated] Run a VBA procedure (opt-in; see --help for what is and isn't supported)
    Run(MacroRunArgs),
}

/// [LLM-generated] Running a macro executes code the workbook's author wrote, so it is never
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
    /// [LLM-generated] Input Excel file path, or a .bas source file, or - for stdin
    pub file: String,
    /// [LLM-generated] Name of the procedure to run
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Module to take the procedure from (defaults to searching all modules)
    #[usage(short, long)]
    pub module: Option<String>,
    /// [LLM-generated] Argument to pass, repeatable and positional in order
    #[usage(short = 'a', long = "arg")]
    pub args: Vec<String>,
    /// [LLM-generated] Where to write the workbook the macro changed (must end in .xlsm)
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Write the changed workbook back over the input file
    #[usage(short, long, conflicts = "--output")]
    pub in_place: bool,
    /// [LLM-generated] Output the result as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct MacroCheckArgs {
    /// [LLM-generated] Input Excel file path, or a .bas source file, or - for stdin
    pub file: String,
    /// [LLM-generated] Check only this module (defaults to every module in the workbook)
    #[usage(short, long)]
    pub name: Option<String>,
    /// [LLM-generated] Treat the input as part of a larger project
    ///
    /// A name used with call syntax that resolves nowhere is accepted
    /// rather than reported, since a module not supplied here -- a sibling
    /// of a loose .bas file, or a referenced project -- may declare it.
    /// Everything the source's own text disproves is still reported.
    #[usage(long)]
    pub partial: bool,
    /// [LLM-generated] Output results as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(ValueEnum, Clone, Copy, Debug, PartialEq, Eq)]
pub enum VbaModuleKindArg {
    /// [LLM-generated] A plain module with no host object binding
    Standard,
    /// [LLM-generated] A class module (not fully validated against real Excel -- see the
    /// VBA feature's known limitations)
    Class,
    /// [LLM-generated] A document module (e.g. a worksheet's code-behind); requires
    /// --sheet
    Document,
}

#[derive(Args, Debug)]
pub struct MacroListArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Output summary as JSON
    #[usage(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct MacroAddArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name for the new module
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Module kind
    #[usage(short, long, value_enum, default = "standard")]
    pub kind: VbaModuleKindArg,
    /// [LLM-generated] Sheet this document module belongs to (required for --kind document)
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// [LLM-generated] Module source code, given inline
    #[usage(long, conflicts = "--source-file")]
    pub source: Option<String>,
    /// [LLM-generated] Module source code, read from a file
    #[usage(long)]
    pub source_file: Option<String>,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroRemoveArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the module to remove
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroRenameArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Current module name
    #[usage(long)]
    pub old: String,
    /// [LLM-generated] New module name
    #[usage(long)]
    pub new: String,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroSetSourceArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Name of the module to update
    #[usage(short, long)]
    pub name: String,
    /// [LLM-generated] New module source code, given inline
    #[usage(long, conflicts = "--source-file")]
    pub source: Option<String>,
    /// [LLM-generated] New module source code, read from a file
    #[usage(long)]
    pub source_file: Option<String>,
    /// [LLM-generated] Write updated workbook to target output file
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ExportArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,
    /// [LLM-generated] Target worksheet name (defaults to first sheet)
    #[usage(short, long)]
    pub sheet: Option<String>,
    /// [LLM-generated] Export format [csv, tsv, json]
    #[usage(short, long, value_enum, default = "csv")]
    pub format: ExportFormat,
    /// [LLM-generated] Output file path (if omitted, writes to stdout)
    #[usage(short, long)]
    pub output: Option<String>,
    /// [LLM-generated] Recalculate formulas before exporting (enabled by default)
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
    /// [LLM-generated] Modify cell font color, background color, font styles, and family/size
    Cell(StyleCellArgs),
    /// [LLM-generated] Modify visual style theme of an Excel Table
    Table(StyleTableArgs),
}

#[derive(Args, Debug)]
pub struct StyleCellArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,

    /// [LLM-generated] Target sheet name (defaults to first sheet or sheet prefix in cell/range)
    #[usage(short, long)]
    pub sheet: Option<String>,

    /// [LLM-generated] Target cell coordinate in A1 notation (e.g. A1)
    #[usage(short, long)]
    pub cell: Option<String>,

    /// [LLM-generated] Target range in A1 notation (e.g. A1:C10)
    #[usage(short, long)]
    pub range: Option<String>,

    /// [LLM-generated] Text/font color in Hex format (e.g. "#FF0000") or color name ("red", "blue")
    #[usage(long = "font-color", alias = "color")]
    pub font_color: Option<String>,

    /// [LLM-generated] Background fill color in Hex format (e.g. "#00FF00") or color name ("green", "yellow")
    #[usage(long = "bg-color", alias = "bg")]
    pub bg_color: Option<String>,

    /// [LLM-generated] Enable bold text style
    #[usage(long)]
    pub bold: bool,

    /// [LLM-generated] Enable italic text style
    #[usage(long)]
    pub italic: bool,

    /// [LLM-generated] Enable underline text style
    #[usage(long)]
    pub underline: bool,

    /// [LLM-generated] Font family name (e.g. "Arial", "Calibri", "Courier New")
    #[usage(long = "font-family")]
    pub font_family: Option<String>,

    /// [LLM-generated] Font size in points (e.g. 11, 12, 14)
    #[usage(long = "font-size")]
    pub font_size: Option<f64>,

    #[usage(long = "num-format", alias = "number-format")]
    pub num_format: Option<String>,

    /// [LLM-generated] Write updated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct StyleTableArgs {
    /// [LLM-generated] Input Excel file path
    pub file: String,

    /// [LLM-generated] Name of the target Excel Table
    #[usage(short, long)]
    pub name: String,

    /// [LLM-generated] Visual style theme name (e.g. "TableStyleMedium9", "TableStyleLight1", "TableStyleDark11")
    #[usage(short, long)]
    pub style: String,

    /// [LLM-generated] Write updated workbook to target output file path
    #[usage(short, long)]
    pub output: Option<String>,

    /// [LLM-generated] Save updated workbook in-place
    #[usage(short = 'i', long)]
    pub in_place: bool,
}
