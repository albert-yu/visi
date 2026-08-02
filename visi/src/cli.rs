use clap::{Args, Parser, Subcommand, ValueEnum};

#[derive(Parser, Debug)]
#[command(
    name = "visi",
    author,
    version,
    about = "Read, evaluate formulas, and update Excel (.xlsx) spreadsheets",
    long_about = "visi is a developer-friendly command line utility to inspect, evaluate formulas in, and update Excel (.xlsx) files powered by the libvisi calculation engine.\n\nExamples:\n  visi info data.xlsx\n  visi read data.xlsx --sheet Sheet1 --format table\n  visi set data.xlsx --sheet Sheet1 --cell A1 --value 100 --in-place\n  visi eval data.xlsx --output calculated.xlsx"
)]
pub struct Cli {
    #[command(subcommand)]
    pub command: Commands,

    /// Enable verbose logging output to stderr
    #[arg(short, long, global = true)]
    pub verbose: bool,

    /// Suppress non-essential informational messages to stderr
    #[arg(short, long, global = true)]
    pub quiet: bool,
}

#[derive(Subcommand, Debug)]
pub enum Commands {
    /// Display structure summary and metadata of an Excel workbook
    Info(InfoArgs),

    /// Read and display contents of a sheet, cell range, or single cell
    #[command(alias = "view")]
    Read(ReadArgs),

    /// Update cell values or formulas in an Excel workbook
    #[command(alias = "update")]
    Set(SetArgs),

    /// Recalculate all formulas across all sheets in the workbook
    #[command(alias = "recalc")]
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
    #[arg(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct ReadArgs {
    /// Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// Target sheet name (defaults to first sheet)
    #[arg(short, long)]
    pub sheet: Option<String>,

    /// Target cell range in A1 notation (e.g. A1:D10)
    #[arg(short, long)]
    pub range: Option<String>,

    /// Target single cell in A1 notation (e.g. A1)
    #[arg(short, long)]
    pub cell: Option<String>,

    /// Output display format [table, csv, tsv, json]
    #[arg(short, long, value_enum, default_value_t = OutputFormat::Table)]
    pub format: OutputFormat,

    /// Recalculate and execute formulas before displaying (enabled by default)
    #[arg(long, default_value_t = true, action = clap::ArgAction::Set)]
    pub eval: bool,

    /// Output raw cell formulas instead of calculated result values
    #[arg(long)]
    pub raw: bool,

    /// Treat the first row of selected cells as header names
    #[arg(long)]
    pub headers: bool,
}

#[derive(Args, Debug)]
pub struct SetArgs {
    /// Input Excel file path (use '-' to read from stdin)
    pub file: String,

    /// Target sheet name (defaults to first sheet)
    #[arg(short, long)]
    pub sheet: Option<String>,

    /// Target cell coordinate in A1 notation (e.g. A1)
    #[arg(short, long, action = clap::ArgAction::Append)]
    pub cell: Vec<String>,

    /// Value or formula to set (e.g. 100, "Hello", "=SUM(A1:A10)")
    #[arg(long, action = clap::ArgAction::Append)]
    pub value: Vec<String>,

    /// Set cell assignments in CELL=VALUE format (e.g. -S A1=100 -S B1="=A1*2")
    #[arg(short = 'S', long = "set", value_name = "CELL=VALUE")]
    pub set_pairs: Vec<String>,

    /// Write updated workbook to target output file path
    #[arg(short, long)]
    pub output: Option<String>,

    /// Save updated workbook in-place, overwriting the input file
    #[arg(short = 'i', long)]
    pub in_place: bool,

    /// Recalculate formulas after setting values (enabled by default)
    #[arg(long, default_value_t = true, action = clap::ArgAction::Set)]
    pub eval: bool,
}

#[derive(Args, Debug)]
pub struct EvalArgs {
    /// Input Excel file path
    pub file: String,

    /// Write calculated workbook to target output file path
    #[arg(short, long)]
    pub output: Option<String>,

    /// Overwrite input file with calculated formula results
    #[arg(short = 'i', long)]
    pub in_place: bool,

    /// Print calculated sheet content(s) to stdout
    #[arg(short, long)]
    pub print: bool,

    /// Specific sheet to print if --print is specified
    #[arg(short, long)]
    pub sheet: Option<String>,

    /// Output display format if --print is specified [table, csv, tsv, json]
    #[arg(short, long, value_enum, default_value_t = OutputFormat::Table)]
    pub format: OutputFormat,
}

#[derive(Args, Debug)]
pub struct SheetArgs {
    #[command(subcommand)]
    pub command: SheetSubcommands,
}

#[derive(Subcommand, Debug)]
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
    #[arg(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct SheetAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Name for the new worksheet
    #[arg(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct SheetDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the worksheet to delete
    #[arg(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct SheetRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current worksheet name
    #[arg(long)]
    pub old: String,
    /// New worksheet name
    #[arg(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct RowArgs {
    #[command(subcommand)]
    pub command: RowSubcommands,
}

#[derive(Subcommand, Debug)]
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
    #[arg(short, long)]
    pub sheet: Option<String>,
    /// 1-based row index (e.g. 1, 5)
    #[arg(short = 'x', long)]
    pub index: usize,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ColArgs {
    #[command(subcommand)]
    pub command: ColSubcommands,
}

#[derive(Subcommand, Debug)]
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
    #[arg(short, long)]
    pub sheet: Option<String>,
    /// Column letter or 1-based index (e.g. "B" or "2")
    #[arg(short = 'x', long)]
    pub index: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartArgs {
    #[command(subcommand)]
    pub command: ChartSubcommands,
}

#[derive(Subcommand, Debug)]
pub enum ChartSubcommands {
    /// List all charts in the workbook
    List(ChartListArgs),
    /// Add a new chart to a worksheet
    Add(ChartAddArgs),
    /// Delete a chart by ID
    Delete(ChartDeleteArgs),
}

#[derive(Args, Debug)]
pub struct ChartListArgs {
    /// Input Excel file path
    pub file: String,
    /// Output summary as JSON
    #[arg(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct ChartAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name
    #[arg(short, long)]
    pub sheet: Option<String>,
    /// Chart type [line, bar, column, pie, scatter, area]
    #[arg(short, long, value_enum)]
    pub chart_type: ChartTypeArg,
    /// Data range reference (e.g. Sheet1!A1:B10)
    #[arg(short, long)]
    pub range: String,
    /// Optional chart title
    #[arg(short, long)]
    pub title: Option<String>,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ChartDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Chart ID to delete
    #[arg(long)]
    pub id: u64,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableArgs {
    #[command(subcommand)]
    pub command: TableSubcommands,
}

#[derive(Subcommand, Debug)]
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
}

#[derive(Args, Debug)]
pub struct TableListArgs {
    /// Input Excel file path
    pub file: String,
    /// Output summary as JSON
    #[arg(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct TableAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name (defaults to first sheet, or the range's own
    /// sheet prefix if given, e.g. Sheet1!A1:D10)
    #[arg(short, long)]
    pub sheet: Option<String>,
    /// Name for the new table
    #[arg(short, long)]
    pub name: String,
    /// Cell range the table should occupy (e.g. A1:D10)
    #[arg(short, long)]
    pub range: String,
    /// Treat the range's first row as plain data, not column headers
    #[arg(long)]
    pub no_header_row: bool,
    /// Reserve the range's last row as a totals row
    #[arg(long)]
    pub totals_row: bool,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the table to delete
    #[arg(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current table name
    #[arg(long)]
    pub old: String,
    /// New table name
    #[arg(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableResizeArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the table to resize
    #[arg(short, long)]
    pub name: String,
    /// New cell range for the table; its top-left corner must match the
    /// table's current top-left corner (e.g. A1:E12)
    #[arg(short, long)]
    pub range: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct TableRenameColumnArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the table containing the column
    #[arg(short, long)]
    pub name: String,
    /// Existing column name or 1-based column index within the table
    #[arg(short, long)]
    pub column: String,
    /// New column name
    #[arg(long)]
    pub new_name: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotArgs {
    #[command(subcommand)]
    pub command: PivotSubcommands,
}

#[derive(Subcommand, Debug)]
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
    #[arg(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct PivotCreateArgs {
    /// Input Excel file path
    pub file: String,
    /// Name for the new pivot table
    #[arg(short, long)]
    pub name: String,
    /// Name of an existing Excel Table to use as the source (mutually
    /// exclusive with --source-range)
    #[arg(long)]
    pub source_table: Option<String>,
    /// Cell range to use as the source, first row treated as headers (e.g.
    /// A1:D100; mutually exclusive with --source-table)
    #[arg(long)]
    pub source_range: Option<String>,
    /// Worksheet the source range lives on (defaults to first sheet, or the
    /// range's own sheet prefix, e.g. Sheet1!A1:D10)
    #[arg(long)]
    pub source_sheet: Option<String>,
    /// Top-left cell of the pivot table's output (e.g. A1)
    #[arg(long, default_value = "A1")]
    pub dest: String,
    /// Worksheet the pivot table's output is written to (defaults to first
    /// sheet)
    #[arg(long)]
    pub dest_sheet: Option<String>,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotDeleteArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to delete
    #[arg(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current pivot table name
    #[arg(long)]
    pub old: String,
    /// New pivot table name
    #[arg(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRefreshArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to refresh (omit with --all to refresh every
    /// pivot table in the workbook)
    #[arg(short, long)]
    pub name: Option<String>,
    /// Refresh every pivot table in the workbook
    #[arg(long)]
    pub all: bool,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotAddFieldArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to modify
    #[arg(short, long)]
    pub name: String,
    /// Area to add the field to [row, column, value, filter]
    #[arg(short, long, value_enum)]
    pub area: PivotAreaArg,
    /// Source column name
    #[arg(short, long)]
    pub column: String,
    /// Aggregation function, only used when --area value [sum, count,
    /// count-numbers, average, max, min] (defaults to sum)
    #[arg(long, value_enum)]
    pub agg: Option<PivotAggArg>,
    /// Custom display label for a value field (defaults to e.g. "Sum of
    /// Amount")
    #[arg(long)]
    pub label: Option<String>,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotRemoveFieldArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to modify
    #[arg(short, long)]
    pub name: String,
    /// Area to remove the field from [row, column, value, filter]
    #[arg(short, long, value_enum)]
    pub area: PivotAreaArg,
    /// Source column name
    #[arg(short, long)]
    pub column: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct PivotFilterArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the pivot table to modify
    #[arg(short, long)]
    pub name: String,
    /// Filter field's source column name
    #[arg(short, long)]
    pub column: String,
    /// Comma-separated list of the only values to include (e.g.
    /// "East,West")
    #[arg(long, value_delimiter = ',')]
    pub values: Vec<String>,
    /// Remove the filter, allowing every value again
    #[arg(long)]
    pub clear: bool,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroArgs {
    #[command(subcommand)]
    pub command: MacroSubcommands,
}

#[derive(Subcommand, Debug)]
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
    #[arg(long)]
    pub json: bool,
}

#[derive(Args, Debug)]
pub struct MacroAddArgs {
    /// Input Excel file path
    pub file: String,
    /// Name for the new module
    #[arg(short, long)]
    pub name: String,
    /// Module kind
    #[arg(short, long, value_enum, default_value_t = VbaModuleKindArg::Standard)]
    pub kind: VbaModuleKindArg,
    /// Sheet this document module belongs to (required for --kind document)
    #[arg(short, long)]
    pub sheet: Option<String>,
    /// Module source code, given inline
    #[arg(long, conflicts_with = "source_file")]
    pub source: Option<String>,
    /// Module source code, read from a file
    #[arg(long)]
    pub source_file: Option<String>,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroRemoveArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the module to remove
    #[arg(short, long)]
    pub name: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroRenameArgs {
    /// Input Excel file path
    pub file: String,
    /// Current module name
    #[arg(long)]
    pub old: String,
    /// New module name
    #[arg(long)]
    pub new: String,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct MacroSetSourceArgs {
    /// Input Excel file path
    pub file: String,
    /// Name of the module to update
    #[arg(short, long)]
    pub name: String,
    /// New module source code, given inline
    #[arg(long, conflicts_with = "source_file")]
    pub source: Option<String>,
    /// New module source code, read from a file
    #[arg(long)]
    pub source_file: Option<String>,
    /// Write updated workbook to target output file
    #[arg(short, long)]
    pub output: Option<String>,
    /// Save updated workbook in-place
    #[arg(short = 'i', long)]
    pub in_place: bool,
}

#[derive(Args, Debug)]
pub struct ExportArgs {
    /// Input Excel file path
    pub file: String,
    /// Target worksheet name (defaults to first sheet)
    #[arg(short, long)]
    pub sheet: Option<String>,
    /// Export format [csv, tsv, json]
    #[arg(short, long, value_enum, default_value_t = ExportFormat::Csv)]
    pub format: ExportFormat,
    /// Output file path (if omitted, writes to stdout)
    #[arg(short, long)]
    pub output: Option<String>,
    /// Recalculate formulas before exporting (enabled by default)
    #[arg(long, default_value_t = true, action = clap::ArgAction::Set)]
    pub eval: bool,
}
