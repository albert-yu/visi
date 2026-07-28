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
