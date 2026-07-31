//! `visi`: Command line interface for the visi spreadsheet application.
//! Consumes `libvisi` to read, evaluate, update, and export Excel (.xlsx) files.

use clap::Parser;
use libvisi::core::chart::ChartType;
use libvisi::core::{PivotAggregation, PivotArea};
use serde_json::json;
use visi::cli::{
    ChartArgs, ChartSubcommands, ChartTypeArg, Cli, ColArgs, ColSubcommands, Commands, EvalArgs,
    ExportArgs, ExportFormat, InfoArgs, OutputFormat, PivotAggArg, PivotAreaArg, PivotArgs,
    PivotSubcommands, ReadArgs, RowArgs, RowSubcommands, SetArgs, SheetArgs, SheetSubcommands,
    TableArgs, TableSubcommands,
};
use visi::engine::WorkbookManager;
use visi::format::{get_cell_display_val, render_grid};
use visi::utils::{
    EXIT_ENGINE_ERROR, EXIT_IO_ERROR, EXIT_USAGE_ERROR, col_idx_to_letters, exit_with_error,
    parse_cell_ref, parse_col_spec, parse_range_ref, parse_row_spec,
};

fn main() {
    let cli = Cli::parse();
    let quiet = cli.quiet;

    match cli.command {
        Commands::Info(args) => handle_info(args),
        Commands::Read(args) => handle_read(args),
        Commands::Set(args) => handle_set(args, quiet),
        Commands::Eval(args) => handle_eval(args, quiet),
        Commands::Sheet(args) => handle_sheet(args, quiet),
        Commands::Row(args) => handle_row(args, quiet),
        Commands::Col(args) => handle_col(args, quiet),
        Commands::Chart(args) => handle_chart(args, quiet),
        Commands::Table(args) => handle_table(args, quiet),
        Commands::Pivot(args) => handle_pivot(args, quiet),
        Commands::Export(args) => handle_export(args),
    }
}

/// Ensure output target (either --output or --in-place) is specified
fn resolve_output_path(output: Option<String>, in_place: bool, input_file: &str) -> String {
    if in_place {
        if input_file == "-" {
            exit_with_error(
                "Cannot use --in-place (-i) when reading from stdin (-)",
                EXIT_USAGE_ERROR,
            );
        }
        input_file.to_string()
    } else if let Some(out) = output {
        out
    } else {
        exit_with_error(
            "Must specify output path with --output <PATH> or use --in-place (-i) to save changes",
            EXIT_USAGE_ERROR,
        );
    }
}

fn handle_info(args: InfoArgs) {
    let wb = WorkbookManager::load_file(&args.file).unwrap_or_else(|e| {
        exit_with_error(e, EXIT_IO_ERROR);
    });

    let summary = wb.get_summary(&args.file);

    if args.json {
        let val = json!({
            "file": summary.file_name,
            "sheets_count": summary.sheet_count,
            "charts_count": summary.chart_count,
            "sheets": summary.sheets.iter().map(|s| json!({
                "name": s.name,
                "rows": s.row_count,
                "cols": s.col_count,
                "formulas": s.formula_count
            })).collect::<Vec<_>>()
        });
        println!("{}", serde_json::to_string_pretty(&val).unwrap());
    } else {
        println!("Workbook: {}", summary.file_name);
        println!("Sheets: {}", summary.sheet_count);
        println!("Charts: {}", summary.chart_count);
        println!();
        println!(
            "{:<20} {:<12} {:<10}",
            "Sheet Name", "Dimensions", "Formulas"
        );
        println!("{}", "-".repeat(45));
        for sheet in &summary.sheets {
            let dims = format!("{}x{}", sheet.row_count, sheet.col_count);
            println!(
                "{:<20} {:<12} {:<10}",
                sheet.name, dims, sheet.formula_count
            );
        }
    }
}

fn handle_read(args: ReadArgs) {
    let mut wb = WorkbookManager::load_file(&args.file).unwrap_or_else(|e| {
        exit_with_error(e, EXIT_IO_ERROR);
    });

    if args.eval {
        wb.evaluate().unwrap_or_else(|e| {
            exit_with_error(
                format!("Formula evaluation failed: {}", e),
                EXIT_ENGINE_ERROR,
            );
        });
    }

    // Determine target sheet and bounds
    let mut sheet_name_opt = args.sheet.clone();
    let mut target_cell = None;
    let mut target_range = None;

    if let Some(ref c_str) = args.cell {
        let (s_name, row, col) = parse_cell_ref(c_str).unwrap_or_else(|e| {
            exit_with_error(e, EXIT_USAGE_ERROR);
        });
        if s_name.is_some() {
            sheet_name_opt = s_name;
        }
        target_cell = Some((row, col));
    } else if let Some(ref r_str) = args.range {
        let (s_name, s_row, s_col, e_row, e_col) = parse_range_ref(r_str).unwrap_or_else(|e| {
            exit_with_error(e, EXIT_USAGE_ERROR);
        });
        if s_name.is_some() {
            sheet_name_opt = s_name;
        }
        target_range = Some((s_row, s_col, e_row, e_col));
    }

    let sheet_idx = wb
        .find_sheet_index(sheet_name_opt.as_deref())
        .unwrap_or_else(|e| {
            exit_with_error(e, EXIT_USAGE_ERROR);
        });

    let sheet = &wb.sheets[sheet_idx];

    // Single cell display
    if let Some((row, col)) = target_cell {
        let val_str = get_cell_display_val(sheet, row, col, args.raw);
        if args.format == OutputFormat::Json {
            let raw_src = sheet
                .get_src(&libvisi::core::CellRef::new(row, col))
                .cloned()
                .unwrap_or_default();
            let json_out = json!({
                "sheet": sheet.name,
                "cell": format!("{}{}", col_idx_to_letters(col), row + 1),
                "row": row + 1,
                "col": col_idx_to_letters(col),
                "value": val_str,
                "formula": if raw_src.starts_with('=') { Some(raw_src) } else { None }
            });
            println!("{}", serde_json::to_string_pretty(&json_out).unwrap());
        } else {
            println!("{}", val_str);
        }
        return;
    }

    // Grid / Range display
    let (min_row, min_col, max_row, max_col) = match target_range {
        Some(range) => range,
        None => (
            0,
            0,
            sheet.row_count().saturating_sub(1),
            sheet.col_count().saturating_sub(1),
        ),
    };

    let rendered = render_grid(
        sheet,
        min_row,
        min_col,
        max_row,
        max_col,
        args.raw,
        args.headers,
        args.format,
    );
    print!("{}", rendered);
}

fn handle_set(args: SetArgs, quiet: bool) {
    let mut wb = WorkbookManager::load_file_or_create(&args.file).unwrap_or_else(|e| {
        exit_with_error(e, EXIT_IO_ERROR);
    });

    let default_sheet_idx = wb
        .find_sheet_index(args.sheet.as_deref())
        .unwrap_or_else(|e| {
            exit_with_error(e, EXIT_USAGE_ERROR);
        });

    let mut updates: Vec<(usize, usize, usize, String)> = Vec::new(); // (sheet_idx, row, col, val)

    // 1. Parallel --cell and --value pairs
    for (cell_str, val_str) in args.cell.iter().zip(args.value.iter()) {
        let (s_name, row, col) = parse_cell_ref(cell_str).unwrap_or_else(|e| {
            exit_with_error(e, EXIT_USAGE_ERROR);
        });
        let sheet_idx = match s_name {
            Some(name) => wb.find_sheet_index(Some(&name)).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            }),
            None => default_sheet_idx,
        };
        updates.push((sheet_idx, row, col, val_str.clone()));
    }

    // 2. Multiple --set pairs (e.g. A1=100)
    for pair in &args.set_pairs {
        let parts: Vec<&str> = pair.splitn(2, '=').collect();
        if parts.len() != 2 {
            exit_with_error(
                format!(
                    "Invalid --set pair '{}', expected format CELL=VALUE (e.g. A1=100)",
                    pair
                ),
                EXIT_USAGE_ERROR,
            );
        }
        let cell_str = parts[0];
        let val_str = parts[1];

        let (s_name, row, col) = parse_cell_ref(cell_str).unwrap_or_else(|e| {
            exit_with_error(e, EXIT_USAGE_ERROR);
        });
        let sheet_idx = match s_name {
            Some(name) => wb.find_sheet_index(Some(&name)).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            }),
            None => default_sheet_idx,
        };
        updates.push((sheet_idx, row, col, val_str.to_string()));
    }

    if updates.is_empty() {
        exit_with_error(
            "Must specify cell assignment using --cell <CELL> --value <VALUE> or --set <CELL=VALUE>",
            EXIT_USAGE_ERROR,
        );
    }

    // Apply updates
    for (s_idx, row, col, val) in &updates {
        wb.set_cell(*s_idx, *row, *col, val.clone());
    }

    // Evaluate formulas if required
    if args.eval {
        wb.evaluate().unwrap_or_else(|e| {
            exit_with_error(
                format!("Formula evaluation failed: {}", e),
                EXIT_ENGINE_ERROR,
            );
        });
    }

    let save_path = resolve_output_path(args.output, args.in_place, &args.file);
    wb.save_file(&save_path).unwrap_or_else(|e| {
        exit_with_error(e, EXIT_IO_ERROR);
    });

    if !quiet {
        eprintln!(
            "Successfully updated {} cell(s) and saved to '{}'.",
            updates.len(),
            save_path
        );
    }
}

fn handle_eval(args: EvalArgs, quiet: bool) {
    let mut wb = WorkbookManager::load_file(&args.file).unwrap_or_else(|e| {
        exit_with_error(e, EXIT_IO_ERROR);
    });

    wb.evaluate().unwrap_or_else(|e| {
        exit_with_error(
            format!("Formula evaluation failed: {}", e),
            EXIT_ENGINE_ERROR,
        );
    });

    if args.print {
        let sheet_idx = wb
            .find_sheet_index(args.sheet.as_deref())
            .unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });
        let sheet = &wb.sheets[sheet_idx];
        let rendered = render_grid(
            sheet,
            0,
            0,
            sheet.row_count().saturating_sub(1),
            sheet.col_count().saturating_sub(1),
            false,
            false,
            args.format,
        );
        print!("{}", rendered);
    }

    // Save if output path or in-place specified
    if args.in_place || args.output.is_some() {
        let save_path = resolve_output_path(args.output, args.in_place, &args.file);
        wb.save_file(&save_path).unwrap_or_else(|e| {
            exit_with_error(e, EXIT_IO_ERROR);
        });
        if !quiet {
            eprintln!(
                "Successfully evaluated workbook and saved to '{}'.",
                save_path
            );
        }
    } else if !args.print {
        exit_with_error(
            "Must specify --output <PATH>, --in-place (-i), or --print (-p)",
            EXIT_USAGE_ERROR,
        );
    }
}

fn handle_sheet(args: SheetArgs, quiet: bool) {
    match args.command {
        SheetSubcommands::List(list_args) => {
            let wb = WorkbookManager::load_file(&list_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let summary = wb.get_summary(&list_args.file);
            if list_args.json {
                let json_val = json!(
                    summary
                        .sheets
                        .iter()
                        .map(|s| json!({
                            "name": s.name,
                            "rows": s.row_count,
                            "cols": s.col_count,
                            "formulas": s.formula_count
                        }))
                        .collect::<Vec<_>>()
                );
                println!("{}", serde_json::to_string_pretty(&json_val).unwrap());
            } else {
                for s in &summary.sheets {
                    println!(
                        "{}\t{}x{}\t{} formulas",
                        s.name, s.row_count, s.col_count, s.formula_count
                    );
                }
            }
        }
        SheetSubcommands::Add(add_args) => {
            let mut wb = WorkbookManager::load_file(&add_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.add_sheet(&add_args.name).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });
            let save_path = resolve_output_path(add_args.output, add_args.in_place, &add_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Added sheet '{}' and saved to '{}'.",
                    add_args.name, save_path
                );
            }
        }
        SheetSubcommands::Delete(del_args) => {
            let mut wb = WorkbookManager::load_file(&del_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.delete_sheet(&del_args.name).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });
            let save_path = resolve_output_path(del_args.output, del_args.in_place, &del_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Deleted sheet '{}' and saved to '{}'.",
                    del_args.name, save_path
                );
            }
        }
        SheetSubcommands::Rename(ren_args) => {
            let mut wb = WorkbookManager::load_file(&ren_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.rename_sheet(&ren_args.old, &ren_args.new)
                .unwrap_or_else(|e| {
                    exit_with_error(e, EXIT_USAGE_ERROR);
                });
            let save_path = resolve_output_path(ren_args.output, ren_args.in_place, &ren_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Renamed sheet '{}' -> '{}' and saved to '{}'.",
                    ren_args.old, ren_args.new, save_path
                );
            }
        }
    }
}

fn handle_row(args: RowArgs, quiet: bool) {
    match args.command {
        RowSubcommands::Insert(row_args) => {
            let mut wb = WorkbookManager::load_file(&row_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let sheet_idx = wb
                .find_sheet_index(row_args.sheet.as_deref())
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let row_idx = parse_row_spec(&row_args.index.to_string()).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });

            wb.insert_row(sheet_idx, row_idx).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_ENGINE_ERROR);
            });

            let save_path = resolve_output_path(row_args.output, row_args.in_place, &row_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Inserted row at index {} and saved to '{}'.",
                    row_args.index, save_path
                );
            }
        }
        RowSubcommands::Delete(row_args) => {
            let mut wb = WorkbookManager::load_file(&row_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let sheet_idx = wb
                .find_sheet_index(row_args.sheet.as_deref())
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let row_idx = parse_row_spec(&row_args.index.to_string()).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });

            wb.delete_row(sheet_idx, row_idx).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_ENGINE_ERROR);
            });

            let save_path = resolve_output_path(row_args.output, row_args.in_place, &row_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Deleted row at index {} and saved to '{}'.",
                    row_args.index, save_path
                );
            }
        }
    }
}

fn handle_col(args: ColArgs, quiet: bool) {
    match args.command {
        ColSubcommands::Insert(col_args) => {
            let mut wb = WorkbookManager::load_file(&col_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let sheet_idx = wb
                .find_sheet_index(col_args.sheet.as_deref())
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let col_idx = parse_col_spec(&col_args.index).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });

            wb.insert_col(sheet_idx, col_idx).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_ENGINE_ERROR);
            });

            let save_path = resolve_output_path(col_args.output, col_args.in_place, &col_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Inserted column '{}' and saved to '{}'.",
                    col_args.index, save_path
                );
            }
        }
        ColSubcommands::Delete(col_args) => {
            let mut wb = WorkbookManager::load_file(&col_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let sheet_idx = wb
                .find_sheet_index(col_args.sheet.as_deref())
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let col_idx = parse_col_spec(&col_args.index).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });

            wb.delete_col(sheet_idx, col_idx).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_ENGINE_ERROR);
            });

            let save_path = resolve_output_path(col_args.output, col_args.in_place, &col_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Deleted column '{}' and saved to '{}'.",
                    col_args.index, save_path
                );
            }
        }
    }
}

fn handle_chart(args: ChartArgs, quiet: bool) {
    match args.command {
        ChartSubcommands::List(list_args) => {
            let wb = WorkbookManager::load_file(&list_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            if list_args.json {
                let json_charts = json!(
                    wb.charts
                        .iter()
                        .map(|c| json!({
                            "id": c.id,
                            "name": c.name,
                            "type": format!("{:?}", c.chart_type),
                            "data_range": c.data_range,
                            "title": c.title
                        }))
                        .collect::<Vec<_>>()
                );
                println!("{}", serde_json::to_string_pretty(&json_charts).unwrap());
            } else {
                if wb.charts.is_empty() {
                    println!("No charts found in workbook.");
                    return;
                }
                println!(
                    "{:<12} {:<15} {:<10} {:<20} {:<20}",
                    "ID", "Name", "Type", "Range", "Title"
                );
                println!("{}", "-".repeat(77));
                for c in &wb.charts {
                    let title = c.title.as_deref().unwrap_or("-");
                    println!(
                        "{:<12} {:<15} {:<10?} {:<20} {:<20}",
                        c.id, c.name, c.chart_type, c.data_range, title
                    );
                }
            }
        }
        ChartSubcommands::Add(add_args) => {
            let mut wb = WorkbookManager::load_file(&add_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let sheet_name = match add_args.sheet {
                Some(name) => name,
                None => {
                    let s_idx = wb
                        .find_sheet_index(None)
                        .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));
                    wb.sheets[s_idx].name.clone()
                }
            };

            let c_type = match add_args.chart_type {
                ChartTypeArg::Line => ChartType::Line,
                ChartTypeArg::Bar => ChartType::Bar,
                ChartTypeArg::Column => ChartType::Column,
                ChartTypeArg::Pie => ChartType::Pie,
                ChartTypeArg::Scatter => ChartType::Scatter,
                ChartTypeArg::Area => ChartType::Area,
            };

            let chart_id = wb
                .add_chart(&sheet_name, c_type, add_args.range, add_args.title)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR));

            let save_path = resolve_output_path(add_args.output, add_args.in_place, &add_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Added chart (ID {}) and saved to '{}'.",
                    chart_id, save_path
                );
            }
        }
        ChartSubcommands::Delete(del_args) => {
            let mut wb = WorkbookManager::load_file(&del_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            wb.delete_chart(del_args.id).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });

            let save_path = resolve_output_path(del_args.output, del_args.in_place, &del_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Deleted chart ID {} and saved to '{}'.",
                    del_args.id, save_path
                );
            }
        }
    }
}

/// Resolves a table column specifier that's either a 1-based index (e.g.
/// "2") or an existing column name (case-insensitive), returning a 0-based
/// index relative to the table.
fn resolve_table_column_index(
    table: &libvisi::core::ExcelTable,
    spec: &str,
) -> Result<usize, String> {
    let trimmed = spec.trim();
    if let Ok(n) = trimmed.parse::<usize>() {
        if n == 0 || n > table.columns.len() {
            return Err(format!(
                "Column index {} out of bounds (table '{}' has {} columns)",
                n,
                table.name,
                table.columns.len()
            ));
        }
        return Ok(n - 1);
    }
    table.local_column_index(trimmed).ok_or_else(|| {
        format!(
            "Column '{}' not found in table '{}' (columns: {})",
            spec,
            table.name,
            table.columns.join(", ")
        )
    })
}

fn handle_table(args: TableArgs, quiet: bool) {
    match args.command {
        TableSubcommands::List(list_args) => {
            let wb = WorkbookManager::load_file(&list_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let tables = wb.list_tables();

            if list_args.json {
                let json_tables = json!(
                    tables
                        .iter()
                        .map(|(sheet_name, t)| json!({
                            "name": t.name,
                            "sheet": sheet_name,
                            "range": format!(
                                "{}{}:{}{}",
                                col_idx_to_letters(t.start_col), t.start_row + 1,
                                col_idx_to_letters(t.end_col), t.end_row + 1
                            ),
                            "columns": t.columns,
                            "has_header_row": t.has_header_row,
                            "has_totals_row": t.has_totals_row
                        }))
                        .collect::<Vec<_>>()
                );
                println!("{}", serde_json::to_string_pretty(&json_tables).unwrap());
            } else {
                if tables.is_empty() {
                    println!("No tables found in workbook.");
                    return;
                }
                println!(
                    "{:<15} {:<15} {:<12} {:<8} {:<8} {}",
                    "Name", "Sheet", "Range", "Header", "Totals", "Columns"
                );
                println!("{}", "-".repeat(90));
                for (sheet_name, t) in &tables {
                    let range = format!(
                        "{}{}:{}{}",
                        col_idx_to_letters(t.start_col),
                        t.start_row + 1,
                        col_idx_to_letters(t.end_col),
                        t.end_row + 1
                    );
                    println!(
                        "{:<15} {:<15} {:<12} {:<8} {:<8} {}",
                        t.name,
                        sheet_name,
                        range,
                        t.has_header_row,
                        t.has_totals_row,
                        t.columns.join(", ")
                    );
                }
            }
        }
        TableSubcommands::Add(add_args) => {
            let mut wb = WorkbookManager::load_file(&add_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let (range_sheet, start_row, start_col, end_row, end_col) =
                parse_range_ref(&add_args.range)
                    .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let sheet_name = match add_args.sheet.or(range_sheet) {
                Some(name) => name,
                None => {
                    let s_idx = wb
                        .find_sheet_index(None)
                        .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));
                    wb.sheets[s_idx].name.clone()
                }
            };

            let table_id = wb
                .add_table(
                    Some(&sheet_name),
                    &add_args.name,
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                    !add_args.no_header_row,
                    add_args.totals_row,
                )
                .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR));

            let save_path = resolve_output_path(add_args.output, add_args.in_place, &add_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Added table '{}' (ID {}) on sheet '{}' and saved to '{}'.",
                    add_args.name, table_id, sheet_name, save_path
                );
            }
        }
        TableSubcommands::Delete(del_args) => {
            let mut wb = WorkbookManager::load_file(&del_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.delete_table(&del_args.name).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });
            let save_path = resolve_output_path(del_args.output, del_args.in_place, &del_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Deleted table '{}' and saved to '{}'.",
                    del_args.name, save_path
                );
            }
        }
        TableSubcommands::Rename(ren_args) => {
            let mut wb = WorkbookManager::load_file(&ren_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.rename_table(&ren_args.old, &ren_args.new)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));
            let save_path = resolve_output_path(ren_args.output, ren_args.in_place, &ren_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Renamed table '{}' -> '{}' and saved to '{}'.",
                    ren_args.old, ren_args.new, save_path
                );
            }
        }
        TableSubcommands::Resize(resize_args) => {
            let mut wb = WorkbookManager::load_file(&resize_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let (_, start_row, start_col, end_row, end_col) = parse_range_ref(&resize_args.range)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let (_, existing) = wb.find_table(&resize_args.name).unwrap_or_else(|| {
                exit_with_error(
                    format!("Table '{}' not found", resize_args.name),
                    EXIT_USAGE_ERROR,
                )
            });
            if (start_row, start_col) != (existing.start_row, existing.start_col) {
                exit_with_error(
                    format!(
                        "Cannot move a table's top-left corner (table '{}' starts at {}{}); \
                         specify a range with the same starting cell",
                        resize_args.name,
                        col_idx_to_letters(existing.start_col),
                        existing.start_row + 1
                    ),
                    EXIT_USAGE_ERROR,
                );
            }

            wb.resize_table(&resize_args.name, end_row, end_col)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR));

            let save_path =
                resolve_output_path(resize_args.output, resize_args.in_place, &resize_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Resized table '{}' and saved to '{}'.",
                    resize_args.name, save_path
                );
            }
        }
        TableSubcommands::RenameColumn(rc_args) => {
            let mut wb = WorkbookManager::load_file(&rc_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let (_, table) = wb.find_table(&rc_args.name).unwrap_or_else(|| {
                exit_with_error(
                    format!("Table '{}' not found", rc_args.name),
                    EXIT_USAGE_ERROR,
                )
            });
            let col_index = resolve_table_column_index(table, &rc_args.column)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            wb.rename_table_column(&rc_args.name, col_index, &rc_args.new_name)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

            let save_path = resolve_output_path(rc_args.output, rc_args.in_place, &rc_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Renamed column '{}' -> '{}' in table '{}' and saved to '{}'.",
                    rc_args.column, rc_args.new_name, rc_args.name, save_path
                );
            }
        }
    }
}

fn pivot_area_arg_to_area(area: PivotAreaArg) -> PivotArea {
    match area {
        PivotAreaArg::Row => PivotArea::Row,
        PivotAreaArg::Column => PivotArea::Column,
        PivotAreaArg::Value => PivotArea::Value,
        PivotAreaArg::Filter => PivotArea::Filter,
    }
}

fn pivot_agg_arg_to_aggregation(agg: PivotAggArg) -> PivotAggregation {
    match agg {
        PivotAggArg::Sum => PivotAggregation::Sum,
        PivotAggArg::Count => PivotAggregation::Count,
        PivotAggArg::CountNumbers => PivotAggregation::CountNumbers,
        PivotAggArg::Average => PivotAggregation::Average,
        PivotAggArg::Max => PivotAggregation::Max,
        PivotAggArg::Min => PivotAggregation::Min,
    }
}

fn handle_pivot(args: PivotArgs, quiet: bool) {
    match args.command {
        PivotSubcommands::List(list_args) => {
            let wb = WorkbookManager::load_file(&list_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            let pivots = wb.list_pivot_tables();

            if list_args.json {
                let json_pivots = json!(
                    pivots
                        .iter()
                        .map(|p| json!({
                            "id": p.id,
                            "name": p.name,
                            "row_fields": p.row_fields.iter().map(|f| &f.column).collect::<Vec<_>>(),
                            "col_fields": p.col_fields.iter().map(|f| &f.column).collect::<Vec<_>>(),
                            "value_fields": p.value_fields.iter().map(|f| f.label()).collect::<Vec<_>>(),
                            "filter_fields": p.filter_fields.iter().map(|f| &f.column).collect::<Vec<_>>(),
                        }))
                        .collect::<Vec<_>>()
                );
                println!("{}", serde_json::to_string_pretty(&json_pivots).unwrap());
            } else {
                if pivots.is_empty() {
                    println!("No pivot tables found in workbook.");
                    return;
                }
                println!(
                    "{:<15} {:<20} {:<20} {:<20} {}",
                    "Name", "Rows", "Columns", "Values", "Filters"
                );
                println!("{}", "-".repeat(90));
                for p in pivots {
                    println!(
                        "{:<15} {:<20} {:<20} {:<20} {}",
                        p.name,
                        p.row_fields
                            .iter()
                            .map(|f| f.column.as_str())
                            .collect::<Vec<_>>()
                            .join(", "),
                        p.col_fields
                            .iter()
                            .map(|f| f.column.as_str())
                            .collect::<Vec<_>>()
                            .join(", "),
                        p.value_fields
                            .iter()
                            .map(|f| f.label())
                            .collect::<Vec<_>>()
                            .join(", "),
                        p.filter_fields
                            .iter()
                            .map(|f| f.column.as_str())
                            .collect::<Vec<_>>()
                            .join(", "),
                    );
                }
            }
        }
        PivotSubcommands::Create(add_args) => {
            let mut wb = WorkbookManager::load_file(&add_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let (dest_sheet_from_cell, dest_row, dest_col) = parse_cell_ref(&add_args.dest)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));
            let dest_sheet = add_args.dest_sheet.clone().or(dest_sheet_from_cell);

            let pivot_id = match (&add_args.source_table, &add_args.source_range) {
                (Some(table_name), None) => wb
                    .add_pivot_table_from_table(
                        &add_args.name,
                        table_name,
                        dest_sheet.as_deref(),
                        dest_row,
                        dest_col,
                    )
                    .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR)),
                (None, Some(range)) => {
                    let (range_sheet, start_row, start_col, end_row, end_col) =
                        parse_range_ref(range)
                            .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));
                    let source_sheet = add_args.source_sheet.clone().or(range_sheet);
                    wb.add_pivot_table_from_range(
                        &add_args.name,
                        source_sheet.as_deref(),
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        dest_sheet.as_deref(),
                        dest_row,
                        dest_col,
                    )
                    .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR))
                }
                (Some(_), Some(_)) => exit_with_error(
                    "Specify only one of --source-table or --source-range, not both",
                    EXIT_USAGE_ERROR,
                ),
                (None, None) => exit_with_error(
                    "Must specify a source using --source-table <NAME> or --source-range <RANGE>",
                    EXIT_USAGE_ERROR,
                ),
            };

            let save_path = resolve_output_path(add_args.output, add_args.in_place, &add_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Created pivot table '{}' (ID {}) and saved to '{}'.",
                    add_args.name, pivot_id, save_path
                );
            }
        }
        PivotSubcommands::Delete(del_args) => {
            let mut wb = WorkbookManager::load_file(&del_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.delete_pivot_table(&del_args.name).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_USAGE_ERROR);
            });
            let save_path = resolve_output_path(del_args.output, del_args.in_place, &del_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Deleted pivot table '{}' and saved to '{}'.",
                    del_args.name, save_path
                );
            }
        }
        PivotSubcommands::Rename(ren_args) => {
            let mut wb = WorkbookManager::load_file(&ren_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            wb.rename_pivot_table(&ren_args.old, &ren_args.new)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));
            let save_path = resolve_output_path(ren_args.output, ren_args.in_place, &ren_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Renamed pivot table '{}' -> '{}' and saved to '{}'.",
                    ren_args.old, ren_args.new, save_path
                );
            }
        }
        PivotSubcommands::Refresh(refresh_args) => {
            let mut wb = WorkbookManager::load_file(&refresh_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let names: Vec<String> = if refresh_args.all {
                wb.list_pivot_tables()
                    .iter()
                    .map(|p| p.name.clone())
                    .collect()
            } else if let Some(name) = &refresh_args.name {
                vec![name.clone()]
            } else {
                exit_with_error("Must specify --name <NAME> or --all", EXIT_USAGE_ERROR)
            };

            for name in &names {
                wb.refresh_pivot_table(name).unwrap_or_else(|e| {
                    exit_with_error(e, EXIT_ENGINE_ERROR);
                });
            }

            let save_path = resolve_output_path(
                refresh_args.output,
                refresh_args.in_place,
                &refresh_args.file,
            );
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Refreshed {} pivot table(s) and saved to '{}'.",
                    names.len(),
                    save_path
                );
            }
        }
        PivotSubcommands::AddField(field_args) => {
            let mut wb = WorkbookManager::load_file(&field_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let area = pivot_area_arg_to_area(field_args.area);
            let agg = field_args.agg.map(pivot_agg_arg_to_aggregation);
            wb.add_pivot_field(&field_args.name, area, &field_args.column, agg)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR));

            if let (Some(label), Some(pivot)) = (
                field_args.label.clone(),
                wb.pivot_tables
                    .iter()
                    .position(|p| p.name.eq_ignore_ascii_case(&field_args.name)),
            ) {
                if let Some(vf) = wb.pivot_tables[pivot]
                    .value_fields
                    .iter_mut()
                    .rev()
                    .find(|f| f.column.eq_ignore_ascii_case(&field_args.column))
                {
                    vf.custom_name = Some(label);
                }
                wb.refresh_pivot_table(&field_args.name)
                    .unwrap_or_else(|e| {
                        exit_with_error(e, EXIT_ENGINE_ERROR);
                    });
            }

            let save_path =
                resolve_output_path(field_args.output, field_args.in_place, &field_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Added field '{}' to pivot table '{}' and saved to '{}'.",
                    field_args.column, field_args.name, save_path
                );
            }
        }
        PivotSubcommands::RemoveField(field_args) => {
            let mut wb = WorkbookManager::load_file(&field_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let area = pivot_area_arg_to_area(field_args.area);
            wb.remove_pivot_field(&field_args.name, area, &field_args.column)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR));

            let save_path =
                resolve_output_path(field_args.output, field_args.in_place, &field_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Removed field '{}' from pivot table '{}' and saved to '{}'.",
                    field_args.column, field_args.name, save_path
                );
            }
        }
        PivotSubcommands::Filter(filter_args) => {
            let mut wb = WorkbookManager::load_file(&filter_args.file).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });

            let values = if filter_args.clear {
                None
            } else if filter_args.values.is_empty() {
                exit_with_error(
                    "Must specify --values <V1,V2,...> or --clear",
                    EXIT_USAGE_ERROR,
                )
            } else {
                Some(filter_args.values.clone())
            };

            wb.set_pivot_filter(&filter_args.name, &filter_args.column, values)
                .unwrap_or_else(|e| exit_with_error(e, EXIT_ENGINE_ERROR));

            let save_path =
                resolve_output_path(filter_args.output, filter_args.in_place, &filter_args.file);
            wb.save_file(&save_path).unwrap_or_else(|e| {
                exit_with_error(e, EXIT_IO_ERROR);
            });
            if !quiet {
                eprintln!(
                    "Updated filter '{}' on pivot table '{}' and saved to '{}'.",
                    filter_args.column, filter_args.name, save_path
                );
            }
        }
    }
}

fn handle_export(args: ExportArgs) {
    let mut wb = WorkbookManager::load_file(&args.file).unwrap_or_else(|e| {
        exit_with_error(e, EXIT_IO_ERROR);
    });

    if args.eval {
        wb.evaluate().unwrap_or_else(|e| {
            exit_with_error(
                format!("Formula evaluation failed: {}", e),
                EXIT_ENGINE_ERROR,
            );
        });
    }

    let sheet_idx = wb
        .find_sheet_index(args.sheet.as_deref())
        .unwrap_or_else(|e| exit_with_error(e, EXIT_USAGE_ERROR));

    let sheet = &wb.sheets[sheet_idx];

    let out_format = match args.format {
        ExportFormat::Csv => OutputFormat::Csv,
        ExportFormat::Tsv => OutputFormat::Tsv,
        ExportFormat::Json => OutputFormat::Json,
    };

    let rendered = render_grid(
        sheet,
        0,
        0,
        sheet.row_count().saturating_sub(1),
        sheet.col_count().saturating_sub(1),
        false,
        false,
        out_format,
    );

    if let Some(ref out_path) = args.output {
        std::fs::write(out_path, rendered).unwrap_or_else(|e| {
            exit_with_error(
                format!("Failed to write export file '{}': {}", out_path, e),
                EXIT_IO_ERROR,
            );
        });
    } else {
        print!("{}", rendered);
    }
}
