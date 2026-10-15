use std::collections::HashMap;

use crate::core::engine::{CellRef, Context, ResultData, Sheet};
use crate::core::function_registry::{is_known_function, normalize_function_name};
use crate::core::parser::{Expr, Op, parse_excel_formula};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormulaDiagnosticSeverity {
    Error,
    Warning,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormulaDiagnosticKind {
    Syntax,
    Reference,
    Function,
    Type,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaDiagnostic {
    pub severity: FormulaDiagnosticSeverity,
    pub kind: FormulaDiagnosticKind,
    pub message: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum FormulaType {
    Any,
    Blank,
    Number,
    Text(Option<String>),
    Boolean,
    Error,
    Array {
        item: Box<FormulaType>,
        rows: Option<usize>,
        cols: Option<usize>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ExpectedType {
    Any,
    Number,
    Text,
    Boolean,
    Reference,
    Array,
}

struct TypeContext<'a> {
    sheet: &'a Sheet,
    workbook_context: Option<&'a Context<'a>>,
    row: Option<usize>,
    bindings: HashMap<String, FormulaType>,
    diagnostics: Vec<FormulaDiagnostic>,
}

pub fn check_formula(
    input: &str,
    sheet: &Sheet,
    context: Option<&Context>,
    row: Option<usize>,
    col: Option<usize>,
) -> Vec<FormulaDiagnostic> {
    let src = input.strip_prefix('=').unwrap_or(input);
    if src.trim().is_empty() || !input.starts_with('=') {
        return Vec::new();
    }
    let ast = match parse_excel_formula(src) {
        Ok(ast) => ast,
        Err(e) => {
            return vec![FormulaDiagnostic {
                severity: FormulaDiagnosticSeverity::Error,
                kind: FormulaDiagnosticKind::Syntax,
                message: e,
            }];
        }
    };
    let mut ctx = TypeContext {
        sheet,
        workbook_context: context,
        row,
        bindings: HashMap::new(),
        diagnostics: Vec::new(),
    };
    let _ = col;
    infer_expr_type(&ast, &mut ctx);
    ctx.diagnostics
}

fn infer_expr_type(expr: &Expr, ctx: &mut TypeContext<'_>) -> FormulaType {
    match expr {
        Expr::Number(_) => FormulaType::Number,
        Expr::String(s) => FormulaType::Text(Some(s.clone())),
        Expr::Boolean(_) => FormulaType::Boolean,
        Expr::Error(_) => FormulaType::Error,
        Expr::Identifier(name) => ctx
            .bindings
            .get(&name.to_ascii_lowercase())
            .cloned()
            .unwrap_or_else(|| {
                push(
                    ctx,
                    FormulaDiagnosticSeverity::Error,
                    FormulaDiagnosticKind::Reference,
                    format!("Unknown name: {name}"),
                );
                FormulaType::Any
            }),
        Expr::CellRef {
            sheet, row, col, ..
        } => infer_cell_ref(sheet.as_deref(), *row, *col, ctx),
        Expr::RangeRef {
            sheet,
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        } => infer_range_ref(
            sheet.as_deref(),
            *start_row,
            *start_col,
            *end_row,
            *end_col,
            ctx,
        ),
        Expr::StructuredRef {
            sheet,
            column,
            is_this_row,
            section,
        } => infer_structured_ref(
            sheet.as_deref(),
            column.as_deref(),
            *is_this_row,
            *section,
            ctx,
        ),
        Expr::List(items) => {
            let mut item_type = FormulaType::Blank;
            for item in items {
                item_type = merge_types(item_type, infer_expr_type(item, ctx));
            }
            FormulaType::Array {
                item: Box::new(item_type),
                rows: Some(1),
                cols: Some(items.len()),
            }
        }
        Expr::Slice { expr, start, end } => {
            let value = infer_expr_type(expr, ctx);
            if let Some(start) = start {
                let ty = infer_expr_type(start, ctx);
                check_expected(ctx, &ty, ExpectedType::Number, "slice start");
            }
            if let Some(end) = end {
                let ty = infer_expr_type(end, ctx);
                check_expected(ctx, &ty, ExpectedType::Number, "slice end");
            }
            match value {
                FormulaType::Array { item, .. } => FormulaType::Array {
                    item,
                    rows: None,
                    cols: None,
                },
                FormulaType::Error => FormulaType::Error,
                _ => FormulaType::Any,
            }
        }
        Expr::UnaryOp { op, expr } => infer_unary(*op, expr, ctx),
        Expr::BinaryOp { op, left, right } => infer_binary(*op, left, right, ctx),
        Expr::FunctionCall { name, args } => infer_function_call(name, args, ctx),
    }
}

fn infer_cell_ref(
    sheet_name: Option<&str>,
    row: usize,
    col: usize,
    ctx: &mut TypeContext<'_>,
) -> FormulaType {
    match resolve_sheet(sheet_name, ctx.sheet, ctx.workbook_context) {
        Ok(sheet) => {
            if row >= sheet.row_count() || col >= sheet.col_count() {
                push(
                    ctx,
                    FormulaDiagnosticSeverity::Warning,
                    FormulaDiagnosticKind::Reference,
                    format!(
                        "Cell reference {}{} is outside sheet bounds",
                        col + 1,
                        row + 1
                    ),
                );
                FormulaType::Blank
            } else {
                result_type(&sheet.get_result_data(&CellRef::new(row, col)))
            }
        }
        Err(name) => {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Reference,
                format!("Sheet not found: {name}"),
            );
            FormulaType::Any
        }
    }
}

fn infer_range_ref(
    sheet_name: Option<&str>,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    ctx: &mut TypeContext<'_>,
) -> FormulaType {
    let sheet = match resolve_sheet(sheet_name, ctx.sheet, ctx.workbook_context) {
        Ok(sheet) => sheet,
        Err(name) => {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Reference,
                format!("Sheet not found: {name}"),
            );
            return FormulaType::Any;
        }
    };
    let actual_end_row = if end_row == usize::MAX {
        sheet.row_count().saturating_sub(1)
    } else {
        end_row
    };
    let actual_end_col = if end_col == usize::MAX {
        sheet.col_count().saturating_sub(1)
    } else {
        end_col
    };
    if start_row >= sheet.row_count()
        || start_col >= sheet.col_count()
        || actual_end_row >= sheet.row_count()
        || actual_end_col >= sheet.col_count()
    {
        push(
            ctx,
            FormulaDiagnosticSeverity::Warning,
            FormulaDiagnosticKind::Reference,
            "Range reference extends outside sheet bounds".to_string(),
        );
    }
    let rows = actual_end_row.saturating_sub(start_row).saturating_add(1);
    let cols = actual_end_col.saturating_sub(start_col).saturating_add(1);
    let item = if rows.saturating_mul(cols) <= 256 {
        let mut ty = FormulaType::Blank;
        for r in start_row..=actual_end_row.min(sheet.row_count().saturating_sub(1)) {
            for c in start_col..=actual_end_col.min(sheet.col_count().saturating_sub(1)) {
                ty = merge_types(ty, result_type(&sheet.get_result_data(&CellRef::new(r, c))));
            }
        }
        ty
    } else {
        FormulaType::Any
    };
    FormulaType::Array {
        item: Box::new(item),
        rows: Some(rows),
        cols: Some(cols),
    }
}

fn infer_structured_ref(
    table_name: Option<&str>,
    column: Option<&str>,
    is_this_row: bool,
    section: crate::core::SheetSection,
    ctx: &mut TypeContext<'_>,
) -> FormulaType {
    let Some(name) = table_name else {
        if let Some(column) = column {
            if is_this_row {
                return FormulaType::Any;
            }
            push(
                ctx,
                FormulaDiagnosticSeverity::Warning,
                FormulaDiagnosticKind::Reference,
                format!("Structured column reference [{column}] has no table name"),
            );
        }
        return FormulaType::Any;
    };
    let mut found = ctx.sheet.find_table(name).map(|table| (ctx.sheet, table));
    if found.is_none()
        && let Some(workbook_context) = ctx.workbook_context
    {
        for sheet in workbook_context.sheets.values() {
            if let Some(table) = sheet.find_table(name) {
                found = Some((sheet, table));
                break;
            }
        }
    }
    let Some((sheet, table)) = found else {
        push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Reference,
            format!("Table not found: {name}"),
        );
        return FormulaType::Any;
    };
    let cols = if let Some(column) = column {
        if table.local_column_index(column).is_none() {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Reference,
                format!("Column not found: {column}"),
            );
            return FormulaType::Any;
        }
        1
    } else {
        table.col_count()
    };
    if section == crate::core::SheetSection::Headers {
        if cols == 1 {
            FormulaType::Text(None)
        } else {
            FormulaType::Array {
                item: Box::new(FormulaType::Text(None)),
                rows: Some(1),
                cols: Some(cols),
            }
        }
    } else if is_this_row && cols == 1 {
        if ctx.row.is_none() {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Reference,
                "This row reference cannot be checked without row context".to_string(),
            );
        }
        FormulaType::Any
    } else {
        let rows = match section {
            crate::core::SheetSection::Totals => usize::from(table.totals_row().is_some()),
            crate::core::SheetSection::Data | crate::core::SheetSection::All => {
                table.data_row_count()
            }
            crate::core::SheetSection::Headers => 1,
        };
        let _ = sheet;
        FormulaType::Array {
            item: Box::new(FormulaType::Any),
            rows: Some(rows),
            cols: Some(cols),
        }
    }
}

fn infer_unary(op: Op, expr: &Expr, ctx: &mut TypeContext<'_>) -> FormulaType {
    let ty = infer_expr_type(expr, ctx);
    match op {
        Op::Sub | Op::Percent => {
            check_expected(ctx, &ty, ExpectedType::Number, "unary numeric operator");
            FormulaType::Number
        }
        Op::ImplicitIntersection => scalarize(ty),
        Op::Spill => match ty {
            FormulaType::Array { .. } => ty,
            FormulaType::Error => FormulaType::Error,
            _ => {
                if !is_reference_expr(expr) {
                    push(
                        ctx,
                        FormulaDiagnosticSeverity::Warning,
                        FormulaDiagnosticKind::Type,
                        "Spill operator is usually applied to a reference or array".to_string(),
                    );
                }
                ty
            }
        },
        _ => ty,
    }
}

fn infer_binary(op: Op, left: &Expr, right: &Expr, ctx: &mut TypeContext<'_>) -> FormulaType {
    if matches!(op, Op::Intersect | Op::Union) {
        if !is_reference_expr(left) {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Type,
                "Left operand must be a reference".to_string(),
            );
        }
        if !is_reference_expr(right) {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Type,
                "Right operand must be a reference".to_string(),
            );
        }
        let _ = infer_expr_type(left, ctx);
        let _ = infer_expr_type(right, ctx);
        return FormulaType::Array {
            item: Box::new(FormulaType::Any),
            rows: None,
            cols: None,
        };
    }
    let l_ty = infer_expr_type(left, ctx);
    let r_ty = infer_expr_type(right, ctx);
    match op {
        Op::Add | Op::Sub | Op::Mul | Op::Div | Op::Exp => {
            check_expected(ctx, &l_ty, ExpectedType::Number, "left operand");
            check_expected(ctx, &r_ty, ExpectedType::Number, "right operand");
            FormulaType::Number
        }
        Op::Concat => {
            check_expected(ctx, &l_ty, ExpectedType::Text, "left operand");
            check_expected(ctx, &r_ty, ExpectedType::Text, "right operand");
            FormulaType::Text(None)
        }
        Op::Eq | Op::Ne | Op::Lt | Op::Gt | Op::Le | Op::Ge => {
            check_scalar(ctx, &l_ty, "left operand");
            check_scalar(ctx, &r_ty, "right operand");
            FormulaType::Boolean
        }
        _ => FormulaType::Any,
    }
}

fn infer_function_call(name: &str, args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    let upper = normalize_function_name(name);
    if !is_known_function(&upper) {
        push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Function,
            format!("Unknown function: {name}"),
        );
        for arg in args {
            let _ = infer_expr_type(arg, ctx);
        }
        return FormulaType::Any;
    }
    match upper.as_str() {
        "IF" => infer_if(args, ctx),
        "IFERROR" | "IFNA" => infer_fixed_lazy(&upper, args, 2, ctx),
        "IFS" => infer_ifs(args, ctx),
        "SWITCH" => infer_switch(args, ctx),
        "CHOOSE" => infer_choose(args, ctx),
        "LET" => infer_let(args, ctx),
        "LAMBDA" => infer_lambda(args, ctx),
        "SUM" | "AVERAGE" | "MIN" | "MAX" | "PRODUCT" | "SUMSQ" => {
            check_arity(ctx, &upper, args.len(), 1, None);
            for arg in args {
                let ty = infer_expr_type(arg, ctx);
                if !matches!(ty, FormulaType::Array { .. }) {
                    check_expected(ctx, &ty, ExpectedType::Number, &upper);
                }
            }
            FormulaType::Number
        }
        "COUNT" | "COUNTA" | "COUNTBLANK" => {
            check_arity(ctx, &upper, args.len(), 1, None);
            for arg in args {
                let _ = infer_expr_type(arg, ctx);
            }
            FormulaType::Number
        }
        "ROUND" | "ROUNDDOWN" | "ROUNDUP" | "MROUND" | "MOD" | "POWER" => infer_signature(
            &upper,
            args,
            2,
            Some(2),
            &[ExpectedType::Number, ExpectedType::Number],
            FormulaType::Number,
            ctx,
        ),
        "ABS" | "SQRT" | "EXP" | "LN" | "LOG10" | "SIN" | "COS" | "TAN" | "ASIN" | "ACOS"
        | "ATAN" | "INT" | "TRUNC" | "SIGN" | "EVEN" | "ODD" | "FACT" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Number],
            FormulaType::Number,
            ctx,
        ),
        "AND" | "OR" | "XOR" => {
            check_arity(ctx, &upper, args.len(), 1, None);
            for arg in args {
                let ty = infer_expr_type(arg, ctx);
                check_expected(ctx, &ty, ExpectedType::Boolean, &upper);
            }
            FormulaType::Boolean
        }
        "NOT" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Boolean],
            FormulaType::Boolean,
            ctx,
        ),
        "LEFT" | "RIGHT" | "LEFTB" | "RIGHTB" => infer_signature(
            &upper,
            args,
            1,
            Some(2),
            &[ExpectedType::Text, ExpectedType::Number],
            FormulaType::Text(None),
            ctx,
        ),
        "MID" | "MIDB" => infer_signature(
            &upper,
            args,
            3,
            Some(3),
            &[
                ExpectedType::Text,
                ExpectedType::Number,
                ExpectedType::Number,
            ],
            FormulaType::Text(None),
            ctx,
        ),
        "LEN" | "LENB" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Text],
            FormulaType::Number,
            ctx,
        ),
        "LOWER" | "UPPER" | "PROPER" | "TRIM" | "CLEAN" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Text],
            FormulaType::Text(None),
            ctx,
        ),
        "FIND" | "FINDB" | "SEARCH" | "SEARCHB" => infer_signature(
            &upper,
            args,
            2,
            Some(3),
            &[ExpectedType::Text, ExpectedType::Text, ExpectedType::Number],
            FormulaType::Number,
            ctx,
        ),
        "CONCAT" => {
            check_arity(ctx, &upper, args.len(), 1, None);
            for arg in args {
                let _ = infer_expr_type(arg, ctx);
            }
            FormulaType::Text(None)
        }
        "CONCATENATE" => {
            check_arity(ctx, &upper, args.len(), 1, None);
            for arg in args {
                let ty = infer_expr_type(arg, ctx);
                check_expected(ctx, &ty, ExpectedType::Text, &upper);
            }
            FormulaType::Text(None)
        }
        "TEXTJOIN" => infer_signature(
            &upper,
            args,
            3,
            None,
            &[ExpectedType::Text, ExpectedType::Boolean, ExpectedType::Any],
            FormulaType::Text(None),
            ctx,
        ),
        "TEXT" => infer_signature(
            &upper,
            args,
            2,
            Some(2),
            &[ExpectedType::Any, ExpectedType::Text],
            FormulaType::Text(None),
            ctx,
        ),
        "VALUE" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Text],
            FormulaType::Number,
            ctx,
        ),
        "NUMBERVALUE" => infer_signature(
            &upper,
            args,
            1,
            Some(3),
            &[ExpectedType::Text],
            FormulaType::Number,
            ctx,
        ),
        "ROW" | "COLUMN" => {
            infer_reference_function(&upper, args, 0, Some(1), FormulaType::Number, ctx)
        }
        "ROWS" | "COLUMNS" | "AREAS" => {
            infer_reference_function(&upper, args, 1, Some(1), FormulaType::Number, ctx)
        }
        "ISREF" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Any],
            FormulaType::Boolean,
            ctx,
        ),
        "FORMULATEXT" => {
            infer_reference_function(&upper, args, 1, Some(1), FormulaType::Text(None), ctx)
        }
        "ISFORMULA" => {
            infer_reference_function(&upper, args, 1, Some(1), FormulaType::Boolean, ctx)
        }
        "SHEET" => infer_signature(
            &upper,
            args,
            0,
            Some(1),
            &[ExpectedType::Any],
            FormulaType::Number,
            ctx,
        ),
        "SHEETS" => infer_reference_function(&upper, args, 0, Some(1), FormulaType::Number, ctx),
        "INDEX" => infer_index(args, ctx),
        "MATCH" => infer_signature(
            &upper,
            args,
            2,
            Some(3),
            &[ExpectedType::Any, ExpectedType::Array, ExpectedType::Number],
            FormulaType::Number,
            ctx,
        ),
        "XMATCH" => infer_signature(
            &upper,
            args,
            2,
            Some(4),
            &[
                ExpectedType::Any,
                ExpectedType::Array,
                ExpectedType::Number,
                ExpectedType::Number,
            ],
            FormulaType::Number,
            ctx,
        ),
        "VLOOKUP" | "HLOOKUP" => infer_signature(
            &upper,
            args,
            3,
            Some(4),
            &[
                ExpectedType::Any,
                ExpectedType::Array,
                ExpectedType::Number,
                ExpectedType::Boolean,
            ],
            FormulaType::Any,
            ctx,
        ),
        "XLOOKUP" => infer_signature(
            &upper,
            args,
            3,
            Some(6),
            &[
                ExpectedType::Any,
                ExpectedType::Array,
                ExpectedType::Array,
                ExpectedType::Any,
                ExpectedType::Number,
                ExpectedType::Number,
            ],
            FormulaType::Any,
            ctx,
        ),
        "FILTER" => infer_signature(
            &upper,
            args,
            2,
            Some(3),
            &[ExpectedType::Array, ExpectedType::Array, ExpectedType::Any],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "SORT" => infer_signature(
            &upper,
            args,
            1,
            Some(4),
            &[
                ExpectedType::Array,
                ExpectedType::Number,
                ExpectedType::Number,
                ExpectedType::Boolean,
            ],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "UNIQUE" => infer_signature(
            &upper,
            args,
            1,
            Some(3),
            &[
                ExpectedType::Array,
                ExpectedType::Boolean,
                ExpectedType::Boolean,
            ],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "TRANSPOSE" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Array],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "TOCOL" | "TOROW" => infer_signature(
            &upper,
            args,
            1,
            Some(3),
            &[
                ExpectedType::Array,
                ExpectedType::Number,
                ExpectedType::Boolean,
            ],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "TAKE" | "DROP" => infer_signature(
            &upper,
            args,
            2,
            Some(3),
            &[ExpectedType::Array, ExpectedType::Number],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "CHOOSECOLS" | "CHOOSEROWS" => infer_signature(
            &upper,
            args,
            2,
            None,
            &[ExpectedType::Array, ExpectedType::Number],
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            },
            ctx,
        ),
        "HSTACK" | "VSTACK" => {
            check_arity(ctx, &upper, args.len(), 1, None);
            for arg in args {
                let ty = infer_expr_type(arg, ctx);
                check_expected(ctx, &ty, ExpectedType::Array, &upper);
            }
            FormulaType::Array {
                item: Box::new(FormulaType::Any),
                rows: None,
                cols: None,
            }
        }
        "ISERROR" | "ISNA" | "ISERR" | "ISBLANK" | "ISTEXT" | "ISNONTEXT" | "ISNUMBER"
        | "ISLOGICAL" | "ISEVEN" | "ISODD" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Any],
            FormulaType::Boolean,
            ctx,
        ),
        "TYPE" | "ERROR.TYPE" => infer_signature(
            &upper,
            args,
            1,
            Some(1),
            &[ExpectedType::Any],
            FormulaType::Number,
            ctx,
        ),
        _ => {
            for arg in args {
                let _ = infer_expr_type(arg, ctx);
            }
            FormulaType::Any
        }
    }
}

fn infer_reference_function(
    name: &str,
    args: &[Expr],
    min: usize,
    max: Option<usize>,
    return_type: FormulaType,
    ctx: &mut TypeContext<'_>,
) -> FormulaType {
    check_arity(ctx, name, args.len(), min, max);
    for arg in args {
        if is_reference_expr(arg) {
            let _ = infer_expr_type(arg, ctx);
        } else {
            let ty = infer_expr_type(arg, ctx);
            check_expected(ctx, &ty, ExpectedType::Reference, name);
        }
    }
    return_type
}

fn infer_signature(
    name: &str,
    args: &[Expr],
    min: usize,
    max: Option<usize>,
    expected: &[ExpectedType],
    return_type: FormulaType,
    ctx: &mut TypeContext<'_>,
) -> FormulaType {
    check_arity(ctx, name, args.len(), min, max);
    for (idx, arg) in args.iter().enumerate() {
        let ty = infer_expr_type(arg, ctx);
        let expected = expected
            .get(idx)
            .copied()
            .or_else(|| expected.last().copied())
            .unwrap_or(ExpectedType::Any);
        check_expected(ctx, &ty, expected, name);
    }
    return_type
}

fn infer_if(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    check_arity(ctx, "IF", args.len(), 2, Some(3));
    if let Some(cond) = args.first() {
        let cond_ty = infer_expr_type(cond, ctx);
        check_expected(ctx, &cond_ty, ExpectedType::Boolean, "IF condition");
    }
    let then_ty = args
        .get(1)
        .map(|expr| infer_expr_type(expr, ctx))
        .unwrap_or(FormulaType::Blank);
    let else_ty = args
        .get(2)
        .map(|expr| infer_expr_type(expr, ctx))
        .unwrap_or(FormulaType::Blank);
    merge_types(then_ty, else_ty)
}

fn infer_fixed_lazy(
    name: &str,
    args: &[Expr],
    count: usize,
    ctx: &mut TypeContext<'_>,
) -> FormulaType {
    check_arity(ctx, name, args.len(), count, Some(count));
    let mut ty = FormulaType::Blank;
    for arg in args {
        ty = merge_types(ty, infer_expr_type(arg, ctx));
    }
    ty
}

fn infer_ifs(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Function,
            "IFS requires condition/value pairs".to_string(),
        );
    }
    let mut out = FormulaType::Blank;
    for pair in args.chunks(2) {
        if let Some(cond) = pair.first() {
            let cond_ty = infer_expr_type(cond, ctx);
            check_expected(ctx, &cond_ty, ExpectedType::Boolean, "IFS condition");
        }
        if let Some(value) = pair.get(1) {
            out = merge_types(out, infer_expr_type(value, ctx));
        }
    }
    out
}

fn infer_switch(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    check_arity(ctx, "SWITCH", args.len(), 3, None);
    let mut out = FormulaType::Blank;
    for (idx, arg) in args.iter().enumerate() {
        let ty = infer_expr_type(arg, ctx);
        if idx >= 2
            && (idx.is_multiple_of(2) || (args.len().is_multiple_of(2) && idx == args.len() - 1))
        {
            out = merge_types(out, ty);
        }
    }
    out
}

fn infer_choose(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    check_arity(ctx, "CHOOSE", args.len(), 2, None);
    if let Some(index) = args.first() {
        let ty = infer_expr_type(index, ctx);
        check_expected(ctx, &ty, ExpectedType::Number, "CHOOSE index");
    }
    let mut out = FormulaType::Blank;
    for arg in args.iter().skip(1) {
        out = merge_types(out, infer_expr_type(arg, ctx));
    }
    out
}

fn infer_let(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Function,
            "LET requires name/value pairs followed by a result expression".to_string(),
        );
    }
    let saved = ctx.bindings.clone();
    let mut i = 0;
    while i + 1 < args.len().saturating_sub(1) {
        if let Expr::Identifier(name) = &args[i] {
            let value_ty = infer_expr_type(&args[i + 1], ctx);
            ctx.bindings.insert(name.to_ascii_lowercase(), value_ty);
        } else {
            let _ = infer_expr_type(&args[i], ctx);
            let _ = infer_expr_type(&args[i + 1], ctx);
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Function,
                "LET binding name must be an identifier".to_string(),
            );
        }
        i += 2;
    }
    let out = args
        .last()
        .map(|expr| infer_expr_type(expr, ctx))
        .unwrap_or(FormulaType::Any);
    ctx.bindings = saved;
    out
}

fn infer_lambda(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    check_arity(ctx, "LAMBDA", args.len(), 2, None);
    let saved = ctx.bindings.clone();
    if args.len() >= 2 {
        for arg in &args[..args.len() - 1] {
            if let Expr::Identifier(name) = arg {
                ctx.bindings
                    .insert(name.to_ascii_lowercase(), FormulaType::Any);
            } else {
                let _ = infer_expr_type(arg, ctx);
                push(
                    ctx,
                    FormulaDiagnosticSeverity::Error,
                    FormulaDiagnosticKind::Function,
                    "LAMBDA parameter must be an identifier".to_string(),
                );
            }
        }
    }
    if let Some(body) = args.last() {
        let _ = infer_expr_type(body, ctx);
    }
    ctx.bindings = saved;
    FormulaType::Any
}

fn infer_index(args: &[Expr], ctx: &mut TypeContext<'_>) -> FormulaType {
    check_arity(ctx, "INDEX", args.len(), 2, Some(4));
    let array_ty = args
        .first()
        .map(|expr| infer_expr_type(expr, ctx))
        .unwrap_or(FormulaType::Any);
    check_expected(ctx, &array_ty, ExpectedType::Array, "INDEX array");
    for arg in args.iter().skip(1) {
        let ty = infer_expr_type(arg, ctx);
        check_expected(ctx, &ty, ExpectedType::Number, "INDEX position");
    }
    match array_ty {
        FormulaType::Array { item, .. } => *item,
        other => scalarize(other),
    }
}

fn check_arity(
    ctx: &mut TypeContext<'_>,
    name: &str,
    actual: usize,
    min: usize,
    max: Option<usize>,
) {
    if actual < min || max.is_some_and(|m| actual > m) {
        let message = match max {
            Some(max) if max == min => format!("{name} requires {min} argument(s), got {actual}"),
            Some(max) => {
                format!("{name} requires between {min} and {max} argument(s), got {actual}")
            }
            None => format!("{name} requires at least {min} argument(s), got {actual}"),
        };
        push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Function,
            message,
        );
    }
}

fn check_expected(
    ctx: &mut TypeContext<'_>,
    actual: &FormulaType,
    expected: ExpectedType,
    label: &str,
) {
    if expected == ExpectedType::Any || matches!(actual, FormulaType::Any | FormulaType::Error) {
        return;
    }
    match expected {
        ExpectedType::Any => {}
        ExpectedType::Number => check_number(ctx, actual, label),
        ExpectedType::Text => check_text(ctx, actual, label),
        ExpectedType::Boolean => check_boolean(ctx, actual, label),
        ExpectedType::Reference => {
            push(
                ctx,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Type,
                format!("{label} expects a reference, got {}", describe_type(actual)),
            );
        }
        ExpectedType::Array => {
            if !matches!(actual, FormulaType::Array { .. }) {
                push(
                    ctx,
                    FormulaDiagnosticSeverity::Error,
                    FormulaDiagnosticKind::Type,
                    format!(
                        "{label} expects an array or range, got {}",
                        describe_type(actual)
                    ),
                );
            }
        }
    }
}

fn check_number(ctx: &mut TypeContext<'_>, actual: &FormulaType, label: &str) {
    match actual {
        FormulaType::Number | FormulaType::Blank | FormulaType::Boolean => {}
        FormulaType::Text(Some(s)) if s.trim().parse::<f64>().is_ok() => push(
            ctx,
            FormulaDiagnosticSeverity::Warning,
            FormulaDiagnosticKind::Type,
            format!("{label} uses numeric text as a number"),
        ),
        FormulaType::Text(_) => push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            format!("{label} expects a number, got text"),
        ),
        FormulaType::Array { .. } => push(
            ctx,
            FormulaDiagnosticSeverity::Warning,
            FormulaDiagnosticKind::Type,
            format!("{label} expects a scalar number, got array"),
        ),
        FormulaType::Any | FormulaType::Error => {}
    }
}

fn check_text(ctx: &mut TypeContext<'_>, actual: &FormulaType, label: &str) {
    match actual {
        FormulaType::Text(_) | FormulaType::Blank | FormulaType::Number | FormulaType::Boolean => {}
        FormulaType::Array { .. } => push(
            ctx,
            FormulaDiagnosticSeverity::Warning,
            FormulaDiagnosticKind::Type,
            format!("{label} expects scalar text, got array"),
        ),
        FormulaType::Any | FormulaType::Error => {}
    }
}

fn check_boolean(ctx: &mut TypeContext<'_>, actual: &FormulaType, label: &str) {
    match actual {
        FormulaType::Boolean | FormulaType::Blank | FormulaType::Number => {}
        FormulaType::Text(Some(s))
            if s.eq_ignore_ascii_case("true") || s.eq_ignore_ascii_case("false") =>
        {
            push(
                ctx,
                FormulaDiagnosticSeverity::Warning,
                FormulaDiagnosticKind::Type,
                format!("{label} uses logical text as a boolean"),
            );
        }
        FormulaType::Text(_) => push(
            ctx,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            format!("{label} expects a boolean, got text"),
        ),
        FormulaType::Array { .. } => push(
            ctx,
            FormulaDiagnosticSeverity::Warning,
            FormulaDiagnosticKind::Type,
            format!("{label} expects a scalar boolean, got array"),
        ),
        FormulaType::Any | FormulaType::Error => {}
    }
}

fn check_scalar(ctx: &mut TypeContext<'_>, actual: &FormulaType, label: &str) {
    if matches!(actual, FormulaType::Array { .. }) {
        push(
            ctx,
            FormulaDiagnosticSeverity::Warning,
            FormulaDiagnosticKind::Type,
            format!("{label} expects a scalar, got array"),
        );
    }
}

fn scalarize(ty: FormulaType) -> FormulaType {
    match ty {
        FormulaType::Array { item, .. } => *item,
        other => other,
    }
}

fn result_type(value: &ResultData) -> FormulaType {
    match value {
        ResultData::None => FormulaType::Blank,
        ResultData::Boolean(_) => FormulaType::Boolean,
        ResultData::Integer(_) | ResultData::Float(_) => FormulaType::Number,
        ResultData::String(s) => FormulaType::Text(Some(s.clone())),
        ResultData::List(values) => {
            let mut item = FormulaType::Blank;
            for value in values {
                item = merge_types(item, result_type(value));
            }
            FormulaType::Array {
                item: Box::new(item),
                rows: None,
                cols: Some(values.len()),
            }
        }
        ResultData::Dict(_) => FormulaType::Any,
        ResultData::Error(_) => FormulaType::Error,
    }
}

fn merge_types(left: FormulaType, right: FormulaType) -> FormulaType {
    if left == right {
        return left;
    }
    match (left, right) {
        (FormulaType::Blank, other) | (other, FormulaType::Blank) => other,
        (FormulaType::Any, _) | (_, FormulaType::Any) => FormulaType::Any,
        (FormulaType::Error, other) | (other, FormulaType::Error) => other,
        (FormulaType::Text(_), FormulaType::Text(_)) => FormulaType::Text(None),
        (
            FormulaType::Array {
                item: l_item,
                rows: l_rows,
                cols: l_cols,
            },
            FormulaType::Array {
                item: r_item,
                rows: r_rows,
                cols: r_cols,
            },
        ) => FormulaType::Array {
            item: Box::new(merge_types(*l_item, *r_item)),
            rows: if l_rows == r_rows { l_rows } else { None },
            cols: if l_cols == r_cols { l_cols } else { None },
        },
        _ => FormulaType::Any,
    }
}

fn resolve_sheet<'a>(
    sheet_name: Option<&str>,
    current_sheet: &'a Sheet,
    workbook_context: Option<&'a Context<'a>>,
) -> Result<&'a Sheet, String> {
    match sheet_name {
        None => Ok(current_sheet),
        Some(name) if name == current_sheet.name => Ok(current_sheet),
        Some(name) => workbook_context
            .and_then(|ctx| ctx.sheets.get(name).copied())
            .ok_or_else(|| name.to_string()),
    }
}

fn is_reference_expr(expr: &Expr) -> bool {
    matches!(
        expr,
        Expr::CellRef { .. } | Expr::RangeRef { .. } | Expr::StructuredRef { .. }
    ) || matches!(
        expr,
        Expr::UnaryOp {
            op: Op::Spill | Op::ImplicitIntersection,
            ..
        }
    ) || matches!(
        expr,
        Expr::BinaryOp {
            op: Op::Intersect | Op::Union,
            ..
        }
    )
}

fn describe_type(ty: &FormulaType) -> &'static str {
    match ty {
        FormulaType::Any => "unknown",
        FormulaType::Blank => "blank",
        FormulaType::Number => "number",
        FormulaType::Text(_) => "text",
        FormulaType::Boolean => "boolean",
        FormulaType::Error => "error",
        FormulaType::Array { .. } => "array",
    }
}

fn push(
    ctx: &mut TypeContext<'_>,
    severity: FormulaDiagnosticSeverity,
    kind: FormulaDiagnosticKind,
    message: String,
) {
    ctx.diagnostics.push(FormulaDiagnostic {
        severity,
        kind,
        message,
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::{Context, Sheet, SheetInit};

    fn sheet() -> Sheet {
        Sheet::new(SheetInit {
            rows: 10,
            cols: 10,
            ..Default::default()
        })
    }

    fn messages(diags: &[FormulaDiagnostic]) -> Vec<String> {
        diags.iter().map(|d| d.message.clone()).collect()
    }

    fn has_diag(
        diags: &[FormulaDiagnostic],
        severity: FormulaDiagnosticSeverity,
        kind: FormulaDiagnosticKind,
        message: &str,
    ) -> bool {
        diags
            .iter()
            .any(|d| d.severity == severity && d.kind == kind && d.message.contains(message))
    }

    fn inferred_type(input: &str, sheet: &Sheet) -> (FormulaType, Vec<FormulaDiagnostic>) {
        let ast = parse_excel_formula(input.strip_prefix('=').unwrap_or(input)).unwrap();
        let mut ctx = TypeContext {
            sheet,
            workbook_context: None,
            row: None,
            bindings: HashMap::new(),
            diagnostics: Vec::new(),
        };
        let ty = infer_expr_type(&ast, &mut ctx);
        (ty, ctx.diagnostics)
    }

    #[test]
    fn non_formula_inputs_are_ignored() {
        let sheet = sheet();
        assert_eq!(check_formula("1+2", &sheet, None, None, None), Vec::new());
        assert_eq!(check_formula("=   ", &sheet, None, None, None), Vec::new());
    }

    #[test]
    fn syntax_errors_report_syntax_diagnostics() {
        let sheet = sheet();
        let diags = check_formula("=SUM(", &sheet, None, None, None);
        assert!(diags.iter().any(|d| {
            d.severity == FormulaDiagnosticSeverity::Error
                && d.kind == FormulaDiagnosticKind::Syntax
        }));
    }

    #[test]
    fn valid_arithmetic_has_no_diagnostics() {
        let sheet = sheet();
        let diags = check_formula("=1+2", &sheet, None, None, None);
        assert_eq!(diags, Vec::new());
    }

    #[test]
    fn invalid_arithmetic_reports_type_error() {
        let sheet = sheet();
        let diags = check_formula("=\"abc\"+1", &sheet, None, None, None);
        assert!(
            messages(&diags)
                .iter()
                .any(|m| m.contains("expects a number"))
        );
    }

    #[test]
    fn numeric_text_reports_coercion_warning() {
        let sheet = sheet();
        let diags = check_formula("=\"1\"+1", &sheet, None, None, None);
        assert!(
            diags
                .iter()
                .any(|d| d.severity == FormulaDiagnosticSeverity::Warning)
        );
    }

    #[test]
    fn bad_arity_reports_function_error() {
        let sheet = sheet();
        let diags = check_formula("=LEFT()", &sheet, None, None, None);
        assert!(messages(&diags).iter().any(|m| m.contains("LEFT requires")));
    }

    #[test]
    fn unknown_function_reports_error() {
        let sheet = sheet();
        let diags = check_formula("=NO_SUCH_FN(1)", &sheet, None, None, None);
        assert!(
            messages(&diags)
                .iter()
                .any(|m| m.contains("Unknown function"))
        );
    }

    #[test]
    fn unknown_sheet_reports_reference_error() {
        let sheet = sheet();
        let diags = check_formula("=Missing!A1", &sheet, None, None, None);
        assert!(
            messages(&diags)
                .iter()
                .any(|m| m.contains("Sheet not found"))
        );
    }

    #[test]
    fn range_to_scalar_function_reports_array_mismatch() {
        let sheet = sheet();
        let diags = check_formula("=LEFT(A1:A2)", &sheet, None, None, None);
        assert!(messages(&diags).iter().any(|m| m.contains("got array")));
    }

    #[test]
    fn let_binding_is_available_to_result_expression() {
        let sheet = sheet();
        let diags = check_formula("=LET(x, 1, x + 2)", &sheet, None, None, None);
        assert_eq!(diags, Vec::new());
    }

    #[test]
    fn cell_reference_types_are_checked_from_evaluated_results() {
        let mut sheet = sheet();
        sheet.set_cell_src(0, 0, "not a number".to_string());
        sheet.commit(None).unwrap();
        let diags = check_formula("=A1+1", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "left operand expects a number"
        ));
    }

    #[test]
    fn cross_sheet_references_use_workbook_context() {
        let sheet1 = Sheet::new(SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 10,
            cols: 10,
            ..Default::default()
        });
        let mut sheet2 = Sheet::new(SheetInit {
            name: Some("Sheet2".to_string()),
            rows: 10,
            cols: 10,
            ..Default::default()
        });
        sheet2.set_cell_src(0, 0, "text".to_string());
        sheet2.commit(None).unwrap();
        let mut context = Context::new();
        context.add_table(sheet2.name.clone(), &sheet2);
        let diags = check_formula("=Sheet2!A1+1", &sheet1, Some(&context), None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "left operand expects a number"
        ));
    }

    #[test]
    fn aggregate_functions_accept_range_arguments() {
        let sheet = sheet();
        for formula in [
            "=SUM(A1:A3)",
            "=AVERAGE(A1:A3)",
            "=MIN(A1:A3)",
            "=MAX(A1:A3)",
            "=PRODUCT(A1:A3)",
            "=SUMSQ(A1:A3)",
        ] {
            assert_eq!(
                check_formula(formula, &sheet, None, None, None),
                Vec::new(),
                "{formula}"
            );
        }

        let diags = check_formula("=SUM(\"not numeric\")", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "SUM expects a number, got text"
        ));
    }

    #[test]
    fn lookup_functions_require_array_arguments() {
        let sheet = sheet();
        let diags = check_formula("=MATCH(1, 2, 0)", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "MATCH expects an array or range"
        ));
    }

    #[test]
    fn reference_functions_reject_scalar_arguments() {
        let sheet = sheet();
        let diags = check_formula("=ROW(1)", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "ROW expects a reference"
        ));
    }

    #[test]
    fn merged_text_branches_remain_text() {
        let sheet = sheet();
        let (ty, diags) = inferred_type("=IF(TRUE, \"left\", \"right\")", &sheet);
        assert_eq!(ty, FormulaType::Text(None));
        assert_eq!(diags, Vec::new());

        let diags = check_formula(
            "=IF(TRUE, \"left\", \"right\") + 1",
            &sheet,
            None,
            None,
            None,
        );
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "left operand expects a number, got text"
        ));
    }

    #[test]
    fn reference_function_signatures_and_return_types_are_distinct() {
        let sheet = sheet();
        assert_eq!(
            inferred_type("=FORMULATEXT(A1)", &sheet).0,
            FormulaType::Text(None)
        );
        assert_eq!(
            inferred_type("=ISFORMULA(A1)", &sheet).0,
            FormulaType::Boolean
        );
        assert_eq!(inferred_type("=ISREF(1)", &sheet).0, FormulaType::Boolean);
        assert_eq!(
            check_formula("=ISREF(1)", &sheet, None, None, None),
            Vec::new()
        );

        for formula in ["=ROWS()", "=AREAS()", "=ISREF()", "=FORMULATEXT()"] {
            let diags = check_formula(formula, &sheet, None, None, None);
            assert!(has_diag(
                &diags,
                FormulaDiagnosticSeverity::Error,
                FormulaDiagnosticKind::Function,
                "requires 1 argument(s)"
            ));
        }

        let diags = check_formula("=FORMULATEXT(1)", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "FORMULATEXT expects a reference"
        ));
    }

    #[test]
    fn textjoin_checks_its_boolean_argument_and_minimum_arity() {
        let sheet = sheet();
        assert_eq!(
            check_formula(
                "=TEXTJOIN(\",\", TRUE, \"a\", \"b\")",
                &sheet,
                None,
                None,
                None,
            ),
            Vec::new()
        );

        let diags = check_formula("=TEXTJOIN(\",\", TRUE)", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Function,
            "requires at least 3 argument(s)"
        ));

        let diags = check_formula(
            "=TEXTJOIN(\",\", \"sometimes\", \"a\")",
            &sheet,
            None,
            None,
            None,
        );
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Type,
            "TEXTJOIN expects a boolean, got text"
        ));
    }

    #[test]
    fn concat_accepts_ranges() {
        let sheet = sheet();
        assert_eq!(
            check_formula("=CONCAT(A1:A3)", &sheet, None, None, None),
            Vec::new()
        );
        assert_eq!(
            check_formula("=CONCATENATE(\"value: \", 1)", &sheet, None, None, None),
            Vec::new()
        );
    }

    #[test]
    fn dynamic_array_optional_arguments_use_their_own_types() {
        let sheet = sheet();
        for formula in [
            "=SORT(A1:A3, 1, -1, FALSE)",
            "=UNIQUE(A1:A3, FALSE, TRUE)",
            "=TOCOL(A1:A3, 1, TRUE)",
            "=TOROW(A1:A3, 2, FALSE)",
        ] {
            assert_eq!(
                check_formula(formula, &sheet, None, None, None),
                Vec::new(),
                "{formula}"
            );
        }

        for formula in ["=TRANSPOSE(A1:A3, A1:A3)", "=TAKE(A1:A3, 1, 1, 1)"] {
            let diags = check_formula(formula, &sheet, None, None, None);
            assert!(
                diags
                    .iter()
                    .any(|d| d.kind == FormulaDiagnosticKind::Function),
                "{formula}: {diags:?}"
            );
        }
    }

    #[test]
    fn similar_functions_keep_separate_arity_rules() {
        let sheet = sheet();
        for formula in ["=VALUE(\"1\", \".\")", "=MATCH(1, A1:A3, 0, 1)"] {
            let diags = check_formula(formula, &sheet, None, None, None);
            assert!(
                diags
                    .iter()
                    .any(|d| d.kind == FormulaDiagnosticKind::Function),
                "{formula}: {diags:?}"
            );
        }
        assert_eq!(
            check_formula(
                "=NUMBERVALUE(\"1,000.5\", \".\", \",\")",
                &sheet,
                None,
                None,
                None,
            ),
            Vec::new()
        );
        assert_eq!(
            check_formula("=XMATCH(1, A1:A3, 0, 1)", &sheet, None, None, None),
            Vec::new()
        );
    }

    #[test]
    fn switch_default_expression_is_checked_once() {
        let sheet = sheet();
        let diags = check_formula("=SWITCH(1, 1, 2, missing)", &sheet, None, None, None);
        assert_eq!(
            diags
                .iter()
                .filter(|d| d.message.contains("Unknown name: missing"))
                .count(),
            1
        );
    }

    #[test]
    fn structured_references_validate_tables_columns_and_row_context() {
        let mut sheet = sheet();
        sheet.set_cell_src(0, 0, "Name".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 1, true, false)
            .unwrap();

        assert_eq!(
            inferred_type("=Sales[[#Headers],[Name]]", &sheet).0,
            FormulaType::Text(None)
        );

        let diags = check_formula("=Sales[Missing]", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Reference,
            "Column not found: Missing"
        ));

        let diags = check_formula("=Sales[@Name]", &sheet, None, None, None);
        assert!(has_diag(
            &diags,
            FormulaDiagnosticSeverity::Error,
            FormulaDiagnosticKind::Reference,
            "without row context"
        ));
    }
}
