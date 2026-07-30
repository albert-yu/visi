use crate::core::{CompiledFormula, FormulaPart, RefType, Sheet, SheetSection, generate_unique_id};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, serde::Serialize, serde::Deserialize)]
pub enum Op {
    Add,
    Sub,
    Mul,
    Div,
    Exp, // ^
    Eq,  // =
    Ne,  // <> or !=
    Lt,  // <
    Gt,  // >
    Le,  // <=
    Ge,  // >=
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    String(String),
    Boolean(bool),
    CellRef {
        sheet: Option<String>,
        row: usize,
        col: usize,
        row_abs: bool,
        col_abs: bool,
    },
    RangeRef {
        sheet: Option<String>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        start_row_abs: bool,
        start_col_abs: bool,
        end_row_abs: bool,
        end_col_abs: bool,
    },
    List(Vec<Expr>),
    FunctionCall {
        name: String,
        args: Vec<Expr>,
    },
    BinaryOp {
        op: Op,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    UnaryOp {
        op: Op,
        expr: Box<Expr>,
    },
    Identifier(String),
    StructuredRef {
        sheet: Option<String>,
        column: Option<String>,
        is_this_row: bool,
        section: SheetSection,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub enum EvalToken {
    Number(f64),
    String(String),
    Boolean(bool),
    Identifier(String),
    Op(Op),
    OpenParen,
    CloseParen,
    OpenBracket,
    CloseBracket,
    Comma,
    Colon,
    Exclamation,
    Dot,
    StructuredRef {
        sheet: Option<String>,
        column: Option<String>,
        is_this_row: bool,
        section: SheetSection,
    },
}

pub fn col_idx_to_letters(mut col: usize) -> String {
    let mut letters = String::new();
    loop {
        let remainder = col % 26;
        letters.insert(0, (b'A' + remainder as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    letters
}

pub fn parse_a1_coordinates(col_str: &str, row_str: &str) -> (usize, usize) {
    let mut col = 0;
    for c in col_str.chars() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
        }
    }
    let col_idx = if col > 0 { col - 1 } else { 0 };

    let row_val: usize = row_str.parse().unwrap_or(1);
    let row_idx = if row_val > 0 { row_val - 1 } else { 0 };

    (row_idx, col_idx)
}

fn parse_cell_ref(s: &str) -> Option<(usize, usize, bool, bool)> {
    let chars: Vec<char> = s.chars().collect();
    let mut idx = 0;

    let mut col_abs = false;
    if idx < chars.len() && chars[idx] == '$' {
        col_abs = true;
        idx += 1;
    }

    let mut col_str = String::new();
    while idx < chars.len() && chars[idx].is_ascii_alphabetic() {
        col_str.push(chars[idx]);
        idx += 1;
    }
    if col_str.is_empty() || col_str.len() > 3 {
        return None;
    }

    let mut row_abs = false;
    if idx < chars.len() && chars[idx] == '$' {
        row_abs = true;
        idx += 1;
    }

    let mut row_str = String::new();
    while idx < chars.len() && chars[idx].is_ascii_digit() {
        row_str.push(chars[idx]);
        idx += 1;
    }
    if row_str.is_empty() {
        return None;
    }

    if idx < chars.len() {
        return None;
    }

    let (row_idx, col_idx) = parse_a1_coordinates(&col_str, &row_str);
    Some((row_idx, col_idx, row_abs, col_abs))
}

fn parse_column_ref(s: &str) -> Option<(usize, bool)> {
    let chars: Vec<char> = s.chars().collect();
    let mut idx = 0;
    let mut col_abs = false;
    if idx < chars.len() && chars[idx] == '$' {
        col_abs = true;
        idx += 1;
    }
    let mut col_str = String::new();
    while idx < chars.len() && chars[idx].is_ascii_alphabetic() {
        col_str.push(chars[idx]);
        idx += 1;
    }
    if col_str.is_empty() || col_str.len() > 3 {
        return None;
    }
    if idx < chars.len() {
        return None;
    }
    let mut col = 0;
    for c in col_str.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    let col_idx = if col > 0 { col - 1 } else { 0 };
    Some((col_idx, col_abs))
}

fn parse_cell_pattern(chars: &[char], mut idx: usize) -> Option<(usize, usize, bool, bool, usize)> {
    let mut col_abs = false;
    if idx < chars.len() && chars[idx] == '$' {
        col_abs = true;
        idx += 1;
    }

    let mut col_str = String::new();
    while idx < chars.len() && chars[idx].is_ascii_alphabetic() {
        col_str.push(chars[idx]);
        idx += 1;
    }
    if col_str.is_empty() || col_str.len() > 3 {
        return None;
    }

    let mut row_abs = false;
    if idx < chars.len() && chars[idx] == '$' {
        row_abs = true;
        idx += 1;
    }

    let mut row_str = String::new();
    while idx < chars.len() && chars[idx].is_ascii_digit() {
        row_str.push(chars[idx]);
        idx += 1;
    }
    if row_str.is_empty() {
        return None;
    }

    if idx < chars.len() && (chars[idx].is_ascii_alphabetic() || chars[idx] == '_') {
        return None;
    }

    let (row_idx, col_idx) = parse_a1_coordinates(&col_str, &row_str);
    Some((row_idx, col_idx, row_abs, col_abs, idx))
}

fn parse_col_pattern(chars: &[char], mut idx: usize) -> Option<(usize, bool, usize)> {
    let mut col_abs = false;
    if idx < chars.len() && chars[idx] == '$' {
        col_abs = true;
        idx += 1;
    }
    let mut col_str = String::new();
    while idx < chars.len() && chars[idx].is_ascii_alphabetic() {
        col_str.push(chars[idx]);
        idx += 1;
    }
    if col_str.is_empty() || col_str.len() > 3 {
        return None;
    }
    if idx < chars.len() && chars[idx].is_ascii_digit() {
        return None;
    }
    if idx < chars.len() && (chars[idx].is_ascii_alphabetic() || chars[idx] == '_') {
        return None;
    }
    let mut col = 0;
    for c in col_str.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    let col_idx = if col > 0 { col - 1 } else { 0 };
    Some((col_idx, col_abs, idx))
}

#[derive(Debug, Clone)]
enum FoundRef {
    Cell {
        sheet: Option<String>,
        row: usize,
        col: usize,
        row_abs: bool,
        col_abs: bool,
    },
    Range {
        sheet: Option<String>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        start_row_abs: bool,
        start_col_abs: bool,
        end_row_abs: bool,
        end_col_abs: bool,
    },
    Structured {
        sheet: Option<String>,
        column: Option<String>,
        is_this_row: bool,
        section: SheetSection,
    },
}

fn parse_bracketed_term(chars: &[char], mut idx: usize) -> Option<(String, usize)> {
    if idx >= chars.len() || chars[idx] != '[' {
        return None;
    }
    idx += 1; // consume '['
    while idx < chars.len() && chars[idx].is_whitespace() {
        idx += 1;
    }
    let mut quote = None;
    if idx < chars.len() && (chars[idx] == '"' || chars[idx] == '\'') {
        quote = Some(chars[idx]);
        idx += 1;
    }
    let mut term = String::new();
    while idx < chars.len() {
        if let Some(q) = quote {
            if chars[idx] == q {
                idx += 1;
                break;
            }
        } else {
            if chars[idx] == ']' {
                break;
            }
            if chars[idx] == ','
                || chars[idx] == ':'
                || chars[idx] == '+'
                || chars[idx] == '-'
                || chars[idx] == '*'
                || chars[idx] == '/'
            {
                return None;
            }
        }
        term.push(chars[idx]);
        idx += 1;
    }
    while idx < chars.len() && chars[idx].is_whitespace() {
        idx += 1;
    }
    if idx < chars.len() && chars[idx] == ']' {
        Some((term.trim().to_string(), idx + 1))
    } else {
        None
    }
}

fn parse_structured_specifier(
    chars: &[char],
    mut idx: usize,
) -> Option<(Option<String>, bool, SheetSection, usize)> {
    if idx >= chars.len() || chars[idx] != '[' {
        return None;
    }
    idx += 1; // consume outer '['

    while idx < chars.len() && chars[idx].is_whitespace() {
        idx += 1;
    }

    let mut is_this_row = false;
    let mut section = SheetSection::Data;
    let mut column = None;

    if idx < chars.len() && chars[idx] == '@' {
        is_this_row = true;
        idx += 1;
        if idx < chars.len() && chars[idx] == '[' {
            if let Some((col_name, next_idx)) = parse_bracketed_term(chars, idx) {
                if !col_name.is_empty() {
                    column = Some(col_name);
                }
                idx = next_idx;
            } else {
                return None;
            }
        } else {
            let mut col_name = String::new();
            while idx < chars.len() && chars[idx] != ']' {
                col_name.push(chars[idx]);
                idx += 1;
            }
            let trimmed = col_name.trim().to_string();
            column = if trimmed.is_empty() {
                None
            } else {
                Some(trimmed)
            };
        }
        if idx < chars.len() && chars[idx] == ']' {
            return Some((column, is_this_row, section, idx + 1));
        } else {
            return None;
        }
    }

    if idx < chars.len() && chars[idx] == '[' {
        if let Some((term1, next_idx)) = parse_bracketed_term(chars, idx) {
            idx = next_idx;
            if term1.starts_with('#') {
                match term1.as_str() {
                    "#This Row" => is_this_row = true,
                    "#Headers" => section = SheetSection::Headers,
                    "#Totals" => section = SheetSection::Totals,
                    "#Data" => section = SheetSection::Data,
                    "#All" => section = SheetSection::All,
                    _ => {}
                }
                while idx < chars.len() && (chars[idx].is_whitespace() || chars[idx] == ',') {
                    idx += 1;
                }
                if idx < chars.len() && chars[idx] == '[' {
                    if let Some((col_name, next_idx2)) = parse_bracketed_term(chars, idx) {
                        if !col_name.is_empty() {
                            column = Some(col_name);
                        }
                        idx = next_idx2;
                    } else {
                        return None;
                    }
                }
            } else if !term1.is_empty() {
                column = Some(term1);
            }
        } else {
            return None;
        }
        while idx < chars.len() && chars[idx].is_whitespace() {
            idx += 1;
        }
        if idx < chars.len() && chars[idx] == ']' {
            return Some((column, is_this_row, section, idx + 1));
        } else {
            return None;
        }
    }

    let mut term = String::new();
    let mut quote = None;
    if idx < chars.len() && (chars[idx] == '"' || chars[idx] == '\'') {
        quote = Some(chars[idx]);
        idx += 1;
    }
    while idx < chars.len() {
        if let Some(q) = quote {
            if chars[idx] == q {
                idx += 1;
                break;
            }
        } else {
            if chars[idx] == ']' {
                break;
            }
            if chars[idx] == ','
                || chars[idx] == ':'
                || chars[idx] == '+'
                || chars[idx] == '-'
                || chars[idx] == '*'
                || chars[idx] == '/'
            {
                return None;
            }
        }
        term.push(chars[idx]);
        idx += 1;
    }
    term = term.trim().to_string();
    if term.starts_with('#') {
        match term.as_str() {
            "#This Row" => is_this_row = true,
            "#Headers" => section = SheetSection::Headers,
            "#Totals" => section = SheetSection::Totals,
            "#Data" => section = SheetSection::Data,
            "#All" => section = SheetSection::All,
            _ => {}
        }
    } else if !term.is_empty() {
        column = Some(term);
    }

    if idx < chars.len() && chars[idx] == ']' {
        return Some((column, is_this_row, section, idx + 1));
    }

    None
}

fn try_parse_ref(chars: &[char], start_idx: usize) -> Option<(FoundRef, usize)> {
    if start_idx > 0 {
        let prev = chars[start_idx - 1];
        if prev.is_ascii_alphanumeric() || prev == '_' {
            return None;
        }
    }
    let mut idx = start_idx;
    let mut sheet = None;

    let mut has_explicit_table = false;
    if idx < chars.len() && chars[idx] == '\'' {
        let mut t_name = String::new();
        let mut j = idx + 1;
        while j < chars.len() && chars[j] != '\'' {
            t_name.push(chars[j]);
            j += 1;
        }
        if j < chars.len() && chars[j] == '\'' {
            if j + 1 < chars.len() && chars[j + 1] == '!' {
                sheet = Some(t_name);
                idx = j + 2;
                has_explicit_table = true;
            } else if j + 1 < chars.len() && chars[j + 1] == '[' {
                sheet = Some(t_name);
                idx = j + 1;
                has_explicit_table = true;
            }
        }
    }

    if !has_explicit_table {
        let mut t_name = String::new();
        let mut j = idx;
        if j < chars.len() && (chars[j].is_ascii_alphabetic() || chars[j] == '_') {
            while j < chars.len()
                && (chars[j].is_ascii_alphanumeric() || chars[j] == '_' || chars[j] == '.')
            {
                t_name.push(chars[j]);
                j += 1;
            }
            if j < chars.len() && chars[j] == '!' {
                if t_name != "iloc" && t_name != "get" {
                    sheet = Some(t_name);
                    idx = j + 1;
                }
            } else if j < chars.len() && chars[j] == '[' {
                if t_name != "iloc" && t_name != "get" {
                    let mut k = j + 1;
                    let mut quote = None;
                    if k < chars.len() && (chars[k] == '"' || chars[k] == '\'') {
                        quote = Some(chars[k]);
                        k += 1;
                    }
                    let mut is_column = true;
                    let mut has_chars = false;
                    while k < chars.len() {
                        if let Some(q) = quote {
                            if chars[k] == q {
                                break;
                            }
                        } else {
                            if chars[k] == ']' {
                                break;
                            }
                            if chars[k] == ':'
                                || chars[k] == ','
                                || chars[k] == '+'
                                || chars[k] == '-'
                                || chars[k] == '*'
                                || chars[k] == '/'
                            {
                                is_column = false;
                                break;
                            }
                        }
                        has_chars = true;
                        k += 1;
                    }
                    if is_column && has_chars {
                        sheet = Some(t_name);
                        idx = j;
                    }
                }
            }
        }
    }

    if idx < chars.len() && chars[idx] == '[' {
        if let Some((column, is_this_row, section, next_idx)) =
            parse_structured_specifier(chars, idx)
        {
            return Some((
                FoundRef::Structured {
                    sheet,
                    column,
                    is_this_row,
                    section,
                },
                next_idx,
            ));
        }
    }

    // Try to parse column range ref like A:A or $A:$B
    if let Some((start_col, start_col_abs, next_idx)) = parse_col_pattern(chars, idx) {
        if next_idx < chars.len() && chars[next_idx] == ':' {
            if let Some((end_col, end_col_abs, end_idx)) = parse_col_pattern(chars, next_idx + 1) {
                return Some((
                    FoundRef::Range {
                        sheet: sheet.clone(),
                        start_row: 0,
                        start_col,
                        end_row: usize::MAX,
                        end_col,
                        start_row_abs: true,
                        start_col_abs,
                        end_row_abs: true,
                        end_col_abs,
                    },
                    end_idx,
                ));
            }
        }
    }

    if let Some((start_row, start_col, start_row_abs, start_col_abs, next_idx)) =
        parse_cell_pattern(chars, idx)
    {
        if next_idx < chars.len() && chars[next_idx] == ':' {
            if let Some((end_row, end_col, end_row_abs, end_col_abs, end_idx)) =
                parse_cell_pattern(chars, next_idx + 1)
            {
                return Some((
                    FoundRef::Range {
                        sheet,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        start_row_abs,
                        start_col_abs,
                        end_row_abs,
                        end_col_abs,
                    },
                    end_idx,
                ));
            }
        }

        return Some((
            FoundRef::Cell {
                sheet,
                row: start_row,
                col: start_col,
                row_abs: start_row_abs,
                col_abs: start_col_abs,
            },
            next_idx,
        ));
    }

    None
}

pub fn compile_formula(code: &str, sheets: &[Sheet]) -> CompiledFormula {
    let chars: Vec<char> = code.chars().collect();
    let mut parts = Vec::new();
    let mut last_idx = 0;

    let mut in_quote = None;
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if let Some(q) = in_quote {
            if c == '\\' {
                i += 2;
                continue;
            } else if c == q {
                in_quote = None;
            }
            i += 1;
            continue;
        }

        if let Some((found_ref, next_i)) = try_parse_ref(&chars, i) {
            if i > last_idx {
                let text: String = chars[last_idx..i].iter().collect();
                parts.push(FormulaPart::Text(text));
            }

            match found_ref {
                FoundRef::Cell {
                    sheet,
                    row,
                    col,
                    row_abs,
                    col_abs,
                } => {
                    let sheet_id = if let Some(ref name) = sheet {
                        if let Some(t) = sheets.iter().find(|t| t.name == *name) {
                            t.id
                        } else if !sheets.is_empty() {
                            sheets[0].id
                        } else {
                            0
                        }
                    } else if !sheets.is_empty() {
                        sheets[0].id
                    } else {
                        0
                    };

                    parts.push(FormulaPart::SheetReference {
                        sheet_id,
                        row,
                        col,
                        row_ref_type: if row_abs {
                            RefType::Absolute
                        } else {
                            RefType::Relative
                        },
                        col_ref_type: if col_abs {
                            RefType::Absolute
                        } else {
                            RefType::Relative
                        },
                    });
                }
                FoundRef::Range {
                    sheet,
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                    start_row_abs,
                    start_col_abs,
                    end_row_abs,
                    end_col_abs,
                } => {
                    let sheet_id = if let Some(ref name) = sheet {
                        if let Some(t) = sheets.iter().find(|t| t.name == *name) {
                            t.id
                        } else if !sheets.is_empty() {
                            sheets[0].id
                        } else {
                            0
                        }
                    } else if !sheets.is_empty() {
                        sheets[0].id
                    } else {
                        0
                    };

                    parts.push(FormulaPart::RangeReference {
                        sheet_id,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        start_row_ref_type: if start_row_abs {
                            RefType::Absolute
                        } else {
                            RefType::Relative
                        },
                        start_col_ref_type: if start_col_abs {
                            RefType::Absolute
                        } else {
                            RefType::Relative
                        },
                        end_row_ref_type: if end_row_abs {
                            RefType::Absolute
                        } else {
                            RefType::Relative
                        },
                        end_col_ref_type: if end_col_abs {
                            RefType::Absolute
                        } else {
                            RefType::Relative
                        },
                    });
                }
                FoundRef::Structured {
                    sheet,
                    column,
                    is_this_row,
                    section,
                } => {
                    let sheet_id = if let Some(ref name) = sheet {
                        if let Some(t) = sheets.iter().find(|t| t.name == *name) {
                            t.id
                        } else if !sheets.is_empty() {
                            sheets[0].id
                        } else {
                            0
                        }
                    } else if !sheets.is_empty() {
                        sheets[0].id
                    } else {
                        0
                    };

                    let col_id = if let Some(ref col_name) = column {
                        if let Some(table_obj) = sheets.iter().find(|t| t.id == sheet_id) {
                            if let Some(col) =
                                table_obj.columns.iter().find(|c| c.name == *col_name)
                            {
                                Some(col.id)
                            } else {
                                Some(generate_unique_id())
                            }
                        } else {
                            Some(generate_unique_id())
                        }
                    } else {
                        None
                    };

                    parts.push(FormulaPart::StructuredReference {
                        sheet_id,
                        col_id,
                        is_this_row,
                        section,
                    });
                }
            }

            i = next_i;
            last_idx = i;
        } else {
            if c == '"' || c == '\'' {
                in_quote = Some(c);
            }
            i += 1;
        }
    }

    if i > last_idx {
        let text: String = chars[last_idx..i].iter().collect();
        parts.push(FormulaPart::Text(text));
    }

    CompiledFormula { parts }
}

pub fn serialize_formula(formula: &CompiledFormula, sheets: &[Sheet]) -> String {
    let get_table_name = |sheet_id: u64, sheets: &[Sheet]| -> String {
        if let Some(sheet) = sheets.iter().find(|t| t.id == sheet_id) {
            if sheet.name.contains(' ') {
                format!("'{}'", sheet.name)
            } else {
                sheet.name.clone()
            }
        } else {
            "table_deleted".to_string()
        }
    };

    let mut result = String::new();
    for part in &formula.parts {
        match part {
            FormulaPart::Text(s) => result.push_str(s),
            FormulaPart::SheetReference {
                sheet_id,
                row,
                col,
                row_ref_type,
                col_ref_type,
            } => {
                let col_letter = col_idx_to_letters(*col);
                let r_prefix = match row_ref_type {
                    RefType::Absolute => "$",
                    RefType::Relative => "",
                };
                let c_prefix = match col_ref_type {
                    RefType::Absolute => "$",
                    RefType::Relative => "",
                };

                let has_prefix = if sheets.is_empty() {
                    false
                } else {
                    *sheet_id != sheets[0].id
                };

                if has_prefix {
                    let sheet_name = get_table_name(*sheet_id, sheets);
                    result.push_str(&format!(
                        "{}!{}{}{}{}",
                        sheet_name,
                        c_prefix,
                        col_letter,
                        r_prefix,
                        row + 1
                    ));
                } else {
                    result.push_str(&format!(
                        "{}{}{}{}",
                        c_prefix,
                        col_letter,
                        r_prefix,
                        row + 1
                    ));
                }
            }
            FormulaPart::RangeReference {
                sheet_id,
                start_row,
                start_col,
                end_row,
                end_col,
                start_row_ref_type,
                start_col_ref_type,
                end_row_ref_type,
                end_col_ref_type,
            } => {
                let start_col_letter = col_idx_to_letters(*start_col);
                let end_col_letter = col_idx_to_letters(*end_col);

                let sc_prefix = match start_col_ref_type {
                    RefType::Absolute => "$",
                    RefType::Relative => "",
                };
                let ec_prefix = match end_col_ref_type {
                    RefType::Absolute => "$",
                    RefType::Relative => "",
                };

                let has_prefix = if sheets.is_empty() {
                    false
                } else {
                    *sheet_id != sheets[0].id
                };

                if *end_row == usize::MAX {
                    if has_prefix {
                        let sheet_name = get_table_name(*sheet_id, sheets);
                        result.push_str(&format!(
                            "{}!{}{}:{}{}",
                            sheet_name, sc_prefix, start_col_letter, ec_prefix, end_col_letter,
                        ));
                    } else {
                        result.push_str(&format!(
                            "{}{}:{}{}",
                            sc_prefix, start_col_letter, ec_prefix, end_col_letter,
                        ));
                    }
                } else {
                    let sr_prefix = match start_row_ref_type {
                        RefType::Absolute => "$",
                        RefType::Relative => "",
                    };
                    let er_prefix = match end_row_ref_type {
                        RefType::Absolute => "$",
                        RefType::Relative => "",
                    };

                    if has_prefix {
                        let sheet_name = get_table_name(*sheet_id, sheets);
                        result.push_str(&format!(
                            "{}!{}{}{}{}:{}{}{}{}",
                            sheet_name,
                            sc_prefix,
                            start_col_letter,
                            sr_prefix,
                            start_row + 1,
                            ec_prefix,
                            end_col_letter,
                            er_prefix,
                            end_row + 1,
                        ));
                    } else {
                        result.push_str(&format!(
                            "{}{}{}{}:{}{}{}{}",
                            sc_prefix,
                            start_col_letter,
                            sr_prefix,
                            start_row + 1,
                            ec_prefix,
                            end_col_letter,
                            er_prefix,
                            end_row + 1,
                        ));
                    }
                }
            }
            FormulaPart::ColumnReference { sheet_id, col_id } => {
                let sheet_name = get_table_name(*sheet_id, sheets);
                let col_name = if let Some(sheet) = sheets.iter().find(|t| t.id == *sheet_id) {
                    if let Some(col) = sheet.columns.iter().find(|c| c.id == *col_id) {
                        col.name.clone()
                    } else {
                        "col_deleted".to_string()
                    }
                } else {
                    "col_deleted".to_string()
                };
                result.push_str(&format!("{}[{}]", sheet_name, col_name));
            }
            FormulaPart::StructuredReference {
                sheet_id,
                col_id,
                is_this_row,
                section,
            } => {
                let sheet_name = get_table_name(*sheet_id, sheets);
                let col_name = if let Some(col_id_val) = col_id {
                    if let Some(sheet) = sheets.iter().find(|t| t.id == *sheet_id) {
                        if let Some(col) = sheet.columns.iter().find(|c| c.id == *col_id_val) {
                            col.name.clone()
                        } else {
                            "col_deleted".to_string()
                        }
                    } else {
                        "col_deleted".to_string()
                    }
                } else {
                    String::new()
                };

                let has_prefix = if sheets.is_empty() {
                    false
                } else {
                    *sheet_id != sheets[0].id
                };

                let prefix = if has_prefix {
                    sheet_name
                } else {
                    String::new()
                };

                if *is_this_row {
                    if col_name.is_empty() {
                        result.push_str(&format!("{}[@]", prefix));
                    } else {
                        result.push_str(&format!("{}[@{}]", prefix, col_name));
                    }
                } else {
                    match section {
                        SheetSection::Headers => {
                            if col_name.is_empty() {
                                result.push_str(&format!("{}[#Headers]", prefix));
                            } else {
                                result
                                    .push_str(&format!("{}[[#Headers], [{}]]", prefix, col_name));
                            }
                        }
                        SheetSection::Totals => {
                            if col_name.is_empty() {
                                result.push_str(&format!("{}[#Totals]", prefix));
                            } else {
                                result.push_str(&format!("{}[[#Totals], [{}]]", prefix, col_name));
                            }
                        }
                        SheetSection::Data => {
                            if col_name.is_empty() {
                                result.push_str(&format!("{}[#Data]", prefix));
                            } else {
                                result.push_str(&format!("{}[{}]", prefix, col_name));
                            }
                        }
                        SheetSection::All => {
                            if col_name.is_empty() {
                                result.push_str(&format!("{}[#All]", prefix));
                            } else {
                                result.push_str(&format!("{}[[#All], [{}]]", prefix, col_name));
                            }
                        }
                    }
                }
            }
        }
    }
    result
}

pub fn lex_eval(input: &str) -> Result<Vec<EvalToken>, String> {
    let chars: Vec<char> = input.chars().collect();
    let mut tokens = Vec::new();
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() {
            i += 1;
            continue;
        }

        if c == '"' || c == '\'' {
            let quote = c;
            i += 1;
            let mut s = String::new();
            while i < chars.len() {
                if chars[i] == '\\' && i + 1 < chars.len() {
                    s.push(chars[i + 1]);
                    i += 2;
                } else if chars[i] == quote {
                    i += 1;
                    break;
                } else {
                    s.push(chars[i]);
                    i += 1;
                }
            }
            tokens.push(EvalToken::String(s));
            continue;
        }

        if c == '[' {
            if let Some((column, is_this_row, section, next_i)) =
                parse_structured_specifier(&chars, i)
            {
                let mut sheet = None;
                if let Some(last_tok) = tokens.last() {
                    match last_tok {
                        EvalToken::Identifier(_) | EvalToken::String(_) => {
                            let last = tokens.pop().unwrap();
                            sheet = match last {
                                EvalToken::Identifier(s) => Some(s),
                                EvalToken::String(s) => Some(s),
                                _ => unreachable!(),
                            };
                        }
                        _ => {}
                    }
                }
                tokens.push(EvalToken::StructuredRef {
                    sheet,
                    column,
                    is_this_row,
                    section,
                });
                i = next_i;
                continue;
            }
        }

        match c {
            '(' => {
                tokens.push(EvalToken::OpenParen);
                i += 1;
                continue;
            }
            ')' => {
                tokens.push(EvalToken::CloseParen);
                i += 1;
                continue;
            }
            '[' => {
                tokens.push(EvalToken::OpenBracket);
                i += 1;
                continue;
            }
            ']' => {
                tokens.push(EvalToken::CloseBracket);
                i += 1;
                continue;
            }
            ',' => {
                tokens.push(EvalToken::Comma);
                i += 1;
                continue;
            }
            ':' => {
                tokens.push(EvalToken::Colon);
                i += 1;
                continue;
            }
            _ => {}
        }

        if c == '<' {
            if i + 1 < chars.len() && chars[i + 1] == '>' {
                tokens.push(EvalToken::Op(Op::Ne));
                i += 2;
            } else if i + 1 < chars.len() && chars[i + 1] == '=' {
                tokens.push(EvalToken::Op(Op::Le));
                i += 2;
            } else {
                tokens.push(EvalToken::Op(Op::Lt));
                i += 1;
            }
            continue;
        }
        if c == '>' {
            if i + 1 < chars.len() && chars[i + 1] == '=' {
                tokens.push(EvalToken::Op(Op::Ge));
                i += 2;
            } else {
                tokens.push(EvalToken::Op(Op::Gt));
                i += 1;
            }
            continue;
        }
        if c == '=' {
            if i + 1 < chars.len() && chars[i + 1] == '=' {
                tokens.push(EvalToken::Op(Op::Eq));
                i += 2;
            } else {
                tokens.push(EvalToken::Op(Op::Eq));
                i += 1;
            }
            continue;
        }
        if c == '!' {
            if i + 1 < chars.len() && chars[i + 1] == '=' {
                tokens.push(EvalToken::Op(Op::Ne));
                i += 2;
            } else {
                tokens.push(EvalToken::Exclamation);
                i += 1;
            }
            continue;
        }

        match c {
            '+' => {
                tokens.push(EvalToken::Op(Op::Add));
                i += 1;
                continue;
            }
            '-' => {
                tokens.push(EvalToken::Op(Op::Sub));
                i += 1;
                continue;
            }
            '*' => {
                if i + 1 < chars.len() && chars[i + 1] == '*' {
                    tokens.push(EvalToken::Op(Op::Exp));
                    i += 2;
                } else {
                    tokens.push(EvalToken::Op(Op::Mul));
                    i += 1;
                }
                continue;
            }
            '/' => {
                tokens.push(EvalToken::Op(Op::Div));
                i += 1;
                continue;
            }
            '^' => {
                tokens.push(EvalToken::Op(Op::Exp));
                i += 1;
                continue;
            }
            _ => {}
        }

        if c == '.' {
            if i + 1 < chars.len() && chars[i + 1].is_ascii_digit() {
                let mut num_str = String::new();
                num_str.push('.');
                i += 1;
                while i < chars.len() && chars[i].is_ascii_digit() {
                    num_str.push(chars[i]);
                    i += 1;
                }
                if let Ok(val) = num_str.parse::<f64>() {
                    tokens.push(EvalToken::Number(val));
                } else {
                    return Err(format!("Invalid number: {}", num_str));
                }
            } else {
                tokens.push(EvalToken::Dot);
                i += 1;
            }
            continue;
        }

        if c.is_ascii_digit() {
            let mut num_str = String::new();
            while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                num_str.push(chars[i]);
                i += 1;
            }
            if let Ok(val) = num_str.parse::<f64>() {
                tokens.push(EvalToken::Number(val));
            } else {
                return Err(format!("Invalid number: {}", num_str));
            }
            continue;
        }

        if c.is_ascii_alphabetic() || c == '_' || c == '$' {
            let mut id_str = String::new();
            while i < chars.len()
                && (chars[i].is_ascii_alphanumeric()
                    || chars[i] == '_'
                    || chars[i] == '$'
                    || chars[i] == '.')
            {
                id_str.push(chars[i]);
                i += 1;
            }

            let upper = id_str.to_uppercase();
            if upper == "TRUE" {
                tokens.push(EvalToken::Boolean(true));
            } else if upper == "FALSE" {
                tokens.push(EvalToken::Boolean(false));
            } else {
                tokens.push(EvalToken::Identifier(id_str));
            }
            continue;
        }

        return Err(format!("Unexpected character: {}", c));
    }

    Ok(tokens)
}

struct Parser<'a> {
    tokens: &'a [EvalToken],
    pos: usize,
}

impl<'a> Parser<'a> {
    fn new(tokens: &'a [EvalToken]) -> Self {
        Self { tokens, pos: 0 }
    }

    fn peek(&self) -> Option<&EvalToken> {
        self.tokens.get(self.pos)
    }

    fn next(&mut self) -> Option<&EvalToken> {
        if self.pos < self.tokens.len() {
            let tok = &self.tokens[self.pos];
            self.pos += 1;
            Some(tok)
        } else {
            None
        }
    }

    fn consume(&mut self, expected: EvalToken) -> Result<(), String> {
        match self.next() {
            Some(tok) if *tok == expected => Ok(()),
            Some(tok) => Err(format!("Expected {:?}, got {:?}", expected, tok)),
            None => Err(format!("Expected {:?}, got EOF", expected)),
        }
    }

    fn parse(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_binary(0)?;

        while let Some(tok) = self.peek() {
            match tok {
                EvalToken::OpenBracket => {
                    self.next(); // consume `[`
                    let mut inner_tokens = Vec::new();
                    let mut depth = 1;
                    while let Some(t) = self.next() {
                        if *t == EvalToken::OpenBracket {
                            depth += 1;
                            inner_tokens.push(t.clone());
                        } else if *t == EvalToken::CloseBracket {
                            depth -= 1;
                            if depth == 0 {
                                break;
                            }
                            inner_tokens.push(t.clone());
                        } else {
                            inner_tokens.push(t.clone());
                        }
                    }

                    if let Some(colon_pos) =
                        inner_tokens.iter().position(|t| *t == EvalToken::Colon)
                    {
                        let start_toks = &inner_tokens[0..colon_pos];
                        let end_toks = &inner_tokens[colon_pos + 1..];

                        let start_val = if start_toks.is_empty() {
                            None
                        } else {
                            let mut p = Parser::new(start_toks);
                            Some(p.parse()?)
                        };

                        let end_val = if end_toks.is_empty() {
                            None
                        } else {
                            let mut p = Parser::new(end_toks);
                            Some(p.parse()?)
                        };

                        let start_expr = start_val.unwrap_or(Expr::Number(0.0));
                        let end_expr = end_val.unwrap_or(Expr::Number(-1.0));

                        lhs = Expr::FunctionCall {
                            name: "SLICE".to_string(),
                            args: vec![lhs, start_expr, end_expr],
                        };
                    } else {
                        let mut p = Parser::new(&inner_tokens);
                        let index_expr = p.parse()?;

                        if let Expr::Identifier(ref sheet_name) = lhs {
                            let col_name_opt = match &index_expr {
                                Expr::String(s) => Some(s.clone()),
                                Expr::Identifier(s) => Some(s.clone()),
                                _ => None,
                            };
                            if let Some(col_name) = col_name_opt {
                                lhs = Expr::FunctionCall {
                                    name: "GET_COL".to_string(),
                                    args: vec![
                                        Expr::String(sheet_name.clone()),
                                        Expr::String(col_name.clone()),
                                    ],
                                };
                                continue;
                            }
                        }

                        lhs = Expr::FunctionCall {
                            name: "INDEX".to_string(),
                            args: vec![lhs, index_expr],
                        };
                    }
                }
                EvalToken::Dot => {
                    return Err("Dot member access is not supported".to_string());
                }
                _ => break,
            }
        }

        Ok(lhs)
    }

    fn parse_binary(&mut self, min_prec: u8) -> Result<Expr, String> {
        let mut lhs = self.parse_prefix()?;

        while let Some(tok) = self.peek() {
            let op = match tok {
                EvalToken::Op(o) => *o,
                _ => break,
            };

            let prec = op_precedence(op);
            if prec < min_prec {
                break;
            }

            self.next(); // consume operator

            let rhs = self.parse_binary(prec + 1)?;
            lhs = Expr::BinaryOp {
                op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            };
        }

        Ok(lhs)
    }

    fn parse_prefix(&mut self) -> Result<Expr, String> {
        let tok = self
            .next()
            .ok_or_else(|| "Unexpected EOF".to_string())?
            .clone();
        match tok {
            EvalToken::Number(val) => Ok(Expr::Number(val)),
            EvalToken::String(val) => {
                if self.peek() == Some(&EvalToken::Exclamation) {
                    self.next();
                    let target_tok = self
                        .next()
                        .ok_or_else(|| "Expected cell or column reference after `!`".to_string())?
                        .clone();
                    let target_str = match target_tok {
                        EvalToken::Identifier(s) => s,
                        _ => {
                            return Err(format!(
                                "Expected cell or column reference after `!`, got {:?}",
                                target_tok
                            ));
                        }
                    };

                    if self.peek() == Some(&EvalToken::Colon) {
                        self.next();
                        let end_tok = self
                            .next()
                            .ok_or_else(|| {
                                "Expected cell or column reference after `:`".to_string()
                            })?
                            .clone();
                        let end_str = match end_tok {
                            EvalToken::Identifier(s) => s,
                            _ => {
                                return Err(format!(
                                    "Expected cell or column reference after `:`, got {:?}",
                                    end_tok
                                ));
                            }
                        };

                        if let (
                            Some((s_row, s_col, s_row_abs, s_col_abs)),
                            Some((e_row, e_col, e_row_abs, e_col_abs)),
                        ) = (parse_cell_ref(&target_str), parse_cell_ref(&end_str))
                        {
                            return Ok(Expr::RangeRef {
                                sheet: Some(val),
                                start_row: s_row,
                                start_col: s_col,
                                end_row: e_row,
                                end_col: e_col,
                                start_row_abs: s_row_abs,
                                start_col_abs: s_col_abs,
                                end_row_abs: e_row_abs,
                                end_col_abs: e_col_abs,
                            });
                        } else if let (Some((s_col, s_col_abs)), Some((e_col, e_col_abs))) =
                            (parse_column_ref(&target_str), parse_column_ref(&end_str))
                        {
                            return Ok(Expr::RangeRef {
                                sheet: Some(val),
                                start_row: 0,
                                start_col: s_col,
                                end_row: usize::MAX,
                                end_col: e_col,
                                start_row_abs: true,
                                start_col_abs: s_col_abs,
                                end_row_abs: true,
                                end_col_abs: e_col_abs,
                            });
                        } else {
                            return Err(format!(
                                "Invalid range reference: {}:{}",
                                target_str, end_str
                            ));
                        }
                    } else {
                        let (row, col, row_abs, col_abs) = parse_cell_ref(&target_str)
                            .ok_or_else(|| format!("Invalid cell: {}", target_str))?;
                        return Ok(Expr::CellRef {
                            sheet: Some(val),
                            row,
                            col,
                            row_abs,
                            col_abs,
                        });
                    }
                }
                Ok(Expr::String(val.clone()))
            }
            EvalToken::Boolean(val) => Ok(Expr::Boolean(val)),
            EvalToken::OpenParen => {
                let expr = self.parse()?;
                self.consume(EvalToken::CloseParen)?;
                Ok(expr)
            }
            EvalToken::OpenBracket => {
                let mut list = Vec::new();
                if self.peek() != Some(&EvalToken::CloseBracket) {
                    loop {
                        list.push(self.parse()?);
                        if self.peek() == Some(&EvalToken::Comma) {
                            self.next();
                        } else {
                            break;
                        }
                    }
                }
                self.consume(EvalToken::CloseBracket)?;
                Ok(Expr::List(list))
            }
            EvalToken::Op(Op::Sub) => {
                let expr = self.parse_binary(100)?;
                Ok(Expr::UnaryOp {
                    op: Op::Sub,
                    expr: Box::new(expr),
                })
            }
            EvalToken::Op(Op::Add) => {
                let expr = self.parse_binary(100)?;
                Ok(expr)
            }
            EvalToken::Identifier(id_name) => {
                if self.peek() == Some(&EvalToken::OpenParen) {
                    self.next();
                    let mut args = Vec::new();
                    if self.peek() != Some(&EvalToken::CloseParen) {
                        loop {
                            args.push(self.parse()?);
                            if self.peek() == Some(&EvalToken::Comma) {
                                self.next();
                            } else {
                                break;
                            }
                        }
                    }
                    self.consume(EvalToken::CloseParen)?;
                    return Ok(Expr::FunctionCall {
                        name: id_name.clone(),
                        args,
                    });
                }

                if self.peek() == Some(&EvalToken::Exclamation) {
                    self.next();
                    let target_tok = self
                        .next()
                        .ok_or_else(|| "Expected cell or column reference after `!`".to_string())?;
                    let target_str = match target_tok {
                        EvalToken::Identifier(s) => s.clone(),
                        _ => {
                            return Err(format!(
                                "Expected cell or column reference after `!`, got {:?}",
                                target_tok
                            ));
                        }
                    };

                    if self.peek() == Some(&EvalToken::Colon) {
                        self.next();
                        let end_tok = self.next().ok_or_else(|| {
                            "Expected cell or column reference after `:`".to_string()
                        })?;
                        let end_str = match end_tok {
                            EvalToken::Identifier(s) => s.clone(),
                            _ => {
                                return Err(format!(
                                    "Expected cell or column reference after `:`, got {:?}",
                                    end_tok
                                ));
                            }
                        };

                        if let (
                            Some((s_row, s_col, s_row_abs, s_col_abs)),
                            Some((e_row, e_col, e_row_abs, e_col_abs)),
                        ) = (parse_cell_ref(&target_str), parse_cell_ref(&end_str))
                        {
                            return Ok(Expr::RangeRef {
                                sheet: Some(id_name.clone()),
                                start_row: s_row,
                                start_col: s_col,
                                end_row: e_row,
                                end_col: e_col,
                                start_row_abs: s_row_abs,
                                start_col_abs: s_col_abs,
                                end_row_abs: e_row_abs,
                                end_col_abs: e_col_abs,
                            });
                        } else if let (Some((s_col, s_col_abs)), Some((e_col, e_col_abs))) =
                            (parse_column_ref(&target_str), parse_column_ref(&end_str))
                        {
                            return Ok(Expr::RangeRef {
                                sheet: Some(id_name.clone()),
                                start_row: 0,
                                start_col: s_col,
                                end_row: usize::MAX,
                                end_col: e_col,
                                start_row_abs: true,
                                start_col_abs: s_col_abs,
                                end_row_abs: true,
                                end_col_abs: e_col_abs,
                            });
                        } else {
                            return Err(format!(
                                "Invalid range reference: {}:{}",
                                target_str, end_str
                            ));
                        }
                    } else {
                        let (row, col, row_abs, col_abs) = parse_cell_ref(&target_str)
                            .ok_or_else(|| format!("Invalid cell: {}", target_str))?;
                        return Ok(Expr::CellRef {
                            sheet: Some(id_name.clone()),
                            row,
                            col,
                            row_abs,
                            col_abs,
                        });
                    }
                }

                if let Some((row, col, row_abs, col_abs)) = parse_cell_ref(&id_name) {
                    if self.peek() == Some(&EvalToken::Colon) {
                        self.next();
                        let end_tok = self
                            .next()
                            .ok_or_else(|| "Expected cell reference after `:`".to_string())?;
                        let end_str = match end_tok {
                            EvalToken::Identifier(s) => s.clone(),
                            _ => {
                                return Err(format!(
                                    "Expected cell reference after `:`, got {:?}",
                                    end_tok
                                ));
                            }
                        };
                        let (e_row, e_col, e_row_abs, e_col_abs) = parse_cell_ref(&end_str)
                            .ok_or_else(|| format!("Invalid end cell: {}", end_str))?;
                        return Ok(Expr::RangeRef {
                            sheet: None,
                            start_row: row,
                            start_col: col,
                            end_row: e_row,
                            end_col: e_col,
                            start_row_abs: row_abs,
                            start_col_abs: col_abs,
                            end_row_abs: e_row_abs,
                            end_col_abs: e_col_abs,
                        });
                    }

                    return Ok(Expr::CellRef {
                        sheet: None,
                        row,
                        col,
                        row_abs,
                        col_abs,
                    });
                }

                if let Some((col, col_abs)) = parse_column_ref(&id_name) {
                    if self.peek() == Some(&EvalToken::Colon) {
                        self.next();
                        let end_tok = self
                            .next()
                            .ok_or_else(|| "Expected column reference after `:`".to_string())?;
                        let end_str = match end_tok {
                            EvalToken::Identifier(s) => s.clone(),
                            _ => {
                                return Err(format!(
                                    "Expected column reference after `:`, got {:?}",
                                    end_tok
                                ));
                            }
                        };
                        let (e_col, e_col_abs) = parse_column_ref(&end_str)
                            .ok_or_else(|| format!("Invalid end column: {}", end_str))?;
                        return Ok(Expr::RangeRef {
                            sheet: None,
                            start_row: 0,
                            start_col: col,
                            end_row: usize::MAX,
                            end_col: e_col,
                            start_row_abs: true,
                            start_col_abs: col_abs,
                            end_row_abs: true,
                            end_col_abs: e_col_abs,
                        });
                    }
                }

                Ok(Expr::Identifier(id_name.clone()))
            }
            EvalToken::StructuredRef {
                sheet,
                column,
                is_this_row,
                section,
            } => Ok(Expr::StructuredRef {
                sheet: sheet.clone(),
                column: column.clone(),
                is_this_row: is_this_row,
                section: section,
            }),
            _ => Err(format!("Unexpected token: {:?}", tok)),
        }
    }
}

fn op_precedence(op: Op) -> u8 {
    match op {
        Op::Add | Op::Sub => 10,
        Op::Mul | Op::Div => 20,
        Op::Exp => 30,
        Op::Eq | Op::Ne | Op::Lt | Op::Gt | Op::Le | Op::Ge => 5,
    }
}

pub fn parse_excel_formula(input: &str) -> Result<Expr, String> {
    let tokens = lex_eval(input)?;
    let mut parser = Parser::new(&tokens);
    let expr = parser.parse()?;
    if parser.peek().is_some() {
        return Err("Trailing tokens after expression".to_string());
    }
    Ok(expr)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compile_and_serialize() {
        let table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 5,
        });
        let table2 = Sheet::new(crate::core::SheetInit {
            id: Some(456),
            name: Some("Sheet2".to_string()),
            rows: 5,
            cols: 5,
        });

        let sheets = vec![table1, table2];

        let formula = compile_formula("=Sheet2!B1 + 10", &sheets);
        assert_eq!(formula.parts.len(), 3);
        match &formula.parts[1] {
            FormulaPart::SheetReference {
                sheet_id, row, col, ..
            } => {
                assert_eq!(*sheet_id, 456);
                assert_eq!(*row, 0);
                assert_eq!(*col, 1);
            }
            _ => panic!("Expected SheetReference"),
        }

        let mut renamed_table2 = sheets[1].clone();
        renamed_table2.name = "Sheet3".to_string();

        let serialized = serialize_formula(&formula, &[sheets[0].clone(), renamed_table2]);
        assert_eq!(serialized, "=Sheet3!B1 + 10");
    }

    #[test]
    fn test_table_names_with_spaces() {
        let table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 5,
        });
        let table2 = Sheet::new(crate::core::SheetInit {
            id: Some(456),
            name: Some("My Sheet".to_string()),
            rows: 5,
            cols: 5,
        });

        let sheets = vec![table1, table2];

        // 1. Single-quote notation
        let formula_quote = compile_formula("='My Sheet'!B1 + 10", &sheets);
        match &formula_quote.parts[1] {
            FormulaPart::SheetReference {
                sheet_id, row, col, ..
            } => {
                assert_eq!(*sheet_id, 456);
                assert_eq!(*row, 0);
                assert_eq!(*col, 1);
            }
            _ => panic!("Expected SheetReference"),
        }

        // 2. Serialization wraps sheet name with spaces in single quotes
        let serialized_quote = serialize_formula(&formula_quote, &sheets);
        assert_eq!(serialized_quote, "='My Sheet'!B1 + 10");

        // 3. parse_excel_formula handles single-quoted sheet references
        let ast_quote = parse_excel_formula("'My Sheet'!B1 + 10").unwrap();
        match ast_quote {
            Expr::BinaryOp { left, .. } => match *left {
                Expr::CellRef {
                    sheet, row, col, ..
                } => {
                    assert_eq!(sheet, Some("My Sheet".to_string()));
                    assert_eq!(row, 0);
                    assert_eq!(col, 1);
                }
                _ => panic!("Expected CellRef"),
            },
            _ => panic!("Expected BinaryOp"),
        }
    }

    #[test]
    fn test_table_names_with_periods() {
        let table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Model_SL_5.5Yr".to_string()),
            rows: 5,
            cols: 5,
        });

        let sheets = vec![table1];

        // 1. Check compile_formula with unquoted period name
        let formula = compile_formula("=Model_SL_5.5Yr!B10 + 10", &sheets);
        assert_eq!(formula.parts.len(), 3);
        match &formula.parts[1] {
            FormulaPart::SheetReference {
                sheet_id, row, col, ..
            } => {
                assert_eq!(*sheet_id, 123);
                assert_eq!(*row, 9);
                assert_eq!(*col, 1);
            }
            _ => panic!("Expected SheetReference"),
        }

        // 2. Check parse_excel_formula with unquoted period name
        let ast = parse_excel_formula("Model_SL_5.5Yr!B10 + 10").unwrap();
        match ast {
            Expr::BinaryOp { left, .. } => match *left {
                Expr::CellRef {
                    sheet, row, col, ..
                } => {
                    assert_eq!(sheet, Some("Model_SL_5.5Yr".to_string()));
                    assert_eq!(row, 9);
                    assert_eq!(col, 1);
                }
                _ => panic!("Expected CellRef"),
            },
            _ => panic!("Expected BinaryOp"),
        }
    }

    #[test]
    fn test_column_ranges() {
        let table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 5,
        });
        let sheets = vec![table1];

        // Test compiling local column range
        let formula = compile_formula("=SUM(A:B)", &sheets);
        assert_eq!(formula.parts.len(), 3);
        match &formula.parts[1] {
            FormulaPart::RangeReference {
                sheet_id,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => {
                assert_eq!(*sheet_id, 123);
                assert_eq!(*start_row, 0);
                assert_eq!(*start_col, 0);
                assert_eq!(*end_row, usize::MAX);
                assert_eq!(*end_col, 1);
            }
            _ => panic!("Expected RangeReference"),
        }

        // Test serializing local column range
        let serialized = serialize_formula(&formula, &sheets);
        assert_eq!(serialized, "=SUM(A:B)");

        // Test compiling cross-sheet column range
        let table2 = Sheet::new(crate::core::SheetInit {
            id: Some(456),
            name: Some("Sheet2".to_string()),
            rows: 5,
            cols: 5,
        });
        let tables_multi = vec![sheets[0].clone(), table2];
        let formula_cross = compile_formula("=SUM(Sheet2!$A:$C)", &tables_multi);
        match &formula_cross.parts[1] {
            FormulaPart::RangeReference {
                sheet_id,
                start_row,
                start_col,
                end_row,
                end_col,
                start_col_ref_type,
                end_col_ref_type,
                ..
            } => {
                assert_eq!(*sheet_id, 456);
                assert_eq!(*start_row, 0);
                assert_eq!(*start_col, 0);
                assert_eq!(*end_row, usize::MAX);
                assert_eq!(*end_col, 2);
                assert_eq!(*start_col_ref_type, RefType::Absolute);
                assert_eq!(*end_col_ref_type, RefType::Absolute);
            }
            _ => panic!("Expected RangeReference for cross sheet"),
        }

        let serialized_cross = serialize_formula(&formula_cross, &tables_multi);
        assert_eq!(serialized_cross, "=SUM(Sheet2!$A:$C)");

        // Test parse_excel_formula with column range
        let ast = parse_excel_formula("SUM(A:A)").unwrap();
        match ast {
            Expr::FunctionCall { name, args } => {
                assert_eq!(name, "SUM");
                assert_eq!(args.len(), 1);
                match &args[0] {
                    Expr::RangeRef {
                        sheet,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        ..
                    } => {
                        assert_eq!(*sheet, None);
                        assert_eq!(*start_row, 0);
                        assert_eq!(*start_col, 0);
                        assert_eq!(*end_row, usize::MAX);
                        assert_eq!(*end_col, 0);
                    }
                    _ => panic!("Expected RangeRef inside SUM"),
                }
            }
            _ => panic!("Expected FunctionCall"),
        }
    }

    #[test]
    fn test_excel_structured_references() {
        let mut table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 2,
        });
        table1.columns[0].name = "Sales".to_string();
        table1.columns[0].id = 1;
        table1.columns[1].name = "Cost".to_string();
        table1.columns[1].id = 2;

        let sheets = vec![table1];

        // 1. Test compile_formula with TableName[ColumnName]
        let f1 = compile_formula("=Sheet1[Sales]", &sheets);
        assert_eq!(f1.parts.len(), 2);
        match &f1.parts[1] {
            FormulaPart::StructuredReference {
                sheet_id,
                col_id,
                is_this_row,
                section,
            } => {
                assert_eq!(*sheet_id, 123);
                assert_eq!(*col_id, Some(1));
                assert!(!is_this_row);
                assert_eq!(*section, SheetSection::Data);
            }
            _ => panic!("Expected StructuredReference, got {:?}", f1.parts[1]),
        }

        // 2. Test serialize_formula (omits prefix for current sheet)
        let s1 = serialize_formula(&f1, &sheets);
        assert_eq!(s1, "=[Sales]");

        // 3. Test compile_formula with [@ColumnName]
        let f2 = compile_formula("=[@Cost]", &sheets);
        match &f2.parts[1] {
            FormulaPart::StructuredReference {
                sheet_id,
                col_id,
                is_this_row,
                section,
            } => {
                assert_eq!(*sheet_id, 123);
                assert_eq!(*col_id, Some(2));
                assert!(is_this_row);
                assert_eq!(*section, SheetSection::Data);
            }
            _ => panic!("Expected StructuredReference"),
        }
        let s2 = serialize_formula(&f2, &sheets);
        assert_eq!(s2, "=[@Cost]");

        // 4. Test parsing special items: TableName[[#Headers],[Sales]]
        let f3 = compile_formula("=Sheet1[[#Headers],[Sales]]", &sheets);
        match &f3.parts[1] {
            FormulaPart::StructuredReference {
                sheet_id,
                col_id,
                section,
                ..
            } => {
                assert_eq!(*sheet_id, 123);
                assert_eq!(*col_id, Some(1));
                assert_eq!(*section, SheetSection::Headers);
            }
            _ => panic!("Expected StructuredReference"),
        }
        let s3 = serialize_formula(&f3, &sheets);
        assert_eq!(s3, "=[[#Headers], [Sales]]");

        // 5. Test retaining prefix when referencing a remote sheet
        let mut table2 = Sheet::new(crate::core::SheetInit {
            id: Some(456),
            name: Some("Sheet2".to_string()),
            rows: 5,
            cols: 2,
        });
        table2.columns[0].name = "Revenue".to_string();
        table2.columns[0].id = 3;
        table2.columns[1].name = "Expenses".to_string();
        table2.columns[1].id = 4;

        let multi_tables = vec![sheets[0].clone(), table2];
        let f4 = compile_formula("=Sheet2[Revenue]", &multi_tables);
        let s4 = serialize_formula(&f4, &multi_tables);
        assert_eq!(s4, "=Sheet2[Revenue]");

        // 5. Test parse_excel_formula for evaluation AST
        let ast = parse_excel_formula("Sheet1[@Sales] + 10").unwrap();
        match ast {
            Expr::BinaryOp { op, left, right: _ } => {
                assert_eq!(op, Op::Add);
                match &*left {
                    Expr::StructuredRef {
                        sheet,
                        column,
                        is_this_row,
                        ..
                    } => {
                        assert_eq!(sheet.as_deref(), Some("Sheet1"));
                        assert_eq!(column.as_deref(), Some("Sales"));
                        assert!(is_this_row);
                    }
                    _ => panic!("Expected StructuredRef"),
                }
            }
            _ => panic!("Expected BinaryOp"),
        }
    }

    #[test]
    fn test_structured_reference_whole_row_no_column() {
        let mut table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 2,
        });
        table1.columns[0].name = "Sales".to_string();
        table1.columns[0].id = 1;
        table1.columns[1].name = "Cost".to_string();
        table1.columns[1].id = 2;
        let sheets = vec![table1];

        // `[@]` (this row, no specific column) should parse with column = None,
        // not a column literally named "".
        let f = compile_formula("=[@]", &sheets);
        match &f.parts[1] {
            FormulaPart::StructuredReference {
                col_id,
                is_this_row,
                section,
                ..
            } => {
                assert_eq!(*col_id, None);
                assert!(is_this_row);
                assert_eq!(*section, SheetSection::Data);
            }
            _ => panic!("Expected StructuredReference, got {:?}", f.parts[1]),
        }
        assert_eq!(serialize_formula(&f, &sheets), "=[@]");

        let ast = parse_excel_formula("[@]").unwrap();
        match ast {
            Expr::StructuredRef {
                column, is_this_row, ..
            } => {
                assert_eq!(column, None);
                assert!(is_this_row);
            }
            _ => panic!("Expected StructuredRef"),
        }
    }

    #[test]
    fn test_structured_reference_whole_table_sections_no_column() {
        let mut table1 = Sheet::new(crate::core::SheetInit {
            id: Some(123),
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 2,
        });
        table1.columns[0].name = "Sales".to_string();
        table1.columns[1].name = "Cost".to_string();
        let sheets = vec![table1];

        for (input, expected_section) in [
            ("=[#Data]", SheetSection::Data),
            ("=[#All]", SheetSection::All),
            ("=[#Headers]", SheetSection::Headers),
            ("=[#Totals]", SheetSection::Totals),
        ] {
            let f = compile_formula(input, &sheets);
            match &f.parts[1] {
                FormulaPart::StructuredReference {
                    col_id,
                    is_this_row,
                    section,
                    ..
                } => {
                    assert_eq!(*col_id, None, "input: {input}");
                    assert!(!is_this_row, "input: {input}");
                    assert_eq!(*section, expected_section, "input: {input}");
                }
                _ => panic!("Expected StructuredReference for {input}, got {:?}", f.parts[1]),
            }
            // Round-trips back to the same whole-section syntax.
            assert_eq!(serialize_formula(&f, &sheets), input);
        }
    }
}
