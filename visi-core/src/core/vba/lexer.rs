use std::fmt;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Pos {
    pub line: u32,

    pub col: u32,
}

impl fmt::Display for Pos {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}:{}", self.line, self.col)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumBase {
    Decimal,

    Hex,

    Octal,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TypeSuffix {
    String,

    Integer,

    Long,

    Single,

    Double,

    Currency,
}

impl TypeSuffix {
    fn from_char(c: char) -> Option<Self> {
        match c {
            '$' => Some(Self::String),
            '%' => Some(Self::Integer),
            '&' => Some(Self::Long),
            '!' => Some(Self::Single),
            '#' => Some(Self::Double),
            '@' => Some(Self::Currency),
            _ => None,
        }
    }

    pub fn as_char(self) -> char {
        match self {
            Self::String => '$',
            Self::Integer => '%',
            Self::Long => '&',
            Self::Single => '!',
            Self::Double => '#',
            Self::Currency => '@',
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    Ident(String),

    Number {
        value: f64,

        base: NumBase,

        suffix: Option<TypeSuffix>,

        is_float: bool,
    },

    Str(String),

    Date(String),

    Punct(&'static str),

    Newline,

    Eof,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokenKind,

    pub pos: Pos,

    pub preceded_by_space: bool,
}

impl Token {
    pub fn ident(&self) -> Option<&str> {
        match &self.kind {
            TokenKind::Ident(s) => Some(s),
            _ => None,
        }
    }

    pub fn is_kw(&self, kw: &str) -> bool {
        self.ident().is_some_and(|s| s.eq_ignore_ascii_case(kw))
    }

    pub fn is_punct(&self, p: &str) -> bool {
        matches!(&self.kind, TokenKind::Punct(x) if *x == p)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct LexError {
    pub message: String,

    pub pos: Pos,
}

impl fmt::Display for LexError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{} at line {}, column {}",
            self.message, self.pos.line, self.pos.col
        )
    }
}

const MULTI_PUNCT: &[&str] = &[":=", "<=", ">=", "<>", "=<", "=>"];
const SINGLE_PUNCT: &[char] = &[
    '(', ')', ',', '.', '=', '+', '-', '*', '/', '\\', '^', '&', '<', '>', ':', ';', '!', '#', '$',
    '%', '@', '?', '{', '}', '[', ']', '~', '|',
];

pub fn lex(src: &str) -> Result<Vec<Token>, LexError> {
    Lexer::new(src).run()
}

struct Lexer {
    chars: Vec<char>,
    i: usize,
    line: u32,
    col: u32,
    out: Vec<Token>,
    space_before: bool,
}

impl Lexer {
    fn new(src: &str) -> Self {
        Self {
            chars: src.chars().collect(),
            i: 0,
            line: 1,
            col: 1,
            out: Vec::new(),
            space_before: false,
        }
    }

    fn pos(&self) -> Pos {
        Pos {
            line: self.line,
            col: self.col,
        }
    }

    fn peek(&self) -> Option<char> {
        self.chars.get(self.i).copied()
    }

    fn peek_at(&self, n: usize) -> Option<char> {
        self.chars.get(self.i + n).copied()
    }

    fn bump(&mut self) -> Option<char> {
        let c = self.chars.get(self.i).copied()?;
        self.i += 1;
        if c == '\n' {
            self.line = self.line.saturating_add(1);
            self.col = 1;
        } else {
            self.col = self.col.saturating_add(1);
        }
        Some(c)
    }

    fn push(&mut self, kind: TokenKind, pos: Pos) {
        let preceded_by_space = self.space_before;
        self.out.push(Token {
            kind,
            pos,
            preceded_by_space,
        });
        self.space_before = false;
    }

    fn last_takes_suffix(&self) -> bool {
        matches!(
            self.out.last().map(|t| &t.kind),
            Some(TokenKind::Ident(_)) | Some(TokenKind::Number { .. })
        )
    }

    fn run(mut self) -> Result<Vec<Token>, LexError> {
        while let Some(c) = self.peek() {
            match c {
                '\r' => {
                    self.bump();
                }
                '\n' => {
                    let pos = self.pos();
                    self.bump();
                    self.push(TokenKind::Newline, pos);
                }
                ' ' | '\t' => {
                    self.bump();
                    self.space_before = true;
                }
                '_' if self.is_line_continuation() => {
                    self.consume_line_continuation();
                    self.space_before = true;
                }
                '_' => {
                    return Err(LexError {
                        message: "Invalid character: a name cannot start with '_'".to_string(),
                        pos: self.pos(),
                    });
                }
                '\'' => self.skip_line_comment(),
                '"' => self.lex_string()?,
                '#' => self.lex_hash()?,
                '&' if self.starts_based_number() => self.lex_based_number()?,
                c if c.is_ascii_digit() => self.lex_number()?,
                '.' if self.peek_at(1).is_some_and(|d| d.is_ascii_digit()) => self.lex_number()?,
                c if is_ident_start(c) => self.lex_ident_or_rem(),
                _ => self.lex_punct()?,
            }
        }
        let pos = self.pos();
        self.push(TokenKind::Eof, pos);
        Ok(self.out)
    }

    fn is_line_continuation(&self) -> bool {
        let prev_ok = self.i == 0
            || self
                .chars
                .get(self.i - 1)
                .is_some_and(|c| *c == ' ' || *c == '\t');
        if !prev_ok {
            return false;
        }
        let mut j = self.i + 1;
        while let Some(c) = self.chars.get(j) {
            match c {
                ' ' | '\t' | '\r' => j += 1,
                '\n' => return true,
                _ => return false,
            }
        }
        true
    }

    fn consume_line_continuation(&mut self) {
        while let Some(c) = self.peek() {
            self.bump();
            if c == '\n' {
                break;
            }
        }
    }

    fn skip_line_comment(&mut self) {
        while let Some(c) = self.peek() {
            if c == '\n' {
                break;
            }
            self.bump();
        }
    }

    fn lex_string(&mut self) -> Result<(), LexError> {
        let pos = self.pos();
        self.bump();
        let mut s = String::new();
        loop {
            match self.peek() {
                None | Some('\n') => {
                    return Err(LexError {
                        message: "unterminated string literal".to_string(),
                        pos,
                    });
                }
                Some('"') => {
                    self.bump();
                    if self.peek() == Some('"') {
                        self.bump();
                        s.push('"');
                    } else {
                        break;
                    }
                }
                Some(c) => {
                    self.bump();
                    s.push(c);
                }
            }
        }
        self.push(TokenKind::Str(s), pos);
        Ok(())
    }

    fn lex_hash(&mut self) -> Result<(), LexError> {
        let pos = self.pos();

        if !self.space_before && self.last_takes_suffix() {
            self.bump();
            self.attach_suffix(TypeSuffix::Double);
            return Ok(());
        }

        let mut j = self.i + 1;
        let mut content = String::new();
        while let Some(&c) = self.chars.get(j) {
            if c == '\n' {
                break;
            }
            if c == '#' {
                if content.trim().is_empty() {
                    break;
                }
                for _ in 0..=(j - self.i) {
                    self.bump();
                }
                self.push(TokenKind::Date(content.trim().to_string()), pos);
                return Ok(());
            }
            content.push(c);
            j += 1;
        }

        self.bump();
        self.push(TokenKind::Punct("#"), pos);
        Ok(())
    }

    fn attach_suffix(&mut self, suffix: TypeSuffix) {
        if let Some(last) = self.out.last_mut() {
            match &mut last.kind {
                TokenKind::Number { suffix: slot, .. } => *slot = Some(suffix),
                TokenKind::Ident(name) => name.push(suffix.as_char()),
                _ => {}
            }
        }
    }

    fn starts_based_number(&self) -> bool {
        if !self.space_before && self.last_takes_suffix() {
            return false;
        }
        matches!(self.peek_at(1), Some(c) if c == 'h' || c == 'H' || c == 'o' || c == 'O')
    }

    fn lex_based_number(&mut self) -> Result<(), LexError> {
        let pos = self.pos();
        self.bump();
        let marker = self.bump().unwrap_or('h');
        let (base, radix) = if marker.eq_ignore_ascii_case(&'h') {
            (NumBase::Hex, 16)
        } else {
            (NumBase::Octal, 8)
        };
        let mut digits = String::new();
        while let Some(c) = self.peek() {
            if c.is_digit(radix) {
                digits.push(c);
                self.bump();
            } else {
                break;
            }
        }
        if digits.is_empty() {
            return Err(LexError {
                message: format!(
                    "expected {} digits after &{}",
                    if radix == 16 { "hexadecimal" } else { "octal" },
                    marker
                ),
                pos,
            });
        }
        let value = u128::from_str_radix(&digits, radix).map_err(|_| LexError {
            message: "numeric literal is too large".to_string(),
            pos,
        })?;
        self.push(
            TokenKind::Number {
                value: value as f64,
                base,
                suffix: None,
                is_float: false,
            },
            pos,
        );
        if let Some(c) = self.peek()
            && let Some(suffix) = TypeSuffix::from_char(c)
        {
            self.bump();
            self.attach_suffix(suffix);
        }
        Ok(())
    }

    fn lex_number(&mut self) -> Result<(), LexError> {
        let pos = self.pos();
        let start = self.i;
        while self.peek().is_some_and(|c| c.is_ascii_digit()) {
            self.bump();
        }
        if self.peek() == Some('.') && self.peek_at(1).is_some_and(|c| c.is_ascii_digit()) {
            self.bump();
            while self.peek().is_some_and(|c| c.is_ascii_digit()) {
                self.bump();
            }
        } else if self.peek() == Some('.') && self.i == start {
            self.bump();
            while self.peek().is_some_and(|c| c.is_ascii_digit()) {
                self.bump();
            }
        }
        if let Some(e) = self.peek()
            && (e == 'e' || e == 'E' || e == 'd' || e == 'D')
        {
            let sign = self.peek_at(1);
            let digit_at = if matches!(sign, Some('+') | Some('-')) {
                2
            } else {
                1
            };
            if self.peek_at(digit_at).is_some_and(|c| c.is_ascii_digit()) {
                self.bump();
                if digit_at == 2 {
                    self.bump();
                }
                while self.peek().is_some_and(|c| c.is_ascii_digit()) {
                    self.bump();
                }
            }
        }
        let text: String = self.chars[start..self.i]
            .iter()
            .map(|c| if *c == 'd' || *c == 'D' { 'e' } else { *c })
            .collect();
        let value = text.parse::<f64>().map_err(|_| LexError {
            message: format!("malformed numeric literal {text:?}"),
            pos,
        })?;
        if !value.is_finite() {
            return Err(LexError {
                message: "numeric literal overflows".to_string(),
                pos,
            });
        }
        let is_float = text.contains('.') || text.contains(['e', 'E']);
        self.push(
            TokenKind::Number {
                value,
                base: NumBase::Decimal,
                suffix: None,
                is_float,
            },
            pos,
        );
        if let Some(c) = self.peek()
            && let Some(suffix) = TypeSuffix::from_char(c)
            && !self.suffix_would_be_operator(c)
        {
            self.bump();
            self.attach_suffix(suffix);
        }
        Ok(())
    }

    fn lex_ident_or_rem(&mut self) {
        let pos = self.pos();
        let start = self.i;
        while self.peek().is_some_and(is_ident_continue) {
            self.bump();
        }
        let name: String = self.chars[start..self.i].iter().collect();

        if name.eq_ignore_ascii_case("rem") && self.starts_statement() {
            self.skip_line_comment();
            return;
        }

        self.push(TokenKind::Ident(name), pos);
        if let Some(c) = self.peek()
            && let Some(suffix) = TypeSuffix::from_char(c)
            && !self.suffix_would_be_operator(c)
        {
            self.bump();
            self.attach_suffix(suffix);
        }
    }

    fn starts_statement(&self) -> bool {
        match self.out.last().map(|t| &t.kind) {
            None | Some(TokenKind::Newline) => true,
            Some(TokenKind::Punct(p)) => *p == ":",
            _ => false,
        }
    }

    fn suffix_would_be_operator(&self, sigil: char) -> bool {
        if sigil != '&' && sigil != '!' {
            return false;
        }
        match self.peek_at(1) {
            Some(c) => is_ident_start(c) || c.is_ascii_digit() || c == '"' || c == '[',
            None => false,
        }
    }

    fn lex_punct(&mut self) -> Result<(), LexError> {
        let pos = self.pos();
        let c = self.peek().unwrap_or('\0');

        if !self.space_before
            && self.last_takes_suffix()
            && let Some(suffix) = TypeSuffix::from_char(c)
            && !self.suffix_would_be_operator(c)
        {
            self.bump();
            self.attach_suffix(suffix);
            return Ok(());
        }

        let two: String = self.chars[self.i..(self.i + 2).min(self.chars.len())]
            .iter()
            .collect();
        for p in MULTI_PUNCT {
            if two == *p {
                self.bump();
                self.bump();
                let canon = match *p {
                    "=<" => "<=",
                    "=>" => ">=",
                    other => other,
                };
                self.push(TokenKind::Punct(canon), pos);
                return Ok(());
            }
        }

        if let Some(idx) = SINGLE_PUNCT.iter().position(|p| *p == c) {
            self.bump();
            self.push(TokenKind::Punct(SINGLE_PUNCT_STR[idx]), pos);
            return Ok(());
        }

        Err(LexError {
            message: format!("unexpected character {c:?}"),
            pos,
        })
    }
}

const SINGLE_PUNCT_STR: &[&str] = &[
    "(", ")", ",", ".", "=", "+", "-", "*", "/", "\\", "^", "&", "<", ">", ":", ";", "!", "#", "$",
    "%", "@", "?", "{", "}", "[", "]", "~", "|",
];

fn is_ident_start(c: char) -> bool {
    c.is_alphabetic() || c == '_'
}

fn is_ident_continue(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(src: &str) -> Vec<TokenKind> {
        lex(src).unwrap().into_iter().map(|t| t.kind).collect()
    }

    fn idents(src: &str) -> Vec<String> {
        lex(src)
            .unwrap()
            .into_iter()
            .filter_map(|t| t.ident().map(|s| s.to_string()))
            .collect()
    }

    fn num(src: &str) -> f64 {
        match &kinds(src)[0] {
            TokenKind::Number { value, .. } => *value,
            other => panic!("not a number: {other:?}"),
        }
    }

    #[test]
    fn newlines_are_tokens_because_they_end_statements() {
        assert_eq!(
            kinds("a\nb"),
            vec![
                TokenKind::Ident("a".into()),
                TokenKind::Newline,
                TokenKind::Ident("b".into()),
                TokenKind::Eof,
            ]
        );
    }

    #[test]
    fn crlf_emits_one_newline() {
        assert_eq!(
            kinds("a\r\nb")
                .iter()
                .filter(|k| **k == TokenKind::Newline)
                .count(),
            1
        );
    }

    #[test]
    fn line_continuation_splices_lines() {
        assert_eq!(
            kinds("a _\n b"),
            vec![
                TokenKind::Ident("a".into()),
                TokenKind::Ident("b".into()),
                TokenKind::Eof,
            ]
        );
    }

    #[test]
    fn underscore_inside_a_name_is_not_a_continuation() {
        assert_eq!(idents("my_var\n"), vec!["my_var"]);
    }

    #[test]
    fn a_name_cannot_start_with_an_underscore() {
        assert!(lex("_leading\n").is_err());
        assert!(lex("_ y = 2\n").is_err());
        assert_eq!(idents("y = 1 + _\n    2\n"), vec!["y"]);
    }

    #[test]
    fn comments_are_dropped_but_their_newline_survives() {
        assert_eq!(
            kinds("a ' trailing\nb"),
            vec![
                TokenKind::Ident("a".into()),
                TokenKind::Newline,
                TokenKind::Ident("b".into()),
                TokenKind::Eof,
            ]
        );
    }

    #[test]
    fn rem_is_a_comment_only_at_the_start_of_a_statement() {
        assert!(idents("Rem this is a comment").is_empty());
        assert_eq!(idents("x: Rem note"), vec!["x"]);
        assert_eq!(idents("Remainder = 1"), vec!["Remainder"]);
        assert_eq!(idents("x = Rem"), vec!["x", "Rem"]);
    }

    #[test]
    fn strings_resolve_doubled_quotes() {
        assert_eq!(kinds(r#""a""b""#)[0], TokenKind::Str(r#"a"b"#.into()));
    }

    #[test]
    fn an_unterminated_string_is_an_error_not_a_swallowed_module() {
        let err = lex("x = \"oops\ny = 1").unwrap_err();
        assert!(err.message.contains("unterminated"));
        assert_eq!(err.pos.line, 1);
    }

    #[test]
    fn numbers_cover_the_shapes_vba_writes() {
        assert_eq!(num("42"), 42.0);
        assert_eq!(num("3.5"), 3.5);
        assert_eq!(num(".5"), 0.5);
        assert_eq!(num("1E3"), 1000.0);
        assert_eq!(num("1.5e-3"), 0.0015);
        assert_eq!(num("1D2"), 100.0);
        assert_eq!(num("&HFF"), 255.0);
        assert_eq!(num("&O17"), 15.0);
        assert_eq!(num("&hff"), 255.0);
    }

    #[test]
    fn a_based_literal_needs_digits() {
        assert!(lex("&H").is_err());
        assert!(lex("&HG").is_err());
    }

    #[test]
    fn type_suffixes_attach_rather_than_becoming_operators() {
        match &kinds("42&")[0] {
            TokenKind::Number { suffix, .. } => assert_eq!(*suffix, Some(TypeSuffix::Long)),
            other => panic!("{other:?}"),
        }
        assert_eq!(idents("name$ = \"x\""), vec!["name$"]);
        assert_eq!(idents("count% = 1"), vec!["count%"]);
    }

    #[test]
    fn ampersand_between_operands_stays_concatenation() {
        assert!(kinds("a & b").contains(&TokenKind::Punct("&")));
        assert!(kinds("a$ & b$").contains(&TokenKind::Punct("&")));
        assert_eq!(idents("a$ & b$"), vec!["a$", "b$"]);
    }

    #[test]
    fn bang_before_a_name_is_dictionary_access_not_a_single_suffix() {
        assert!(kinds("rs!Field").contains(&TokenKind::Punct("!")));
        assert_eq!(idents("rs!Field"), vec!["rs", "Field"]);
    }

    #[test]
    fn hash_is_a_date_a_suffix_and_a_directive() {
        assert_eq!(kinds("#1/1/2000#")[0], TokenKind::Date("1/1/2000".into()));
        match &kinds("x#")[0] {
            TokenKind::Ident(name) => assert_eq!(name, "x#"),
            other => panic!("{other:?}"),
        }
        assert_eq!(kinds("#If")[0], TokenKind::Punct("#"));
        assert_eq!(kinds("#If")[1], TokenKind::Ident("If".into()));
    }

    #[test]
    fn multi_char_operators_win_over_their_prefixes() {
        assert_eq!(kinds("a <= b")[1], TokenKind::Punct("<="));
        assert_eq!(kinds("a <> b")[1], TokenKind::Punct("<>"));
        assert_eq!(kinds("a >= b")[1], TokenKind::Punct(">="));
        assert_eq!(kinds("a =< b")[1], TokenKind::Punct("<="));
        assert_eq!(kinds("a => b")[1], TokenKind::Punct(">="));
    }

    #[test]
    fn keywords_keep_their_spelling_and_are_not_reserved() {
        assert_eq!(idents("ws.Name = \"x\""), vec!["ws", "Name"]);
    }

    #[test]
    fn space_before_distinguishes_a_call_from_an_index() {
        let toks = lex("Foo (a)").unwrap();
        assert!(toks[1].preceded_by_space);
        let toks = lex("Foo(a)").unwrap();
        assert!(!toks[1].preceded_by_space);
    }

    #[test]
    fn positions_are_one_based_and_track_lines() {
        let toks = lex("a\n  b").unwrap();
        assert_eq!(toks[0].pos, Pos { line: 1, col: 1 });
        assert_eq!(toks[2].pos, Pos { line: 2, col: 3 });
    }

    #[test]
    fn never_panics_on_odd_input() {
        for src in [
            "",
            "\0",
            "\"",
            "#",
            "&",
            "_",
            ":",
            "\r",
            "é",
            "#\n#",
            "&H",
            "1e",
            "1e+",
            ".",
            "'unterminated comment",
            "a _",
            "\u{feff}x",
        ] {
            let _ = lex(src);
        }
    }
}
