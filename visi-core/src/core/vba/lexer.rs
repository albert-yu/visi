//! Tokenizer for VBA source text.
//!
//! Never panics and never allocates unboundedly: it is reachable from
//! `visi-core/fuzz`'s `vba_parse` target over arbitrary bytes, and from
//! `visi macro check` over whatever a user pasted into a module.
//!
//! Three things about VBA make this more than a keyword scanner, and each is
//! a place a naive lexer gets it wrong:
//!
//! **Newlines are significant.** VBA is line-oriented -- a statement ends at
//! the end of a line, not at a delimiter -- so [`TokenKind::Newline`] is a
//! real token the parser matches on, not whitespace to skip. `:` is the
//! explicit statement separator and is emitted as a distinct token, since
//! `Foo: Bar` (two statements) and `Foo:` alone (a line label) differ only in
//! what follows.
//!
//! **A line can be continued.** ` _` at end of line splices the next line on.
//! The underscore must be preceded by whitespace and followed only by the
//! line break, which is what keeps it from swallowing an identifier ending
//! in `_`.
//!
//! **Keywords are not reserved here.** They are lexed as plain
//! [`TokenKind::Ident`]s preserving their original spelling, and the parser
//! matches them case-insensitively. VBA's keyword set is contextual (`Line`,
//! `Name`, and `Type` are all keywords in one position and ordinary
//! identifiers in another), so a lexer that promoted them to distinct tokens
//! would have to un-promote them again constantly.

use std::fmt;

/// A source position, 1-based in both axes so it can be printed as-is.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Pos {
    /// 1-based line number.
    pub line: u32,
    /// 1-based column number, counted in characters rather than bytes.
    pub col: u32,
}

impl fmt::Display for Pos {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}:{}", self.line, self.col)
    }
}

/// The numeric base a literal was written in, kept so a round trip can tell
/// `&H10` from `16`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumBase {
    /// Ordinary decimal, possibly with a fraction and exponent.
    Decimal,
    /// `&H`-prefixed hexadecimal.
    Hex,
    /// `&O`-prefixed (or bare `&`-prefixed) octal.
    Octal,
}

/// A VBA type-declaration character: the trailing sigil in `count%`, `name$`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TypeSuffix {
    /// `$` -- String.
    String,
    /// `%` -- Integer.
    Integer,
    /// `&` -- Long.
    Long,
    /// `!` -- Single.
    Single,
    /// `#` -- Double.
    Double,
    /// `@` -- Currency.
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

    /// The character this suffix is written as.
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

/// What a token is.
#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    /// An identifier or a keyword, in its original spelling. Keywords are not
    /// distinguished here -- see this module's docs.
    Ident(String),
    /// A numeric literal.
    Number {
        /// The value. Hex/octal literals are already converted.
        value: f64,
        /// How it was written.
        base: NumBase,
        /// A trailing type-declaration character, if any.
        suffix: Option<TypeSuffix>,
        /// Whether it was written with a decimal point or an exponent, which
        /// forces `Double` regardless of the value: `1E3` is a `Double` even
        /// though the same value written `1000` is a `Long`.
        is_float: bool,
    },
    /// A string literal, with `""` escapes already resolved to `"`.
    Str(String),
    /// A `#...#` date literal, holding the raw text between the hashes. Not
    /// parsed into a serial here: that is `date.rs`'s job and it needs the
    /// workbook's date system, which the lexer has no business knowing.
    Date(String),
    /// Punctuation or an operator, as its canonical spelling (`"<="`, `"&"`).
    /// Word operators (`And`, `Mod`, `Is`) arrive as `Ident` instead.
    Punct(&'static str),
    /// An end of line, which in VBA ends a statement.
    Newline,
    /// End of input.
    Eof,
}

/// A token and where it came from.
#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    /// What the token is.
    pub kind: TokenKind,
    /// Where it starts.
    pub pos: Pos,
    /// Whether whitespace (or a line continuation) preceded it on this line.
    /// The parser needs this to tell `Foo (a)` -- a call whose one argument
    /// is parenthesised -- from `Foo(a)`, and the lexer needs it to tell a
    /// type suffix from an operator.
    pub preceded_by_space: bool,
}

impl Token {
    /// The identifier text, if this is an `Ident`.
    pub fn ident(&self) -> Option<&str> {
        match &self.kind {
            TokenKind::Ident(s) => Some(s),
            _ => None,
        }
    }

    /// Whether this is the given keyword, compared case-insensitively as VBA
    /// does.
    pub fn is_kw(&self, kw: &str) -> bool {
        self.ident().is_some_and(|s| s.eq_ignore_ascii_case(kw))
    }

    /// Whether this is the given punctuation.
    pub fn is_punct(&self, p: &str) -> bool {
        matches!(&self.kind, TokenKind::Punct(x) if *x == p)
    }
}

/// A lexing failure, with the position of the offending character.
#[derive(Debug, Clone, PartialEq)]
pub struct LexError {
    /// What went wrong, phrased for a user reading CLI output.
    pub message: String,
    /// Where it went wrong.
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

/// Multi-character operators, longest first so `<=` wins over `<`.
///
/// `:=` is lexed as one token rather than `:` + `=` so that a line label
/// (`Failed:`) and a named argument (`Foo bar:=1`) cannot be confused: both
/// start with an identifier followed by a colon.
const MULTI_PUNCT: &[&str] = &[":=", "<=", ">=", "<>", "=<", "=>"];
const SINGLE_PUNCT: &[char] = &[
    '(', ')', ',', '.', '=', '+', '-', '*', '/', '\\', '^', '&', '<', '>', ':', ';', '!', '#', '$',
    '%', '@', '?', '{', '}', '[', ']', '~', '|',
];

/// Splits VBA source into tokens.
///
/// Returns every token including a final [`TokenKind::Eof`]. Comments are
/// discarded (VBA has no doc-comment convention that the parser needs), but
/// the newline that ends a comment is kept, since it still ends a statement.
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

    /// Whether the token just emitted can carry a trailing type suffix.
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
                    // Swallow CR so CRLF (what real .bas files carry) emits
                    // exactly one Newline.
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
                // A `_` starting a token that is *not* a continuation would
                // begin an identifier, and VBA has none: a name must start
                // with a letter (the same rule `validate_vba_module_name`
                // enforces for modules). Measured -- real Excel refuses to
                // compile `_ y = 2`, while a trailing `_` continuation is
                // fine (`fuzz/vba_compile_probe.py --only continuation`).
                // Left as an identifier, it silently became an implicit-call
                // statement on a name spelled `_`, which is issue #78's
                // iter_24 false negative.
                //
                // Note `is_ident_start` still admits `_`: it answers a
                // different question at `suffix_would_be_operator`, where a
                // following continuation must keep `&` an operator rather
                // than turn it into a type suffix.
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

    /// A `_` is a continuation only when whitespace precedes it and nothing
    /// but whitespace follows it on the line. Otherwise it is part of an
    /// identifier (`my_var`) -- it can never *start* one, since VBA has no
    /// name beginning with an underscore (measured; see the `'_'` arm in
    /// [`Lexer::run`]).
    fn is_line_continuation(&self) -> bool {
        // Something like `a_` must not split: the underscore has to be its
        // own token position, i.e. preceded by whitespace or line start.
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
        // Trailing `_` at end of input: treat as a continuation of nothing,
        // which is harmless and avoids an identifier named `_`.
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
        self.bump(); // opening quote
        let mut s = String::new();
        loop {
            match self.peek() {
                // A string literal cannot span lines; an unterminated one is
                // an error rather than a silent swallow of the rest of the
                // module.
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

    /// `#` is three different things: a type suffix (`x#`), a date literal
    /// (`#1/1/2000#`), and the lead-in to a compiler directive (`#If`).
    fn lex_hash(&mut self) -> Result<(), LexError> {
        let pos = self.pos();

        if !self.space_before && self.last_takes_suffix() {
            self.bump();
            self.attach_suffix(TypeSuffix::Double);
            return Ok(());
        }

        // A date literal needs a closing `#` on the same line. Anything else
        // (a directive, a stray hash) falls through to punctuation, so `#If`
        // reaches the parser as `#` + `If`.
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

    /// Rewrites the token just emitted to carry a type suffix.
    fn attach_suffix(&mut self, suffix: TypeSuffix) {
        if let Some(last) = self.out.last_mut() {
            match &mut last.kind {
                TokenKind::Number { suffix: slot, .. } => *slot = Some(suffix),
                // An identifier's suffix is part of its name in VBA (`a$` and
                // `a` are the same variable, but `a$` is how it was written),
                // so fold it back into the spelling rather than dropping it.
                TokenKind::Ident(name) => name.push(suffix.as_char()),
                _ => {}
            }
        }
    }

    fn starts_based_number(&self) -> bool {
        // `&H1F` / `&O17`, but `a & b` is concatenation. Also require that
        // this `&` isn't a type suffix, which `lex_punct` handles.
        if !self.space_before && self.last_takes_suffix() {
            return false;
        }
        matches!(self.peek_at(1), Some(c) if c == 'h' || c == 'H' || c == 'o' || c == 'O')
    }

    fn lex_based_number(&mut self) -> Result<(), LexError> {
        let pos = self.pos();
        self.bump(); // &
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
        // VBA wraps &H literals into signed 16/32-bit, but that is a value
        // question rather than a syntax one; Phase 0 only needs the digits to
        // be well-formed. u128 keeps an absurdly long literal from wrapping
        // silently here.
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
        // A based literal may still carry `&` (Long) as a suffix: `&HFF&`.
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
            // A leading `.5`.
            self.bump();
            while self.peek().is_some_and(|c| c.is_ascii_digit()) {
                self.bump();
            }
        }
        // `1E5`, `1.5e-3`, and VBA's `D` exponent for Double literals.
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

        // `Rem` is a comment keyword, but only as a whole word starting a
        // statement -- `Remainder` is an ordinary identifier.
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

    /// Whether the token about to be emitted is the first of a statement.
    fn starts_statement(&self) -> bool {
        match self.out.last().map(|t| &t.kind) {
            None | Some(TokenKind::Newline) => true,
            Some(TokenKind::Punct(p)) => *p == ":",
            _ => false,
        }
    }

    /// A sigil directly after an identifier or number is a type suffix unless
    /// it reads as an operator instead.
    ///
    /// `&` and `!` are the two that overlap: `a & b` concatenates, `a$ & b$`
    /// concatenates two suffixed names, `rs!Field` is a dictionary access.
    /// The rule is that a suffix cannot be followed by something that would
    /// begin an operand, which is what separates `a& = 1` from `a&b` and
    /// `rs!F`.
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

        // A sigil reaching here directly after an identifier/number is a type
        // suffix, not an operator.
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
                // `=<` and `=>` are accepted spellings of `<=` and `>=`.
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

/// `&'static str` spellings parallel to [`SINGLE_PUNCT`].
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
        // Measured against real Excel, which refuses to compile `_ y = 2`
        // (`fuzz/vba_compile_probe.py --only continuation`). This test used
        // to assert the opposite -- that `_leading` lexed as an identifier --
        // which is what let issue #78's iter_24 through as a false negative.
        assert!(lex("_leading\n").is_err());
        assert!(lex("_ y = 2\n").is_err());
        // A *trailing* `_` is a real continuation and stays one.
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
        // Not a comment: an ordinary name that merely begins with "rem".
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
        // VBA's Double exponent marker.
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
        // The overlap that a naive "sigil after ident is a suffix" rule gets
        // wrong: `a & b` and `a$ & b$` both concatenate.
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
        // `#If` must reach the parser as punctuation plus a name.
        assert_eq!(kinds("#If")[0], TokenKind::Punct("#"));
        assert_eq!(kinds("#If")[1], TokenKind::Ident("If".into()));
    }

    #[test]
    fn multi_char_operators_win_over_their_prefixes() {
        assert_eq!(kinds("a <= b")[1], TokenKind::Punct("<="));
        assert_eq!(kinds("a <> b")[1], TokenKind::Punct("<>"));
        assert_eq!(kinds("a >= b")[1], TokenKind::Punct(">="));
        // VBA also accepts the reversed spellings, canonicalised here.
        assert_eq!(kinds("a =< b")[1], TokenKind::Punct("<="));
        assert_eq!(kinds("a => b")[1], TokenKind::Punct(">="));
    }

    #[test]
    fn keywords_keep_their_spelling_and_are_not_reserved() {
        // `Name` is a keyword in `Name x As y` and a property everywhere else;
        // the lexer must not decide which.
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
