//! Recursive-descent parser from VBA tokens to [`ast::Module`].
//!
//! Phase 0 of the plan in `docs/vba-macro-support.md`: this decides whether
//! source *parses*, and nothing else. No name resolution, no types, no
//! evaluation. A `Call` node does not know whether it is a procedure call or
//! an array index, because telling those apart needs a symbol table.
//!
//! # Precedence
//!
//! The operator table below was confirmed against real Excel (16.112) rather
//! than taken from documentation, via a macro returning each result -- the
//! two cases that actually discriminate between plausible tables are
//! `False Imp False Eqv False`, which is `True` only if `Eqv` binds tighter
//! than `Imp`, and `True Xor True Eqv False`, likewise for `Xor` over `Eqv`.
//! Loosest to tightest:
//!
//! | Level | Operators | Confirming case |
//! | --- | --- | --- |
//! | 1 | `Imp` | `False Imp False Eqv False` = `True` |
//! | 2 | `Eqv` | `True Xor True Eqv False` = `True` |
//! | 3 | `Xor` | as above |
//! | 4 | `Or` | `False And False Or True` = `True` |
//! | 5 | `And` | as above |
//! | 6 | `Not` (prefix) | `Not 1 = 0` = `True`, i.e. `Not (1 = 0)` |
//! | 7 | `=` `<>` `<` `>` `<=` `>=` `Is` `Like` | `1 = 1 And 1 = 0` = `False` |
//! | 8 | `&` | `2 + 3 & 4` = `"54"` |
//! | 9 | `+` `-` | `1 + 7 Mod 3` = `2` |
//! | 10 | `Mod` | `20 \ 3 Mod 4` = `2` |
//! | 11 | `\` | `2 * 10 \ 3` = `6` |
//! | 12 | `*` `/` | as above |
//! | 13 | unary `-` `+` | `-2 ^ 2` = `-4`, i.e. `-(2 ^ 2)` |
//! | 14 | `^` | `2 ^ 3 ^ 2` = `64`, i.e. **left**-associative |
//!
//! `^` being left-associative is the one most likely to be got wrong: most
//! languages with a `^`/`**` operator make it right-associative, which would
//! give `512` here.

use super::ast::*;
use super::lexer::{LexError, NumBase, Pos, Token, TokenKind, lex};
use std::fmt;

/// A syntax error, with the position to point a user at.
#[derive(Debug, Clone, PartialEq)]
pub struct ParseError {
    /// What went wrong, phrased for someone reading CLI output.
    pub message: String,
    /// Where it went wrong.
    pub pos: Pos,
}

impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{} at line {}, column {}",
            self.message, self.pos.line, self.pos.col
        )
    }
}

impl std::error::Error for ParseError {}

impl From<LexError> for ParseError {
    fn from(e: LexError) -> Self {
        Self {
            message: e.message,
            pos: e.pos,
        }
    }
}

/// Parses a complete VBA module.
///
/// ```
/// use visi_core::core::vba::parser::parse_module;
/// let m = parse_module("Public Sub Hello()\n    MsgBox \"hi\"\nEnd Sub\n").unwrap();
/// assert_eq!(m.procedures()[0].name, "Hello");
/// ```
pub fn parse_module(src: &str) -> Result<Module, ParseError> {
    let tokens = lex(src)?;
    Parser::new(tokens).parse_module()
}

/// Keywords that, at the start of a statement, close an enclosing block.
fn closes_block(t: &Token, next: &Token) -> bool {
    if t.is_kw("end") {
        return next.is_kw("sub")
            || next.is_kw("function")
            || next.is_kw("property")
            || next.is_kw("if")
            || next.is_kw("with")
            || next.is_kw("select")
            || next.is_kw("type")
            || next.is_kw("enum");
    }
    t.is_kw("loop")
        || t.is_kw("next")
        || t.is_kw("wend")
        || t.is_kw("else")
        || t.is_kw("elseif")
        || t.is_kw("case")
}

/// Statement keywords whose syntax is irregular enough that Phase 0 records
/// them as [`Stmt::Opaque`] rather than modelling them. See that variant.
const OPAQUE_IO_KEYWORDS: &[&str] = &[
    "print", "write", "input", "put", "get", "seek", "lock", "unlock", "width",
];

/// The built-in scalar type keywords, which unlike most VBA keywords
/// (contextual and freely reusable as identifiers -- see this module's own
/// doc comment on that) cannot themselves be used as a declared name: a
/// `Dim`/`Const` variable or a parameter. Measured directly against real
/// Excel (Windows), found via `fuzz/fuzz_vba_parse.py`: `Dim Long As
/// Integer`, `Const Long = 5`, and a parameter named `Long` all fail to
/// compile, and the same holds for every other name here. `Object` is
/// deliberately not in this list -- `Dim Object As Long` compiles fine,
/// even though `Object` is itself a valid type in an `As` clause; real
/// Excel does not treat it as reserved the way it treats these.
const RESERVED_TYPE_NAMES: &[&str] = &[
    "boolean", "byte", "currency", "date", "double", "integer", "long", "single", "string",
    "variant",
];

struct Parser {
    toks: Vec<Token>,
    i: usize,
    /// How many further `Next`s a `Next i, j` still has to close. VBA lets one
    /// `Next` close several nested `For`s; the innermost loop consumes the
    /// token and leaves this counter for its enclosing loops to drain.
    pending_next: usize,
}

impl Parser {
    fn new(toks: Vec<Token>) -> Self {
        Self {
            toks,
            i: 0,
            pending_next: 0,
        }
    }

    // ---- token access ---------------------------------------------------

    fn peek(&self) -> &Token {
        // `lex` always terminates the stream with Eof, so this never wraps.
        self.toks.get(self.i).unwrap_or_else(|| self.eof_token())
    }

    fn at(&self, n: usize) -> &Token {
        self.toks
            .get(self.i + n)
            .unwrap_or_else(|| self.eof_token())
    }

    fn eof_token(&self) -> &Token {
        self.toks.last().expect("lex always emits Eof")
    }

    fn pos(&self) -> Pos {
        self.peek().pos
    }

    fn is_eof(&self) -> bool {
        matches!(self.peek().kind, TokenKind::Eof)
    }

    fn bump(&mut self) -> Token {
        let t = self.peek().clone();
        if !matches!(t.kind, TokenKind::Eof) {
            self.i += 1;
        }
        t
    }

    fn eat_punct(&mut self, p: &str) -> bool {
        if self.peek().is_punct(p) {
            self.i += 1;
            true
        } else {
            false
        }
    }

    fn eat_kw(&mut self, kw: &str) -> bool {
        if self.peek().is_kw(kw) {
            self.i += 1;
            true
        } else {
            false
        }
    }

    fn expect_punct(&mut self, p: &str) -> Result<(), ParseError> {
        if self.eat_punct(p) {
            Ok(())
        } else {
            Err(self.expected(&format!("{p:?}")))
        }
    }

    fn expect_kw(&mut self, kw: &str) -> Result<(), ParseError> {
        if self.eat_kw(kw) {
            Ok(())
        } else {
            Err(self.expected(kw))
        }
    }

    fn expect_ident(&mut self) -> Result<(String, Pos), ParseError> {
        let pos = self.pos();
        match self.peek().ident() {
            Some(name) => {
                let name = name.to_string();
                self.i += 1;
                Ok((name, pos))
            }
            None => Err(self.expected("a name")),
        }
    }

    /// Like [`Self::expect_ident`], but for a name that is being *declared*
    /// (a `Dim`/`Const` variable, a parameter) rather than merely
    /// referenced -- those positions additionally reject
    /// [`RESERVED_TYPE_NAMES`], which real Excel refuses to compile there.
    fn expect_declarable_ident(&mut self) -> Result<(String, Pos), ParseError> {
        let (name, pos) = self.expect_ident()?;
        if RESERVED_TYPE_NAMES.contains(&name.to_ascii_lowercase().as_str()) {
            return Err(ParseError {
                message: format!("'{name}' is a built-in type name and cannot be declared"),
                pos,
            });
        }
        Ok((name, pos))
    }

    fn expected(&self, what: &str) -> ParseError {
        ParseError {
            message: format!("expected {what}, found {}", describe(self.peek())),
            pos: self.pos(),
        }
    }

    fn error(&self, message: impl Into<String>) -> ParseError {
        ParseError {
            message: message.into(),
            pos: self.pos(),
        }
    }

    /// A block that was opened and never closed.
    ///
    /// Reported at the **opener**, not at whatever turned up where the closer
    /// was due, and worded the way VBA's own compiler words it ("Block If
    /// without End If", "For without Next", ...). Pointing at the opener is
    /// the more useful of the two: in
    ///
    /// ```text
    /// Sub S()
    ///     If x Then
    /// End Sub
    /// ```
    ///
    /// the defect is the `If` on line 2, and blaming the `End Sub` on line 3
    /// sends the reader to a line that is perfectly correct. It matters most
    /// in a long procedure, where the two can be hundreds of lines apart.
    ///
    /// The wording and the choice of position follow VBA's documented
    /// compile errors; they could not be read back from Excel directly,
    /// because a compile error surfaces only as a modal dialog, which is
    /// unreadable to the AppleScript bridge (it is the hang the fuzz harness
    /// works around) and, in the environment this was written in, to UI
    /// scripting and screenshots as well. Worth re-checking by hand in the
    /// VBE if these strings ever matter for more than readability.
    fn unclosed(&self, what: &str, pos: Pos) -> ParseError {
        ParseError {
            message: what.to_string(),
            pos,
        }
    }

    // ---- separators -----------------------------------------------------

    fn skip_newlines(&mut self) {
        while matches!(self.peek().kind, TokenKind::Newline) || self.peek().is_punct(":") {
            self.i += 1;
        }
    }

    fn at_stmt_end(&self) -> bool {
        matches!(self.peek().kind, TokenKind::Newline | TokenKind::Eof) || self.peek().is_punct(":")
    }

    fn expect_stmt_end(&mut self) -> Result<(), ParseError> {
        if self.at_stmt_end() {
            if !self.is_eof() {
                self.i += 1;
            }
            Ok(())
        } else {
            Err(self.expected("end of statement"))
        }
    }

    /// Consumes tokens to the end of the current statement, tracking nesting
    /// so a `:` or newline inside parentheses does not end it early.
    fn skip_to_stmt_end(&mut self) {
        let mut depth = 0usize;
        loop {
            let t = self.peek();
            match &t.kind {
                TokenKind::Eof => return,
                TokenKind::Newline if depth == 0 => return,
                TokenKind::Punct("(") => depth += 1,
                TokenKind::Punct(")") => depth = depth.saturating_sub(1),
                TokenKind::Punct(":") if depth == 0 => return,
                _ => {}
            }
            self.i += 1;
        }
    }

    // ---- module ---------------------------------------------------------

    fn parse_module(mut self) -> Result<Module, ParseError> {
        let items = self.parse_module_items(false)?;
        if !self.is_eof() {
            return Err(self.error(format!("unexpected {}", describe(self.peek()))));
        }
        Ok(Module { items })
    }

    fn parse_module_items(&mut self, nested: bool) -> Result<Vec<ModuleItem>, ParseError> {
        let mut items = Vec::new();
        loop {
            self.skip_newlines();
            if self.is_eof() {
                break;
            }
            if nested && self.at_conditional_close() {
                break;
            }
            items.push(self.parse_module_item()?);
        }
        Ok(items)
    }

    fn at_conditional_close(&self) -> bool {
        self.peek().is_punct("#")
            && (self.at(1).is_kw("else") || self.at(1).is_kw("elseif") || self.at(1).is_kw("end"))
    }

    fn parse_module_item(&mut self) -> Result<ModuleItem, ParseError> {
        let pos = self.pos();

        if self.peek().is_kw("attribute") {
            self.i += 1;
            let (name, _) = self.parse_dotted_name()?;
            self.expect_punct("=")?;
            let mut values = vec![self.parse_expr()?];
            while self.eat_punct(",") {
                values.push(self.parse_expr()?);
            }
            self.expect_stmt_end()?;
            return Ok(ModuleItem::Attribute { name, values, pos });
        }

        if self.peek().is_kw("option") {
            self.i += 1;
            let mut words = Vec::new();
            while !self.at_stmt_end() {
                let t = self.bump();
                words.push(match t.kind {
                    TokenKind::Ident(s) => s,
                    TokenKind::Number { value, .. } => format!("{value}"),
                    other => {
                        return Err(ParseError {
                            message: format!("unexpected {} in Option", describe_kind(&other)),
                            pos: t.pos,
                        });
                    }
                });
            }
            self.expect_stmt_end()?;
            if words.is_empty() {
                return Err(ParseError {
                    message: "Option needs a setting, e.g. Option Explicit".to_string(),
                    pos,
                });
            }
            return Ok(ModuleItem::Option { words, pos });
        }

        if self.peek().is_punct("#") {
            return self.parse_conditional();
        }

        // A visibility keyword can lead either a procedure or a declaration,
        // so it is read here and handed to whichever follows.
        let visibility = self.eat_visibility();
        let is_static = self.eat_kw("static");

        if self.peek().is_kw("sub")
            || self.peek().is_kw("function")
            || self.peek().is_kw("property")
        {
            return Ok(ModuleItem::Procedure(
                self.parse_procedure(visibility, is_static, pos)?,
            ));
        }

        if self.peek().is_kw("declare") {
            return Ok(ModuleItem::Declaration(self.parse_declare(pos)?));
        }

        let stmt = self.parse_declaration_stmt(visibility, is_static, pos)?;
        self.expect_stmt_end()?;
        Ok(ModuleItem::Declaration(stmt))
    }

    /// `#If` / `#ElseIf` / `#Else` / `#End If`, and `#Const`.
    fn parse_conditional(&mut self) -> Result<ModuleItem, ParseError> {
        let pos = self.pos();
        self.expect_punct("#")?;

        if self.eat_kw("const") {
            let vars = self.parse_var_decls(true)?;
            self.expect_stmt_end()?;
            return Ok(ModuleItem::Declaration(Stmt::Const {
                visibility: None,
                vars,
                pos,
            }));
        }

        self.expect_kw("if")?;
        let mut branches = Vec::new();
        let cond = self.parse_expr()?;
        self.expect_kw("then")?;
        self.expect_stmt_end()?;
        branches.push((cond, self.parse_module_items(true)?));

        let mut else_items = None;
        loop {
            if !self.peek().is_punct("#") {
                return Err(self.expected("#End If"));
            }
            self.i += 1;
            if self.eat_kw("elseif") {
                let cond = self.parse_expr()?;
                self.expect_kw("then")?;
                self.expect_stmt_end()?;
                branches.push((cond, self.parse_module_items(true)?));
            } else if self.eat_kw("else") {
                self.expect_stmt_end()?;
                else_items = Some(self.parse_module_items(true)?);
            } else if self.eat_kw("end") {
                self.expect_kw("if")?;
                if !self.is_eof() {
                    self.expect_stmt_end()?;
                }
                break;
            } else {
                return Err(self.expected("#ElseIf, #Else or #End If"));
            }
        }

        Ok(ModuleItem::Conditional {
            branches,
            else_items,
            pos,
        })
    }

    fn eat_visibility(&mut self) -> Option<Visibility> {
        let v = if self.peek().is_kw("public") {
            Visibility::Public
        } else if self.peek().is_kw("private") {
            Visibility::Private
        } else if self.peek().is_kw("friend") {
            Visibility::Friend
        } else if self.peek().is_kw("global") {
            Visibility::Global
        } else {
            return None;
        };
        // `Private` also leads `Private Module` in an Option line, but that
        // is handled before this is reached.
        self.i += 1;
        Some(v)
    }

    // ---- procedures -----------------------------------------------------

    fn parse_procedure(
        &mut self,
        visibility: Option<Visibility>,
        is_static: bool,
        pos: Pos,
    ) -> Result<Procedure, ParseError> {
        let kind = if self.eat_kw("sub") {
            ProcKind::Sub
        } else if self.eat_kw("function") {
            ProcKind::Function
        } else {
            self.expect_kw("property")?;
            if self.eat_kw("get") {
                ProcKind::PropertyGet
            } else if self.eat_kw("let") {
                ProcKind::PropertyLet
            } else if self.eat_kw("set") {
                ProcKind::PropertySet
            } else {
                return Err(self.expected("Get, Let or Set after Property"));
            }
        };

        let (name, _) = self.expect_ident()?;
        let params = if self.peek().is_punct("(") {
            self.parse_param_list()?
        } else {
            Vec::new()
        };
        let return_type = if self.eat_kw("as") {
            Some(self.parse_type_ref()?)
        } else {
            None
        };
        self.expect_stmt_end()?;

        let body = self.parse_block()?;

        let closer = match kind {
            ProcKind::Sub => "Sub",
            ProcKind::Function => "Function",
            _ => "Property",
        };
        if !(self.peek().is_kw("end") && self.at(1).is_kw(closer)) {
            return Err(self.unclosed(&format!("Expected End {closer}"), pos));
        }
        self.i += 2;
        if !self.is_eof() {
            self.expect_stmt_end()?;
        }

        Ok(Procedure {
            kind,
            name,
            visibility,
            is_static,
            params,
            return_type,
            body,
            pos,
        })
    }

    fn parse_param_list(&mut self) -> Result<Vec<Param>, ParseError> {
        self.expect_punct("(")?;
        let mut params = Vec::new();
        if self.eat_punct(")") {
            return Ok(params);
        }
        loop {
            params.push(self.parse_param()?);
            if self.eat_punct(",") {
                continue;
            }
            self.expect_punct(")")?;
            break;
        }
        Ok(params)
    }

    fn parse_param(&mut self) -> Result<Param, ParseError> {
        let optional = self.eat_kw("optional");
        let param_array = self.eat_kw("paramarray");
        let by = if self.eat_kw("byval") {
            Some(PassBy::Value)
        } else if self.eat_kw("byref") {
            Some(PassBy::Reference)
        } else {
            None
        };
        // `Optional ByVal x` -- either order appears in real code.
        let optional = optional || self.eat_kw("optional");

        let (name, _) = self.expect_declarable_ident()?;
        let is_array = if self.peek().is_punct("(") && self.at(1).is_punct(")") {
            self.i += 2;
            true
        } else {
            false
        };
        let ty = if self.eat_kw("as") {
            Some(self.parse_type_ref()?)
        } else {
            None
        };
        let default = if self.eat_punct("=") {
            Some(self.parse_expr()?)
        } else {
            None
        };
        Ok(Param {
            name,
            by,
            optional,
            param_array,
            is_array,
            ty,
            default,
        })
    }

    fn parse_type_ref(&mut self) -> Result<TypeRef, ParseError> {
        let is_new = self.eat_kw("new");
        let mut path = vec![self.expect_ident()?.0];
        while self.peek().is_punct(".") && self.at(1).ident().is_some() {
            self.i += 1;
            path.push(self.expect_ident()?.0);
        }
        // `As String * 10` -- a fixed-length string.
        let string_length = if self.peek().is_punct("*") {
            self.i += 1;
            Some(Box::new(self.parse_expr()?))
        } else {
            None
        };
        Ok(TypeRef {
            path,
            is_new,
            string_length,
        })
    }

    fn parse_dotted_name(&mut self) -> Result<(String, Pos), ParseError> {
        let (first, pos) = self.expect_ident()?;
        let mut name = first;
        while self.peek().is_punct(".") && self.at(1).ident().is_some() {
            self.i += 1;
            name.push('.');
            name.push_str(&self.expect_ident()?.0);
        }
        Ok((name, pos))
    }

    // ---- statements -----------------------------------------------------

    fn parse_block(&mut self) -> Result<Vec<Stmt>, ParseError> {
        let mut out = Vec::new();
        loop {
            self.skip_newlines();
            if self.is_eof() || self.pending_next > 0 {
                break;
            }
            if closes_block(self.peek(), self.at(1)) {
                break;
            }
            let stmt = self.parse_stmt()?;
            let is_label = matches!(stmt, Stmt::Label { .. });
            out.push(stmt);
            if is_label {
                // A label's `:` already terminated it; a statement may follow
                // on the same line.
                continue;
            }
            if self.at_stmt_end() {
                self.expect_stmt_end()?;
            } else if !closes_block(self.peek(), self.at(1)) {
                return Err(self.expected("end of statement"));
            }
        }
        Ok(out)
    }

    fn parse_stmt(&mut self) -> Result<Stmt, ParseError> {
        let pos = self.pos();
        let t = self.peek().clone();

        // A line number is a jump target, exactly like a label.
        if let TokenKind::Number { value, base, .. } = &t.kind
            && *base == NumBase::Decimal
            && value.fract() == 0.0
            && *value >= 0.0
        {
            self.i += 1;
            return Ok(Stmt::Label {
                name: format!("{value}"),
                pos,
            });
        }

        // `Failed:` -- a label. `:=` lexes as one token, so a named argument
        // cannot be mistaken for one.
        if t.ident().is_some() && self.at(1).is_punct(":") {
            let (name, _) = self.expect_ident()?;
            self.expect_punct(":")?;
            return Ok(Stmt::Label { name, pos });
        }

        if let Some(kw) = t.ident() {
            let lower = kw.to_ascii_lowercase();
            match lower.as_str() {
                "dim" | "static" | "private" | "public" | "global" => {
                    let visibility = self.eat_visibility();
                    let is_static = visibility.is_none() && self.eat_kw("static");
                    return self.parse_declaration_stmt(visibility, is_static, pos);
                }
                "const" => {
                    self.i += 1;
                    let vars = self.parse_var_decls(true)?;
                    return Ok(Stmt::Const {
                        visibility: None,
                        vars,
                        pos,
                    });
                }
                "redim" => {
                    self.i += 1;
                    let preserve = self.eat_kw("preserve");
                    let vars = self.parse_var_decls(false)?;
                    return Ok(Stmt::ReDim {
                        preserve,
                        vars,
                        pos,
                    });
                }
                "if" => return self.parse_if(pos),
                "select" => return self.parse_select(pos),
                "for" => return self.parse_for(pos),
                "do" => return self.parse_do(pos),
                "while" => return self.parse_while(pos),
                "with" => return self.parse_with(pos),
                "exit" => return self.parse_exit(pos),
                "goto" => {
                    self.i += 1;
                    let label = self.parse_label_ref()?;
                    return Ok(Stmt::GoTo { label, pos });
                }
                "gosub" => {
                    self.i += 1;
                    let label = self.parse_label_ref()?;
                    return Ok(Stmt::GoSub { label, pos });
                }
                "return" => {
                    self.i += 1;
                    return Ok(Stmt::Return { pos });
                }
                "on" => return self.parse_on(pos),
                "resume" => {
                    self.i += 1;
                    let kind = if self.eat_kw("next") {
                        ResumeKind::Next
                    } else if self.at_stmt_end() {
                        ResumeKind::Retry
                    } else {
                        ResumeKind::Label(self.parse_label_ref()?)
                    };
                    return Ok(Stmt::Resume { kind, pos });
                }
                "erase" => {
                    self.i += 1;
                    let mut targets = vec![self.parse_expr()?];
                    while self.eat_punct(",") {
                        targets.push(self.parse_expr()?);
                    }
                    return Ok(Stmt::Erase { targets, pos });
                }
                "type" => return self.parse_type_def(None, pos),
                "enum" => return self.parse_enum_def(None, pos),
                "declare" => return self.parse_declare(pos),
                "event" => {
                    self.i += 1;
                    let (name, _) = self.expect_ident()?;
                    let params = self.parse_param_list()?;
                    return Ok(Stmt::EventDef { name, params, pos });
                }
                "raiseevent" => {
                    self.i += 1;
                    let (name, _) = self.expect_ident()?;
                    let args = if self.peek().is_punct("(") {
                        self.parse_paren_args()?
                    } else {
                        Vec::new()
                    };
                    return Ok(Stmt::RaiseEvent { name, args, pos });
                }
                "implements" => {
                    self.i += 1;
                    let (name, _) = self.parse_dotted_name()?;
                    return Ok(Stmt::Implements { name, pos });
                }
                "stop" => {
                    self.i += 1;
                    return Ok(Stmt::Stop { pos });
                }
                "end" => {
                    self.i += 1;
                    // `End Sub` and friends are consumed by their block; an
                    // `End` reaching here is the halt statement.
                    return Ok(Stmt::End { pos });
                }
                "set" => {
                    self.i += 1;
                    // `parse_postfix`, not `parse_expr`: `=` is a comparison
                    // operator in an expression, so parsing the target with
                    // the full grammar would swallow `a = b` whole and then
                    // find no `=` left for the assignment.
                    let target = self.parse_postfix()?;
                    self.expect_punct("=")?;
                    let value = self.parse_expr()?;
                    return Ok(Stmt::Assign {
                        target,
                        value,
                        set: true,
                        pos,
                    });
                }
                "let" => {
                    self.i += 1;
                    let target = self.parse_postfix()?;
                    self.expect_punct("=")?;
                    let value = self.parse_expr()?;
                    return Ok(Stmt::Assign {
                        target,
                        value,
                        set: false,
                        pos,
                    });
                }
                "call" => {
                    self.i += 1;
                    let expr = self.parse_expr()?;
                    return Ok(Stmt::Call { expr, pos });
                }
                "attribute" => {
                    self.i += 1;
                    let (name, _) = self.parse_dotted_name()?;
                    self.expect_punct("=")?;
                    let mut values = vec![self.parse_expr()?];
                    while self.eat_punct(",") {
                        values.push(self.parse_expr()?);
                    }
                    return Ok(Stmt::Attribute { name, values, pos });
                }
                _ => {
                    if let Some(stmt) = self.try_parse_opaque(&lower, pos) {
                        return Ok(stmt);
                    }
                }
            }
        }

        self.parse_assign_or_call(pos)
    }

    /// The file-I/O and legacy statements Phase 0 records verbatim.
    ///
    /// Each guard is deliberately narrow, because every one of these words is
    /// also an ordinary identifier or method name: `Print` is a statement only
    /// before a `#`, `Line` only before `Input`, `Name` only in
    /// `Name a As b`. Without the guards, `Application.Width = 100` or a
    /// variable called `Get` would stop parsing.
    fn try_parse_opaque(&mut self, lower: &str, pos: Pos) -> Option<Stmt> {
        let matched = match lower {
            "open" => true,
            "close" | "reset" => {
                self.at(1).is_punct("#")
                    || matches!(self.at(1).kind, TokenKind::Newline | TokenKind::Eof)
            }
            "line" => self.at(1).is_kw("input"),
            "name" => self.statement_has_kw("as"),
            _ if OPAQUE_IO_KEYWORDS.contains(&lower) => self.at(1).is_punct("#"),
            _ => false,
        };
        if !matched {
            return None;
        }
        let keyword = self.peek().ident().unwrap_or_default().to_string();
        self.skip_to_stmt_end();
        Some(Stmt::Opaque { keyword, pos })
    }

    /// Whether the rest of this statement contains the given keyword outside
    /// parentheses.
    fn statement_has_kw(&self, kw: &str) -> bool {
        let mut j = self.i;
        let mut depth = 0usize;
        while let Some(t) = self.toks.get(j) {
            match &t.kind {
                TokenKind::Eof | TokenKind::Newline => return false,
                TokenKind::Punct("(") => depth += 1,
                TokenKind::Punct(")") => depth = depth.saturating_sub(1),
                TokenKind::Punct(":") if depth == 0 => return false,
                _ if depth == 0 && t.is_kw(kw) => return true,
                _ => {}
            }
            j += 1;
        }
        false
    }

    /// A jump target, which VBA lets be either a name or a line number.
    fn parse_label_ref(&mut self) -> Result<String, ParseError> {
        match &self.peek().kind {
            TokenKind::Ident(s) => {
                let s = s.clone();
                self.i += 1;
                Ok(s)
            }
            TokenKind::Number { value, .. } => {
                let v = *value;
                self.i += 1;
                Ok(format!("{v}"))
            }
            _ => Err(self.expected("a label or line number")),
        }
    }

    fn parse_declaration_stmt(
        &mut self,
        visibility: Option<Visibility>,
        is_static: bool,
        pos: Pos,
    ) -> Result<Stmt, ParseError> {
        if self.peek().is_kw("type") {
            return self.parse_type_def(visibility, pos);
        }
        if self.peek().is_kw("enum") {
            return self.parse_enum_def(visibility, pos);
        }
        if self.peek().is_kw("event") {
            self.i += 1;
            let (name, _) = self.expect_ident()?;
            let params = self.parse_param_list()?;
            return Ok(Stmt::EventDef { name, params, pos });
        }
        if self.eat_kw("const") {
            let vars = self.parse_var_decls(true)?;
            return Ok(Stmt::Const {
                visibility,
                vars,
                pos,
            });
        }

        let kind = if self.eat_kw("dim") {
            DimKind::Dim
        } else if is_static || self.eat_kw("static") {
            DimKind::Static
        } else {
            match visibility {
                Some(Visibility::Private) => DimKind::Private,
                Some(Visibility::Public) => DimKind::Public,
                Some(Visibility::Global) => DimKind::Global,
                _ => return Err(self.expected("a declaration")),
            }
        };
        let with_events = self.eat_kw("withevents");
        let vars = self.parse_var_decls(false)?;
        Ok(Stmt::Dim {
            kind,
            with_events,
            vars,
            pos,
        })
    }

    fn parse_var_decls(&mut self, allow_value: bool) -> Result<Vec<VarDecl>, ParseError> {
        let mut vars = vec![self.parse_var_decl(allow_value)?];
        while self.eat_punct(",") {
            vars.push(self.parse_var_decl(allow_value)?);
        }
        Ok(vars)
    }

    fn parse_var_decl(&mut self, allow_value: bool) -> Result<VarDecl, ParseError> {
        let (name, pos) = self.expect_declarable_ident()?;
        let bounds = if self.peek().is_punct("(") {
            self.i += 1;
            if self.eat_punct(")") {
                // `x()` -- a dynamic array, distinct from not an array.
                Some(Vec::new())
            } else {
                let mut list = Vec::new();
                loop {
                    let first = self.parse_expr()?;
                    let bound = if self.eat_kw("to") {
                        ArrayBound {
                            lower: Some(first),
                            upper: self.parse_expr()?,
                        }
                    } else {
                        ArrayBound {
                            lower: None,
                            upper: first,
                        }
                    };
                    list.push(bound);
                    if !self.eat_punct(",") {
                        break;
                    }
                }
                self.expect_punct(")")?;
                Some(list)
            }
        } else {
            None
        };
        let ty = if self.eat_kw("as") {
            Some(self.parse_type_ref()?)
        } else {
            None
        };
        let value = if allow_value && self.eat_punct("=") {
            Some(self.parse_expr()?)
        } else {
            None
        };
        Ok(VarDecl {
            name,
            bounds,
            ty,
            value,
            pos,
        })
    }

    fn parse_type_def(
        &mut self,
        visibility: Option<Visibility>,
        pos: Pos,
    ) -> Result<Stmt, ParseError> {
        self.expect_kw("type")?;
        let (name, _) = self.expect_ident()?;
        self.expect_stmt_end()?;
        let mut fields = Vec::new();
        loop {
            self.skip_newlines();
            if self.peek().is_kw("end") && self.at(1).is_kw("type") {
                break;
            }
            if self.is_eof() {
                return Err(self.unclosed("Expected End Type", pos));
            }
            fields.push(self.parse_var_decl(false)?);
            self.expect_stmt_end()?;
        }
        self.i += 2;
        Ok(Stmt::TypeDef {
            name,
            visibility,
            fields,
            pos,
        })
    }

    fn parse_enum_def(
        &mut self,
        visibility: Option<Visibility>,
        pos: Pos,
    ) -> Result<Stmt, ParseError> {
        self.expect_kw("enum")?;
        let (name, _) = self.expect_ident()?;
        self.expect_stmt_end()?;
        let mut members = Vec::new();
        loop {
            self.skip_newlines();
            if self.peek().is_kw("end") && self.at(1).is_kw("enum") {
                break;
            }
            if self.is_eof() {
                return Err(self.unclosed("Expected End Enum", pos));
            }
            let (member, _) = self.expect_ident()?;
            let value = if self.eat_punct("=") {
                Some(self.parse_expr()?)
            } else {
                None
            };
            members.push(EnumMember {
                name: member,
                value,
            });
            self.expect_stmt_end()?;
        }
        self.i += 2;
        Ok(Stmt::EnumDef {
            name,
            visibility,
            members,
            pos,
        })
    }

    fn parse_declare(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("declare")?;
        // 64-bit Office writes `Declare PtrSafe`.
        self.eat_kw("ptrsafe");
        let is_function = if self.eat_kw("function") {
            true
        } else {
            self.expect_kw("sub")?;
            false
        };
        let (name, _) = self.expect_ident()?;
        self.expect_kw("lib")?;
        let lib = self.expect_string()?;
        let alias = if self.eat_kw("alias") {
            Some(self.expect_string()?)
        } else {
            None
        };
        let params = if self.peek().is_punct("(") {
            self.parse_param_list()?
        } else {
            Vec::new()
        };
        let return_type = if self.eat_kw("as") {
            Some(self.parse_type_ref()?)
        } else {
            None
        };
        Ok(Stmt::Declare {
            name,
            is_function,
            lib,
            alias,
            params,
            return_type,
            pos,
        })
    }

    fn expect_string(&mut self) -> Result<String, ParseError> {
        match &self.peek().kind {
            TokenKind::Str(s) => {
                let s = s.clone();
                self.i += 1;
                Ok(s)
            }
            _ => Err(self.expected("a string literal")),
        }
    }

    fn parse_if(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("if")?;
        let cond = self.parse_expr()?;
        self.expect_kw("then")?;

        // Single-line form: something other than a line break follows `Then`.
        if !self.at_stmt_end() {
            let then_body = self.parse_inline_stmts()?;
            let else_body = if self.eat_kw("else") {
                Some(self.parse_inline_stmts()?)
            } else {
                None
            };
            return Ok(Stmt::If {
                branches: vec![(cond, then_body)],
                else_body,
                pos,
            });
        }

        self.expect_stmt_end()?;
        let mut branches = vec![(cond, self.parse_block()?)];
        let mut else_body = None;
        loop {
            if self.eat_kw("elseif") {
                let cond = self.parse_expr()?;
                self.expect_kw("then")?;
                self.expect_stmt_end()?;
                branches.push((cond, self.parse_block()?));
            } else if self.eat_kw("else") {
                self.expect_stmt_end()?;
                else_body = Some(self.parse_block()?);
            } else if self.peek().is_kw("end") && self.at(1).is_kw("if") {
                self.i += 2;
                break;
            } else {
                return Err(self.unclosed("Block If without End If", pos));
            }
        }
        Ok(Stmt::If {
            branches,
            else_body,
            pos,
        })
    }

    /// Statements on one line, as in the single-line `If` form, stopping at
    /// `Else` or the line break.
    fn parse_inline_stmts(&mut self) -> Result<Vec<Stmt>, ParseError> {
        let mut out = Vec::new();
        loop {
            if self.at_stmt_end() || self.peek().is_kw("else") {
                break;
            }
            out.push(self.parse_stmt()?);
            if self.peek().is_punct(":") {
                self.i += 1;
                continue;
            }
            break;
        }
        Ok(out)
    }

    fn parse_select(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("select")?;
        self.expect_kw("case")?;
        let subject = self.parse_expr()?;
        self.expect_stmt_end()?;
        self.skip_newlines();

        let mut cases = Vec::new();
        let mut case_else = None;
        loop {
            if self.peek().is_kw("end") && self.at(1).is_kw("select") {
                self.i += 2;
                break;
            }
            if !self.eat_kw("case") {
                return Err(self.unclosed("Select Case without End Select", pos));
            }
            if self.eat_kw("else") {
                self.expect_stmt_end()?;
                case_else = Some(self.parse_block()?);
                continue;
            }
            let mut matches = Vec::new();
            loop {
                matches.push(self.parse_case_match()?);
                if !self.eat_punct(",") {
                    break;
                }
            }
            self.expect_stmt_end()?;
            let body = self.parse_block()?;
            cases.push(CaseClause { matches, body });
        }

        Ok(Stmt::SelectCase {
            subject,
            cases,
            case_else,
            pos,
        })
    }

    fn parse_case_match(&mut self) -> Result<CaseMatch, ParseError> {
        if self.eat_kw("is") {
            let op = self
                .eat_comparison_op()
                .ok_or_else(|| self.expected("a comparison operator after Is"))?;
            return Ok(CaseMatch::Is(op, self.parse_expr()?));
        }
        let first = self.parse_expr()?;
        if self.eat_kw("to") {
            return Ok(CaseMatch::Range(first, self.parse_expr()?));
        }
        Ok(CaseMatch::Value(first))
    }

    fn eat_comparison_op(&mut self) -> Option<BinOp> {
        let op = match &self.peek().kind {
            TokenKind::Punct("=") => BinOp::Eq,
            TokenKind::Punct("<>") => BinOp::Ne,
            TokenKind::Punct("<") => BinOp::Lt,
            TokenKind::Punct(">") => BinOp::Gt,
            TokenKind::Punct("<=") => BinOp::Le,
            TokenKind::Punct(">=") => BinOp::Ge,
            _ => return None,
        };
        self.i += 1;
        Some(op)
    }

    fn parse_for(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("for")?;

        if self.eat_kw("each") {
            let var = self.parse_postfix()?;
            self.expect_kw("in")?;
            let iterable = self.parse_expr()?;
            self.expect_stmt_end()?;
            let body = self.parse_block()?;
            self.finish_next(pos)?;
            return Ok(Stmt::ForEach {
                var,
                iterable,
                body,
                pos,
            });
        }

        let var = self.parse_postfix()?;
        self.expect_punct("=")?;
        let from = self.parse_expr()?;
        self.expect_kw("to")?;
        let to = self.parse_expr()?;
        let step = if self.eat_kw("step") {
            Some(self.parse_expr()?)
        } else {
            None
        };
        self.expect_stmt_end()?;
        let body = self.parse_block()?;
        self.finish_next(pos)?;
        Ok(Stmt::For {
            var,
            from,
            to,
            step,
            body,
            pos,
        })
    }

    /// Consumes the `Next` closing a loop, handling `Next i, j` closing
    /// several at once by leaving a count for the enclosing loops.
    fn finish_next(&mut self, opener: Pos) -> Result<(), ParseError> {
        if self.pending_next > 0 {
            self.pending_next -= 1;
            return Ok(());
        }
        if !self.eat_kw("next") {
            return Err(self.unclosed("For without Next", opener));
        }
        let mut names = 0usize;
        while !self.at_stmt_end() {
            self.parse_postfix()?;
            names += 1;
            if !self.eat_punct(",") {
                break;
            }
        }
        self.pending_next = names.saturating_sub(1);
        Ok(())
    }

    fn parse_do(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("do")?;
        let pre = self.eat_do_test()?;
        self.expect_stmt_end()?;
        let body = self.parse_block()?;
        if !self.eat_kw("loop") {
            return Err(self.unclosed("Do without Loop", pos));
        }
        let post = if pre.is_none() {
            self.eat_do_test()?
        } else {
            None
        };
        Ok(Stmt::DoLoop {
            pre,
            post,
            body,
            pos,
        })
    }

    fn eat_do_test(&mut self) -> Result<Option<(DoTest, Expr)>, ParseError> {
        if self.eat_kw("while") {
            Ok(Some((DoTest::While, self.parse_expr()?)))
        } else if self.eat_kw("until") {
            Ok(Some((DoTest::Until, self.parse_expr()?)))
        } else {
            Ok(None)
        }
    }

    /// `While ... Wend`, normalised into the equivalent `Do While ... Loop`.
    fn parse_while(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("while")?;
        let cond = self.parse_expr()?;
        self.expect_stmt_end()?;
        let body = self.parse_block()?;
        if !self.eat_kw("wend") {
            return Err(self.unclosed("While without Wend", pos));
        }
        Ok(Stmt::DoLoop {
            pre: Some((DoTest::While, cond)),
            post: None,
            body,
            pos,
        })
    }

    fn parse_with(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("with")?;
        let subject = self.parse_expr()?;
        self.expect_stmt_end()?;
        let body = self.parse_block()?;
        if !(self.peek().is_kw("end") && self.at(1).is_kw("with")) {
            return Err(self.unclosed("With without End With", pos));
        }
        self.i += 2;
        Ok(Stmt::With { subject, body, pos })
    }

    fn parse_exit(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("exit")?;
        let kind = if self.eat_kw("sub") {
            ExitKind::Sub
        } else if self.eat_kw("function") {
            ExitKind::Function
        } else if self.eat_kw("property") {
            ExitKind::Property
        } else if self.eat_kw("for") {
            ExitKind::For
        } else if self.eat_kw("do") {
            ExitKind::Do
        } else if self.eat_kw("while") {
            ExitKind::While
        } else {
            return Err(self.expected("Sub, Function, Property, For, Do or While after Exit"));
        };
        Ok(Stmt::Exit { kind, pos })
    }

    fn parse_on(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        self.expect_kw("on")?;

        if self.eat_kw("error") {
            let kind = if self.eat_kw("resume") {
                self.expect_kw("next")?;
                OnErrorKind::ResumeNext
            } else if self.eat_kw("goto") {
                let label = self.parse_label_ref()?;
                // `On Error GoTo 0` turns the handler off rather than jumping
                // to a label named "0".
                if label == "0" {
                    OnErrorKind::Disable
                } else {
                    OnErrorKind::GoTo(label)
                }
            } else {
                return Err(self.expected("Resume Next or GoTo after On Error"));
            };
            return Ok(Stmt::OnError { kind, pos });
        }

        let subject = self.parse_expr()?;
        let gosub = if self.eat_kw("goto") {
            false
        } else if self.eat_kw("gosub") {
            true
        } else {
            return Err(self.expected("GoTo or GoSub"));
        };
        let mut labels = vec![self.parse_label_ref()?];
        while self.eat_punct(",") {
            labels.push(self.parse_label_ref()?);
        }
        Ok(Stmt::OnGoto {
            subject,
            labels,
            gosub,
            pos,
        })
    }

    /// An assignment, or a call with or without parentheses.
    ///
    /// The two are told apart only after parsing the left side: `x = 1` is an
    /// assignment, `MsgBox "hi"` is a call whose arguments are bare, and
    /// `Foo` alone is a call with none.
    fn parse_assign_or_call(&mut self, pos: Pos) -> Result<Stmt, ParseError> {
        let target = self.parse_postfix()?;

        if self.peek().is_punct("=") {
            self.i += 1;
            let value = self.parse_expr()?;
            return Ok(Stmt::Assign {
                target,
                value,
                set: false,
                pos,
            });
        }

        if self.at_stmt_end() || self.peek().is_kw("else") {
            return Ok(Stmt::Call { expr: target, pos });
        }

        // A `Print` method's output list is its own grammar, not an
        // argument list -- see `parse_print_output_list`.
        let args = if is_print_member(&target) {
            self.parse_print_output_list()?
        } else {
            // `Debug.Print a, b` -- arguments without parentheses.
            let mut args = Vec::new();
            loop {
                args.push(self.parse_arg()?);
                if !self.eat_punct(",") {
                    break;
                }
            }
            args
        };
        Ok(Stmt::Call {
            expr: Expr::Call {
                target: Box::new(target),
                args,
                pos,
            },
            pos,
        })
    }

    /// A `Print` method's output list: `Debug.Print "a"; 1`.
    ///
    /// This is deliberately **not** the bare-argument list above. VBA gives
    /// `Print` its own grammar (MS-VBAL's `output-item = [output-clause]
    /// [output-item-separator]`, with both halves optional), and every way
    /// it differs was measured with `fuzz/vba_compile_probe.py` against
    /// real Excel rather than read off the spec:
    ///
    /// - `;` separates items as `,` does (`Debug.Print "a"; 1`).
    /// - A **trailing** separator is legal, and meaningful -- it suppresses
    ///   the newline: `Debug.Print "a";`, `Debug.Print "a",`.
    /// - So is a leading or a repeated one, which prints an empty item:
    ///   `Debug.Print , "a"`, `Debug.Print ; "a"`, `Debug.Print "a";; "b"`.
    /// - A separator may be omitted between two items entirely:
    ///   `Debug.Print "a" "b"` compiles.
    ///
    /// Which separator was written is not recorded. It only decides output
    /// spacing, and `Print` is outside the interpreter's scope by the
    /// security posture in `docs/vba-macro-support.md`, so the only
    /// question this path answers is `macro check`'s: does it compile.
    fn parse_print_output_list(&mut self) -> Result<Vec<Arg>, ParseError> {
        let mut args = Vec::new();
        while !self.at_stmt_end() && !self.peek().is_kw("else") {
            let arg = if self.peek().is_punct(";") || self.peek().is_punct(",") {
                // An empty item, when a separator comes first.
                Arg {
                    name: None,
                    value: None,
                }
            } else if self.peek().ident().is_some() && self.at(1).is_punct(":=") {
                // Nothing says the `Print` here is VBA's: a class module may
                // define one, and a named argument to it parsed before this
                // path existed. Keeping it is the same no-false-positives
                // rule that motivated the rest of this function.
                self.parse_arg()?
            } else {
                Arg {
                    name: None,
                    value: Some(self.parse_expr()?),
                }
            };
            args.push(arg);
            // At most one separator, and it may be absent altogether.
            let _ = self.eat_punct(";") || self.eat_punct(",");
        }
        Ok(args)
    }

    // ---- expressions ----------------------------------------------------

    /// Entry point; see this module's docs for the precedence table.
    fn parse_expr(&mut self) -> Result<Expr, ParseError> {
        self.parse_imp()
    }

    fn parse_imp(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_eqv()?;
        while self.peek().is_kw("imp") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_eqv()?;
            lhs = binary(BinOp::Imp, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_eqv(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_xor()?;
        while self.peek().is_kw("eqv") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_xor()?;
            lhs = binary(BinOp::Eqv, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_xor(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_or()?;
        while self.peek().is_kw("xor") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_or()?;
            lhs = binary(BinOp::Xor, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_or(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_and()?;
        while self.peek().is_kw("or") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_and()?;
            lhs = binary(BinOp::Or, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_and(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_not()?;
        while self.peek().is_kw("and") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_not()?;
            lhs = binary(BinOp::And, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    /// `Not` binds looser than comparison, so `Not a = b` is `Not (a = b)`.
    fn parse_not(&mut self) -> Result<Expr, ParseError> {
        if self.peek().is_kw("not") {
            let pos = self.pos();
            self.i += 1;
            let expr = self.parse_not()?;
            return Ok(Expr::Unary {
                op: UnOp::Not,
                expr: Box::new(expr),
                pos,
            });
        }
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_concat()?;
        loop {
            let pos = self.pos();
            let op = if let Some(op) = self.eat_comparison_op() {
                op
            } else if self.peek().is_kw("is") {
                self.i += 1;
                BinOp::Is
            } else if self.peek().is_kw("like") {
                self.i += 1;
                BinOp::Like
            } else {
                break;
            };
            let rhs = self.parse_concat()?;
            lhs = binary(op, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_concat(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_additive()?;
        while self.peek().is_punct("&") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_additive()?;
            lhs = binary(BinOp::Concat, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_additive(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_modulo()?;
        loop {
            let pos = self.pos();
            let op = if self.peek().is_punct("+") {
                BinOp::Add
            } else if self.peek().is_punct("-") {
                BinOp::Sub
            } else {
                break;
            };
            self.i += 1;
            let rhs = self.parse_modulo()?;
            lhs = binary(op, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_modulo(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_int_div()?;
        while self.peek().is_kw("mod") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_int_div()?;
            lhs = binary(BinOp::Mod, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_int_div(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_multiplicative()?;
        while self.peek().is_punct("\\") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_multiplicative()?;
            lhs = binary(BinOp::IntDiv, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_multiplicative(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_unary()?;
        loop {
            let pos = self.pos();
            let op = if self.peek().is_punct("*") {
                BinOp::Mul
            } else if self.peek().is_punct("/") {
                BinOp::Div
            } else {
                break;
            };
            self.i += 1;
            let rhs = self.parse_unary()?;
            lhs = binary(op, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    /// Unary sign binds *looser* than `^`, which is why `-2 ^ 2` is `-4`.
    fn parse_unary(&mut self) -> Result<Expr, ParseError> {
        let pos = self.pos();
        if self.peek().is_punct("-") {
            self.i += 1;
            let expr = self.parse_unary()?;
            return Ok(Expr::Unary {
                op: UnOp::Neg,
                expr: Box::new(expr),
                pos,
            });
        }
        if self.peek().is_punct("+") {
            self.i += 1;
            let expr = self.parse_unary()?;
            return Ok(Expr::Unary {
                op: UnOp::Pos,
                expr: Box::new(expr),
                pos,
            });
        }
        self.parse_pow()
    }

    /// `^`, left-associative -- `2 ^ 3 ^ 2` is `64`, not `512`.
    ///
    /// The right operand is [`Self::parse_pow_operand`] rather than
    /// [`Self::parse_pow`], which is what keeps it left-associative while
    /// still allowing a signed exponent (`2 ^ -3`).
    fn parse_pow(&mut self) -> Result<Expr, ParseError> {
        let mut lhs = self.parse_postfix()?;
        while self.peek().is_punct("^") {
            let pos = self.pos();
            self.i += 1;
            let rhs = self.parse_pow_operand()?;
            lhs = binary(BinOp::Pow, lhs, rhs, pos);
        }
        Ok(lhs)
    }

    fn parse_pow_operand(&mut self) -> Result<Expr, ParseError> {
        let pos = self.pos();
        if self.peek().is_punct("-") {
            self.i += 1;
            let expr = self.parse_pow_operand()?;
            return Ok(Expr::Unary {
                op: UnOp::Neg,
                expr: Box::new(expr),
                pos,
            });
        }
        if self.peek().is_punct("+") {
            self.i += 1;
            let expr = self.parse_pow_operand()?;
            return Ok(Expr::Unary {
                op: UnOp::Pos,
                expr: Box::new(expr),
                pos,
            });
        }
        self.parse_postfix()
    }

    fn parse_postfix(&mut self) -> Result<Expr, ParseError> {
        let mut expr = self.parse_primary()?;
        loop {
            let pos = self.pos();
            if self.peek().is_punct(".") {
                self.i += 1;
                let (name, _) = self.expect_ident()?;
                expr = Expr::Member {
                    target: Some(Box::new(expr)),
                    name,
                    pos,
                };
            } else if self.peek().is_punct("!") {
                self.i += 1;
                let (name, _) = self.expect_ident()?;
                expr = Expr::Bang {
                    target: Box::new(expr),
                    name,
                    pos,
                };
            } else if self.peek().is_punct("(") {
                let args = self.parse_paren_args()?;
                expr = Expr::Call {
                    target: Box::new(expr),
                    args,
                    pos,
                };
            } else {
                break;
            }
        }
        Ok(expr)
    }

    fn parse_paren_args(&mut self) -> Result<Vec<Arg>, ParseError> {
        self.expect_punct("(")?;
        let mut args = Vec::new();
        if self.eat_punct(")") {
            return Ok(args);
        }
        loop {
            args.push(self.parse_arg()?);
            if self.eat_punct(",") {
                continue;
            }
            self.expect_punct(")")?;
            break;
        }
        Ok(args)
    }

    fn parse_arg(&mut self) -> Result<Arg, ParseError> {
        // An omitted positional argument: `f(1, , 3)`. Not the same as
        // passing Empty, so it is modelled as an absent value.
        if self.peek().is_punct(",") || self.peek().is_punct(")") {
            return Ok(Arg {
                name: None,
                value: None,
            });
        }
        if self.peek().ident().is_some() && self.at(1).is_punct(":=") {
            let (name, _) = self.expect_ident()?;
            self.i += 1;
            return Ok(Arg {
                name: Some(name),
                value: Some(self.parse_expr()?),
            });
        }
        Ok(Arg {
            name: None,
            value: Some(self.parse_expr()?),
        })
    }

    fn parse_primary(&mut self) -> Result<Expr, ParseError> {
        let pos = self.pos();
        let t = self.peek().clone();

        match &t.kind {
            TokenKind::Number {
                value,
                base,
                suffix,
                is_float,
            } => {
                self.i += 1;
                Ok(Expr::Literal(Literal::Number {
                    value: *value,
                    base: *base,
                    suffix: *suffix,
                    is_float: *is_float,
                }))
            }
            TokenKind::Str(s) => {
                self.i += 1;
                Ok(Expr::Literal(Literal::Str(s.clone())))
            }
            TokenKind::Date(d) => {
                self.i += 1;
                Ok(Expr::Literal(Literal::Date(d.clone())))
            }
            TokenKind::Punct("(") => {
                self.i += 1;
                let inner = self.parse_expr()?;
                self.expect_punct(")")?;
                Ok(Expr::Paren {
                    expr: Box::new(inner),
                    pos,
                })
            }
            // A leading `.` inside a `With` block.
            TokenKind::Punct(".") => {
                self.i += 1;
                let (name, _) = self.expect_ident()?;
                Ok(Expr::Member {
                    target: None,
                    name,
                    pos,
                })
            }
            TokenKind::Ident(name) => {
                let lower = name.to_ascii_lowercase();
                match lower.as_str() {
                    "true" => {
                        self.i += 1;
                        Ok(Expr::Literal(Literal::Bool(true)))
                    }
                    "false" => {
                        self.i += 1;
                        Ok(Expr::Literal(Literal::Bool(false)))
                    }
                    "nothing" => {
                        self.i += 1;
                        Ok(Expr::Literal(Literal::Nothing))
                    }
                    "empty" => {
                        self.i += 1;
                        Ok(Expr::Literal(Literal::Empty))
                    }
                    "null" => {
                        self.i += 1;
                        Ok(Expr::Literal(Literal::Null))
                    }
                    "me" => {
                        self.i += 1;
                        Ok(Expr::Me { pos })
                    }
                    "new" => {
                        self.i += 1;
                        Ok(Expr::New {
                            ty: self.parse_type_ref()?,
                            pos,
                        })
                    }
                    "typeof" => {
                        self.i += 1;
                        let expr = self.parse_postfix()?;
                        self.expect_kw("is")?;
                        Ok(Expr::TypeOf {
                            expr: Box::new(expr),
                            ty: self.parse_type_ref()?,
                            pos,
                        })
                    }
                    "addressof" => {
                        self.i += 1;
                        let (name, _) = self.parse_dotted_name()?;
                        Ok(Expr::AddressOf { name, pos })
                    }
                    _ => {
                        self.i += 1;
                        Ok(Expr::Ident {
                            name: name.clone(),
                            pos,
                        })
                    }
                }
            }
            _ => Err(self.expected("an expression")),
        }
    }
}

/// Whether a bare-argument statement's target is a `Print` method, and so
/// takes an output list rather than an argument list.
///
/// The gate is the *member name*, measured both ways: `x.Print "a"; 1`
/// compiles for an object that is not `Debug`, while `Debug.Assert "a"; 1`
/// does not, and neither does a bare `Print "a"; 1` -- unqualified `Print`
/// is a statement only before a `#`, which `try_parse_opaque` already takes.
fn is_print_member(target: &Expr) -> bool {
    matches!(target, Expr::Member { name, .. } if name.eq_ignore_ascii_case("print"))
}

fn binary(op: BinOp, lhs: Expr, rhs: Expr, pos: Pos) -> Expr {
    Expr::Binary {
        op,
        lhs: Box::new(lhs),
        rhs: Box::new(rhs),
        pos,
    }
}

fn describe(t: &Token) -> String {
    describe_kind(&t.kind)
}

fn describe_kind(k: &TokenKind) -> String {
    match k {
        TokenKind::Ident(s) => format!("{s:?}"),
        TokenKind::Number { value, .. } => format!("the number {value}"),
        TokenKind::Str(_) => "a string literal".to_string(),
        TokenKind::Date(_) => "a date literal".to_string(),
        TokenKind::Punct(p) => format!("{p:?}"),
        TokenKind::Newline => "end of line".to_string(),
        TokenKind::Eof => "end of file".to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(src: &str) -> Module {
        parse_module(src).unwrap_or_else(|e| panic!("{e}\n--- source ---\n{src}"))
    }

    fn expr(src: &str) -> Expr {
        // Wrap in a body so the statement parser is exercised the same way it
        // is in real code.
        let m = parse(&format!("Sub S()\n    x = {src}\nEnd Sub\n"));
        match &m.procedures()[0].body[0] {
            Stmt::Assign { value, .. } => value.clone(),
            other => panic!("{other:?}"),
        }
    }

    /// Renders an expression as a fully-parenthesised string, so a
    /// precedence test states the tree shape rather than poking at it.
    fn shape(e: &Expr) -> String {
        match e {
            Expr::Literal(Literal::Number { value, .. }) => format!("{value}"),
            Expr::Literal(Literal::Str(s)) => format!("{s:?}"),
            Expr::Literal(Literal::Bool(b)) => b.to_string(),
            Expr::Literal(l) => format!("{l:?}"),
            Expr::Ident { name, .. } => name.clone(),
            Expr::Me { .. } => "Me".to_string(),
            Expr::Member { target, name, .. } => match target {
                Some(t) => format!("{}.{name}", shape(t)),
                None => format!(".{name}"),
            },
            Expr::Bang { target, name, .. } => format!("{}!{name}", shape(target)),
            Expr::Call { target, args, .. } => {
                let inner: Vec<String> = args
                    .iter()
                    .map(|a| match (&a.name, &a.value) {
                        (Some(n), Some(v)) => format!("{n}:={}", shape(v)),
                        (None, Some(v)) => shape(v),
                        _ => "_".to_string(),
                    })
                    .collect();
                format!("{}({})", shape(target), inner.join(", "))
            }
            Expr::Unary { op, expr, .. } => format!("({op:?} {})", shape(expr)),
            Expr::Binary { op, lhs, rhs, .. } => {
                format!("({:?} {} {})", op, shape(lhs), shape(rhs))
            }
            Expr::Paren { expr, .. } => format!("[{}]", shape(expr)),
            Expr::New { ty, .. } => format!("New {}", ty.path.join(".")),
            Expr::TypeOf { expr, ty, .. } => {
                format!("(TypeOf {} Is {})", shape(expr), ty.path.join("."))
            }
            Expr::AddressOf { name, .. } => format!("AddressOf {name}"),
        }
    }

    // ---- precedence, as confirmed against real Excel --------------------
    //
    // Each case below is one this parser could plausibly get wrong, and each
    // was run through Excel 16.112 to get the answer rather than assumed.
    // The comment gives the value Excel produced.

    #[test]
    fn pow_is_left_associative() {
        // Excel: 2 ^ 3 ^ 2 = 64. Right-associativity would give 512.
        assert_eq!(shape(&expr("2 ^ 3 ^ 2")), "(Pow (Pow 2 3) 2)");
    }

    #[test]
    fn pow_binds_tighter_than_unary_minus() {
        // Excel: -2 ^ 2 = -4.
        assert_eq!(shape(&expr("-2 ^ 2")), "(Neg (Pow 2 2))");
    }

    #[test]
    fn a_signed_exponent_still_parses() {
        assert_eq!(shape(&expr("2 ^ -3")), "(Pow 2 (Neg 3))");
    }

    #[test]
    fn concat_binds_looser_than_addition() {
        // Excel: 2 + 3 & 4 = "54".
        assert_eq!(shape(&expr("2 + 3 & 4")), "(Concat (Add 2 3) 4)");
    }

    #[test]
    fn comparison_binds_tighter_than_and() {
        // Excel: 1 = 1 And 1 = 0 -> False.
        assert_eq!(shape(&expr("1 = 1 And 1 = 0")), "(And (Eq 1 1) (Eq 1 0))");
    }

    #[test]
    fn not_binds_looser_than_comparison() {
        // Excel: Not 1 = 0 -> True, i.e. Not (1 = 0).
        assert_eq!(shape(&expr("Not 1 = 0")), "(Not (Eq 1 0))");
    }

    #[test]
    fn the_three_division_operators_nest_correctly() {
        // Excel: 2 * 10 \ 3 = 6, 10 \ 3 * 2 = 1, 20 \ 3 Mod 4 = 2,
        //        1 + 7 Mod 3 = 2.
        assert_eq!(shape(&expr("2 * 10 \\ 3")), "(IntDiv (Mul 2 10) 3)");
        assert_eq!(shape(&expr("10 \\ 3 * 2")), "(IntDiv 10 (Mul 3 2))");
        assert_eq!(shape(&expr("20 \\ 3 Mod 4")), "(Mod (IntDiv 20 3) 4)");
        assert_eq!(shape(&expr("1 + 7 Mod 3")), "(Add 1 (Mod 7 3))");
    }

    #[test]
    fn the_logical_operators_nest_correctly() {
        // Excel: True Xor True Eqv False -> True (so Xor binds tighter),
        //        False And False Or True -> True,
        //        False Imp False Eqv False -> True (so Eqv binds tighter).
        assert_eq!(
            shape(&expr("True Xor True Eqv False")),
            "(Eqv (Xor true true) false)"
        );
        assert_eq!(
            shape(&expr("False And False Or True")),
            "(Or (And false false) true)"
        );
        assert_eq!(
            shape(&expr("False Imp False Eqv False")),
            "(Imp false (Eqv false false))"
        );
    }

    #[test]
    fn arithmetic_is_left_associative() {
        // Excel: 2 - 3 - 4 = -5, 8 / 4 / 2 = 1.
        assert_eq!(shape(&expr("2 - 3 - 4")), "(Sub (Sub 2 3) 4)");
        assert_eq!(shape(&expr("8 / 4 / 2")), "(Div (Div 8 4) 2)");
    }

    // ---- expressions ----------------------------------------------------

    #[test]
    fn member_index_and_dictionary_access_chain() {
        assert_eq!(
            shape(&expr("ws.Range(\"A1\").Value")),
            "ws.Range(\"A1\").Value"
        );
        assert_eq!(shape(&expr("rs!Field.Name")), "rs!Field.Name");
    }

    #[test]
    fn parentheses_are_kept_because_they_force_by_value() {
        assert_eq!(shape(&expr("(a)")), "[a]");
    }

    #[test]
    fn named_and_omitted_arguments() {
        assert_eq!(shape(&expr("f(1, , 3)")), "f(1, _, 3)");
        assert_eq!(shape(&expr("f(Key:=1)")), "f(Key:=1)");
    }

    #[test]
    fn object_expressions() {
        assert_eq!(shape(&expr("New Collection")), "New Collection");
        assert_eq!(
            shape(&expr("TypeOf x Is Worksheet")),
            "(TypeOf x Is Worksheet)"
        );
        assert_eq!(shape(&expr("AddressOf Foo")), "AddressOf Foo");
    }

    // ---- statements -----------------------------------------------------

    #[test]
    fn a_call_can_have_bare_arguments() {
        let m = parse("Sub S()\n    MsgBox \"hi\", vbOK\nEnd Sub\n");
        match &m.procedures()[0].body[0] {
            Stmt::Call { expr, .. } => assert_eq!(shape(expr), "MsgBox(\"hi\", vbOK)"),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn set_is_distinguished_from_let() {
        let m = parse("Sub S()\n    Set a = b\n    Let c = d\n    e = f\nEnd Sub\n");
        let body = &m.procedures()[0].body;
        assert!(matches!(body[0], Stmt::Assign { set: true, .. }));
        assert!(matches!(body[1], Stmt::Assign { set: false, .. }));
        assert!(matches!(body[2], Stmt::Assign { set: false, .. }));
    }

    #[test]
    fn both_if_forms_parse() {
        let m = parse(
            "Sub S()\n\
             If a Then b = 1\n\
             If a Then b = 1 Else b = 2\n\
             If a Then\n    b = 1\nElseIf c Then\n    b = 2\nElse\n    b = 3\nEnd If\n\
             End Sub\n",
        );
        let body = &m.procedures()[0].body;
        assert_eq!(body.len(), 3);
        match &body[2] {
            Stmt::If {
                branches,
                else_body,
                ..
            } => {
                assert_eq!(branches.len(), 2);
                assert!(else_body.is_some());
            }
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn select_case_covers_values_ranges_and_is() {
        let m = parse(
            "Sub S()\n\
             Select Case x\n\
             Case 1, 2\n    a = 1\n\
             Case 3 To 5\n    a = 2\n\
             Case Is >= 6\n    a = 3\n\
             Case Else\n    a = 4\n\
             End Select\n\
             End Sub\n",
        );
        match &m.procedures()[0].body[0] {
            Stmt::SelectCase {
                cases, case_else, ..
            } => {
                assert_eq!(cases.len(), 3);
                assert_eq!(cases[0].matches.len(), 2);
                assert!(matches!(cases[1].matches[0], CaseMatch::Range(_, _)));
                assert!(matches!(cases[2].matches[0], CaseMatch::Is(BinOp::Ge, _)));
                assert!(case_else.is_some());
            }
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn every_loop_form_parses() {
        parse(
            "Sub S()\n\
             For i = 1 To 10 Step 2\n    a = i\nNext i\n\
             For Each c In rng\n    a = c\nNext\n\
             Do While x\n    a = 1\nLoop\n\
             Do Until x\n    a = 1\nLoop\n\
             Do\n    a = 1\nLoop While x\n\
             Do\n    a = 1\nLoop Until x\n\
             While x\n    a = 1\nWend\n\
             End Sub\n",
        );
    }

    #[test]
    fn one_next_can_close_several_for_loops() {
        // `Next j, i` closes both loops; a parser that consumed it once would
        // then look for a second `Next` that isn't there.
        let m = parse(
            "Sub S()\n\
             For i = 1 To 2\n\
             For j = 1 To 2\n\
             a = 1\n\
             Next j, i\n\
             End Sub\n",
        );
        assert_eq!(m.procedures()[0].body.len(), 1);
    }

    #[test]
    fn while_is_normalised_into_a_do_loop() {
        let m = parse("Sub S()\n    While x\n        a = 1\n    Wend\nEnd Sub\n");
        match &m.procedures()[0].body[0] {
            Stmt::DoLoop { pre, post, .. } => {
                assert!(matches!(pre, Some((DoTest::While, _))));
                assert!(post.is_none());
            }
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn with_blocks_allow_leading_dot_references() {
        let m = parse(
            "Sub S()\n    With ws\n        .Range(\"A1\").Value = 1\n    End With\nEnd Sub\n",
        );
        match &m.procedures()[0].body[0] {
            Stmt::With { body, .. } => match &body[0] {
                Stmt::Assign { target, .. } => assert_eq!(shape(target), ".Range(\"A1\").Value"),
                other => panic!("{other:?}"),
            },
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn error_handling_statements_parse() {
        let m = parse(
            "Sub S()\n\
             On Error GoTo Failed\n\
             On Error Resume Next\n\
             On Error GoTo 0\n\
             Resume Next\n\
             Exit Sub\n\
             Failed:\n\
             a = 1\n\
             End Sub\n",
        );
        let body = &m.procedures()[0].body;
        assert!(matches!(
            body[0],
            Stmt::OnError {
                kind: OnErrorKind::GoTo(_),
                ..
            }
        ));
        assert!(matches!(
            body[1],
            Stmt::OnError {
                kind: OnErrorKind::ResumeNext,
                ..
            }
        ));
        // `GoTo 0` disables the handler; it is not a jump to a label "0".
        assert!(matches!(
            body[2],
            Stmt::OnError {
                kind: OnErrorKind::Disable,
                ..
            }
        ));
        assert!(matches!(body[5], Stmt::Label { .. }));
    }

    #[test]
    fn a_label_and_a_named_argument_are_not_confused() {
        // Both are an identifier followed by a colon; `:=` lexing as one
        // token is what separates them.
        let m = parse("Sub S()\n    Foo bar:=1\nBaz:\n    a = 1\nEnd Sub\n");
        let body = &m.procedures()[0].body;
        assert!(matches!(body[0], Stmt::Call { .. }));
        assert!(matches!(&body[1], Stmt::Label { name, .. } if name == "Baz"));
    }

    #[test]
    fn line_numbers_are_labels() {
        let m = parse("Sub S()\n10 a = 1\n20 GoTo 10\nEnd Sub\n");
        let body = &m.procedures()[0].body;
        assert!(matches!(&body[0], Stmt::Label { name, .. } if name == "10"));
        assert!(matches!(&body[3], Stmt::GoTo { label, .. } if label == "10"));
    }

    #[test]
    fn statements_can_share_a_line_via_colons() {
        let m = parse("Sub S()\n    a = 1: b = 2: c = 3\nEnd Sub\n");
        assert_eq!(m.procedures()[0].body.len(), 3);
    }

    // ---- declarations ---------------------------------------------------

    #[test]
    fn declarations_cover_the_shapes_real_modules_use() {
        let m = parse(
            "Option Explicit\n\
             Private Const MAX As Long = 10\n\
             Public x As String, y(1 To 5) As Double, z()\n\
             Dim WithEvents app As Application\n\
             Private Type Point\n    X As Long\n    Y As Long\nEnd Type\n\
             Public Enum Color\n    Red = 1\n    Green\nEnd Enum\n\
             Sub S()\n\
                 Dim a As New Collection\n\
                 Dim buf As String * 10\n\
                 ReDim Preserve arr(1 To 5)\n\
                 Static counter As Long\n\
                 Erase arr\n\
             End Sub\n",
        );
        assert_eq!(m.procedures().len(), 1);
        assert_eq!(m.items.len(), 7);
    }

    #[test]
    fn every_procedure_form_parses() {
        let m = parse(
            "Public Sub A()\nEnd Sub\n\
             Private Function B(ByVal x As Long, Optional y As String = \"d\") As Long\nEnd Function\n\
             Friend Property Get C() As Long\nEnd Property\n\
             Property Let C(v As Long)\nEnd Property\n\
             Property Set D(v As Object)\nEnd Property\n\
             Public Static Sub E(ParamArray args() As Variant)\nEnd Sub\n",
        );
        let procs = m.procedures();
        assert_eq!(procs.len(), 6);
        assert_eq!(procs[0].kind, ProcKind::Sub);
        assert_eq!(procs[1].kind, ProcKind::Function);
        assert_eq!(procs[1].params[0].by, Some(PassBy::Value));
        assert!(procs[1].params[1].optional);
        assert!(procs[1].params[1].default.is_some());
        assert_eq!(procs[2].kind, ProcKind::PropertyGet);
        assert_eq!(procs[3].kind, ProcKind::PropertyLet);
        assert_eq!(procs[4].kind, ProcKind::PropertySet);
        assert!(procs[5].is_static);
        assert!(procs[5].params[0].param_array);
    }

    #[test]
    fn attribute_lines_are_syntax_not_metadata() {
        // Every module stream real Excel writes starts with one.
        let m = parse("Attribute VB_Name = \"Module1\"\nSub S()\nEnd Sub\n");
        assert!(matches!(&m.items[0], ModuleItem::Attribute { name, .. } if name == "VB_Name"));
    }

    #[test]
    fn a_win32_declare_parses_even_though_calling_one_is_out_of_scope() {
        parse(
            "Private Declare PtrSafe Function Sleep Lib \"kernel32\" \
             Alias \"Sleep\" (ByVal ms As Long) As Long\n",
        );
    }

    #[test]
    fn conditional_compilation_parses_both_branches() {
        // Excel reports a syntax error inside an inactive branch, so this
        // must not skip them.
        let m = parse(
            "#Const DEBUGGING = 1\n\
             #If VBA7 Then\n\
             Sub A()\nEnd Sub\n\
             #ElseIf Mac Then\n\
             Sub B()\nEnd Sub\n\
             #Else\n\
             Sub C()\nEnd Sub\n\
             #End If\n",
        );
        assert_eq!(m.procedures().len(), 3);
    }

    #[test]
    fn line_continuations_join_a_split_signature() {
        let m = parse(
            "Public Function F( _\n    ByVal a As Long, _\n    ByVal b As Long) As Long\n\
             F = a + _\n    b\n\
             End Function\n",
        );
        assert_eq!(m.procedures()[0].params.len(), 2);
    }

    // ---- opaque statements ----------------------------------------------

    #[test]
    fn file_io_statements_are_recorded_rather_than_modelled() {
        let m = parse(
            "Sub S()\n\
             Open \"f.txt\" For Input As #1\n\
             Line Input #1, buf\n\
             Print #1, \"x\"\n\
             Close #1\n\
             Name \"a\" As \"b\"\n\
             End Sub\n",
        );
        let body = &m.procedures()[0].body;
        assert_eq!(body.len(), 5);
        assert!(body.iter().all(|s| matches!(s, Stmt::Opaque { .. })));
    }

    #[test]
    fn the_opaque_keywords_stay_usable_as_ordinary_names() {
        // The reason each opaque guard is narrow: every one of these words is
        // also a property or a variable somewhere in real code.
        let m = parse(
            "Sub S()\n\
             Application.Width = 100\n\
             ws.Name = \"Sheet1\"\n\
             Debug.Print x\n\
             Get = 1\n\
             a = rs.Line\n\
             End Sub\n",
        );
        let body = &m.procedures()[0].body;
        assert_eq!(body.len(), 5);
        assert!(!body.iter().any(|s| matches!(s, Stmt::Opaque { .. })));
    }

    // ---- Print output lists ---------------------------------------------

    /// Shapes the single statement of a one-line `Sub`.
    fn stmt_shape(src: &str) -> String {
        let m = parse(&format!(
            "Sub S()
    {src}
End Sub
"
        ));
        match &m.procedures()[0].body[0] {
            Stmt::Call { expr, .. } => shape(expr),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn print_output_lists_accept_every_separator_excel_does() {
        // `;` is a `Print` output separator. Every case here is one
        // `fuzz/vba_compile_probe.py` measurement that real Excel compiled;
        // `_` is an empty output item.
        for (src, want) in [
            (r#"Debug.Print "a"; 1"#, r#"Debug.Print("a", 1)"#),
            (r#"Debug.Print "a"; "b"; 1"#, r#"Debug.Print("a", "b", 1)"#),
            (r#"Debug.Print "a", 1"#, r#"Debug.Print("a", 1)"#),
            // A trailing separator suppresses the newline, so it is legal
            // and load-bearing -- in either spelling.
            (r#"Debug.Print "a";"#, r#"Debug.Print("a")"#),
            (r#"Debug.Print "a","#, r#"Debug.Print("a")"#),
            // A leading or repeated separator prints an empty item.
            (r#"Debug.Print , "a""#, r#"Debug.Print(_, "a")"#),
            (r#"Debug.Print ; "a""#, r#"Debug.Print(_, "a")"#),
            (r#"Debug.Print "a";; "b""#, r#"Debug.Print("a", _, "b")"#),
            // No separator at all between two items also compiles.
            (r#"Debug.Print "a" "b""#, r#"Debug.Print("a", "b")"#),
            (
                r#"Debug.Print Spc(3); "a"; Tab(10); "b""#,
                r#"Debug.Print(Spc(3), "a", Tab(10), "b")"#,
            ),
            // The gate is the member name, not the `Debug` object.
            (r#"x.Print "a"; 1"#, r#"x.Print("a", 1)"#),
        ] {
            assert_eq!(stmt_shape(src), want, "{src}");
        }
    }

    #[test]
    fn a_trailing_print_separator_can_be_followed_by_else() {
        // `If True Then Debug.Print "a"; Else Debug.Print "b"` compiles in
        // Excel, so the output list has to stop at `Else` the way the
        // bare-argument list it replaced did.
        let m = parse(
            "Sub S()
If True Then Debug.Print \"a\"; Else Debug.Print \"b\"
End Sub
",
        );
        assert!(matches!(m.procedures()[0].body[0], Stmt::If { .. }));
    }

    #[test]
    fn a_semicolon_separator_is_only_a_print_thing() {
        // Measured: Excel rejects all three. `;` is not a general
        // bare-argument separator, and unqualified `Print` is a statement
        // only before a `#` -- which `try_parse_opaque` takes first.
        for src in [
            "Sub S()
MsgBox \"a\"; 1
End Sub
",
            "Sub S()
Debug.Assert \"a\"; 1
End Sub
",
            "Sub S()
Print \"a\"; 1
End Sub
",
        ] {
            assert!(parse_module(src).is_err(), "{src:?} should not parse");
        }
    }

    #[test]
    fn a_user_defined_print_keeps_its_named_arguments() {
        // A class module may define its own `Print`, and a named argument to
        // it parsed before the output-list path existed.
        assert_eq!(
            stmt_shape("obj.Print value:=1, style:=2"),
            "obj.Print(value:=1, style:=2)"
        );
    }

    #[test]
    fn fuzz_reserved_type_names_cannot_be_declared() {
        // Harvested from fuzz/fuzz_vba_parse.py: `Dim Long As x` compiled
        // under check_syntax but real Excel refuses it. Measured directly
        // (win32com, real Windows Excel) which half of that is the real,
        // fixable gap: `Dim Long As Integer`, `Const Long = 5`, and a
        // parameter named `Long` all fail to compile in Excel too -- `Long`
        // (like the rest of `RESERVED_TYPE_NAMES`) can never be a declared
        // name, unlike most VBA keywords (contextual and reusable
        // elsewhere, see `the_opaque_keywords_stay_usable_as_ordinary_names`
        // above). `Dim Long As x` itself is left accepting: whether `x` is a
        // valid type needs name resolution Phase 0 doesn't do, so that half
        // of the original case is correctly out of scope, not fixed here.
        for src in [
            "Sub S()\nDim Long As Integer\nEnd Sub\n",
            "Sub S()\nConst Long = 5\nEnd Sub\n",
            "Sub S(Long As Integer)\nEnd Sub\n",
            "Sub S()\nDim Boolean\nEnd Sub\n",
        ] {
            let err = parse_module(src).unwrap_err();
            assert!(err.message.contains("built-in type name"), "{src:?}: {err}");
        }
        // `Object` is not in the reserved set -- it compiles as a name even
        // though it's also a valid type in an `As` clause (measured: `Dim
        // Object As Long` compiles in real Excel).
        parse_module("Sub S()\nDim Object As Long\nEnd Sub\n").unwrap();
        // The unresolvable-type half of the original case stays accepted,
        // deliberately -- Phase 0 does no name resolution.
        parse_module("Sub S()\nDim y As x\nEnd Sub\n").unwrap();
    }

    // ---- errors ---------------------------------------------------------

    #[test]
    fn errors_point_at_the_offending_line() {
        let err = parse_module("Sub S()\n    a = = 1\nEnd Sub\n").unwrap_err();
        assert_eq!(err.pos.line, 2);
        assert!(err.message.contains("expected an expression"), "{err}");
    }

    /// An unclosed block is blamed on the line that opened it, matching how
    /// VBA words and places its own compile errors -- not on whatever token
    /// arrived where the closer was due, which is usually a correct line.
    #[test]
    fn an_unclosed_block_is_blamed_on_its_opener() {
        // (source, expected message, expected line)
        let cases = [
            ("Sub S()\n    a = 1\n", "Expected End Sub", 1),
            ("Function F()\n    a = 1\n", "Expected End Function", 1),
            (
                "Sub S()\n    If a Then\n        b = 1\nEnd Sub\n",
                "Block If without End If",
                2,
            ),
            (
                "Sub S()\n    b = 1\n    If a Then\n        b = 2\nEnd Sub\n",
                "Block If without End If",
                3,
            ),
            (
                "Sub S()\n    For i = 1 To 2\n        a = 1\nEnd Sub\n",
                "For without Next",
                2,
            ),
            (
                "Sub S()\n    For Each c In r\n        a = 1\nEnd Sub\n",
                "For without Next",
                2,
            ),
            (
                "Sub S()\n    Do While a\n        b = 1\nEnd Sub\n",
                "Do without Loop",
                2,
            ),
            (
                "Sub S()\n    While a\n        b = 1\nEnd Sub\n",
                "While without Wend",
                2,
            ),
            (
                "Sub S()\n    With x\n        .a = 1\nEnd Sub\n",
                "With without End With",
                2,
            ),
            (
                "Sub S()\n    Select Case x\n    Case 1\n        a = 1\nEnd Sub\n",
                "Select Case without End Select",
                2,
            ),
            ("Type P\n    X As Long\n", "Expected End Type", 1),
            ("Enum C\n    Red\n", "Expected End Enum", 1),
        ];
        for (src, message, line) in cases {
            let err = parse_module(src).expect_err(src);
            assert_eq!(err.message, message, "for {src:?}");
            assert_eq!(err.pos.line, line, "for {src:?}");
        }
    }

    #[test]
    fn a_mismatched_block_closer_is_reported() {
        // `End Function` does not close a `Sub`, so the Sub is unclosed and
        // is what gets blamed -- again matching VBA's "Expected End Sub".
        let err = parse_module("Sub S()\nEnd Function\n").unwrap_err();
        assert_eq!(err.message, "Expected End Sub");
        assert_eq!(err.pos.line, 1);

        let err = parse_module("Function F()\nEnd Sub\n").unwrap_err();
        assert_eq!(err.message, "Expected End Function");
        assert_eq!(err.pos.line, 1);
    }

    #[test]
    fn the_innermost_unclosed_block_is_the_one_blamed() {
        // Both the For and the If are unclosed; the If is nearer, and fixing
        // it is what lets the next error surface.
        let err = parse_module(
            "Sub S()\n    For i = 1 To 2\n        If a Then\n            b = 1\nEnd Sub\n",
        )
        .unwrap_err();
        assert_eq!(err.message, "Block If without End If");
        assert_eq!(err.pos.line, 3);
    }

    // ---- real-world corpus ----------------------------------------------

    #[test]
    fn parses_the_repos_own_pivot_fuzzing_macro() {
        // ~300 lines of real, Excel-authored VBA: the best available check
        // that this grammar covers what people actually write.
        let src = include_str!("../../../../fuzz/BuildFuzzPivot.bas");
        let m = parse(src);
        assert!(
            m.procedures().len() >= 2,
            "expected several procedures, got {}",
            m.procedures().len()
        );
    }

    #[test]
    fn never_panics_on_arbitrary_input() {
        for src in [
            "",
            "\n",
            "Sub",
            "Sub S(",
            "End Sub",
            "If",
            "For",
            "Do",
            "Loop",
            "Next",
            "Case",
            "#If",
            "Attribute",
            "Option",
            "Dim",
            "a = ",
            "((((((",
            "))))))",
            "Set",
            "On Error",
            "Type",
            "Enum",
            "Declare",
            "Property",
            "Exit",
            "Resume",
            "1",
            ":",
            ".",
            "&",
        ] {
            let _ = parse_module(src);
        }
    }
}
