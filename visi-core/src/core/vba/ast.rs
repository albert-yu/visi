//! The VBA syntax tree.
//!
//! Shaped for a tree-walking interpreter to come (see the phased plan in
//! `docs/vba-macro-support.md`), so nodes carry the source positions an error
//! at runtime will need, and constructs are kept apart where evaluation will
//! treat them differently -- `For` and `For Each` are distinct variants
//! rather than one loop with an optional collection, because almost nothing
//! about executing them is shared.
//!
//! Where VBA offers several spellings of one thing, this normalises: `While`
//! becomes a [`Stmt::DoLoop`], `Let x = 1` and `x = 1` are both
//! [`Stmt::Assign`], and `=<` has already become `<=` in the lexer. Where the
//! difference is semantic it is preserved: `Set x = y` keeps its `set` flag,
//! since VBA assigns a reference rather than a value.

use super::lexer::{NumBase, Pos, TypeSuffix};

/// A whole module: the unit `visi macro check` validates.
#[derive(Debug, Clone, PartialEq)]
pub struct Module {
    /// Top-level items, in source order. Order is preserved rather than
    /// bucketed by kind so a future formatter or round trip can reproduce the
    /// file, and so an error can point at the right one.
    pub items: Vec<ModuleItem>,
}

/// Anything that can appear at module level.
#[derive(Debug, Clone, PartialEq)]
pub enum ModuleItem {
    /// `Attribute VB_Name = "Module1"`. Real Excel writes these into every
    /// module stream, so they are syntax to accept, not metadata to strip.
    Attribute {
        /// The attribute name, e.g. `VB_Name`.
        name: String,
        /// Its values -- `Attribute VB_Ext_KEY = "a", "b"` takes several.
        values: Vec<Expr>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Option Explicit`, `Option Base 1`, ...
    Option {
        /// The words after `Option`, preserved verbatim.
        words: Vec<String>,
        /// Where it starts.
        pos: Pos,
    },
    /// A module-level declaration: `Dim`, `Const`, `Type`, `Enum`, `Declare`,
    /// `Event`, `Implements`.
    Declaration(Stmt),
    /// A `Sub`, `Function` or `Property`.
    Procedure(Procedure),
    /// A `#If` / `#Const` conditional-compilation block. Its branches are
    /// parsed as ordinary items rather than skipped, so a syntax error inside
    /// an inactive branch is still reported -- which is what Excel does.
    Conditional {
        /// One entry per `#If` / `#ElseIf` branch.
        branches: Vec<(Expr, Vec<ModuleItem>)>,
        /// The `#Else` branch, if present.
        else_items: Option<Vec<ModuleItem>>,
        /// Where it starts.
        pos: Pos,
    },
}

/// What kind of procedure something is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind {
    /// `Sub` -- no return value.
    Sub,
    /// `Function` -- returns a value by assigning to its own name.
    Function,
    /// `Property Get`.
    PropertyGet,
    /// `Property Let`.
    PropertyLet,
    /// `Property Set`.
    PropertySet,
}

/// A declaration's visibility.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    /// `Public`.
    Public,
    /// `Private`.
    Private,
    /// `Friend` -- class-module visibility with no direct Rust analogue.
    Friend,
    /// `Global`, the pre-VB6 spelling of `Public`.
    Global,
}

/// A `Sub`, `Function` or `Property`.
#[derive(Debug, Clone, PartialEq)]
pub struct Procedure {
    /// Which of the five forms this is.
    pub kind: ProcKind,
    /// The procedure name.
    pub name: String,
    /// `Public` / `Private` / `Friend`, if written.
    pub visibility: Option<Visibility>,
    /// Whether it was declared `Static`.
    pub is_static: bool,
    /// Its parameters, in order.
    pub params: Vec<Param>,
    /// A `Function`'s or `Property Get`'s declared return type.
    pub return_type: Option<TypeRef>,
    /// The body, as statements.
    pub body: Vec<Stmt>,
    /// Where the declaration starts.
    pub pos: Pos,
}

/// How a parameter is passed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PassBy {
    /// `ByVal`.
    Value,
    /// `ByRef`, which is also VBA's default when neither is written.
    Reference,
}

/// One parameter of a [`Procedure`].
#[derive(Debug, Clone, PartialEq)]
pub struct Param {
    /// The parameter name.
    pub name: String,
    /// `ByVal` / `ByRef`, if written. `None` means VBA's default (`ByRef`),
    /// kept distinct from an explicit `ByRef` so a round trip is faithful.
    pub by: Option<PassBy>,
    /// Whether it is `Optional`.
    pub optional: bool,
    /// Whether it is a `ParamArray`.
    pub param_array: bool,
    /// Whether it was declared as an array (`x() As Long`).
    pub is_array: bool,
    /// Its declared type, if any.
    pub ty: Option<TypeRef>,
    /// An `Optional` parameter's default value.
    pub default: Option<Expr>,
}

/// A type name in an `As` clause: `Long`, `Excel.Range`, `Collection`.
#[derive(Debug, Clone, PartialEq)]
pub struct TypeRef {
    /// The dotted name parts, e.g. `["Excel", "Range"]`.
    pub path: Vec<String>,
    /// `New` in `Dim x As New Collection`.
    pub is_new: bool,
    /// A fixed-length string's size: `As String * 10`. Boxed because a
    /// `TypeRef` is reachable from [`Expr::New`], so the two types are
    /// mutually recursive.
    pub string_length: Option<Box<Expr>>,
}

/// One name in a `Dim` / `Const` / `ReDim` list.
#[derive(Debug, Clone, PartialEq)]
pub struct VarDecl {
    /// The variable name.
    pub name: String,
    /// Array bounds, if declared as an array. An empty vector means `x()`,
    /// a dynamic array, which is distinct from not being an array at all.
    pub bounds: Option<Vec<ArrayBound>>,
    /// Its declared type.
    pub ty: Option<TypeRef>,
    /// A `Const`'s value.
    pub value: Option<Expr>,
    /// Where the name is.
    pub pos: Pos,
}

/// One dimension of an array declaration.
#[derive(Debug, Clone, PartialEq)]
pub struct ArrayBound {
    /// The lower bound in `1 To 10`, absent in a bare `10`.
    pub lower: Option<Expr>,
    /// The upper bound.
    pub upper: Expr,
}

/// Which keyword introduced a variable declaration, since it decides scope
/// and lifetime.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DimKind {
    /// `Dim`.
    Dim,
    /// `Static` -- persists across calls.
    Static,
    /// `Private` at module level.
    Private,
    /// `Public` at module level.
    Public,
    /// `Global`, the older spelling of `Public`.
    Global,
}

/// What `Exit` leaves.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind {
    /// `Exit Sub`.
    Sub,
    /// `Exit Function`.
    Function,
    /// `Exit Property`.
    Property,
    /// `Exit For`.
    For,
    /// `Exit Do`.
    Do,
    /// `Exit While`.
    While,
}

/// The three forms of `On Error`.
#[derive(Debug, Clone, PartialEq)]
pub enum OnErrorKind {
    /// `On Error GoTo <label>`.
    GoTo(String),
    /// `On Error Resume Next`.
    ResumeNext,
    /// `On Error GoTo 0` -- disables the active handler.
    Disable,
}

/// The three forms of `Resume`.
#[derive(Debug, Clone, PartialEq)]
pub enum ResumeKind {
    /// Bare `Resume` -- retries the failing statement.
    Retry,
    /// `Resume Next`.
    Next,
    /// `Resume <label>`.
    Label(String),
}

/// Whether a `Do` loop tests before or after the body, and on which sense.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DoTest {
    /// `While <cond>` -- run while true.
    While,
    /// `Until <cond>` -- run while false.
    Until,
}

/// One `Case` of a `Select Case`.
#[derive(Debug, Clone, PartialEq)]
pub struct CaseClause {
    /// The values this case matches.
    pub matches: Vec<CaseMatch>,
    /// Its body.
    pub body: Vec<Stmt>,
}

/// One alternative within a `Case`.
#[derive(Debug, Clone, PartialEq)]
pub enum CaseMatch {
    /// A plain value: `Case 1`.
    Value(Expr),
    /// `Case 1 To 5`.
    Range(Expr, Expr),
    /// `Case Is >= 5`, holding the operator and its right side.
    Is(BinOp, Expr),
}

/// A member of an `Enum`.
#[derive(Debug, Clone, PartialEq)]
pub struct EnumMember {
    /// The member name.
    pub name: String,
    /// Its explicit value, if written.
    pub value: Option<Expr>,
}

/// A statement.
#[derive(Debug, Clone, PartialEq)]
pub enum Stmt {
    /// `Dim` / `Static` / `Private` / `Public` / `Global` variable declarations.
    Dim {
        /// Which keyword introduced it.
        kind: DimKind,
        /// `Dim WithEvents x As Foo`.
        with_events: bool,
        /// The declared names.
        vars: Vec<VarDecl>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Const NAME = value`.
    Const {
        /// Its visibility, if written.
        visibility: Option<Visibility>,
        /// The declared constants.
        vars: Vec<VarDecl>,
        /// Where it starts.
        pos: Pos,
    },
    /// `ReDim [Preserve] x(bounds)`.
    ReDim {
        /// Whether `Preserve` was written.
        preserve: bool,
        /// The arrays being resized.
        vars: Vec<VarDecl>,
        /// Where it starts.
        pos: Pos,
    },
    /// An assignment. `Let` is normalised away; `Set` is not, because it
    /// binds a reference rather than a value.
    Assign {
        /// The left-hand side.
        target: Expr,
        /// The right-hand side.
        value: Expr,
        /// Whether this was a `Set`.
        set: bool,
        /// Where it starts.
        pos: Pos,
    },
    /// A procedure call used as a statement, with or without `Call`.
    Call {
        /// The call expression.
        expr: Expr,
        /// Where it starts.
        pos: Pos,
    },
    /// `If` / `ElseIf` / `Else`, both the block and single-line forms.
    If {
        /// One entry per `If` / `ElseIf` branch, condition first.
        branches: Vec<(Expr, Vec<Stmt>)>,
        /// The `Else` body, if present.
        else_body: Option<Vec<Stmt>>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Select Case`.
    SelectCase {
        /// The value being switched on.
        subject: Expr,
        /// The `Case` clauses, in order.
        cases: Vec<CaseClause>,
        /// The `Case Else` body, if present.
        case_else: Option<Vec<Stmt>>,
        /// Where it starts.
        pos: Pos,
    },
    /// `For i = a To b [Step c]`.
    For {
        /// The loop variable, as an expression so `For obj.i` parses.
        var: Expr,
        /// The starting value.
        from: Expr,
        /// The limit.
        to: Expr,
        /// The step, if written.
        step: Option<Expr>,
        /// The body.
        body: Vec<Stmt>,
        /// Where it starts.
        pos: Pos,
    },
    /// `For Each x In collection`.
    ForEach {
        /// The element variable.
        var: Expr,
        /// The collection.
        iterable: Expr,
        /// The body.
        body: Vec<Stmt>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Do`/`Loop` in all four forms, plus `While`/`Wend` normalised into it.
    DoLoop {
        /// A test at the top, if any.
        pre: Option<(DoTest, Expr)>,
        /// A test at the bottom, if any.
        post: Option<(DoTest, Expr)>,
        /// The body.
        body: Vec<Stmt>,
        /// Where it starts.
        pos: Pos,
    },
    /// `With obj ... End With`.
    With {
        /// The object the leading-dot references resolve against.
        subject: Expr,
        /// The body.
        body: Vec<Stmt>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Exit Sub` and friends.
    Exit {
        /// What is being exited.
        kind: ExitKind,
        /// Where it starts.
        pos: Pos,
    },
    /// `GoTo label`.
    GoTo {
        /// The target label.
        label: String,
        /// Where it starts.
        pos: Pos,
    },
    /// `GoSub label`.
    GoSub {
        /// The target label.
        label: String,
        /// Where it starts.
        pos: Pos,
    },
    /// `Return`, which in VBA returns from a `GoSub`.
    Return {
        /// Where it starts.
        pos: Pos,
    },
    /// `On Error ...`.
    OnError {
        /// Which form.
        kind: OnErrorKind,
        /// Where it starts.
        pos: Pos,
    },
    /// `On <expr> GoTo l1, l2` / `On <expr> GoSub l1, l2`.
    OnGoto {
        /// The selector.
        subject: Expr,
        /// The candidate labels.
        labels: Vec<String>,
        /// Whether this was `GoSub` rather than `GoTo`.
        gosub: bool,
        /// Where it starts.
        pos: Pos,
    },
    /// `Resume ...`.
    Resume {
        /// Which form.
        kind: ResumeKind,
        /// Where it starts.
        pos: Pos,
    },
    /// A line label (`Failed:`) or a line number, both of which name a jump
    /// target.
    Label {
        /// The label text; a line number arrives as its digits.
        name: String,
        /// Where it starts.
        pos: Pos,
    },
    /// `Erase a, b`.
    Erase {
        /// The arrays being cleared.
        targets: Vec<Expr>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Type Foo ... End Type`.
    TypeDef {
        /// The type name.
        name: String,
        /// Its visibility, if written.
        visibility: Option<Visibility>,
        /// Its fields.
        fields: Vec<VarDecl>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Enum Foo ... End Enum`.
    EnumDef {
        /// The enum name.
        name: String,
        /// Its visibility, if written.
        visibility: Option<Visibility>,
        /// Its members.
        members: Vec<EnumMember>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Declare [PtrSafe] Sub/Function ... Lib "..."` -- a Win32 API import.
    /// Parsed so real modules do not fail to check; calling one is explicitly
    /// out of scope for the interpreter.
    Declare {
        /// The aliased procedure name.
        name: String,
        /// Whether it returns a value.
        is_function: bool,
        /// The library name.
        lib: String,
        /// An `Alias` clause, if written.
        alias: Option<String>,
        /// Its parameters.
        params: Vec<Param>,
        /// Its return type.
        return_type: Option<TypeRef>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Event Foo(args)` in a class module.
    EventDef {
        /// The event name.
        name: String,
        /// Its parameters.
        params: Vec<Param>,
        /// Where it starts.
        pos: Pos,
    },
    /// `RaiseEvent Foo(args)`.
    RaiseEvent {
        /// The event name.
        name: String,
        /// The arguments.
        args: Vec<Arg>,
        /// Where it starts.
        pos: Pos,
    },
    /// `Implements IFoo`.
    Implements {
        /// The interface name.
        name: String,
        /// Where it starts.
        pos: Pos,
    },
    /// `Stop` -- breaks into the debugger.
    Stop {
        /// Where it starts.
        pos: Pos,
    },
    /// `End` as a statement -- halts execution. Distinct from the `End X`
    /// that closes a block, which never reaches the AST.
    End {
        /// Where it starts.
        pos: Pos,
    },
    /// A statement this parser accepts but models only as its source text.
    ///
    /// VBA's file I/O and legacy graphics statements (`Open`, `Print #`,
    /// `Line Input`, `Name x As y`) have irregular syntax that would triple
    /// the grammar for constructs the interpreter is never going to run --
    /// they are out of scope by the plan's own security posture. Keeping them
    /// as opaque text means a real-world module still passes `macro check`
    /// instead of failing on a line nobody will execute.
    Opaque {
        /// The leading keyword, for a later "unsupported" diagnostic.
        keyword: String,
        /// Where it starts.
        pos: Pos,
    },
    /// `Attribute Item.VB_UserMemId = 0` or other statement-level attribute.
    Attribute {
        /// The attribute name.
        name: String,
        /// Its values.
        values: Vec<Expr>,
        /// Where it starts.
        pos: Pos,
    },
}

/// A unary operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOp {
    /// Arithmetic negation.
    Neg,
    /// Unary plus, which VBA accepts and which is not a no-op on a String.
    Pos,
    /// Logical/bitwise `Not`.
    Not,
}

/// A binary operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    /// `^`.
    Pow,
    /// `*`.
    Mul,
    /// `/` -- floating-point division.
    Div,
    /// `\` -- integer division.
    IntDiv,
    /// `Mod`.
    Mod,
    /// `+`.
    Add,
    /// `-`.
    Sub,
    /// `&` -- string concatenation.
    Concat,
    /// `=` used as a comparison.
    Eq,
    /// `<>`.
    Ne,
    /// `<`.
    Lt,
    /// `>`.
    Gt,
    /// `<=`.
    Le,
    /// `>=`.
    Ge,
    /// `Is` -- reference identity.
    Is,
    /// `Like` -- pattern match.
    Like,
    /// `And`.
    And,
    /// `Or`.
    Or,
    /// `Xor`.
    Xor,
    /// `Eqv` -- logical equivalence.
    Eqv,
    /// `Imp` -- logical implication.
    Imp,
}

/// A literal value.
#[derive(Debug, Clone, PartialEq)]
pub enum Literal {
    /// A number, with how it was written.
    Number {
        /// Its value.
        value: f64,
        /// Decimal, hex or octal.
        base: NumBase,
        /// A trailing type-declaration character.
        suffix: Option<TypeSuffix>,
        /// Whether it was written with a fraction or exponent, which forces
        /// `Double`. See [`super::lexer::TokenKind::Number`].
        is_float: bool,
    },
    /// A string.
    Str(String),
    /// A `#...#` date, as its raw text.
    Date(String),
    /// `True` or `False`.
    Bool(bool),
    /// `Nothing`.
    Nothing,
    /// `Empty`.
    Empty,
    /// `Null`.
    Null,
}

/// One argument at a call site.
#[derive(Debug, Clone, PartialEq)]
pub struct Arg {
    /// A `name:=value` argument's name.
    pub name: Option<String>,
    /// The value, or `None` for an omitted positional argument (`f(1, , 3)`),
    /// which VBA allows and which is not the same as passing `Empty`.
    pub value: Option<Expr>,
}

/// An expression.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    /// A literal.
    Literal(Literal),
    /// A bare name.
    Ident {
        /// The name, in its original spelling.
        name: String,
        /// Where it is.
        pos: Pos,
    },
    /// `Me`.
    Me {
        /// Where it is.
        pos: Pos,
    },
    /// Member access: `a.b`, or `.b` inside a `With` when `target` is `None`.
    Member {
        /// What is being accessed, or `None` for a leading-dot reference.
        target: Option<Box<Expr>>,
        /// The member name.
        name: String,
        /// Where it is.
        pos: Pos,
    },
    /// Dictionary access: `rs!Field`.
    Bang {
        /// The object.
        target: Box<Expr>,
        /// The key.
        name: String,
        /// Where it is.
        pos: Pos,
    },
    /// A call or an array index -- indistinguishable in VBA's syntax, and
    /// deliberately not distinguished here, since telling them apart needs
    /// the symbol table that Phase 1 will have and Phase 0 does not.
    Call {
        /// What is being called or indexed.
        target: Box<Expr>,
        /// The arguments.
        args: Vec<Arg>,
        /// Where it is.
        pos: Pos,
    },
    /// A unary operation.
    Unary {
        /// Which operator.
        op: UnOp,
        /// Its operand.
        expr: Box<Expr>,
        /// Where it is.
        pos: Pos,
    },
    /// A binary operation.
    Binary {
        /// Which operator.
        op: BinOp,
        /// The left operand.
        lhs: Box<Expr>,
        /// The right operand.
        rhs: Box<Expr>,
        /// Where the operator is.
        pos: Pos,
    },
    /// A parenthesised expression, kept rather than folded away because in
    /// VBA `Foo (x)` forces `x` to be passed by value even when the parameter
    /// is `ByRef` -- the parentheses are semantic, not just grouping.
    Paren {
        /// The inner expression.
        expr: Box<Expr>,
        /// Where it is.
        pos: Pos,
    },
    /// `New Foo`.
    New {
        /// The type being constructed.
        ty: TypeRef,
        /// Where it is.
        pos: Pos,
    },
    /// `TypeOf x Is Foo`.
    TypeOf {
        /// The value being tested.
        expr: Box<Expr>,
        /// The type it is tested against.
        ty: TypeRef,
        /// Where it is.
        pos: Pos,
    },
    /// `AddressOf Foo`.
    AddressOf {
        /// The procedure name.
        name: String,
        /// Where it is.
        pos: Pos,
    },
}

impl Expr {
    /// Where this expression starts.
    pub fn pos(&self) -> Pos {
        match self {
            Expr::Literal(_) => Pos::default(),
            Expr::Ident { pos, .. }
            | Expr::Me { pos }
            | Expr::Member { pos, .. }
            | Expr::Bang { pos, .. }
            | Expr::Call { pos, .. }
            | Expr::Unary { pos, .. }
            | Expr::Binary { pos, .. }
            | Expr::Paren { pos, .. }
            | Expr::New { pos, .. }
            | Expr::TypeOf { pos, .. }
            | Expr::AddressOf { pos, .. } => *pos,
        }
    }
}

impl Module {
    /// Every procedure declared in this module, in source order.
    ///
    /// Convenience for the common "what can `run VB macro` invoke?" question;
    /// conditional-compilation branches are included, since which branch is
    /// live depends on `#Const` values this phase does not evaluate.
    pub fn procedures(&self) -> Vec<&Procedure> {
        fn walk<'a>(items: &'a [ModuleItem], out: &mut Vec<&'a Procedure>) {
            for item in items {
                match item {
                    ModuleItem::Procedure(p) => out.push(p),
                    ModuleItem::Conditional {
                        branches,
                        else_items,
                        ..
                    } => {
                        for (_, body) in branches {
                            walk(body, out);
                        }
                        if let Some(body) = else_items {
                            walk(body, out);
                        }
                    }
                    _ => {}
                }
            }
        }
        let mut out = Vec::new();
        walk(&self.items, &mut out);
        out
    }
}
