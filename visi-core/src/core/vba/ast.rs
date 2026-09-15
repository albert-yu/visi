use super::lexer::{NumBase, Pos, TypeSuffix};

/// [LLM-generated] A whole module: the unit `visi macro check` validates.
#[derive(Debug, Clone, PartialEq)]
pub struct Module {
    /// [LLM-generated] Top-level items, in source order. Order is preserved rather than
    /// bucketed by kind so a future formatter or round trip can reproduce the
    /// file, and so an error can point at the right one.
    pub items: Vec<ModuleItem>,
}

/// [LLM-generated] Anything that can appear at module level.
#[derive(Debug, Clone, PartialEq)]
pub enum ModuleItem {
    /// [LLM-generated] `Attribute VB_Name = "Module1"`. Real Excel writes these into every
    /// module stream, so they are syntax to accept, not metadata to strip.
    Attribute {
        /// [LLM-generated] The attribute name, e.g. `VB_Name`.
        name: String,
        /// [LLM-generated] Its values -- `Attribute VB_Ext_KEY = "a", "b"` takes several.
        values: Vec<Expr>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Option Explicit`, `Option Base 1`, ...
    Option {
        /// [LLM-generated] The words after `Option`, preserved verbatim.
        words: Vec<String>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] A module-level declaration: `Dim`, `Const`, `Type`, `Enum`, `Declare`,
    /// `Event`, `Implements`.
    Declaration(Stmt),
    /// [LLM-generated] A `Sub`, `Function` or `Property`.
    Procedure(Procedure),
    /// [LLM-generated] A `#If` / `#Const` conditional-compilation block. Its branches are
    /// parsed as ordinary items rather than skipped, so a syntax error inside
    /// an inactive branch is still reported -- which is what Excel does.
    Conditional {
        /// [LLM-generated] One entry per `#If` / `#ElseIf` branch.
        branches: Vec<(Expr, Vec<ModuleItem>)>,
        /// [LLM-generated] The `#Else` branch, if present.
        else_items: Option<Vec<ModuleItem>>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
}

/// [LLM-generated] What kind of procedure something is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind {
    /// [LLM-generated] `Sub` -- no return value.
    Sub,
    /// [LLM-generated] `Function` -- returns a value by assigning to its own name.
    Function,
    /// [LLM-generated] `Property Get`.
    PropertyGet,
    /// [LLM-generated] `Property Let`.
    PropertyLet,
    /// [LLM-generated] `Property Set`.
    PropertySet,
}

/// [LLM-generated] A declaration's visibility.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    /// [LLM-generated] `Public`.
    Public,
    /// [LLM-generated] `Private`.
    Private,
    /// [LLM-generated] `Friend` -- class-module visibility with no direct Rust analogue.
    Friend,
    /// [LLM-generated] `Global`, the pre-VB6 spelling of `Public`.
    Global,
}

/// [LLM-generated] A `Sub`, `Function` or `Property`.
#[derive(Debug, Clone, PartialEq)]
pub struct Procedure {
    /// [LLM-generated] Which of the five forms this is.
    pub kind: ProcKind,
    /// [LLM-generated] The procedure name.
    pub name: String,
    /// [LLM-generated] `Public` / `Private` / `Friend`, if written.
    pub visibility: Option<Visibility>,
    /// [LLM-generated] Whether it was declared `Static`.
    pub is_static: bool,
    /// [LLM-generated] Its parameters, in order.
    pub params: Vec<Param>,
    /// [LLM-generated] A `Function`'s or `Property Get`'s declared return type.
    pub return_type: Option<TypeRef>,
    /// [LLM-generated] The body, as statements.
    pub body: Vec<Stmt>,
    /// [LLM-generated] Where the declaration starts.
    pub pos: Pos,
}

/// [LLM-generated] How a parameter is passed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PassBy {
    /// [LLM-generated] `ByVal`.
    Value,
    /// [LLM-generated] `ByRef`, which is also VBA's default when neither is written.
    Reference,
}

/// [LLM-generated] One parameter of a [`Procedure`].
#[derive(Debug, Clone, PartialEq)]
pub struct Param {
    /// [LLM-generated] The parameter name.
    pub name: String,
    /// [LLM-generated] `ByVal` / `ByRef`, if written. `None` means VBA's default (`ByRef`),
    /// kept distinct from an explicit `ByRef` so a round trip is faithful.
    pub by: Option<PassBy>,
    /// [LLM-generated] Whether it is `Optional`.
    pub optional: bool,
    /// [LLM-generated] Whether it is a `ParamArray`.
    pub param_array: bool,
    /// [LLM-generated] Whether it was declared as an array (`x() As Long`).
    pub is_array: bool,
    /// [LLM-generated] Its declared type, if any.
    pub ty: Option<TypeRef>,
    /// [LLM-generated] An `Optional` parameter's default value.
    pub default: Option<Expr>,
}

/// [LLM-generated] A type name in an `As` clause: `Long`, `Excel.Range`, `Collection`.
#[derive(Debug, Clone, PartialEq)]
pub struct TypeRef {
    /// [LLM-generated] The dotted name parts, e.g. `["Excel", "Range"]`.
    pub path: Vec<String>,
    /// [LLM-generated] `New` in `Dim x As New Collection`.
    pub is_new: bool,
    /// [LLM-generated] A fixed-length string's size: `As String * 10`. Boxed because a
    /// `TypeRef` is reachable from [`Expr::New`], so the two types are
    /// mutually recursive.
    pub string_length: Option<Box<Expr>>,
}

/// [LLM-generated] One name in a `Dim` / `Const` / `ReDim` list.
#[derive(Debug, Clone, PartialEq)]
pub struct VarDecl {
    /// [LLM-generated] The variable name.
    pub name: String,
    /// [LLM-generated] Array bounds, if declared as an array. An empty vector means `x()`,
    /// a dynamic array, which is distinct from not being an array at all.
    pub bounds: Option<Vec<ArrayBound>>,
    /// [LLM-generated] Its declared type.
    pub ty: Option<TypeRef>,
    /// [LLM-generated] A `Const`'s value.
    pub value: Option<Expr>,
    /// [LLM-generated] Where the name is.
    pub pos: Pos,
}

/// [LLM-generated] One dimension of an array declaration.
#[derive(Debug, Clone, PartialEq)]
pub struct ArrayBound {
    /// [LLM-generated] The lower bound in `1 To 10`, absent in a bare `10`.
    pub lower: Option<Expr>,
    /// [LLM-generated] The upper bound.
    pub upper: Expr,
}

/// [LLM-generated] Which keyword introduced a variable declaration, since it decides scope
/// and lifetime.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DimKind {
    /// [LLM-generated] `Dim`.
    Dim,
    /// [LLM-generated] `Static` -- persists across calls.
    Static,
    /// [LLM-generated] `Private` at module level.
    Private,
    /// [LLM-generated] `Public` at module level.
    Public,
    /// [LLM-generated] `Global`, the older spelling of `Public`.
    Global,
}

/// [LLM-generated] What `Exit` leaves.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind {
    /// [LLM-generated] `Exit Sub`.
    Sub,
    /// [LLM-generated] `Exit Function`.
    Function,
    /// [LLM-generated] `Exit Property`.
    Property,
    /// [LLM-generated] `Exit For`.
    For,
    /// [LLM-generated] `Exit Do`.
    Do,
    /// [LLM-generated] `Exit While`.
    While,
}

/// [LLM-generated] The three forms of `On Error`.
#[derive(Debug, Clone, PartialEq)]
pub enum OnErrorKind {
    /// [LLM-generated] `On Error GoTo <label>`.
    GoTo(String),
    /// [LLM-generated] `On Error Resume Next`.
    ResumeNext,
    /// [LLM-generated] `On Error GoTo 0` -- disables the active handler.
    Disable,
}

/// [LLM-generated] The three forms of `Resume`.
#[derive(Debug, Clone, PartialEq)]
pub enum ResumeKind {
    /// [LLM-generated] Bare `Resume` -- retries the failing statement.
    Retry,
    /// [LLM-generated] `Resume Next`.
    Next,
    /// [LLM-generated] `Resume <label>`.
    Label(String),
}

/// [LLM-generated] Whether a `Do` loop tests before or after the body, and on which sense.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DoTest {
    /// [LLM-generated] `While <cond>` -- run while true.
    While,
    /// [LLM-generated] `Until <cond>` -- run while false.
    Until,
}

/// [LLM-generated] One `Case` of a `Select Case`.
#[derive(Debug, Clone, PartialEq)]
pub struct CaseClause {
    /// [LLM-generated] The values this case matches.
    pub matches: Vec<CaseMatch>,
    /// [LLM-generated] Its body.
    pub body: Vec<Stmt>,
}

/// [LLM-generated] One alternative within a `Case`.
#[derive(Debug, Clone, PartialEq)]
pub enum CaseMatch {
    /// [LLM-generated] A plain value: `Case 1`.
    Value(Expr),
    /// [LLM-generated] `Case 1 To 5`.
    Range(Expr, Expr),
    /// [LLM-generated] `Case Is >= 5`, holding the operator and its right side.
    Is(BinOp, Expr),
}

/// [LLM-generated] A member of an `Enum`.
#[derive(Debug, Clone, PartialEq)]
pub struct EnumMember {
    /// [LLM-generated] The member name.
    pub name: String,
    /// [LLM-generated] Its explicit value, if written.
    pub value: Option<Expr>,
}

/// [LLM-generated] A statement.
#[derive(Debug, Clone, PartialEq)]
pub enum Stmt {
    /// [LLM-generated] `Dim` / `Static` / `Private` / `Public` / `Global` variable declarations.
    Dim {
        /// [LLM-generated] Which keyword introduced it.
        kind: DimKind,
        /// [LLM-generated] `Dim WithEvents x As Foo`.
        with_events: bool,
        /// [LLM-generated] The declared names.
        vars: Vec<VarDecl>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Const NAME = value`.
    Const {
        /// [LLM-generated] Its visibility, if written.
        visibility: Option<Visibility>,
        /// [LLM-generated] The declared constants.
        vars: Vec<VarDecl>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `ReDim [Preserve] x(bounds)`.
    ReDim {
        /// [LLM-generated] Whether `Preserve` was written.
        preserve: bool,
        /// [LLM-generated] The arrays being resized.
        vars: Vec<VarDecl>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] An assignment. `Let` is normalised away; `Set` is not, because it
    /// binds a reference rather than a value.
    Assign {
        /// [LLM-generated] The left-hand side.
        target: Expr,
        /// [LLM-generated] The right-hand side.
        value: Expr,
        /// [LLM-generated] Whether this was a `Set`.
        set: bool,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] A procedure call used as a statement, with or without `Call`.
    Call {
        /// [LLM-generated] The call expression.
        expr: Expr,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `If` / `ElseIf` / `Else`, both the block and single-line forms.
    If {
        /// [LLM-generated] One entry per `If` / `ElseIf` branch, condition first.
        branches: Vec<(Expr, Vec<Stmt>)>,
        /// [LLM-generated] The `Else` body, if present.
        else_body: Option<Vec<Stmt>>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Select Case`.
    SelectCase {
        /// [LLM-generated] The value being switched on.
        subject: Expr,
        /// [LLM-generated] The `Case` clauses, in order.
        cases: Vec<CaseClause>,
        /// [LLM-generated] The `Case Else` body, if present.
        case_else: Option<Vec<Stmt>>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `For i = a To b [Step c]`.
    For {
        /// [LLM-generated] The loop variable, as an expression so `For obj.i` parses.
        var: Expr,
        /// [LLM-generated] The starting value.
        from: Expr,
        /// [LLM-generated] The limit.
        to: Expr,
        /// [LLM-generated] The step, if written.
        step: Option<Expr>,
        /// [LLM-generated] The body.
        body: Vec<Stmt>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `For Each x In collection`.
    ForEach {
        /// [LLM-generated] The element variable.
        var: Expr,
        /// [LLM-generated] The collection.
        iterable: Expr,
        /// [LLM-generated] The body.
        body: Vec<Stmt>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Do`/`Loop` in all four forms, plus `While`/`Wend` normalised into it.
    DoLoop {
        /// [LLM-generated] A test at the top, if any.
        pre: Option<(DoTest, Expr)>,
        /// [LLM-generated] A test at the bottom, if any.
        post: Option<(DoTest, Expr)>,
        /// [LLM-generated] The body.
        body: Vec<Stmt>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `With obj ... End With`.
    With {
        /// [LLM-generated] The object the leading-dot references resolve against.
        subject: Expr,
        /// [LLM-generated] The body.
        body: Vec<Stmt>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Exit Sub` and friends.
    Exit {
        /// [LLM-generated] What is being exited.
        kind: ExitKind,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `GoTo label`.
    GoTo {
        /// [LLM-generated] The target label.
        label: String,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `GoSub label`.
    GoSub {
        /// [LLM-generated] The target label.
        label: String,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Return`, which in VBA returns from a `GoSub`.
    Return {
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `On Error ...`.
    OnError {
        /// [LLM-generated] Which form.
        kind: OnErrorKind,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `On <expr> GoTo l1, l2` / `On <expr> GoSub l1, l2`.
    OnGoto {
        /// [LLM-generated] The selector.
        subject: Expr,
        /// [LLM-generated] The candidate labels.
        labels: Vec<String>,
        /// [LLM-generated] Whether this was `GoSub` rather than `GoTo`.
        gosub: bool,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Resume ...`.
    Resume {
        /// [LLM-generated] Which form.
        kind: ResumeKind,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] A line label (`Failed:`) or a line number, both of which name a jump
    /// target.
    Label {
        /// [LLM-generated] The label text; a line number arrives as its digits.
        name: String,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Erase a, b`.
    Erase {
        /// [LLM-generated] The arrays being cleared.
        targets: Vec<Expr>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Type Foo ... End Type`.
    TypeDef {
        /// [LLM-generated] The type name.
        name: String,
        /// [LLM-generated] Its visibility, if written.
        visibility: Option<Visibility>,
        /// [LLM-generated] Its fields.
        fields: Vec<VarDecl>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Enum Foo ... End Enum`.
    EnumDef {
        /// [LLM-generated] The enum name.
        name: String,
        /// [LLM-generated] Its visibility, if written.
        visibility: Option<Visibility>,
        /// [LLM-generated] Its members.
        members: Vec<EnumMember>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Declare [PtrSafe] Sub/Function ... Lib "..."` -- a Win32 API import.
    /// Parsed so real modules do not fail to check; calling one is explicitly
    /// out of scope for the interpreter.
    Declare {
        /// [LLM-generated] The aliased procedure name.
        name: String,
        /// [LLM-generated] Whether it returns a value.
        is_function: bool,
        /// [LLM-generated] The library name.
        lib: String,
        /// [LLM-generated] An `Alias` clause, if written.
        alias: Option<String>,
        /// [LLM-generated] Its parameters.
        params: Vec<Param>,
        /// [LLM-generated] Its return type.
        return_type: Option<TypeRef>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Event Foo(args)` in a class module.
    EventDef {
        /// [LLM-generated] The event name.
        name: String,
        /// [LLM-generated] Its parameters.
        params: Vec<Param>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `RaiseEvent Foo(args)`.
    RaiseEvent {
        /// [LLM-generated] The event name.
        name: String,
        /// [LLM-generated] The arguments.
        args: Vec<Arg>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Implements IFoo`.
    Implements {
        /// [LLM-generated] The interface name.
        name: String,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Stop` -- breaks into the debugger.
    Stop {
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `End` as a statement -- halts execution. Distinct from the `End X`
    /// that closes a block, which never reaches the AST.
    End {
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] A statement this parser accepts but models only as its source text.
    ///
    /// VBA's file I/O and legacy graphics statements (`Open`, `Print #`,
    /// `Line Input`, `Name x As y`) have irregular syntax that would triple
    /// the grammar for constructs the interpreter is never going to run --
    /// they are out of scope by the plan's own security posture. Keeping them
    /// as opaque text means a real-world module still passes `macro check`
    /// instead of failing on a line nobody will execute.
    Opaque {
        /// [LLM-generated] The leading keyword, for a later "unsupported" diagnostic.
        keyword: String,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
    /// [LLM-generated] `Attribute Item.VB_UserMemId = 0` or other statement-level attribute.
    Attribute {
        /// [LLM-generated] The attribute name.
        name: String,
        /// [LLM-generated] Its values.
        values: Vec<Expr>,
        /// [LLM-generated] Where it starts.
        pos: Pos,
    },
}

/// [LLM-generated] A unary operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOp {
    /// [LLM-generated] Arithmetic negation.
    Neg,
    /// [LLM-generated] Unary plus, which VBA accepts and which is not a no-op on a String.
    Pos,
    /// [LLM-generated] Logical/bitwise `Not`.
    Not,
}

/// [LLM-generated] A binary operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    /// [LLM-generated] `^`.
    Pow,
    /// [LLM-generated] `*`.
    Mul,
    /// [LLM-generated] `/` -- floating-point division.
    Div,
    /// [LLM-generated] `\` -- integer division.
    IntDiv,
    /// [LLM-generated] `Mod`.
    Mod,
    /// [LLM-generated] `+`.
    Add,
    /// [LLM-generated] `-`.
    Sub,
    /// [LLM-generated] `&` -- string concatenation.
    Concat,
    /// [LLM-generated] `=` used as a comparison.
    Eq,
    /// [LLM-generated] `<>`.
    Ne,
    /// [LLM-generated] `<`.
    Lt,
    /// [LLM-generated] `>`.
    Gt,
    /// [LLM-generated] `<=`.
    Le,
    /// [LLM-generated] `>=`.
    Ge,
    /// [LLM-generated] `Is` -- reference identity.
    Is,
    /// [LLM-generated] `Like` -- pattern match.
    Like,
    /// [LLM-generated] `And`.
    And,
    /// [LLM-generated] `Or`.
    Or,
    /// [LLM-generated] `Xor`.
    Xor,
    /// [LLM-generated] `Eqv` -- logical equivalence.
    Eqv,
    /// [LLM-generated] `Imp` -- logical implication.
    Imp,
}

/// [LLM-generated] A literal value.
#[derive(Debug, Clone, PartialEq)]
pub enum Literal {
    /// [LLM-generated] A number, with how it was written.
    Number {
        /// [LLM-generated] Its value.
        value: f64,
        /// [LLM-generated] Decimal, hex or octal.
        base: NumBase,
        /// [LLM-generated] A trailing type-declaration character.
        suffix: Option<TypeSuffix>,
        /// [LLM-generated] Whether it was written with a fraction or exponent, which forces
        /// `Double`. See [`super::lexer::TokenKind::Number`].
        is_float: bool,
    },
    /// [LLM-generated] A string.
    Str(String),
    /// [LLM-generated] A `#...#` date, as its raw text.
    Date(String),
    /// [LLM-generated] `True` or `False`.
    Bool(bool),
    /// [LLM-generated] `Nothing`.
    Nothing,
    /// [LLM-generated] `Empty`.
    Empty,
    /// [LLM-generated] `Null`.
    Null,
}

/// [LLM-generated] One argument at a call site.
#[derive(Debug, Clone, PartialEq)]
pub struct Arg {
    /// [LLM-generated] A `name:=value` argument's name.
    pub name: Option<String>,
    /// [LLM-generated] The value, or `None` for an omitted positional argument (`f(1, , 3)`),
    /// which VBA allows and which is not the same as passing `Empty`.
    pub value: Option<Expr>,
}

/// [LLM-generated] An expression.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    /// [LLM-generated] A literal.
    Literal(Literal),
    /// [LLM-generated] A bare name.
    Ident {
        /// [LLM-generated] The name, in its original spelling.
        name: String,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] `Me`.
    Me {
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] Member access: `a.b`, or `.b` inside a `With` when `target` is `None`.
    Member {
        /// [LLM-generated] What is being accessed, or `None` for a leading-dot reference.
        target: Option<Box<Expr>>,
        /// [LLM-generated] The member name.
        name: String,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] Dictionary access: `rs!Field`.
    Bang {
        /// [LLM-generated] The object.
        target: Box<Expr>,
        /// [LLM-generated] The key.
        name: String,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] A call or an array index -- indistinguishable in VBA's syntax, and
    /// deliberately not distinguished here, since telling them apart needs
    /// the symbol table that Phase 1 will have and Phase 0 does not.
    Call {
        /// [LLM-generated] What is being called or indexed.
        target: Box<Expr>,
        /// [LLM-generated] The arguments.
        args: Vec<Arg>,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] A unary operation.
    Unary {
        /// [LLM-generated] Which operator.
        op: UnOp,
        /// [LLM-generated] Its operand.
        expr: Box<Expr>,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] A binary operation.
    Binary {
        /// [LLM-generated] Which operator.
        op: BinOp,
        /// [LLM-generated] The left operand.
        lhs: Box<Expr>,
        /// [LLM-generated] The right operand.
        rhs: Box<Expr>,
        /// [LLM-generated] Where the operator is.
        pos: Pos,
    },
    /// [LLM-generated] A parenthesised expression, kept rather than folded away because in
    /// VBA `Foo (x)` forces `x` to be passed by value even when the parameter
    /// is `ByRef` -- the parentheses are semantic, not just grouping.
    Paren {
        /// [LLM-generated] The inner expression.
        expr: Box<Expr>,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] `New Foo`.
    New {
        /// [LLM-generated] The type being constructed.
        ty: TypeRef,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] `TypeOf x Is Foo`.
    TypeOf {
        /// [LLM-generated] The value being tested.
        expr: Box<Expr>,
        /// [LLM-generated] The type it is tested against.
        ty: TypeRef,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
    /// [LLM-generated] `AddressOf Foo`.
    AddressOf {
        /// [LLM-generated] The procedure name.
        name: String,
        /// [LLM-generated] Where it is.
        pos: Pos,
    },
}

impl Expr {
    /// [LLM-generated] Where this expression starts.
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
    /// [LLM-generated] Every procedure declared in this module, in source order.
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
