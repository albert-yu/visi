use super::lexer::{NumBase, Pos, TypeSuffix};

#[derive(Debug, Clone, PartialEq)]
pub struct Module {
    pub items: Vec<ModuleItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ModuleItem {
    Attribute {
        name: String,
        values: Vec<Expr>,
        pos: Pos,
    },
    Option {
        words: Vec<String>,
        pos: Pos,
    },
    Declaration(Stmt),
    Procedure(Procedure),
    Conditional {
        branches: Vec<(Expr, Vec<ModuleItem>)>,
        else_items: Option<Vec<ModuleItem>>,
        pos: Pos,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Public,
    Private,
    Friend,
    Global,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Procedure {
    pub kind: ProcKind,
    pub name: String,
    pub visibility: Option<Visibility>,
    pub is_static: bool,
    pub params: Vec<Param>,
    pub return_type: Option<TypeRef>,
    pub body: Vec<Stmt>,
    pub pos: Pos,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PassBy {
    Value,
    Reference,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Param {
    pub name: String,
    pub by: Option<PassBy>,
    pub optional: bool,
    pub param_array: bool,
    pub is_array: bool,
    pub ty: Option<TypeRef>,
    pub default: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TypeRef {
    pub path: Vec<String>,
    pub is_new: bool,
    pub string_length: Option<Box<Expr>>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct VarDecl {
    pub name: String,
    pub bounds: Option<Vec<ArrayBound>>,
    pub ty: Option<TypeRef>,
    pub value: Option<Expr>,
    pub pos: Pos,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ArrayBound {
    pub lower: Option<Expr>,
    pub upper: Expr,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DimKind {
    Dim,
    Static,
    Private,
    Public,
    Global,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    Property,
    For,
    Do,
    While,
}

#[derive(Debug, Clone, PartialEq)]
pub enum OnErrorKind {
    GoTo(String),
    ResumeNext,
    Disable,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ResumeKind {
    Retry,
    Next,
    Label(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DoTest {
    While,
    Until,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CaseClause {
    pub matches: Vec<CaseMatch>,
    pub body: Vec<Stmt>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CaseMatch {
    Value(Expr),
    Range(Expr, Expr),
    Is(BinOp, Expr),
}

#[derive(Debug, Clone, PartialEq)]
pub struct EnumMember {
    pub name: String,
    pub value: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Stmt {
    Dim {
        kind: DimKind,
        with_events: bool,
        vars: Vec<VarDecl>,
        pos: Pos,
    },
    Const {
        visibility: Option<Visibility>,
        vars: Vec<VarDecl>,
        pos: Pos,
    },
    ReDim {
        preserve: bool,
        vars: Vec<VarDecl>,
        pos: Pos,
    },
    Assign {
        target: Expr,
        value: Expr,
        set: bool,
        pos: Pos,
    },
    Call {
        expr: Expr,
        pos: Pos,
    },
    If {
        branches: Vec<(Expr, Vec<Stmt>)>,
        else_body: Option<Vec<Stmt>>,
        pos: Pos,
    },
    SelectCase {
        subject: Expr,
        cases: Vec<CaseClause>,
        case_else: Option<Vec<Stmt>>,
        pos: Pos,
    },
    For {
        var: Expr,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
        pos: Pos,
    },
    ForEach {
        var: Expr,
        iterable: Expr,
        body: Vec<Stmt>,
        pos: Pos,
    },
    DoLoop {
        pre: Option<(DoTest, Expr)>,
        post: Option<(DoTest, Expr)>,
        body: Vec<Stmt>,
        pos: Pos,
    },
    With {
        subject: Expr,
        body: Vec<Stmt>,
        pos: Pos,
    },
    Exit {
        kind: ExitKind,
        pos: Pos,
    },
    GoTo {
        label: String,
        pos: Pos,
    },
    GoSub {
        label: String,
        pos: Pos,
    },
    Return {
        pos: Pos,
    },
    OnError {
        kind: OnErrorKind,
        pos: Pos,
    },
    OnGoto {
        subject: Expr,
        labels: Vec<String>,
        gosub: bool,
        pos: Pos,
    },
    Resume {
        kind: ResumeKind,
        pos: Pos,
    },
    Label {
        name: String,
        pos: Pos,
    },
    Erase {
        targets: Vec<Expr>,
        pos: Pos,
    },
    TypeDef {
        name: String,
        visibility: Option<Visibility>,
        fields: Vec<VarDecl>,
        pos: Pos,
    },
    EnumDef {
        name: String,
        visibility: Option<Visibility>,
        members: Vec<EnumMember>,
        pos: Pos,
    },
    Declare {
        name: String,
        is_function: bool,
        lib: String,
        alias: Option<String>,
        params: Vec<Param>,
        return_type: Option<TypeRef>,
        pos: Pos,
    },
    EventDef {
        name: String,
        params: Vec<Param>,
        pos: Pos,
    },
    RaiseEvent {
        name: String,
        args: Vec<Arg>,
        pos: Pos,
    },
    Implements {
        name: String,
        pos: Pos,
    },
    Stop {
        pos: Pos,
    },
    End {
        pos: Pos,
    },
    Opaque {
        keyword: String,
        pos: Pos,
    },
    Attribute {
        name: String,
        values: Vec<Expr>,
        pos: Pos,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOp {
    Neg,
    Pos,
    Not,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    Pow,
    Mul,
    Div,
    IntDiv,
    Mod,
    Add,
    Sub,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    Is,
    Like,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Literal {
    Number {
        value: f64,
        base: NumBase,
        suffix: Option<TypeSuffix>,
        is_float: bool,
    },
    Str(String),
    Date(String),
    Bool(bool),
    Nothing,
    Empty,
    Null,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Arg {
    pub name: Option<String>,
    pub value: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Literal(Literal),
    Ident {
        name: String,
        pos: Pos,
    },
    Me {
        pos: Pos,
    },
    Member {
        target: Option<Box<Expr>>,
        name: String,
        pos: Pos,
    },
    Bang {
        target: Box<Expr>,
        name: String,
        pos: Pos,
    },
    Call {
        target: Box<Expr>,
        args: Vec<Arg>,
        pos: Pos,
    },
    Unary {
        op: UnOp,
        expr: Box<Expr>,
        pos: Pos,
    },
    Binary {
        op: BinOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
        pos: Pos,
    },
    Paren {
        expr: Box<Expr>,
        pos: Pos,
    },
    New {
        ty: TypeRef,
        pos: Pos,
    },
    TypeOf {
        expr: Box<Expr>,
        ty: TypeRef,
        pos: Pos,
    },
    AddressOf {
        name: String,
        pos: Pos,
    },
}

impl Expr {
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
