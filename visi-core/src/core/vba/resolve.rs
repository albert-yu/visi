use super::ast::{
    Arg, CaseMatch, Expr, Literal, Module, ModuleItem, Param, Procedure, Stmt, TypeRef, VarDecl,
};
use super::builtin_names::is_builtin;
use super::parser::ParseError;
use std::collections::HashMap;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    /// A `Sub`, `Function`, `Property` or `Declare` -- a legitimate call
    /// target.
    Callable,
    /// Declared with array bounds (`x()`, `x(10)`) -- `x(i)` is indexing,
    /// not a call, so this is left alone rather than rejected.
    Array,
    /// A plain scalar: no array bounds, and either untyped (defaults to
    /// `Variant`) or typed as one of VBA's primitive scalar types. This is
    /// the only kind an implicit-call statement is rejected for.
    PlainScalar,
    /// An object-shaped declared type (`As New X`, `As SomeClass`, a dotted
    /// path) -- could have a default member callable with arguments, and
    /// this pass cannot resolve user-defined class shapes, so it is left
    /// alone.
    Opaque,
}

const PRIMITIVE_SCALAR_TYPES: &[&str] = &[
    "integer", "long", "single", "double", "currency", "string", "boolean", "byte", "date",
    "variant",
];

fn kind_for_decl(is_array: bool, ty: &Option<TypeRef>) -> Kind {
    if is_array {
        return Kind::Array;
    }
    match ty {
        None => Kind::PlainScalar,
        Some(t) if t.is_new => Kind::Opaque,
        Some(t) if t.path.len() == 1 && is_primitive_scalar(&t.path[0]) => Kind::PlainScalar,
        Some(_) => Kind::Opaque,
    }
}

fn is_primitive_scalar(name: &str) -> bool {
    PRIMITIVE_SCALAR_TYPES.contains(&name.to_ascii_lowercase().as_str())
}

const TYPE_SUFFIXES: [char; 6] = ['$', '%', '&', '!', '#', '@'];

fn norm(name: &str) -> String {
    name.strip_suffix(TYPE_SUFFIXES)
        .unwrap_or(name)
        .to_ascii_lowercase()
}

/// What a call target may resolve against, beyond the module's own text.
pub(super) struct Scope<'a> {
    /// Names declared by *other* modules in the same project. Empty when
    /// the caller has no project to consult.
    pub external: &'a std::collections::HashSet<String>,
    /// Whether [`Scope::external`] is known to cover every other module in
    /// the project.
    ///
    /// This gates the whole undeclared-name rule, and is the safety valve
    /// the design turns on. Excel compiles a *project*, so `x = arr(1)` is
    /// legal whenever any module declares `arr` -- meaning a checker
    /// looking at one module of several genuinely cannot tell an
    /// undeclared name from a cross-module reference. With this `false`,
    /// an unresolvable name is accepted and only the
    /// definitely-not-callable rule applies.
    pub complete_project: bool,
}

impl Scope<'_> {
    /// A scope for source that is the whole project as far as anyone knows
    /// -- a standalone `.bas`, or the single generated module the
    /// differential harness asks Excel about.
    pub fn self_contained(empty: &std::collections::HashSet<String>) -> Scope<'_> {
        Scope {
            external: empty,
            complete_project: true,
        }
    }

    /// A scope for one module of a project whose other modules were not
    /// supplied. Never rejects an unresolvable name.
    pub fn partial(empty: &std::collections::HashSet<String>) -> Scope<'_> {
        Scope {
            external: empty,
            complete_project: false,
        }
    }

    fn knows(&self, lower: &str) -> bool {
        self.external.contains(lower) || is_builtin(lower)
    }
}

/// Checks `module`'s call targets against the symbol table built from its
/// own text plus `scope`, per this module's doc.
pub(super) fn check_module(module: &Module, scope: &Scope<'_>) -> Result<(), ParseError> {
    let module_syms = collect_module_symbols(module);
    check_items(&module.items, &module_syms, scope)
}

/// Every name `module` declares at module level, already [`norm`]alised --
/// what a sibling module's [`Scope::external`] is built from, and so keyed
/// the same way [`Ctx::known`] will look them up.
pub(super) fn declared_names(module: &Module) -> Vec<String> {
    collect_module_symbols(module).into_keys().collect()
}

fn check_items(
    items: &[ModuleItem],
    module_syms: &HashMap<String, Kind>,
    scope: &Scope<'_>,
) -> Result<(), ParseError> {
    for item in items {
        match item {
            ModuleItem::Procedure(p) => check_procedure(p, module_syms, scope)?,
            ModuleItem::Conditional {
                branches,
                else_items,
                ..
            } => {
                for (_, body) in branches {
                    check_items(body, module_syms, scope)?;
                }
                if let Some(body) = else_items {
                    check_items(body, module_syms, scope)?;
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn collect_module_symbols(module: &Module) -> HashMap<String, Kind> {
    let mut syms = HashMap::new();
    collect_module_items(&module.items, &mut syms);
    syms
}

fn collect_module_items(items: &[ModuleItem], syms: &mut HashMap<String, Kind>) {
    for item in items {
        match item {
            ModuleItem::Procedure(p) => {
                syms.insert(norm(&p.name), Kind::Callable);
            }
            ModuleItem::Declaration(stmt) => collect_decl_stmt(stmt, syms),
            ModuleItem::Conditional {
                branches,
                else_items,
                ..
            } => {
                for (_, body) in branches {
                    collect_module_items(body, syms);
                }
                if let Some(body) = else_items {
                    collect_module_items(body, syms);
                }
            }
            ModuleItem::Attribute { .. } | ModuleItem::Option { .. } => {}
        }
    }
}

fn collect_decl_stmt(stmt: &Stmt, syms: &mut HashMap<String, Kind>) {
    match stmt {
        Stmt::Dim { vars, .. } => insert_var_decls(vars, syms),
        Stmt::Const { vars, .. } => insert_var_decls(vars, syms),
        Stmt::Declare { name, .. } => {
            syms.insert(norm(name), Kind::Callable);
        }
        Stmt::TypeDef { name, .. } => {
            syms.insert(norm(name), Kind::Opaque);
        }
        Stmt::EnumDef { name, members, .. } => {
            syms.insert(norm(name), Kind::Opaque);
            for m in members {
                syms.insert(norm(&m.name), Kind::Opaque);
            }
        }
        Stmt::EventDef { name, .. } => {
            syms.insert(norm(name), Kind::Callable);
        }
        _ => {}
    }
}

fn insert_var_decls(vars: &[VarDecl], syms: &mut HashMap<String, Kind>) {
    for v in vars {
        syms.insert(norm(&v.name), kind_for_decl(v.bounds.is_some(), &v.ty));
    }
}

fn check_procedure(
    proc: &Procedure,
    module_syms: &HashMap<String, Kind>,
    scope: &Scope<'_>,
) -> Result<(), ParseError> {
    let mut locals = HashMap::new();
    locals.insert(norm(&proc.name), Kind::Callable);
    for param in &proc.params {
        locals.insert(
            norm(&param.name),
            kind_for_decl(param.is_array || param.param_array, &param.ty),
        );
    }
    collect_locals(&proc.body, &mut locals);
    collect_implicit_locals(&proc.body, &mut locals);
    check_duplicate_declarations(&proc.body, &proc.params)?;
    let ctx = Ctx {
        module: module_syms,
        locals: &locals,
        scope,
    };
    check_block(&proc.body, &ctx)
}

struct Ctx<'a> {
    module: &'a HashMap<String, Kind>,
    locals: &'a HashMap<String, Kind>,
    scope: &'a Scope<'a>,
}

impl Ctx<'_> {
    fn kind_of(&self, lower: &str) -> Option<Kind> {
        self.locals
            .get(lower)
            .or_else(|| self.module.get(lower))
            .copied()
    }

    fn known(&self, lower: &str) -> bool {
        self.kind_of(lower).is_some() || self.scope.knows(lower)
    }
}

fn collect_locals(body: &[Stmt], locals: &mut HashMap<String, Kind>) {
    for stmt in body {
        match stmt {
            Stmt::Dim { vars, .. } => insert_var_decls(vars, locals),
            Stmt::Const { vars, .. } => insert_var_decls(vars, locals),
            Stmt::ReDim {
                preserve: false,
                vars,
                ..
            } => {
                for v in vars {
                    locals.entry(norm(&v.name)).or_insert(Kind::Array);
                }
            }
            Stmt::If {
                branches,
                else_body,
                ..
            } => {
                for (_, b) in branches {
                    collect_locals(b, locals);
                }
                if let Some(b) = else_body {
                    collect_locals(b, locals);
                }
            }
            Stmt::SelectCase {
                cases, case_else, ..
            } => {
                for c in cases {
                    collect_locals(&c.body, locals);
                }
                if let Some(b) = case_else {
                    collect_locals(b, locals);
                }
            }
            Stmt::For { body, .. }
            | Stmt::ForEach { body, .. }
            | Stmt::DoLoop { body, .. }
            | Stmt::With { body, .. } => collect_locals(body, locals),
            _ => {}
        }
    }
}

fn collect_implicit_locals(body: &[Stmt], locals: &mut HashMap<String, Kind>) {
    for stmt in body {
        match stmt {
            Stmt::Assign {
                target: Expr::Ident { name, .. },
                set,
                ..
            } => {
                let kind = if *set {
                    Kind::Opaque
                } else {
                    Kind::PlainScalar
                };
                locals.entry(norm(name)).or_insert(kind);
            }
            Stmt::For { var, body, .. } => {
                if let Expr::Ident { name, .. } = var {
                    locals.entry(norm(name)).or_insert(Kind::PlainScalar);
                }
                collect_implicit_locals(body, locals);
            }
            Stmt::ForEach { var, body, .. } => {
                if let Expr::Ident { name, .. } = var {
                    locals.entry(norm(name)).or_insert(Kind::Opaque);
                }
                collect_implicit_locals(body, locals);
            }
            Stmt::If {
                branches,
                else_body,
                ..
            } => {
                for (_, b) in branches {
                    collect_implicit_locals(b, locals);
                }
                if let Some(b) = else_body {
                    collect_implicit_locals(b, locals);
                }
            }
            Stmt::SelectCase {
                cases, case_else, ..
            } => {
                for c in cases {
                    collect_implicit_locals(&c.body, locals);
                }
                if let Some(b) = case_else {
                    collect_implicit_locals(b, locals);
                }
            }
            Stmt::DoLoop { body, .. } | Stmt::With { body, .. } => {
                collect_implicit_locals(body, locals);
            }
            _ => {}
        }
    }
}

fn check_duplicate_declarations(body: &[Stmt], params: &[Param]) -> Result<(), ParseError> {
    let mut in_scope: std::collections::HashSet<String> =
        params.iter().map(|p| norm(&p.name)).collect();
    walk_declarations(body, &mut in_scope)
}

fn walk_declarations(
    body: &[Stmt],
    in_scope: &mut std::collections::HashSet<String>,
) -> Result<(), ParseError> {
    for stmt in body {
        match stmt {
            Stmt::Dim { vars, .. } | Stmt::Const { vars, .. } => {
                for v in vars {
                    if !in_scope.insert(norm(&v.name)) {
                        return Err(ParseError {
                            message: format!("Duplicate declaration in current scope: {}", v.name),
                            pos: v.pos,
                        });
                    }
                }
            }
            Stmt::Assign {
                target: Expr::Ident { name, .. },
                ..
            } => {
                in_scope.insert(norm(name));
            }
            Stmt::ReDim { vars, .. } => {
                for v in vars {
                    in_scope.insert(norm(&v.name));
                }
            }
            Stmt::If {
                branches,
                else_body,
                ..
            } => {
                for (_, b) in branches {
                    walk_declarations(b, in_scope)?;
                }
                if let Some(b) = else_body {
                    walk_declarations(b, in_scope)?;
                }
            }
            Stmt::SelectCase {
                cases, case_else, ..
            } => {
                for c in cases {
                    walk_declarations(&c.body, in_scope)?;
                }
                if let Some(b) = case_else {
                    walk_declarations(b, in_scope)?;
                }
            }
            Stmt::For { var, body, .. } | Stmt::ForEach { var, body, .. } => {
                if let Expr::Ident { name, .. } = var {
                    in_scope.insert(norm(name));
                }
                walk_declarations(body, in_scope)?;
            }
            Stmt::DoLoop { body, .. } | Stmt::With { body, .. } => {
                walk_declarations(body, in_scope)?
            }
            _ => {}
        }
    }
    Ok(())
}

fn check_block(body: &[Stmt], ctx: &Ctx<'_>) -> Result<(), ParseError> {
    for stmt in body {
        check_stmt(stmt, ctx)?;
    }
    Ok(())
}

fn check_stmt(stmt: &Stmt, ctx: &Ctx<'_>) -> Result<(), ParseError> {
    match stmt {
        Stmt::Call {
            expr: Expr::Ident { name, pos },
            ..
        } => check_call_name(name, *pos, ctx),
        Stmt::Call { expr, .. } => check_expr(expr, ctx),
        Stmt::Assign { target, value, .. } => {
            check_expr(target, ctx)?;
            check_expr(value, ctx)
        }
        Stmt::ReDim { preserve, vars, .. } => {
            for v in vars {
                if *preserve && ctx.scope.complete_project {
                    let lower = norm(&v.name);
                    if !ctx.known(&lower) {
                        return Err(ParseError {
                            message: format!("Variable not defined: {}", v.name),
                            pos: v.pos,
                        });
                    }
                }
                check_var_decl_exprs(v, ctx)?;
            }
            Ok(())
        }
        Stmt::Dim { vars, .. } | Stmt::Const { vars, .. } => {
            for v in vars {
                check_var_decl_exprs(v, ctx)?;
            }
            Ok(())
        }
        Stmt::If {
            branches,
            else_body,
            ..
        } => {
            for (cond, b) in branches {
                check_expr(cond, ctx)?;
                if !expr_is_literal_false(cond) {
                    check_block(b, ctx)?;
                }
            }
            if let Some(b) = else_body {
                check_block(b, ctx)?;
            }
            Ok(())
        }
        Stmt::SelectCase {
            subject,
            cases,
            case_else,
            ..
        } => {
            check_expr(subject, ctx)?;
            for c in cases {
                for m in &c.matches {
                    match m {
                        CaseMatch::Value(e) | CaseMatch::Is(_, e) => check_expr(e, ctx)?,
                        CaseMatch::Range(a, b) => {
                            check_expr(a, ctx)?;
                            check_expr(b, ctx)?;
                        }
                    }
                }
                check_block(&c.body, ctx)?;
            }
            if let Some(b) = case_else {
                check_block(b, ctx)?;
            }
            Ok(())
        }
        Stmt::For {
            from,
            to,
            step,
            body,
            ..
        } => {
            check_expr(from, ctx)?;
            check_expr(to, ctx)?;
            if let Some(s) = step {
                check_expr(s, ctx)?;
            }
            check_block(body, ctx)
        }
        Stmt::ForEach { iterable, body, .. } => {
            check_expr(iterable, ctx)?;
            check_block(body, ctx)
        }
        Stmt::DoLoop {
            pre, post, body, ..
        } => {
            for (_, e) in pre.iter().chain(post.iter()) {
                check_expr(e, ctx)?;
            }
            check_block(body, ctx)
        }
        Stmt::With { subject, body, .. } => {
            check_expr(subject, ctx)?;
            check_block(body, ctx)
        }
        Stmt::Erase { targets, .. } => {
            for t in targets {
                check_expr(t, ctx)?;
            }
            Ok(())
        }
        Stmt::RaiseEvent { args, .. } => check_args(args, ctx),
        _ => Ok(()),
    }
}

fn check_var_decl_exprs(v: &VarDecl, ctx: &Ctx<'_>) -> Result<(), ParseError> {
    if let Some(bounds) = &v.bounds {
        for b in bounds {
            if let Some(l) = &b.lower {
                check_expr(l, ctx)?;
            }
            check_expr(&b.upper, ctx)?;
        }
    }
    if let Some(val) = &v.value {
        check_expr(val, ctx)?;
    }
    Ok(())
}

fn check_args(args: &[Arg], ctx: &Ctx<'_>) -> Result<(), ParseError> {
    for a in args {
        if let Some(v) = &a.value {
            check_expr(v, ctx)?;
        }
    }
    Ok(())
}

fn expr_is_literal_false(expr: &Expr) -> bool {
    matches!(expr, Expr::Literal(Literal::Bool(false)))
}

fn check_expr(expr: &Expr, ctx: &Ctx<'_>) -> Result<(), ParseError> {
    match expr {
        Expr::Call { target, args, .. } => {
            if let Expr::Ident { name, pos } = target.as_ref() {
                check_call_name(name, *pos, ctx)?;
            } else {
                check_expr(target, ctx)?;
            }
            check_args(args, ctx)
        }
        Expr::Member { target, .. } => {
            if let Some(t) = target {
                check_expr(t, ctx)?;
            }
            Ok(())
        }
        Expr::Bang { target, .. } => check_expr(target, ctx),
        Expr::Unary { expr, .. } | Expr::Paren { expr, .. } | Expr::TypeOf { expr, .. } => {
            check_expr(expr, ctx)
        }
        Expr::Binary { lhs, rhs, .. } => {
            check_expr(lhs, ctx)?;
            check_expr(rhs, ctx)
        }
        _ => Ok(()),
    }
}

fn check_call_name(name: &str, pos: super::lexer::Pos, ctx: &Ctx<'_>) -> Result<(), ParseError> {
    let lower = norm(name);
    match ctx.kind_of(&lower) {
        Some(Kind::PlainScalar) => Err(ParseError {
            message: format!("Sub or Function not defined: {name}"),
            pos,
        }),
        Some(_) => Ok(()),
        None if ctx.scope.complete_project && !ctx.scope.knows(&lower) => Err(ParseError {
            message: format!("Sub or Function not defined: {name}"),
            pos,
        }),
        None => Ok(()),
    }
}

#[cfg(test)]
mod tests {
    use super::super::parser::parse_module;
    use super::{Scope, check_module, norm};
    use std::collections::HashSet;

    fn check(src: &str) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        let empty = HashSet::new();
        check_module(&module, &Scope::self_contained(&empty)).map_err(|e| e.message)
    }

    fn check_partial(src: &str) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        let empty = HashSet::new();
        check_module(&module, &Scope::partial(&empty)).map_err(|e| e.message)
    }

    fn check_with_external(src: &str, names: &[&str]) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        let external: HashSet<String> = names.iter().map(|n| norm(n)).collect();
        let scope = Scope {
            external: &external,
            complete_project: true,
        };
        check_module(&module, &scope).map_err(|e| e.message)
    }

    #[test]
    fn rejects_local_plain_scalar_as_call_target() {
        let err = check("Sub Test()\n    Dim x As Long\n    x 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn rejects_module_level_plain_scalar_as_call_target() {
        let src = "Dim g As Long\n\nSub Test()\n    g 5\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: g"), "{err}");
    }

    #[test]
    fn rejects_untyped_dim_as_call_target() {
        let err = check("Sub Test()\n    Dim x\n    x 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn rejects_parameter_as_call_target() {
        let err = check("Sub Test(x As Long)\n    x 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn accepts_call_to_declared_procedure() {
        let src = "Sub Test()\n    Foo 5\nEnd Sub\n\nSub Foo(n As Long)\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_call_to_a_builtin_name() {
        assert!(check("Sub Test()\n    MsgBox \"hi\"\nEnd Sub\n").is_ok());
        assert!(check("Sub Test()\n    x = Split(\"a,b\", \",\")\nEnd Sub\n").is_ok());
        assert!(check("Sub Test()\n    x = Range(\"A1\")\nEnd Sub\n").is_ok());
    }

    #[test]
    fn accepts_array_indexing() {
        let src = "Sub Test()\n    Dim arr(10) As Long\n    arr(3) = 1\n    y = arr(3)\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_object_shaped_local_as_call_target() {
        let src = "Sub Test()\n    Dim obj As Collection\n    obj 3\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn rejects_a_bare_identifier_statement_that_is_not_callable() {
        for src in [
            "Sub Test()\n    x\nEnd Sub\n",
            "Sub Test()\n    Dim x As Long\n    x\nEnd Sub\n",
            "Sub Test()\n    x = 1\n    x\nEnd Sub\n",
        ] {
            assert!(check(src).is_err(), "should have been rejected: {src}");
        }
    }

    #[test]
    fn accepts_a_bare_statement_naming_something_callable() {
        assert!(check("Sub Test()\n    Helper\nEnd Sub\n\nSub Helper()\nEnd Sub\n").is_ok());
        assert!(check("Sub Test()\n    Beep\nEnd Sub\n").is_ok());
    }

    #[test]
    fn a_type_suffix_does_not_hide_a_name() {
        assert!(check("Sub Test()\n    x = Trim$(\" a \")\nEnd Sub\n").is_ok());
        assert!(check("Sub Test()\n    x = Left$(\"ab\", 1)\nEnd Sub\n").is_ok());
        assert!(check("Sub Test()\n    Dim s$\n    s = Trim(s$)\nEnd Sub\n").is_ok());
        let src =
            "Function F$(a%)\n    F = CStr(a)\nEnd Function\n\nSub T()\n    x = F(1)\nEnd Sub\n";
        assert!(check(src).is_ok());
        let err = check("Sub Test()\n    Dim s$\n    s$ 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined"), "{err}");
    }

    #[test]
    fn rejects_a_duplicate_declaration() {
        for src in [
            "Sub T()\n    Dim x As Long\n    Dim x As Long\nEnd Sub\n",
            "Sub T()\n    x = 1\n    Dim x As Long\nEnd Sub\n",
            "Sub T()\n    x = Helper(1)\n    Dim x As Long\nEnd Sub\n\nFunction Helper(a)\nEnd Function\n",
        ] {
            let err = check(src).unwrap_err();
            assert!(err.contains("Duplicate declaration"), "{src} gave {err}");
        }
    }

    #[test]
    fn every_route_into_scope_collides_with_a_later_dim() {
        for src in [
            "Sub T(ByVal x As Long)\n    Dim x As Long\nEnd Sub\n",
            "Sub T()\n    For x = 1 To 3\n        y = x\n    Next x\n    Dim x As Long\nEnd Sub\n",
            "Sub T()\n    For Each x In rng\n        y = 1\n    Next\n    Dim x As Long\nEnd Sub\n",
            "Sub T()\n    Set x = New Collection\n    Dim x As Object\nEnd Sub\n",
            "Sub T()\n    ReDim arr(1 To 5)\n    Dim arr()\nEnd Sub\n",
        ] {
            let err = check(src).unwrap_err();
            assert!(err.contains("Duplicate declaration"), "{src} gave {err}");
        }
    }

    #[test]
    fn a_route_into_scope_is_not_itself_a_declaration_error() {
        for src in [
            "Sub T(ByVal x As Long)\n    y = x\nEnd Sub\n",
            "Sub T()\n    For x = 1 To 3\n        y = x\n    Next x\nEnd Sub\n",
            "Sub T()\n    For Each x In rng\n        y = 1\n    Next\nEnd Sub\n",
            "Sub T()\n    Set x = New Collection\nEnd Sub\n",
            "Sub T()\n    Dim arr()\n    ReDim arr(1 To 5)\nEnd Sub\n",
        ] {
            assert!(check(src).is_ok(), "should have been accepted: {src}");
        }
    }

    #[test]
    fn declaring_before_assigning_is_ordinary_code() {
        assert!(check("Sub T()\n    Dim x As Long\n    x = 1\nEnd Sub\n").is_ok());
        let src =
            "Sub T()\n    Dim x As Long\n    If True Then\n        x = 1\n    End If\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn a_duplicate_is_caught_across_block_boundaries() {
        let src = "Sub T()\n    If True Then\n        Dim x As Long\n    Else\n        Dim x As Long\n    End If\nEnd Sub\n";
        assert!(check(src).unwrap_err().contains("Duplicate declaration"));
        let src = "Sub T()\n    Do While x < 10\n        x = x + 1\n    Loop\n    Dim x As Long\nEnd Sub\n";
        assert!(check(src).unwrap_err().contains("Duplicate declaration"));
    }

    #[test]
    fn a_module_level_name_is_not_a_duplicate_of_a_local() {
        let src = "Dim x As Long\n\nSub T()\n    Dim x As Long\nEnd Sub\n";
        assert!(check(src).is_ok());
        let src = "Sub A()\n    Dim x As Long\nEnd Sub\n\nSub B()\n    Dim x As Long\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn a_bare_name_inside_an_expression_stays_unchecked() {
        assert!(check("Sub Test()\n    x = a + b\nEnd Sub\n").is_ok());
    }

    #[test]
    fn accepts_call_keyword_to_declared_procedure() {
        let src = "Sub Test()\n    Call Foo(5)\nEnd Sub\n\nSub Foo(n As Long)\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn rejects_local_shadowing_a_module_procedure_name() {
        let src = "Sub Foo()\nEnd Sub\n\nSub Test()\n    Dim Bar As Long\n    Bar 5\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: Bar"), "{err}");
    }

    #[test]
    fn rejects_a_name_implicitly_declared_by_plain_assignment() {
        let src = "Sub Test()\n    x = 1\n    x 5\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn rejects_a_name_implicitly_declared_before_its_first_assignment() {
        let src = "Sub Test()\n    x 5\n    x = 1\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn rejects_a_for_loop_counter_used_as_a_call_target() {
        let src = "Sub Test()\n    For i = 1 To 3\n        i 5\n    Next i\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: i"), "{err}");
    }

    #[test]
    fn rejects_an_undeclared_name_used_with_call_syntax() {
        let err = check("Sub Test()\n    arr 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: arr"), "{err}");
        let err = check("Sub Test()\n    d #1/1/2000#\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: d"), "{err}");
    }

    #[test]
    fn rejects_an_undeclared_call_in_expression_position() {
        let err = check("Sub Test()\n    x = arr(1, 2)\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: arr"), "{err}");
    }

    #[test]
    fn name_resolution_skips_statically_false_if_bodies() {
        assert!(
            check("Sub Test()\n    If False Then\n        x = arr(1, 2)\n    End If\nEnd Sub\n")
                .is_ok()
        );
    }

    #[test]
    fn a_partial_scope_never_rejects_an_unresolvable_name() {
        assert!(check_partial("Sub Test()\n    arr 5\nEnd Sub\n").is_ok());
        assert!(check_partial("Sub Test()\n    x = arr(1, 2)\nEnd Sub\n").is_ok());
        let err = check_partial("Sub Test()\n    Dim x As Long\n    x 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn a_sibling_modules_name_resolves() {
        assert!(check_with_external("Sub Test()\n    x = arr(1, 2)\nEnd Sub\n", &["arr"]).is_ok());
        assert!(check_with_external("Sub Test()\n    Helper 5\nEnd Sub\n", &["helper"]).is_ok());
    }

    #[test]
    fn rejects_redim_preserve_on_an_undeclared_name() {
        let err = check("Sub Test()\n    ReDim Preserve arr(1 To 5)\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Variable not defined: arr"), "{err}");
    }

    #[test]
    fn accepts_plain_redim_which_declares_the_array() {
        assert!(check("Sub Test()\n    ReDim arr(1 To 5)\nEnd Sub\n").is_ok());
        let src = "Sub Test()\n    ReDim arr(1 To 5)\n    ReDim Preserve arr(1 To 9)\nEnd Sub\n";
        assert!(check(src).is_ok());
        let src = "Sub Test()\n    Dim arr()\n    ReDim Preserve arr(1 To 5)\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_an_undeclared_bare_name_with_no_call_syntax() {
        assert!(check("Sub Test()\n    x = a + b\nEnd Sub\n").is_ok());
        let src = "Sub Test()\n    For Each c In rng\n        x = 1\n    Next\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_a_set_assignment_target_as_a_call_target() {
        let src = "Sub Test()\n    Set x = Nothing\n    x 5\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_a_for_each_element_variable_as_a_call_target() {
        let src = "Sub Test()\n    For Each c In rng\n        c 5\n    Next c\nEnd Sub\n";
        assert!(check(src).is_ok());
    }
}
