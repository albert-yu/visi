//! A narrow, conservative name-resolution pass on top of Phase 0's AST.
//!
//! `parser.rs` cannot "even tell a procedure call from an array index
//! without a symbol table" (see its module doc and
//! `docs/vba-macro-support.md`). [issue #78] quantified the fallout: 11 of
//! 100 win32com-generated cases were false negatives -- `check_syntax`
//! accepted source real Excel refuses to compile -- and traced essentially
//! all of them to one shape: an **implicit-call statement** (`x y`, no
//! parentheses, no leading `Call`) where `x` resolves to something that is
//! definitely not callable. Measured directly against Excel:
//!
//! ```text
//! x y = 1        ' compiles -- x is a real Sub taking a ParamArray
//! x (y = 1)      ' compiles -- a parenthesised argument, not "call syntax"
//! x y            ' FAILS -- x undeclared
//! Dim x : x y    ' FAILS -- x declared, but as a plain Variant, not callable
//! MsgBox "hi"    ' compiles -- MsgBox is a real Sub
//! ```
//!
//! This pass only ever *rejects*; it never turns an acceptable module into a
//! rejected one beyond that. The risk the issue calls out explicitly is
//! getting "undeclared" and "declared but non-callable" confused: rejecting
//! a name this pass simply hasn't seen would produce a false positive (a
//! macro `check_syntax` breaks even though Excel accepts it), which the
//! project's own docs call the worse failure mode. So the rule implemented
//! here is deliberately one-sided --
//!
//! - A call target that resolves (locally, or at module level) to a
//!   `Sub`/`Function`/`Property`/`Declare` is accepted.
//! - A call target that resolves to a plain scalar -- a `Dim`/`Static`/
//!   `Const`/parameter with no array bounds and no object-shaped declared
//!   type -- is rejected, mirroring the measured `Dim x : x y` case and the
//!   interpreter's own runtime rule in `interp.rs::eval_call` (a variable
//!   that is not an `Array` or an `Object` with a default member falls
//!   through to error 35, "Sub or Function not defined").
//! - A call target that resolves to an array, or to something declared with
//!   an object-shaped type (`As New X`, `As SomeClass`, a dotted type path),
//!   is left alone: real VBA lets a default-member call or array indexing
//!   use exactly this syntax, and this pass has no way to tell those apart
//!   from a non-callable use without full type resolution.
//! - A name that resolves **nowhere** -- not in this module, not in a
//!   sibling module, not in [`super::builtin_names`] -- is rejected, but
//!   *only* under a [`Scope`] that can see the whole project. See below.
//!
//! A bare identifier is checked **in statement position only**. `x` alone on
//! a line is a call -- VBA reads it as "call x, no arguments" -- and resolves
//! by the same rule; measured, it is rejected whether `x` is undeclared,
//! `Dim`'d as a `Long`, or created by assignment, while `Helper` and `Beep`
//! are fine. Inside an expression it is not: `x = a + b` with nothing
//! declared compiles, because `Option Explicit` being off makes `a` and `b`
//! implicit Variants.
//!
//! ## Undeclared names need whole-project scope
//!
//! Excel compiles a *project*, not a file, so `x = arr(1)` is legal as soon
//! as any module declares `arr`. A checker holding one module of several
//! therefore cannot distinguish an undeclared name from a cross-module
//! reference, and guessing would produce exactly the false positive the
//! issue warns about. [`Scope::complete_project`] gates the rule: with it
//! `false`, an unresolvable name is accepted and only the
//! definitely-not-callable rule above applies. `check_syntax` sets it (a
//! standalone `.bas` is the whole story), `VbaProject::check_modules` sets
//! it after unioning every module's names, and `VbaModule::check_syntax`
//! does not, having no project to consult.
//!
//! ## Rules here that are not about resolving a name
//!
//! The first two were assumed to be name-resolution failures and turned out
//! not to be, which is why they are called out. All measured with
//! `fuzz/vba_compile_probe.py`:
//!
//! - **`ReDim Preserve` requires an existing array.** Plain
//!   `ReDim arr(1 To 5)` *declares* `arr` and compiles with no `Dim`
//!   anywhere; adding `Preserve` makes the same line a compile error, since
//!   there is nothing to preserve. Handled in `collect_locals` (the plain
//!   form contributes a symbol) and `check_stmt` (the `Preserve` form is
//!   checked instead).
//! - **A name cannot start with `_`.** That one is not in this module at
//!   all -- it is a lexer rule, see the `'_'` arm of `lexer.rs`'s main loop.
//! - **Duplicate declaration**, in [`check_duplicate_declarations`] -- the
//!   one order-sensitive rule here, since `Dim x` then `x = 1` is ordinary
//!   code while the reverse is a compile error.
//!
//! ## Where this leaves issue #78
//!
//! All 12 saved false-negative reproductions
//! (`fuzz_results/failures/vba_parse_iter_*`) are now reported. They split
//! four ways, and only the first group is what the issue predicted:
//!
//! | Cause | Cases |
//! | --- | --- |
//! | undeclared name used with call syntax | iter_22, 40, 50, 64, 93 |
//! | a plain local used as a call target | iter_12, 15, 46 |
//! | `ReDim Preserve` on an undeclared name | iter_14, 57 |
//! | a name starting with `_` | iter_24 |
//! | built-in type keyword as a declared name (fixed earlier) | iter_4 |
//!
//! Validating that against two unseen seeds of 60 generated cases each
//! (4242 and 91177) gave **0 false positives on both**, which is the number
//! that matters here -- this is the first thing in the checker that can
//! reject a module for a reason other than its own syntax.
//!
//! Those runs turned up two more shapes not in the original twelve, both
//! since measured and implemented: the bare-identifier statement above, and
//! duplicate declaration. On seed 91177 *all five* surviving false
//! negatives were the latter -- one shape, not five gaps -- because the
//! generator emits its `Dim` last while ordinary code declares first.
//!
//! What the fuzzer cannot bound is worth stating plainly: its grammar never
//! emits a type-declaration suffix, so it could not have caught the one real
//! false positive this work introduced (`Trim$` resolving as `"trim$"`; see
//! [`norm`]). A differential run bounds the error rate *over the shapes it
//! generates*, which is narrower than what real modules contain.
//!
//! [issue #78]: https://github.com/albert-yu/visi/issues/78

use super::ast::{Arg, CaseMatch, Expr, Module, ModuleItem, Procedure, Stmt, TypeRef, VarDecl};
use super::builtin_names::is_builtin;
use super::parser::ParseError;
use std::collections::HashMap;

/// What a declared name is known to be, as far as this pass can tell.
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

/// VBA's primitive scalar type keywords. An untyped `Dim` defaults to
/// `Variant`, and `Variant` is on this list too -- measured directly: a bare
/// `Dim x` is exactly the case `docs/vba-macro-support.md`'s transcript
/// shows Excel rejecting as a call target.
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

/// VBA's type-declaration characters, in the spelling the lexer folds into
/// an identifier's name.
const TYPE_SUFFIXES: [char; 6] = ['$', '%', '&', '!', '#', '@'];

/// A name as this pass keys it: lowercased, with any trailing
/// type-declaration character removed.
///
/// `lexer.rs` deliberately folds a type suffix back into an identifier's
/// spelling, since `a$` is how the name was written -- but, as its own
/// comment says, **`a$` and `a` are the same variable**, and `Trim$` is the
/// same function as `Trim`. Resolution therefore has to strip it, and has to
/// do so on the declaring side and the referencing side alike or the two
/// stop meeting in the middle.
///
/// Getting this wrong rejected every `$` string intrinsic there is --
/// `Left$`, `Trim$`, `Mid$`, `Format$`, `UCase$` -- none of which the
/// generated fuzz grammar happens to emit, so it took hand-written
/// real-world VBA to surface.
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

/// Every `Sub`/`Function`/`Property`/`Declare` and module-level
/// `Dim`/`Const`/`Private`/`Public`/`Global` in the module, flattened across
/// `#If` branches exactly as [`Module::procedures`] already does -- which
/// branch is live depends on `#Const` values parsing alone cannot decide.
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
        // A `Type`/`Enum` name and an enum's members are all `Opaque`: they
        // are names that exist, which is all the undeclared-name rule needs
        // to know, and none of them is a plain scalar this pass would
        // reject a call on.
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
    // A `Function`'s own name is assignable inside it (`Harness = "OK"`) and
    // is also a legitimate recursive call target.
    locals.insert(norm(&proc.name), Kind::Callable);
    for param in &proc.params {
        locals.insert(
            norm(&param.name),
            kind_for_decl(param.is_array || param.param_array, &param.ty),
        );
    }
    collect_locals(&proc.body, &mut locals);
    collect_implicit_locals(&proc.body, &mut locals);
    check_duplicate_declarations(&proc.body)?;
    let ctx = Ctx {
        module: module_syms,
        locals: &locals,
        scope,
    };
    check_block(&proc.body, &ctx)
}

/// Everything one procedure's body resolves a name against.
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

    /// Whether the name exists at all, anywhere this pass can see.
    fn known(&self, lower: &str) -> bool {
        self.kind_of(lower).is_some() || self.scope.knows(lower)
    }
}

/// `Dim`/`Static`/`Const` are procedure-scoped in VBA, not block-scoped, so
/// this is a flat walk of every statement the procedure body contains,
/// regardless of how deeply nested in `If`/`For`/`Do`/`With`/`Select Case`
/// it is. VBA has no `#If` inside a procedure body (only at module level),
/// so unlike [`collect_module_items`] there is no conditional branch to
/// flatten here.
fn collect_locals(body: &[Stmt], locals: &mut HashMap<String, Kind>) {
    for stmt in body {
        match stmt {
            Stmt::Dim { vars, .. } => insert_var_decls(vars, locals),
            Stmt::Const { vars, .. } => insert_var_decls(vars, locals),
            // A plain `ReDim arr(1 To 5)` **declares** `arr` when nothing
            // else did -- measured, it compiles with no `Dim` in sight.
            // `ReDim Preserve` does not: it needs an array already there to
            // preserve, and Excel rejects it outright on an unknown name
            // (`fuzz/vba_compile_probe.py --only redim`). So only the
            // non-`Preserve` form contributes a symbol; the `Preserve` form
            // is *checked* instead, in `check_stmt`.
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

/// Names VBA creates implicitly, with no `Dim` at all, when `Option
/// Explicit` is off: the target of a plain (non-`Set`) assignment, and a
/// `For` loop's counter. Both are unambiguously plain scalars -- a `Set`
/// target holds an object reference and is left `Opaque` by omission here,
/// and a `For Each` element variable is left alone for the same reason,
/// since either could legitimately be an object with a default member.
///
/// This is a second, separate walk (rather than folded into
/// [`collect_locals`]) so an explicit `Dim`/`Const`/parameter -- collected
/// first -- always wins regardless of where in the procedure text an
/// assignment to the same name happens to sit; VBA's own scoping does not
/// care about textual order either. It is exactly this shape --
/// `x = ws.Range("A1").Value` earlier in a procedure, then a bare `x i`
/// later -- that accounts for most of the false negatives issue #78
/// measured: real VBA scoping makes a procedure-local shadow anything
/// external unconditionally, so treating it as a known plain scalar here
/// carries the same safety argument as the explicit-`Dim` case, not a new
/// risk of a false positive.
fn collect_implicit_locals(body: &[Stmt], locals: &mut HashMap<String, Kind>) {
    for stmt in body {
        match stmt {
            Stmt::Assign {
                target: Expr::Ident { name, .. },
                set,
                ..
            } => {
                // A plain assignment makes a plain scalar; `Set` makes an
                // object reference, which may have a default member and so
                // is `Opaque` -- it exists, but nothing here can say it is
                // uncallable.
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
            // `For Each c In rng` introduces `c`. It is whatever the
            // collection yields -- very often an object -- so `Opaque`.
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

/// VBA's "Duplicate declaration in current scope", for the two ways into it
/// that have been measured.
///
/// Unlike everything else here this pass is **order-sensitive**, and has to
/// be: `Dim x As Long` followed by `x = 1` is ordinary code, while the same
/// two lines the other way round is a compile error. Measured
/// (`fuzz/vba_compile_probe.py --only dup:`):
///
/// ```text
/// Dim x As Long : Dim x As Long   ' FAILS -- declared twice
/// x = 1         : Dim x As Long   ' FAILS -- assigning first creates it
/// x = Helper(1) : Dim x As Long   ' FAILS -- likewise
/// Dim x As Long : x = 1           ' compiles
/// ```
///
/// So a name enters procedure scope either by being declared or by being
/// assigned to, and declaring one that is already there is the error.
///
/// Only those two entry routes are treated as bringing a name into scope,
/// because only those two were measured. A `For` counter, a `For Each`
/// element, a `Set` target, a plain `ReDim` and a parameter all plausibly
/// count as well -- VBA very likely rejects `Sub F(x)` plus `Dim x` -- but
/// the harness splices its snippet into a parameterless `Sub`, so the
/// parameter case in particular could not be put to Excel. Leaving them out
/// under-reports, which is the safe direction; guessing them in would risk
/// rejecting working code.
///
/// The walk is flat because VBA scoping is: a `Dim` inside an `If` is
/// procedure-scoped, not block-scoped, so two of them in opposite branches
/// still collide. Statements are visited in source order, nested bodies
/// included, which is what makes the ordering rule fall out.
fn check_duplicate_declarations(body: &[Stmt]) -> Result<(), ParseError> {
    let mut in_scope: std::collections::HashSet<String> = std::collections::HashSet::new();
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
                set: false,
                ..
            } => {
                in_scope.insert(norm(name));
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
            Stmt::For { body, .. }
            | Stmt::ForEach { body, .. }
            | Stmt::DoLoop { body, .. }
            | Stmt::With { body, .. } => walk_declarations(body, in_scope)?,
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
        // A bare identifier standing as a whole statement is still a call --
        // VBA reads `x` as "call x, no arguments" -- so it resolves by the
        // same rule as `x 5` does. Measured: `x` alone is rejected whether
        // `x` is undeclared, `Dim`'d as a Long, or created by assignment,
        // while a declared Sub (`Helper`) and a built-in (`Beep`) are
        // accepted (`fuzz/vba_compile_probe.py --only bare:`).
        //
        // Statement position only. A bare `x` *inside* an expression is an
        // ordinary implicit-Variant read and stays unchecked -- also
        // measured, `x = a + b` with nothing declared compiles.
        Stmt::Call {
            expr: Expr::Ident { name, pos },
            ..
        } => check_call_name(name, *pos, ctx),
        Stmt::Call { expr, .. } => check_expr(expr, ctx),
        Stmt::Assign { target, value, .. } => {
            check_expr(target, ctx)?;
            check_expr(value, ctx)
        }
        // `ReDim Preserve arr(...)` needs `arr` to already exist -- measured;
        // the plain form declares it instead (see `collect_locals`). The
        // bounds are ordinary expressions either way.
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
                check_block(b, ctx)?;
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

/// Walks an expression, checking every **call target** in it.
///
/// A bare `Expr::Ident` on its own is deliberately never checked: an
/// undeclared name with no call syntax is a legal implicit Variant, which
/// real Excel compiles happily (measured -- `x = a _ + b` with neither `a`
/// nor `b` declared is accepted). Only `name(...)` and the bare-argument
/// statement form force resolution.
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

/// The one rule, applied to a name used with call syntax.
fn check_call_name(name: &str, pos: super::lexer::Pos, ctx: &Ctx<'_>) -> Result<(), ParseError> {
    let lower = norm(name);
    match ctx.kind_of(&lower) {
        // Declared, and definitely not callable or indexable.
        Some(Kind::PlainScalar) => Err(ParseError {
            message: format!("Sub or Function not defined: {name}"),
            pos,
        }),
        Some(_) => Ok(()),
        // Nowhere in this module. Only a scope that can see the whole
        // project may conclude anything from that -- otherwise the name may
        // perfectly well live in a module this pass was not given.
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

    /// The self-contained scope: `src` is the whole project, so an
    /// unresolvable name is an error. What `check_syntax` does.
    fn check(src: &str) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        let empty = HashSet::new();
        check_module(&module, &Scope::self_contained(&empty)).map_err(|e| e.message)
    }

    /// The partial scope: other modules exist but were not supplied, so an
    /// unresolvable name must be accepted. What `VbaModule::check_syntax`
    /// does.
    fn check_partial(src: &str) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        let empty = HashSet::new();
        check_module(&module, &Scope::partial(&empty)).map_err(|e| e.message)
    }

    /// A self-contained scope that additionally knows `names` from siblings.
    fn check_with_external(src: &str, names: &[&str]) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        let external: HashSet<String> = names.iter().map(|n| norm(n)).collect();
        let scope = Scope {
            external: &external,
            complete_project: true,
        };
        check_module(&module, &scope).map_err(|e| e.message)
    }

    // `Dim x : x y` -- the exact case measured against real Excel in issue
    // #78: a locally declared plain-scalar local used as a call target.
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
        // An untyped `Dim` defaults to `Variant`, and Variant is still
        // rejected here -- measured directly, not a guess.
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

    // `MsgBox "hi"` -- a real VBA intrinsic this module never declares. It
    // resolves through the built-in registry, not the module's own text.
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
        // Could have a default member -- this pass cannot resolve
        // user-defined class shapes, so it stays silent rather than guess.
        let src = "Sub Test()\n    Dim obj As Collection\n    obj 3\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn rejects_a_bare_identifier_statement_that_is_not_callable() {
        // This shape was left unchecked at first as unmeasured; a later fuzz
        // run found Excel rejecting it, and the probe then settled every
        // variant. All three of these are compile errors in real Excel.
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
        // The lexer folds `$` into the identifier's spelling, so a lookup
        // that does not strip it fails to find `Trim`. That rejected every
        // `$` string intrinsic there is; found with hand-written VBA, since
        // the generated fuzz grammar never emits one.
        assert!(check("Sub Test()\n    x = Trim$(\" a \")\nEnd Sub\n").is_ok());
        assert!(check("Sub Test()\n    x = Left$(\"ab\", 1)\nEnd Sub\n").is_ok());
        // It has to be stripped on the *declaring* side too, or a name
        // written one way and used the other stops matching.
        assert!(check("Sub Test()\n    Dim s$\n    s = Trim(s$)\nEnd Sub\n").is_ok());
        let src =
            "Function F$(a%)\n    F = CStr(a)\nEnd Function\n\nSub T()\n    x = F(1)\nEnd Sub\n";
        assert!(check(src).is_ok());
        // ...and a suffixed plain scalar is still not callable.
        let err = check("Sub Test()\n    Dim s$\n    s$ 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined"), "{err}");
    }

    #[test]
    fn rejects_a_duplicate_declaration() {
        // All measured. The last is the shape every remaining false negative
        // on fuzz seed 91177 turned out to be.
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
    fn declaring_before_assigning_is_ordinary_code() {
        // The other order is what every well-written procedure does, and it
        // must stay legal -- measured accept. This is why the pass has to be
        // order-sensitive rather than just counting names.
        assert!(check("Sub T()\n    Dim x As Long\n    x = 1\nEnd Sub\n").is_ok());
        // Nested bodies are walked in source order too.
        let src =
            "Sub T()\n    Dim x As Long\n    If True Then\n        x = 1\n    End If\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn a_duplicate_is_caught_across_block_boundaries() {
        // VBA has no block scope, so two `Dim`s in opposite branches of one
        // `If` still collide.
        let src = "Sub T()\n    If True Then\n        Dim x As Long\n    Else\n        Dim x As Long\n    End If\nEnd Sub\n";
        assert!(check(src).unwrap_err().contains("Duplicate declaration"));
        // ...and an assignment buried in a loop body counts as creating the
        // name, which is fuzz iter_11's shape.
        let src = "Sub T()\n    Do While x < 10\n        x = x + 1\n    Loop\n    Dim x As Long\nEnd Sub\n";
        assert!(check(src).unwrap_err().contains("Duplicate declaration"));
    }

    #[test]
    fn a_module_level_name_is_not_a_duplicate_of_a_local() {
        // Procedure scope shadows module scope in VBA; only same-scope
        // redeclaration is the error.
        let src = "Dim x As Long\n\nSub T()\n    Dim x As Long\nEnd Sub\n";
        assert!(check(src).is_ok());
        // And two procedures may each declare the same local name.
        let src = "Sub A()\n    Dim x As Long\nEnd Sub\n\nSub B()\n    Dim x As Long\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn a_bare_name_inside_an_expression_stays_unchecked() {
        // The counterpart to the rule above: only *statement* position is a
        // call. Measured -- `x = a + b` with nothing declared compiles.
        assert!(check("Sub Test()\n    x = a + b\nEnd Sub\n").is_ok());
    }

    #[test]
    fn accepts_call_keyword_to_declared_procedure() {
        let src = "Sub Test()\n    Call Foo(5)\nEnd Sub\n\nSub Foo(n As Long)\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn rejects_local_shadowing_a_module_procedure_name() {
        // A local `Dim` with the same name as a real procedure shadows it,
        // and using the local as a call target is still an error.
        let src = "Sub Foo()\nEnd Sub\n\nSub Test()\n    Dim Bar As Long\n    Bar 5\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: Bar"), "{err}");
    }

    #[test]
    fn rejects_a_name_implicitly_declared_by_plain_assignment() {
        // No `Dim` anywhere -- `x` becomes a plain-Variant local purely by
        // being assigned to, which is how VBA creates one when `Option
        // Explicit` is off. Minimized from
        // fuzz_results/failures/vba_parse_iter_12/source.bas.
        let src = "Sub Test()\n    x = 1\n    x 5\nEnd Sub\n";
        let err = check(src).unwrap_err();
        assert!(err.contains("Sub or Function not defined: x"), "{err}");
    }

    #[test]
    fn rejects_a_name_implicitly_declared_before_its_first_assignment() {
        // Same rule, but the assignment establishing `x` as a plain local
        // comes *after* the offending call textually -- VBA scoping is not
        // sensitive to statement order within a procedure.
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
        // `arr` appears nowhere and is no built-in. Excel refuses to compile
        // this; measured with `fuzz/vba_compile_probe.py --only undeclared`.
        // Minimized from vba_parse_iter_50 and iter_22 respectively.
        let err = check("Sub Test()\n    arr 5\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: arr"), "{err}");
        let err = check("Sub Test()\n    d #1/1/2000#\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: d"), "{err}");
    }

    #[test]
    fn rejects_an_undeclared_call_in_expression_position() {
        // Not a statement -- the call is inside the assigned value.
        // Minimized from vba_parse_iter_40 / 64 / 93.
        let err = check("Sub Test()\n    x = arr(1, 2)\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Sub or Function not defined: arr"), "{err}");
    }

    #[test]
    fn a_partial_scope_never_rejects_an_unresolvable_name() {
        // The same source, checked as one module of a project whose others
        // were not supplied: `arr` may live in a sibling, so it must be
        // accepted. This is the false-positive guard the whole design turns
        // on.
        assert!(check_partial("Sub Test()\n    arr 5\nEnd Sub\n").is_ok());
        assert!(check_partial("Sub Test()\n    x = arr(1, 2)\nEnd Sub\n").is_ok());
        // ...but a name that *is* declared here, as a plain scalar, is still
        // rejected -- a sibling cannot make a local Long callable.
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
        // Measured: `ReDim Preserve arr(1 To 5)` with no `Dim` is a compile
        // error, while the plain form below declares the array and is fine
        // (`fuzz/vba_compile_probe.py --only redim`). Minimized from
        // vba_parse_iter_14 and iter_57.
        let err = check("Sub Test()\n    ReDim Preserve arr(1 To 5)\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Variable not defined: arr"), "{err}");
    }

    #[test]
    fn accepts_plain_redim_which_declares_the_array() {
        assert!(check("Sub Test()\n    ReDim arr(1 To 5)\nEnd Sub\n").is_ok());
        // ...and having declared it, a later `Preserve` is fine.
        let src = "Sub Test()\n    ReDim arr(1 To 5)\n    ReDim Preserve arr(1 To 9)\nEnd Sub\n";
        assert!(check(src).is_ok());
        // As is `Preserve` after an explicit `Dim`.
        let src = "Sub Test()\n    Dim arr()\n    ReDim Preserve arr(1 To 5)\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_an_undeclared_bare_name_with_no_call_syntax() {
        // Measured accept: with `Option Explicit` off these are implicit
        // Variants. Only *call syntax* forces resolution.
        assert!(check("Sub Test()\n    x = a + b\nEnd Sub\n").is_ok());
        let src = "Sub Test()\n    For Each c In rng\n        x = 1\n    Next\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_a_set_assignment_target_as_a_call_target() {
        // `Set` binds an object reference; unlike a plain assignment this
        // does not prove the name is a non-callable scalar, so it must not
        // be treated the same as an implicit `Dim`.
        let src = "Sub Test()\n    Set x = Nothing\n    x 5\nEnd Sub\n";
        assert!(check(src).is_ok());
    }

    #[test]
    fn accepts_a_for_each_element_variable_as_a_call_target() {
        // Could legitimately be an object with a default member -- left
        // alone for the same reason a `Set` target is.
        let src = "Sub Test()\n    For Each c In rng\n        c 5\n    Next c\nEnd Sub\n";
        assert!(check(src).is_ok());
    }
}
