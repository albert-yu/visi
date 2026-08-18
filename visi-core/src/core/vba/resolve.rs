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
//! - A name this pass cannot find *anywhere* -- another module's procedure,
//!   a real VBA intrinsic like `MsgBox`, a host object -- is left alone.
//!   `check_syntax` operates one module at a time and does not enumerate
//!   VBA's built-in surface, so "not found" is not evidence of anything.
//!
//! A bare identifier used as a statement with **no** call syntax at all
//! (`x` alone, or `x` as the value of `Stmt::Call { expr: Expr::Ident }`)
//! is deliberately left unchecked -- unlike `x y`, that shape was never
//! measured against Excel, so guessing at its rule risks exactly the false
//! positive this pass exists to avoid.
//!
//! ## What this does and does not fix
//!
//! Re-running `check_syntax` over the 11 saved false-negative reproductions
//! from issue #78's investigation (`fuzz_results/failures/vba_parse_iter_*`)
//! after this pass landed: 3 now correctly fail to parse (a 4th of the 11,
//! the keyword-declaration case, was already fixed separately). All 3 share
//! the same shape -- `Dim` is absent, but the name is unambiguously a plain
//! local because it is the target of a plain assignment (`x = ...`)
//! elsewhere in the *same procedure*, which is how VBA creates a variable
//! implicitly when `Option Explicit` is off. That is exactly the measured
//! `Dim x : x y` rule with the `Dim` replaced by an equivalent implicit one
//! -- real VBA scoping makes a local shadow anything external
//! unconditionally regardless of how the local came to exist, so this
//! carries the same safety argument, not a new risk.
//!
//! The other 7 remain unfixed, and deliberately so: every one of them is a
//! name that is **never** assigned or declared anywhere in the module --
//! `x Function = 1 + 2 * 3` where `x` appears nowhere else, `d #1/1/2000#`
//! where `d` appears nowhere else, `x = arr(1, 2)` where `arr` appears
//! nowhere else. This pass cannot tell "genuinely undeclared, therefore a
//! compile error" apart from "a legitimate call to a real VBA intrinsic or
//! another module's public procedure, which this single-module pass simply
//! has not seen" -- which is precisely the ambiguity the issue's own
//! writeup calls out as needing "its own design pass," not a guess bolted
//! onto this one. Closing that gap needs either a real VBA built-in-name
//! registry or whole-project (not single-module) resolution.
//!
//! [issue #78]: https://github.com/albert-yu/visi/issues/78

use super::ast::{Expr, Module, ModuleItem, Procedure, Stmt, TypeRef, VarDecl};
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

/// Checks every implicit-call statement in `module` against the symbol
/// table this pass can build from its own text, per this module's doc.
pub(super) fn check_implicit_calls(module: &Module) -> Result<(), ParseError> {
    let module_syms = collect_module_symbols(module);
    check_items(&module.items, &module_syms)
}

fn check_items(
    items: &[ModuleItem],
    module_syms: &HashMap<String, Kind>,
) -> Result<(), ParseError> {
    for item in items {
        match item {
            ModuleItem::Procedure(p) => check_procedure(p, module_syms)?,
            ModuleItem::Conditional {
                branches,
                else_items,
                ..
            } => {
                for (_, body) in branches {
                    check_items(body, module_syms)?;
                }
                if let Some(body) = else_items {
                    check_items(body, module_syms)?;
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
                syms.insert(p.name.to_ascii_lowercase(), Kind::Callable);
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
            syms.insert(name.to_ascii_lowercase(), Kind::Callable);
        }
        // `Type`/`Enum`/`Event`/`Implements` don't declare a callable or a
        // plain-scalar name, so they cannot be an implicit-call target this
        // pass would flag either way.
        _ => {}
    }
}

fn insert_var_decls(vars: &[VarDecl], syms: &mut HashMap<String, Kind>) {
    for v in vars {
        syms.insert(
            v.name.to_ascii_lowercase(),
            kind_for_decl(v.bounds.is_some(), &v.ty),
        );
    }
}

fn check_procedure(
    proc: &Procedure,
    module_syms: &HashMap<String, Kind>,
) -> Result<(), ParseError> {
    let mut locals = HashMap::new();
    for param in &proc.params {
        locals.insert(
            param.name.to_ascii_lowercase(),
            kind_for_decl(param.is_array || param.param_array, &param.ty),
        );
    }
    collect_locals(&proc.body, &mut locals);
    collect_implicit_locals(&proc.body, &mut locals);
    check_block(&proc.body, module_syms, &locals)
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
                set: false,
                ..
            } => {
                locals
                    .entry(name.to_ascii_lowercase())
                    .or_insert(Kind::PlainScalar);
            }
            Stmt::For { var, body, .. } => {
                if let Expr::Ident { name, .. } = var {
                    locals
                        .entry(name.to_ascii_lowercase())
                        .or_insert(Kind::PlainScalar);
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
            Stmt::ForEach { body, .. } | Stmt::DoLoop { body, .. } | Stmt::With { body, .. } => {
                collect_implicit_locals(body, locals);
            }
            _ => {}
        }
    }
}

fn check_block(
    body: &[Stmt],
    module_syms: &HashMap<String, Kind>,
    locals: &HashMap<String, Kind>,
) -> Result<(), ParseError> {
    for stmt in body {
        check_stmt(stmt, module_syms, locals)?;
    }
    Ok(())
}

fn check_stmt(
    stmt: &Stmt,
    module_syms: &HashMap<String, Kind>,
    locals: &HashMap<String, Kind>,
) -> Result<(), ParseError> {
    match stmt {
        Stmt::Call { expr, .. } => check_call_target(expr, module_syms, locals),
        Stmt::If {
            branches,
            else_body,
            ..
        } => {
            for (_, b) in branches {
                check_block(b, module_syms, locals)?;
            }
            if let Some(b) = else_body {
                check_block(b, module_syms, locals)?;
            }
            Ok(())
        }
        Stmt::SelectCase {
            cases, case_else, ..
        } => {
            for c in cases {
                check_block(&c.body, module_syms, locals)?;
            }
            if let Some(b) = case_else {
                check_block(b, module_syms, locals)?;
            }
            Ok(())
        }
        Stmt::For { body, .. }
        | Stmt::ForEach { body, .. }
        | Stmt::DoLoop { body, .. }
        | Stmt::With { body, .. } => check_block(body, module_syms, locals),
        _ => Ok(()),
    }
}

/// Only `Stmt::Call { expr: Expr::Call { target: Expr::Ident, .. }, .. }` --
/// explicit call syntax, whether `Foo(x)`, bare `Foo x`, or `Call Foo(x)` --
/// is in scope. A bare `Expr::Ident` alone (no call syntax at all) is left
/// unchecked; see this module's doc for why.
fn check_call_target(
    expr: &Expr,
    module_syms: &HashMap<String, Kind>,
    locals: &HashMap<String, Kind>,
) -> Result<(), ParseError> {
    let Expr::Call { target, .. } = expr else {
        return Ok(());
    };
    let Expr::Ident { name, pos } = target.as_ref() else {
        return Ok(());
    };
    let lower = name.to_ascii_lowercase();
    let kind = locals.get(&lower).or_else(|| module_syms.get(&lower));
    if kind == Some(&Kind::PlainScalar) {
        return Err(ParseError {
            message: format!("Sub or Function not defined: {name}"),
            pos: *pos,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::parser::parse_module;
    use super::check_implicit_calls;

    fn check(src: &str) -> Result<(), String> {
        let module = parse_module(src).expect("should parse");
        check_implicit_calls(&module).map_err(|e| e.message)
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

    // `MsgBox "hi"` -- a real VBA intrinsic this module never declares.
    // Absence from the symbol table must not be treated as evidence of
    // anything, per this module's doc.
    #[test]
    fn accepts_call_to_unknown_name() {
        assert!(check("Sub Test()\n    MsgBox \"hi\"\nEnd Sub\n").is_ok());
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
    fn accepts_bare_identifier_statement_with_no_call_syntax() {
        // Unmeasured shape -- deliberately left unchecked.
        let src = "Sub Test()\n    Dim x As Long\n    x\nEnd Sub\n";
        assert!(check(src).is_ok());
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
    fn accepts_an_undeclared_name_never_assigned_anywhere() {
        // `arr` appears nowhere but here -- genuinely undeclared, not
        // implicitly declared by assignment. Left unchecked: this pass
        // cannot tell "compile error" apart from "a real intrinsic or
        // another module's procedure it has not seen." Minimized from
        // fuzz_results/failures/vba_parse_iter_50/source.bas.
        assert!(check("Sub Test()\n    arr 5\nEnd Sub\n").is_ok());
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
