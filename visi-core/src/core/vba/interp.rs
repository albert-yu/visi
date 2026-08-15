//! A tree-walking interpreter for the VBA subset in Phase 1 of
//! `docs/vba-macro-support.md`.
//!
//! Scope is expressions, control flow, `Sub`/`Function` calls, and
//! `On Error`. There is **no host object model**: `Range`, `Worksheets`,
//! `ThisWorkbook` and everything else that touches a workbook belong to
//! Phase 2 and raise [`VbaError`] 438 here rather than silently doing
//! nothing. That refusal is deliberate — a macro that skips a line it does
//! not understand and then reports success has produced a wrong answer in
//! the most dangerous way available.
//!
//! Two structural notes:
//!
//! **`On Error Resume Next` is handled where the statement fails, not at the
//! procedure level.** [`Interpreter::exec_block`] catches the error from each
//! statement it runs, so resumption continues with the next statement *in
//! that block* — inside the loop body, if that is where the failure was.
//! Handling it only at the top would resume in the wrong place for anything
//! nested, which is most real error handling.
//!
//! **Every loop and every call is bounded.** `max_ops` caps total statement
//! executions and `max_depth` caps recursion, because this runs on source the
//! user did not necessarily write, and `Do While True` is one keystroke away.

use std::collections::HashMap;
use std::rc::Rc;

use super::ast::*;
use super::builtins;
use super::value::{self, ArithMode, VResult, Variant, VbaError};

/// How many statements a single `run` may execute before giving up.
const DEFAULT_MAX_OPS: u64 = 5_000_000;
/// How deep procedure calls may nest.
///
/// Well below VBA's own limit, and deliberately so: each VBA frame costs
/// several Rust frames, and an unbounded-recursion test overflowed the real
/// stack at 256 before this guard could fire. A guard that aborts the process
/// instead of returning an error is not a guard.
const DEFAULT_MAX_DEPTH: usize = 64;

/// Error 438 — the "object doesn't support this property or method" that
/// everything outside Phase 1's scope reports.
fn out_of_scope(what: &str) -> VbaError {
    VbaError::new(
        438,
        format!(
            "Object doesn't support this property or method: {what} is not available yet (Phase 1 has no host object model)"
        ),
    )
}

/// Non-local control flow out of a statement.
#[derive(Debug, Clone, PartialEq)]
enum Flow {
    /// Fall through to the next statement.
    Normal,
    /// `Exit Sub` / `Exit Function` / `Exit Property`.
    ExitProc,
    /// `Exit For`.
    ExitFor,
    /// `Exit Do` (and `Exit While`).
    ExitDo,
    /// `GoTo`, or a jump into an error handler. Unwinds to the procedure
    /// body, where labels live.
    Goto(String),
}

/// What `On Error` is currently set to.
#[derive(Debug, Clone, PartialEq)]
enum Handler {
    /// No handler: an error propagates out of the procedure.
    None,
    /// `On Error Resume Next`.
    ResumeNext,
    /// `On Error GoTo <label>`.
    Goto(String),
}

/// One procedure activation.
struct Frame {
    locals: HashMap<String, Variant>,
    handler: Handler,
    /// Whether we are running inside an error handler, during which VBA
    /// disables the active handler so a second error propagates instead of
    /// looping back into the same handler forever.
    in_handler: bool,
    /// The index, in the procedure body, of the top-level statement that
    /// raised the error a handler is currently dealing with. `Resume Next`
    /// continues after it.
    failed_at: Option<usize>,
}

impl Frame {
    fn new() -> Self {
        Self {
            locals: HashMap::new(),
            handler: Handler::None,
            in_handler: false,
            failed_at: None,
        }
    }
}

/// The state `Err` exposes.
#[derive(Debug, Clone, Default)]
struct ErrState {
    number: i32,
    description: String,
}

/// Runs VBA procedures from a parsed [`Module`].
pub struct Interpreter {
    module: Module,
    /// Procedures indexed by lowercased name. Behind an `Rc` so a call can
    /// hold one while `&mut self` runs its body, without cloning the body.
    procs: HashMap<String, Rc<Procedure>>,
    /// Module-level variables, keyed by lowercased name (VBA is
    /// case-insensitive).
    globals: HashMap<String, Variant>,
    err: ErrState,
    ops: u64,
    max_ops: u64,
    depth: usize,
    max_depth: usize,
}

impl Interpreter {
    /// Builds an interpreter over a parsed module.
    pub fn new(module: Module) -> Self {
        let procs = module
            .procedures()
            .into_iter()
            .map(|p| (p.name.to_ascii_lowercase(), Rc::new(p.clone())))
            .collect();
        Self {
            module,
            procs,
            globals: HashMap::new(),
            err: ErrState::default(),
            ops: 0,
            max_ops: DEFAULT_MAX_OPS,
            depth: 0,
            max_depth: DEFAULT_MAX_DEPTH,
        }
    }

    /// Caps how many statements a run may execute.
    pub fn with_max_ops(mut self, max_ops: u64) -> Self {
        self.max_ops = max_ops;
        self
    }

    /// Runs the named procedure and returns its value (`Empty` for a `Sub`).
    pub fn run(&mut self, name: &str, args: Vec<Variant>) -> VResult<Variant> {
        self.ops = 0;
        self.init_module_level()?;
        self.call_procedure(name, args)
    }

    /// Executes module-level declarations so their initialisers are in scope.
    fn init_module_level(&mut self) -> VResult<()> {
        let items = std::mem::take(&mut self.module.items);
        for item in &items {
            if let ModuleItem::Declaration(stmt) = item {
                let mut frame = Frame::new();
                // Module-level declarations write to globals, so run them
                // against a throwaway frame and lift the results.
                let r = self.exec_stmt(stmt, &mut frame, true);
                for (k, v) in frame.locals {
                    self.globals.insert(k, v);
                }
                r?;
            }
        }
        self.module.items = items;
        Ok(())
    }

    fn find_procedure(&self, name: &str) -> Option<Rc<Procedure>> {
        self.procs.get(&name.to_ascii_lowercase()).cloned()
    }

    fn call_procedure(&mut self, name: &str, args: Vec<Variant>) -> VResult<Variant> {
        let Some(proc) = self.find_procedure(name) else {
            return Err(VbaError::new(
                35,
                format!("Sub or Function not defined: {name}"),
            ));
        };
        self.depth += 1;
        if self.depth > self.max_depth {
            self.depth -= 1;
            return Err(VbaError::new(28, "Out of stack space"));
        }
        let result = self.call_body(&proc, args);
        self.depth -= 1;
        result
    }

    fn call_body(&mut self, proc: &Procedure, args: Vec<Variant>) -> VResult<Variant> {
        let mut frame = Frame::new();
        for (i, param) in proc.params.iter().enumerate() {
            let value = args.get(i).cloned().unwrap_or(Variant::Empty);
            frame.locals.insert(param.name.to_ascii_lowercase(), value);
        }
        // A Function returns by assigning to its own name, so seed a slot.
        let ret_key = proc.name.to_ascii_lowercase();
        if proc.kind != ProcKind::Sub {
            frame
                .locals
                .entry(ret_key.clone())
                .or_insert(Variant::Empty);
        }

        self.exec_procedure_body(&proc.body, &mut frame)?;

        Ok(if proc.kind == ProcKind::Sub {
            Variant::Empty
        } else {
            frame
                .locals
                .get(&ret_key)
                .cloned()
                .unwrap_or(Variant::Empty)
        })
    }

    /// Runs a procedure body, resolving `GoTo` against its top-level labels.
    ///
    /// Labels live at procedure level, so a jump out of a nested block
    /// unwinds to here as [`Flow::Goto`] and resumes at the label's index.
    fn exec_procedure_body(&mut self, body: &[Stmt], frame: &mut Frame) -> VResult<()> {
        let mut pc = 0usize;
        while pc < body.len() {
            let flow = match self.exec_stmt(&body[pc], frame, false) {
                Ok(f) => f,
                Err(e) => {
                    frame.failed_at = Some(pc);
                    match self.take_handler(frame) {
                        Handler::ResumeNext => {
                            self.set_err(&e);
                            pc += 1;
                            continue;
                        }
                        Handler::Goto(label) => {
                            self.set_err(&e);
                            frame.in_handler = true;
                            Flow::Goto(label)
                        }
                        Handler::None => return Err(e),
                    }
                }
            };
            match flow {
                Flow::Normal => pc += 1,
                Flow::ExitProc => return Ok(()),
                // `Exit For`/`Exit Do` outside a loop is a no-op rather than
                // an error, matching how VBA compiles it.
                Flow::ExitFor | Flow::ExitDo => pc += 1,
                Flow::Goto(label) => {
                    if label == "\0resume-next" {
                        pc = frame.failed_at.map(|i| i + 1).unwrap_or(pc + 1);
                        frame.in_handler = false;
                        continue;
                    }
                    match Self::find_label(body, &label) {
                        Some(i) => pc = i,
                        None => {
                            return Err(VbaError::new(
                                erl_label_error(),
                                format!("Label not defined: {label}"),
                            ));
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn find_label(body: &[Stmt], label: &str) -> Option<usize> {
        body.iter()
            .position(|s| matches!(s, Stmt::Label { name, .. } if name.eq_ignore_ascii_case(label)))
    }

    /// The handler to apply to an error, honouring VBA's rule that a handler
    /// is disabled while it is running.
    fn take_handler(&self, frame: &Frame) -> Handler {
        if frame.in_handler {
            Handler::None
        } else {
            frame.handler.clone()
        }
    }

    fn set_err(&mut self, e: &VbaError) {
        self.err = ErrState {
            number: e.number,
            description: e.description.clone(),
        };
    }

    /// Runs a block, applying `On Error Resume Next` at the point of failure.
    fn exec_block(&mut self, body: &[Stmt], frame: &mut Frame) -> VResult<Flow> {
        for stmt in body {
            match self.exec_stmt(stmt, frame, false) {
                Ok(Flow::Normal) => {}
                Ok(other) => return Ok(other),
                Err(e) => match self.take_handler(frame) {
                    // Resume where the failure happened, which for a nested
                    // statement means the next statement in *this* block.
                    Handler::ResumeNext => {
                        self.set_err(&e);
                        continue;
                    }
                    Handler::Goto(label) => {
                        self.set_err(&e);
                        frame.in_handler = true;
                        return Ok(Flow::Goto(label));
                    }
                    Handler::None => return Err(e),
                },
            }
        }
        Ok(Flow::Normal)
    }

    fn tick(&mut self) -> VResult<()> {
        self.ops += 1;
        if self.ops > self.max_ops {
            return Err(VbaError::new(
                16,
                "Expression too complex: statement limit exceeded (possible infinite loop)",
            ));
        }
        Ok(())
    }

    fn exec_stmt(&mut self, stmt: &Stmt, frame: &mut Frame, module_level: bool) -> VResult<Flow> {
        self.tick()?;
        match stmt {
            Stmt::Label { .. } => Ok(Flow::Normal),

            Stmt::Dim { vars, .. } => {
                for v in vars {
                    let initial = default_for(v.ty.as_ref());
                    frame.locals.insert(v.name.to_ascii_lowercase(), initial);
                }
                Ok(Flow::Normal)
            }

            Stmt::Const { vars, .. } => {
                for v in vars {
                    let value = match &v.value {
                        Some(e) => self.eval(e, frame)?,
                        None => Variant::Empty,
                    };
                    frame.locals.insert(v.name.to_ascii_lowercase(), value);
                }
                Ok(Flow::Normal)
            }

            Stmt::Assign { target, value, .. } => {
                let v = self.eval(value, frame)?;
                self.assign(target, v, frame, module_level)?;
                Ok(Flow::Normal)
            }

            Stmt::Call { expr, .. } => {
                self.eval(expr, frame)?;
                Ok(Flow::Normal)
            }

            Stmt::If {
                branches,
                else_body,
                ..
            } => {
                for (cond, body) in branches {
                    let c = self.eval(cond, frame)?;
                    // A Null condition is false, not an error.
                    if !c.is_null() && c.to_bool()? {
                        return self.exec_block(body, frame);
                    }
                }
                match else_body {
                    Some(body) => self.exec_block(body, frame),
                    None => Ok(Flow::Normal),
                }
            }

            Stmt::SelectCase {
                subject,
                cases,
                case_else,
                ..
            } => {
                let subject = self.eval(subject, frame)?;
                for case in cases {
                    for m in &case.matches {
                        if self.case_matches(&subject, m, frame)? {
                            return self.exec_block(&case.body, frame);
                        }
                    }
                }
                match case_else {
                    Some(body) => self.exec_block(body, frame),
                    None => Ok(Flow::Normal),
                }
            }

            Stmt::For {
                var,
                from,
                to,
                step,
                body,
                ..
            } => self.exec_for(var, from, to, step.as_ref(), body, frame),

            Stmt::ForEach { .. } => Err(out_of_scope("For Each")),

            Stmt::DoLoop {
                pre, post, body, ..
            } => self.exec_do(pre.as_ref(), post.as_ref(), body, frame),

            Stmt::With { .. } => Err(out_of_scope("With")),

            Stmt::Exit { kind, .. } => Ok(match kind {
                ExitKind::Sub | ExitKind::Function | ExitKind::Property => Flow::ExitProc,
                ExitKind::For => Flow::ExitFor,
                ExitKind::Do | ExitKind::While => Flow::ExitDo,
            }),

            Stmt::GoTo { label, .. } => Ok(Flow::Goto(label.clone())),

            Stmt::OnError { kind, .. } => {
                frame.handler = match kind {
                    OnErrorKind::GoTo(label) => Handler::Goto(label.clone()),
                    OnErrorKind::ResumeNext => Handler::ResumeNext,
                    OnErrorKind::Disable => Handler::None,
                };
                // Re-arming the handler leaves the handler context.
                frame.in_handler = false;
                Ok(Flow::Normal)
            }

            Stmt::Resume { kind, .. } => {
                frame.in_handler = false;
                Ok(match kind {
                    ResumeKind::Label(label) => Flow::Goto(label.clone()),
                    // Sentinel the procedure loop turns into "the statement
                    // after the one that failed".
                    ResumeKind::Next => Flow::Goto("\0resume-next".to_string()),
                    ResumeKind::Retry => Flow::Goto("\0resume-next".to_string()),
                })
            }

            Stmt::Stop { .. } | Stmt::End { .. } => Ok(Flow::ExitProc),

            // Everything below is out of Phase 1's scope. Each reports what
            // it was rather than being skipped.
            Stmt::ReDim { .. } => Err(out_of_scope("ReDim")),
            Stmt::Erase { .. } => Err(out_of_scope("Erase")),
            Stmt::GoSub { .. } | Stmt::Return { .. } => Err(out_of_scope("GoSub")),
            Stmt::OnGoto { .. } => Err(out_of_scope("On ... GoTo")),
            Stmt::TypeDef { .. } => Err(out_of_scope("Type")),
            Stmt::EnumDef { .. } => Err(out_of_scope("Enum")),
            Stmt::Declare { .. } => Err(out_of_scope("Declare")),
            Stmt::EventDef { .. } | Stmt::RaiseEvent { .. } => Err(out_of_scope("events")),
            Stmt::Implements { .. } => Err(out_of_scope("Implements")),
            Stmt::Opaque { keyword, .. } => Err(out_of_scope(keyword)),
        }
    }

    fn case_matches(
        &mut self,
        subject: &Variant,
        m: &CaseMatch,
        frame: &mut Frame,
    ) -> VResult<bool> {
        Ok(match m {
            CaseMatch::Value(e) => {
                let v = self.eval(e, frame)?;
                value::compare(subject, &v)? == Some(std::cmp::Ordering::Equal)
            }
            CaseMatch::Range(lo, hi) => {
                let lo = self.eval(lo, frame)?;
                let hi = self.eval(hi, frame)?;
                let a = value::compare(subject, &lo)?;
                let b = value::compare(subject, &hi)?;
                matches!(a, Some(o) if o != std::cmp::Ordering::Less)
                    && matches!(b, Some(o) if o != std::cmp::Ordering::Greater)
            }
            CaseMatch::Is(op, e) => {
                let v = self.eval(e, frame)?;
                let ord = value::compare(subject, &v)?;
                match ord {
                    None => false,
                    Some(o) => compare_with(*op, o),
                }
            }
        })
    }

    fn exec_for(
        &mut self,
        var: &Expr,
        from: &Expr,
        to: &Expr,
        step: Option<&Expr>,
        body: &[Stmt],
        frame: &mut Frame,
    ) -> VResult<Flow> {
        let start = self.eval(from, frame)?.to_f64()?;
        let limit = self.eval(to, frame)?.to_f64()?;
        let step_v = match step {
            Some(e) => self.eval(e, frame)?.to_f64()?,
            None => 1.0,
        };
        // A zero step would spin forever; VBA runs it as an infinite loop,
        // which the op budget would eventually catch, but failing fast is
        // more useful than burning five million ops first.
        if step_v == 0.0 {
            return Err(VbaError::new(
                5,
                "Invalid procedure call or argument: For step is 0",
            ));
        }

        let mut current = start;
        loop {
            self.tick()?;
            let done = if step_v > 0.0 {
                current > limit
            } else {
                current < limit
            };
            if done {
                break;
            }
            self.assign(var, number_like(current, start, step_v), frame, false)?;
            match self.exec_block(body, frame)? {
                Flow::Normal => {}
                Flow::ExitFor => break,
                other => return Ok(other),
            }
            current += step_v;
        }
        Ok(Flow::Normal)
    }

    fn exec_do(
        &mut self,
        pre: Option<&(DoTest, Expr)>,
        post: Option<&(DoTest, Expr)>,
        body: &[Stmt],
        frame: &mut Frame,
    ) -> VResult<Flow> {
        loop {
            self.tick()?;
            if let Some((test, cond)) = pre {
                let c = self.eval(cond, frame)?.to_bool()?;
                let go = match test {
                    DoTest::While => c,
                    DoTest::Until => !c,
                };
                if !go {
                    break;
                }
            }
            match self.exec_block(body, frame)? {
                Flow::Normal => {}
                Flow::ExitDo => break,
                other => return Ok(other),
            }
            if let Some((test, cond)) = post {
                let c = self.eval(cond, frame)?.to_bool()?;
                let go = match test {
                    DoTest::While => c,
                    DoTest::Until => !c,
                };
                if !go {
                    break;
                }
            }
        }
        Ok(Flow::Normal)
    }

    fn assign(
        &mut self,
        target: &Expr,
        v: Variant,
        frame: &mut Frame,
        module_level: bool,
    ) -> VResult<()> {
        match target {
            Expr::Ident { name, .. } => {
                let key = name.to_ascii_lowercase();
                // A module-level statement writes a global; inside a
                // procedure, a name already local (or not global at all)
                // stays local, and only an existing global is written
                // through -- which is VBA's shadowing rule.
                let writes_global = !module_level
                    && !frame.locals.contains_key(&key)
                    && self.globals.contains_key(&key);
                if writes_global {
                    self.globals.insert(key, v);
                } else {
                    frame.locals.insert(key, v);
                }
                Ok(())
            }
            Expr::Member { .. } | Expr::Bang { .. } => Err(out_of_scope("property assignment")),
            Expr::Call { .. } => Err(out_of_scope("array or property assignment")),
            other => Err(VbaError::new(
                erl_assign_error(),
                format!("Cannot assign to this expression ({other:?})"),
            )),
        }
    }

    fn lookup(&self, name: &str, frame: &Frame) -> Option<Variant> {
        let key = name.to_ascii_lowercase();
        frame
            .locals
            .get(&key)
            .or_else(|| self.globals.get(&key))
            .cloned()
    }

    fn eval(&mut self, e: &Expr, frame: &mut Frame) -> VResult<Variant> {
        self.tick()?;
        match e {
            Expr::Literal(l) => Ok(literal_to_variant(l)),

            Expr::Paren { expr, .. } => self.eval(expr, frame),

            Expr::Ident { name, .. } => {
                if let Some(v) = self.lookup(name, frame) {
                    return Ok(v);
                }
                if let Some(v) = self.builtin_constant(name) {
                    return Ok(v);
                }
                // A zero-argument call written without parentheses.
                if self.find_procedure(name).is_some() {
                    return self.call_procedure(name, Vec::new());
                }
                if let Some(v) = builtins::call(name, &[])? {
                    return Ok(v);
                }
                // An undeclared name is Empty in VBA without Option Explicit.
                Ok(Variant::Empty)
            }

            Expr::Unary { op, expr, .. } => {
                let v = self.eval(expr, frame)?;
                match op {
                    UnOp::Neg => value::neg(&v),
                    UnOp::Pos => value::pos(&v),
                    UnOp::Not => value::not(&v),
                }
            }

            Expr::Binary { op, lhs, rhs, .. } => {
                let a = self.eval(lhs, frame)?;
                let b = self.eval(rhs, frame)?;
                // Two compile-time constants use fixed-width arithmetic and
                // overflow; anything involving a variable promotes. See
                // `value::ArithMode`.
                let mode = if is_constant(lhs) && is_constant(rhs) {
                    ArithMode::Constant
                } else {
                    ArithMode::Promote
                };
                eval_binary(*op, &a, &b, mode)
            }

            Expr::Call { target, args, .. } => self.eval_call(target, args, frame),

            Expr::Member { name, .. } => {
                // `Err.Number` / `Err.Description` are the one member access
                // Phase 1 supports, because error handling is in scope and
                // reading them is how a handler reports anything.
                if let Expr::Member {
                    target: Some(t), ..
                } = e
                    && let Expr::Ident { name: obj, .. } = t.as_ref()
                    && obj.eq_ignore_ascii_case("err")
                {
                    return Ok(match name.to_ascii_lowercase().as_str() {
                        "number" => Variant::Long(self.err.number),
                        "description" => Variant::Str(self.err.description.clone()),
                        other => return Err(out_of_scope(&format!("Err.{other}"))),
                    });
                }
                Err(out_of_scope(&format!(".{name}")))
            }

            Expr::Bang { name, .. } => Err(out_of_scope(&format!("!{name}"))),
            Expr::Me { .. } => Err(out_of_scope("Me")),
            Expr::New { .. } => Err(out_of_scope("New")),
            Expr::TypeOf { .. } => Err(out_of_scope("TypeOf")),
            Expr::AddressOf { .. } => Err(out_of_scope("AddressOf")),
        }
    }

    fn eval_call(&mut self, target: &Expr, args: &[Arg], frame: &mut Frame) -> VResult<Variant> {
        // `Err.Raise n` -- in scope because raising is half of error handling.
        if let Expr::Member {
            target: Some(obj),
            name,
            ..
        } = target
            && let Expr::Ident { name: o, .. } = obj.as_ref()
            && o.eq_ignore_ascii_case("err")
            && name.eq_ignore_ascii_case("raise")
        {
            let values = self.eval_args(args, frame)?;
            let number = values
                .first()
                .map(|v| v.to_f64())
                .transpose()?
                .unwrap_or(0.0) as i32;
            let description = match values.get(2) {
                Some(v) => v.to_vba_string()?,
                None => describe_error(number),
            };
            return Err(VbaError::new(number, description));
        }

        let Expr::Ident { name, .. } = target else {
            return Err(out_of_scope("this call target"));
        };

        // A local array or variable indexed like a call is out of scope; a
        // user procedure wins over a builtin of the same name, as in VBA.
        if self.find_procedure(name).is_some() {
            let values = self.eval_args(args, frame)?;
            return self.call_procedure(name, values);
        }
        let values = self.eval_args(args, frame)?;
        if let Some(v) = builtins::call(name, &values)? {
            return Ok(v);
        }
        Err(VbaError::new(
            35,
            format!("Sub or Function not defined: {name}"),
        ))
    }

    fn eval_args(&mut self, args: &[Arg], frame: &mut Frame) -> VResult<Vec<Variant>> {
        let mut out = Vec::with_capacity(args.len());
        for a in args {
            match &a.value {
                Some(e) => out.push(self.eval(e, frame)?),
                // An omitted argument arrives as Empty, matching what
                // `IsMissing` would report for an Optional Variant.
                None => out.push(Variant::Empty),
            }
        }
        Ok(out)
    }

    fn builtin_constant(&self, name: &str) -> Option<Variant> {
        Some(match name.to_ascii_lowercase().as_str() {
            "vbnullstring" => Variant::Str(String::new()),
            "vbcrlf" => Variant::Str("\r\n".to_string()),
            "vbcr" => Variant::Str("\r".to_string()),
            "vblf" => Variant::Str("\n".to_string()),
            "vbtab" => Variant::Str("\t".to_string()),
            "vbnewline" => Variant::Str("\n".to_string()),
            "vbobjecterror" => Variant::Long(-2147221504),
            _ => return None,
        })
    }
}

/// VBA reports an undefined label and a bad assignment target as compile
/// errors, which have no `Err.Number`. 13 is the closest runtime analogue and
/// keeps the differential comparison meaningful rather than inventing a
/// number Excel would never produce.
fn erl_label_error() -> i32 {
    13
}
fn erl_assign_error() -> i32 {
    13
}

fn describe_error(number: i32) -> String {
    match number {
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        9 => "Subscript out of range",
        11 => "Division by zero",
        13 => "Type mismatch",
        94 => "Invalid use of Null",
        _ => "Application-defined or object-defined error",
    }
    .to_string()
}

fn compare_with(op: BinOp, ord: std::cmp::Ordering) -> bool {
    use std::cmp::Ordering::*;
    match op {
        BinOp::Eq => ord == Equal,
        BinOp::Ne => ord != Equal,
        BinOp::Lt => ord == Less,
        BinOp::Gt => ord == Greater,
        BinOp::Le => ord != Greater,
        BinOp::Ge => ord != Less,
        _ => false,
    }
}

/// Whether an expression is a compile-time constant, which decides whether
/// arithmetic over it overflows or promotes.
fn is_constant(e: &Expr) -> bool {
    match e {
        Expr::Literal(_) => true,
        Expr::Paren { expr, .. } => is_constant(expr),
        Expr::Unary { expr, .. } => is_constant(expr),
        Expr::Binary { lhs, rhs, .. } => is_constant(lhs) && is_constant(rhs),
        _ => false,
    }
}

fn eval_binary(op: BinOp, a: &Variant, b: &Variant, mode: ArithMode) -> VResult<Variant> {
    use BinOp::*;
    match op {
        Add => value::add(a, b, mode),
        Sub => value::sub(a, b, mode),
        Mul => value::mul(a, b, mode),
        Div => value::div(a, b),
        IntDiv => value::int_div(a, b),
        Mod => value::modulo(a, b),
        Pow => value::pow(a, b),
        Concat => value::concat(a, b),
        Eq | Ne | Lt | Gt | Le | Ge => match value::compare(a, b)? {
            None => Ok(Variant::Null),
            Some(ord) => Ok(Variant::Boolean(compare_with(op, ord))),
        },
        And => value::logical(a, b, |x, y| x & y),
        Or => value::logical(a, b, |x, y| x | y),
        Xor => value::logical(a, b, |x, y| x ^ y),
        Eqv => value::logical(a, b, |x, y| !(x ^ y)),
        Imp => value::logical(a, b, |x, y| !x | y),
        Like => Err(out_of_scope("Like")),
        Is => Err(out_of_scope("Is")),
    }
}

fn literal_to_variant(l: &Literal) -> Variant {
    use super::lexer::TypeSuffix;
    match l {
        Literal::Number {
            value,
            base,
            suffix,
            is_float,
        } => match suffix {
            Some(TypeSuffix::Integer) => Variant::Integer(*value as i16),
            Some(TypeSuffix::Long) => Variant::Long(*value as i32),
            Some(TypeSuffix::Single) => Variant::Single(*value as f32),
            Some(TypeSuffix::Double) => Variant::Double(*value),
            Some(TypeSuffix::Currency) => Variant::Currency((value * 10_000.0).round() as i64),
            Some(TypeSuffix::String) => Variant::Str(value::format_number(*value)),
            None => {
                // A fraction or exponent forces Double, which the lexer
                // records: `1E3` is a Double even though `1000` is a Long.
                let _ = base;
                Variant::from_literal(*value, *is_float || value.fract() != 0.0)
            }
        },
        Literal::Str(s) => Variant::Str(s.clone()),
        Literal::Bool(b) => Variant::Boolean(*b),
        Literal::Empty => Variant::Empty,
        Literal::Null => Variant::Null,
        // Dates and Nothing are out of Phase 1's value scope; a date literal
        // becomes Empty rather than a wrong number.
        Literal::Date(_) | Literal::Nothing => Variant::Empty,
    }
}

/// A `For` counter keeps the type its bounds imply, so `For i = 1 To 3`
/// counts in `Integer`s and `For x = 1.5 To 3` in `Double`s.
fn number_like(current: f64, start: f64, step: f64) -> Variant {
    let integral = current.fract() == 0.0 && start.fract() == 0.0 && step.fract() == 0.0;
    Variant::from_literal(current, !integral)
}

fn default_for(ty: Option<&TypeRef>) -> Variant {
    let Some(ty) = ty else {
        return Variant::Empty;
    };
    let Some(last) = ty.path.last() else {
        return Variant::Empty;
    };
    // A typed variable starts at its type's zero, not Empty -- which is
    // observable, since `Dim s As String` makes `s` `""` rather than Empty.
    match last.to_ascii_lowercase().as_str() {
        "integer" => Variant::Integer(0),
        "long" => Variant::Long(0),
        "single" => Variant::Single(0.0),
        "double" => Variant::Double(0.0),
        "currency" => Variant::Currency(0),
        "boolean" => Variant::Boolean(false),
        "string" => Variant::Str(String::new()),
        _ => Variant::Empty,
    }
}

#[cfg(test)]
mod tests {
    use super::super::parser::parse_module;
    use super::*;

    /// Runs a body inside a Function and reports `TypeName|CStr` -- the same
    /// pair `fuzz/vba_variant_probe.bas` prints from Excel, so a test's
    /// expected string can be pasted straight from a probe run.
    fn run(body: &str) -> String {
        let src = format!("Function F()\n{body}\nEnd Function\n");
        let module = parse_module(&src).unwrap_or_else(|e| panic!("{e}\n{src}"));
        match Interpreter::new(module).run("F", Vec::new()) {
            Ok(v) => format!(
                "{}|{}",
                v.type_name(),
                v.to_vba_string().unwrap_or_default()
            ),
            Err(e) => format!("ERR|{}", e.number),
        }
    }

    fn expr(e: &str) -> String {
        run(&format!("    F = {e}"))
    }

    // ---- expressions ----------------------------------------------------

    #[test]
    fn arithmetic_and_types_match_the_excel_probe() {
        // Each expectation is what `fuzz/vba_variant_probe.bas` returned from
        // Excel 16.112 for the same expression.
        assert_eq!(expr("1 + 1"), "Integer|2");
        assert_eq!(expr("32767 + 1"), "ERR|6");
        assert_eq!(expr("1 / 2"), "Double|0.5");
        assert_eq!(expr("4 / 2"), "Double|2");
        assert_eq!(expr("7 \\ 2"), "Integer|3");
        assert_eq!(expr("-7 \\ 2"), "Integer|-3");
        assert_eq!(expr("7.6 \\ 2"), "Long|4");
        assert_eq!(expr("7 Mod 2"), "Integer|1");
        assert_eq!(expr("-7 Mod 2"), "Integer|-1");
        assert_eq!(expr("7.6 Mod 2"), "Long|0");
        assert_eq!(expr("2 ^ 2"), "Double|4");
        assert_eq!(expr("1.5 + 1"), "Double|2.5");
        assert_eq!(expr("1 / 0"), "ERR|11");
    }

    #[test]
    fn precedence_is_the_one_measured_in_phase_0() {
        // The parser's table, exercised through evaluation.
        assert_eq!(expr("2 ^ 3 ^ 2"), "Double|64");
        assert_eq!(expr("-2 ^ 2"), "Double|-4");
        assert_eq!(expr("2 + 3 & 4"), "String|54");
        assert_eq!(expr("1 = 1 And 1 = 0"), "Boolean|False");
        assert_eq!(expr("Not 1 = 0"), "Boolean|True");
        assert_eq!(expr("2 * 10 \\ 3"), "Integer|6");
        assert_eq!(expr("1 + 7 Mod 3"), "Integer|2");
    }

    #[test]
    fn string_coercion_matches_the_probe() {
        assert_eq!(expr("\"1\" + 1"), "Double|2");
        assert_eq!(expr("\"1\" + \"2\""), "String|12");
        assert_eq!(expr("\"abc\" + 1"), "ERR|13");
        assert_eq!(expr("1 & 2"), "String|12");
        assert_eq!(expr("\"  3  \" + 1"), "Double|4");
    }

    #[test]
    fn booleans_and_bitwise_operators_match_the_probe() {
        assert_eq!(expr("True + 1"), "Integer|0");
        assert_eq!(expr("True + True"), "Integer|-2");
        assert_eq!(expr("True And False"), "Boolean|False");
        assert_eq!(expr("5 And 3"), "Integer|1");
        assert_eq!(expr("Not 5"), "Integer|-6");
        assert_eq!(expr("CInt(True)"), "Integer|-1");
    }

    #[test]
    fn empty_and_null_behave_as_measured() {
        assert_eq!(expr("Empty + 1"), "Integer|1");
        assert_eq!(expr("Empty & \"a\""), "String|a");
        assert_eq!(expr("Null & \"a\""), "String|a");
        assert_eq!(expr("IsNull(Null + 1)"), "Boolean|True");
        assert_eq!(expr("Empty = 0"), "Boolean|True");
        assert_eq!(expr("Empty = \"\""), "Boolean|True");
    }

    #[test]
    fn conversions_use_bankers_rounding() {
        assert_eq!(expr("CLng(0.5)"), "Long|0");
        assert_eq!(expr("CLng(1.5)"), "Long|2");
        assert_eq!(expr("CLng(2.5)"), "Long|2");
        assert_eq!(expr("CLng(-1.5)"), "Long|-2");
        assert_eq!(expr("CInt(32768)"), "ERR|6");
        assert_eq!(expr("Int(-1.5)"), "Double|-2");
        assert_eq!(expr("Fix(-1.5)"), "Double|-1");
        assert_eq!(expr("CDbl(\"1e3\")"), "Double|1000");
    }

    // ---- control flow ---------------------------------------------------

    #[test]
    fn for_loops_run_and_can_be_exited() {
        assert_eq!(
            run("    Dim t\n    For i = 1 To 5\n        t = t + i * i\n    Next i\n    F = t"),
            "Integer|55"
        );
        assert_eq!(
            run(
                "    Dim t\n    For i = 1 To 10\n        If i > 3 Then Exit For\n        t = t + 1\n    Next i\n    F = t"
            ),
            "Integer|3"
        );
        // A negative step counts down.
        assert_eq!(
            run("    Dim t\n    For i = 5 To 1 Step -1\n        t = t + i\n    Next i\n    F = t"),
            "Integer|15"
        );
        // A loop whose bounds exclude the start never runs.
        assert_eq!(
            run(
                "    Dim t\n    t = 0\n    For i = 5 To 1\n        t = t + 1\n    Next i\n    F = t"
            ),
            "Integer|0"
        );
    }

    #[test]
    fn every_do_form_terminates_correctly() {
        assert_eq!(
            run("    Dim i\n    i = 0\n    Do While i < 5\n        i = i + 1\n    Loop\n    F = i"),
            "Integer|5"
        );
        assert_eq!(
            run(
                "    Dim i\n    i = 0\n    Do Until i >= 5\n        i = i + 1\n    Loop\n    F = i"
            ),
            "Integer|5"
        );
        // A post-tested loop always runs its body at least once.
        assert_eq!(
            run("    Dim i\n    i = 9\n    Do\n        i = i + 1\n    Loop While i < 5\n    F = i"),
            "Integer|10"
        );
        assert_eq!(
            run("    Dim i\n    i = 0\n    While i < 3\n        i = i + 1\n    Wend\n    F = i"),
            "Integer|3"
        );
    }

    #[test]
    fn select_case_covers_values_ranges_and_is() {
        let body = |x: &str| {
            format!(
                "    Dim r\n    Select Case {x}\n    Case 1, 2\n        r = \"a\"\n    \
                 Case 3 To 5\n        r = \"b\"\n    Case Is >= 6\n        r = \"c\"\n    \
                 Case Else\n        r = \"d\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(run(&body("2")), "String|a");
        assert_eq!(run(&body("4")), "String|b");
        assert_eq!(run(&body("9")), "String|c");
        assert_eq!(run(&body("0")), "String|d");
    }

    #[test]
    fn if_elseif_else_picks_one_branch() {
        let body = |x: &str| {
            format!(
                "    Dim r\n    If {x} > 5 Then\n        r = 1\n    ElseIf {x} > 2 Then\n        \
                 r = 2\n    Else\n        r = 3\n    End If\n    F = r"
            )
        };
        assert_eq!(run(&body("9")), "Integer|1");
        assert_eq!(run(&body("4")), "Integer|2");
        assert_eq!(run(&body("1")), "Integer|3");
    }

    #[test]
    fn goto_jumps_to_a_procedure_level_label() {
        assert_eq!(
            run("    Dim t\n    t = 1\n    GoTo Skip\n    t = 99\nSkip:\n    F = t"),
            "Integer|1"
        );
    }

    // ---- procedures -----------------------------------------------------

    #[test]
    fn functions_call_each_other_and_return_by_name() {
        let src = "Function Outer()\n    Outer = Inner(3) + Inner(4)\nEnd Function\n\
                   Function Inner(n)\n    Inner = n * n\nEnd Function\n";
        let m = parse_module(src).unwrap();
        let v = Interpreter::new(m).run("Outer", Vec::new()).unwrap();
        assert_eq!(v, Variant::Integer(25));
    }

    #[test]
    fn recursion_works_and_is_bounded() {
        let src = "Function Fact(n)\n    If n <= 1 Then\n        Fact = 1\n    Else\n        \
                   Fact = n * Fact(n - 1)\n    End If\nEnd Function\n";
        let m = parse_module(src).unwrap();
        let v = Interpreter::new(m)
            .run("Fact", vec![Variant::Integer(5)])
            .unwrap();
        assert_eq!(v, Variant::Integer(120));

        // Unbounded recursion stops rather than blowing the Rust stack.
        let src = "Function Boom()\n    Boom = Boom()\nEnd Function\n";
        let m = parse_module(src).unwrap();
        let e = Interpreter::new(m).run("Boom", Vec::new()).unwrap_err();
        assert_eq!(e.number, 28);
    }

    #[test]
    fn a_sub_returns_empty_and_exits_early() {
        let src = "Sub S()\n    Exit Sub\nEnd Sub\n";
        let m = parse_module(src).unwrap();
        assert_eq!(
            Interpreter::new(m).run("S", Vec::new()).unwrap(),
            Variant::Empty
        );
    }

    #[test]
    fn an_infinite_loop_hits_the_op_budget_instead_of_hanging() {
        let src = "Function F()\n    Do While True\n    Loop\nEnd Function\n";
        let m = parse_module(src).unwrap();
        let e = Interpreter::new(m)
            .with_max_ops(10_000)
            .run("F", Vec::new())
            .unwrap_err();
        assert_eq!(e.number, 16);
    }

    // ---- error handling -------------------------------------------------

    #[test]
    fn on_error_goto_runs_the_handler_and_exposes_err() {
        assert_eq!(
            run(
                "    On Error GoTo Failed\n    F = 1 / 0\n    Exit Function\nFailed:\n    \
                 F = \"ERR|\" & Err.Number"
            ),
            "String|ERR|11"
        );
        assert_eq!(
            run(
                "    On Error GoTo Failed\n    F = CLng(\"nope\")\n    Exit Function\nFailed:\n    \
                 F = Err.Description"
            ),
            "String|Type mismatch"
        );
    }

    #[test]
    fn on_error_resume_next_continues_at_the_failing_statement() {
        assert_eq!(
            run("    Dim t\n    On Error Resume Next\n    t = 1 / 0\n    t = 7\n    F = t"),
            "Integer|7"
        );
    }

    /// The reason `exec_block` handles errors rather than only the procedure
    /// loop: resuming has to continue inside the loop body, not after it.
    #[test]
    fn resume_next_resumes_inside_a_nested_block() {
        assert_eq!(
            run(
                "    Dim t\n    t = 0\n    On Error Resume Next\n    For i = 1 To 3\n        \
                 t = t + 1 / 0\n        t = t + 1\n    Next i\n    F = t"
            ),
            "Integer|3"
        );
    }

    #[test]
    fn on_error_goto_0_disarms_the_handler() {
        let src = "Function F()\n    On Error Resume Next\n    On Error GoTo 0\n    \
                   F = 1 / 0\nEnd Function\n";
        let m = parse_module(src).unwrap();
        assert_eq!(
            Interpreter::new(m).run("F", Vec::new()).unwrap_err().number,
            11
        );
    }

    #[test]
    fn an_error_inside_a_handler_is_not_caught_by_the_same_handler() {
        // Without this, a handler that itself fails loops forever.
        let src = "Function F()\n    On Error GoTo Failed\n    F = 1 / 0\n    Exit Function\n\
                   Failed:\n    F = 1 / 0\nEnd Function\n";
        let m = parse_module(src).unwrap();
        assert_eq!(
            Interpreter::new(m).run("F", Vec::new()).unwrap_err().number,
            11
        );
    }

    #[test]
    fn err_raise_produces_a_catchable_error() {
        assert_eq!(
            run(
                "    On Error GoTo Failed\n    Err.Raise 5\n    Exit Function\nFailed:\n    \
                 F = Err.Number"
            ),
            "Long|5"
        );
    }

    // ---- builtins -------------------------------------------------------

    #[test]
    fn string_builtins_are_one_based_like_vba() {
        assert_eq!(expr("Len(\"abcd\")"), "Long|4");
        assert_eq!(expr("Left(\"abcd\", 2)"), "String|ab");
        assert_eq!(expr("Right(\"abcd\", 2)"), "String|cd");
        assert_eq!(expr("Mid(\"abcd\", 2, 2)"), "String|bc");
        assert_eq!(expr("Mid(\"abcd\", 3)"), "String|cd");
        assert_eq!(expr("InStr(\"abcd\", \"cd\")"), "Long|3");
        assert_eq!(expr("InStr(\"abcd\", \"z\")"), "Long|0");
        assert_eq!(expr("InStr(3, \"abcabc\", \"a\")"), "Long|4");
        assert_eq!(expr("UCase(\"aB\")"), "String|AB");
        assert_eq!(expr("Trim(\"  a  \")"), "String|a");
        assert_eq!(expr("Replace(\"aXbXc\", \"X\", \"-\")"), "String|a-b-c");
        assert_eq!(expr("Chr(65)"), "String|A");
        assert_eq!(expr("Asc(\"A\")"), "Integer|65");
        // Mid is 1-based, so 0 is an error rather than a clamp.
        assert_eq!(expr("Mid(\"abcd\", 0)"), "ERR|5");
    }

    #[test]
    fn inspection_builtins_report_the_subtype() {
        assert_eq!(expr("TypeName(1)"), "String|Integer");
        assert_eq!(expr("TypeName(1.5)"), "String|Double");
        assert_eq!(expr("TypeName(\"a\")"), "String|String");
        assert_eq!(expr("TypeName(True)"), "String|Boolean");
        assert_eq!(expr("TypeName(100000)"), "String|Long");
        assert_eq!(expr("IsNumeric(\"12\")"), "Boolean|True");
        assert_eq!(expr("IsNumeric(\"ab\")"), "Boolean|False");
        assert_eq!(expr("IsEmpty(Empty)"), "Boolean|True");
    }

    #[test]
    fn math_builtins_keep_the_arguments_width() {
        assert_eq!(expr("Abs(-3)"), "Integer|3");
        assert_eq!(expr("Abs(-3.5)"), "Double|3.5");
        assert_eq!(expr("Sgn(-9)"), "Integer|-1");
        assert_eq!(expr("Sqr(9)"), "Double|3");
        assert_eq!(expr("Sqr(-1)"), "ERR|5");
    }

    #[test]
    fn a_typed_dim_starts_at_its_types_zero_not_empty() {
        // Observable: `Dim s As String` makes s "" rather than Empty.
        assert_eq!(
            run("    Dim s As String\n    F = TypeName(s)"),
            "String|String"
        );
        assert_eq!(run("    Dim n As Long\n    F = TypeName(n)"), "String|Long");
        assert_eq!(run("    Dim v\n    F = TypeName(v)"), "String|Empty");
    }

    // ---- out of scope ---------------------------------------------------

    #[test]
    fn host_object_access_errors_rather_than_silently_doing_nothing() {
        // The refusal that matters: a macro that skips a line it cannot
        // understand and then reports success is wrong in the worst way.
        for body in [
            "    F = Range(\"A1\").Value",
            "    F = ThisWorkbook.Name",
            "    With x\n        F = .a\n    End With",
            "    Dim c\n    For Each c In r\n    Next",
        ] {
            let out = run(body);
            assert!(out.starts_with("ERR|438"), "{body:?} gave {out}");
        }
    }

    #[test]
    fn an_unknown_function_is_reported_not_ignored() {
        assert_eq!(expr("NoSuchFunction(1)"), "ERR|35");
    }
}
