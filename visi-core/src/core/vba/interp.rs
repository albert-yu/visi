//! A tree-walking interpreter for the VBA subset in Phase 1 of
//! `docs/vba-macro-support.md`.
//!
//! Scope is expressions, control flow, `Sub`/`Function` calls, `On Error`,
//! and — when a [`Host`] is attached — the Phase 2 object model in
//! [`super::host`]. Without a host, `Range`, `Worksheets`, `ThisWorkbook` and
//! everything else that touches a workbook raise [`VbaError`] 438, and so
//! does anything outside the object model's allow-list even with one. That
//! refusal is deliberate — a macro that skips a line it does not understand
//! and then reports success has produced a wrong answer in the most
//! dangerous way available.
//!
//! **A plain `=` reads an object's default member; `Set` does not.**
//! `x = ws.Range("A1")` puts the cell's *value* in `x` while
//! `Set r = ws.Range("A1")` puts the object there, and the parser cannot tell
//! the two apart — it sees the same `Member` expression. So every context
//! that wants a scalar funnels through [`Interpreter::scalar`], and the few
//! that want the object (`Set`, `Is`, `TypeName`, a `With` subject, an
//! argument to a user procedure) deliberately skip it. Getting this wrong
//! does not raise; it silently produces the wrong kind of value.
//!
//! Two more structural notes:
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
use super::host::{Host, ObjRef};
use super::value::{self, ArithMode, Operand, VResult, Variant, VbaError};

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
/// everything outside the implemented scope reports.
fn out_of_scope(what: &str) -> VbaError {
    VbaError::new(
        438,
        format!("Object doesn't support this property or method: {what} is not available"),
    )
}

/// The same refusal, for something that needs a workbook when none is
/// attached.
///
/// Distinct wording from [`out_of_scope`] on purpose: "no workbook is
/// attached" is a fixable mistake by the *caller* (run the macro against a
/// file, not a bare `.bas`), where a plain 438 means the construct is not
/// implemented at all and never will be by trying harder.
fn needs_workbook(what: &str) -> VbaError {
    VbaError::new(
        438,
        format!(
            "Object doesn't support this property or method: {what} needs a workbook, and this run has none"
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
    /// Enclosing `With` subjects, innermost last, against which a
    /// leading-dot member reference resolves.
    ///
    /// Per frame rather than per interpreter because a `With` block does not
    /// reach into a procedure it calls: a bare `.Value` inside the callee is
    /// a compile error in VBA, not a reference to the caller's subject.
    with_stack: Vec<Variant>,
}

impl Frame {
    fn new() -> Self {
        Self {
            locals: HashMap::new(),
            handler: Handler::None,
            in_handler: false,
            failed_at: None,
            with_stack: Vec::new(),
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
///
/// The lifetime is the workbook's: an interpreter with a [`Host`] borrows it
/// mutably for the whole run. `Interpreter::new` alone leaves the parameter
/// free, so a host-free run has no lifetime obligations at all.
pub struct Interpreter<'w> {
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
    /// The workbook, if this run has one. `None` is a real mode, not a
    /// degraded one: `visi macro run` over a bare `.bas` file has no workbook
    /// to offer, and every host construct then reports so.
    host: Option<Host<'w>>,
}

impl<'w> Interpreter<'w> {
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
            host: None,
        }
    }

    /// Binds a workbook, enabling the host object model.
    pub fn with_host(mut self, host: Host<'w>) -> Self {
        self.host = Some(host);
        self
    }

    /// Whether the run changed the workbook, and so whether the caller has
    /// something worth writing back.
    pub fn mutated(&self) -> bool {
        self.host.as_ref().is_some_and(|h| h.mutated())
    }

    /// Settles any outstanding recalculation so a workbook about to be saved
    /// holds what a reader inside the macro would have seen.
    pub fn finish(&mut self) {
        if let Some(h) = self.host.as_mut() {
            h.finish();
        }
    }

    fn host(&mut self, what: &str) -> VResult<&mut Host<'w>> {
        self.host.as_mut().ok_or_else(|| needs_workbook(what))
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

            Stmt::Assign {
                target, value, set, ..
            } => {
                let v = self.eval(value, frame)?;
                // `Set` assigns the reference; a plain `=` reads the object's
                // default member on both sides. See the module doc comment.
                let v = if *set { v } else { self.scalar(v)? };
                self.assign_with(target, v, frame, module_level, *set)?;
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
                // A constant String subject compares as text, even against
                // numeric cases -- `Select Case "32768abc"` takes
                // `Case 2 To 5` because "32768abc" sorts between "2" and "5".
                // The same string in a *variable* does not, and the plain `=`
                // operator does not either (`"" = 0` is error 13, while
                // `Select Case ""` against `Case 0` is simply no match). All
                // measured; Select Case genuinely has its own comparison.
                let subject_is_const_text = is_constant(subject);
                // A *statically* Boolean subject converts every case value
                // with `CBool` before comparing; a Variant that merely holds a
                // Boolean does not. See `is_statically_boolean`.
                let subject_is_static_bool = is_statically_boolean(subject);
                let subject = self.eval(subject, frame)?;
                let text_compare = subject_is_const_text && matches!(subject, Variant::Str(_));
                let bool_compare = subject_is_static_bool && matches!(subject, Variant::Boolean(_));
                for case in cases {
                    for m in &case.matches {
                        if self.case_matches(&subject, m, frame, text_compare, bool_compare)? {
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

            Stmt::ForEach {
                var,
                iterable,
                body,
                ..
            } => self.exec_for_each(var, iterable, body, frame),

            Stmt::DoLoop {
                pre, post, body, ..
            } => self.exec_do(pre.as_ref(), post.as_ref(), body, frame),

            Stmt::With { subject, body, .. } => {
                // The subject is evaluated once, on entry, and *not* through
                // `scalar`: `With ws.Range("A2")` binds the Range, which is
                // what makes `.Value` inside the block mean the cell.
                let subject = self.eval(subject, frame)?;
                frame.with_stack.push(subject);
                let flow = self.exec_block(body, frame);
                frame.with_stack.pop();
                flow
            }

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
        text_compare: bool,
        bool_compare: bool,
    ) -> VResult<bool> {
        // See `Stmt::SelectCase` for why a constant String subject compares
        // as text.
        let cmp =
            |lhs: &Variant, rhs: &Variant, kind: Operand| -> VResult<Option<std::cmp::Ordering>> {
                if text_compare {
                    return Ok(Some(lhs.to_vba_string()?.cmp(&rhs.to_vba_string()?)));
                }
                value::compare_ctx(lhs, rhs, Operand::Runtime, kind)
            };
        // Against a statically Boolean subject every case value is converted
        // with `CBool` and the comparison then runs on the Booleans, in all
        // three case forms. That single rule produces the whole measured
        // table, including the parts that look inconsistent:
        //
        //   Case 1        matches True   -- CBool(1) is True
        //   Case 0        misses  True
        //   Case 2 To 5   matches True   -- both ends become True
        //   Case 0 To 1   misses  True   -- the range is False To True, i.e.
        //                                   0 To -1, which is empty
        //   Case Is < 0   matches True   -- True is -1
        //   Case Null     is error 94    -- CBool(Null) raises it
        //
        // The `Case 0 To 1` and `Case Null` rows are what rule out "compare
        // the case value as a Boolean": the conversion happens first, and
        // everything after it is the ordinary comparison.
        let cast = |v: Variant| -> VResult<Variant> {
            if bool_compare {
                return Ok(Variant::Boolean(v.to_bool()?));
            }
            Ok(v)
        };
        Ok(match m {
            CaseMatch::Value(e) => {
                let v = cast(self.eval(e, frame)?)?;
                // The case value carries its own constant-ness, which is what
                // makes `Select Case "10"` match `Case 10`.
                cmp(subject, &v, operand_kind(e))? == Some(std::cmp::Ordering::Equal)
            }
            CaseMatch::Range(lo_e, hi_e) => {
                // A `To` range matches a Null subject, which no other case
                // form does: `Select Case Null` skips `Case 0, 1` and
                // `Case Is > 2` but takes `Case 2 To 5`. Measured directly,
                // and it does not follow from the comparisons -- `Null >= 2`
                // is Null. Excel quirk, matched deliberately.
                if subject.is_null() {
                    return Ok(true);
                }
                let lo = cast(self.eval(lo_e, frame)?)?;
                let hi = cast(self.eval(hi_e, frame)?)?;
                let a = cmp(subject, &lo, operand_kind(lo_e))?;
                let b = cmp(subject, &hi, operand_kind(hi_e))?;
                matches!(a, Some(o) if o != std::cmp::Ordering::Less)
                    && matches!(b, Some(o) if o != std::cmp::Ordering::Greater)
            }
            CaseMatch::Is(op, e) => {
                let v = cast(self.eval(e, frame)?)?;
                let ord = cmp(subject, &v, operand_kind(e))?;
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

        // The counter is assigned *before* the test, not after it, so that
        // after the loop it holds the value that failed -- `For i = 1 To 3`
        // leaves `i` at 4, and `Step 2` leaves it at 5. `Exit For` leaves it
        // at the value the body was running with. All measured.
        let mut current = start;
        loop {
            self.tick()?;
            self.assign(var, number_like(current, start, step_v), frame, false)?;
            let done = if step_v > 0.0 {
                current > limit
            } else {
                current < limit
            };
            if done {
                break;
            }
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

    /// `For`'s counter assignment and every other internal write, which are
    /// never `Set`.
    fn assign(
        &mut self,
        target: &Expr,
        v: Variant,
        frame: &mut Frame,
        module_level: bool,
    ) -> VResult<()> {
        self.assign_with(target, v, frame, module_level, false)
    }

    fn assign_with(
        &mut self,
        target: &Expr,
        v: Variant,
        frame: &mut Frame,
        module_level: bool,
        set: bool,
    ) -> VResult<()> {
        // A property write is the one place the target is *not* evaluated:
        // `ws.Range("A1").Value = 5` has to reach the Range and set a member
        // on it, not read `.Value` and throw the result away.
        match target {
            Expr::Member {
                target: obj, name, ..
            } => {
                let obj = self.member_owner(obj.as_deref(), frame)?;
                let Variant::Object(obj) = obj else {
                    return Err(VbaError::new(
                        424,
                        format!("Object required: .{name} on a {}", obj.type_name()),
                    ));
                };
                return self
                    .host(&format!(".{name}"))?
                    .set_member(&obj, name, &[], &v);
            }
            // `ws.Range("A1") = 5` and `ws.Cells(1, 2) = 5`: the call
            // produces an object, and assigning to it writes its default
            // member. `Set` on the same shape is a property *set*, which VBA
            // needs `Set` + a `Property Set` to mean and this does not have.
            Expr::Call {
                target: t, args, ..
            } if !set => {
                if let Expr::Member { .. } | Expr::Ident { .. } = t.as_ref() {
                    let obj = self.eval(target, frame);
                    if let Ok(Variant::Object(obj)) = obj {
                        return self
                            .host("assignment to an object")?
                            .assign_default(&obj, &v);
                    }
                    // Fall through to the error below, but only after the
                    // call has had its chance -- an *array* element write is
                    // a different, still-unsupported thing.
                    let _ = args;
                }
                return Err(out_of_scope("array or property assignment"));
            }
            _ => {}
        }
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
            Expr::Bang { .. } => Err(out_of_scope("property assignment")),
            Expr::Member { .. } | Expr::Call { .. } => {
                Err(out_of_scope("array or property assignment"))
            }
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
                if let Some(h) = self.host.as_mut()
                    && let Some(obj) = h.global(name)
                {
                    return Ok(Variant::Object(obj));
                }
                if self.host.is_none() && super::host::is_host_name(name) {
                    return Err(needs_workbook(name));
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
                let v = self.scalar(v)?;
                // Unary sign promotes on overflow at runtime and does not
                // between constants, exactly as the binary operators do.
                let mode = if is_constant(expr) {
                    ArithMode::Constant
                } else {
                    ArithMode::Promote
                };
                match op {
                    UnOp::Neg => value::neg(&v, mode),
                    UnOp::Pos => value::pos(&v, mode),
                    UnOp::Not => value::not(&v),
                }
            }

            Expr::Binary { op, lhs, rhs, .. } => {
                let a = self.eval(lhs, frame)?;
                let b = self.eval(rhs, frame)?;
                // `Is` is the one operator that wants the references
                // themselves; everything else reads through the default
                // member first.
                if *op == BinOp::Is {
                    return is_comparison(&a, &b);
                }
                let a = self.scalar(a)?;
                let b = self.scalar(b)?;
                // Two operands the compiler has *typed* use fixed-width
                // arithmetic and overflow; anything Variant promotes. See
                // `value::ArithMode`.
                //
                // This is `is_statically_typed`, not `is_constant`, and §28
                // measured both halves of the difference: `CInt(32767) + 1`
                // and `Sgn(1) + 32767` are error 6 though neither is a
                // constant expression, while `(Empty + 32767) + 1` is 32768
                // though it is one -- `Empty` is a `Variant`, so the
                // expression promotes like a variable.
                let mode = if is_statically_typed(lhs) && is_statically_typed(rhs) {
                    ArithMode::Constant
                } else {
                    ArithMode::Promote
                };
                // Comparison needs each side's constant-ness separately, not
                // just whether both are -- see `value::compare_ctx`.
                let kinds = (operand_kind(lhs), operand_kind(rhs));
                eval_binary(*op, &a, &b, mode, kinds)
            }

            Expr::Call { target, args, .. } => self.eval_call(target, args, frame),

            Expr::Member { target, name, .. } => {
                // `Err.Number` / `Err.Description` come first: `Err` is an
                // interpreter object, not a host one, because error handling
                // works with or without a workbook.
                if let Some(t) = target
                    && let Expr::Ident { name: obj, .. } = t.as_ref()
                    && obj.eq_ignore_ascii_case("err")
                {
                    return Ok(match name.to_ascii_lowercase().as_str() {
                        "number" => Variant::Long(self.err.number),
                        "description" => Variant::Str(self.err.description.clone()),
                        other => return Err(out_of_scope(&format!("Err.{other}"))),
                    });
                }
                self.member(target.as_deref(), name, &[], frame)
            }

            Expr::Bang { name, .. } => Err(out_of_scope(&format!("!{name}"))),
            Expr::Me { .. } => Err(out_of_scope("Me")),
            Expr::New { .. } => Err(out_of_scope("New")),
            Expr::TypeOf { .. } => Err(out_of_scope("TypeOf")),
            Expr::AddressOf { .. } => Err(out_of_scope("AddressOf")),
        }
    }

    /// The object a member reference hangs off: an explicit target, or the
    /// innermost `With` subject for a leading dot.
    fn member_owner(&mut self, target: Option<&Expr>, frame: &mut Frame) -> VResult<Variant> {
        match target {
            Some(e) => self.eval(e, frame),
            None => frame
                .with_stack
                .last()
                .cloned()
                // A leading dot outside a `With` is a compile error in VBA,
                // which has no error number; 91 is the closest runtime
                // analogue and says the same thing.
                .ok_or_else(|| {
                    VbaError::new(
                        91,
                        "Object variable or With block variable not set: a leading '.' outside a With block",
                    )
                }),
        }
    }

    /// Reads `<target>.<name>`, with or without arguments.
    fn member(
        &mut self,
        target: Option<&Expr>,
        name: &str,
        args: &[Variant],
        frame: &mut Frame,
    ) -> VResult<Variant> {
        let owner = self.member_owner(target, frame)?;
        let Variant::Object(obj) = owner else {
            return Err(VbaError::new(
                424,
                format!("Object required: .{name} on a {}", owner.type_name()),
            ));
        };
        self.host(&format!(".{name}"))?.get_member(&obj, name, args)
    }

    /// A value in a context that wants a scalar, reading an object's default
    /// member if that is what it is.
    ///
    /// The single funnel the module doc comment describes. Everything that
    /// computes with a value goes through here; `Set`, `Is`, `TypeName`, a
    /// `With` subject and a user procedure's arguments deliberately do not.
    fn scalar(&mut self, v: Variant) -> VResult<Variant> {
        let Variant::Object(obj) = v else {
            return Ok(v);
        };
        self.host("using an object as a value")?.default_value(&obj)
    }

    fn exec_for_each(
        &mut self,
        var: &Expr,
        iterable: &Expr,
        body: &[Stmt],
        frame: &mut Frame,
    ) -> VResult<Flow> {
        let subject = self.eval(iterable, frame)?;
        // The elements are materialised up front rather than streamed. That
        // is a real constraint -- it is why iterating a range is capped -- but
        // the alternative is holding a borrow of the workbook across the loop
        // body, which is exactly what the value-not-pointer object model
        // exists to avoid.
        let items = match subject {
            Variant::Object(obj) => self.host("For Each")?.iterate(&obj)?,
            Variant::Array(a) => a.values.clone(),
            other => {
                return Err(VbaError::new(
                    438,
                    format!(
                        "Object doesn't support this property or method: For Each over a {}",
                        other.type_name()
                    ),
                ));
            }
        };
        for item in items {
            self.tick()?;
            // The element variable holds a reference, so this is a `Set`-like
            // assignment: `For Each c In ws.Range(...)` makes `c` a Range,
            // not the cell's value.
            self.assign_with(var, item, frame, false, true)?;
            match self.exec_block(body, frame)? {
                Flow::Normal => {}
                Flow::ExitFor => break,
                other => return Ok(other),
            }
        }
        Ok(Flow::Normal)
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

        // `obj.Method(args)` and `.Method(args)` inside a `With`.
        if let Expr::Member {
            target: obj, name, ..
        } = target
        {
            let values = self.eval_args(args, frame)?;
            return self.member(obj.as_deref(), name, &values, frame);
        }

        let Expr::Ident { name, .. } = target else {
            return Err(out_of_scope("this call target"));
        };

        // A user procedure wins over everything else of the same name, as in
        // VBA. Its arguments keep their objects: `Foo ws.Range("A1")` passes
        // the Range, not the cell's value.
        if self.find_procedure(name).is_some() {
            let values = self.eval_args(args, frame)?;
            return self.call_procedure(name, values);
        }

        // A local holding an array, indexed. The parser cannot tell this from
        // a call -- that needs a symbol table -- so it is decided here, by
        // what the name actually holds.
        if let Some(Variant::Array(a)) = self.lookup(name, frame) {
            let values = self.eval_args(args, frame)?;
            let row = values
                .first()
                .map(|v| v.to_f64())
                .transpose()?
                .unwrap_or(0.0);
            let col = match values.get(1) {
                Some(v) => v.to_f64()?,
                // A 2-D array read with one index is error 9 in VBA, which
                // `VarArray::get` reports for column 0.
                None => 0.0,
            };
            return a.get(row as usize, col as usize);
        }

        let values = self.eval_args(args, frame)?;
        // A builtin sees scalars: `Len(ws.Range("A1"))` measures the cell's
        // value. The type-inspection builtins are the exception, and are the
        // reason this is a list rather than a blanket conversion -- deref
        // `TypeName`'s argument and it can only ever answer about the value.
        let values = if OBJECT_AWARE_BUILTINS.contains(&name.to_ascii_lowercase().as_str()) {
            values
        } else {
            values
                .into_iter()
                .map(|v| self.scalar(v))
                .collect::<VResult<Vec<_>>>()?
        };
        if let Some(v) = builtins::call(name, &values)? {
            return Ok(v);
        }
        // `Range("A1")`, `Cells(2, 3)`, `Worksheets(1)` -- the host's own
        // unqualified constructors, tried last so a user procedure or a
        // builtin of the same name still shadows them.
        if let Some(h) = self.host.as_mut()
            && let Some(r) = h.global_call(name, &values)
        {
            return r;
        }
        if self.host.is_none() && super::host::is_host_name(name) {
            return Err(needs_workbook(name));
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

/// Whether an expression is a compile-time constant.
///
/// This is *constness*, not static typing, and the two come apart in both
/// directions -- see [`is_statically_typed`], which is what decides whether
/// arithmetic overflows or promotes (§28). `CInt(32767)` is statically typed
/// and not constant; `(Empty + 1)` is constant and not statically typed.
fn is_constant(e: &Expr) -> bool {
    match e {
        // `Null` is not foldable, so nothing containing it is constant.
        // `(False & Null) = (0.1 / -2.5)` is simply False, where the same
        // comparison with a foldable string is error 13.
        Expr::Literal(Literal::Null) => false,
        Expr::Literal(_) => true,
        Expr::Paren { expr, .. } => is_constant(expr),
        Expr::Unary { expr, .. } => is_constant(expr),
        Expr::Binary { lhs, rhs, .. } => is_constant(lhs) && is_constant(rhs),
        _ => false,
    }
}

/// Intrinsics whose return type is declared numeric rather than `Variant`.
///
/// This matters for comparison, not for arithmetic. `value::compare_ctx`'s
/// "constant" case is really "the compiler knows this side's numeric type
/// statically", and a call to one of these qualifies just as a literal does:
/// `(1.5 & "abc") <> CLng(a)` is error 13, while `(1.5 & "abc") <> a` with
/// `a = -1` compares fine, because `a` is a `Variant` and the runtime
/// number-sorts-before-string rule applies instead. Measured.
///
/// `Len`, `Val` and `Sgn` belong here alongside the `C*` conversions, and the
/// discriminating case has to hold the *string* side constant to show it:
/// against `(-32768 & -2.5)` all four raise error 13 while `Int(a)`, `Abs(a)`
/// and a bare `a` do not. An earlier round put `Len` in on the strength of
/// its `As Long` signature, tested it against a *runtime* string -- where
/// nothing is strict, see `compare_ctx` -- concluded it did not belong, and
/// took it out again. `Int` and `Abs` stay out for a reason that is visible
/// in their signatures: they return the type they were handed, so a Variant
/// argument makes them Variant, where `Len` is always `Long`.
const STATICALLY_NUMERIC: &[&str] = &[
    "cint", "clng", "cdbl", "csng", "ccur", "cbool", "cbyte", "len", "val", "sgn",
];

/// Intrinsics whose return type is declared `Boolean`.
///
/// The same "the compiler knows this statically" idea as
/// [`STATICALLY_NUMERIC`] (which lists `cbool` too, for the numeric
/// comparison rule), used by `Select Case` to decide whether to convert its
/// case values with `CBool`. Measured for `CBool`, `IsNumeric`, `IsNull`,
/// `IsEmpty`, `IsDate` and `IsObject`; `IsArray` and `IsError` measure the
/// same way in Excel but are not implemented here yet, and are listed so they
/// arrive with the right behaviour rather than silently as Variants.
const STATICALLY_BOOLEAN: &[&str] = &[
    "cbool",
    "isnumeric",
    "isnull",
    "isempty",
    "isdate",
    "isobject",
    "isarray",
    "iserror",
];

/// Intrinsics whose return type is declared `String`.
///
/// The pair `True Eqv CStr(True)` (error 13) against `LCase("TRUE") Eqv True`
/// (True) is what pins the distinction down -- see [`value::logical_pair`].
///
/// `TypeName` was added on the strength of a measurement, not its signature:
/// `TypeName(32767) >= False` is error 13 in Excel while
/// `LCase("Integer") >= (Not True)` is True, and the difference is exactly
/// that `TypeName` returns `String` where `LCase` returns `Variant`. Found by
/// `fuzz/fuzz_vba.py`. `LCase`, `UCase`, `Left` and the rest stay out for the
/// reason above -- it is their `$`-suffixed forms that are typed `String`.
///
/// `StrReverse`, `Replace` and `Join` are the members of that same family with
/// **no** `$` form, so the plain name is the typed-`String` one. Measured, and
/// the contrast with their Variant-returning neighbours is what places them:
///
///   StrReverse("abc")        > True    error 13
///   Replace("abc", "a", "z") > True    error 13
///   Join(Array("a", "b"))    > True    error 13
///   Trim("abc")              > True    True
///   LTrim("abc")             > True    True
///
/// `StrReverse` is the one `fuzz/fuzz_vba.py` found, as a whole procedure
/// diverging on which error it raised: Excel stopped at the comparison with
/// 13 while visi ran on to a later division by zero and raised 11.
///
/// `Join` is listed though it is not implemented yet (the call raises 35
/// first), for the reason `IsArray` is listed in [`STATICALLY_BOOLEAN`] -- so
/// it arrives with the right type rather than silently as a Variant.
const STATICALLY_STRING: &[&str] = &["cstr", "typename", "strreverse", "replace", "join"];

/// Whether an expression's *static* type is `Boolean`, as the VBA compiler
/// would know it.
///
/// This is the distinction `Select Case` turns on, and it is invisible in the
/// value: `Select Case CBool(a)` matches `Case 1`, while `Select Case a` with
/// `a = True` does not, though both subjects are `True` at run time. A
/// constant expression qualifies because the compiler folds it (`Select Case
/// (1 = 1)` behaves as `Select Case True`); a Variant never does, whatever it
/// happens to hold.
fn is_statically_boolean(e: &Expr) -> bool {
    match e {
        Expr::Paren { expr, .. } => is_statically_boolean(expr),
        // `Not` of a Boolean is a Boolean, so it carries the static type
        // through: `Select Case (Not IsEmpty("Z"))` takes `Case 0, 1` -- the
        // case values convert with `CBool` -- where the same subject read as
        // a plain -1 takes `Case Else`. `Not` of a *number* is a number and
        // does not, which the `Variant::Boolean` check at the use site
        // enforces anyway: `Select Case (Not 5)` is -6 and matches neither.
        // Measured; found by `fuzz/fuzz_vba.py`.
        Expr::Unary {
            op: UnOp::Not,
            expr,
            ..
        } => is_statically_boolean(expr),
        Expr::Call { target, .. } => matches!(target.as_ref(), Expr::Ident { name, .. }
            if STATICALLY_BOOLEAN.contains(&name.to_ascii_lowercase().as_str())),
        _ => is_constant(e),
    }
}

/// Whether the compiler knows this expression's type without its value.
///
/// A call to one of the declared-return-type intrinsics qualifies, and so
/// does **arithmetic over them** -- `Len(CStr(a)) / 2` is a `Double` as
/// surely as `Len(CStr(a))` is a `Long`, because every operand's type is
/// known. One `Variant` operand loses it for the whole expression, which is
/// why `Len(CStr(a)) + a` is not static.
///
/// The propagation is measured, not assumed:
///
/// ```text
/// a = -3 : Len(CStr(a))       = "-7False"   error 13   (bare call)
/// a = -3 : (Len(CStr(a)) / 2) = "-7False"   error 13   (propagated)
/// a = -3 : (Len(CStr(a)) + 1) = "-7False"   error 13   (propagated)
/// a = -3 : (Len(CStr(a)) + a) = "-7False"   False      (a Variant operand)
/// a = -3 : (a / (-32768))     = "-7False"   False      (no static operand)
/// a = -3 : (CLng(a) * 2)      = "-6.0"      True       (numeric, not text)
/// ```
///
/// That last row is the positive half: against a statically typed number the
/// string must parse *and then compares numerically*, where a `Variant`
/// partner would compare it as text and say False.
///
/// **Comparison and `&` propagate too**, which §18 left open for want of a
/// measurement and §24 supplied. A comparison is statically `Boolean` only
/// when both its operands are statically typed, because a `Variant` operand
/// could make it `Null` -- and that is the whole of the rule §16 had written
/// as a 4x4 table with an unexplained cell:
///
/// ```text
/// "0" >= (3# >= CDbl(0))       True    every operand statically typed
/// "0" >= (Len(CStr(0)) >= 1)   True    likewise
/// "0" >= ("1" >= -7)           True    a string *literal* is statically typed
/// "0" >= (3# >= Empty)         False   `Empty` is a Variant, so this is not
/// b = 1 : "0" >= (3# >= b)     False   nor is a variable
/// "0" >= IsEmpty(Empty)        True    but a declared-Boolean call is
/// ```
///
/// The last two rows are what say this is about the static *type* rather than
/// about `Empty` appearing anywhere: `IsEmpty(Empty)` is declared `Boolean`
/// and converts, while `(3# >= Empty)` does not.
fn is_statically_typed(e: &Expr) -> bool {
    match e {
        // `Empty` and `Null` are `Variant`, not statically typed values. This
        // is the distinction the whole §24 rule turns on: `(3# >= Empty)` is
        // a compile-time *constant* and still not a compile-time `Boolean`.
        Expr::Literal(Literal::Empty | Literal::Null) => false,
        Expr::Literal(_) => true,
        Expr::Paren { expr, .. } | Expr::Unary { expr, .. } => is_statically_typed(expr),
        Expr::Binary { op, lhs, rhs, .. } => {
            matches!(
                op,
                BinOp::Add
                    | BinOp::Sub
                    | BinOp::Mul
                    | BinOp::Div
                    | BinOp::IntDiv
                    | BinOp::Mod
                    | BinOp::Pow
                    // `&` yields a `String` whatever it is handed, and a
                    // comparison a `Boolean` -- provided every operand is
                    // itself statically typed. `("1" & "3")` is strict
                    // where `(Empty & "13")` is not.
                    | BinOp::Concat
                    | BinOp::Eq
                    | BinOp::Ne
                    | BinOp::Lt
                    | BinOp::Gt
                    | BinOp::Le
                    | BinOp::Ge
            ) && is_statically_typed(lhs)
                && is_statically_typed(rhs)
        }
        // Boolean- and String-returning intrinsics count for the same reason
        // the numeric ones do: the compiler knows the type without the value.
        Expr::Call { target, .. } => matches!(target.as_ref(), Expr::Ident { name, .. }
        if {
            let name = name.to_ascii_lowercase();
            STATICALLY_NUMERIC.contains(&name.as_str())
                || STATICALLY_BOOLEAN.contains(&name.as_str())
                || STATICALLY_STRING.contains(&name.as_str())
        }),
        _ => false,
    }
}

/// How `value::compare_ctx` should treat an operand.
fn operand_kind(e: &Expr) -> Operand {
    let statically_typed = is_statically_typed(e);
    match e {
        Expr::Literal(_) => Operand::Literal,
        // A parenthesised or signed literal is still just a literal, however
        // many layers deep: `(Not True)` behaves as `False` does, where the
        // *folded* `(3# >= Empty)` does not, and the two differ only in that
        // one bottoms out at a literal through unary operators and the other
        // through a comparison. Measured -- `TypeName(32767) >= (Not True)`
        // is error 13 while `TypeName(0) >= (3# >= Empty)` compares as text.
        // This used to check one level, which put `(Not True)` and `(-7)` in
        // the wrong bucket.
        Expr::Paren { expr, .. } | Expr::Unary { expr, .. }
            if operand_kind(expr) == Operand::Literal =>
        {
            Operand::Literal
        }
        // Constant *and* statically typed. A constant expression over `Empty`
        // is neither one thing nor the other -- the compiler can fold it but
        // its type is `Variant`, so it behaves exactly as a variable does.
        // `(3# >= Empty)` and `(Empty & "13")` are the measured cases (§24).
        _ if is_constant(e) && statically_typed => Operand::ConstExpr,
        _ if statically_typed => Operand::Static,
        _ => Operand::Runtime,
    }
}

/// The one constant-folding quirk this interpreter reproduces.
///
/// `True Mod "12"` is the **Boolean** `False`, and `True \ "12"` is `True`,
/// where the same expressions with either operand in a variable give the
/// ordinary `Long` results. The model that fits every measurement is that
/// when the *left* operand is a constant `Boolean` and the right is a
/// constant `String`, `\` and `Mod` convert **both** sides with `CBool` and
/// return a `Boolean`.
///
/// Confirmed against eighteen cases, including the ones that pin down how
/// narrow it is: `"12" Mod True` is `Long 0` (so it is left-specific),
/// `True Mod 12` is `Integer -1` (so the partner must be a String),
/// `a = True : a Mod "12"` is `Long -1` (so both must be constants), and
/// `True And "12"` is `Long 12` (so it is only `\` and `Mod`).
/// `True Mod "0"` is error 11, which the CBool conversion explains: `"0"`
/// becomes `False`, i.e. zero.
fn constant_bool_int_op(op: BinOp, a: &Variant, b: &Variant, mode: ArithMode) -> Option<()> {
    (mode == ArithMode::Constant
        && matches!(op, BinOp::IntDiv | BinOp::Mod)
        && matches!(a, Variant::Boolean(_))
        && matches!(b, Variant::Str(_)))
    .then_some(())
}

fn eval_binary(
    op: BinOp,
    a: &Variant,
    b: &Variant,
    mode: ArithMode,
    kinds: (Operand, Operand),
) -> VResult<Variant> {
    use BinOp::*;
    if constant_bool_int_op(op, a, b, mode).is_some() {
        let l: i64 = if a.to_bool()? { -1 } else { 0 };
        let r: i64 = if b.to_bool()? { -1 } else { 0 };
        if r == 0 {
            return Err(VbaError::div_by_zero());
        }
        let v = if op == IntDiv { l / r } else { l % r };
        return Ok(Variant::Boolean(v != 0));
    }
    match op {
        Add => value::add(a, b, mode),
        Sub => value::sub(a, b, mode),
        Mul => value::mul(a, b, mode),
        Div => value::div(a, b),
        IntDiv => value::int_div(a, b),
        Mod => value::modulo(a, b),
        Pow => value::pow(a, b, mode),
        Concat => value::concat(a, b),
        Eq | Ne | Lt | Gt | Le | Ge => match value::compare_ctx(a, b, kinds.0, kinds.1)? {
            None => Ok(Variant::Null),
            Some(ord) => Ok(Variant::Boolean(compare_with(op, ord))),
        },
        // And/Or/Imp are three-valued; Xor and Eqv are not (a Null operand
        // always makes their result unknown).
        And => null_on_the_right(a, b, kinds, value::and(a, b, kinds)),
        Or => null_on_the_right(a, b, kinds, value::or(a, b, kinds)),
        Xor => null_on_the_right(a, b, kinds, value::logical(a, b, kinds, |x, y| x ^ y)),
        Eqv => null_on_the_right(a, b, kinds, value::logical(a, b, kinds, |x, y| !(x ^ y))),
        Imp => null_on_the_right(a, b, kinds, value::imp(a, b, kinds)),
        Like => Err(out_of_scope("Like")),
        // Handled before the operands are dereferenced -- see `eval`.
        Is => is_comparison(a, b),
    }
}

/// Builtins that must see an object rather than its default member.
///
/// Short on purpose. `TypeName` and `VarType` exist to report *what a value
/// is*, and `IsObject` to report whether it is one at all, so dereferencing
/// their argument would make them structurally unable to answer. Everything
/// else -- `Len`, `IsNumeric`, `CStr` -- is asking about the value, which for
/// a `Range` means the cell.
const OBJECT_AWARE_BUILTINS: &[&str] = &["typename", "vartype", "isobject"];

/// `Is`: reference identity.
///
/// Both operands must be objects. `Nothing` is one, which is what makes
/// `r Is Nothing` the ordinary way to test an unset reference; anything else
/// is error 424, VBA's "Object required".
fn is_comparison(a: &Variant, b: &Variant) -> VResult<Variant> {
    match (a.as_object(), b.as_object()) {
        (Some(x), Some(y)) => Ok(Variant::Boolean(x.same_object(y))),
        _ => Err(VbaError::new(424, "Object required: Is compares objects")),
    }
}

/// A statically typed `String` on the **left** of a logical operator, with
/// `Null` on the right, is error 94.
///
/// | Expression | Excel |
/// | --- | --- |
/// | `"  3  " Imp Null`, `"3" And Null`, `"1.5" Or Null`, `"0" Or Null` | error 94 |
/// | `("  " & "3") Or Null`, `CStr(3) Or Null` | error 94 |
/// | `a = Null : "  3  " Or a` | error 94 -- the *Null* may be a variable |
/// | `a = "  3  " : a Imp Null` | not an error -- the **String** may not |
/// | `Null Or "  3  "`, `Null And "  3  "`, `Null Xor "  3  "` | not an error -- it is left-specific |
/// | `"abc" Imp Null`, `"True" Or Null` | error 13 -- the string's own conversion is checked first |
/// | `3 Imp Null`, `255 Imp Null` | not an error -- the operand must be a String |
///
/// Which is why this wraps the operator rather than short-circuiting it: the
/// conversion failures have to surface as themselves, and only a *successful*
/// operation becomes the 94. Measured with `fuzz/vba_expr_probe.py`.
fn null_on_the_right(
    lhs: &Variant,
    rhs: &Variant,
    kinds: (Operand, Operand),
    computed: VResult<Variant>,
) -> VResult<Variant> {
    let statically_string = matches!(lhs, Variant::Str(_)) && kinds.0 != Operand::Runtime;
    if statically_string && rhs.is_null() {
        computed?;
        return Err(VbaError::invalid_null());
    }
    computed
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
        // `#6/22/2026#` is the Date 46195, and `CStr` of it is `6/22/26`.
        // The engine's own date parser reads the literal, so a date written
        // in a macro and a date typed into a cell go through one
        // implementation. A literal it cannot read is Empty rather than a
        // wrong number -- the same refusal Phase 1 made for every date.
        Literal::Date(text) => match crate::core::date::parse_date(text) {
            Some((d, _)) => Variant::Date(crate::core::date::date_to_excel_serial(d)),
            None => Variant::Empty,
        },
        Literal::Nothing => Variant::Object(ObjRef::Nothing),
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
        "date" => Variant::Date(0.0),
        // Measured: `Dim r As Range` leaves `r` reporting `TypeName` of
        // "Nothing" and `r Is Nothing` True, where an untyped `Dim r` is
        // Empty. Any object type behaves the same way, so this matches the
        // host's classes rather than `Range` alone.
        "range" | "worksheet" | "workbook" | "object" | "application" | "sheets" => {
            Variant::Object(ObjRef::Nothing)
        }
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

    // ---- three-valued logic, comparison, loop counters -------------------
    //
    // All measured against Excel 16.112 after fuzz/fuzz_vba.py flagged them.

    #[test]
    fn and_or_and_imp_are_three_valued() {
        // A falsy operand determines And; a truthy one determines Or. The
        // deciding operand is returned unchanged, keeping its type.
        assert_eq!(expr("False And Null"), "Boolean|False");
        assert_eq!(expr("True Or Null"), "Boolean|True");
        assert_eq!(expr("IsNull(True And Null)"), "Boolean|True");
        assert_eq!(expr("IsNull(False Or Null)"), "Boolean|True");
        // Numeric operands keep their own subtype through the same rule.
        assert_eq!(
            run("    Dim a\n    a = 0\n    F = (a And Null)"),
            "Integer|0"
        );
        assert_eq!(
            run("    Dim a\n    a = 5\n    F = (a Or Null)"),
            "Integer|5"
        );
        assert_eq!(
            run("    Dim a\n    a = 5\n    F = IsNull(a And Null)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = 0\n    F = IsNull(a Or Null)"),
            "Boolean|True"
        );
        // Imp is determined by a true consequent or a false antecedent.
        assert_eq!(expr("Null Imp True"), "Boolean|True");
        assert_eq!(expr("False Imp Null"), "Boolean|True");
        // Xor and Eqv are not three-valued: Null always wins.
        assert_eq!(expr("IsNull(Null Xor True)"), "Boolean|True");
        assert_eq!(expr("IsNull(Null Eqv True)"), "Boolean|True");
        assert_eq!(expr("IsNull(Not Null)"), "Boolean|True");
    }

    #[test]
    fn string_versus_number_comparison_depends_on_constant_ness() {
        // The four rules in `value::compare_ctx`, each with the Excel result
        // that established it.

        // Both constant: numeric, and error 13 if the string will not parse.
        assert_eq!(expr("\"10\" = 10"), "Boolean|True");
        assert_eq!(expr("\"2\" > 10"), "Boolean|False");
        assert_eq!(expr("\"\" = 0"), "ERR|13");
        assert_eq!(expr("\"abc\" > 1"), "ERR|13");

        // Numeric constant, string variable: numeric, falling back rather
        // than erroring when the string will not parse.
        assert_eq!(
            run("    Dim a\n    a = \"2\"\n    F = (a > 10)"),
            "Boolean|False"
        );
        assert_eq!(
            run("    Dim a\n    a = \"1.5\"\n    F = (a = 1.5)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = \"\"\n    F = (a = 0)"),
            "Boolean|False"
        );
        assert_eq!(
            run("    Dim a\n    a = \"abc\"\n    F = (a = 1)"),
            "Boolean|False"
        );

        // String constant, numeric variable: string comparison.
        assert_eq!(
            run("    Dim b\n    b = 10\n    F = (\"2\" > b)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim b\n    b = 1\n    F = (\"abc\" > b)"),
            "Boolean|True"
        );

        // A call whose return type is declared numeric counts as statically
        // typed, exactly as a literal does -- see `STATICALLY_NUMERIC`.
        assert_eq!(
            run("    Dim a\n    a = True\n    F = ((1.5 & \"abc\") <> CLng(a))"),
            "ERR|13"
        );
        // The same comparison against a plain Variant uses the runtime rule
        // and does not error.
        assert_eq!(
            run("    Dim a\n    a = -1\n    F = ((1.5 & \"abc\") <> a)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = (\"Z\" <> a)"),
            "Boolean|True"
        );

        // Against a numeric constant the string is coerced by its numeric
        // *prefix*, as Val takes it -- which is what separates these two,
        // identical by every structural property: "1.5False" has the prefix
        // 1.5, "True255" has none.
        assert_eq!(expr("(Not 2!) <= (\"1.5\" & False)"), "Boolean|True");
        assert_eq!(expr("(-True) <> (True & &HFF)"), "ERR|13");
        assert_eq!(expr("\"False\" = -0.04"), "ERR|13");
        assert_eq!(expr("\"1.5abc\" > 1"), "Boolean|True");
        // Null is not foldable, so nothing containing it is constant, and
        // this falls back to the runtime ordering instead of erroring.
        assert_eq!(expr("(False & Null) = (0.1 / -2.5)"), "Boolean|False");

        // A statically-typed numeric partner is strict whatever the string
        // side looks like -- but only the C* conversions qualify. `Len` does
        // not, despite its documented `As Long` signature.
        assert_eq!(
            run("    Dim a\n    a = True\n    F = ((1.5 & \"abc\") <> CLng(a))"),
            "ERR|13"
        );
        assert_eq!(
            run("    Dim a\n    a = 1\n    F = ((\"abc\" & a) <> Len(CStr(\"Z\")))"),
            "Boolean|True"
        );

        // A Boolean partner converts the string with CBool, so the words
        // compare as booleans. A numeric partner does not: `"True" = -1` is
        // error 13.
        assert_eq!(
            run("    Dim a\n    a = \"True\"\n    F = (a = True)"),
            "Boolean|True"
        );
        assert_eq!(expr("\"True\" = -1"), "ERR|13");

        // Both variables: a number sorts before a string, whatever it is.
        // This is the row that defeats every simpler theory -- "1.5" and 1.5
        // are equal both numerically and textually, and Excel says False.
        assert_eq!(
            run("    Dim a, b\n    a = \"1.5\"\n    b = 1.5\n    F = (a = b)"),
            "Boolean|False"
        );
        assert_eq!(
            run("    Dim a, b\n    a = \"2\"\n    b = 10\n    F = (a > b)"),
            "Boolean|True"
        );
    }

    /// A `Select Case` whose subject is a *constant* string compares as
    /// text, even against numeric cases -- and the same string held in a
    /// variable does not. Both halves measured; the split is the same
    /// constant-vs-runtime one the arithmetic and comparison rules have.
    #[test]
    fn a_constant_string_select_subject_compares_as_text() {
        let sel = |subject: &str| {
            format!(
                "    Dim r\n    Select Case {subject}\n    Case 2 To 5\n        r = \"range\"\n    \
                 Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        // Constant subjects: "32768abc" sorts between "2" and "5" as text.
        assert_eq!(run(&sel("\"32768abc\"")), "String|range");
        assert_eq!(run(&sel("(32768 & \"abc\")")), "String|range");
        assert_eq!(run(&sel("\"3\"")), "String|range");
        assert_eq!(run(&sel("\"abc\"")), "String|else");
        assert_eq!(run(&sel("\"7\"")), "String|else");
        assert_eq!(run(&sel("\"1x\"")), "String|else");
        assert_eq!(run(&sel("\"\"")), "String|else");
        // Numeric constant subjects are unaffected.
        assert_eq!(run(&sel("3")), "String|range");
        assert_eq!(run(&sel("7")), "String|else");

        // The same strings in a *variable* use the numeric rule instead, so
        // "32768abc" no longer matches while "3" still does.
        let sel_var = |value: &str| {
            format!(
                "    Dim a, r\n    a = {value}\n    Select Case a\n    Case 2 To 5\n        \
                 r = \"range\"\n    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(run(&sel_var("\"32768abc\"")), "String|else");
        assert_eq!(run(&sel_var("\"3\"")), "String|range");
        assert_eq!(run(&sel_var("\"7\"")), "String|else");
        assert_eq!(run(&sel_var("\"abc\"")), "String|else");
    }

    #[test]
    fn a_constant_string_subject_also_governs_value_and_is_cases() {
        let sel = |cases: &str| {
            format!(
                "    Dim r\n    Select Case \"abc\"\n{cases}    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(
            run(&sel("    Case 3\n        r = \"value\"\n")),
            "String|else"
        );
        // "abc" >= "2" as text, so this one matches.
        assert_eq!(
            run(&sel("    Case Is >= 2\n        r = \"is\"\n")),
            "String|is"
        );
    }

    #[test]
    fn a_case_range_matches_a_null_subject_but_no_other_case_form_does() {
        // Measured, and deliberately not derived: `Null >= 2` is Null, so
        // nothing about the comparisons predicts this.
        let sel = |cases: &str| {
            format!(
                "    Dim r\n    Select Case Null\n{cases}    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(
            run(&sel("    Case 2 To 5\n        r = \"range\"\n")),
            "String|range"
        );
        assert_eq!(
            run(&sel("    Case 0, 1\n        r = \"value\"\n")),
            "String|else"
        );
        assert_eq!(
            run(&sel("    Case Is > 2\n        r = \"is\"\n")),
            "String|else"
        );
    }

    // ---- error ordering (docs/vba-error-ordering.md) --------------------

    #[test]
    fn zero_divided_by_zero_is_overflow_not_division_by_zero() {
        // Measured: only floating-point `/` makes the distinction.
        assert_eq!(expr("1 / 0"), "ERR|11");
        assert_eq!(expr("-1 / 0"), "ERR|11");
        assert_eq!(expr("1.5 / 0"), "ERR|11");
        assert_eq!(expr("0 / 0"), "ERR|6");
        assert_eq!(expr("False / 0"), "ERR|6");
        // `\` and `Mod` stay at 11 even for 0 op 0.
        assert_eq!(expr("0 \\ 0"), "ERR|11");
        assert_eq!(expr("0 Mod 0"), "ERR|11");
    }

    #[test]
    fn division_coerces_both_operands_before_testing_the_divisor() {
        // A type mismatch beats a division by zero. Testing the divisor
        // first masked the real error.
        assert_eq!(expr("\"xxxx\" / 0"), "ERR|13");
        assert_eq!(expr("\"\" / 0"), "ERR|13");
        assert_eq!(expr("0 / \"xxxx\""), "ERR|13");
        assert_eq!(expr("\"abc\" / Null"), "ERR|13");
    }

    #[test]
    fn a_static_string_over_a_null_is_invalid_use_of_null() {
        // Left-specific, and only for a statically typed string. See
        // `null_on_the_right` for the measured table.
        for e in [
            "\"  3  \" Imp Null",
            "\"3\" And Null",
            "\"1.5\" Or Null",
            "\"0\" Or Null",
            "\"  3  \" Xor Null",
            "\"  3  \" Eqv Null",
            "(\"  \" & \"3\") Or Null",
            "CStr(3) Or Null",
        ] {
            assert_eq!(expr(e), "ERR|94", "{e}");
        }
        assert_eq!(
            run("    Dim a\n    a = Null\n    F = IsNull(\"  3  \" Or a)"),
            "ERR|94"
        );
        // A runtime string does not trigger it, and neither does a Null on
        // the left.
        assert_eq!(
            run("    Dim a\n    a = \"  3  \"\n    F = IsNull(a Imp Null)"),
            "Boolean|False"
        );
        assert_eq!(expr("IsNull(Null Or \"  3  \")"), "Boolean|False");
        assert_eq!(expr("IsNull(Null Xor \"  3  \")"), "Boolean|True");
        // The string's own conversion is checked first: these stay 13.
        assert_eq!(expr("\"abc\" Imp Null"), "ERR|13");
        assert_eq!(expr("\"True\" Or Null"), "ERR|13");
        // A numeric operand is unaffected.
        assert_eq!(expr("IsNull(255 Imp Null)"), "Boolean|False");
    }

    #[test]
    fn a_statically_typed_numeric_partner_is_strict_only_against_a_constant_string() {
        // `Len`, `Val` and `Sgn` are declared numeric like the `C*`
        // conversions, so a constant string compared against one has to parse
        // whole; `Int` and `Abs` return their argument's type and do not.
        let with = |e: &str| run(&format!("    Dim va\n    va = 1\n    F = {e}"));
        for f in ["CLng(va)", "Len(CStr(va))", "Val(CStr(va))", "Sgn(va)"] {
            assert_eq!(with(&format!("({f} > (-32768 & -2.5))")), "ERR|13", "{f}");
        }
        for f in ["Int(va)", "Abs(va)", "va"] {
            assert_eq!(
                with(&format!("({f} > (-32768 & -2.5))")),
                "Boolean|True",
                "{f}"
            );
        }
        // A *runtime* string is not held to that: it compares numerically
        // when it parses, and falls back to the ordering when it does not,
        // rather than erroring.
        assert_eq!(
            run("    Dim va, vb\n    va = 5\n    vb = \"1\"\n    F = (CLng(va) < vb)"),
            "Boolean|False"
        );
        assert_eq!(with("(CLng(va) < (\"abc\" & va))"), "Boolean|True");
    }

    #[test]
    fn negating_the_long_minimum_between_constants_wraps_to_itself() {
        // `-(-2147483648)` is arithmetically 2147483648, and Excel gives back
        // the Long -2147483648 -- plain two's complement, and wrong. Narrow:
        // the Integer minimum errors instead, and at run time the whole thing
        // widens to a Double. All three measured, and matched deliberately,
        // since a macro doing this should behave the same way here.
        assert_eq!(expr("TypeName(-(Not 2147483647))"), "String|Long");
        assert_eq!(expr("CStr(-(Not 2147483647))"), "String|-2147483648");
        assert_eq!(expr("CStr(-(Not 32767))"), "ERR|6");
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = CStr(-(Not a))"),
            "String|2147483648"
        );
    }

    #[test]
    fn select_case_sees_not_of_a_boolean_as_statically_boolean() {
        // §7's rule -- `Select Case` converts its case values to the
        // subject's *static* type -- carries through `Not`, because `Not` of
        // a Boolean is a Boolean. Measured; `fuzz/fuzz_vba.py` found it as a
        // case that took `Case Else` here and `Case 0, 1` in Excel, which
        // then raised on an expression the other arm never evaluates.
        let sel = |subject: &str| {
            run(&format!(
                "    Dim c\n    Select Case {subject}\n    Case 0, 1\n        c = \"one\"\n                     Case 2 To 5\n        c = \"range\"\n    Case Else\n        c = \"else\"\n                     End Select\n    F = c"
            ))
        };
        assert_eq!(sel("(Not IsEmpty(\"Z\"))"), "String|one");
        assert_eq!(sel("(Not IsEmpty(\"\"))"), "String|one");
        assert_eq!(sel("(Not (IsEmpty(\"Z\")))"), "String|one");
        assert_eq!(sel("(Not CBool(0))"), "String|one");
        assert_eq!(sel("IsEmpty(\"Z\")"), "String|one");
        // `Not` of a *number* is a number, so this stays on the numeric path
        // and matches nothing.
        assert_eq!(sel("(Not 5)"), "String|else");
    }

    #[test]
    fn overflow_between_constants_is_really_between_statically_typed_operands() {
        // §28. The fixed-width arithmetic that makes `32767 + 1` error 6 is
        // chosen by static *typing*, not by constness, and the two come apart
        // in both directions. Measured; `fuzz/fuzz_vba.py` found it on seed
        // 314159 through `CInt(vb) ^ (vb Mod va)`.
        //
        // Typed but not constant: these overflow.
        assert_eq!(expr("CStr(CInt(32767) + 1)"), "ERR|6");
        assert_eq!(expr("CStr(CInt(32767) * 2)"), "ERR|6");
        assert_eq!(expr("CStr(CInt(32767) + CInt(1))"), "ERR|6");
        assert_eq!(expr("CStr(Sgn(1) + 32767)"), "ERR|6");
        assert_eq!(expr("CStr(CLng(2147483647) + 1)"), "ERR|6");
        assert_eq!(expr("CStr(CInt(32767) ^ 4652)"), "ERR|6");
        assert_eq!(expr("CStr(CDbl(32767) ^ 4652)"), "ERR|6");
        assert_eq!(expr("CStr(Len(\"abcde\") ^ 4652)"), "ERR|6");
        // Constant but not typed: `Empty` is a `Variant`, so this promotes
        // exactly as a variable does. visi had this backwards.
        assert_eq!(expr("CStr((Empty + 32767) + 1)"), "String|32768");
        // Unchanged: literals overflow, a variable promotes.
        assert_eq!(expr("CStr(32767 + 1)"), "ERR|6");
        assert_eq!(
            run("    Dim a\n    a = 32767\n    F = CStr(a + 1)"),
            "String|32768"
        );
        assert_eq!(
            run("    Dim a\n    a = 1\n    F = CStr(CInt(32767) + a)"),
            "String|32768"
        );
        // Typed, but the width is `Long`, so there is nothing to overflow.
        assert_eq!(expr("CStr(Len(\"abcde\") + 32763)"), "String|32768");
        assert_eq!(expr("CStr(CInt(Empty) + 32768)"), "String|32768");
        // `^` yields an infinity rather than overflowing (§4) whenever the
        // expression is not statically typed -- which is what makes the
        // first `^` row above a change and not a contradiction.
        assert_eq!(
            run("    Dim vb\n    vb = 4652\n    F = CStr(32767 ^ vb)"),
            "String|INF"
        );
    }

    #[test]
    fn a_statically_string_value_compares_as_text_against_a_runtime_number() {
        // §27. The other half of §23's split: a declared `String` against a
        // *runtime* number compares as text with the number via `CStr`,
        // exactly as a literal does, where a `Variant`-returning intrinsic
        // orders. Measured; `fuzz/fuzz_vba.py` found it on seed 987654.
        //
        // `a` is a variable throughout, so the number is never static and the
        // strictness of §13/§23 never applies -- these differ only in how
        // well the compiler knows the *string*.
        let with = |setup: &str, e: &str| run(&format!("    Dim a, b\n{setup}\n    F = CStr({e})"));
        assert_eq!(with("    a = 5", "(a < \"10\")"), "String|False");
        assert_eq!(with("    a = 5", "(a < CStr(10))"), "String|False");
        assert_eq!(with("    a = 5", "(a < (CStr(1) & \"0\"))"), "String|False");
        assert_eq!(
            with("    a = 5\n    b = 10", "(a < CStr(b))"),
            "String|False"
        );
        // A `Variant`-returning intrinsic is not statically `String`, so the
        // runtime rule applies instead: the number sorts first.
        assert_eq!(with("    a = 5", "(a < Trim(\"10\"))"), "String|True");
        // Text, and never an error, even when the string will not convert --
        // this is the row the ordering got wrong in both directions.
        assert_eq!(with("    a = -2", "(a < CStr(\"\"))"), "String|False");
        assert_eq!(with("    a = -2", "(a < StrReverse(\"\"))"), "String|False");
        assert_eq!(with("    a = -2", "(a < CStr(\"abc\"))"), "String|True");
        assert_eq!(with("    a = -2", "(a < \"\")"), "String|False");
        // The case as the fuzzer found it: the fold is statically `String`
        // through `&`, so it compares as text and the whole thing is False.
        assert_eq!(
            with(
                "    a = 1\n    b = 1",
                "(((True * 1E3) & Len(CStr(\"Z\"))) > ((-a) - (b ^ 255)))"
            ),
            "String|False"
        );
    }

    #[test]
    fn is_numeric_of_empty_is_true_and_of_null_is_false() {
        // §26. `Empty` answers as the 0 it coerces to; `Null` answers for
        // nothing; `""` is not numeric despite comparing equal to `Empty`.
        // Measured. `fuzz/fuzz_vba.py` found it on seed 862021 as
        // `(Not vc) Xor IsNumeric(Empty)`, which is 1 when the operand is
        // False and -2 when it is True.
        assert_eq!(expr("CStr(IsNumeric(Empty))"), "String|True");
        assert_eq!(expr("CStr(IsNumeric(Null))"), "String|False");
        assert_eq!(expr("CStr(IsNumeric(\"\"))"), "String|False");
        assert_eq!(
            run("    Dim vc\n    vc = -2.5\n    F = CStr((Not vc) Xor IsNumeric(Empty))"),
            "String|-2"
        );
    }

    #[test]
    fn instr_of_an_empty_haystack_is_zero() {
        // `InStr("", "")` is 0 while `InStr("a", "")` is 1: an empty needle
        // matches at the start position only when there is a string to match
        // in. Measured; this used to report 1 for the empty/empty pair.
        assert_eq!(expr("CStr(InStr(\"\", \"\"))"), "String|0");
        assert_eq!(expr("CStr(InStr(Empty, \"\"))"), "String|0");
        assert_eq!(expr("CStr(InStr(\"a\", \"\"))"), "String|1");
        assert_eq!(expr("CStr(InStr(\"\", \"a\"))"), "String|0");
    }

    #[test]
    fn static_typing_propagates_through_arithmetic() {
        // `Len(CStr(a)) / 2` is a Double as surely as `Len(CStr(a))` is a
        // Long -- every operand's type is known -- so the strictness of the
        // test above applies to the whole expression. One Variant operand
        // loses it.
        //
        // Found by `fuzz/fuzz_vba.py` on an unseen seed, which is worth
        // noting: the rule itself is §13, already implemented and tested, and
        // what was missing was only that it stopped at the top-level call.
        let with = |e: &str| run(&format!("    Dim a\n    a = -3\n    F = {e}"));
        assert_eq!(with("(Len(CStr(a)) = \"-7False\")"), "ERR|13");
        assert_eq!(with("((Len(CStr(a)) / 2) = \"-7False\")"), "ERR|13");
        assert_eq!(with("((Len(CStr(a)) + 1) = \"-7False\")"), "ERR|13");
        assert_eq!(with("((CLng(a) / 2) = \"abc\")"), "ERR|13");
        assert_eq!(
            with("((Len(CStr(a)) / (-32768)) = ((-7) & (0 > \"1.5\")))"),
            "ERR|13"
        );
        // A Variant operand anywhere in the arithmetic makes the whole
        // expression a Variant, and then the string compares as text.
        assert_eq!(with("((Len(CStr(a)) + a) = \"-7False\")"), "Boolean|False");
        assert_eq!(with("((a / (-32768)) = \"-7False\")"), "Boolean|False");
        assert_eq!(with("((a + 1) = \"-7False\")"), "Boolean|False");
        // The positive half: against a statically typed number a string that
        // *does* parse compares **numerically**, where a Variant partner
        // would compare it as text and say False.
        assert_eq!(with("((CLng(a) * 2) = \"-6.0\")"), "Boolean|True");
    }

    #[test]
    fn static_typing_propagates_through_comparison_and_concatenation() {
        // §24. The half of §18 it explicitly left open: `&` yields a `String`
        // and a comparison a `Boolean`, provided every operand is itself
        // statically typed. `Empty` is a `Variant`, so a fold over it is
        // neither -- which is what §16's "one cell that resists explanation"
        // actually was.
        //
        // The Boolean side. All measured with `fuzz/vba_expr_probe.py`
        // against the same literal string, so only the partner varies:
        // convert says True (CBool("0") is 0, and 0 >= -1), text says False
        // ("0" sorts below "True").
        assert_eq!(expr("(\"0\" >= (3# >= CDbl(0)))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (Len(CStr(0)) >= 1))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (\"1\" >= -7))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (2 >= 1))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (1 = 1))"), "Boolean|True");
        // ...and the same shapes with a `Variant` operand, which could yield
        // `Null` and so is not statically `Boolean`. A declared-Boolean call
        // over `Empty` still is, which is what says this is about the static
        // type and not about `Empty` appearing anywhere.
        assert_eq!(expr("(\"0\" >= (3# >= Empty))"), "Boolean|False");
        assert_eq!(expr("(\"0\" >= (Empty = Empty))"), "Boolean|False");
        assert_eq!(expr("(\"0\" < (3# >= Empty))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= IsEmpty(Empty))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= CBool(Empty))"), "Boolean|True");
        assert_eq!(
            run("    Dim b\n    b = 1\n    F = (\"0\" >= (3# >= b))"),
            "Boolean|False"
        );
        // The String side, against a Boolean that is *not* static, where a
        // statically typed String compares as text and a Variant takes the
        // numeric rules. `("1" + "3")` is the case the fuzzer reduced to:
        // a fold of two string literals is a `String` as surely as a literal
        // is, where a fold over `Empty` is not.
        let folded = |s: &str| expr(&format!("({s} <= (\"\" <> Empty))"));
        assert_eq!(folded("\"13\""), "Boolean|True");
        assert_eq!(folded("(\"1\" + \"3\")"), "Boolean|True");
        assert_eq!(folded("(\"1\" & \"3\")"), "Boolean|True");
        assert_eq!(folded("CStr(13)"), "Boolean|True");
        assert_eq!(folded("(CStr(13) & CStr(0))"), "Boolean|True");
        assert_eq!(folded("(Empty & \"13\")"), "Boolean|False");
        assert_eq!(
            run("    Dim a\n    a = \"13\"\n    F = (a <= (\"\" <> Empty))"),
            "Boolean|False"
        );
        // Against a Boolean that *is* static, every string kind converts --
        // including the fold, which used to take the numeric path here and
        // is what `fuzz/fuzz_vba.py` found on seed 271828.
        assert_eq!(expr("((\"1\" + \"3\") <= False)"), "Boolean|True");
        assert_eq!(expr("((Empty & \"13\") <= False)"), "Boolean|True");
        assert_eq!(expr("((\"1\" + \"3\") = True)"), "Boolean|True");
        assert_eq!(expr("((\"1\" + \"3\") > False)"), "Boolean|False");
        // A fold that will not convert is error 13, exactly as the literal
        // it is: the strictness follows the static `String` type.
        assert_eq!(expr("((\"1\" + \"  3  \") <= False)"), "ERR|13");
        assert_eq!(expr("((\"abc\" + \"d\") > True)"), "ERR|13");
        // ...while the same unconvertible string through a Variant orders
        // above the number (§23) instead of raising.
        assert_eq!(expr("((Empty & \"1  3  \") <= False)"), "Boolean|False");
        // Text, not an error, when the Boolean is not static -- even though
        // the string will not convert.
        assert_eq!(
            expr("((\"1\" & \"  3  \") <= (\"\" <> Empty))"),
            "Boolean|True"
        );
    }

    #[test]
    fn a_string_converts_with_cbool_against_a_static_boolean() {
        // Measured with `fuzz/vba_expr_probe.py`. This used to read "compares
        // as text", which fit `("011" < False)` -- True under both readings --
        // and was wrong about every case that discriminates them:
        // `a = "-1"` makes `a = True` **True**, which no text comparison
        // produces. `fuzz/fuzz_vba.py` found it on a generated case whose
        // visible symptom was a *cell* holding the wrong value.
        //
        // The rule: convert the string with `CBool`, compare as Booleans, and
        // fall back to the ordinary runtime ordering only when the conversion
        // fails. Ordering is numeric, so True (-1) sorts below False (0).
        let with = |setup: &str, e: &str| run(&format!("    Dim va, vb\n{setup}\n    F = {e}"));
        assert_eq!(with("    va = \"011\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"0\"", "(va = False)"), "Boolean|True");
        assert_eq!(with("    va = \"2\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"-1\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"1.5\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"-1\"", "(va <> True)"), "Boolean|False");
        assert_eq!(with("    va = \"011\"", "(va < False)"), "Boolean|True");
        assert_eq!(with("    va = \"011\"", "(va > False)"), "Boolean|False");
        assert_eq!(with("    va = \"011\"", "(va > True)"), "Boolean|False");
        // The two that pin down *that* there is a fallback: `CBool` raises for
        // both, yet neither comparison does -- they are simply unequal. What
        // the fallback *is* takes the ordering cases in
        // `an_unconvertible_runtime_string_sorts_above_a_static_boolean`;
        // equality cannot tell text from ordering.
        assert_eq!(with("    va = \"abc\"", "(va = True)"), "Boolean|False");
        assert_eq!(with("    va = \"\"", "(va = False)"), "Boolean|False");
        // A *statically* String operand takes the same conversion but does
        // **not** get that fallback -- it is error 13 instead. The
        // discriminating rows, all measured: the same string reaches the
        // fallback through a Variant or through a Variant-returning
        // intrinsic, and error 13 only through one declared `As String`.
        assert_eq!(expr("CStr(32767) >= (Not True)"), "Boolean|False");
        assert_eq!(expr("TypeName(32767) >= False"), "ERR|13");
        assert_eq!(expr("(TypeName(32767) >= (Not True))"), "ERR|13");
        assert_eq!(expr("LCase(\"Integer\") >= (Not True)"), "Boolean|True");
        assert_eq!(
            run("    Dim va\n    va = TypeName(32767)\n    F = (va >= (Not True))"),
            "Boolean|True"
        );
        assert_eq!(with("    va = \"011\"", "(va < CBool(0))"), "Boolean|True");
        assert_eq!(
            with("    va = \"011\"", "(va < IsNull(32768))"),
            "Boolean|True"
        );
        // A string *literal* converts too, and a conversion failure is
        // error 13 rather than the text fallback.
        assert_eq!(expr("(\"abc\" < True)"), "ERR|13");
        assert_eq!(expr("(\"Z\" < True)"), "ERR|13");
        assert_eq!(expr("(False >= \"abc\")"), "ERR|13");
        assert_eq!(expr("(\"\" = False)"), "ERR|13");
        // These two are what the numeric reading got wrong, and they are the
        // same rule: `CBool("011")` and `CBool("12")` are both True (-1),
        // which sorts *below* False (0). `("011" < False)` in particular sat
        // in this file as an unexplained divergence for the whole of Phase 1;
        // `fuzz/fuzz_vba.py` re-surfaced it as `(False > "12")` and the
        // `CBool` model accounts for both.
        assert_eq!(expr("(\"011\" < False)"), "Boolean|True");
        assert_eq!(expr("(False > \"12\")"), "Boolean|True");
        assert_eq!(expr("(\"0\" = False)"), "Boolean|True");
        // Neither side is statically typed here -- `Empty` is a `Variant`, so
        // the fold over it is not a `String` and the comparison is not a
        // `Boolean` -- so both fall to the numeric rules: `"1"` becomes 1, and
        // `1 <= 0` is False, where the conversion would say True. See §24.
        assert_eq!(
            expr("((Empty & \"1\") <= (\"\" <> Empty))"),
            "Boolean|False"
        );
        // Against a Boolean that is *not* statically `Boolean` -- here a
        // comparison with an `Empty` operand -- a statically typed `String`
        // compares as text, with the Boolean rendered "True"/"False".
        assert_eq!(expr("TypeName(0) >= (3# >= Empty)"), "Boolean|False");
        assert_eq!(expr("(3# >= Empty) >= TypeName(0)"), "Boolean|True");
        assert_eq!(expr("CStr(0) >= (3# >= Empty)"), "Boolean|False");
        assert_eq!(expr("(Not True) <= CStr(32767)"), "Boolean|False");
        assert_eq!(expr("False >= TypeName(0)"), "ERR|13");
        // ...where a comparison whose operands *are* all statically typed is
        // a statically known Boolean, and converts. This pair looks like an
        // exception about static strings and is not one: the two differ in
        // the **Boolean**, not the string -- see §24 and the test below.
        assert_eq!(expr("(\"000\" < (\"1\" >= -7))"), "Boolean|False");
        assert_eq!(
            run("    Dim va\n    va = \"000\"\n    F = (va < (\"1\" >= -7))"),
            "Boolean|False"
        );
        assert_eq!(expr("(Right(100000, 3) < (\"1\" >= -7))"), "Boolean|False");
        // A *static* Boolean partner converts against every string kind,
        // including a static one.
        assert_eq!(expr("(CStr(0) >= CBool(1))"), "Boolean|True");
        assert_eq!(expr("(\"000\" < CBool(1))"), "Boolean|False");
        assert_eq!(expr("(TypeName(0) >= CBool(1))"), "ERR|13");
        // A Boolean *variable* is not static at all, so none of this applies
        // and the runtime rule takes over: a number sorts before a string.
        assert_eq!(
            with("    va = \"011\"\n    vb = False", "(va < vb)"),
            "Boolean|False"
        );
        // The words take the same path -- `CBool` accepts them too --
        // case-insensitively, and order as the Booleans they become.
        assert_eq!(with("    va = \"True\"", "(va < False)"), "Boolean|True");
        assert_eq!(with("    va = \"true\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"TRUE\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"true\"", "(va = False)"), "Boolean|False");
        // A Boolean variable is not static: the number sorts before the
        // string, so "011" is Greater and `<` is False.
        assert_eq!(
            with("    va = \"011\"\n    vb = False", "(va < vb)"),
            "Boolean|False"
        );
        // A numeric partner is unaffected, and still refuses the words.
        assert_eq!(with("    va = \"011\"", "(va < 0)"), "Boolean|False");
        assert_eq!(expr("(\"True\" = -1)"), "ERR|13");
    }

    #[test]
    fn an_unconvertible_runtime_string_sorts_above_a_static_boolean() {
        // When `CBool` will not take the string, a *runtime* one falls back to
        // the ordinary runtime rule -- the number sorts first, so the string
        // is Greater whatever the two spell. This was written as a text
        // comparison, which every case available at the time agreed with:
        // `"abc"`, `"Integer"` and `""` all sort on the same side of
        // `"True"`/`"False"` as the ordering rule puts them, so the two
        // readings only come apart on a string that does not -- `"ABC"`,
        // whose `A` sorts below both words.
        //
        // Measured with `fuzz/vba_expr_probe.py`; every expectation here is
        // what Excel returned. Found while reducing the `StrReverse` case in
        // `statically_string_intrinsics_are_strict_against_a_boolean`.
        let with = |setup: &str, e: &str| run(&format!("    Dim va\n{setup}\n    F = {e}"));
        assert_eq!(with("    va = \"ABC\"", "(va > True)"), "Boolean|True");
        assert_eq!(with("    va = \"ABC\"", "(va < True)"), "Boolean|False");
        assert_eq!(with("    va = \"ABC\"", "(va >= False)"), "Boolean|True");
        // The same through a `Variant`-returning intrinsic, which is how the
        // fuzzer's generated code reaches it.
        assert_eq!(expr("Chr(65) > True"), "Boolean|True");
        assert_eq!(expr("Chr(65) > False"), "Boolean|True");
        assert_eq!(expr("Hex(255) > True"), "Boolean|True");
        assert_eq!(expr("Space(2) > True"), "Boolean|True");
        // The cases the text reading was derived from still hold -- they are
        // simply blind to the difference.
        assert_eq!(with("    va = \"abc\"", "(va = True)"), "Boolean|False");
        assert_eq!(expr("LCase(\"Integer\") >= (Not True)"), "Boolean|True");
    }

    #[test]
    fn statically_string_intrinsics_are_strict_against_a_boolean() {
        // `StrReverse`, `Replace` and `Join` are declared `As String` and have
        // no `$` form, so the plain name is the typed one -- an unconvertible
        // result against a statically known Boolean is error 13, where the
        // Variant-returning neighbours fall back to ordering instead.
        //
        // `fuzz/fuzz_vba.py` found this as a whole-procedure divergence: Excel
        // stopped at `StrReverse(False) > (Not False)` with 13 while visi took
        // the comparison as True, ran into the other branch, and raised 11 on
        // a division by zero Excel never reached. Measured with
        // `fuzz/vba_expr_probe.py`; see `STATICALLY_STRING`.
        assert_eq!(expr("StrReverse(False) > (Not False)"), "ERR|13");
        assert_eq!(expr("StrReverse(\"abc\") > True"), "ERR|13");
        assert_eq!(expr("StrReverse(\"abc\") > 5"), "ERR|13");
        assert_eq!(expr("True > StrReverse(\"abc\")"), "ERR|13");
        assert_eq!(expr("Replace(\"abc\", \"a\", \"z\") > True"), "ERR|13");
        // A *numeric* partner is strict the same way, and was already wrong
        // for `CStr`/`TypeName` before `StrReverse` joined them: the strictness
        // keyed off the string being *constant* rather than merely typed.
        assert_eq!(expr("CStr(\"abc\") > 5"), "ERR|13");
        assert_eq!(expr("TypeName(1) > 5"), "ERR|13");
        assert_eq!(expr("CStr(\"abc\") >= 0"), "ERR|13");
        assert_eq!(expr("CStr(\"abc\") > CLng(1)"), "ERR|13");
        assert_eq!(expr("TypeName(1) > CLng(5)"), "ERR|13");
        // Variant-returning neighbours are unaffected, against either partner.
        assert_eq!(expr("Trim(\"abc\") > True"), "Boolean|True");
        assert_eq!(expr("LTrim(\"abc\") > True"), "Boolean|True");
        assert_eq!(expr("Trim(\"abc\") > 5"), "Boolean|True");
        assert_eq!(expr("Chr(65) > 5"), "Boolean|True");
        // A typed string that *does* convert is not an error.
        assert_eq!(expr("CStr(\"11\") > 5"), "Boolean|True");
        // Two strings still compare as text, whatever their kinds.
        assert_eq!(expr("TypeName(1) > \"5\""), "Boolean|True");
        // The same string through a Variant is not statically typed, so it
        // compares rather than erroring -- the row that makes this about the
        // declared return type and not the value.
        assert_eq!(
            run("    Dim va\n    va = StrReverse(\"abc\")\n    F = (va > True)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim va\n    va = 5\n    F = (StrReverse(\"abc\") > va)"),
            "Boolean|True"
        );
        // A convertible result still converts: `CBool("11")` is True (-1),
        // which sorts below False (0).
        assert_eq!(expr("StrReverse(\"11\") > False"), "Boolean|False");
    }

    #[test]
    fn division_overflows_rather_than_returning_an_infinity() {
        // `/` was the last operator handing back an INF where Excel raises
        // error 6, at run time as well as between constants. Measured with
        // `fuzz/vba_expr_probe.py`; `^` remains the one operator that does
        // produce infinities, and feeding one of those to `/` raises too.
        assert_eq!(expr("1E308 / 1E-308"), "ERR|6");
        assert_eq!(
            run("    Dim a, b\n    a = 1E308\n    b = 1E-308\n    F = a / b"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 3.75\n    b = a ^ 32767\n    F = b / 2"),
            "ERR|6"
        );
        // Ordinary division is untouched, and so are the two zero cases.
        assert_eq!(expr("1 / 2"), "Double|0.5");
        assert_eq!(expr("1 / 0"), "ERR|11");
        assert_eq!(expr("0 / 0"), "ERR|6");
    }

    #[test]
    fn pow_overflows_between_constants_and_yields_infinity_at_runtime() {
        // The same ArithMode split the other operators have. Phase 1
        // measured the runtime half and missed this one.
        assert_eq!(expr("3.75 ^ 32767"), "ERR|6");
        assert_eq!(expr("255 ^ 255"), "ERR|6");
        assert_eq!(
            run("    Dim a\n    a = 3.75\n    F = (a ^ 32767)"),
            "Double|INF"
        );
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = (a ^ 255)"),
            "Double|INF"
        );
        // A finite constant result is unaffected.
        assert_eq!(expr("2 ^ 10"), "Double|1024");
    }

    #[test]
    fn infinity_is_a_value_for_pow_but_not_for_arithmetic() {
        // `^` produces it, negation preserves it, CStr renders it "INF" --
        // but +, - and * refuse to produce or consume one. The runtime form
        // is used throughout: between constants `^` overflows instead, which
        // `pow_overflows_between_constants_and_yields_infinity_at_runtime`
        // covers.
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = (a ^ 255)"),
            "Double|INF"
        );
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = -(a ^ 255)"),
            "Double|-INF"
        );
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = ((a ^ 255) & \"x\")"),
            "String|INFx"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 255\n    b = (a ^ 255)\n    F = (b + 1)"),
            "ERR|6"
        );
        assert_eq!(run("    Dim a\n    a = 1E300\n    F = (a * a)"), "ERR|6");
        // Finite overflow of an addition is still fine.
        assert_eq!(
            run("    Dim a, b\n    a = 1E300\n    b = 1E300\n    F = (a + b)"),
            "Double|2E+300"
        );
    }

    #[test]
    fn imp_follows_its_definition_rather_than_a_hand_rolled_table() {
        // `255 Imp Null` is `Not 255 Or Null` = `-256 Or Null` = -256,
        // because -256 is truthy. A hand-rolled three-valued table said Null.
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = (a Imp Null)"),
            "Integer|-256"
        );
        // The measured endpoints still hold.
        assert_eq!(expr("Null Imp True"), "Boolean|True");
        assert_eq!(expr("False Imp Null"), "Boolean|True");
        assert_eq!(expr("5 Imp 3"), "Integer|-5");
    }

    #[test]
    fn single_combined_with_long_widens_past_both() {
        // A Single cannot hold every Long, so VBA goes to Double -- but a
        // Single with an Integer stays Single. Both measured.
        assert_eq!(run("    Dim a\n    a = 2!\n    F = (a + 1)"), "Single|3");
        assert_eq!(
            run("    Dim a, b\n    a = 2!\n    b = 1&\n    F = (a * b)"),
            "Double|2"
        );
        assert_eq!(
            run("    Dim a\n    a = 2!\n    F = (a - 0.5)"),
            "Double|1.5"
        );
    }

    /// Which operators coerce a `Null`'s partner before propagating, and
    /// which short-circuit. Measured in both directions with `IsNull`.
    #[test]
    fn only_plus_short_circuits_past_a_bad_partner() {
        // `+` alone returns Null without looking at the other side --
        // plausibly because it cannot tell addition from concatenation
        // without inspecting both, so it gives up first.
        assert_eq!(expr("IsNull(Null + \"Z\")"), "Boolean|True");
        assert_eq!(expr("IsNull(\"Z\" + Null)"), "Boolean|True");
        assert_eq!(expr("IsNull(Null + \"12\")"), "Boolean|True");

        // Every other operator coerces the partner, and a bad string wins.
        for e in [
            "\"Z\" - Null",
            "Null - \"Z\"",
            "\"Z\" * Null",
            "\"Z\" / Null",
            "\"Z\" ^ Null",
            "\"Z\" Mod Null",
            "Null Mod \"Z\"",
            "\"Z\" \\ Null",
            "\"Z\" And Null",
            "Null Or \"Z\"",
        ] {
            assert_eq!(expr(e), "ERR|13", "for {e}");
        }

        // `&` keeps the non-Null side rather than propagating at all.
        assert_eq!(expr("\"Z\" & Null"), "String|Z");

        // A well-formed partner still propagates.
        assert_eq!(expr("IsNull(1 - Null)"), "Boolean|True");
        assert_eq!(expr("IsNull(Null Mod 3)"), "Boolean|True");
    }

    #[test]
    fn unary_sign_promotes_on_overflow_at_runtime() {
        // Same constant-vs-runtime split the binary operators have.
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = (-(Not a))"),
            "Double|2147483648"
        );
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = TypeName(-(Not a))"),
            "String|Double"
        );
        // Integer widens to Long the same way.
        assert_eq!(
            run("    Dim a\n    a = 32767\n    F = (-(Not a))"),
            "Long|32768"
        );
    }

    #[test]
    fn a_statically_boolean_select_subject_converts_its_cases_with_cbool() {
        // Every row measured against Excel 16.112 with
        // `fuzz/vba_expr_probe.py`. A *statically* Boolean subject converts
        // each case value with CBool and compares the Booleans; a Variant
        // holding a Boolean does not, and compares numerically with True as
        // -1. The two halves of this test are the same subject value either
        // side of that line.
        let sel = |subject: &str, cases: &str| {
            format!(
                "    Dim r\n    Select Case {subject}\n{cases}    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        let hit = |subject: &str, case: &str| {
            run(&sel(
                subject,
                &format!("    Case {case}\n        r = \"a\"\n"),
            ))
        };

        // Statically Boolean: a folded constant, or a Boolean-returning
        // intrinsic over a variable.
        for subject in ["(1 = 1)", "True", "CBool(1)", "IsNumeric(0)"] {
            assert_eq!(hit(subject, "1"), "String|a", "{subject} vs Case 1");
            assert_eq!(hit(subject, "0"), "String|else", "{subject} vs Case 0");
            assert_eq!(hit(subject, "0, 1"), "String|a", "{subject} vs Case 0, 1");
            // Both ends become True, so the range is True To True.
            assert_eq!(hit(subject, "2 To 5"), "String|a", "{subject} vs 2 To 5");
            // ... while `0 To 1` becomes False To True, i.e. 0 To -1, which
            // is empty. This row is why the conversion cannot be "compare as
            // Booleans" -- it has to happen before the comparison.
            assert_eq!(hit(subject, "0 To 1"), "String|else", "{subject} vs 0 To 1");
            assert_eq!(hit(subject, "Is = 1"), "String|a", "{subject} vs Is = 1");
            assert_eq!(hit(subject, "Is > 0"), "String|else", "{subject} vs Is > 0");
            assert_eq!(hit(subject, "Is < 0"), "String|a", "{subject} vs Is < 0");
        }
        assert_eq!(hit("(1 = 2)", "0, 1"), "String|a");
        assert_eq!(hit("(1 = 2)", "2 To 5"), "String|else");
        // CBool(Null) is error 94, and the case value goes through CBool.
        assert_eq!(
            run(&sel("CBool(1)", "    Case Null\n        r = \"a\"\n")),
            "ERR|94"
        );

        // The same values in a Variant compare numerically instead.
        let via_var = |value: &str, case: &str| {
            run(&format!(
                "    Dim a, r\n    a = {value}\n    Select Case a\n    Case {case}\n        \
                 r = \"a\"\n    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            ))
        };
        assert_eq!(via_var("True", "0, 1"), "String|else");
        assert_eq!(via_var("True", "-1"), "String|a");
        assert_eq!(via_var("True", "2 To 5"), "String|else");
        assert_eq!(via_var("True", "Is < 0"), "String|a");
        assert_eq!(via_var("False", "0, 1"), "String|a");
    }

    /// The constant-folding quirk in `constant_bool_int_op`, with the
    /// negative controls that pin down how narrow it is.
    #[test]
    fn a_constant_boolean_over_a_constant_string_folds_to_a_boolean() {
        assert_eq!(expr("True Mod \"12\""), "Boolean|False");
        assert_eq!(expr("True \\ \"12\""), "Boolean|True");
        assert_eq!(expr("False \\ \"12\""), "Boolean|False");
        // "0" becomes False, i.e. zero, so these divide by zero.
        assert_eq!(expr("True Mod \"0\""), "ERR|11");
        assert_eq!(expr("True \\ \"0\""), "ERR|11");

        // Left-specific.
        assert_eq!(expr("\"12\" Mod True"), "Long|0");
        assert_eq!(expr("\"12\" \\ True"), "Long|-12");
        // The partner has to be a String.
        assert_eq!(expr("True Mod 12"), "Integer|-1");
        assert_eq!(expr("True \\ 12"), "Integer|0");
        // Both have to be constants.
        assert_eq!(
            run("    Dim a\n    a = True\n    F = (a Mod \"12\")"),
            "Long|-1"
        );
        assert_eq!(
            run("    Dim b\n    b = \"12\"\n    F = (True Mod b)"),
            "Long|-1"
        );
        // Only `\\` and `Mod`.
        assert_eq!(expr("True And \"12\""), "Long|12");
        assert_eq!(expr("True Or \"12\""), "Long|-1");
        assert_eq!(expr("True Eqv \"12\""), "Long|12");
    }

    #[test]
    fn integer_operators_process_the_left_operand_first() {
        // Which error surfaces depends on the order: the left operand
        // overflowing a Long beats a bad string on the right, and vice versa.
        assert_eq!(
            run("    Dim a\n    a = \"32768100000\"\n    F = (a Mod \"Double\")"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a\n    a = \"Double\"\n    F = (a Mod \"32768100000\")"),
            "ERR|13"
        );
        assert_eq!(
            run("    Dim a\n    a = \"32768100000\"\n    F = (a Mod 3)"),
            "ERR|6"
        );
    }

    /// The whole `Null` table, from a sweep of every intrinsic against real
    /// Excel. There is no principle behind the split, so the test enumerates
    /// it -- `Hex` propagates but `Chr` rejects, `String` propagates but
    /// `Space` rejects, `CVar` propagates where every other `C*` rejects.
    #[test]
    fn every_intrinsic_handles_null_the_way_excel_does() {
        for f in [
            "CVar", "Abs", "Int", "Fix", "Round", "Len", "UCase", "LCase", "Trim", "LTrim",
            "RTrim", "Hex", "Oct",
        ] {
            assert_eq!(
                expr(&format!("IsNull({f}(Null))")),
                "Boolean|True",
                "{f} should propagate"
            );
        }
        for e in [
            "Left(Null, 1)",
            "Right(Null, 1)",
            "Mid(Null, 1, 1)",
            "InStr(Null, \"a\")",
            "String(2, Null)",
            "StrComp(Null, \"a\")",
        ] {
            assert_eq!(
                expr(&format!("IsNull({e})")),
                "Boolean|True",
                "{e} should propagate"
            );
        }
        for f in [
            "CStr",
            "CInt",
            "CLng",
            "CDbl",
            "CSng",
            "CBool",
            "CCur",
            "Val",
            "Sgn",
            "Sqr",
            "Exp",
            "Log",
            "Sin",
            "Cos",
            "Tan",
            "Atn",
            "Space",
            "StrReverse",
            "Chr",
            "Asc",
        ] {
            assert_eq!(expr(&format!("{f}(Null)")), "ERR|94", "{f} should reject");
        }
        assert_eq!(expr("Replace(Null, \"a\", \"b\")"), "ERR|94");
        // Inspection functions look at it rather than propagating or rejecting.
        assert_eq!(expr("TypeName(Null)"), "String|Null");
        assert_eq!(expr("IsNull(Null)"), "Boolean|True");
        assert_eq!(expr("IsNumeric(Null)"), "Boolean|False");
        assert_eq!(expr("IsEmpty(Null)"), "Boolean|False");
    }

    #[test]
    fn conversions_reject_null_rather_than_propagating_it() {
        // `CStr(Null)` raises error 94. Propagating a Null instead was a real
        // mismatch: callers put it under `On Error Resume Next` expecting the
        // assignment to be skipped, and a returned Null poisoned everything
        // downstream of it.
        assert_eq!(expr("CStr(Null)"), "ERR|94");
        assert_eq!(expr("CDbl(Null)"), "ERR|94");
        assert_eq!(expr("CLng(Null)"), "ERR|94");
        // String functions do propagate.
        assert_eq!(expr("IsNull(UCase(Null))"), "Boolean|True");
        assert_eq!(expr("IsNull(Left(Null, 1))"), "Boolean|True");
        // Inspection functions look at it rather than propagating.
        assert_eq!(expr("TypeName(Null)"), "String|Null");
        assert_eq!(expr("IsNull(Null)"), "Boolean|True");
    }

    #[test]
    fn the_words_true_and_false_coerce_on_the_integer_path_only() {
        // Measured. The integer/logical path accepts them as -1 and 0; the
        // floating-point path has never heard of them.
        assert_eq!(expr("\"True\" Xor 1"), "Integer|-2");
        assert_eq!(expr("\"False\" Xor 1"), "Integer|1");
        assert_eq!(expr("\"True\" \\ 1"), "Integer|-1");
        assert_eq!(expr("\"True\" Mod 2"), "Integer|-1");
        assert_eq!(expr("CBool(\"True\")"), "Boolean|True");
        // Case-insensitive, and space-tolerant.
        assert_eq!(expr("\"true\" Xor 1"), "Integer|-2");
        assert_eq!(expr("\"TRUE\" Xor 1"), "Integer|-2");
        // `Not` keeps it a Boolean, because both sides of the operation are
        // one; `Xor` with a number goes bitwise and yields an Integer.
        assert_eq!(expr("Not \"True\""), "Boolean|False");

        // Against a Boolean partner the fold is suppressed only when *both*
        // sides are statically typed -- a literal, or a call with a declared
        // return type. `CStr` is declared `As String`; `LCase` returns a
        // Variant, and that pair is what separates the two halves.
        assert_eq!(expr("True Eqv \"True\""), "ERR|13");
        assert_eq!(expr("\"True\" Eqv True"), "ERR|13");
        assert_eq!(expr("True Eqv CStr(True)"), "ERR|13");
        assert_eq!(
            run("    Dim a\n    a = 3.75\n    F = (IsNumeric(a) Eqv CStr(True))"),
            "ERR|13"
        );
        // ... and happens as soon as either side is a Variant.
        assert_eq!(expr("LCase(\"TRUE\") Eqv True"), "Boolean|True");
        assert_eq!(expr("LCase(False) Eqv IsNull(True)"), "Boolean|True");
        assert_eq!(
            run("    Dim a\n    a = True\n    F = (a Eqv \"True\")"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = \"false\"\n    F = (a Eqv False)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = \"true\"\n    F = (a Eqv False)"),
            "Boolean|False"
        );
        assert_eq!(
            run("    Dim a, b\n    a = \"true\"\n    b = False\n    F = (a Eqv b)"),
            "Boolean|False"
        );

        // The floating-point path still rejects them.
        for e in [
            "\"True\" + 1",
            "\"False\" + 1",
            "\"True\" * 2",
            "CDbl(\"True\")",
        ] {
            assert_eq!(expr(e), "ERR|13", "for {e}");
        }
        assert_eq!(expr("IsNumeric(\"True\")"), "Boolean|False");

        // The exact shape the fuzzer hit: Trim of a comparison yields the
        // word, which then has to work as a logical operand.
        assert_eq!(expr("Trim((1 >= 2)) Xor 5"), "Integer|5");
    }

    #[test]
    fn a_string_outside_double_range_fails_to_convert() {
        // Error 6 from the *conversion*, not a quiet infinity -- and not the
        // 13 an unparseable string gives.
        assert_eq!(
            run("    Dim a\n    a = \"1E+2923\"\n    F = (a ^ 255)"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a\n    a = \"1E400\"\n    F = (a + 1)"),
            "ERR|6"
        );
        // The power itself still overflows to infinity happily.
        assert_eq!(
            run("    Dim a\n    a = \"255\"\n    F = (a ^ 255)"),
            "Double|INF"
        );
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = (a ^ 255)"),
            "Double|INF"
        );
    }

    #[test]
    fn an_empty_string_never_coerces_to_a_number() {
        // Measured across every operator: `"" - 3`, `"" + 3`, `"" * 3`,
        // `"" \ 3`, `"" And 1`, `Not ""` and `CDbl("")` are all error 13.
        for e in [
            "\"\" - 3",
            "\"\" + 3",
            "\"\" * 3",
            "\"\" \\ 3",
            "Not \"\"",
            "CDbl(\"\")",
        ] {
            assert_eq!(expr(e), "ERR|13", "for {e}");
        }
    }

    #[test]
    fn val_always_returns_a_double() {
        // Measured directly. A previous version typed the result like a
        // literal, inferred from a fuzz case where `Val` may never have run.
        assert_eq!(expr("Val(255)"), "Double|255");
        assert_eq!(expr("Val(\"1.5\")"), "Double|1.5");
        assert_eq!(expr("Val(\"100000\")"), "Double|100000");
        assert_eq!(run("    Dim a\n    a = 1%\n    F = Val(a)"), "Double|1");
    }

    #[test]
    fn a_zero_base_with_a_negative_exponent_is_an_error() {
        assert_eq!(
            run("    Dim a, b\n    a = 0\n    b = -1\n    F = (a ^ b)"),
            "ERR|5"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 0\n    b = -246\n    F = (a ^ b)"),
            "ERR|5"
        );
        // Zero and positive exponents are fine, as is a negative exponent
        // over a non-zero base.
        assert_eq!(
            run("    Dim a, b\n    a = 0\n    b = 0\n    F = (a ^ b)"),
            "Double|1"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 0\n    b = 2\n    F = (a ^ b)"),
            "Double|0"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 2\n    b = -2\n    F = (a ^ b)"),
            "Double|0.25"
        );
    }

    #[test]
    fn logical_operators_range_check_their_operands_too() {
        // Same rule as `\\` and `Mod`: the operands must fit a Long.
        assert_eq!(expr("True Or \"2147483648\""), "ERR|6");
        assert_eq!(expr("1 And \"2147483648\""), "ERR|6");
        // Operands that round into a Long are fine.
        assert_eq!(expr("True Or \"3.752147483647\""), "Long|-1");
        assert_eq!(expr("1 And \"12\""), "Long|0");
    }

    #[test]
    fn int_div_and_mod_range_check_their_operands_not_just_the_result() {
        // `254 Mod "22147483647"` is error 6 even though the answer is 254:
        // the operand is not a Long. Checking only the result let it through.
        assert_eq!(
            run("    Dim a, b\n    a = 254\n    b = \"22147483647\"\n    F = (a Mod b)"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 254\n    b = \"22147483647\"\n    F = (a \\ b)"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 3000000000#\n    b = 3\n    F = (a Mod b)"),
            "ERR|6"
        );
        // Operands that do fit a Long still work.
        assert_eq!(
            run("    Dim a, b\n    a = 254\n    b = 2147483647\n    F = (a Mod b)"),
            "Long|254"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 40000\n    b = 3\n    F = (a Mod b)"),
            "Long|1"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 40000\n    b = 3\n    F = (a \\ b)"),
            "Long|13333"
        );
    }

    #[test]
    fn a_negative_base_with_a_fractional_exponent_is_an_error() {
        // Excel raises error 5 rather than returning NaN.
        assert_eq!(expr("(-1) ^ 1.5"), "ERR|5");
        assert_eq!(expr("(-8) ^ (1 / 3)"), "ERR|5");
        // Integral exponents are fine.
        assert_eq!(expr("(-2) ^ 2"), "Double|4");
        assert_eq!(expr("(-2) ^ 3"), "Double|-8");
    }

    #[test]
    fn select_case_matches_a_numeric_case_against_a_string_subject() {
        // `Select Case "10"` matches `Case 10`, but `Select Case ""` does not
        // match `Case 0` -- the numeric-constant rule, not an error.
        let body = |x: &str| {
            format!(
                "    Dim r\n    Select Case {x}\n    Case 0\n        r = \"zero\"\n    \
                     Case 10\n        r = \"ten\"\n    Case Else\n        r = \"else\"\n    \
                     End Select\n    F = r"
            )
        };
        assert_eq!(run(&body("\"10\"")), "String|ten");
        assert_eq!(run(&body("\"\"")), "String|else");
    }

    #[test]
    fn a_for_counter_is_left_at_the_value_that_failed_the_test() {
        assert_eq!(
            run("    Dim c\n    For c = 1 To 3\n    Next c\n    F = c"),
            "Integer|4"
        );
        assert_eq!(
            run("    Dim c\n    For c = 1 To 3 Step 2\n    Next c\n    F = c"),
            "Integer|5"
        );
        // A loop that never runs leaves the counter at its start value.
        assert_eq!(
            run("    Dim c\n    For c = 5 To 1\n    Next c\n    F = c"),
            "Integer|5"
        );
        assert_eq!(
            run("    Dim c\n    For c = 3 To 1 Step -1\n    Next c\n    F = c"),
            "Integer|0"
        );
        // Exit For leaves it at the value the body was running with.
        assert_eq!(
            run("    Dim c\n    For c = 1 To 3\n        Exit For\n    Next c\n    F = c"),
            "Integer|1"
        );
    }

    #[test]
    fn count_arguments_round_rather_than_truncate() {
        // Space(2.6) is three spaces, not two.
        assert_eq!(expr("Len(Space(2.6))"), "Long|3");
        assert_eq!(expr("Space(-1)"), "ERR|5");
        assert_eq!(expr("String(-1, \"x\")"), "ERR|5");
        assert_eq!(expr("Left(\"abc\", -1)"), "ERR|5");
        assert_eq!(expr("Right(\"abc\", 99)"), "String|abc");
        assert_eq!(expr("InStr(0, \"abc\", \"b\")"), "ERR|5");
        assert_eq!(expr("String(2, 65)"), "String|AA");
    }

    // ---- out of scope ---------------------------------------------------

    #[test]
    fn host_object_access_errors_rather_than_silently_doing_nothing() {
        // The refusal that matters: a macro that skips a line it cannot
        // understand and then reports success is wrong in the worst way.
        // These run with no workbook attached, which is what `visi macro run`
        // over a bare `.bas` file does.
        for body in [
            "    F = Range(\"A1\").Value",
            "    F = ThisWorkbook.Name",
            "    F = Worksheets(1).Name",
            "    F = Application.WorksheetFunction.Sum(1, 2)",
            "    Dim c\n    For Each c In r\n    Next",
        ] {
            let out = run(body);
            assert!(out.starts_with("ERR|438"), "{body:?} gave {out}");
        }
    }

    #[test]
    fn a_member_of_a_non_object_is_error_424() {
        // Not 438: the construct *is* supported, the value just is not an
        // object. VBA calls this "Object required", and distinguishing it
        // from "not implemented" is the difference between a macro bug and a
        // gap in this interpreter.
        assert_eq!(run("    With x\n        F = .a\n    End With"), "ERR|424");
        assert_eq!(expr("x.Name"), "ERR|424");
        assert_eq!(expr("x Is Nothing"), "ERR|424");
    }

    #[test]
    fn an_unknown_function_is_reported_not_ignored() {
        assert_eq!(expr("NoSuchFunction(1)"), "ERR|35");
    }
}
