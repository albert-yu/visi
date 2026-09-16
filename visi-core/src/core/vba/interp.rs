use std::collections::HashMap;
use std::rc::Rc;

use super::ast::*;
use super::builtins;
use super::host::{Host, ObjRef};
use super::value::{self, ArithMode, Operand, VResult, Variant, VbaError};
use super::{VbaModule, VbaModuleKind, VbaProject};

const DEFAULT_MAX_OPS: u64 = 5_000_000;
const DEFAULT_MAX_DEPTH: usize = 64;

fn out_of_scope(what: &str) -> VbaError {
    VbaError::new(
        438,
        format!("Object doesn't support this property or method: {what} is not available"),
    )
}

fn needs_workbook(what: &str) -> VbaError {
    VbaError::new(
        438,
        format!(
            "Object doesn't support this property or method: {what} needs a workbook, and this run has none"
        ),
    )
}

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

#[derive(Debug, Clone, PartialEq)]
enum Handler {
    /// No handler: an error propagates out of the procedure.
    None,
    /// `On Error Resume Next`.
    ResumeNext,
    /// `On Error GoTo <label>`.
    Goto(String),
}

/// Procedures sharing a name in a module (e.g. Sub/Function vs Property Get/Let/Set).
#[derive(Debug, Clone, Default)]
pub struct MemberProcs {
    pub sub_or_func: Option<Rc<Procedure>>,
    pub prop_get: Option<Rc<Procedure>>,
    pub prop_let: Option<Rc<Procedure>>,
    pub prop_set: Option<Rc<Procedure>>,
}

impl MemberProcs {
    pub fn insert(&mut self, proc: Rc<Procedure>) {
        match proc.kind {
            ProcKind::Sub | ProcKind::Function => self.sub_or_func = Some(proc),
            ProcKind::PropertyGet => self.prop_get = Some(proc),
            ProcKind::PropertyLet => self.prop_let = Some(proc),
            ProcKind::PropertySet => self.prop_set = Some(proc),
        }
    }

    pub fn first(&self) -> Option<Rc<Procedure>> {
        self.sub_or_func
            .as_ref()
            .or(self.prop_get.as_ref())
            .or(self.prop_let.as_ref())
            .or(self.prop_set.as_ref())
            .cloned()
    }
}

/// A parsed module environment in the VBA project.
#[derive(Debug, Clone)]
pub struct ModuleEnv {
    pub name: String,
    pub kind: VbaModuleKind,
    pub bound_sheet_id: Option<u64>,
    pub procs: HashMap<String, MemberProcs>,
    pub events: HashMap<String, Rc<Stmt>>,
    pub globals: HashMap<String, Variant>,
    pub auto_new_vars: HashMap<String, String>,
    pub with_events_vars: HashMap<String, String>,
    pub default_member: Option<String>,
    pub ast: Module,
}

impl ModuleEnv {
    pub fn new(
        name: String,
        kind: VbaModuleKind,
        bound_sheet_id: Option<u64>,
        ast: Module,
    ) -> Self {
        let mut procs: HashMap<String, MemberProcs> = HashMap::new();
        let mut events: HashMap<String, Rc<Stmt>> = HashMap::new();
        let mut globals: HashMap<String, Variant> = HashMap::new();
        let mut auto_new_vars: HashMap<String, String> = HashMap::new();
        let mut with_events_vars: HashMap<String, String> = HashMap::new();
        let mut default_member: Option<String> = None;

        for item in &ast.items {
            match item {
                ModuleItem::Attribute {
                    name: attr_name,
                    values,
                    ..
                } => {
                    if attr_name.to_ascii_lowercase().ends_with(".vb_usermemid")
                        && let Some(val_expr) = values.first()
                        && is_zero_expr(val_expr)
                        && let Some((member, _)) = attr_name.split_once('.')
                    {
                        default_member = Some(member.to_ascii_lowercase());
                    }
                }
                ModuleItem::Declaration(stmt) => match stmt {
                    Stmt::Dim {
                        vars, with_events, ..
                    } => {
                        for v in vars {
                            let key = v.name.to_ascii_lowercase();
                            if *with_events {
                                let ty_name =
                                    v.ty.as_ref()
                                        .and_then(|t| t.path.last())
                                        .cloned()
                                        .unwrap_or_default();
                                with_events_vars.insert(key.clone(), ty_name);
                                globals.insert(key, Variant::Object(ObjRef::Nothing));
                            } else if v.ty.as_ref().is_some_and(|t| t.is_new) {
                                let cls_name =
                                    v.ty.as_ref()
                                        .unwrap()
                                        .path
                                        .last()
                                        .cloned()
                                        .unwrap_or_default();
                                auto_new_vars.insert(key.clone(), cls_name);
                                globals.insert(key, Variant::Object(ObjRef::Nothing));
                            } else {
                                globals.insert(key, default_for(v.ty.as_ref()));
                            }
                        }
                    }
                    Stmt::Const { vars, .. } => {
                        for v in vars {
                            let key = v.name.to_ascii_lowercase();
                            globals.insert(key, default_for(v.ty.as_ref()));
                        }
                    }
                    Stmt::EventDef {
                        name: event_name, ..
                    } => {
                        events.insert(event_name.to_ascii_lowercase(), Rc::new(stmt.clone()));
                    }
                    _ => {}
                },
                ModuleItem::Procedure(p) => {
                    let key = p.name.to_ascii_lowercase();
                    for s in &p.body {
                        if let Stmt::Attribute {
                            name: attr_name,
                            values,
                            ..
                        } = s
                            && (attr_name.to_ascii_lowercase().ends_with(".vb_usermemid")
                                || attr_name.eq_ignore_ascii_case("vb_usermemid"))
                            && let Some(val_expr) = values.first()
                            && is_zero_expr(val_expr)
                        {
                            default_member = Some(key.clone());
                        }
                    }
                    procs.entry(key).or_default().insert(Rc::new(p.clone()));
                }
                ModuleItem::Conditional {
                    branches,
                    else_items,
                    ..
                } => {
                    for (_, b_items) in branches {
                        for b_item in b_items {
                            if let ModuleItem::Procedure(p) = b_item {
                                procs
                                    .entry(p.name.to_ascii_lowercase())
                                    .or_default()
                                    .insert(Rc::new(p.clone()));
                            }
                        }
                    }
                    if let Some(e_items) = else_items {
                        for e_item in e_items {
                            if let ModuleItem::Procedure(p) = e_item {
                                procs
                                    .entry(p.name.to_ascii_lowercase())
                                    .or_default()
                                    .insert(Rc::new(p.clone()));
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        Self {
            name,
            kind,
            bound_sheet_id,
            procs,
            events,
            globals,
            auto_new_vars,
            with_events_vars,
            default_member,
            ast,
        }
    }
}

fn is_zero_expr(e: &Expr) -> bool {
    match e {
        Expr::Literal(Literal::Number { value, .. }) => *value == 0.0,
        Expr::Literal(Literal::Str(s)) => s == "0",
        _ => false,
    }
}

#[derive(Debug, Clone)]
pub struct EventSubscription {
    pub event_name: String,
    pub listener_module: String,
    pub listener_var_name: String,
    pub listener_instance: Option<u64>,
}

#[derive(Debug, Clone)]
pub struct UserClassInstance {
    pub id: u64,
    pub class_name: String,
    pub fields: HashMap<String, Variant>,
    pub auto_new_fields: HashMap<String, String>,
    pub with_events_fields: HashMap<String, String>,
    pub event_sinks: Vec<EventSubscription>,
    pub ref_count: usize,
    pub terminating: bool,
}

struct Frame {
    locals: HashMap<String, Variant>,
    auto_new_locals: HashMap<String, String>,
    with_events_locals: HashMap<String, String>,
    handler: Handler,
    in_handler: bool,
    failed_at: Option<usize>,
    with_stack: Vec<Variant>,
    me: Option<ObjRef>,
    module_name: String,
}

impl Frame {
    fn new() -> Self {
        Self {
            locals: HashMap::new(),
            auto_new_locals: HashMap::new(),
            with_events_locals: HashMap::new(),
            handler: Handler::None,
            in_handler: false,
            failed_at: None,
            with_stack: Vec::new(),
            me: None,
            module_name: String::new(),
        }
    }
}

#[derive(Debug, Clone, Default)]
struct ErrState {
    number: i32,
    description: String,
}

/// Runs VBA procedures across single-module or multi-module projects.
pub struct Interpreter<'w> {
    modules: HashMap<String, ModuleEnv>,
    instances: HashMap<u64, UserClassInstance>,
    next_instance_id: u64,
    active_module: String,
    err: ErrState,
    ops: u64,
    max_ops: u64,
    depth: usize,
    max_depth: usize,
    host: Option<Host<'w>>,
    event_depth: usize,
}

impl<'w> Interpreter<'w> {
    /// Builds an interpreter over a single parsed module.
    pub fn new(module: Module) -> Self {
        let mut name = "Module1".to_string();
        for item in &module.items {
            if let ModuleItem::Attribute {
                name: attr_name,
                values,
                ..
            } = item
                && attr_name.eq_ignore_ascii_case("vb_name")
                && let Some(Expr::Literal(Literal::Str(n))) = values.first()
            {
                name = n.clone();
            }
        }
        let env = ModuleEnv::new(name.clone(), VbaModuleKind::Standard, None, module);
        let mut modules = HashMap::new();
        modules.insert(name.to_ascii_lowercase(), env);
        Self {
            modules,
            instances: HashMap::new(),
            next_instance_id: 1,
            active_module: name.to_ascii_lowercase(),
            err: ErrState::default(),
            ops: 0,
            max_ops: DEFAULT_MAX_OPS,
            depth: 0,
            max_depth: DEFAULT_MAX_DEPTH,
            host: None,
            event_depth: 0,
        }
    }

    /// Builds an interpreter from a list of project modules.
    pub fn from_modules(modules_list: Vec<VbaModule>, target_module: Option<&str>) -> Self {
        let mut modules = HashMap::new();
        let mut default_active = String::new();

        for m in modules_list {
            let parsed = super::parser::parse_module(&m.source)
                .unwrap_or_else(|_| Module { items: Vec::new() });
            let env = ModuleEnv::new(m.name.clone(), m.kind, m.bound_sheet_id, parsed);
            let lower = m.name.to_ascii_lowercase();
            if default_active.is_empty() || m.kind == VbaModuleKind::Standard {
                default_active = lower.clone();
            }
            modules.insert(lower, env);
        }

        let active_module = target_module
            .map(|t| t.to_ascii_lowercase())
            .unwrap_or(default_active);

        Self {
            modules,
            instances: HashMap::new(),
            next_instance_id: 1,
            active_module,
            err: ErrState::default(),
            ops: 0,
            max_ops: DEFAULT_MAX_OPS,
            depth: 0,
            max_depth: DEFAULT_MAX_DEPTH,
            host: None,
            event_depth: 0,
        }
    }

    /// Builds an interpreter from a `VbaProject`.
    pub fn from_project(project: &VbaProject, target_module: Option<&str>) -> VResult<Self> {
        let mut modules = HashMap::new();
        let mut default_active = String::new();

        for m in &project.modules {
            let parsed = super::parser::parse_module(&m.source).map_err(|e| {
                VbaError::new(
                    13,
                    format!("Syntax error in module {}: {}", m.name, e.message),
                )
            })?;
            let env = ModuleEnv::new(m.name.clone(), m.kind, m.bound_sheet_id, parsed);
            let lower = m.name.to_ascii_lowercase();
            if default_active.is_empty() || m.kind == VbaModuleKind::Standard {
                default_active = lower.clone();
            }
            modules.insert(lower, env);
        }

        let active_module = if let Some(t) = target_module {
            let lower_t = t.to_ascii_lowercase();
            if !modules.contains_key(&lower_t) {
                return Err(VbaError::new(35, format!("Module not found: {t}")));
            }
            lower_t
        } else {
            default_active
        };

        Ok(Self {
            modules,
            instances: HashMap::new(),
            next_instance_id: 1,
            active_module,
            err: ErrState::default(),
            ops: 0,
            max_ops: DEFAULT_MAX_OPS,
            depth: 0,
            max_depth: DEFAULT_MAX_DEPTH,
            host: None,
            event_depth: 0,
        })
    }

    /// Adds a parsed module to this interpreter.
    pub fn add_module(
        &mut self,
        name: &str,
        kind: VbaModuleKind,
        bound_sheet_id: Option<u64>,
        ast: Module,
    ) {
        let env = ModuleEnv::new(name.to_string(), kind, bound_sheet_id, ast);
        self.modules.insert(name.to_ascii_lowercase(), env);
    }

    /// Binds a workbook, enabling the host object model.
    pub fn with_host(mut self, host: Host<'w>) -> Self {
        self.host = Some(host);
        self
    }

    /// Whether the run changed the workbook.
    pub fn mutated(&self) -> bool {
        self.host.as_ref().is_some_and(|h| h.mutated())
    }

    /// Settles any outstanding recalculation.
    pub fn finish(&mut self) {
        if let Some(h) = self.host.as_mut() {
            h.finish();
        }
        let mut all_globals = Vec::new();
        for m in self.modules.values_mut() {
            for v in m.globals.values() {
                all_globals.push(v.clone());
            }
        }
        for g in all_globals {
            self.dec_ref(&g);
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

    /// Whether events are currently enabled.
    pub fn enable_events(&self) -> bool {
        self.host.as_ref().is_none_or(|h| h.enable_events)
    }

    /// Returns the resolved type name for a variant (e.g. "Class1" for UserClass).
    pub fn type_name_of(&self, v: &Variant) -> String {
        match v {
            Variant::Object(ObjRef::UserClass(id)) => {
                if let Some(inst) = self.instances.get(id) {
                    inst.class_name.clone()
                } else {
                    "Object".to_string()
                }
            }
            Variant::Object(ObjRef::Nothing) => "Nothing".to_string(),
            _ => v.type_name().to_string(),
        }
    }

    pub fn class_name_of(&self, id: u64) -> String {
        self.instances
            .get(&id)
            .map(|i| i.class_name.clone())
            .unwrap_or_else(|| "Object".to_string())
    }

    /// Runs the named procedure and returns its value.
    pub fn run(&mut self, name: &str, args: Vec<Variant>) -> VResult<Variant> {
        self.ops = 0;
        self.init_all_modules()?;
        let res = self.call_procedure(name, args)?;
        self.drain_and_fire_events()?;
        Ok(res)
    }

    /// Runs startup macro events (`Workbook_Open` in `ThisWorkbook` then `Auto_Open` in standard modules).
    pub fn run_open_events(&mut self) -> VResult<()> {
        self.ops = 0;
        self.init_all_modules()?;

        if let Some(m_env) = self.modules.get("thisworkbook")
            && let Some(mp) = m_env.procs.get("workbook_open")
            && let Some(sub) = mp.first()
        {
            let mut frame = Frame::new();
            frame.module_name = "thisworkbook".to_string();
            frame.me = Some(ObjRef::Workbook);
            let old_active = self.active_module.clone();
            self.active_module = "thisworkbook".to_string();
            let res = self.call_body_with_frame(&sub, Vec::new(), &mut frame);
            self.active_module = old_active;
            res?;
        }

        let std_mods: Vec<String> = self
            .modules
            .values()
            .filter(|m| m.kind == VbaModuleKind::Standard)
            .map(|m| m.name.clone())
            .collect();

        for mod_name in std_mods {
            if let Some(m_env) = self.modules.get(&mod_name.to_ascii_lowercase())
                && let Some(mp) = m_env.procs.get("auto_open")
                && let Some(sub) = mp.first()
            {
                let mut frame = Frame::new();
                frame.module_name = mod_name.clone();
                let old_active = self.active_module.clone();
                self.active_module = mod_name.to_ascii_lowercase();
                let res = self.call_body_with_frame(&sub, Vec::new(), &mut frame);
                self.active_module = old_active;
                res?;
            }
        }

        self.drain_and_fire_events()?;
        Ok(())
    }

    /// Fires `Workbook_BeforeClose` event. Returns true if canceled.
    pub fn fire_workbook_before_close(&mut self) -> VResult<bool> {
        if !self.enable_events() {
            return Ok(false);
        }
        if let Some(m_env) = self.modules.get("thisworkbook")
            && let Some(mp) = m_env.procs.get("workbook_beforeclose")
            && let Some(sub) = mp.first()
        {
            let mut frame = Frame::new();
            frame.module_name = "thisworkbook".to_string();
            frame.me = Some(ObjRef::Workbook);
            let param_name = sub
                .params
                .first()
                .map(|p| p.name.to_ascii_lowercase())
                .unwrap_or_else(|| "cancel".to_string());
            frame
                .locals
                .insert(param_name.clone(), Variant::Boolean(false));
            let old_active = self.active_module.clone();
            self.active_module = "thisworkbook".to_string();
            self.exec_procedure_body(&sub.body, &mut frame)?;
            self.active_module = old_active;
            let canceled = frame
                .locals
                .get(&param_name)
                .is_some_and(|v| v.to_bool().unwrap_or(false));
            return Ok(canceled);
        }
        Ok(false)
    }

    /// Fires `Workbook_BeforeSave` event. Returns true if canceled.
    pub fn fire_workbook_before_save(&mut self, save_as_ui: bool) -> VResult<bool> {
        if !self.enable_events() {
            return Ok(false);
        }
        if let Some(m_env) = self.modules.get("thisworkbook")
            && let Some(mp) = m_env.procs.get("workbook_beforesave")
            && let Some(sub) = mp.first()
        {
            let mut frame = Frame::new();
            frame.module_name = "thisworkbook".to_string();
            frame.me = Some(ObjRef::Workbook);
            if let Some(p1) = sub.params.first() {
                frame
                    .locals
                    .insert(p1.name.to_ascii_lowercase(), Variant::Boolean(save_as_ui));
            }
            let cancel_name = sub
                .params
                .get(1)
                .map(|p| p.name.to_ascii_lowercase())
                .unwrap_or_else(|| "cancel".to_string());
            frame
                .locals
                .insert(cancel_name.clone(), Variant::Boolean(false));
            let old_active = self.active_module.clone();
            self.active_module = "thisworkbook".to_string();
            self.exec_procedure_body(&sub.body, &mut frame)?;
            self.active_module = old_active;
            let canceled = frame
                .locals
                .get(&cancel_name)
                .is_some_and(|v| v.to_bool().unwrap_or(false));
            return Ok(canceled);
        }
        Ok(false)
    }

    fn init_all_modules(&mut self) -> VResult<()> {
        let mod_names: Vec<String> = self.modules.keys().cloned().collect();
        for mod_name in mod_names {
            let items = if let Some(m) = self.modules.get(&mod_name) {
                m.ast.items.clone()
            } else {
                continue;
            };
            for item in &items {
                if let ModuleItem::Declaration(stmt) = item {
                    let mut frame = Frame::new();
                    frame.module_name = mod_name.clone();
                    let old_active = self.active_module.clone();
                    self.active_module = mod_name.clone();
                    let r = self.exec_stmt(stmt, &mut frame, true);
                    self.active_module = old_active;
                    if let Some(m) = self.modules.get_mut(&mod_name) {
                        for (k, v) in frame.locals {
                            m.globals.insert(k, v);
                        }
                    }
                    r?;
                }
            }
        }
        Ok(())
    }

    fn find_class_module(&self, class_name: &str) -> Option<&ModuleEnv> {
        let lower = class_name.to_ascii_lowercase();
        self.modules
            .get(&lower)
            .filter(|m| m.kind == VbaModuleKind::Class)
    }

    fn find_document_module_name_by_sheet_id(&self, sheet_id: u64) -> Option<String> {
        self.modules
            .values()
            .find(|m| m.kind == VbaModuleKind::Document && m.bound_sheet_id == Some(sheet_id))
            .map(|m| m.name.to_ascii_lowercase())
    }

    pub fn instantiate_class(&mut self, class_name: &str) -> VResult<Variant> {
        let module = self.find_class_module(class_name).cloned().ok_or_else(|| {
            VbaError::new(
                424,
                format!("Object required: Class '{class_name}' not defined"),
            )
        })?;

        let id = self.next_instance_id;
        self.next_instance_id += 1;

        let mut fields = HashMap::new();
        let mut auto_new_fields = HashMap::new();
        let mut with_events_fields = HashMap::new();

        for (name, val) in &module.globals {
            fields.insert(name.clone(), val.clone());
        }
        for (name, cls) in &module.auto_new_vars {
            auto_new_fields.insert(name.clone(), cls.clone());
        }
        for (name, cls) in &module.with_events_vars {
            with_events_fields.insert(name.clone(), cls.clone());
        }

        let instance = UserClassInstance {
            id,
            class_name: module.name.clone(),
            fields,
            auto_new_fields,
            with_events_fields,
            event_sinks: Vec::new(),
            ref_count: 1,
            terminating: false,
        };
        self.instances.insert(id, instance);

        let init_proc = module
            .procs
            .get("class_initialize")
            .and_then(|mp| mp.first());
        if let Some(proc) = init_proc {
            let mut frame = Frame::new();
            frame.module_name = module.name.to_ascii_lowercase();
            frame.me = Some(ObjRef::UserClass(id));
            let old_active = self.active_module.clone();
            self.active_module = module.name.to_ascii_lowercase();
            let res = self.call_body_with_frame(&proc, Vec::new(), &mut frame);
            self.active_module = old_active;
            res?;
        }

        Ok(Variant::Object(ObjRef::UserClass(id)))
    }

    fn inc_ref(&mut self, val: &Variant) {
        if let Variant::Object(ObjRef::UserClass(id)) = val
            && let Some(inst) = self.instances.get_mut(id)
        {
            inst.ref_count += 1;
        }
    }

    fn dec_ref(&mut self, val: &Variant) {
        if let Variant::Object(ObjRef::UserClass(id)) = val {
            let mut terminate_proc = None;
            let mut class_mod_name = String::new();
            let mut fields_to_dec = Vec::new();

            if let Some(inst) = self.instances.get_mut(id) {
                if inst.ref_count > 0 {
                    inst.ref_count -= 1;
                }
                if inst.ref_count == 0 && !inst.terminating {
                    inst.terminating = true;
                    class_mod_name = inst.class_name.to_ascii_lowercase();
                    if let Some(m_env) = self.modules.get(&class_mod_name)
                        && let Some(mp) = m_env.procs.get("class_terminate")
                    {
                        terminate_proc = mp.first();
                    }
                    fields_to_dec = inst.fields.values().cloned().collect();
                }
            }

            if let Some(proc) = terminate_proc {
                let mut frame = Frame::new();
                frame.module_name = class_mod_name.clone();
                frame.me = Some(ObjRef::UserClass(*id));
                let old_active = self.active_module.clone();
                self.active_module = class_mod_name;
                let _ = self.call_body_with_frame(&proc, Vec::new(), &mut frame);
                self.active_module = old_active;
            }

            if let Some(inst) = self.instances.get(id)
                && inst.terminating
                && inst.ref_count == 0
            {
                for f in fields_to_dec {
                    self.dec_ref(&f);
                }
                self.instances.remove(id);
            }
        }
    }

    fn add_event_subscriptions(
        &mut self,
        target_inst_id: u64,
        listener_module: &str,
        listener_var_name: &str,
        listener_inst: Option<ObjRef>,
    ) {
        let listener_inst_id = match listener_inst {
            Some(ObjRef::UserClass(lid)) => Some(lid),
            _ => None,
        };
        let class_name = if let Some(inst) = self.instances.get(&target_inst_id) {
            inst.class_name.to_ascii_lowercase()
        } else {
            return;
        };

        let events: Vec<String> = if let Some(m_env) = self.modules.get(&class_name) {
            m_env.events.keys().cloned().collect()
        } else {
            Vec::new()
        };

        if let Some(inst) = self.instances.get_mut(&target_inst_id) {
            for evt in events {
                inst.event_sinks.push(EventSubscription {
                    event_name: evt,
                    listener_module: listener_module.to_string(),
                    listener_var_name: listener_var_name.to_string(),
                    listener_instance: listener_inst_id,
                });
            }
        }
    }

    fn remove_event_subscriptions(
        &mut self,
        target_inst_id: u64,
        listener_module: &str,
        listener_var_name: &str,
    ) {
        if let Some(inst) = self.instances.get_mut(&target_inst_id) {
            inst.event_sinks.retain(|s| {
                !(s.listener_module.eq_ignore_ascii_case(listener_module)
                    && s.listener_var_name.eq_ignore_ascii_case(listener_var_name))
            });
        }
    }

    fn drain_and_fire_events(&mut self) -> VResult<()> {
        if !self.enable_events() {
            if let Some(h) = self.host.as_mut() {
                h.pending_cell_changes.clear();
                h.pending_calculate_sheets.clear();
            }
            return Ok(());
        }

        self.event_depth += 1;
        if self.event_depth > self.max_depth {
            self.event_depth -= 1;
            return Err(VbaError::new(28, "Out of stack space"));
        }

        while let Some(change) = self
            .host
            .as_mut()
            .and_then(|h| h.pending_cell_changes.pop())
        {
            let sheet_id = change.sheet_id;
            let token = self.host("event range")?.new_range(
                change.sheet_id,
                change.row,
                change.col,
                change.height,
                change.width,
            );

            if let Some(sheet_mod_name) = self.find_document_module_name_by_sheet_id(sheet_id)
                && let Some(m_env) = self.modules.get(&sheet_mod_name)
                && let Some(mp) = m_env.procs.get("worksheet_change")
                && let Some(sub) = mp.first()
            {
                let mut frame = Frame::new();
                frame.module_name = sheet_mod_name.clone();
                frame.me = Some(ObjRef::Worksheet(sheet_id));
                let old_active = self.active_module.clone();
                self.active_module = sheet_mod_name;
                let res = self.call_body_with_frame(&sub, vec![Variant::Object(token)], &mut frame);
                self.active_module = old_active;
                res?;
            }

            if let Some(m_env) = self.modules.get("thisworkbook")
                && let Some(mp) = m_env.procs.get("workbook_sheetchange")
                && let Some(sub) = mp.first()
            {
                let mut frame = Frame::new();
                frame.module_name = "thisworkbook".to_string();
                frame.me = Some(ObjRef::Workbook);
                let old_active = self.active_module.clone();
                self.active_module = "thisworkbook".to_string();
                let res = self.call_body_with_frame(
                    &sub,
                    vec![
                        Variant::Object(ObjRef::Worksheet(sheet_id)),
                        Variant::Object(token),
                    ],
                    &mut frame,
                );
                self.active_module = old_active;
                res?;
            }
        }

        while let Some(sheet_id) = self
            .host
            .as_mut()
            .and_then(|h| h.pending_calculate_sheets.pop())
        {
            if let Some(sheet_mod_name) = self.find_document_module_name_by_sheet_id(sheet_id)
                && let Some(m_env) = self.modules.get(&sheet_mod_name)
                && let Some(mp) = m_env.procs.get("worksheet_calculate")
                && let Some(sub) = mp.first()
            {
                let mut frame = Frame::new();
                frame.module_name = sheet_mod_name.clone();
                frame.me = Some(ObjRef::Worksheet(sheet_id));
                let old_active = self.active_module.clone();
                self.active_module = sheet_mod_name;
                let res = self.call_body_with_frame(&sub, Vec::new(), &mut frame);
                self.active_module = old_active;
                res?;
            }

            if let Some(m_env) = self.modules.get("thisworkbook")
                && let Some(mp) = m_env.procs.get("workbook_sheetcalculate")
                && let Some(sub) = mp.first()
            {
                let mut frame = Frame::new();
                frame.module_name = "thisworkbook".to_string();
                frame.me = Some(ObjRef::Workbook);
                let old_active = self.active_module.clone();
                self.active_module = "thisworkbook".to_string();
                let res = self.call_body_with_frame(
                    &sub,
                    vec![Variant::Object(ObjRef::Worksheet(sheet_id))],
                    &mut frame,
                );
                self.active_module = old_active;
                res?;
            }
        }

        self.event_depth -= 1;
        Ok(())
    }

    fn find_procedure_in_scope(
        &self,
        name: &str,
        frame: &Frame,
    ) -> Option<(String, Rc<Procedure>)> {
        let lower = name.to_ascii_lowercase();

        if let Some(me_obj) = frame.me {
            match me_obj {
                ObjRef::UserClass(id) => {
                    if let Some(inst) = self.instances.get(&id) {
                        let cls = inst.class_name.to_ascii_lowercase();
                        if let Some(m_env) = self.modules.get(&cls)
                            && let Some(mp) = m_env.procs.get(&lower)
                            && let Some(p) = mp.first()
                        {
                            return Some((cls, p));
                        }
                    }
                }
                ObjRef::Worksheet(sheet_id) => {
                    if let Some(doc_name) = self.find_document_module_name_by_sheet_id(sheet_id)
                        && let Some(m_env) = self.modules.get(&doc_name)
                        && let Some(mp) = m_env.procs.get(&lower)
                        && let Some(p) = mp.first()
                    {
                        return Some((doc_name, p));
                    }
                }
                ObjRef::Workbook => {
                    if let Some(m_env) = self.modules.get("thisworkbook")
                        && let Some(mp) = m_env.procs.get(&lower)
                        && let Some(p) = mp.first()
                    {
                        return Some(("thisworkbook".to_string(), p));
                    }
                }
                _ => {}
            }
        }

        if let Some(m_env) = self.modules.get(&self.active_module)
            && let Some(mp) = m_env.procs.get(&lower)
            && let Some(p) = mp.first()
        {
            return Some((self.active_module.clone(), p));
        }

        for (m_name, m_env) in &self.modules {
            if m_env.kind == VbaModuleKind::Standard
                && *m_name != self.active_module
                && let Some(mp) = m_env.procs.get(&lower)
                && let Some(p) = mp.first()
            {
                return Some((m_name.clone(), p));
            }
        }

        None
    }

    fn call_procedure(&mut self, name: &str, args: Vec<Variant>) -> VResult<Variant> {
        let lower = name.to_ascii_lowercase();
        let mut target_mod = None;
        let mut target_proc = None;

        if let Some(m_env) = self.modules.get(&self.active_module)
            && let Some(mp) = m_env.procs.get(&lower)
            && let Some(p) = mp.first()
        {
            target_mod = Some(self.active_module.clone());
            target_proc = Some(p);
        }

        if target_proc.is_none() {
            for (m_name, m_env) in &self.modules {
                if m_env.kind == VbaModuleKind::Standard
                    && let Some(mp) = m_env.procs.get(&lower)
                    && let Some(p) = mp.first()
                {
                    target_mod = Some(m_name.clone());
                    target_proc = Some(p);
                    break;
                }
            }
        }

        if target_proc.is_none() {
            for (m_name, m_env) in &self.modules {
                if let Some(mp) = m_env.procs.get(&lower)
                    && let Some(p) = mp.first()
                {
                    target_mod = Some(m_name.clone());
                    target_proc = Some(p);
                    break;
                }
            }
        }

        let (mod_name, proc) = match (target_mod, target_proc) {
            (Some(m), Some(p)) => (m, p),
            _ => {
                return Err(VbaError::new(
                    35,
                    format!("Sub or Function not defined: {name}"),
                ));
            }
        };

        let mut frame = Frame::new();
        frame.module_name = mod_name.clone();
        if let Some(m_env) = self.modules.get(&mod_name)
            && m_env.kind == VbaModuleKind::Document
        {
            if let Some(sheet_id) = m_env.bound_sheet_id {
                frame.me = Some(ObjRef::Worksheet(sheet_id));
            } else {
                frame.me = Some(ObjRef::Workbook);
            }
        }

        let old_active = self.active_module.clone();
        self.active_module = mod_name;
        let result = self.call_body_with_frame(&proc, args, &mut frame);
        self.active_module = old_active;
        result
    }

    fn call_body_with_frame(
        &mut self,
        proc: &Procedure,
        args: Vec<Variant>,
        frame: &mut Frame,
    ) -> VResult<Variant> {
        self.depth += 1;
        if self.depth > self.max_depth {
            self.depth -= 1;
            return Err(VbaError::new(28, "Out of stack space"));
        }

        for (i, param) in proc.params.iter().enumerate() {
            let value = args.get(i).cloned().unwrap_or(Variant::Empty);
            let key = param.name.to_ascii_lowercase();
            self.inc_ref(&value);
            if param.ty.as_ref().is_some_and(|t| t.is_new) {
                let cls = param
                    .ty
                    .as_ref()
                    .unwrap()
                    .path
                    .last()
                    .cloned()
                    .unwrap_or_default();
                frame.auto_new_locals.insert(key.clone(), cls);
            }
            frame.locals.insert(key, value);
        }

        let ret_key = proc.name.to_ascii_lowercase();
        if proc.kind != ProcKind::Sub {
            frame
                .locals
                .entry(ret_key.clone())
                .or_insert(Variant::Empty);
        }

        let body_res = self.exec_procedure_body(&proc.body, frame);

        let ret = if proc.kind == ProcKind::Sub {
            Variant::Empty
        } else {
            frame
                .locals
                .get(&ret_key)
                .cloned()
                .unwrap_or(Variant::Empty)
        };

        let locals = std::mem::take(&mut frame.locals);
        for (_, val) in locals {
            self.dec_ref(&val);
        }

        self.depth -= 1;
        body_res?;
        Ok(ret)
    }

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

    fn exec_block(&mut self, body: &[Stmt], frame: &mut Frame) -> VResult<Flow> {
        for stmt in body {
            match self.exec_stmt(stmt, frame, false) {
                Ok(Flow::Normal) => {}
                Ok(other) => return Ok(other),
                Err(e) => match self.take_handler(frame) {
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

            Stmt::Dim {
                vars, with_events, ..
            } => {
                for v in vars {
                    let key = v.name.to_ascii_lowercase();
                    if *with_events {
                        let ty_name =
                            v.ty.as_ref()
                                .and_then(|t| t.path.last())
                                .cloned()
                                .unwrap_or_default();
                        frame.with_events_locals.insert(key.clone(), ty_name);
                        frame.locals.insert(key, Variant::Object(ObjRef::Nothing));
                    } else if v.ty.as_ref().is_some_and(|t| t.is_new) {
                        let cls =
                            v.ty.as_ref()
                                .unwrap()
                                .path
                                .last()
                                .cloned()
                                .unwrap_or_default();
                        frame.auto_new_locals.insert(key.clone(), cls);
                        frame.locals.insert(key, Variant::Object(ObjRef::Nothing));
                    } else {
                        let initial = default_for(v.ty.as_ref());
                        frame.locals.insert(key, initial);
                    }
                }
                Ok(Flow::Normal)
            }

            Stmt::Const { vars, .. } => {
                for v in vars {
                    let value = match &v.value {
                        Some(e) => self.eval(e, frame)?,
                        None => Variant::Empty,
                    };
                    self.inc_ref(&value);
                    frame.locals.insert(v.name.to_ascii_lowercase(), value);
                }
                Ok(Flow::Normal)
            }

            Stmt::Assign {
                target, value, set, ..
            } => {
                let v = self.eval(value, frame)?;
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
                    if self.scalar(c)?.to_bool_condition()? {
                        return self.exec_block(body, frame);
                    }
                }
                if let Some(body) = else_body {
                    return self.exec_block(body, frame);
                }
                Ok(Flow::Normal)
            }

            Stmt::SelectCase {
                subject,
                cases,
                case_else,
                ..
            } => {
                let s = self.eval(subject, frame)?;
                let s = self.scalar(s)?;
                let text_compare = matches!(s, Variant::Str(_)) && is_constant(subject);
                let bool_compare =
                    matches!(s, Variant::Boolean(_)) && is_statically_boolean(subject);
                for clause in cases {
                    for m in &clause.matches {
                        if self.case_matches(&s, m, frame, text_compare, bool_compare)? {
                            return self.exec_block(&clause.body, frame);
                        }
                    }
                }
                if let Some(body) = case_else {
                    return self.exec_block(body, frame);
                }
                Ok(Flow::Normal)
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
                frame.in_handler = false;
                Ok(Flow::Normal)
            }

            Stmt::Resume { kind, .. } => {
                frame.in_handler = false;
                Ok(match kind {
                    ResumeKind::Label(label) => Flow::Goto(label.clone()),
                    ResumeKind::Next => Flow::Goto("\0resume-next".to_string()),
                    ResumeKind::Retry => Flow::Goto("\0resume-next".to_string()),
                })
            }

            Stmt::Stop { .. } | Stmt::End { .. } => Ok(Flow::ExitProc),

            Stmt::EventDef { .. } => Ok(Flow::Normal),

            Stmt::RaiseEvent { name, args, .. } => {
                let mut arg_vals = self.eval_args(args, frame)?;
                let Some(ObjRef::UserClass(inst_id)) = frame.me else {
                    return Err(VbaError::new(
                        438,
                        "RaiseEvent must be called within a class instance",
                    ));
                };

                let sinks = self
                    .instances
                    .get(&inst_id)
                    .map(|inst| inst.event_sinks.clone())
                    .unwrap_or_default();
                let lower_event = name.to_ascii_lowercase();

                for sink in sinks {
                    if sink.event_name == lower_event {
                        let handler_name = format!("{}_{}", sink.listener_var_name, name);
                        if let Some(m_env) = self.modules.get(&sink.listener_module)
                            && let Some(mp) = m_env.procs.get(&handler_name.to_ascii_lowercase())
                            && let Some(proc) = mp.first()
                        {
                            let mut handler_frame = Frame::new();
                            handler_frame.module_name = sink.listener_module.clone();
                            handler_frame.me = sink.listener_instance.map(ObjRef::UserClass);
                            let old_active = self.active_module.clone();
                            self.active_module = sink.listener_module.clone();

                            for (i, p) in proc.params.iter().enumerate() {
                                let v = arg_vals.get(i).cloned().unwrap_or(Variant::Empty);
                                handler_frame.locals.insert(p.name.to_ascii_lowercase(), v);
                            }

                            self.exec_procedure_body(&proc.body, &mut handler_frame)?;

                            for (i, p) in proc.params.iter().enumerate() {
                                if p.by != Some(PassBy::Value)
                                    && let Some(new_val) =
                                        handler_frame.locals.get(&p.name.to_ascii_lowercase())
                                    && i < arg_vals.len()
                                {
                                    arg_vals[i] = new_val.clone();
                                }
                            }

                            self.active_module = old_active;
                        }
                    }
                }

                for (i, a) in args.iter().enumerate() {
                    if let Some(Expr::Ident { name: arg_var, .. }) = &a.value
                        && i < arg_vals.len()
                    {
                        let pos = a.value.as_ref().unwrap().pos();
                        self.assign_with(
                            &Expr::Ident {
                                name: arg_var.clone(),
                                pos,
                            },
                            arg_vals[i].clone(),
                            frame,
                            false,
                            false,
                        )?;
                    }
                }

                Ok(Flow::Normal)
            }

            Stmt::Attribute { .. } => Ok(Flow::Normal),

            Stmt::ReDim { .. } => Err(out_of_scope("ReDim")),
            Stmt::Erase { .. } => Err(out_of_scope("Erase")),
            Stmt::GoSub { .. } | Stmt::Return { .. } => Err(out_of_scope("GoSub")),
            Stmt::OnGoto { .. } => Err(out_of_scope("On ... GoTo")),
            Stmt::TypeDef { .. } => Err(out_of_scope("Type")),
            Stmt::EnumDef { .. } => Err(out_of_scope("Enum")),
            Stmt::Declare { .. } => Err(out_of_scope("Declare")),
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
        let cmp =
            |lhs: &Variant, rhs: &Variant, kind: Operand| -> VResult<Option<std::cmp::Ordering>> {
                if text_compare {
                    return Ok(Some(lhs.to_vba_string()?.cmp(&rhs.to_vba_string()?)));
                }
                value::compare_ctx(lhs, rhs, Operand::Runtime, kind)
            };
        let cast = |v: Variant| -> VResult<Variant> {
            if bool_compare {
                return Ok(Variant::Boolean(v.to_bool()?));
            }
            Ok(v)
        };
        Ok(match m {
            CaseMatch::Value(e) => {
                let v = cast(self.eval(e, frame)?)?;
                cmp(subject, &v, operand_kind(e))? == Some(std::cmp::Ordering::Equal)
            }
            CaseMatch::Range(lo_e, hi_e) => {
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
        if step_v == 0.0 {
            return Err(VbaError::new(
                5,
                "Invalid procedure call or argument: For step is 0",
            ));
        }

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
                let c = self.eval(cond, frame)?;
                let c = self.scalar(c)?.to_bool_condition()?;
                let stop = match test {
                    DoTest::While => !c,
                    DoTest::Until => c,
                };
                if stop {
                    break;
                }
            }
            match self.exec_block(body, frame)? {
                Flow::Normal => {}
                Flow::ExitDo => break,
                other => return Ok(other),
            }
            if let Some((test, cond)) = post {
                let c = self.eval(cond, frame)?;
                let c = self.scalar(c)?.to_bool_condition()?;
                let stop = match test {
                    DoTest::While => !c,
                    DoTest::Until => c,
                };
                if stop {
                    break;
                }
            }
        }
        Ok(Flow::Normal)
    }

    fn assign(
        &mut self,
        target: &Expr,
        value: Variant,
        frame: &mut Frame,
        set: bool,
    ) -> VResult<()> {
        self.assign_with(target, value, frame, false, set)
    }

    fn assign_with(
        &mut self,
        target: &Expr,
        v: Variant,
        frame: &mut Frame,
        module_level: bool,
        set: bool,
    ) -> VResult<()> {
        match target {
            Expr::Member {
                target: obj, name, ..
            } => {
                let owner = self.member_owner(obj.as_deref(), frame)?;
                let Variant::Object(owner_obj) = owner else {
                    return Err(VbaError::new(
                        424,
                        format!("Object required: .{name} on a {}", owner.type_name()),
                    ));
                };
                return self.set_member_on_object(&owner_obj, name, &[], &v, set);
            }
            Expr::Call {
                target: t, args, ..
            } => {
                if let Expr::Member {
                    target: inner_obj,
                    name: prop_name,
                    ..
                } = t.as_ref()
                {
                    let owner = self.member_owner(inner_obj.as_deref(), frame)?;
                    if let Variant::Object(owner_obj) = &owner
                        && matches!(owner_obj, ObjRef::UserClass(_))
                    {
                        let arg_vals = self.eval_args(args, frame)?;
                        return self.set_member_on_object(owner_obj, prop_name, &arg_vals, &v, set);
                    }
                }
                if let Expr::Ident {
                    name: ident_name, ..
                } = t.as_ref()
                    && let Some(Variant::Object(obj)) = self.lookup(ident_name, frame)
                    && let ObjRef::UserClass(id) = obj
                {
                    let arg_vals = self.eval_args(args, frame)?;
                    let cls_name = self
                        .instances
                        .get(&id)
                        .map(|inst| inst.class_name.clone())
                        .unwrap_or_default();
                    if let Some(def_member) = self
                        .modules
                        .get(&cls_name.to_ascii_lowercase())
                        .and_then(|m| m.default_member.clone())
                    {
                        return self.set_member_on_object(&obj, &def_member, &arg_vals, &v, set);
                    }
                }
                if !set {
                    let obj = self.eval(target, frame);
                    if let Ok(Variant::Object(obj)) = obj {
                        self.host("assignment to an object")?
                            .assign_default(&obj, &v)?;
                        self.drain_and_fire_events()?;
                        return Ok(());
                    }
                }
                return Err(out_of_scope("array or property assignment"));
            }
            Expr::Bang {
                target: obj_expr,
                name,
                ..
            } => {
                let owner = self.eval(obj_expr, frame)?;
                let Variant::Object(owner_obj) = owner else {
                    return Err(VbaError::new(
                        424,
                        format!("Object required: !{name} on a {}", owner.type_name()),
                    ));
                };
                return self.set_member_on_object(&owner_obj, name, &[], &v, set);
            }
            _ => {}
        }

        match target {
            Expr::Ident { name, .. } => {
                let key = name.to_ascii_lowercase();
                self.inc_ref(&v);

                let is_with_events = frame.with_events_locals.contains_key(&key)
                    || self
                        .modules
                        .get(&self.active_module)
                        .is_some_and(|m| m.with_events_vars.contains_key(&key));

                if is_with_events {
                    let active_mod = self.active_module.clone();
                    if let Some(old_val) = self.lookup(&key, frame)
                        && let Variant::Object(ObjRef::UserClass(old_id)) = old_val
                    {
                        self.remove_event_subscriptions(old_id, &active_mod, &key);
                    }
                    if let Variant::Object(ObjRef::UserClass(new_id)) = &v {
                        self.add_event_subscriptions(*new_id, &active_mod, &key, frame.me);
                    }
                }

                if module_level {
                    if let Some(m) = self.modules.get_mut(&self.active_module)
                        && let Some(old) = m.globals.insert(key, v)
                    {
                        self.dec_ref(&old);
                    }
                } else if frame.locals.contains_key(&key)
                    || frame.auto_new_locals.contains_key(&key)
                    || (!self.module_has_global(&self.active_module, &key)
                        && !self.instance_has_field(frame.me, &key))
                {
                    if let Some(old) = frame.locals.insert(key, v) {
                        self.dec_ref(&old);
                    }
                } else if self.instance_has_field(frame.me, &key) {
                    if let Some(ObjRef::UserClass(id)) = frame.me
                        && let Some(inst) = self.instances.get_mut(&id)
                        && let Some(old) = inst.fields.insert(key, v)
                    {
                        self.dec_ref(&old);
                    }
                } else if self.module_has_global(&self.active_module, &key) {
                    if let Some(m) = self.modules.get_mut(&self.active_module)
                        && let Some(old) = m.globals.insert(key, v)
                    {
                        self.dec_ref(&old);
                    }
                } else {
                    frame.locals.insert(key, v);
                }
                Ok(())
            }
            other => Err(VbaError::new(
                erl_assign_error(),
                format!("Cannot assign to this expression ({other:?})"),
            )),
        }
    }

    fn module_has_global(&self, mod_name: &str, key: &str) -> bool {
        self.modules
            .get(mod_name)
            .is_some_and(|m| m.globals.contains_key(key))
    }

    fn instance_has_field(&self, me: Option<ObjRef>, key: &str) -> bool {
        if let Some(ObjRef::UserClass(id)) = me {
            self.instances
                .get(&id)
                .is_some_and(|i| i.fields.contains_key(key))
        } else {
            false
        }
    }

    fn set_member_on_object(
        &mut self,
        obj: &ObjRef,
        name: &str,
        args: &[Variant],
        value: &Variant,
        set: bool,
    ) -> VResult<()> {
        match obj {
            ObjRef::Nothing => Err(VbaError::new(
                91,
                format!("Object variable or With block variable not set: .{name}"),
            )),
            ObjRef::UserClass(id) => {
                let cls_name = self
                    .instances
                    .get(id)
                    .map(|inst| inst.class_name.clone())
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                let lower_cls = cls_name.to_ascii_lowercase();
                let lower_name = name.to_ascii_lowercase();

                let mut prop_proc = None;
                if let Some(m_env) = self.modules.get(&lower_cls)
                    && let Some(mp) = m_env.procs.get(&lower_name)
                {
                    if set {
                        prop_proc = mp.prop_set.clone().or_else(|| mp.prop_let.clone());
                    } else {
                        prop_proc = mp.prop_let.clone().or_else(|| mp.prop_set.clone());
                    }
                }

                if let Some(proc) = prop_proc {
                    let mut full_args = args.to_vec();
                    full_args.push(value.clone());
                    let mut frame = Frame::new();
                    frame.module_name = cls_name.clone();
                    frame.me = Some(ObjRef::UserClass(*id));
                    let old_active = self.active_module.clone();
                    self.active_module = cls_name;
                    let res = self.call_body_with_frame(&proc, full_args, &mut frame);
                    self.active_module = old_active;
                    res?;
                    return Ok(());
                }

                if args.is_empty() {
                    self.inc_ref(value);
                    let old = if let Some(inst) = self.instances.get_mut(id) {
                        inst.fields.insert(lower_name, value.clone())
                    } else {
                        None
                    };
                    if let Some(old) = old {
                        self.dec_ref(&old);
                    }
                    return Ok(());
                }

                Err(VbaError::new(
                    438,
                    format!("Object doesn't support this property or method: .{name}"),
                ))
            }
            ObjRef::Worksheet(sheet_id) => {
                if let Some(mod_name) = self.find_document_module_name_by_sheet_id(*sheet_id) {
                    let mut prop_proc = None;
                    if let Some(m_env) = self.modules.get(&mod_name)
                        && let Some(mp) = m_env.procs.get(&name.to_ascii_lowercase())
                    {
                        if set {
                            prop_proc = mp.prop_set.clone().or_else(|| mp.prop_let.clone());
                        } else {
                            prop_proc = mp.prop_let.clone().or_else(|| mp.prop_set.clone());
                        }
                    }
                    if let Some(proc) = prop_proc {
                        let mut full_args = args.to_vec();
                        full_args.push(value.clone());
                        let mut frame = Frame::new();
                        frame.module_name = mod_name.clone();
                        frame.me = Some(ObjRef::Worksheet(*sheet_id));
                        let old_active = self.active_module.clone();
                        self.active_module = mod_name;
                        let res = self.call_body_with_frame(&proc, full_args, &mut frame);
                        self.active_module = old_active;
                        res?;
                        return Ok(());
                    }
                }
                self.host(&format!(".{name}"))?
                    .set_member(obj, name, args, value)?;
                self.drain_and_fire_events()?;
                Ok(())
            }
            ObjRef::Workbook => {
                if let Some(m_env) = self.modules.get("thisworkbook") {
                    let mut prop_proc = None;
                    if let Some(mp) = m_env.procs.get(&name.to_ascii_lowercase()) {
                        if set {
                            prop_proc = mp.prop_set.clone().or_else(|| mp.prop_let.clone());
                        } else {
                            prop_proc = mp.prop_let.clone().or_else(|| mp.prop_set.clone());
                        }
                    }
                    if let Some(proc) = prop_proc {
                        let mut full_args = args.to_vec();
                        full_args.push(value.clone());
                        let mut frame = Frame::new();
                        frame.module_name = "thisworkbook".to_string();
                        frame.me = Some(ObjRef::Workbook);
                        let old_active = self.active_module.clone();
                        self.active_module = "thisworkbook".to_string();
                        let res = self.call_body_with_frame(&proc, full_args, &mut frame);
                        self.active_module = old_active;
                        res?;
                        return Ok(());
                    }
                }
                self.host(&format!(".{name}"))?
                    .set_member(obj, name, args, value)?;
                self.drain_and_fire_events()?;
                Ok(())
            }
            other => {
                self.host(&format!(".{name}"))?
                    .set_member(other, name, args, value)?;
                self.drain_and_fire_events()?;
                Ok(())
            }
        }
    }

    fn lookup(&mut self, name: &str, frame: &mut Frame) -> Option<Variant> {
        let key = name.to_ascii_lowercase();

        if frame.auto_new_locals.contains_key(&key) {
            let val = frame.locals.get(&key).cloned();
            if val
                .as_ref()
                .is_none_or(|v| matches!(v, Variant::Empty | Variant::Object(ObjRef::Nothing)))
            {
                let cls = frame.auto_new_locals.get(&key).unwrap().clone();
                if let Ok(new_inst) = self.instantiate_class(&cls) {
                    self.inc_ref(&new_inst);
                    frame.locals.insert(key.clone(), new_inst.clone());
                    return Some(new_inst);
                }
            }
        }
        if let Some(v) = frame.locals.get(&key) {
            return Some(v.clone());
        }

        if let Some(ObjRef::UserClass(id)) = frame.me {
            let auto_cls = self.instances.get(&id).and_then(|inst| {
                if inst.auto_new_fields.contains_key(&key) {
                    let val = inst.fields.get(&key);
                    if val.is_none_or(|v| {
                        matches!(v, Variant::Empty | Variant::Object(ObjRef::Nothing))
                    }) {
                        return inst.auto_new_fields.get(&key).cloned();
                    }
                }
                None
            });
            if let Some(cls) = auto_cls
                && let Ok(new_inst) = self.instantiate_class(&cls)
            {
                self.inc_ref(&new_inst);
                if let Some(inst) = self.instances.get_mut(&id) {
                    inst.fields.insert(key.clone(), new_inst.clone());
                }
                return Some(new_inst);
            }
            if let Some(inst) = self.instances.get(&id)
                && let Some(v) = inst.fields.get(&key)
            {
                return Some(v.clone());
            }
        }

        let auto_cls = self.modules.get(&self.active_module).and_then(|m| {
            if m.auto_new_vars.contains_key(&key) {
                let val = m.globals.get(&key);
                if val
                    .is_none_or(|v| matches!(v, Variant::Empty | Variant::Object(ObjRef::Nothing)))
                {
                    return m.auto_new_vars.get(&key).cloned();
                }
            }
            None
        });
        if let Some(cls) = auto_cls
            && let Ok(new_inst) = self.instantiate_class(&cls)
        {
            self.inc_ref(&new_inst);
            if let Some(m) = self.modules.get_mut(&self.active_module) {
                m.globals.insert(key.clone(), new_inst.clone());
            }
            return Some(new_inst);
        }
        if let Some(m) = self.modules.get(&self.active_module)
            && let Some(v) = m.globals.get(&key)
        {
            return Some(v.clone());
        }

        if key == "thisworkbook" {
            return Some(Variant::Object(ObjRef::Workbook));
        }
        for m in self.modules.values() {
            if m.kind == VbaModuleKind::Document && m.name.eq_ignore_ascii_case(name) {
                if let Some(sheet_id) = m.bound_sheet_id {
                    return Some(Variant::Object(ObjRef::Worksheet(sheet_id)));
                } else {
                    return Some(Variant::Object(ObjRef::Workbook));
                }
            }
        }

        for m in self.modules.values() {
            if m.kind == VbaModuleKind::Standard
                && m.name != self.active_module
                && let Some(v) = m.globals.get(&key)
            {
                return Some(v.clone());
            }
        }

        None
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
                if let Some((proc_mod, proc)) = self.find_procedure_in_scope(name, frame) {
                    let mut call_frame = Frame::new();
                    call_frame.module_name = proc_mod.clone();
                    if proc_mod == frame.module_name {
                        call_frame.me = frame.me;
                    }
                    let old_active = self.active_module.clone();
                    self.active_module = proc_mod;
                    let res = self.call_body_with_frame(&proc, Vec::new(), &mut call_frame);
                    self.active_module = old_active;
                    return res;
                }
                if let Some(v) = builtins::call(name, &[])? {
                    return Ok(v);
                }
                Ok(Variant::Empty)
            }

            Expr::Unary { op, expr, .. } => {
                let v = self.eval(expr, frame)?;
                let v = self.scalar(v)?;
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
                if *op == BinOp::Is {
                    let b = self.eval(rhs, frame)?;
                    return is_comparison(&a, &b);
                }
                let a = self.scalar(a)?;
                if matches!(
                    op,
                    BinOp::And | BinOp::Or | BinOp::Xor | BinOp::Eqv | BinOp::Imp
                ) && matches!(a, Variant::Str(_))
                    && operand_kind(lhs) != Operand::Runtime
                {
                    a.to_bool()?;
                }
                let b = self.eval(rhs, frame)?;
                let b = self.scalar(b)?;
                let mode = if is_statically_typed(lhs) && is_statically_typed(rhs) {
                    ArithMode::Constant
                } else {
                    ArithMode::Promote
                };
                let kinds = (operand_kind(lhs), operand_kind(rhs));
                eval_binary(*op, &a, &b, mode, kinds)
            }

            Expr::Call { target, args, .. } => self.eval_call(target, args, frame),

            Expr::Member { target, name, .. } => {
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
                if let Some(t) = target
                    && let Expr::Ident {
                        name: mod_ident, ..
                    } = t.as_ref()
                {
                    let lower_mod = mod_ident.to_ascii_lowercase();
                    if let Some(m_env) = self.modules.get(&lower_mod)
                        && m_env.kind == VbaModuleKind::Standard
                    {
                        if let Some(v) = m_env.globals.get(&name.to_ascii_lowercase()) {
                            return Ok(v.clone());
                        }
                        if let Some(mp) = m_env.procs.get(&name.to_ascii_lowercase())
                            && let Some(proc) = mp.first()
                        {
                            let mut call_frame = Frame::new();
                            call_frame.module_name = lower_mod.clone();
                            let old_active = self.active_module.clone();
                            self.active_module = lower_mod;
                            let res = self.call_body_with_frame(&proc, Vec::new(), &mut call_frame);
                            self.active_module = old_active;
                            return res;
                        }
                    }
                }
                self.member(target.as_deref(), name, &[], frame)
            }

            Expr::Bang { target, name, .. } => {
                let owner = self.eval(target, frame)?;
                let Variant::Object(obj) = owner else {
                    return Err(VbaError::new(
                        424,
                        format!("Object required: !{name} on a {}", owner.type_name()),
                    ));
                };
                self.get_member_on_object(&obj, name, &[])
            }

            Expr::Me { .. } => {
                let me = frame
                    .me
                    .ok_or_else(|| VbaError::new(543, "Invalid use of Me keyword"))?;
                Ok(Variant::Object(me))
            }

            Expr::New { ty, .. } => {
                let class_name = ty.path.last().map(|s| s.as_str()).unwrap_or("");
                self.instantiate_class(class_name)
            }

            Expr::TypeOf { expr, ty, .. } => {
                let val = self.eval(expr, frame)?;
                let type_name = ty.path.last().map(|s| s.as_str()).unwrap_or("");
                Ok(Variant::Boolean(self.type_of_matches(&val, type_name)))
            }

            Expr::AddressOf { .. } => Ok(Variant::Long(0)),
        }
    }

    fn member_owner(&mut self, target: Option<&Expr>, frame: &mut Frame) -> VResult<Variant> {
        match target {
            Some(e) => self.eval(e, frame),
            None => frame
                .with_stack
                .last()
                .cloned()
                .ok_or_else(|| {
                    VbaError::new(
                        91,
                        "Object variable or With block variable not set: a leading '.' outside a With block",
                    )
                }),
        }
    }

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
        self.get_member_on_object(&obj, name, args)
    }

    fn get_member_on_object(
        &mut self,
        obj: &ObjRef,
        name: &str,
        args: &[Variant],
    ) -> VResult<Variant> {
        match obj {
            ObjRef::Nothing => Err(VbaError::new(
                91,
                format!("Object variable or With block variable not set: .{name}"),
            )),
            ObjRef::UserClass(id) => {
                let cls_name = self
                    .instances
                    .get(id)
                    .map(|inst| inst.class_name.clone())
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                let lower_cls = cls_name.to_ascii_lowercase();
                let lower_name = name.to_ascii_lowercase();

                if let Some(m_env) = self.modules.get(&lower_cls)
                    && let Some(mp) = m_env.procs.get(&lower_name)
                    && let Some(proc) = mp.prop_get.clone().or_else(|| mp.sub_or_func.clone())
                {
                    let mut frame = Frame::new();
                    frame.module_name = cls_name.clone();
                    frame.me = Some(ObjRef::UserClass(*id));
                    let old_active = self.active_module.clone();
                    self.active_module = cls_name;
                    let res = self.call_body_with_frame(&proc, args.to_vec(), &mut frame);
                    self.active_module = old_active;
                    return res;
                }

                let auto_cls = self.instances.get(id).and_then(|inst| {
                    if inst.auto_new_fields.contains_key(&lower_name) {
                        let val = inst.fields.get(&lower_name);
                        if val.is_none_or(|v| {
                            matches!(v, Variant::Empty | Variant::Object(ObjRef::Nothing))
                        }) {
                            return inst.auto_new_fields.get(&lower_name).cloned();
                        }
                    }
                    None
                });
                if let Some(cls) = auto_cls
                    && let Ok(new_inst) = self.instantiate_class(&cls)
                {
                    self.inc_ref(&new_inst);
                    if let Some(inst) = self.instances.get_mut(id) {
                        inst.fields.insert(lower_name.clone(), new_inst.clone());
                    }
                    return Ok(new_inst);
                }
                if let Some(inst) = self.instances.get(id)
                    && let Some(v) = inst.fields.get(&lower_name).cloned()
                {
                    if args.is_empty() {
                        return Ok(v);
                    }
                    if let Variant::Object(nested_obj) = &v {
                        return self.get_member_on_object(nested_obj, "item", args);
                    }
                }

                Err(VbaError::new(
                    438,
                    format!("Object doesn't support this property or method: .{name}"),
                ))
            }
            ObjRef::Worksheet(sheet_id) => {
                if let Some(mod_name) = self.find_document_module_name_by_sheet_id(*sheet_id)
                    && let Some(m_env) = self.modules.get(&mod_name)
                    && let Some(mp) = m_env.procs.get(&name.to_ascii_lowercase())
                    && let Some(proc) = mp.prop_get.clone().or_else(|| mp.sub_or_func.clone())
                {
                    let mut frame = Frame::new();
                    frame.module_name = mod_name.clone();
                    frame.me = Some(ObjRef::Worksheet(*sheet_id));
                    let old_active = self.active_module.clone();
                    self.active_module = mod_name;
                    let res = self.call_body_with_frame(&proc, args.to_vec(), &mut frame);
                    self.active_module = old_active;
                    return res;
                }
                self.host(&format!(".{name}"))?.get_member(obj, name, args)
            }
            ObjRef::Workbook => {
                if let Some(m_env) = self.modules.get("thisworkbook")
                    && let Some(mp) = m_env.procs.get(&name.to_ascii_lowercase())
                    && let Some(proc) = mp.prop_get.clone().or_else(|| mp.sub_or_func.clone())
                {
                    let mut frame = Frame::new();
                    frame.module_name = "thisworkbook".to_string();
                    frame.me = Some(ObjRef::Workbook);
                    let old_active = self.active_module.clone();
                    self.active_module = "thisworkbook".to_string();
                    let res = self.call_body_with_frame(&proc, args.to_vec(), &mut frame);
                    self.active_module = old_active;
                    return res;
                }
                self.host(&format!(".{name}"))?.get_member(obj, name, args)
            }
            other => self
                .host(&format!(".{name}"))?
                .get_member(other, name, args),
        }
    }

    fn scalar(&mut self, v: Variant) -> VResult<Variant> {
        match v {
            Variant::Object(ObjRef::UserClass(id)) => {
                let cls_name = self
                    .instances
                    .get(&id)
                    .map(|inst| inst.class_name.clone())
                    .unwrap_or_default();
                let def_member = self
                    .modules
                    .get(&cls_name.to_ascii_lowercase())
                    .and_then(|m| m.default_member.clone());
                if let Some(def_member) = def_member {
                    let res =
                        self.get_member_on_object(&ObjRef::UserClass(id), &def_member, &[])?;
                    return self.scalar(res);
                }
                Err(VbaError::new(
                    438,
                    "Object doesn't support this property or method: default property not found",
                ))
            }
            Variant::Object(obj) => self.host("using an object as a value")?.default_value(&obj),
            other => Ok(other),
        }
    }

    fn type_of_matches(&self, v: &Variant, type_name: &str) -> bool {
        let Variant::Object(obj) = v else {
            return false;
        };
        match obj {
            ObjRef::Nothing => false,
            ObjRef::UserClass(id) => {
                if let Some(inst) = self.instances.get(id) {
                    inst.class_name.eq_ignore_ascii_case(type_name)
                        || type_name.eq_ignore_ascii_case("object")
                } else {
                    false
                }
            }
            ObjRef::Worksheet(_) => {
                type_name.eq_ignore_ascii_case("worksheet")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::Workbook => {
                type_name.eq_ignore_ascii_case("workbook")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::Range(_) => {
                type_name.eq_ignore_ascii_case("range") || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::Application => {
                type_name.eq_ignore_ascii_case("application")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::ListObject(_) => {
                type_name.eq_ignore_ascii_case("listobject")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::ListColumn(..) => {
                type_name.eq_ignore_ascii_case("listcolumn")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::ListRow(..) => {
                type_name.eq_ignore_ascii_case("listrow")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::PivotTable(_) => {
                type_name.eq_ignore_ascii_case("pivottable")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::PivotField(..) => {
                type_name.eq_ignore_ascii_case("pivotfield")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::Interior(_) => {
                type_name.eq_ignore_ascii_case("interior")
                    || type_name.eq_ignore_ascii_case("object")
            }
            ObjRef::Font(_) => {
                type_name.eq_ignore_ascii_case("font") || type_name.eq_ignore_ascii_case("object")
            }
            _ => type_name.eq_ignore_ascii_case("object"),
        }
    }

    fn exec_for_each(
        &mut self,
        var: &Expr,
        iterable: &Expr,
        body: &[Stmt],
        frame: &mut Frame,
    ) -> VResult<Flow> {
        let subject = self.eval(iterable, frame)?;
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

        if let Expr::Member {
            target: obj, name, ..
        } = target
        {
            if let Some(t) = obj
                && let Expr::Ident {
                    name: mod_ident, ..
                } = t.as_ref()
            {
                let lower_mod = mod_ident.to_ascii_lowercase();
                if let Some(m_env) = self.modules.get(&lower_mod)
                    && m_env.kind == VbaModuleKind::Standard
                    && let Some(mp) = m_env.procs.get(&name.to_ascii_lowercase())
                    && let Some(proc) = mp.first()
                {
                    let values = self.eval_args(args, frame)?;
                    let mut call_frame = Frame::new();
                    call_frame.module_name = lower_mod.clone();
                    let old_active = self.active_module.clone();
                    self.active_module = lower_mod;
                    let res = self.call_body_with_frame(&proc, values, &mut call_frame);
                    self.active_module = old_active;
                    return res;
                }
            }
            let values = self.eval_args(args, frame)?;
            return self.member(obj.as_deref(), name, &values, frame);
        }

        let Expr::Ident { name, .. } = target else {
            return Err(out_of_scope("this call target"));
        };

        if let Some((proc_mod, proc)) = self.find_procedure_in_scope(name, frame) {
            let values = self.eval_args(args, frame)?;
            let mut call_frame = Frame::new();
            call_frame.module_name = proc_mod.clone();
            if proc_mod == frame.module_name {
                call_frame.me = frame.me;
            }
            let old_active = self.active_module.clone();
            self.active_module = proc_mod;
            let res = self.call_body_with_frame(&proc, values, &mut call_frame);
            self.active_module = old_active;
            return res;
        }

        if let Some(val) = self.lookup(name, frame) {
            if let Variant::Object(ObjRef::UserClass(id)) = val {
                let values = self.eval_args(args, frame)?;
                let cls_name = self
                    .instances
                    .get(&id)
                    .map(|inst| inst.class_name.clone())
                    .unwrap_or_default();
                if let Some(def_member) = self
                    .modules
                    .get(&cls_name.to_ascii_lowercase())
                    .and_then(|m| m.default_member.clone())
                {
                    return self.get_member_on_object(&ObjRef::UserClass(id), &def_member, &values);
                }
                return Err(VbaError::new(
                    438,
                    "Object doesn't support this property or method: default member not found",
                ));
            }
            if let Variant::Array(a) = val {
                let values = self.eval_args(args, frame)?;
                let row = values
                    .first()
                    .map(|v| v.to_f64())
                    .transpose()?
                    .unwrap_or(0.0);
                let col = match values.get(1) {
                    Some(v) => v.to_f64()?,
                    None => 0.0,
                };
                return a.get(row as usize, col as usize);
            }
        }

        let values = self.eval_args(args, frame)?;

        if name.eq_ignore_ascii_case("typename")
            && let Some(arg) = values.first()
        {
            return Ok(Variant::Str(self.type_name_of(arg)));
        }

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

fn is_constant(e: &Expr) -> bool {
    match e {
        Expr::Literal(Literal::Null) => false,
        Expr::Literal(_) => true,
        Expr::Paren { expr, .. } => is_constant(expr),
        Expr::Unary { expr, .. } => is_constant(expr),
        Expr::Binary { lhs, rhs, .. } => is_constant(lhs) && is_constant(rhs),
        _ => false,
    }
}

const STATICALLY_NUMERIC: &[&str] = &[
    "cint", "clng", "cdbl", "csng", "ccur", "cbool", "cbyte", "len", "val", "sgn",
];

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

const STATICALLY_STRING: &[&str] = &["cstr", "typename", "strreverse", "replace", "join"];

fn is_statically_boolean(e: &Expr) -> bool {
    match e {
        Expr::Paren { expr, .. } => is_statically_boolean(expr),
        Expr::Unary {
            op: UnOp::Not,
            expr,
            ..
        } => is_statically_boolean(expr),
        Expr::Call { target, .. } => matches!(target.as_ref(), Expr::Ident { name, .. }
            if STATICALLY_BOOLEAN.contains(&name.to_ascii_lowercase().as_str())),
        Expr::Binary { op, lhs, rhs, .. }
            if matches!(op, BinOp::IntDiv | BinOp::Mod)
                && is_literal_bool(lhs)
                && is_literal_string(rhs) =>
        {
            false
        }
        Expr::Binary {
            op: BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Gt | BinOp::Le | BinOp::Ge,
            ..
        } => is_statically_typed(e),
        _ => is_constant(e),
    }
}

fn is_literal_bool(e: &Expr) -> bool {
    match e {
        Expr::Paren { expr, .. } => is_literal_bool(expr),
        Expr::Literal(Literal::Bool(_)) => true,
        _ => false,
    }
}

fn is_literal_string(e: &Expr) -> bool {
    match e {
        Expr::Paren { expr, .. } => is_literal_string(expr),
        Expr::Literal(Literal::Str(_)) => true,
        _ => false,
    }
}

fn is_statically_typed(e: &Expr) -> bool {
    match e {
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
                    | BinOp::Concat
                    | BinOp::Eq
                    | BinOp::Ne
                    | BinOp::Lt
                    | BinOp::Gt
                    | BinOp::Le
                    | BinOp::Ge
            ) && is_statically_typed(lhs)
                && is_statically_typed(rhs)
                || matches!(
                    op,
                    BinOp::And | BinOp::Or | BinOp::Xor | BinOp::Eqv | BinOp::Imp
                )
        }
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

fn operand_kind(e: &Expr) -> Operand {
    let statically_typed = is_statically_typed(e);
    match e {
        Expr::Literal(_) => Operand::Literal,
        Expr::Paren { expr, .. } | Expr::Unary { expr, .. }
            if operand_kind(expr) == Operand::Literal =>
        {
            Operand::Literal
        }
        Expr::Paren { expr, .. } if operand_kind(expr) == Operand::Runtime => Operand::Runtime,
        Expr::Binary { op, lhs, rhs, .. }
            if matches!(
                op,
                BinOp::And | BinOp::Or | BinOp::Xor | BinOp::Eqv | BinOp::Imp
            ) && (operand_kind(lhs) == Operand::Runtime
                || operand_kind(rhs) == Operand::Runtime) =>
        {
            Operand::Runtime
        }
        _ if is_constant(e) && statically_typed => Operand::ConstExpr,
        _ if statically_typed => Operand::Static,
        _ => Operand::Runtime,
    }
}

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
        And => null_on_the_right(a, b, kinds, value::and(a, b, kinds)),
        Or => null_on_the_right(a, b, kinds, value::or(a, b, kinds)),
        Xor => null_on_the_right(a, b, kinds, value::logical(a, b, kinds, |x, y| x ^ y)),
        Eqv => null_on_the_right(a, b, kinds, value::logical(a, b, kinds, |x, y| !(x ^ y))),
        Imp => null_on_the_right(a, b, kinds, value::imp(a, b, kinds)),
        Like => Err(out_of_scope("Like")),
        Is => is_comparison(a, b),
    }
}

const OBJECT_AWARE_BUILTINS: &[&str] = &["typename", "vartype", "isobject"];

fn is_comparison(a: &Variant, b: &Variant) -> VResult<Variant> {
    match (a.as_object(), b.as_object()) {
        (Some(x), Some(y)) => Ok(Variant::Boolean(x.same_object(y))),
        _ => Err(VbaError::new(424, "Object required: Is compares objects")),
    }
}

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
                let _ = base;
                Variant::from_literal(*value, *is_float || value.fract() != 0.0)
            }
        },
        Literal::Str(s) => Variant::Str(s.clone()),
        Literal::Bool(b) => Variant::Boolean(*b),
        Literal::Empty => Variant::Empty,
        Literal::Null => Variant::Null,
        Literal::Date(text) => match crate::core::date::parse_date(text) {
            Some((d, _)) => Variant::Date(crate::core::date::date_to_excel_serial(d)),
            None => Variant::Empty,
        },
        Literal::Nothing => Variant::Object(ObjRef::Nothing),
    }
}

fn number_like(current: f64, start: f64, step: f64) -> Variant {
    let integral = current.fract() == 0.0 && start.fract() == 0.0 && step.fract() == 0.0;
    Variant::from_literal(current, !integral)
}

fn default_for(ty: Option<&TypeRef>) -> Variant {
    let Some(ty) = ty else {
        return Variant::Empty;
    };
    if ty.is_new {
        return Variant::Object(ObjRef::Nothing);
    }
    let Some(last) = ty.path.last() else {
        return Variant::Empty;
    };
    match last.to_ascii_lowercase().as_str() {
        "integer" => Variant::Integer(0),
        "long" => Variant::Long(0),
        "single" => Variant::Single(0.0),
        "double" => Variant::Double(0.0),
        "currency" => Variant::Currency(0),
        "boolean" => Variant::Boolean(false),
        "string" => Variant::Str(String::new()),
        "date" => Variant::Date(0.0),
        "variant" => Variant::Empty,
        _ => Variant::Object(ObjRef::Nothing),
    }
}

#[cfg(test)]
mod tests {
    use super::super::parser::parse_module;
    use super::*;

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

    #[test]
    fn arithmetic_and_types_match_the_excel_probe() {
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
        assert_eq!(
            run("    Dim t\n    For i = 5 To 1 Step -1\n        t = t + i\n    Next i\n    F = t"),
            "Integer|15"
        );
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
        assert_eq!(
            run("    Dim s As String\n    F = TypeName(s)"),
            "String|String"
        );
        assert_eq!(run("    Dim n As Long\n    F = TypeName(n)"), "String|Long");
        assert_eq!(run("    Dim v\n    F = TypeName(v)"), "String|Empty");
    }

    #[test]
    fn and_or_and_imp_are_three_valued() {
        assert_eq!(expr("False And Null"), "Boolean|False");
        assert_eq!(expr("True Or Null"), "Boolean|True");
        assert_eq!(expr("IsNull(True And Null)"), "Boolean|True");
        assert_eq!(expr("IsNull(False Or Null)"), "Boolean|True");
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
        assert_eq!(expr("Null Imp True"), "Boolean|True");
        assert_eq!(expr("False Imp Null"), "Boolean|True");
        assert_eq!(expr("IsNull(Null Xor True)"), "Boolean|True");
        assert_eq!(expr("IsNull(Null Eqv True)"), "Boolean|True");
        assert_eq!(expr("IsNull(Not Null)"), "Boolean|True");
        assert_eq!(expr("UCase(\"False\") And Null"), "Boolean|False");
        assert_eq!(expr("UCase(\"0\") And Null"), "Long|0");
        assert_eq!(expr("IsNull(UCase(\"True\") And Null)"), "Boolean|True");
        assert_eq!(expr("UCase(\"abc\") And Null"), "ERR|13");
    }

    #[test]
    fn string_versus_number_comparison_depends_on_constant_ness() {
        assert_eq!(expr("\"10\" = 10"), "Boolean|True");
        assert_eq!(expr("\"2\" > 10"), "Boolean|False");
        assert_eq!(expr("\"\" = 0"), "ERR|13");
        assert_eq!(expr("\"abc\" > 1"), "ERR|13");

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

        assert_eq!(
            run("    Dim b\n    b = 10\n    F = (\"2\" > b)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim b\n    b = 1\n    F = (\"abc\" > b)"),
            "Boolean|True"
        );

        assert_eq!(
            run("    Dim a\n    a = True\n    F = ((1.5 & \"abc\") <> CLng(a))"),
            "ERR|13"
        );
        assert_eq!(
            run("    Dim a\n    a = -1\n    F = ((1.5 & \"abc\") <> a)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = (\"Z\" <> a)"),
            "Boolean|True"
        );

        assert_eq!(expr("(Not 2!) <= (\"1.5\" & False)"), "Boolean|True");
        assert_eq!(expr("(-True) <> (True & &HFF)"), "ERR|13");
        assert_eq!(expr("\"False\" = -0.04"), "ERR|13");
        assert_eq!(expr("\"1.5abc\" > 1"), "Boolean|True");
        assert_eq!(expr("(False & Null) = (0.1 / -2.5)"), "Boolean|False");

        assert_eq!(
            run("    Dim a\n    a = True\n    F = ((1.5 & \"abc\") <> CLng(a))"),
            "ERR|13"
        );
        assert_eq!(
            run("    Dim a\n    a = 1\n    F = ((\"abc\" & a) <> Len(CStr(\"Z\")))"),
            "Boolean|True"
        );

        assert_eq!(
            run("    Dim a\n    a = \"True\"\n    F = (a = True)"),
            "Boolean|True"
        );
        assert_eq!(expr("\"True\" = -1"), "ERR|13");

        assert_eq!(
            run("    Dim a, b\n    a = \"1.5\"\n    b = 1.5\n    F = (a = b)"),
            "Boolean|False"
        );
        assert_eq!(
            run("    Dim a, b\n    a = \"2\"\n    b = 10\n    F = (a > b)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim a\n    a = \"1True\"\n    F = (a = 1)"),
            "Boolean|False"
        );
        assert_eq!(
            run("    Dim a\n    a = \"3abc\"\n    F = (a < 5)"),
            "Boolean|False"
        );
    }

    #[test]
    fn a_constant_string_select_subject_compares_as_text() {
        let sel = |subject: &str| {
            format!(
                "    Dim r\n    Select Case {subject}\n    Case 2 To 5\n        r = \"range\"\n    \
                 Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(run(&sel("\"32768abc\"")), "String|range");
        assert_eq!(run(&sel("(32768 & \"abc\")")), "String|range");
        assert_eq!(run(&sel("\"3\"")), "String|range");
        assert_eq!(run(&sel("\"abc\"")), "String|else");
        assert_eq!(run(&sel("\"7\"")), "String|else");
        assert_eq!(run(&sel("\"1x\"")), "String|else");
        assert_eq!(run(&sel("\"\"")), "String|else");
        assert_eq!(run(&sel("3")), "String|range");
        assert_eq!(run(&sel("7")), "String|else");

        let sel_var = |value: &str| {
            format!(
                "    Dim a, r\n    a = {value}\n    Select Case a\n    Case 2 To 5\n        \
                 r = \"range\"\n    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(run(&sel_var("\"32768abc\"")), "String|else");
        assert_eq!(run(&sel_var("\"3abc\"")), "String|else");
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
        assert_eq!(
            run(&sel("    Case Is >= 2\n        r = \"is\"\n")),
            "String|is"
        );
    }

    #[test]
    fn select_case_null_subject_matches_no_case_form() {
        let sel = |cases: &str| {
            format!(
                "    Dim r\n    Select Case Null\n{cases}    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            )
        };
        assert_eq!(
            run(&sel("    Case 2 To 5\n        r = \"range\"\n")),
            "String|else"
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

    #[test]
    fn zero_divided_by_zero_is_overflow_not_division_by_zero() {
        assert_eq!(expr("1 / 0"), "ERR|11");
        assert_eq!(expr("-1 / 0"), "ERR|11");
        assert_eq!(expr("1.5 / 0"), "ERR|11");
        assert_eq!(expr("0 / 0"), "ERR|6");
        assert_eq!(expr("False / 0"), "ERR|6");
        assert_eq!(expr("0 \\ 0"), "ERR|11");
        assert_eq!(expr("0 Mod 0"), "ERR|11");
    }

    #[test]
    fn logical_operators_convert_the_left_operand_before_evaluating_the_right() {
        assert_eq!(expr("(\"a\" + \"Z\") Eqv (\"1\" \\ 0)"), "ERR|13");
    }

    #[test]
    fn division_coerces_both_operands_before_testing_the_divisor() {
        assert_eq!(expr("\"xxxx\" / 0"), "ERR|13");
        assert_eq!(expr("\"\" / 0"), "ERR|13");
        assert_eq!(expr("0 / \"xxxx\""), "ERR|13");
        assert_eq!(expr("\"abc\" / Null"), "ERR|13");
    }

    #[test]
    fn a_static_string_over_a_null_is_invalid_use_of_null() {
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
        assert_eq!(
            run("    Dim a\n    a = \"  3  \"\n    F = IsNull(a Imp Null)"),
            "Boolean|False"
        );
        assert_eq!(expr("IsNull(Null Or \"  3  \")"), "Boolean|False");
        assert_eq!(expr("IsNull(Null Xor \"  3  \")"), "Boolean|True");
        assert_eq!(expr("\"abc\" Imp Null"), "ERR|13");
        assert_eq!(expr("\"True\" Or Null"), "ERR|13");
        assert_eq!(expr("IsNull(255 Imp Null)"), "Boolean|False");
    }

    #[test]
    fn a_statically_typed_numeric_partner_is_strict_only_against_a_constant_string() {
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
        assert_eq!(
            run("    Dim va, vb\n    va = 5\n    vb = \"1\"\n    F = (CLng(va) < vb)"),
            "Boolean|False"
        );
        assert_eq!(with("(CLng(va) < (\"abc\" & va))"), "Boolean|True");
    }

    #[test]
    fn negating_the_long_minimum_between_constants_wraps_to_itself() {
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
        assert_eq!(sel("(Not 5)"), "String|else");
    }

    #[test]
    fn logical_expression_width_is_static_for_arithmetic_overflow() {
        assert_eq!(
            run("    Dim vc\n    vc = 10\n    F = ((\"  3  \" And vc) - (Not 2147483647))"),
            "ERR|6"
        );
        assert_eq!(expr("(Trim(-1) And (1E3 * 1%)) >= \"7\""), "Boolean|False");
    }

    #[test]
    fn overflow_between_constants_is_really_between_statically_typed_operands() {
        assert_eq!(expr("CStr(CInt(32767) + 1)"), "ERR|6");
        assert_eq!(expr("CStr(CInt(32767) * 2)"), "ERR|6");
        assert_eq!(expr("CStr(CInt(32767) + CInt(1))"), "ERR|6");
        assert_eq!(expr("CStr(Sgn(1) + 32767)"), "ERR|6");
        assert_eq!(expr("CStr(CLng(2147483647) + 1)"), "ERR|6");
        assert_eq!(expr("CStr(CInt(32767) ^ 4652)"), "ERR|6");
        assert_eq!(expr("CStr(CDbl(32767) ^ 4652)"), "ERR|6");
        assert_eq!(expr("CStr(Len(\"abcde\") ^ 4652)"), "ERR|6");
        assert_eq!(expr("CStr((Empty + 32767) + 1)"), "String|32768");
        assert_eq!(expr("CStr(32767 + 1)"), "ERR|6");
        assert_eq!(
            run("    Dim a\n    a = 32767\n    F = CStr(a + 1)"),
            "String|32768"
        );
        assert_eq!(
            run("    Dim a\n    a = 1\n    F = CStr(CInt(32767) + a)"),
            "String|32768"
        );
        assert_eq!(expr("CStr(Len(\"abcde\") + 32763)"), "String|32768");
        assert_eq!(expr("CStr(CInt(Empty) + 32768)"), "String|32768");
        assert_eq!(
            run("    Dim vb\n    vb = 4652\n    F = CStr(32767 ^ vb)"),
            "ERR|6"
        );
    }

    #[test]
    fn a_statically_string_value_compares_as_text_against_a_runtime_number() {
        let with = |setup: &str, e: &str| run(&format!("    Dim a, b\n{setup}\n    F = CStr({e})"));
        assert_eq!(with("    a = 5", "(a < \"10\")"), "String|False");
        assert_eq!(with("    a = 5", "(a < CStr(10))"), "String|False");
        assert_eq!(with("    a = 5", "(a < (CStr(1) & \"0\"))"), "String|False");
        assert_eq!(
            with("    a = 5\n    b = 10", "(a < CStr(b))"),
            "String|False"
        );
        assert_eq!(with("    a = 5", "(a < Trim(\"10\"))"), "String|True");
        assert_eq!(with("    a = -2", "(a < CStr(\"\"))"), "String|False");
        assert_eq!(with("    a = -2", "(a < StrReverse(\"\"))"), "String|False");
        assert_eq!(with("    a = -2", "(a < CStr(\"abc\"))"), "String|True");
        assert_eq!(with("    a = -2", "(a < \"\")"), "String|False");
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
        assert_eq!(expr("CStr(InStr(\"\", \"\"))"), "String|0");
        assert_eq!(expr("CStr(InStr(Empty, \"\"))"), "String|0");
        assert_eq!(expr("CStr(InStr(\"a\", \"\"))"), "String|1");
        assert_eq!(expr("CStr(InStr(\"\", \"a\"))"), "String|0");
    }

    #[test]
    fn static_typing_propagates_through_arithmetic() {
        let with = |e: &str| run(&format!("    Dim a\n    a = -3\n    F = {e}"));
        assert_eq!(with("(Len(CStr(a)) = \"-7False\")"), "ERR|13");
        assert_eq!(with("((Len(CStr(a)) / 2) = \"-7False\")"), "ERR|13");
        assert_eq!(with("((Len(CStr(a)) + 1) = \"-7False\")"), "ERR|13");
        assert_eq!(with("((CLng(a) / 2) = \"abc\")"), "ERR|13");
        assert_eq!(
            with("((Len(CStr(a)) / (-32768)) = ((-7) & (0 > \"1.5\")))"),
            "ERR|13"
        );
        assert_eq!(with("((Len(CStr(a)) + a) = \"-7False\")"), "Boolean|False");
        assert_eq!(with("((a / (-32768)) = \"-7False\")"), "Boolean|False");
        assert_eq!(with("((a + 1) = \"-7False\")"), "Boolean|False");
        assert_eq!(with("((CLng(a) * 2) = \"-6.0\")"), "Boolean|True");
    }

    #[test]
    fn static_typing_propagates_through_comparison_and_concatenation() {
        assert_eq!(expr("(\"0\" >= (3# >= CDbl(0)))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (Len(CStr(0)) >= 1))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (\"1\" >= -7))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (2 >= 1))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (1 = 1))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= (3# >= Empty))"), "Boolean|False");
        assert_eq!(expr("(\"0\" >= (Empty = Empty))"), "Boolean|False");
        assert_eq!(expr("(\"0\" < (3# >= Empty))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= IsEmpty(Empty))"), "Boolean|True");
        assert_eq!(expr("(\"0\" >= CBool(Empty))"), "Boolean|True");
        assert_eq!(
            run("    Dim b\n    b = 1\n    F = (\"0\" >= (3# >= b))"),
            "Boolean|False"
        );
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
        assert_eq!(expr("((\"1\" + \"3\") <= False)"), "Boolean|True");
        assert_eq!(expr("((Empty & \"13\") <= False)"), "Boolean|True");
        assert_eq!(expr("((\"1\" + \"3\") = True)"), "Boolean|True");
        assert_eq!(expr("((\"1\" + \"3\") > False)"), "Boolean|False");
        assert_eq!(expr("((\"1\" + \"  3  \") <= False)"), "ERR|13");
        assert_eq!(expr("((\"abc\" + \"d\") > True)"), "ERR|13");
        assert_eq!(expr("((Empty & \"1  3  \") <= False)"), "Boolean|False");
        assert_eq!(
            expr("((\"1\" & \"  3  \") <= (\"\" <> Empty))"),
            "Boolean|True"
        );
    }

    #[test]
    fn a_string_converts_with_cbool_against_a_static_boolean() {
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
        assert_eq!(with("    va = \"abc\"", "(va = True)"), "Boolean|False");
        assert_eq!(with("    va = \"\"", "(va = False)"), "Boolean|False");
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
        assert_eq!(expr("(\"abc\" < True)"), "ERR|13");
        assert_eq!(expr("(\"Z\" < True)"), "ERR|13");
        assert_eq!(expr("(False >= \"abc\")"), "ERR|13");
        assert_eq!(expr("(\"\" = False)"), "ERR|13");
        assert_eq!(expr("(\"011\" < False)"), "Boolean|True");
        assert_eq!(expr("(False > \"12\")"), "Boolean|True");
        assert_eq!(expr("(\"0\" = False)"), "Boolean|True");
        assert_eq!(
            expr("((Empty & \"1\") <= (\"\" <> Empty))"),
            "Boolean|False"
        );
        assert_eq!(expr("TypeName(0) >= (3# >= Empty)"), "Boolean|False");
        assert_eq!(expr("(3# >= Empty) >= TypeName(0)"), "Boolean|True");
        assert_eq!(expr("CStr(0) >= (3# >= Empty)"), "Boolean|False");
        assert_eq!(expr("(Not True) <= CStr(32767)"), "Boolean|False");
        assert_eq!(expr("False >= TypeName(0)"), "ERR|13");
        assert_eq!(expr("(\"000\" < (\"1\" >= -7))"), "Boolean|False");
        assert_eq!(
            run("    Dim va\n    va = \"000\"\n    F = (va < (\"1\" >= -7))"),
            "Boolean|False"
        );
        assert_eq!(expr("(Right(100000, 3) < (\"1\" >= -7))"), "Boolean|False");
        assert_eq!(expr("(CStr(0) >= CBool(1))"), "Boolean|True");
        assert_eq!(expr("(\"000\" < CBool(1))"), "Boolean|False");
        assert_eq!(expr("(TypeName(0) >= CBool(1))"), "ERR|13");
        assert_eq!(
            with("    va = \"011\"\n    vb = False", "(va < vb)"),
            "Boolean|False"
        );
        assert_eq!(with("    va = \"True\"", "(va < False)"), "Boolean|True");
        assert_eq!(with("    va = \"true\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"TRUE\"", "(va = True)"), "Boolean|True");
        assert_eq!(with("    va = \"true\"", "(va = False)"), "Boolean|False");
        assert_eq!(
            with("    va = \"011\"\n    vb = False", "(va < vb)"),
            "Boolean|False"
        );
        assert_eq!(with("    va = \"011\"", "(va < 0)"), "Boolean|False");
        assert_eq!(expr("(\"True\" = -1)"), "ERR|13");
    }

    #[test]
    fn an_unconvertible_runtime_string_sorts_above_a_static_boolean() {
        let with = |setup: &str, e: &str| run(&format!("    Dim va\n{setup}\n    F = {e}"));
        assert_eq!(with("    va = \"ABC\"", "(va > True)"), "Boolean|True");
        assert_eq!(with("    va = \"ABC\"", "(va < True)"), "Boolean|False");
        assert_eq!(with("    va = \"ABC\"", "(va >= False)"), "Boolean|True");
        assert_eq!(expr("Chr(65) > True"), "Boolean|True");
        assert_eq!(expr("Chr(65) > False"), "Boolean|True");
        assert_eq!(expr("Hex(255) > True"), "Boolean|True");
        assert_eq!(expr("Space(2) > True"), "Boolean|True");
        assert_eq!(with("    va = \"abc\"", "(va = True)"), "Boolean|False");
        assert_eq!(expr("LCase(\"Integer\") >= (Not True)"), "Boolean|True");
    }

    #[test]
    fn statically_string_intrinsics_are_strict_against_a_boolean() {
        assert_eq!(expr("StrReverse(False) > (Not False)"), "ERR|13");
        assert_eq!(expr("StrReverse(\"abc\") > True"), "ERR|13");
        assert_eq!(expr("StrReverse(\"abc\") > 5"), "ERR|13");
        assert_eq!(expr("True > StrReverse(\"abc\")"), "ERR|13");
        assert_eq!(expr("Replace(\"abc\", \"a\", \"z\") > True"), "ERR|13");
        assert_eq!(expr("CStr(\"abc\") > 5"), "ERR|13");
        assert_eq!(expr("TypeName(1) > 5"), "ERR|13");
        assert_eq!(expr("CStr(\"abc\") >= 0"), "ERR|13");
        assert_eq!(expr("CStr(\"abc\") > CLng(1)"), "ERR|13");
        assert_eq!(expr("TypeName(1) > CLng(5)"), "ERR|13");
        assert_eq!(expr("Trim(\"abc\") > True"), "Boolean|True");
        assert_eq!(expr("LTrim(\"abc\") > True"), "Boolean|True");
        assert_eq!(expr("Trim(\"abc\") > 5"), "Boolean|True");
        assert_eq!(expr("Chr(65) > 5"), "Boolean|True");
        assert_eq!(expr("CStr(\"11\") > 5"), "Boolean|True");
        assert_eq!(expr("TypeName(1) > \"5\""), "Boolean|True");
        assert_eq!(
            run("    Dim va\n    va = StrReverse(\"abc\")\n    F = (va > True)"),
            "Boolean|True"
        );
        assert_eq!(
            run("    Dim va\n    va = 5\n    F = (StrReverse(\"abc\") > va)"),
            "Boolean|True"
        );
        assert_eq!(expr("StrReverse(\"11\") > False"), "Boolean|False");
    }

    #[test]
    fn division_overflows_rather_than_returning_an_infinity() {
        assert_eq!(expr("1E308 / 1E-308"), "ERR|6");
        assert_eq!(
            run("    Dim a, b\n    a = 1E308\n    b = 1E-308\n    F = a / b"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 3.75\n    b = a ^ 32767\n    F = b / 2"),
            "ERR|6"
        );
        assert_eq!(expr("1 / 2"), "Double|0.5");
        assert_eq!(expr("1 / 0"), "ERR|11");
        assert_eq!(expr("0 / 0"), "ERR|6");
    }

    #[test]
    fn pow_overflow_raises_at_runtime_too() {
        assert_eq!(expr("3.75 ^ 32767"), "ERR|6");
        assert_eq!(expr("255 ^ 255"), "ERR|6");
        assert_eq!(run("    Dim a\n    a = 3.75\n    F = (a ^ 32767)"), "ERR|6");
        assert_eq!(run("    Dim a\n    a = 255\n    F = (a ^ 255)"), "ERR|6");
        assert_eq!(expr("2 ^ 10"), "Double|1024");
    }

    #[test]
    fn overflowing_pow_raises_before_arithmetic_can_observe_infinity() {
        assert_eq!(run("    Dim a\n    a = 255\n    F = (a ^ 255)"), "ERR|6");
        assert_eq!(run("    Dim a\n    a = 255\n    F = -(a ^ 255)"), "ERR|6");
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = ((a ^ 255) & \"x\")"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a, b\n    a = 255\n    b = (a ^ 255)\n    F = (b + 1)"),
            "ERR|6"
        );
        assert_eq!(run("    Dim a\n    a = 1E300\n    F = (a * a)"), "ERR|6");
        assert_eq!(
            run("    Dim a, b\n    a = 1E300\n    b = 1E300\n    F = (a + b)"),
            "Double|2E+300"
        );
    }

    #[test]
    fn imp_follows_its_definition_rather_than_a_hand_rolled_table() {
        assert_eq!(
            run("    Dim a\n    a = 255\n    F = (a Imp Null)"),
            "Integer|-256"
        );
        assert_eq!(expr("Null Imp True"), "Boolean|True");
        assert_eq!(expr("False Imp Null"), "Boolean|True");
        assert_eq!(expr("5 Imp 3"), "Integer|-5");
    }

    #[test]
    fn single_combined_with_long_widens_past_both() {
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

    #[test]
    fn only_plus_short_circuits_past_a_bad_partner() {
        assert_eq!(expr("IsNull(Null + \"Z\")"), "Boolean|True");
        assert_eq!(expr("IsNull(\"Z\" + Null)"), "Boolean|True");
        assert_eq!(expr("IsNull(Null + \"12\")"), "Boolean|True");

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

        assert_eq!(expr("\"Z\" & Null"), "String|Z");

        assert_eq!(expr("IsNull(1 - Null)"), "Boolean|True");
        assert_eq!(expr("IsNull(Null Mod 3)"), "Boolean|True");
    }

    #[test]
    fn unary_sign_promotes_on_overflow_at_runtime() {
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = (-(Not a))"),
            "Double|2147483648"
        );
        assert_eq!(
            run("    Dim a\n    a = 2147483647\n    F = TypeName(-(Not a))"),
            "String|Double"
        );
        assert_eq!(
            run("    Dim a\n    a = 32767\n    F = (-(Not a))"),
            "Long|32768"
        );
    }

    #[test]
    fn a_statically_boolean_select_subject_converts_its_cases_with_cbool() {
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

        for subject in [
            "(1 = 1)",
            "True",
            "CBool(1)",
            "IsNumeric(0)",
            "(Val(&O17) >= (True > 3#))",
        ] {
            assert_eq!(hit(subject, "1"), "String|a", "{subject} vs Case 1");
            assert_eq!(hit(subject, "0"), "String|else", "{subject} vs Case 0");
            assert_eq!(hit(subject, "0, 1"), "String|a", "{subject} vs Case 0, 1");
            assert_eq!(hit(subject, "2 To 5"), "String|a", "{subject} vs 2 To 5");
            assert_eq!(hit(subject, "0 To 1"), "String|else", "{subject} vs 0 To 1");
            assert_eq!(hit(subject, "Is = 1"), "String|a", "{subject} vs Is = 1");
            assert_eq!(hit(subject, "Is > 0"), "String|else", "{subject} vs Is > 0");
            assert_eq!(hit(subject, "Is < 0"), "String|a", "{subject} vs Is < 0");
        }
        assert_eq!(hit("(1 = 2)", "0, 1"), "String|a");
        assert_eq!(hit("(1 = 2)", "2 To 5"), "String|else");
        assert_eq!(
            run(&sel("CBool(1)", "    Case Null\n        r = \"a\"\n")),
            "ERR|94"
        );

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

    #[test]
    fn select_case_constant_bool_int_op_subject_is_not_statically_boolean() {
        assert_eq!(
            run(
                "    Dim r\n    Select Case (True \\ \"12\")\n    Case 0, 1\n        r = \"value\"\n    Case Else\n        r = \"else\"\n    End Select\n    F = r"
            ),
            "String|else"
        );
    }

    #[test]
    fn a_constant_boolean_over_a_constant_string_folds_to_a_boolean() {
        assert_eq!(expr("True Mod \"12\""), "Boolean|False");
        assert_eq!(expr("True \\ \"12\""), "Boolean|True");
        assert_eq!(expr("False \\ \"12\""), "Boolean|False");
        assert_eq!(expr("True Mod \"0\""), "ERR|11");
        assert_eq!(expr("True \\ \"0\""), "ERR|11");

        assert_eq!(expr("\"12\" Mod True"), "Long|0");
        assert_eq!(expr("\"12\" \\ True"), "Long|-12");
        assert_eq!(expr("True Mod 12"), "Integer|-1");
        assert_eq!(expr("True \\ 12"), "Integer|0");
        assert_eq!(
            run("    Dim a\n    a = True\n    F = (a Mod \"12\")"),
            "Long|-1"
        );
        assert_eq!(
            run("    Dim b\n    b = \"12\"\n    F = (True Mod b)"),
            "Long|-1"
        );
        assert_eq!(expr("True And \"12\""), "Long|12");
        assert_eq!(expr("True Or \"12\""), "Long|-1");
        assert_eq!(expr("True Eqv \"12\""), "Long|12");
    }

    #[test]
    fn integer_operators_process_the_left_operand_first() {
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
        assert_eq!(expr("TypeName(Null)"), "String|Null");
        assert_eq!(expr("IsNull(Null)"), "Boolean|True");
        assert_eq!(expr("IsNumeric(Null)"), "Boolean|False");
        assert_eq!(expr("IsEmpty(Null)"), "Boolean|False");
    }

    #[test]
    fn conversions_reject_null_rather_than_propagating_it() {
        assert_eq!(expr("CStr(Null)"), "ERR|94");
        assert_eq!(expr("CDbl(Null)"), "ERR|94");
        assert_eq!(expr("CLng(Null)"), "ERR|94");
        assert_eq!(expr("IsNull(UCase(Null))"), "Boolean|True");
        assert_eq!(expr("IsNull(Left(Null, 1))"), "Boolean|True");
        assert_eq!(expr("TypeName(Null)"), "String|Null");
        assert_eq!(expr("IsNull(Null)"), "Boolean|True");
    }

    #[test]
    fn fuzz_statement_conditions_read_null_as_false_unlike_cbool() {
        assert_eq!(
            run("    If Null Then\n        F = \"T\"\n    Else\n        F = \"F\"\n    End If"),
            "String|F"
        );
        assert_eq!(
            run(
                "    Dim n As Integer\n    n = 0\n    Do While Null\n        n = n + 1\n    Loop\n    F = n"
            ),
            "Integer|0"
        );
        assert_eq!(
            run(
                "    Dim n As Integer\n    n = 0\n    Do Until Null\n        n = n + 1\n        If n > 3 Then Exit Do\n    Loop\n    F = n"
            ),
            "Integer|4"
        );
        assert_eq!(expr("CBool(Null)"), "ERR|94");
    }

    #[test]
    fn fuzz_not_null_propagates_until_observed() {
        assert_eq!(expr("IsNull(Not Null)"), "Boolean|True");
        assert_eq!(
            run("    Dim a, b\n    a = Not Null\n    b = \"x\" & a\n    F = b"),
            "String|x"
        );
    }

    #[test]
    fn the_words_true_and_false_coerce_on_the_integer_path_only() {
        assert_eq!(expr("\"True\" Xor 1"), "Integer|-2");
        assert_eq!(expr("\"False\" Xor 1"), "Integer|1");
        assert_eq!(expr("\"True\" \\ 1"), "Integer|-1");
        assert_eq!(expr("\"True\" Mod 2"), "Integer|-1");
        assert_eq!(expr("CBool(\"True\")"), "Boolean|True");
        assert_eq!(expr("\"true\" Xor 1"), "Integer|-2");
        assert_eq!(expr("\"TRUE\" Xor 1"), "Integer|-2");
        assert_eq!(expr("Not \"True\""), "Boolean|False");

        assert_eq!(expr("True Eqv \"True\""), "ERR|13");
        assert_eq!(expr("\"True\" Eqv True"), "ERR|13");
        assert_eq!(expr("True Eqv CStr(True)"), "ERR|13");
        assert_eq!(
            run("    Dim a\n    a = 3.75\n    F = (IsNumeric(a) Eqv CStr(True))"),
            "ERR|13"
        );
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

        for e in [
            "\"True\" + 1",
            "\"False\" + 1",
            "\"True\" * 2",
            "CDbl(\"True\")",
        ] {
            assert_eq!(expr(e), "ERR|13", "for {e}");
        }
        assert_eq!(expr("IsNumeric(\"True\")"), "Boolean|False");

        assert_eq!(expr("Trim((1 >= 2)) Xor 5"), "Integer|5");
    }

    #[test]
    fn a_string_outside_double_range_fails_to_convert() {
        assert_eq!(
            run("    Dim a\n    a = \"1E+2923\"\n    F = (a ^ 255)"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a\n    a = \"1E400\"\n    F = (a + 1)"),
            "ERR|6"
        );
        assert_eq!(
            run("    Dim a\n    a = \"255\"\n    F = (a ^ 255)"),
            "ERR|6"
        );
        assert_eq!(run("    Dim a\n    a = 255\n    F = (a ^ 255)"), "ERR|6");
    }

    #[test]
    fn an_empty_string_never_coerces_to_a_number() {
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
        assert_eq!(expr("True Or \"2147483648\""), "ERR|6");
        assert_eq!(expr("1 And \"2147483648\""), "ERR|6");
        assert_eq!(expr("True Or \"3.752147483647\""), "Long|-1");
        assert_eq!(expr("1 And \"12\""), "Long|0");
    }

    #[test]
    fn int_div_and_mod_range_check_their_operands_not_just_the_result() {
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
        assert_eq!(expr("(-1) ^ 1.5"), "ERR|5");
        assert_eq!(expr("(-8) ^ (1 / 3)"), "ERR|5");
        assert_eq!(expr("(-2) ^ 2"), "Double|4");
        assert_eq!(expr("(-2) ^ 3"), "Double|-8");
    }

    #[test]
    fn select_case_matches_a_numeric_case_against_a_string_subject() {
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
        assert_eq!(
            run("    Dim c\n    For c = 5 To 1\n    Next c\n    F = c"),
            "Integer|5"
        );
        assert_eq!(
            run("    Dim c\n    For c = 3 To 1 Step -1\n    Next c\n    F = c"),
            "Integer|0"
        );
        assert_eq!(
            run("    Dim c\n    For c = 1 To 3\n        Exit For\n    Next c\n    F = c"),
            "Integer|1"
        );
    }

    #[test]
    fn count_arguments_round_rather_than_truncate() {
        assert_eq!(expr("Len(Space(2.6))"), "Long|3");
        assert_eq!(expr("Space(-1)"), "ERR|5");
        assert_eq!(expr("String(-1, \"x\")"), "ERR|5");
        assert_eq!(expr("Left(\"abc\", -1)"), "ERR|5");
        assert_eq!(expr("Right(\"abc\", 99)"), "String|abc");
        assert_eq!(expr("InStr(0, \"abc\", \"b\")"), "ERR|5");
        assert_eq!(expr("String(2, 65)"), "String|AA");
    }

    #[test]
    fn host_object_access_errors_rather_than_silently_doing_nothing() {
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
        assert_eq!(run("    With x\n        F = .a\n    End With"), "ERR|424");
        assert_eq!(expr("x.Name"), "ERR|424");
        assert_eq!(expr("x Is Nothing"), "ERR|424");
    }

    #[test]
    fn an_unknown_function_is_reported_not_ignored() {
        assert_eq!(expr("NoSuchFunction(1)"), "ERR|35");
    }

    #[test]
    fn class_module_instantiation_properties_and_methods() {
        let class_src = "Attribute VB_Name = \"Person\"\n\
                         Private m_name As String\n\
                         Public Property Get Name() As String\n\
                             Name = m_name\n\
                         End Property\n\
                         Public Property Let Name(val As String)\n\
                             m_name = val\n\
                         End Property\n\
                         Public Function Greet() As String\n\
                             Greet = \"Hello, \" & Me.Name\n\
                         End Function\n";

        let main_src = "Attribute VB_Name = \"Main\"\n\
                        Function Test()\n\
                            Dim p As Person\n\
                            Set p = New Person\n\
                            p.Name = \"Alice\"\n\
                            Test = p.Greet() & \"|\" & TypeName(p) & \"|\" & (TypeOf p Is Person) & \"|\" & (TypeOf p Is Object)\n\
                        End Function\n";

        let p_cls = VbaModule {
            name: "Person".to_string(),
            kind: VbaModuleKind::Class,
            source: class_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };
        let p_main = VbaModule {
            name: "Main".to_string(),
            kind: VbaModuleKind::Standard,
            source: main_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };

        let mut interp = Interpreter::from_modules(vec![p_cls, p_main], Some("Main"));
        let res = interp.run("Test", Vec::new()).unwrap();
        assert_eq!(
            res,
            Variant::Str("Hello, Alice|Person|True|True".to_string())
        );
    }

    #[test]
    fn class_module_lifecycle_and_auto_new() {
        let class_src = "Attribute VB_Name = \"Counter\"\n\
                         Public Value As Long\n\
                         Private Sub Class_Initialize()\n\
                             Value = 100\n\
                         End Sub\n\
                         Private Sub Class_Terminate()\n\
                             Value = 0\n\
                         End Sub\n";

        let main_src = "Attribute VB_Name = \"Main\"\n\
                        Function TestAutoNew()\n\
                            Dim c As New Counter\n\
                            Dim v1 As Long, v2 As Long\n\
                            v1 = c.Value\n\
                            c.Value = 200\n\
                            Set c = Nothing\n\
                            v2 = c.Value\n\
                            TestAutoNew = v1 & \"|\" & v2\n\
                        End Function\n";

        let p_cls = VbaModule {
            name: "Counter".to_string(),
            kind: VbaModuleKind::Class,
            source: class_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };
        let p_main = VbaModule {
            name: "Main".to_string(),
            kind: VbaModuleKind::Standard,
            source: main_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };

        let mut interp = Interpreter::from_modules(vec![p_cls, p_main], Some("Main"));
        let res = interp.run("TestAutoNew", Vec::new()).unwrap();
        assert_eq!(res, Variant::Str("100|100".to_string()));
    }

    #[test]
    fn class_default_member_dispatch() {
        let class_src = "Attribute VB_Name = \"Bag\"\n\
                         Private m_val As Long\n\
                         Public Property Get Item(idx As Long) As Long\n\
                             Attribute Item.VB_UserMemId = 0\n\
                             Item = m_val * idx\n\
                         End Property\n\
                         Public Property Let Item(idx As Long, val As Long)\n\
                             Attribute Item.VB_UserMemId = 0\n\
                             m_val = val + idx\n\
                         End Property\n";

        let main_src = "Attribute VB_Name = \"Main\"\n\
                        Function TestDefault()\n\
                            Dim b As Bag\n\
                            Set b = New Bag\n\
                            b(2) = 10\n\
                            TestDefault = b(3)\n\
                        End Function\n";

        let p_cls = VbaModule {
            name: "Bag".to_string(),
            kind: VbaModuleKind::Class,
            source: class_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };
        let p_main = VbaModule {
            name: "Main".to_string(),
            kind: VbaModuleKind::Standard,
            source: main_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };

        let mut interp = Interpreter::from_modules(vec![p_cls, p_main], Some("Main"));
        let res = interp.run("TestDefault", Vec::new()).unwrap();
        assert_eq!(res, Variant::Integer(36));
    }

    #[test]
    fn custom_events_and_with_events() {
        let emitter_src = "Attribute VB_Name = \"Emitter\"\n\
                           Public Event OnChange(ByRef num As Long)\n\
                           Public Sub Trigger(n As Long)\n\
                               RaiseEvent OnChange(n)\n\
                           End Sub\n";

        let listener_src = "Attribute VB_Name = \"Listener\"\n\
                            Public Received As Long\n\
                            Public WithEvents em As Emitter\n\
                            Private Sub em_OnChange(ByRef num As Long)\n\
                                Received = num * 2\n\
                                num = num + 1\n\
                            End Sub\n\
                            Public Function TestEvent()\n\
                                Set em = New Emitter\n\
                                Dim x As Long\n\
                                x = 21\n\
                                em.Trigger x\n\
                                TestEvent = Received & \"|\" & x\n\
                            End Function\n";

        let p_em = VbaModule {
            name: "Emitter".to_string(),
            kind: VbaModuleKind::Class,
            source: emitter_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };
        let p_lis = VbaModule {
            name: "Listener".to_string(),
            kind: VbaModuleKind::Standard,
            source: listener_src.to_string(),
            bound_sheet_id: None,
            prefix_bytes: Vec::new(),
            cached_compressed_source: None,
            module_cookie: 0,
        };

        let mut interp = Interpreter::from_modules(vec![p_em, p_lis], Some("Listener"));
        let res = interp.run("TestEvent", Vec::new()).unwrap();
        assert_eq!(res, Variant::Str("42|21".to_string()));
    }

    #[test]
    fn host_worksheet_change_events() {
        let mut wb = crate::core::WorkbookManager::new_empty().unwrap();
        wb.ensure_vba_project().unwrap();
        let sheet_id = wb.sheets[0].id;

        let sheet1_src = "Attribute VB_Name = \"Sheet1\"\n\
                          Private Sub Worksheet_Change(ByVal Target As Range)\n\
                              If Target.Address = \"$A$1\" Then\n\
                                  Application.EnableEvents = False\n\
                                  Range(\"B1\").Value = Target.Value * 10\n\
                                  Application.EnableEvents = True\n\
                              End If\n\
                          End Sub\n";

        let main_src = "Attribute VB_Name = \"Main\"\n\
                        Sub TriggerChange()\n\
                            Range(\"A1\").Value = 5\n\
                        End Sub\n";

        wb.add_vba_module(
            "Sheet1".to_string(),
            VbaModuleKind::Document,
            sheet1_src.to_string(),
            Some(sheet_id),
        )
        .unwrap();

        wb.add_vba_module(
            "Main".to_string(),
            VbaModuleKind::Standard,
            main_src.to_string(),
            None,
        )
        .unwrap();

        let res = wb.run_macro(Some("Main"), "TriggerChange", &[]).unwrap();
        assert!(res.mutated);
        assert_eq!(
            wb.sheets[0].get_display_string(&crate::core::CellRef::new(0, 0)),
            "5"
        );
        assert_eq!(
            wb.sheets[0].get_display_string(&crate::core::CellRef::new(0, 1)),
            "50"
        );
    }

    #[test]
    fn open_events_execution() {
        let mut wb = crate::core::WorkbookManager::new_empty().unwrap();
        wb.ensure_vba_project().unwrap();

        let thisworkbook_src = "Attribute VB_Name = \"ThisWorkbook\"\n\
                                Private Sub Workbook_Open()\n\
                                    Range(\"A1\").Value = \"Opened\"\n\
                                End Sub\n\
                                Private Sub Workbook_BeforeClose(Cancel As Boolean)\n\
                                    Cancel = True\n\
                                End Sub\n";

        let mod_src = "Attribute VB_Name = \"Module1\"\n\
                       Public Sub Auto_Open()\n\
                           Range(\"A2\").Value = \"Auto\"\n\
                       End Sub\n";

        wb.add_vba_module(
            "ThisWorkbook".to_string(),
            VbaModuleKind::Document,
            thisworkbook_src.to_string(),
            None,
        )
        .unwrap();

        wb.add_vba_module(
            "Module1".to_string(),
            VbaModuleKind::Standard,
            mod_src.to_string(),
            None,
        )
        .unwrap();

        let res = wb.run_open_events().unwrap();
        assert!(res.mutated);
        assert_eq!(
            wb.sheets[0].get_display_string(&crate::core::CellRef::new(0, 0)),
            "Opened"
        );
        assert_eq!(
            wb.sheets[0].get_display_string(&crate::core::CellRef::new(1, 0)),
            "Auto"
        );
    }
}
