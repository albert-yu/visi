//! VBA macro project data model.
//!
//! A `VbaProject` is workbook-level (like `Chart`/`PivotTable`), not
//! sheet-scoped like `ExcelTable`, since it's a single `vbaProject.bin` part
//! per workbook holding potentially many modules, some of which (document
//! modules) happen to bind to individual sheets.
//!
//! Unlike tables/pivots, round-tripping this through xlsx doesn't mean
//! re-deriving every byte from these fields on export: `raw_donor` holds the
//! `vbaProject.bin` bytes export (`vba_xlsx.rs`) patches only what changed
//! into, rather than synthesizing a full CFB container from scratch every
//! time. For a project imported from a real file, that's the file's own
//! original bytes (preserving whatever PROJECTREFERENCES it already had --
//! e.g. MSForms, Office -- which this codebase doesn't yet synthesize). For
//! a brand-new project, `VbaProject::new_empty` builds `raw_donor` (and the
//! per-module `prefix_bytes` new modules borrow) entirely synthetically via
//! `vba_synth.rs`, with no real Excel-authored file involved. See
//! `vba_xlsx.rs` and `vba_synth.rs` for why that used to require one, and
//! the design notes in this crate's VBA feature plan for the full rationale
//! (proven via a scratchpad proof-of-concept against real Excel).

// The syntax layer. These are `#[doc(hidden)] pub` for the same reason
// `ovba` and `vba_xlsx` are: `visi-core/fuzz`'s `vba_parse` target needs to
// reach `parse_module` from outside the crate. The supported surface is
// [`check_syntax`] and [`ModuleSyntax`] below, which is what `core`'s
// `pub use` list carries -- the AST is an implementation detail until the
// interpreter phases need it, and pinning its shape now would be a semver
// commitment made a phase too early.
#[doc(hidden)]
pub mod ast;
#[doc(hidden)]
pub mod builtins;
pub(crate) mod color;
#[doc(hidden)]
pub mod host;
#[doc(hidden)]
pub mod interp;
#[doc(hidden)]
pub mod lexer;
#[doc(hidden)]
pub mod parser;
#[doc(hidden)]
pub mod value;

use crate::{Error, ObjectKind};
use serde::{Deserialize, Serialize};

/// What [`check_syntax`] found in a module that parsed.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
#[non_exhaustive]
pub struct ModuleSyntax {
    /// The names of every `Sub`, `Function` and `Property` declared, in source
    /// order. Procedures inside a `#If` branch are all included: which branch
    /// is live depends on `#Const` values, which parsing alone cannot decide.
    pub procedures: Vec<String>,
}

/// Checks a VBA module's source for syntax errors.
///
/// This is Phase 0 of the plan in `docs/vba-macro-support.md`: it answers
/// whether the source *parses*, and nothing more. It does not resolve names,
/// check types, or evaluate anything, so it will accept a module that fails
/// at run time -- and, being an independent implementation, it may still
/// differ from Excel's own compiler at the edges.
///
/// ```
/// use visi_core::core::check_syntax;
/// assert!(check_syntax("Sub Hello()\n    MsgBox \"hi\"\nEnd Sub\n").is_ok());
/// assert!(check_syntax("Sub Hello()\n").is_err());
/// ```
pub fn check_syntax(source: &str) -> Result<ModuleSyntax, Error> {
    let module = parser::parse_module(source).map_err(|e| Error::VbaSyntax {
        message: e.message,
        module: None,
        line: e.pos.line,
        column: e.pos.col,
    })?;
    Ok(ModuleSyntax {
        procedures: module.procedures().iter().map(|p| p.name.clone()).collect(),
    })
}

/// The outcome of running a VBA procedure: its return value, rendered the way
/// VBA would render it, plus the subtype name `TypeName()` reports.
///
/// Both halves matter. An interpreter that computes the right number with the
/// wrong subtype has a real bug -- `1 + 1` is an `Integer` and `1 / 1` is a
/// `Double` -- so the differential fuzzer compares the type as well as the
/// value.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct RunOutcome {
    /// `TypeName()` of the returned value.
    pub type_name: String,
    /// `CStr()` of the returned value, or `None` where VBA itself cannot
    /// stringify it (`Null`).
    pub value: Option<String>,
    /// Whether the run changed the workbook.
    ///
    /// Always `false` from [`run_macro`], which has no workbook to change.
    /// From [`crate::core::WorkbookManager::run_macro`] this is what tells a caller
    /// whether it has something worth saving -- and, for the `visi` CLI,
    /// whether discarding the result silently would be a data loss rather
    /// than a no-op.
    pub mutated: bool,
}

/// Turns command-line argument text into the `Variant`s a procedure receives.
///
/// Arguments arrive as text -- they come from a CLI or a fuzz harness -- and
/// are given the type VBA would give the same literal, so `-a 1` is an
/// `Integer` and `-a 1.5` a `Double`.
fn parse_args(args: &[&str]) -> Vec<value::Variant> {
    args.iter()
        .map(|a| match value::parse_vba_number(a) {
            Ok(n) if !a.trim().is_empty() => {
                value::Variant::from_literal(n, a.contains('.') || a.contains(['e', 'E']))
            }
            _ => value::Variant::Str((*a).to_string()),
        })
        .collect()
}

fn to_outcome(result: value::Variant, mutated: bool) -> RunOutcome {
    RunOutcome {
        type_name: result.type_name().to_string(),
        value: result.to_vba_string().ok(),
        mutated,
    }
}

fn parse_or_error(source: &str, module: Option<&str>) -> Result<ast::Module, Error> {
    parser::parse_module(source).map_err(|e| Error::VbaSyntax {
        message: e.message,
        module: module.map(str::to_string),
        line: e.pos.line,
        column: e.pos.col,
    })
}

fn to_runtime_error(e: value::VbaError) -> Error {
    Error::VbaRuntime {
        message: e.description,
        number: e.number,
    }
}

impl crate::core::WorkbookManager {
    /// Runs one of this workbook's own VBA procedures **against** this
    /// workbook.
    ///
    /// Phase 2 of `docs/vba-macro-support.md`, and the entry point that
    /// separates it from Phase 1: the interpreter borrows the workbook for
    /// the duration, so a macro can read and write cells, walk the sheets,
    /// and call worksheet functions. [`run_macro`] stays as the text-only
    /// form -- it is what `visi_core.run_macro` and `fuzz/fuzz_vba.py` drive,
    /// and a macro that touches no workbook has no reason to need one.
    ///
    /// `module` picks which module to take the procedure from; `None`
    /// searches every module for one that declares it, which is the common
    /// single-module case. Resolving it here rather than in each caller is
    /// deliberate: the CLI and the Python bindings would otherwise each have
    /// their own copy of the rule, and only `fuzz/test_backend_parity.py`
    /// would notice them drifting apart.
    ///
    /// **This executes code the workbook's author wrote.** Nothing calls it
    /// implicitly -- not loading a file, not evaluating formulas, and not a
    /// `Workbook_Open` handler. See the security posture in the feature plan.
    ///
    /// The workbook is left recalculated, so a caller that saves afterwards
    /// writes the values the macro itself would have read.
    pub fn run_macro(
        &mut self,
        module: Option<&str>,
        procedure: &str,
        args: &[&str],
    ) -> Result<RunOutcome, Error> {
        let source = self.macro_source_for(module, procedure)?;
        let parsed = parse_or_error(&source, module)?;
        let args = parse_args(args);

        let host = host::Host::new(self).map_err(to_runtime_error)?;
        let mut interp = interp::Interpreter::new(parsed).with_host(host);
        let result = interp.run(procedure, args);
        // The recalculation runs whether or not the procedure succeeded: a
        // macro that wrote three cells and then raised has still written
        // them, and leaving the workbook holding stale computed values would
        // make the failure look like corruption.
        interp.finish();
        let mutated = interp.mutated();
        let result = result.map_err(to_runtime_error)?;
        Ok(to_outcome(result, mutated))
    }

    /// The source text to run, resolving `module` the way
    /// [`WorkbookManager::run_macro`] documents.
    fn macro_source_for(&self, module: Option<&str>, procedure: &str) -> Result<String, Error> {
        let project = self
            .vba_project
            .as_ref()
            .ok_or_else(|| Error::not_found(ObjectKind::VbaModule, module.unwrap_or(procedure)))?;
        let available = || project.modules.iter().map(|m| m.name.clone()).collect();
        if let Some(name) = module {
            return project
                .find_module(name)
                .map(|m| m.source.clone())
                .ok_or_else(|| Error::not_found_among(ObjectKind::VbaModule, name, available()));
        }
        project
            .modules
            .iter()
            // A module that does not parse is skipped rather than fatal: it
            // cannot be the one declaring the procedure, and reporting its
            // syntax error here would blame the wrong module entirely.
            .find(|m| {
                m.check_syntax().is_ok_and(|s| {
                    s.procedures
                        .iter()
                        .any(|p| p.eq_ignore_ascii_case(procedure))
                })
            })
            .map(|m| m.source.clone())
            .ok_or_else(|| {
                Error::not_found_among(
                    ObjectKind::VbaModule,
                    format!("a module declaring '{procedure}'"),
                    available(),
                )
            })
    }
}

/// Parses `source` and runs one of its procedures.
///
/// Phase 1 of `docs/vba-macro-support.md`: expressions, control flow,
/// `Sub`/`Function` calls and `On Error`. There is **no host object model**,
/// so anything touching a workbook raises a run-time error naming what it
/// was rather than silently doing nothing.
///
/// Execution is bounded -- a statement budget stops a runaway loop and a
/// depth limit stops unbounded recursion -- because this runs source the
/// caller did not necessarily write.
///
/// ```
/// use visi_core::core::run_macro;
/// let src = "Function Add2(a, b)\n    Add2 = a + b\nEnd Function\n";
/// let out = run_macro(src, "Add2", &["1", "2"]).unwrap();
/// assert_eq!(out.type_name, "Integer");
/// assert_eq!(out.value.as_deref(), Some("3"));
/// ```
pub fn run_macro(source: &str, procedure: &str, args: &[&str]) -> Result<RunOutcome, Error> {
    let module = parser::parse_module(source).map_err(|e| Error::VbaSyntax {
        message: e.message,
        module: None,
        line: e.pos.line,
        column: e.pos.col,
    })?;
    let result = interp::Interpreter::new(module)
        .run(procedure, parse_args(args))
        .map_err(to_runtime_error)?;

    Ok(to_outcome(result, false))
}

impl VbaModule {
    /// Checks this module's source, naming it in any error.
    ///
    /// The name matters more than it looks: a workbook can hold many modules
    /// and `visi macro check` reports on all of them, so an error that does
    /// not say which one it came from is close to useless.
    pub fn check_syntax(&self) -> Result<ModuleSyntax, Error> {
        check_syntax(&self.source).map_err(|e| match e {
            Error::VbaSyntax {
                message,
                line,
                column,
                ..
            } => Error::VbaSyntax {
                message,
                module: Some(self.name.clone()),
                line,
                column,
            },
            other => other,
        })
    }
}

/// What kind of VBA module a [`VbaModule`] is, which decides how it binds to
/// the workbook.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum VbaModuleKind {
    /// A `.bas`-equivalent module with no host object binding.
    Standard,
    /// A `.cls`-equivalent module (not validated end-to-end against real
    /// Excel yet -- see the feature plan's open-risk notes).
    Class,
    /// `ThisWorkbook` or a worksheet's code-behind module. Must correspond
    /// 1:1 with an existing sheet (or the workbook itself) via
    /// `bound_sheet_id`, mirroring Excel's own codeName wiring.
    Document,
}

/// A single VBA module's editable content plus the opaque bytes needed to
/// keep Excel happy on export.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct VbaModule {
    /// VB_Name -- must satisfy `validate_vba_module_name`.
    pub name: String,
    /// What kind of module this is, and so how it binds to the workbook.
    pub kind: VbaModuleKind,
    /// Plain VBA source text (no compression, no Attribute-line management
    /// beyond what the caller writes -- callers are expected to include the
    /// `Attribute VB_Name = "..."` line themselves, matching how real
    /// Excel-authored module streams are shaped).
    pub source: String,
    /// Required iff `kind == Document`: the sheet this module's code
    /// belongs to (or `None`/ignored for `ThisWorkbook`, which isn't tied to
    /// a specific sheet). Kept as a stable id (not a name) so sheet renames
    /// don't silently orphan the binding -- deliberately NOT cascaded the
    /// other direction (renaming this module does not rename the sheet, and
    /// vice versa; Excel allows the two names to diverge).
    pub bound_sheet_id: Option<u64>,
    /// Opaque bytes forming the pre-TextOffset "p-code prefix" of this
    /// module's stream. Never reparsed or validated by this codebase --
    /// proven (via the POC) that its *content* doesn't need to correspond
    /// to this module's actual source, only its presence matters, as long
    /// as it's shaped the way real Excel's module loader expects (a
    /// naively zero-filled placeholder of the same length is NOT enough).
    /// For an imported module these are the real bytes read back from the
    /// original file; for a module created in this codebase they're
    /// `vba_synth::synthetic_module_prefix()`'s from-scratch, self-consistent
    /// zero-procedure cache -- see that module's doc comment.
    #[serde(default)]
    pub prefix_bytes: Vec<u8>,
    /// The module stream's MODULECOOKIE record (`0x002C`) value. MS-OVBA
    /// documents this as implementation-specific and ignorable on read, but
    /// this codebase used to blindly overwrite every module's (including
    /// untouched, imported ones') cookie with a hardcoded `0xFFFF` on every
    /// export -- discovered while investigating why every workbook this
    /// codebase produces failed `has vb project` in real Excel, by diffing
    /// a re-exported real donor project's `dir` stream against the
    /// original's record-by-record and finding this was the one place real
    /// data was being discarded and replaced rather than round-tripped
    /// verbatim. Preserved here instead so an imported module's original
    /// value survives re-export.
    #[serde(default = "default_module_cookie")]
    pub module_cookie: u16,
    /// This module stream's already-compressed source, as read back
    /// verbatim from an imported file -- `None` for a module created fresh
    /// in this session (nothing to cache yet). `set_vba_module_source`
    /// clears this whenever `source` is replaced. Export reuses the cached
    /// bytes instead of recompressing `source` from scratch for every
    /// module untouched by the CRUD operation that triggered the save.
    #[serde(default)]
    pub cached_compressed_source: Option<Vec<u8>>,
}

fn default_module_cookie() -> u16 {
    0xFFFF
}

impl VbaModule {
    /// Whether this is a document module -- `ThisWorkbook` or a worksheet's
    /// code-behind -- as opposed to a standard or class module.
    pub fn is_document(&self) -> bool {
        self.kind == VbaModuleKind::Document
    }
}

/// A workbook's VBA project: its modules plus the raw material needed to
/// patch (not rebuild from scratch) a `vbaProject.bin` on export.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct VbaProject {
    /// Project ID GUID, e.g. `"{7B4E3A2C-1F5D-4A6B-9C8E-2D3F4A5B6C7D}"`.
    /// Must stay internally consistent with `protection_lines` -- never
    /// mutated after import/creation, so it always is. If `CMG`/`DPB`/`GC`
    /// protection-state lines are ever made independently settable, they
    /// must correspond to this exact ID or Excel reports the whole project
    /// "unviewable" (a real finding from the POC, not a hypothetical).
    pub project_id: String,
    /// The project's modules, in no particular order. Names are unique
    /// case-insensitively.
    pub modules: Vec<VbaModule>,
    /// The full original `vbaProject.bin` bytes this project was imported
    /// from, or (for a project created fresh in this session)
    /// `vba_synth::synthetic_raw_donor()`'s from-scratch bytes -- export's
    /// patch base. See `vba_xlsx.rs`.
    #[serde(default)]
    pub raw_donor: Vec<u8>,
    /// P-code prefix bytes to donate to the first module ever added to a
    /// project that started with none -- kept separate from `modules`
    /// rather than as a phantom placeholder module, so it never shows up in
    /// `list_vba_modules`/export. Once a project has at least one real
    /// module, new modules instead borrow prefix bytes from an existing
    /// one, and this field goes unused.
    #[serde(default)]
    pub seed_prefix_bytes: Vec<u8>,
    /// `VbaModule::module_cookie` to donate to the first module ever added
    /// to a project that started with none -- same donation scheme as
    /// `seed_prefix_bytes`, see there for why.
    #[serde(default = "default_module_cookie")]
    pub seed_module_cookie: u16,
    /// The donor's original `PROJECT` stream `CMG=`/`DPB=`/`GC=` lines
    /// (joined with `\r\n`), reproduced verbatim on export -- `None` for a
    /// project created fresh in this session, which never had any. See
    /// `vba_xlsx::build_project_stream` for why these must be preserved
    /// rather than dropped.
    #[serde(default)]
    pub protection_lines: Option<String>,
}

impl VbaProject {
    /// A brand-new, empty VBA project with no real Excel-authored file
    /// behind it anywhere -- `raw_donor` and `seed_prefix_bytes` are built
    /// by `vba_synth` entirely from scratch. See `vba_synth`'s doc comment
    /// for why that's now possible.
    pub fn new_empty() -> Self {
        VbaProject {
            project_id: new_project_guid(),
            modules: Vec::new(),
            raw_donor: crate::core::vba_synth::synthetic_raw_donor(),
            seed_prefix_bytes: crate::core::vba_synth::synthetic_module_prefix(),
            seed_module_cookie: default_module_cookie(),
            protection_lines: None,
        }
    }

    /// Finds a module by name, matched case-insensitively as VBA does.
    pub fn find_module(&self, name: &str) -> Option<&VbaModule> {
        self.modules
            .iter()
            .find(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// [`VbaProject::find_module`], mutably.
    pub fn find_module_mut(&mut self, name: &str) -> Option<&mut VbaModule> {
        self.modules
            .iter_mut()
            .find(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// Whether a module of this name already exists, matched
    /// case-insensitively.
    pub fn module_name_taken(&self, name: &str) -> bool {
        self.find_module(name).is_some()
    }
}

/// A GUID-shaped project id (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`) for
/// a brand-new project, built from two `generate_unique_id()` draws rather
/// than duplicating its getrandom/fallback logic.
fn new_project_guid() -> String {
    let hi = crate::core::engine::generate_unique_id();
    let lo = crate::core::engine::generate_unique_id();
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:04X}-{:012X}}}",
        (hi >> 32) as u32,
        (hi >> 16) as u16,
        hi as u16,
        (lo >> 48) as u16,
        lo & 0xFFFF_FFFF_FFFF,
    )
}

/// VBA identifiers: must start with a letter, contain only letters/digits/
/// underscore, and be at most 31 characters (the real VBE module-name
/// limit).
pub fn validate_vba_module_name(name: &str) -> Result<(), String> {
    let trimmed = name.trim();
    if trimmed.is_empty() {
        return Err("Module name cannot be empty".to_string());
    }
    if trimmed.chars().count() > 31 {
        return Err(format!(
            "Module name '{}' exceeds VBA's 31-character limit",
            name
        ));
    }
    let first = trimmed.chars().next().unwrap();
    if !first.is_alphabetic() {
        return Err(format!("Module name '{}' must start with a letter", name));
    }
    if !trimmed.chars().all(|c| c.is_alphanumeric() || c == '_') {
        return Err(format!(
            "Module name '{}' may only contain letters, digits, and underscores",
            name
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_project() -> VbaProject {
        VbaProject {
            project_id: "{00000000-0000-0000-0000-000000000000}".to_string(),
            modules: vec![
                VbaModule {
                    name: "ThisWorkbook".to_string(),
                    kind: VbaModuleKind::Document,
                    source: "Attribute VB_Name = \"ThisWorkbook\"\r\n".to_string(),
                    bound_sheet_id: None,
                    prefix_bytes: vec![0xAA; 16],
                    module_cookie: 0xFFFF,
                    cached_compressed_source: None,
                },
                VbaModule {
                    name: "Module1".to_string(),
                    kind: VbaModuleKind::Standard,
                    source: "Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n"
                        .to_string(),
                    bound_sheet_id: None,
                    prefix_bytes: vec![0xBB; 16],
                    module_cookie: 0xFFFF,
                    cached_compressed_source: None,
                },
            ],
            raw_donor: Vec::new(),
            seed_prefix_bytes: Vec::new(),
            seed_module_cookie: 0xFFFF,
            protection_lines: None,
        }
    }

    #[test]
    fn validate_name_rules() {
        assert!(validate_vba_module_name("Module1").is_ok());
        assert!(validate_vba_module_name("_Bad").is_err());
        assert!(validate_vba_module_name("1Bad").is_err());
        assert!(validate_vba_module_name("").is_err());
        assert!(validate_vba_module_name("Has Space").is_err());
        assert!(validate_vba_module_name("Has-Dash").is_err());
        assert!(validate_vba_module_name(&"A".repeat(32)).is_err());
        assert!(validate_vba_module_name(&"A".repeat(31)).is_ok());
    }

    #[test]
    fn find_module_case_insensitive() {
        let project = sample_project();
        assert!(project.find_module("module1").is_some());
        assert!(project.find_module("MODULE1").is_some());
        assert!(project.find_module("Module2").is_none());
    }

    #[test]
    fn module_name_taken_case_insensitive() {
        let project = sample_project();
        assert!(project.module_name_taken("module1"));
        assert!(!project.module_name_taken("Module2"));
    }

    #[test]
    fn set_source_leaves_prefix_bytes_untouched() {
        let mut project = sample_project();
        let original_prefix = project.find_module("Module1").unwrap().prefix_bytes.clone();
        project.find_module_mut("Module1").unwrap().source =
            "Attribute VB_Name = \"Module1\"\r\nSub Bar()\r\nEnd Sub\r\n".to_string();
        assert_eq!(
            project.find_module("Module1").unwrap().prefix_bytes,
            original_prefix
        );
    }
}
