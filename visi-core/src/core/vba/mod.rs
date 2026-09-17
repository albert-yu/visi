#[doc(hidden)]
pub mod ast;
pub(crate) mod builtin_names;
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
pub(crate) mod resolve;
#[doc(hidden)]
pub mod value;

use crate::{Error, ObjectKind};
use serde::{Deserialize, Serialize};

/// [`check_syntax`] result
#[derive(Debug, Clone, PartialEq, Eq, Default)]
#[non_exhaustive]
pub struct ModuleSyntax {
    /// Every `Sub`, `Function` and `Property` in source order
    pub procedures: Vec<String>,
}

/// Checks a VBA module's source for syntax errors.
///
/// ```
/// use visi_core::core::check_syntax;
/// assert!(check_syntax("Sub Hello()\n    MsgBox \"hi\"\nEnd Sub\n").is_ok());
/// assert!(check_syntax("Sub Hello()\n").is_err());
/// ```
pub fn check_syntax(source: &str) -> Result<ModuleSyntax, Error> {
    let empty = std::collections::HashSet::new();
    check_source(source, None, &resolve::Scope::self_contained(&empty))
}

/// Check source that is part of a larger project
/// where other modules may not be available.
///
/// ```
/// use visi_core::core::{check_syntax, check_syntax_partial};
/// // `DoWork` is declared by some other module of the project.
/// let src = "Sub Caller()\n    DoWork 1\nEnd Sub\n";
/// assert!(check_syntax(src).is_err());
/// assert!(check_syntax_partial(src).is_ok());
/// assert!(check_syntax_partial("Sub Caller()\n").is_err());
/// ```
pub fn check_syntax_partial(source: &str) -> Result<ModuleSyntax, Error> {
    let empty = std::collections::HashSet::new();
    check_source(source, None, &resolve::Scope::partial(&empty))
}

fn check_source(
    source: &str,
    module_name: Option<&str>,
    scope: &resolve::Scope<'_>,
) -> Result<ModuleSyntax, Error> {
    let to_err = |e: parser::ParseError| Error::VbaSyntax {
        message: e.message,
        module: module_name.map(str::to_string),
        line: e.pos.line,
        column: e.pos.col,
    };
    let module = parser::parse_module(source).map_err(to_err)?;
    resolve::check_module(&module, scope).map_err(to_err)?;
    Ok(ModuleSyntax {
        procedures: module.procedures().iter().map(|p| p.name.clone()).collect(),
    })
}

/// The outcome of running a VBA procedure
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct RunOutcome {
    /// `TypeName()`
    pub type_name: String,
    /// `CStr()`, or `None` where VBA itself cannot
    /// stringify it (`Null`).
    pub value: Option<String>,
    /// Whether the run changed the workbook
    pub mutated: bool,
}

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

fn to_outcome(result: value::Variant, mutated: bool, interp: &interp::Interpreter) -> RunOutcome {
    RunOutcome {
        type_name: interp.type_name_of(&result),
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
    /// Runs one of this workbook's own VBA procedures
    pub fn run_macro(
        &mut self,
        module: Option<&str>,
        procedure: &str,
        args: &[&str],
    ) -> Result<RunOutcome, Error> {
        let args = parse_args(args);

        let interp = if let Some(project) = &self.vba_project {
            if let Some(name) = module
                && project.find_module(name).is_none()
            {
                let available = project.modules.iter().map(|m| m.name.clone()).collect();
                return Err(Error::not_found_among(
                    ObjectKind::VbaModule,
                    name,
                    available,
                ));
            }
            interp::Interpreter::from_project(project, module).map_err(to_runtime_error)?
        } else {
            let source = self.macro_source_for(module, procedure)?;
            let parsed = parse_or_error(&source, module)?;
            interp::Interpreter::new(parsed)
        };

        let host = host::Host::new(self).map_err(to_runtime_error)?;
        let mut interp = interp.with_host(host);

        let result = interp.run(procedure, args);
        interp.finish();
        let mutated = interp.mutated();
        let result = result.map_err(to_runtime_error)?;
        Ok(to_outcome(result, mutated, &interp))
    }

    /// Runs startup macro events (`Workbook_Open` in `ThisWorkbook` then `Auto_Open` in standard modules).
    pub fn run_open_events(&mut self) -> Result<RunOutcome, Error> {
        let interp = if let Some(project) = &self.vba_project {
            interp::Interpreter::from_project(project, None).map_err(to_runtime_error)?
        } else {
            return Err(Error::not_found(
                ObjectKind::VbaModule,
                "Workbook_Open or Auto_Open",
            ));
        };

        let host = host::Host::new(self).map_err(to_runtime_error)?;
        let mut interp = interp.with_host(host);

        interp.run_open_events().map_err(to_runtime_error)?;
        interp.finish();
        let mutated = interp.mutated();
        Ok(RunOutcome {
            type_name: "Empty".to_string(),
            value: Some(String::new()),
            mutated,
        })
    }

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
            .find(|m| {
                parser::parse_module(&m.source).is_ok_and(|module| {
                    module
                        .procedures()
                        .iter()
                        .any(|p| p.name.eq_ignore_ascii_case(procedure))
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

/// Parses `source` and runs one of its procedures
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
    let mut interp = interp::Interpreter::new(module);
    let result = interp
        .run(procedure, parse_args(args))
        .map_err(to_runtime_error)?;

    Ok(to_outcome(result, false, &interp))
}

impl VbaModule {
    /// Checks this module's source, naming it in any error
    pub fn check_syntax(&self) -> Result<ModuleSyntax, Error> {
        let empty = std::collections::HashSet::new();
        check_source(
            &self.source,
            Some(&self.name),
            &resolve::Scope::partial(&empty),
        )
    }
}

/// What kind [`VbaModule`] is, which decides how it binds to
/// the workbook.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum VbaModuleKind {
    /// A `.bas`-equivalent module with no host object binding.
    Standard,
    /// A `.cls`-equivalent module).
    /// TODO: fuzz this against real Excel
    Class,
    /// `ThisWorkbook` or a worksheet's code-behind module
    Document,
}

/// A single VBA module's editable content plus the opaque bytes needed to
/// keep Excel happy on export.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct VbaModule {
    /// VB_Name -- must satisfy [`validate_vba_module_name`]
    pub name: String,
    /// What kind of module this is
    pub kind: VbaModuleKind,
    /// Plain VBA source text
    pub source: String,
    /// Required iff `kind == Document`
    pub bound_sheet_id: Option<u64>,
    /// Opaque bytes forming the pre-TextOffset "p-code prefix" of this
    /// module's stream. Has nothing to do with the actual content of the
    /// module
    #[serde(default)]
    pub prefix_bytes: Vec<u8>,
    /// The module stream's MODULECOOKIE record (`0x002C`) value. MS-OVBA
    /// documents this as implementation-specific and ignorable on read.
    /// Preserved here so an imported module's original value survives re-export.
    #[serde(default = "default_module_cookie")]
    pub module_cookie: u16,
    /// This module stream's already-compressed source, as read back
    /// verbatim from an imported file. Is `None` (empty cache) for a freshly
    /// created module.
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

/// VBA project for a workbook
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct VbaProject {
    /// Project ID GUID, e.g. `"{7B4E3A2C-1F5D-4A6B-9C8E-2D3F4A5B6C7D}"`.
    pub project_id: String,
    /// The project's modules, in no particular order
    pub modules: Vec<VbaModule>,
    /// The full original `vbaProject.bin` bytes this project was imported
    /// from `vba_synth::synthetic_raw_donor()`'s from-scratch bytes -- export's
    /// patch base. See `vba_xlsx.rs`.
    #[serde(default)]
    pub raw_donor: Vec<u8>,
    /// P-code prefix bytes to donate to the first module
    /// (weird vibe-coded workaround in order to create a valid module).
    #[serde(default)]
    pub seed_prefix_bytes: Vec<u8>,
    /// Same donation scheme as `seed_prefix_bytes`
    #[serde(default = "default_module_cookie")]
    pub seed_module_cookie: u16,
    /// See [`vba_xlsx::build_project_stream`] for why these must be preserved
    /// rather than dropped.
    #[serde(default)]
    pub protection_lines: Option<String>,
}

impl VbaProject {
    /// Create an empty project
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

    /// Finds a module by name, matched case-insensitively
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

    /// Checks every module, resolving names against the **whole project**
    pub fn check_modules(&self) -> Vec<(String, Result<ModuleSyntax, Error>)> {
        self.check_modules_scoped(true)
    }

    /// [`check_modules`](Self::check_modules) for a project that may contain
    /// other modules
    pub fn check_modules_partial(&self) -> Vec<(String, Result<ModuleSyntax, Error>)> {
        self.check_modules_scoped(false)
    }

    fn check_modules_scoped(&self, complete: bool) -> Vec<(String, Result<ModuleSyntax, Error>)> {
        let mut declared: std::collections::HashSet<String> = std::collections::HashSet::new();
        let parsed: Vec<_> = self
            .modules
            .iter()
            .map(|m| (m, parser::parse_module(&m.source).ok()))
            .collect();
        for (_, module) in &parsed {
            if let Some(module) = module {
                declared.extend(resolve::declared_names(module));
            }
        }

        parsed
            .iter()
            .map(|(m, _)| {
                let scope = resolve::Scope {
                    external: &declared,
                    complete_project: complete,
                };
                (
                    m.name.clone(),
                    check_source(&m.source, Some(&m.name), &scope),
                )
            })
            .collect()
    }
}

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

    fn project_of(sources: &[(&str, &str)]) -> VbaProject {
        let mut project = sample_project();
        project.modules = sources
            .iter()
            .map(|(name, source)| VbaModule {
                name: (*name).to_string(),
                kind: VbaModuleKind::Standard,
                source: (*source).to_string(),
                bound_sheet_id: None,
                prefix_bytes: vec![0xBB; 16],
                module_cookie: 0xFFFF,
                cached_compressed_source: None,
            })
            .collect();
        project
    }

    const CALLER: &str = "Public Sub Caller()\n    DoWork 1\nEnd Sub\n";
    const CALLEE: &str = "Public Sub DoWork(n As Long)\nEnd Sub\n";

    #[test]
    fn partial_scope_accepts_a_call_into_source_not_supplied() {
        assert!(check_syntax(CALLER).is_err());
        assert!(check_syntax_partial(CALLER).is_ok());

        let dup = "Sub Test()\n    Dim x As Long\n    Dim x As Long\nEnd Sub\n";
        assert!(check_syntax(dup).is_err());
        assert!(check_syntax_partial(dup).is_err());
    }

    #[test]
    fn check_modules_resolves_across_siblings() {
        let project = project_of(&[("Module1", CALLER), ("Module2", CALLEE)]);
        for (name, result) in project.check_modules() {
            assert!(result.is_ok(), "{name} should be clean: {result:?}");
        }

        let alone = project_of(&[("Module1", CALLER)]);
        let results = alone.check_modules();
        assert_eq!(results.len(), 1);
        match &results[0].1 {
            Err(Error::VbaSyntax {
                message, module, ..
            }) => {
                assert!(message.contains("DoWork"), "{message}");
                assert_eq!(module.as_deref(), Some("Module1"));
            }
            other => panic!("expected a syntax error, got {other:?}"),
        }

        assert!(alone.check_modules_partial()[0].1.is_ok());
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
