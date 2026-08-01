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

use serde::{Deserialize, Serialize};

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
}

impl VbaModule {
    pub fn is_document(&self) -> bool {
        self.kind == VbaModuleKind::Document
    }
}

/// A workbook's VBA project: its modules plus the raw material needed to
/// patch (not rebuild from scratch) a `vbaProject.bin` on export.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct VbaProject {
    /// Project ID GUID, e.g. `"{7B4E3A2C-1F5D-4A6B-9C8E-2D3F4A5B6C7D}"`.
    /// Must stay internally consistent with `raw_donor`'s `PROJECT` stream
    /// -- if `CMG`/`DPB`/`GC` protection-state lines are ever reintroduced,
    /// they must correspond to this exact ID or Excel reports the whole
    /// project "unviewable" (a real finding from the POC, not a
    /// hypothetical).
    pub project_id: String,
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
        }
    }

    pub fn find_module(&self, name: &str) -> Option<&VbaModule> {
        self.modules.iter().find(|m| m.name.eq_ignore_ascii_case(name))
    }

    pub fn find_module_mut(&mut self, name: &str) -> Option<&mut VbaModule> {
        self.modules
            .iter_mut()
            .find(|m| m.name.eq_ignore_ascii_case(name))
    }

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
        return Err(format!(
            "Module name '{}' must start with a letter",
            name
        ));
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
                },
                VbaModule {
                    name: "Module1".to_string(),
                    kind: VbaModuleKind::Standard,
                    source: "Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n"
                        .to_string(),
                    bound_sheet_id: None,
                    prefix_bytes: vec![0xBB; 16],
                },
            ],
            raw_donor: Vec::new(),
            seed_prefix_bytes: Vec::new(),
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
        project.find_module_mut("Module1").unwrap().source = "Attribute VB_Name = \"Module1\"\r\nSub Bar()\r\nEnd Sub\r\n".to_string();
        assert_eq!(project.find_module("Module1").unwrap().prefix_bytes, original_prefix);
    }
}
