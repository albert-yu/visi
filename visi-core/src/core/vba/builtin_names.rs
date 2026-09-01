//! Names VBA and the Excel host provide, for name resolution in
//! [`super::resolve`].
//!
//! # Why this exists
//!
//! Excel refuses to compile a name used with *call syntax* -- `arr(1, 2)`,
//! or a bare-argument statement `d #1/1/2000#` -- unless it resolves to
//! something callable or indexable. Measured with `fuzz/vba_compile_probe.py`
//! (`undeclared:*` cases); a bare identifier with **no** call syntax is
//! accepted, since `Option Explicit` being off makes it an implicit Variant.
//!
//! To act on that, a checker has to distinguish "this name is nowhere"
//! from "this name is VBA's, not yours." That list is what this module is.
//!
//! # The bias is deliberate: when unsure, include
//!
//! Getting a name wrong is not symmetric.
//!
//! - **Omitting** a real built-in makes `macro check` reject working code --
//!   a false positive, which `docs/vba-macro-support.md` calls the worse
//!   failure mode.
//! - **Including** a name that is not really a built-in only means one
//!   undeclared-name mistake goes unreported -- a false negative.
//!
//! So this list is deliberately over-broad. It carries names whose spelling
//! is contextual (`Name`, `Line`, `Width`, `Get`, `Put` are statements in
//! one position and ordinary property names in another -- see `parser.rs`'s
//! note on VBA's keywords not being reserved), names from libraries a
//! workbook may or may not reference, and the whole Excel `Application`
//! global surface. **Adding a plausible name here needs no justification;
//! removing one does.**
//!
//! It is not, and cannot be, complete: VBA's surface grows with whatever
//! references a project carries (MSForms, Scripting, ADO, ...), which a
//! single module's text does not reveal. That is a bounded, one-directional
//! inaccuracy by the argument above.

/// Whether `name` is a VBA or Excel-host built-in, matched
/// case-insensitively as VBA does.
pub(super) fn is_builtin(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    BUILTIN_NAMES.contains(&lower.as_str())
}

/// Every built-in name, lowercased, grouped by where it comes from.
///
/// Deliberately a **linear** scan in [`is_builtin`] rather than a sorted
/// binary search. Sorting would mean one flat alphabetical run, throwing
/// away the grouping that documents each name's provenance -- and a
/// binary search over a hand-maintained list fails *silently* when an entry
/// lands out of order, which is a false positive on working code, the one
/// direction this module must never fail in. (That is not hypothetical: the
/// first draft here was a binary search over a list grouped exactly like
/// this, and it could not find `MsgBox`.) A couple of hundred string
/// compares once per call target is nothing next to parsing the module.
#[rustfmt::skip]
static BUILTIN_NAMES: &[&str] = &[
    // ---- VBA: conversion ------------------------------------------------
    "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint", "clng", "clnglng", "clngptr",
    "csng", "cstr", "cvar", "cvdate", "cverr", "hex", "oct", "val",
    // ---- VBA: maths -----------------------------------------------------
    "abs", "atn", "cos", "exp", "fix", "int", "log", "rnd", "round", "sgn", "sin", "sqr", "tan",
    // ---- VBA: strings ---------------------------------------------------
    "asc", "ascb", "ascw", "chr", "chrb", "chrw", "filter", "format", "formatcurrency",
    "formatdatetime", "formatnumber", "formatpercent", "instr", "instrb", "instrrev", "join",
    "lcase", "left", "leftb", "len", "lenb", "ltrim", "mid", "midb", "replace", "right", "rightb",
    "rtrim", "space", "split", "str", "strcomp", "strconv", "string", "strreverse", "trim",
    "ucase",
    // ---- VBA: date and time ---------------------------------------------
    "date", "dateadd", "datediff", "datepart", "dateserial", "datevalue", "day", "hour", "minute",
    "month", "monthname", "now", "second", "time", "timer", "timeserial", "timevalue", "weekday",
    "weekdayname", "year",
    // ---- VBA: type inspection -------------------------------------------
    "isarray", "isdate", "isempty", "iserror", "ismissing", "isnull", "isnumeric", "isobject",
    "typename", "vartype",
    // ---- VBA: arrays ----------------------------------------------------
    "array", "lbound", "ubound",
    // ---- VBA: financial -------------------------------------------------
    "ddb", "fv", "ipmt", "irr", "mirr", "nper", "npv", "pmt", "ppmt", "pv", "rate", "sln", "syd",
    // ---- VBA: file and device I/O ---------------------------------------
    // Out of scope for the interpreter, but perfectly compilable, which is
    // the only question this module answers.
    "chdir", "chdrive", "close", "curdir", "dir", "eof", "fileattr", "filecopy", "filedatetime",
    "filelen", "freefile", "get", "getattr", "input", "inputb", "kill", "line", "loc", "lock",
    "lof", "mkdir", "name", "open", "print", "put", "reset", "rmdir", "savepicture", "seek",
    "setattr", "unlock", "width", "write",
    // ---- VBA: interaction and system ------------------------------------
    "appactivate", "beep", "callbyname", "choose", "command", "createobject", "deletesetting",
    "doevents", "environ", "getallsettings", "getobject", "getsetting", "iif", "inputbox", "load",
    "msgbox", "partition", "qbcolor", "randomize", "rgb", "savesetting", "sendkeys", "shell",
    "spc", "switch", "tab", "unload",
    // ---- VBA: errors and debugging --------------------------------------
    "debug", "err", "error", "raise",
    // ---- VBA: pointer and assignment intrinsics --------------------------
    // Hidden but perfectly callable, and `LSet`/`RSet` are statements whose
    // leading word the parser hands over as an ordinary identifier.
    "lset", "objptr", "rset", "strptr", "varptr",
    // ---- VBA: root library objects --------------------------------------
    // `VBA.Left(...)`, `Excel.Range(...)` -- a qualified call's leading name
    // is an ordinary identifier as far as the parser is concerned.
    "collection",
    "excel",
    "office",
    "stdole",
    "vba",
    "vbacceleratorbuttons",
    // ---- Excel: the Application global surface ---------------------------
    // Every one of these is reachable as a bare name in a standard module,
    // because Excel exposes Application's members globally.
    "activecell",
    "activechart",
    "activeprinter",
    "activesheet",
    "activewindow",
    "activeworkbook",
    "addins",
    "application",
    "assistant",
    "calculate",
    "cells",
    "charts",
    "columns",
    "commandbars",
    "creator",
    "dialogs",
    "evaluate",
    "executeexcel4macro",
    "intersect",
    "names",
    "parent",
    "range",
    "rows",
    "run",
    "selection",
    "sheets",
    "shortcutmenus",
    "thisworkbook",
    "toolbars",
    "union",
    "windows",
    "workbooks",
    "worksheetfunction",
    "worksheets",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builtin_names_are_lowercase_and_unique() {
        // `is_builtin` lowercases before comparing, so a mixed-case entry
        // here would be unfindable -- a false positive on working code, the
        // one direction this module must never fail in. Ordering is
        // deliberately *not* required; see `BUILTIN_NAMES`.
        for n in BUILTIN_NAMES {
            assert_eq!(*n, n.to_ascii_lowercase(), "not lowercase: {n}");
        }
        let mut seen = std::collections::HashSet::new();
        for n in BUILTIN_NAMES {
            assert!(seen.insert(*n), "duplicate entry: {n}");
        }
    }

    #[test]
    fn lookup_is_case_insensitive() {
        assert!(is_builtin("MsgBox"));
        assert!(is_builtin("msgbox"));
        assert!(is_builtin("MSGBOX"));
        assert!(is_builtin("Range"));
        assert!(is_builtin("Worksheets"));
    }

    #[test]
    fn a_name_that_is_not_a_builtin_is_not_found() {
        assert!(!is_builtin("arr"));
        assert!(!is_builtin("Helper"));
        assert!(!is_builtin("MyOwnSub"));
    }

    /// Every intrinsic the interpreter implements must also be a name this
    /// registry knows.
    ///
    /// The two lists are maintained separately and answer different
    /// questions -- `builtins.rs` is "what can this engine evaluate?",
    /// this is "what will Excel's compiler accept?" -- but one direction is
    /// not optional: an intrinsic `builtins::call` handles is by definition
    /// a real VBA name, so omitting it here rejects working code. Adding
    /// `Hex`, `Oct` and `Val` to `builtins.rs` without adding them here is
    /// exactly the slip this catches; it had already happened when the test
    /// was written.
    #[test]
    fn every_implemented_intrinsic_is_a_known_builtin() {
        for name in super::super::builtins::implemented_names() {
            assert!(
                is_builtin(name),
                "`{name}` is implemented in builtins.rs but missing from BUILTIN_NAMES, \
                 so `macro check` would reject working code that calls it"
            );
        }
    }
}
