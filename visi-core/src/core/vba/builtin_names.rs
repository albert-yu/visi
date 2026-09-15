/// [LLM-generated] Whether `name` is a VBA or Excel-host built-in, matched
/// case-insensitively as VBA does.
pub(super) fn is_builtin(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    BUILTIN_NAMES.contains(&lower.as_str())
}

/// [LLM-generated] Every built-in name, lowercased, grouped by where it comes from.
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
    "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint", "clng", "clnglng", "clngptr",
    "csng", "cstr", "cvar", "cvdate", "cverr", "hex", "oct", "val",
    "abs", "atn", "cos", "exp", "fix", "int", "log", "rnd", "round", "sgn", "sin", "sqr", "tan",
    "asc", "ascb", "ascw", "chr", "chrb", "chrw", "filter", "format", "formatcurrency",
    "formatdatetime", "formatnumber", "formatpercent", "instr", "instrb", "instrrev", "join",
    "lcase", "left", "leftb", "len", "lenb", "ltrim", "mid", "midb", "replace", "right", "rightb",
    "rtrim", "space", "split", "str", "strcomp", "strconv", "string", "strreverse", "trim",
    "ucase",
    "date", "dateadd", "datediff", "datepart", "dateserial", "datevalue", "day", "hour", "minute",
    "month", "monthname", "now", "second", "time", "timer", "timeserial", "timevalue", "weekday",
    "weekdayname", "year",
    "isarray", "isdate", "isempty", "iserror", "ismissing", "isnull", "isnumeric", "isobject",
    "typename", "vartype",
    "array", "lbound", "ubound",
    "ddb", "fv", "ipmt", "irr", "mirr", "nper", "npv", "pmt", "ppmt", "pv", "rate", "sln", "syd",
    "chdir", "chdrive", "close", "curdir", "dir", "eof", "fileattr", "filecopy", "filedatetime",
    "filelen", "freefile", "get", "getattr", "input", "inputb", "kill", "line", "loc", "lock",
    "lof", "mkdir", "name", "open", "print", "put", "reset", "rmdir", "savepicture", "seek",
    "setattr", "unlock", "width", "write",
    "appactivate", "beep", "callbyname", "choose", "command", "createobject", "deletesetting",
    "doevents", "environ", "getallsettings", "getobject", "getsetting", "iif", "inputbox", "load",
    "msgbox", "partition", "qbcolor", "randomize", "rgb", "savesetting", "sendkeys", "shell",
    "spc", "switch", "tab", "unload",
    "debug", "err", "error", "raise",
    "lset", "objptr", "rset", "strptr", "varptr",
    "collection",
    "excel",
    "office",
    "stdole",
    "vba",
    "vbacceleratorbuttons",
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

    /// [LLM-generated] Every intrinsic the interpreter implements must also be a name this
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
