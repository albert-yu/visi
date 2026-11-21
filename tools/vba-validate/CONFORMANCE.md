_The following content is LLM-generated._

# VBA Specification Conformance Matrix

Status legend: **Supported**, **Partial**, **Out of Scope (Sandboxed)**.

| Area | Spec Reference | Status | Notes |
| --- | --- | --- | --- |
| RLE Compression & Raw Chunk Fallback | `[MS-OVBA] §2.4.1` | Supported | `ovba::compress` falls back to raw uncompressed chunks on incompressible 4096-byte blocks |
| Codepage / MBCS Encoding (`PROJECTCODEPAGE`) | `[MS-OVBA] §2.3.4.2.1.4` | Supported | CP1252, MacRoman, UTF-8 decoding/encoding; synthetic donor uses Win32 + CP1252 (`0x04E4`) |
| `dir`, `PROJECT` (`Class=`/`Module=`/`Document=`), `PROJECTwm` | `[MS-OVBA] §2.3` | Supported | Validated by `vba_validate::validate_vba_project_bin`; preserves donor metadata & opaque storages |
| Digital Signatures (`vbaProjectSignature.bin`) | `[MS-OVBA] §2.3.2` | Supported | Stale signature parts are stripped on VBA mutation; idempotent OPC XML patching |
| `ByRef` / `ByVal`, Named Args, `Optional`, `IsMissing`, `ParamArray` | `[MS-VBAL] §5.3.1 / §5.6.13` | Supported | Parenthesized `(x)` forces `ByVal`; write-back for locals, globals, array elements, UDT fields |
| Primitive Types & Coercion (`Byte`, `Decimal`, `LongLong`, `String * N`) | `[MS-VBAL] §2.1 / §5.5` | Supported | Banker's rounding and overflow enforcement on declared variables, params, and return values |
| N-Dimensional Arrays, `Option Base`, `ReDim [Preserve]`, `Erase` | `[MS-VBAL] §5.4.3` | Supported | Arbitrary `(lower..=upper)` bounds; fixed vs dynamic array rules |
| `Type ... End Type` (UDTs) & `Enum` | `[MS-VBAL] §4.2` | Supported | Value-copy UDT semantics, qualified & unqualified enum constants |
| `Like`, `Option Compare Text|Binary`, `Option Explicit`, `#If` | `[MS-VBAL] §3.4 / §5.6.9.4` | Supported | Static `Option Explicit` in `resolve.rs`; `#If` supported at module and procedure scope |
| `On Error`, `Resume` (retry/next/label), `On Error GoTo -1`, `Err.*` | `[MS-VBAL] §5.4.5` | Supported | `GoSub`/`Return`, `On n GoTo|GoSub`, `Mid() =`, `Implements`, `End` |
| Standard Library (`VBA.*`) | `[MS-VBAL] §6.1` | Supported | Strings, Formatting, Date/Time, Math, `Rnd`/`Randomize`, Financial, Arrays, `VBA.Collection`, `vb*` constants |
| In-Memory COM Emulation | Excel Macro Compatibility | Supported | `Scripting.Dictionary`, `VBScript.RegExp` (`MatchCollection`/`Match`/`SubMatches`), `Debug.Print` capture |
| Excel Object Model (`Application`, `Workbook`, `Worksheet`, `Range`, `Names`, `ListObject`, `PivotTable`) | Excel VBA Reference | Supported | `Range.End`, `CurrentRegion`, `UsedRange`, `Copy`, `Clear*`, `HasFormula`, `Intersect`, `Union`, `Evaluate`, `Xl*` constants, 2D array writes |
| Worksheet Formula UDFs | Excel VBA Reference | Supported | Opt-in via `evaluate_with_macros()` / `visi eval --enable-macros` with interactive TTY prompt |
| OS / Native DLL / File I/O (`Declare`, `Shell`, `Open`, `Kill`) | `[MS-VBAL] §5.4.4` | Out of Scope (Sandboxed) | Parsed for syntax checks; blocked at runtime by design |
