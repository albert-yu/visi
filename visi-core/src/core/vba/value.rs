//! VBA's `Variant` value model and its coercion rules.
//!
//! Phase 1 of `docs/vba-macro-support.md`. Every rule here was measured
//! against real Excel (16.112) with `fuzz/vba_variant_probe.bas`, which
//! returns `TypeName(v) & "|" & CStr(v)` for 64 expressions, rather than
//! taken from documentation. The measurements matter because several are the
//! opposite of what a reasonable implementer would assume:
//!
//! **Overflow promotes at runtime but not between literals.** This one was
//! measured wrong the first time and corrected by `fuzz/fuzz_vba.py`. The
//! literal expression `32767 + 1` really is error 6 -- but
//! `a = 32767 : a + 1` is the `Long` 32768, `a = 2147483647 : a + 1` is the
//! `Double` 2147483648, and `a = 100000 : a * a` is the `Double` 1e10. So
//! arithmetic between two *compile-time constants* uses fixed-width operand
//! types and errors, while anything involving a variable widens
//! Integer -> Long -> Double. [`ArithMode`] carries which of the two applies.
//!
//! **`+` is overloaded on strings.** `"1" + 1` is the `Double` `2` -- the
//! string coerces to a number -- but `"1" + "2"` is the `String` `"12"`,
//! because two string operands make `+` concatenate. `"abc" + 1` is error 13.
//!
//! **`\` and `Mod` round their operands first, and widen.** `7.6 \ 2` is
//! `4`, not `3`: 7.6 rounds (banker's) to 8 first. Its type is `Long`, not
//! `Integer`, because a non-integral operand forces the `Long` conversion.
//!
//! **Rounding is banker's, everywhere.** `CLng(0.5)` is `0`, `CLng(1.5)` is
//! `2`, `CLng(2.5)` is `2`. This is Excel's `ROUND`-to-even, not the
//! round-half-away-from-zero that `f64::round` gives.
//!
//! **`Null` propagates through arithmetic but not through `&`.**
//! `Null + 1` is `Null`; `Null & "a"` is `"a"`.
//!
//! **`Empty` is `0` and `""` at once.** `Empty + 1` is `1`, `Empty & "a"` is
//! `"a"`, and both `Empty = 0` and `Empty = ""` are `True`.
//!
//! **An empty string is not a zero.** `"" = 0`, `"" < 0` and `Not ""` are all
//! error 13, unlike `Empty`, which is genuinely zero. Only a string that
//! parses as a number coerces; `"abc"` and `""` alike are type mismatches.
//!
//! The doc comment on each method names the probe case it came from. Do not
//! "correct" one of these from memory -- re-measure instead.

use std::fmt;
use std::rc::Rc;

use super::host::ObjRef;

/// A VBA runtime error: a number and a description, as `Err.Number` and
/// `Err.Description` expose them.
///
/// Modelled on VBA's own error numbers rather than a Rust enum so that
/// `On Error` handlers, and the differential fuzzer, can compare them
/// directly against Excel's.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaError {
    /// `Err.Number`.
    pub number: i32,
    /// `Err.Description`.
    pub description: String,
}

impl VbaError {
    /// Error 5 -- Invalid procedure call or argument.
    pub fn invalid_call() -> Self {
        Self::new(5, "Invalid procedure call or argument")
    }
    /// Error 6 -- Overflow.
    pub fn overflow() -> Self {
        Self::new(6, "Overflow")
    }
    /// Error 9 -- Subscript out of range.
    pub fn subscript() -> Self {
        Self::new(9, "Subscript out of range")
    }
    /// Error 11 -- Division by zero.
    pub fn div_by_zero() -> Self {
        Self::new(11, "Division by zero")
    }
    /// Error 13 -- Type mismatch.
    pub fn type_mismatch() -> Self {
        Self::new(13, "Type mismatch")
    }
    /// Error 94 -- Invalid use of Null.
    pub fn invalid_null() -> Self {
        Self::new(94, "Invalid use of Null")
    }

    /// An error with an explicit number and description.
    pub fn new(number: i32, description: impl Into<String>) -> Self {
        Self {
            number,
            description: description.into(),
        }
    }
}

impl fmt::Display for VbaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "run-time error {}: {}", self.number, self.description)
    }
}

impl std::error::Error for VbaError {}

/// The result of evaluating VBA, which is either a value or a runtime error.
pub type VResult<T> = Result<T, VbaError>;

/// A VBA value.
///
/// `Byte`, `LongLong` and `Decimal` are deliberately absent: nothing in the
/// implemented scope constructs one, and a variant no path can produce makes
/// every `match` pay for a case that cannot happen. [`Variant::Object`],
/// [`Variant::ErrValue`] and [`Variant::Array`] earned their place in Phase 2,
/// where a cell read produces all three.
#[derive(Debug, Clone, PartialEq)]
pub enum Variant {
    /// An uninitialised variable. Behaves as `0` and `""` depending on
    /// context.
    Empty,
    /// SQL-style unknown. Propagates through arithmetic, is skipped by `&`.
    Null,
    /// `True` is `-1`, not `1` -- which is why `True + 1` is `0`.
    Boolean(bool),
    /// 16-bit. The default type of a small integer literal.
    Integer(i16),
    /// 32-bit.
    Long(i32),
    /// 32-bit float, from a `!` suffix.
    Single(f32),
    /// 64-bit float. The default type of any literal with a fraction or
    /// exponent.
    Double(f64),
    /// Fixed-point with 4 decimal places, stored scaled by 10_000 so that
    /// the decimal arithmetic it exists for stays exact.
    Currency(i64),
    /// A date serial. Numerically a `Double`; the difference is only in how
    /// it renders and what `TypeName` says.
    ///
    /// This is a VBA-side type, deliberately *not* mirrored by a
    /// `ResultData::Date` in the engine -- `core/date.rs` explains why the
    /// engine has no date value type at all. The conversion happens at the
    /// host boundary: a cell whose style carries a date `num_format` reads
    /// back through `.Value` as one of these, and through `.Value2` as a
    /// plain `Double`. Both halves measured (`fuzz/vba_host_probe.py`).
    Date(f64),
    /// A string.
    Str(String),
    /// An Excel error value, as `CVErr` builds one, `Application.VLookup`
    /// returns on failure, and a cell holding `=1/0` reads back as.
    ///
    /// The payload is the `CVErr` number (2007 for `#DIV/0!`, 2042 for
    /// `#N/A`, ...), which is what `CLng` on one gives back. Measured: it
    /// stringifies as `"Error 2042"` but is error 13 in arithmetic,
    /// concatenation and comparison alike.
    ErrValue(i32),
    /// An object reference, or `Nothing`.
    ///
    /// Reference semantics: `Set` assigns one, `Is` compares identity, and a
    /// plain `=` reads the object's default member instead. See
    /// [`ObjRef`](super::host::ObjRef) for why identity is a token rather
    /// than the range coordinates.
    Object(ObjRef),
    /// A 2-D `Variant` array, which in this scope only a multi-cell
    /// `Range.Value` produces.
    ///
    /// Behind an `Rc` because a `Variant` is cloned constantly and a range
    /// read can be large. Deliberately not general VBA arrays: `Dim x(10)`,
    /// `ReDim` and `Erase` are still out of scope and still report so.
    Array(Rc<VarArray>),
}

/// A 2-D `Variant` array, indexed from 1 as VBA's are.
///
/// Measured shape for a range read: `ws.Range("A1:A3").Value` has
/// `UBound(v, 1) = 3` and `UBound(v, 2) = 1`, i.e. `(row, column)` with rows
/// first, even for a single column.
#[derive(Debug, Clone, PartialEq)]
pub struct VarArray {
    /// Number of rows; `UBound(v, 1)`.
    pub rows: usize,
    /// Number of columns; `UBound(v, 2)`.
    pub cols: usize,
    /// The elements, row-major.
    pub values: Vec<Variant>,
}

impl VarArray {
    /// The element at a 1-based `(row, column)`, or error 9 if either index
    /// is outside the array.
    pub fn get(&self, row: usize, col: usize) -> VResult<Variant> {
        if row < 1 || col < 1 || row > self.rows || col > self.cols {
            return Err(VbaError::subscript());
        }
        Ok(self.values[(row - 1) * self.cols + (col - 1)].clone())
    }

    /// `UBound(v, dim)` for a 1-based `dim`.
    pub fn ubound(&self, dim: usize) -> VResult<usize> {
        match dim {
            1 => Ok(self.rows),
            2 => Ok(self.cols),
            _ => Err(VbaError::subscript()),
        }
    }
}

/// Whether an arithmetic operation may widen its result type on overflow.
///
/// The distinction is real and measured: `32767 + 1` written with two
/// literals is error 6, but the same addition with a variable on either side
/// promotes to `Long`. VBA compiles a **statically typed** expression with
/// its operands' own fixed widths and evaluates a `Variant` expression
/// through a path that widens.
///
/// It is static typing that decides this and not constness, which §28
/// measured in both directions -- see `interp::is_statically_typed`:
///
/// ```text
/// CInt(32767) + 1        error 6    typed, not constant
/// Sgn(1) + 32767         error 6    likewise
/// CInt(32767) + CInt(1)  error 6    likewise
/// (Empty + 32767) + 1    32768      constant, not typed -- `Empty` is Variant
/// a = 32767 : a + 1      32768      a variable, as before
/// Len("abcde") + 32763   32768      typed, but `Len` is Long: no overflow
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArithMode {
    /// Both operands are compile-time constants: overflow is an error.
    Constant,
    /// At least one operand is a variable: overflow widens the result.
    Promote,
}

/// Where a numeric result's type comes from, ordered by width so the wider
/// of two operands wins.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum NumClass {
    Integer,
    Long,
    Currency,
    Single,
    Double,
}

impl Variant {
    /// What `TypeName()` returns for this value.
    ///
    /// Observable from VBA, and therefore something the differential fuzzer
    /// compares -- an interpreter that computes the right number with the
    /// wrong subtype has a real bug.
    pub fn type_name(&self) -> &'static str {
        match self {
            Variant::Empty => "Empty",
            Variant::Null => "Null",
            Variant::Boolean(_) => "Boolean",
            Variant::Integer(_) => "Integer",
            Variant::Long(_) => "Long",
            Variant::Single(_) => "Single",
            Variant::Double(_) => "Double",
            Variant::Currency(_) => "Currency",
            Variant::Date(_) => "Date",
            Variant::Str(_) => "String",
            Variant::ErrValue(_) => "Error",
            Variant::Object(o) => o.type_name(),
            // Excel for Mac 16.112 reports this through the AppleScript
            // bridge as `V()` rather than `Variant()`, consistently and even
            // when the result is bracketed to prove nothing is truncating it.
            // `Variant()` is what the language documents and what every other
            // host reports, so that is what this returns; the discrepancy is
            // recorded in `fuzz/vba_host_probe.py` rather than matched, since
            // matching it would be encoding one bridge's quirk as a rule.
            Variant::Array(_) => "Variant()",
        }
    }

    /// The `CVErr` number if this is an error value.
    ///
    /// Separate from [`Variant::to_f64`] on purpose: `CLng(CVErr(2042))` is
    /// `2042`, but `CVErr(2042) + 1` is error 13. The explicit conversions
    /// reach for this; arithmetic must not.
    pub fn error_number(&self) -> Option<i32> {
        match self {
            Variant::ErrValue(n) => Some(*n),
            _ => None,
        }
    }

    /// The object this holds, if it is one.
    pub fn as_object(&self) -> Option<&ObjRef> {
        match self {
            Variant::Object(o) => Some(o),
            _ => None,
        }
    }

    /// Whether this is `Null`, which most operations propagate.
    pub fn is_null(&self) -> bool {
        matches!(self, Variant::Null)
    }

    /// Whether this is `Empty`.
    pub fn is_empty(&self) -> bool {
        matches!(self, Variant::Empty)
    }

    fn num_class(&self) -> Option<NumClass> {
        Some(match self {
            // Probe case 48: `Empty + 1` is the Integer 1, so Empty enters
            // arithmetic as an Integer zero.
            Variant::Empty | Variant::Boolean(_) | Variant::Integer(_) => NumClass::Integer,
            Variant::Long(_) => NumClass::Long,
            Variant::Currency(_) => NumClass::Currency,
            Variant::Single(_) => NumClass::Single,
            // Probe case 35: a String operand coerces to Double, not to
            // whatever the other operand is.
            Variant::Double(_) | Variant::Date(_) | Variant::Str(_) => NumClass::Double,
            Variant::Null => return None,
            // These three have no numeric class at all. They report `Double`
            // rather than `None` because `None` means `Null` to every caller
            // here, and the real refusal happens a line later in `to_f64`,
            // which is error 13 for all three -- measured for `ErrValue`
            // (`CVErr(2042) + 1`), and the only defensible answer for an
            // object or an array.
            Variant::ErrValue(_) | Variant::Object(_) | Variant::Array(_) => NumClass::Double,
        })
    }

    /// This value as an `f64`, for arithmetic.
    ///
    /// `Null` is rejected rather than defaulted: an operation that reaches
    /// here with a `Null` has failed to propagate it, and silently treating
    /// it as zero would be worse than an error.
    pub fn to_f64(&self) -> VResult<f64> {
        Ok(match self {
            Variant::Empty => 0.0,
            // True is -1. Probe case 41: `True + 1` is 0.
            Variant::Boolean(b) => {
                if *b {
                    -1.0
                } else {
                    0.0
                }
            }
            Variant::Integer(v) => *v as f64,
            Variant::Long(v) => *v as f64,
            Variant::Single(v) => *v as f64,
            Variant::Double(v) | Variant::Date(v) => *v,
            Variant::Currency(v) => *v as f64 / 10_000.0,
            Variant::Str(s) => parse_vba_number(s)?,
            Variant::Null => return Err(VbaError::invalid_null()),
            // Measured: `v = CVErr(2042)` makes `v + 1`, `v & ""` and `v = 1`
            // all error 13. `CLng(v)` still gives 2042 -- that path goes
            // through `error_number`, not here. An object reaching this point
            // has not had its default member read, and an array has none.
            Variant::ErrValue(_) | Variant::Object(_) | Variant::Array(_) => {
                return Err(VbaError::type_mismatch());
            }
        })
    }

    /// This value as a string, as `CStr` and `&` produce it.
    pub fn to_vba_string(&self) -> VResult<String> {
        Ok(match self {
            // Probe case 14: `CStr(Empty)` is the empty string.
            Variant::Empty => String::new(),
            Variant::Null => return Err(VbaError::invalid_null()),
            Variant::Boolean(b) => if *b { "True" } else { "False" }.to_string(),
            Variant::Integer(v) => v.to_string(),
            Variant::Long(v) => v.to_string(),
            Variant::Single(v) => format_number(*v as f64),
            Variant::Double(v) => format_number(*v),
            Variant::Currency(v) => format_currency(*v),
            Variant::Date(v) => format_vba_date(*v),
            Variant::Str(s) => s.clone(),
            // Measured: `CStr(CVErr(2042))` is the string "Error 2042", even
            // though `CVErr(2042) & ""` is error 13. `CStr` is the one
            // conversion that renders an error value rather than refusing it,
            // so this is reachable only from there -- `concat` rejects an
            // `ErrValue` before it gets here.
            Variant::ErrValue(n) => format!("Error {n}"),
            Variant::Object(_) | Variant::Array(_) => return Err(VbaError::type_mismatch()),
        })
    }

    /// This value as a `Boolean`, as `CBool` and the logical operators read
    /// it.
    ///
    /// Any non-zero number is true, which is why `If 5 Then` runs. `Null`
    /// is error 94 here -- measured against real Excel (Windows):
    /// `CBool(Null)` and `Not Null` both raise 94. A *statement condition*
    /// (`If`/`Do While`/`Do Until`) is a different coercion that treats
    /// `Null` as `False` instead -- see `to_bool_condition`, which is what
    /// those statements actually use.
    pub fn to_bool(&self) -> VResult<bool> {
        match self {
            Variant::Boolean(b) => Ok(*b),
            Variant::Null => Err(VbaError::invalid_null()),
            // `CBool("True")` is True even though `CDbl("True")` is error 13.
            Variant::Str(s) if bool_word(s).is_some() => Ok(bool_word(s).unwrap_or(false)),
            other => Ok(other.to_f64()? != 0.0),
        }
    }

    /// This value as a `Boolean`, as an `If`/`Do While`/`Do Until` statement
    /// condition reads it -- unlike [`Self::to_bool`], a `Null` condition is
    /// `False` rather than error 94. Measured against real Excel (Windows):
    /// `If Null Then` takes the `Else` branch, `Do While Null` never loops,
    /// and `Do Until Null` loops until an explicit exit (i.e. the condition
    /// reads as `False`, never `True`) -- while `CBool(Null)` and `Not Null`
    /// still raise 94 in the same session. Two different coercions behind
    /// what looks like one "read as boolean" idea, confirmed separately
    /// rather than assumed to be the same rule (fuzz/fuzz_vba.py, whose
    /// win32com driver made this measurable on Windows for the first time).
    pub fn to_bool_condition(&self) -> VResult<bool> {
        match self {
            Variant::Null => Ok(false),
            other => other.to_bool(),
        }
    }

    /// Builds the narrowest [`Variant`] that a numeric literal of this text
    /// should have.
    ///
    /// Probe cases 1/2/15--20: `1` is `Integer`, `32768` is `Long`,
    /// `2147483648` is `Double`, and `1E5` is `Double` even though `100000`
    /// alone is `Long` -- exponent notation forces floating point regardless
    /// of the value.
    pub fn from_literal(value: f64, has_fraction_or_exponent: bool) -> Variant {
        if has_fraction_or_exponent {
            return Variant::Double(value);
        }
        if value >= i16::MIN as f64 && value <= i16::MAX as f64 {
            Variant::Integer(value as i16)
        } else if value >= i32::MIN as f64 && value <= i32::MAX as f64 {
            Variant::Long(value as i32)
        } else {
            Variant::Double(value)
        }
    }

    /// Packs an `f64` into the given numeric class.
    ///
    /// In [`ArithMode::Promote`] a value that does not fit widens to the next
    /// class up rather than erroring, which is what runtime Variant
    /// arithmetic does. In [`ArithMode::Constant`] -- both operands literal --
    /// it errors instead, which is what `32767 + 1` does.
    fn pack_mode(value: f64, class: NumClass, mode: ArithMode) -> VResult<Variant> {
        if mode == ArithMode::Constant {
            return Self::pack(value, class);
        }
        let mut class = class;
        loop {
            match Self::pack(value, class) {
                Ok(v) => return Ok(v),
                Err(e) => {
                    class = match class {
                        NumClass::Integer => NumClass::Long,
                        NumClass::Long | NumClass::Currency | NumClass::Single => NumClass::Double,
                        // Nothing is wider than Double, so a value that does
                        // not fit there really is an overflow.
                        NumClass::Double => return Err(e),
                    };
                }
            }
        }
    }

    fn pack(value: f64, class: NumClass) -> VResult<Variant> {
        Ok(match class {
            NumClass::Integer => {
                if value < i16::MIN as f64 || value > i16::MAX as f64 {
                    return Err(VbaError::overflow());
                }
                Variant::Integer(value as i16)
            }
            NumClass::Long => {
                if value < i32::MIN as f64 || value > i32::MAX as f64 {
                    return Err(VbaError::overflow());
                }
                Variant::Long(value as i32)
            }
            NumClass::Currency => {
                let scaled = bankers_round(value * 10_000.0);
                if scaled < i64::MIN as f64 || scaled > i64::MAX as f64 {
                    return Err(VbaError::overflow());
                }
                Variant::Currency(scaled as i64)
            }
            NumClass::Single => {
                let as_single = value as f32;
                if !as_single.is_finite() && value.is_finite() {
                    return Err(VbaError::overflow());
                }
                Variant::Single(as_single)
            }
            // A Double may hold an infinity: `255 ^ 255` is `INF`, not an
            // error, and negating it gives `-INF`. Only the arithmetic
            // operators below refuse to produce or consume one.
            NumClass::Double => Variant::Double(value),
        })
    }

    /// The result class of an arithmetic operation on two operands.
    ///
    /// Normally the wider of the two, but `Single` combined with `Long` is
    /// `Double` rather than `Single` -- a `Single` cannot hold every `Long`,
    /// so VBA widens past both. Measured: `2! + 1` is a `Single` (the other
    /// side is an `Integer`) while `2! * 1&` is a `Double`.
    fn arith_type(lhs: &Variant, rhs: &Variant) -> VResult<NumClass> {
        let l = lhs.num_class().ok_or_else(VbaError::invalid_null)?;
        let r = rhs.num_class().ok_or_else(VbaError::invalid_null)?;
        let pair = (l.min(r), l.max(r));
        if pair == (NumClass::Long, NumClass::Single) {
            return Ok(NumClass::Double);
        }
        Ok(l.max(r))
    }
}

/// `"True"` / `"False"` as a boolean, case-insensitively and ignoring
/// surrounding space.
///
/// VBA accepts these words on the *integer* conversion path only. `"True"
/// Xor 1` is `-2` and `CBool("True")` is `True`, but `"True" + 1` and
/// `CDbl("True")` are both error 13 -- the floating-point path has never
/// heard of them. That asymmetry is why this is a separate function rather
/// than a case inside [`parse_vba_number`].
pub fn bool_word(s: &str) -> Option<bool> {
    let t = s.trim();
    if t.eq_ignore_ascii_case("true") {
        Some(true)
    } else if t.eq_ignore_ascii_case("false") {
        Some(false)
    } else {
        None
    }
}

/// A logical/bitwise operand, with a `"True"`/`"False"` string folded to a
/// `Boolean` so the ordinary rules take over from there.
///
/// Folding to `Boolean` rather than to a number is what makes
/// `Not "True"` a `Boolean` while `"True" Xor 1` is an `Integer`: the first
/// stays inside [`not`]'s boolean branch, the second falls into the bitwise
/// one because only one side is a `Boolean`.
fn logical_operand(v: &Variant) -> Variant {
    match v {
        Variant::Str(s) => match bool_word(s) {
            Some(b) => Variant::Boolean(b),
            None => v.clone(),
        },
        _ => v.clone(),
    }
}

/// The pair of operands for a logical operator, with the `"True"`/`"False"`
/// fold applied only where Excel applies it.
///
/// Against a `Boolean` partner the fold is suppressed exactly when **both**
/// operands' types are known statically -- a literal, a constant expression,
/// or a call whose declared return type says so:
///
/// | Expression | Excel | Why |
/// | --- | --- | --- |
/// | `True Eqv "True"` | error 13 | Boolean literal, String literal |
/// | `True Eqv CStr(True)` | error 13 | `CStr` is declared `As String` |
/// | `a = 3.75 : IsNumeric(a) Eqv CStr(True)` | error 13 | both declared |
/// | `LCase("TRUE") Eqv True` | `True` | `LCase` returns a *Variant* |
/// | `a = True : a Eqv "True"` | `True` | `a` is a Variant |
/// | `a = "false" : a Eqv False` | `True` | same, other way round |
///
/// The `CStr`/`LCase` pair is the one that pins it down, and it is not
/// arbitrary: `CStr` returns `String`, while `LCase` (like `UCase`, `Left`
/// and friends, whose `$`-suffixed forms are the String-typed ones) returns
/// `Variant`. A non-Boolean partner always folds -- `"True" Xor 1` is `-2`
/// between two literals.
fn logical_pair(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> (Variant, Variant) {
    let both_static = kinds.0 != Operand::Runtime && kinds.1 != Operand::Runtime;
    let l_bool = matches!(lhs, Variant::Boolean(_));
    let r_bool = matches!(rhs, Variant::Boolean(_));
    (
        if r_bool && both_static {
            lhs.clone()
        } else {
            logical_operand(lhs)
        },
        if l_bool && both_static {
            rhs.clone()
        } else {
            logical_operand(rhs)
        },
    )
}

/// Round-half-to-even, which is what every VBA numeric conversion uses.
///
/// Probe cases 55--59: `CLng(0.5)` is `0`, `CLng(1.5)` is `2`, `CLng(2.5)`
/// is `2`, `CLng(-0.5)` is `0`, `CLng(-1.5)` is `-2`. Rust's `f64::round`
/// rounds half away from zero and would give 1, 2, 3, -1, -2.
pub fn bankers_round(v: f64) -> f64 {
    let floor = v.floor();
    let diff = v - floor;
    if diff > 0.5 {
        floor + 1.0
    } else if diff < 0.5 {
        floor
    } else {
        // Exactly halfway: go to the even neighbour.
        if (floor / 2.0).fract() == 0.0 {
            floor
        } else {
            floor + 1.0
        }
    }
}

/// Parses a string the way VBA's implicit string-to-number coercion does.
///
/// Probe case 40: leading and trailing whitespace is ignored (`"  3  " + 1`
/// is `4`). Probe case 37: anything else that is not a number is error 13,
/// not zero.
pub fn parse_vba_number(s: &str) -> VResult<f64> {
    let t = s.trim();
    if t.is_empty() {
        // Measured: `"" = 0`, `"" < 0` and `Not ""` are all error 13. An
        // empty string is NOT a zero -- that is `Empty`'s job.
        return Err(VbaError::type_mismatch());
    }
    // `&H`/`&O` literals are accepted in string form too.
    if let Some(hex) = t.strip_prefix("&H").or_else(|| t.strip_prefix("&h")) {
        return i64::from_str_radix(hex, 16)
            .map(|v| v as f64)
            .map_err(|_| VbaError::type_mismatch());
    }
    if let Some(oct) = t.strip_prefix("&O").or_else(|| t.strip_prefix("&o")) {
        return i64::from_str_radix(oct, 8)
            .map(|v| v as f64)
            .map_err(|_| VbaError::type_mismatch());
    }
    // VBA's `D` exponent marker is equivalent to `E`.
    let normalised = t.replace(['d', 'D'], "e");
    // A *trailing* sign negates (or confirms) the number, so `CDbl("1-")` is
    // -1 and `CDbl("1+")` is 1 -- as is `"1 -"`, the space being trimmed, and
    // `"1E2-"`, which is -100. It is a suffix, not a second sign: `"-1-"` and
    // `"1--"` are both error 13. Measured; this is what makes `CBool("1-")`
    // True where visi used to raise 13 and take the other `If` branch.
    let negated_by_suffix = normalised.ends_with('-');
    let body = match normalised.strip_suffix(['-', '+']) {
        Some(rest) => {
            let rest = rest.trim_end();
            // The suffix replaces a sign rather than adding to one.
            if rest.starts_with(['-', '+']) {
                return Err(VbaError::type_mismatch());
            }
            rest
        }
        None => normalised.as_str(),
    };
    let value = body
        .parse::<f64>()
        .map(|v| if negated_by_suffix { -v } else { v })
        .map_err(|_| VbaError::type_mismatch())?;
    // A string whose value is outside Double range fails to *convert* --
    // error 6, not 13, and not a quiet infinity. `a = "1E+2923" : a ^ 255`
    // is error 6 for this reason, while `a = "255" : a ^ 255` is INF: the
    // power overflows happily, the conversion does not.
    if !value.is_finite() {
        return Err(VbaError::overflow());
    }
    Ok(value)
}

/// The longest leading run of `s` that parses as a number, as `Val` takes it.
///
/// Comparison against a numeric *constant* coerces the string this way rather
/// than demanding the whole string parse, which is what separates
/// `(Not 2!) <= ("1.5" & False)` -- `"1.5False"` has the numeric prefix
/// `1.5`, so the comparison succeeds -- from `(-True) <> (True & &HFF)`,
/// where `"True255"` has none and the comparison is error 13.
pub fn numeric_prefix(s: &str) -> Option<f64> {
    let t: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    let mut best = None;
    for (i, _) in t.char_indices() {
        if let Ok(v) = t[..=i].parse::<f64>()
            && v.is_finite()
        {
            best = Some(v);
        }
    }
    best
}

/// Renders a number the way VBA's `CStr` does.
///
/// Not the same as Rust's `{}`: VBA prints up to 15 significant digits and
/// drops a trailing `.0`, and writes exponents as `1E+20`.
pub fn format_number(v: f64) -> String {
    if v == 0.0 {
        // `-0.0 == 0.0`, so the sign has to come from the bit pattern.
        // Excel's CStr does render negative zero as "-0".
        return if v.is_sign_negative() { "-0" } else { "0" }.to_string();
    }
    if v.is_nan() {
        return "NaN".to_string();
    }
    if v.is_infinite() {
        // Measured: `CStr(255 ^ 255)` is "INF", not VB6's "1.#INF".
        return if v > 0.0 { "INF" } else { "-INF" }.to_string();
    }

    // Everything is derived from a 15-significant-digit rendering, because
    // that is exactly what Excel shows. Getting this from `log10().floor()`
    // and a decimal count was wrong in both directions -- 16 digits in
    // exponent form and 14 in fixed form -- and the fuzzer caught it on
    // every seed. Asking Rust for the exponent instead removes the
    // arithmetic that was getting it wrong.
    let sci = format!("{v:.14e}");
    let (mantissa, exp_text) = sci.split_once('e').unwrap_or((sci.as_str(), "0"));
    let exp: i32 = exp_text.parse().unwrap_or(0);

    // Excel stays in fixed notation over this range and switches outside it.
    if (-4..16).contains(&exp) {
        let decimals = (14 - exp).max(0) as usize;
        let mut r = format!("{v:.decimals$}");
        if r.contains('.') {
            r = r.trim_end_matches('0').trim_end_matches('.').to_string();
        }
        return if r == "-0" { "-0".to_string() } else { r };
    }

    let mantissa = if mantissa.contains('.') {
        mantissa.trim_end_matches('0').trim_end_matches('.')
    } else {
        mantissa
    };
    let sign = if exp < 0 { "-" } else { "+" };
    format!("{mantissa}E{sign}{:02}", exp.abs())
}

fn format_currency(scaled: i64) -> String {
    let whole = scaled / 10_000;
    let frac = (scaled % 10_000).abs();
    if frac == 0 {
        whole.to_string()
    } else {
        let frac_str = format!("{frac:04}");
        let frac_str = frac_str.trim_end_matches('0');
        let sign = if scaled < 0 && whole == 0 { "-" } else { "" };
        format!("{sign}{whole}.{frac_str}")
    }
}

// ---------------------------------------------------------------------------
// Operators
// ---------------------------------------------------------------------------

/// `+`, which is arithmetic *or* concatenation depending on the operands.
///
/// Probe cases 35--37: `"1" + 1` is the `Double` 2, `"1" + "2"` is the
/// `String` "12", and `"abc" + 1` is error 13. Only when *both* sides are
/// strings does `+` concatenate.
pub fn add(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        // `+` alone does *not* coerce the other operand first: `Null + "Z"`
        // is Null, while `Null - "Z"` is error 13. The likely reason is that
        // `+` is overloaded -- it cannot know whether it is addition or
        // concatenation without inspecting both sides -- so it short-circuits
        // before deciding. Measured in both directions.
        return Ok(Variant::Null);
    }
    if let (Variant::Str(a), Variant::Str(b)) = (lhs, rhs) {
        return Ok(Variant::Str(format!("{a}{b}")));
    }
    // `Empty` against a String concatenates instead of forcing the string
    // into a number, so `Empty + "a"` is the String "a" rather than error 13,
    // and `Empty + "1"` is the String "1" rather than the Double 1. Empty
    // takes the other operand's type here, the same way it does against a
    // number -- and only for `+`: `Empty - "a"` and `Empty * "a"` are both
    // still error 13. Measured with `fuzz/vba_expr_probe.py`.
    if let (Variant::Empty, Variant::Str(s)) | (Variant::Str(s), Variant::Empty) = (lhs, rhs) {
        return Ok(Variant::Str(s.clone()));
    }
    // `Empty + Empty` falls through to `arith` and lands on the Integer 0,
    // which is what `TypeName(Empty + Empty)` reports in Excel. It used to
    // short-circuit to Empty here; that came from reading the result back
    // through the fuzz harness, which cannot see the difference -- Empty and
    // the Integer 0 both render as "0" once assigned onward.
    keep_date(lhs, rhs, arith(lhs, rhs, mode, |a, b| a + b)?)
}

/// `-`.
pub fn sub(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    keep_date(lhs, rhs, arith(lhs, rhs, mode, |a, b| a - b)?)
}

/// A date plus or minus a number is still a date; a date minus a date is a
/// count of days and is not.
///
/// Measured: `TypeName(#6/22/2026# + 1)` is `Date` and `CStr` of it is
/// `6/23/26`, while `#6/22/2026# - #6/21/2026#` is `1`. The rule keys off
/// *exactly one* operand being a `Date`, which is the same shape as the
/// engine's own `Sheet::inherited_date_format` -- and, as there, it applies
/// to `+` and `-` only. `*` and `/` are not measured and do not preserve the
/// subtype here; a date multiplied by anything is not a date in any reading.
fn keep_date(lhs: &Variant, rhs: &Variant, result: Variant) -> VResult<Variant> {
    let one_date = matches!(lhs, Variant::Date(_)) != matches!(rhs, Variant::Date(_));
    if !one_date {
        return Ok(result);
    }
    Ok(match result {
        Variant::Integer(_)
        | Variant::Long(_)
        | Variant::Single(_)
        | Variant::Double(_)
        | Variant::Currency(_) => Variant::Date(result.to_f64()?),
        other => other,
    })
}

/// A `Date` as `CStr` renders it: the system short date, plus a time when the
/// serial carries one, and the time alone when it carries no date.
///
/// Measured against Excel for Mac 16.112 on a machine set to en-US:
/// `CStr(#6/22/2026#)` is `6/22/26`, `CStr(#6/22/2026 12:00:00 PM#)` is
/// `6/22/26 12:00:00 PM`, and `CStr(CDate(0.5))` is `12:00:00 PM`. The
/// two-digit year and the `m/d/yy` order come from the *system* short-date
/// setting rather than from VBA, so a machine configured differently will
/// disagree -- that is a property of the language, not a bug here, and it is
/// why the fuzz harness compares dates on a machine it also measured on.
pub fn format_vba_date(serial: f64) -> String {
    let days = serial.floor();
    let frac = serial - days;
    // Excel's serial 0 is the (fictional) 1900-01-00, so a pure time has no
    // date part to render.
    let date_part = if days == 0.0 {
        None
    } else {
        let d = crate::core::date::excel_serial_to_date(days);
        Some(format!(
            "{}/{}/{:02}",
            d.month,
            d.day,
            d.year.rem_euclid(100)
        ))
    };
    // Rounded to the nearest second before splitting, so that a serial a
    // hair under a whole day does not render as `23:59:60`.
    let total_seconds = (frac * 86_400.0).round() as i64;
    let time_part = if total_seconds == 0 && date_part.is_some() {
        None
    } else {
        let (h24, m, s) = (
            total_seconds / 3600,
            (total_seconds / 60) % 60,
            total_seconds % 60,
        );
        let (h12, ampm) = match h24 {
            0 => (12, "AM"),
            1..=11 => (h24, "AM"),
            12 => (12, "PM"),
            _ => (h24 - 12, "PM"),
        };
        Some(format!("{h12}:{m:02}:{s:02} {ampm}"))
    };
    match (date_part, time_part) {
        (Some(d), Some(t)) => format!("{d} {t}"),
        (Some(d), None) => d,
        (None, Some(t)) => t,
        // Serial 0 exactly: VBA renders the epoch's own date.
        (None, None) => "12:00:00 AM".to_string(),
    }
}

/// `*`.
pub fn mul(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    arith(lhs, rhs, mode, |a, b| a * b)
}

/// `/`, which is always floating point.
///
/// Probe case 25: `4 / 2` is the `Double` 2, not an `Integer`.
pub fn div(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        // Coerce the non-Null side first, as the other operators do:
        // `"abc" / Null` is error 13, not Null.
        if !lhs.is_null() {
            lhs.to_f64()?;
        }
        if !rhs.is_null() {
            rhs.to_f64()?;
        }
        return Ok(Variant::Null);
    }
    // Both operands are coerced *before* the divisor is tested, so a type
    // mismatch beats a division by zero: `"xxxx" / 0` is error 13, not 11.
    // Testing the divisor first masked the real error.
    let a = lhs.to_f64()?;
    let b = rhs.to_f64()?;
    if b == 0.0 {
        // `0 / 0` is Overflow, not Division by zero -- measured, and specific
        // to floating-point `/`: `0 \ 0` and `0 Mod 0` are both error 11.
        return Err(if a == 0.0 {
            VbaError::overflow()
        } else {
            VbaError::div_by_zero()
        });
    }
    let r = a / b;
    // `/` refuses infinity exactly as `arith` does for `+`, `-` and `*`:
    // `1E308 / 1E-308` is error 6, and so is `b / 2` when `b` is already
    // infinite. It is only `^` that hands back an infinity. Measured both
    // ways round -- this was the last operator still returning INF, which is
    // how `(-7 ^ a) / (Not c)` came out INF here and error 6 in Excel.
    if !r.is_finite() || !a.is_finite() || !b.is_finite() {
        return Err(VbaError::overflow());
    }
    Ok(Variant::Double(r))
}

/// `\` -- integer division.
///
/// Probe cases 26--28: operands are rounded to integers *first*, so
/// `7.6 \ 2` is `4` rather than `3`, and a non-integral operand widens the
/// result to `Long`.
pub fn int_div(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    let (a, b, class) = int_operands(lhs, rhs)?;
    let Some((a, b, class)) = zip3(a, b, class) else {
        return Ok(Variant::Null);
    };
    if b == 0 {
        return Err(VbaError::div_by_zero());
    }
    // Truncating division: probe case 27 gives -3 for `-7 \ 2`.
    Variant::pack((a / b) as f64, class)
}

/// `Mod`, with the same operand rounding and widening as `\`.
///
/// Probe case 31: `7.6 Mod 2` is `0`, because 7.6 rounds to 8 first.
pub fn modulo(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    let (a, b, class) = int_operands(lhs, rhs)?;
    let Some((a, b, class)) = zip3(a, b, class) else {
        return Ok(Variant::Null);
    };
    if b == 0 {
        return Err(VbaError::div_by_zero());
    }
    // Sign follows the dividend: probe case 30 gives -1 for `-7 Mod 2`.
    Variant::pack((a % b) as f64, class)
}

fn zip3(a: Option<i64>, b: Option<i64>, class: NumClass) -> Option<(i64, i64, NumClass)> {
    Some((a?, b?, class))
}

/// Rounds both operands to integers and decides the result class, shared by
/// `\` and `Mod`.
fn int_operands(lhs: &Variant, rhs: &Variant) -> VResult<(Option<i64>, Option<i64>, NumClass)> {
    // Each operand is coerced, rounded and range-checked in turn, left
    // before right, because which error surfaces depends on the order:
    // `"32768100000" Mod "Double"` is error 6 (the left overflows a Long)
    // while `"Double" Mod "32768100000"` is error 13. Coercing both and then
    // checking both reported the wrong one.
    //
    // A Null operand does not short-circuit past its partner either:
    // `Null Mod "Z"` is error 13, not Null.
    fn one(v: &Variant) -> VResult<Option<i64>> {
        if v.is_null() {
            return Ok(None);
        }
        let r = bankers_round(v.to_f64()?);
        if !r.is_finite() || r < i32::MIN as f64 || r > i32::MAX as f64 {
            return Err(VbaError::overflow());
        }
        Ok(Some(r as i64))
    }

    // The `"True"`/`"False"` fold has to happen *before* coercion, or the
    // words would fail `to_f64` on the way in: `"True" \\ 1` is -1.
    // `\` and `Mod` claim both operands are runtime, so the fold always
    // happens for them: their constant Boolean-against-String case is handled
    // ahead of this by `interp::constant_bool_int_op`, and the suppression
    // `Eqv` and friends need has not been measured for them.
    let (l, r) = logical_pair(lhs, rhs, (Operand::Runtime, Operand::Runtime));

    let a = one(&l)?;
    let b = one(&r)?;
    if a.is_none() || b.is_none() {
        return Ok((None, None, NumClass::Long));
    }

    // A non-integral operand forces Long; two small integers stay Integer.
    let class = match Variant::arith_type(&l, &r)? {
        NumClass::Integer => NumClass::Integer,
        _ => NumClass::Long,
    };
    Ok((a, b, class))
}

/// `^`, which is always `Double`.
///
/// Probe case 32: `2 ^ 2` is the `Double` 4.
pub fn pow(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        if !lhs.is_null() {
            lhs.to_f64()?;
        }
        if !rhs.is_null() {
            rhs.to_f64()?;
        }
        return Ok(Variant::Null);
    }
    let base = lhs.to_f64()?;
    let exp = rhs.to_f64()?;
    // A negative base with a fractional exponent has no real result, and VBA
    // raises error 5 rather than returning NaN. Found by fuzz/fuzz_vba.py via
    // `(-1) ^ 1.5`, which this used to return as a quiet NaN.
    if base < 0.0 && exp.fract() != 0.0 {
        return Err(VbaError::invalid_call());
    }
    // `0 ^ -1` has no value either, and is the same error. `0 ^ 0` is 1 and
    // `0 ^ 2` is 0, so it is specifically a negative exponent over a zero
    // base.
    if base == 0.0 && exp < 0.0 {
        return Err(VbaError::invalid_call());
    }
    let r = base.powf(exp);
    // The same constant-vs-runtime split the other operators have:
    // `3.75 ^ 32767` written with literals is error 6, while the same
    // exponentiation with a variable base yields `INF`.
    if mode == ArithMode::Constant && !r.is_finite() && base.is_finite() && exp.is_finite() {
        return Err(VbaError::overflow());
    }
    Ok(Variant::Double(r))
}

/// `&` -- concatenation, which skips `Null` operands rather than
/// propagating them.
///
/// Probe case 51: `Null & "a"` is `"a"`. Probe case 38: `1 & 2` is `"12"`.
pub fn concat(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    if lhs.is_null() && rhs.is_null() {
        return Ok(Variant::Null);
    }
    // `CStr(CVErr(2042))` renders as "Error 2042", but `CVErr(2042) & ""` is
    // error 13 -- `&` refuses an error value rather than stringifying it.
    // Both measured, and the pair is why `to_vba_string` cannot be the only
    // gate here.
    if matches!(lhs, Variant::ErrValue(_)) || matches!(rhs, Variant::ErrValue(_)) {
        return Err(VbaError::type_mismatch());
    }
    let a = if lhs.is_null() {
        String::new()
    } else {
        lhs.to_vba_string()?
    };
    let b = if rhs.is_null() {
        String::new()
    } else {
        rhs.to_vba_string()?
    };
    Ok(Variant::Str(format!("{a}{b}")))
}

fn arith(
    lhs: &Variant,
    rhs: &Variant,
    mode: ArithMode,
    f: impl Fn(f64, f64) -> f64,
) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        // Null propagates, but only *after* the other operand has been
        // coerced: `"Z" - Null` is error 13, not Null. `+` is the one
        // exception -- see `add`.
        if !lhs.is_null() {
            lhs.to_f64()?;
        }
        if !rhs.is_null() {
            rhs.to_f64()?;
        }
        return Ok(Variant::Null);
    }
    let class = Variant::arith_type(lhs, rhs)?;
    let (a, b) = (lhs.to_f64()?, rhs.to_f64()?);
    let r = f(a, b);
    // Measured: `1E300 * 1E300` is error 6, and so is `b + 1` when `b` is
    // already infinite -- but `1E300 + 1E300` is fine, and `^` produces
    // infinities happily. So +, - and * refuse infinity on either side.
    if !r.is_finite() || !a.is_finite() || !b.is_finite() {
        return Err(VbaError::overflow());
    }
    Variant::pack_mode(r, class, mode)
}

/// Unary `-`.
pub fn neg(v: &Variant, mode: ArithMode) -> VResult<Variant> {
    if v.is_null() {
        return Ok(Variant::Null);
    }
    // Negating the `Long` minimum between constants **wraps to itself**
    // rather than overflowing or widening: `-(Not 2147483647)` is the `Long`
    // -2147483648, which is arithmetically wrong and is what Excel does.
    // Plain two's complement, and narrow -- the `Integer` minimum does *not*
    // do it (`-(Not 32767)` is error 6 on both sides), and at run time the
    // whole thing widens instead (`a = 2147483647 : -(Not a)` is the Double
    // 2147483648). All three measured.
    if mode == ArithMode::Constant
        && let Variant::Long(n) = v
        && *n == i32::MIN
    {
        return Ok(Variant::Long(i32::MIN));
    }
    // Boolean negation widens to Integer: `-True` is 1.
    let class = v.num_class().ok_or_else(VbaError::invalid_null)?;
    // Promotes on overflow at runtime, like the binary operators:
    // `a = 2147483647 : -(Not a)` is the Double 2147483648, not an overflow.
    Variant::pack_mode(-v.to_f64()?, class, mode)
}

/// Unary `+`, which still coerces to a number.
pub fn pos(v: &Variant, mode: ArithMode) -> VResult<Variant> {
    if v.is_null() {
        return Ok(Variant::Null);
    }
    let class = v.num_class().ok_or_else(VbaError::invalid_null)?;
    Variant::pack_mode(v.to_f64()?, class, mode)
}

/// `Not`, which is bitwise on numbers and logical on `Boolean`s.
///
/// Probe case 46: `Not 5` is `-6`, the bitwise complement.
///
/// Propagates a `Null` operand as `Null` -- this is the primitive `imp`
/// builds on (`Imp` is defined as `Not a Or b`, so `Null Imp True` needs
/// `Not Null` to come back `Null` here, then `three_valued`'s Or logic
/// picks the truthy `True`). The user-facing `Not` *operator* is a
/// different rule: measured directly against real Excel (Windows),
/// `Not Null` typed directly by a caller raises error 94, the same 94
/// `CBool(Null)` does -- see `UnOp::Not`'s own dispatch in `interp.rs`,
/// which checks for `Null` before ever calling this function, rather than
/// this shared primitive raising and breaking `imp`'s use of it.
pub fn not(v: &Variant) -> VResult<Variant> {
    let v = &logical_operand(v);
    match v {
        Variant::Null => Ok(Variant::Null),
        Variant::Boolean(b) => Ok(Variant::Boolean(!b)),
        other => {
            let n = bankers_round(other.to_f64()?);
            let class = match other.num_class().ok_or_else(VbaError::invalid_null)? {
                NumClass::Integer => NumClass::Integer,
                _ => NumClass::Long,
            };
            Variant::pack(!(n as i64) as f64, class)
        }
    }
}

/// The bitwise/logical binary operators.
///
/// Probe cases 43/45: `True And False` is the `Boolean` `False`, but
/// `5 And 3` is the `Integer` `1` -- the operation is bitwise unless both
/// operands are already `Boolean`.
pub fn logical(
    lhs: &Variant,
    rhs: &Variant,
    kinds: (Operand, Operand),
    f: impl Fn(i64, i64) -> i64,
) -> VResult<Variant> {
    // `"True" Xor 1` is -2: the integer path accepts the words.
    let (l, r) = logical_pair(lhs, rhs, kinds);
    let (lhs, rhs) = (&l, &r);
    if let (Variant::Boolean(a), Variant::Boolean(b)) = (lhs, rhs) {
        let r = f(if *a { -1 } else { 0 }, if *b { -1 } else { 0 });
        return Ok(Variant::Boolean(r != 0));
    }
    // Each operand is coerced and range-checked in turn, left before right,
    // and a Null does not short-circuit past its partner: `Null And "Z"` is
    // error 13. Same rule as `int_operands`, for the same measured reason.
    let one = |v: &Variant| -> VResult<Option<f64>> {
        if v.is_null() {
            return Ok(None);
        }
        let r = bankers_round(v.to_f64()?);
        // The operands have to fit a Long, as `\\` and `Mod` require:
        // `True Or "2147483648"` is error 6 even though the answer would fit.
        if !r.is_finite() || r < i32::MIN as f64 || r > i32::MAX as f64 {
            return Err(VbaError::overflow());
        }
        Ok(Some(r))
    };
    let (a, b) = (one(lhs)?, one(rhs)?);
    let (Some(a), Some(b)) = (a, b) else {
        return Ok(Variant::Null);
    };
    let class = match Variant::arith_type(lhs, rhs)? {
        NumClass::Integer => NumClass::Integer,
        _ => NumClass::Long,
    };
    Variant::pack(f(a as i64, b as i64) as f64, class)
}

/// `And`, which is three-valued: a `Null` operand does not always poison the
/// result.
///
/// Measured: `False And Null` is `False`, and `0 And Null` is the `Integer`
/// `0` -- a falsy operand *determines* the answer, so the result is that
/// operand, returned unchanged (type included). Only when the known operand
/// is truthy is the answer genuinely unknown: `5 And Null` and `-1 And Null`
/// are both `Null`.
pub fn and(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> VResult<Variant> {
    if let Some(v) = three_valued(lhs, rhs, false)? {
        return Ok(v);
    }
    logical(lhs, rhs, kinds, |x, y| x & y)
}

/// `Or`, three-valued in the mirrored way.
///
/// Measured: `True Or Null` is `True`, `5 Or Null` is the `Integer` `5`, and
/// `0 Or Null` is `Null`. A *truthy* operand determines the answer here.
pub fn or(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> VResult<Variant> {
    if let Some(v) = three_valued(lhs, rhs, true)? {
        return Ok(v);
    }
    logical(lhs, rhs, kinds, |x, y| x | y)
}

/// `Imp`, evaluated as its definition: `Not a Or b`.
///
/// Deriving it rather than hand-rolling a three-valued table is not just
/// tidier, it is what makes it *correct*. A hand-rolled version said
/// `255 Imp Null` was `Null`; the definition gives `Not 255 Or Null` =
/// `-256 Or Null`, and since `-256` is truthy [`or`] returns it. Excel
/// agrees. The measured endpoints still hold: `Null Imp True` is `True`
/// (a truthy consequent decides it) and `False Imp Null` is `True`
/// (`Not False` is truthy).
pub fn imp(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> VResult<Variant> {
    or(&not(lhs)?, rhs, kinds)
}

/// The shared half of [`and`] and [`or`]: when one side is `Null`, the other
/// decides the result if its truthiness is the deciding one.
///
/// Returns the deciding operand converted the way the bitwise operation
/// would have converted it: a `Boolean` stays a `Boolean`, and anything else
/// becomes the `Integer` or `Long` the operator works in.
///
/// The conversion is not cosmetic. `vb = 0.1 - 2147483647` then
/// `vb Or Null` is the **`Long`** `-2147483647` in Excel, not the `Double`
/// `-2147483646.9` -- the operand is rounded and narrowed before `Or` looks
/// at it, and returning it unchanged was a mismatch fuzz/fuzz_vba.py caught.
fn three_valued(lhs: &Variant, rhs: &Variant, deciding: bool) -> VResult<Option<Variant>> {
    let known = match (lhs.is_null(), rhs.is_null()) {
        (true, true) => return Ok(Some(Variant::Null)),
        (true, false) => rhs,
        (false, true) => lhs,
        (false, false) => return Ok(None),
    };
    if known.to_bool()? != deciding {
        return Ok(Some(Variant::Null));
    }
    if matches!(known, Variant::Boolean(_)) {
        return Ok(Some(known.clone()));
    }
    let class = match known.num_class().ok_or_else(VbaError::invalid_null)? {
        NumClass::Integer => NumClass::Integer,
        _ => NumClass::Long,
    };
    Variant::pack(bankers_round(known.to_f64()?), class).map(Some)
}

/// Whether a comparison operand was a compile-time constant.
///
/// Comparison between a string and a number depends on this, in the same way
/// arithmetic overflow does (see [`ArithMode`]) -- and the dependence is what
/// makes the rules look contradictory until you separate the cases.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operand {
    /// A plain literal. Only this makes the *string* side of a comparison
    /// strict; a constant expression that merely evaluates to a string does
    /// not.
    Literal,
    /// An expression built only from literals. Counts as constant on the
    /// numeric side, but not as a literal on the string side.
    ConstExpr,
    /// A call whose return type is declared numeric, so the compiler knows
    /// the type statically without the value being constant.
    Static,
    /// Anything involving a variable.
    Runtime,
}

impl Operand {
    fn is_const(self) -> bool {
        matches!(self, Operand::Literal | Operand::ConstExpr)
    }
}

/// Comparison, returning `None` when either side is `Null`.
///
/// Probe cases 53/54: `Empty = 0` and `Empty = ""` are both `True`, because
/// `Empty` compares as whichever the other operand is.
///
/// # String against number
///
/// Four measured rules, which only make sense once the literal-vs-variable
/// split is separated out. Every one of these was run against Excel:
///
/// | Operands | Rule | Evidence |
/// | --- | --- | --- |
/// | both constant | numeric; error 13 if the string does not parse | `"10" = 10` is `True`, `"" = 0` is error 13 |
/// | numeric constant, string variable | numeric, falling back below if it does not parse | `a = "2"` makes `a > 10` `False`; `a = ""` makes `a = 0` `False`, not an error |
/// | string constant, numeric variable | string, with the number via `CStr` | `b = 10` makes `"2" > b` `True` |
/// | both variables | **a number always sorts before a string** | `a = "1.5"`, `b = 1.5` makes `a = b` **`False`** |
///
/// That last row is the one that defeats every simpler theory: `"1.5"` and
/// `1.5` are equal both numerically and as text, and Excel still says they
/// differ -- because at runtime VBA does not convert either side, it orders
/// numbers before strings wholesale.
pub fn compare_ctx(
    lhs: &Variant,
    rhs: &Variant,
    lhs_kind: Operand,
    rhs_kind: Operand,
) -> VResult<Option<std::cmp::Ordering>> {
    use std::cmp::Ordering;
    if lhs.is_null() || rhs.is_null() {
        return Ok(None);
    }

    // `Empty` takes the other side's shape, so it is settled before the
    // string/number split below.
    let both_stringy = matches!(
        (lhs, rhs),
        (Variant::Str(_), Variant::Str(_))
            | (Variant::Str(_), Variant::Empty)
            | (Variant::Empty, Variant::Str(_))
    );
    if both_stringy {
        return Ok(Some(lhs.to_vba_string()?.cmp(&rhs.to_vba_string()?)));
    }

    let numeric = |v: &Variant| -> VResult<f64> { v.to_f64() };

    match (lhs, rhs) {
        (Variant::Str(text), other) | (other, Variant::Str(text)) => {
            let str_on_left = matches!(lhs, Variant::Str(_));
            let (str_kind, num_kind) = if str_on_left {
                (lhs_kind, rhs_kind)
            } else {
                (rhs_kind, lhs_kind)
            };
            // A String against a Boolean asks **two independent questions,
            // one per side**, and neither is about the values:
            //
            //   is the Boolean statically `Boolean`?  -> convert with `CBool`
            //   is the String  statically `String`?   -> if not, be lenient
            //
            // §16 of docs/vba-error-ordering.md had this as a 4x4 table with
            // one cell it could not explain. The cell was an artifact: it
            // read `(3# >= Empty)` as a folded constant, when what actually
            // matters is that `Empty` is a `Variant`, so the comparison is
            // not statically `Boolean` at all. See `interp::is_statically_typed`.
            //
            // The two questions, measured (§24):
            //
            //   "0" >= (3# >= CDbl(0))   True    static Boolean -> convert
            //   "0" >= ("1" >= -7)       True    likewise
            //   "0" >= IsEmpty(Empty)    True    a declared-Boolean call
            //   "0" >= (3# >= Empty)     False   a Variant operand -> text
            //   b = 1 : "0" >= (3# >= b) False   likewise
            //
            // and on the String side, against a Boolean that is *not* static:
            //
            //   "13"          <= ("" <> Empty)   True    text: "13" < "False"
            //   ("1" + "3")   <= ("" <> Empty)   True    a fold of literals is
            //   CStr(13)      <= ("" <> Empty)   True    a declared String is
            //   a = "13" : a  <= ("" <> Empty)   False   a Variant is not
            //   (Empty & "13")<= ("" <> Empty)   False   nor is a fold over Empty
            if matches!(other, Variant::Boolean(_)) {
                // Everything but `Runtime` is a Boolean the compiler knows:
                // a literal, `(Not True)`, a `CBool`/`IsEmpty` call, or a
                // comparison whose every operand is statically typed.
                if num_kind != Operand::Runtime {
                    // The conversion is `CBool`, and the comparison then runs
                    // on the Booleans as *numbers*, so True (-1) sorts below
                    // False (0). Measured:
                    //
                    //   a = "011"  makes  a = True    True
                    //   a = "0"    makes  a = False   True
                    //   a = "-1"   makes  a = True    True
                    //   a = "1.5"  makes  a = True    True
                    //   a = "011"  makes  a < False   True   -- -1 < 0
                    //   a = "011"  makes  a > False   False
                    //   ("011" < False)             True     -- a literal converts too
                    //   (False > "12")              True
                    //
                    // `bool_word` is folded in rather than handled beside
                    // this: `CBool` takes the words "True"/"False" as well.
                    let as_bool =
                        bool_word(text).or_else(|| parse_vba_number(text).ok().map(|n| n != 0.0));
                    if let Some(a) = as_bool {
                        let ord = cmp_f64(bool_as_number(a), bool_as_number(other.to_bool()?));
                        return Ok(Some(if str_on_left { ord } else { ord.reverse() }));
                    }
                    // What a string that will *not* convert does depends on
                    // how well the compiler knows its type. A declared
                    // `String` (`CStr`, `TypeName`) or a literal is error 13;
                    // an ordinary runtime Variant falls back to the runtime
                    // rule below -- **the number sorts first**, whatever the
                    // two actually are. Measured:
                    //
                    //   TypeName(32767) >= False              error 13
                    //   ("abc" < True)                        error 13
                    //   a = "abc" : a = True                  False
                    //   a = ""    : a = False                 False
                    //   a = TypeName(32767) : a >= (Not True) True
                    //   LCase("Integer") >= (Not True)        True
                    //   a = "ABC" : a > True                  True
                    //   a = "ABC" : a < True                  False
                    //   a = "ABC" : a >= False                True
                    //   Chr(65) > True                        True
                    //   Chr(65) > False                       True
                    //   StrConv("abc", 1) > True              True
                    //
                    // The first six are what make this about the *declared*
                    // type rather than the value: the same string through a
                    // Variant, and through a `Variant`-returning intrinsic,
                    // both reach this fallback rather than erroring.
                    //
                    // The last five are what say the fallback orders rather
                    // than compares as text, and they had to be chosen to
                    // tell those apart: this was written as a text comparison
                    // on the strength of the six above, every one of which
                    // holds either way, because their strings all happen to
                    // sort on the same side of `"True"`/`"False"` as the
                    // ordering rule puts them. `"ABC"` does not -- text makes
                    // `a > True` False, and Excel says True. Found by the
                    // reduction of the `StrReverse` case above.
                    if str_kind != Operand::Runtime {
                        return Err(VbaError::type_mismatch());
                    }
                    let ord = Ordering::Greater;
                    return Ok(Some(if str_on_left { ord } else { ord.reverse() }));
                }
                // The Boolean is *not* statically `Boolean` -- a variable, or
                // a comparison with a `Variant` operand. A statically typed
                // `String` then compares as text, with the Boolean rendered
                // "True"/"False", and never errors:
                //
                //   "0"          >= (3# >= Empty)   False   "0"    < "True"
                //   CStr(0)      >= (3# >= Empty)   False   "0"    < "True"
                //   TypeName(0)  >= (3# >= Empty)   False   "Integer" < "True"
                //   "0"          <  (3# >= Empty)   True    the same, inverted
                //   ("1" & "  3  ") <= ("" <> Empty) True   unconvertible, still text
                //
                // Anything else -- a Variant, or a fold over one -- falls
                // through to the numeric rules below, which is what makes
                // `((Empty & "1") <= ("" <> Empty))` False.
                if str_kind != Operand::Runtime {
                    let ord = text.as_str().cmp(other.to_vba_string()?.as_str());
                    return Ok(Some(if str_on_left { ord } else { ord.reverse() }));
                }
            }

            // A string whose type the compiler knows -- a constant, or a call
            // declared `As String` -- must convert. Only a Variant gets to
            // fall back to the ordering. This is the same "declared type, not
            // value" split the Boolean branch above turns on, and the two are
            // deliberately spelled the same way. Measured:
            //
            //   CStr("abc")       > 5        error 13
            //   TypeName(1)       > 5        error 13
            //   CStr("abc")       > CLng(1)  error 13
            //   StrReverse("abc") > 5        error 13
            //   Trim("abc")       > 5        True
            //   Chr(65)           > 5        True
            //   CStr("11")        > 5        True   -- it converts
            //
            // The `Static` half of this was missing, so every declared-String
            // intrinsic compared as an ordering against a number instead of
            // raising. `fuzz/fuzz_vba.py` reached it through `StrReverse`,
            // but `CStr` and `TypeName` were already wrong the same way.
            let str_typed = str_kind.is_const() || str_kind == Operand::Static;
            // Against a *statically typed* numeric partner the whole string
            // must parse; against a numeric constant only a leading run need
            // parse, as `Val` takes it.
            let ord = if num_kind == Operand::Static {
                match parse_vba_number(text) {
                    Ok(a) => cmp_f64(a, numeric(other)?),
                    // Only a string the compiler has typed has to parse. A
                    // runtime one that does not falls back to the ordering
                    // below, exactly as it does against a numeric constant:
                    // `CLng(a) < ("abc" & a)` is True, not error 13. An
                    // out-of-range string is a different failure (error 6,
                    // from the conversion) and still propagates.
                    Err(e) if str_typed || e.number != 13 => return Err(e),
                    Err(_) => Ordering::Greater,
                }
            } else if num_kind.is_const() {
                match numeric_prefix(text) {
                    Some(a) => cmp_f64(a, numeric(other)?),
                    // No numeric prefix at all. A typed string is an error; a
                    // runtime one falls back to the ordering below.
                    None if str_typed => return Err(VbaError::type_mismatch()),
                    None => Ordering::Greater,
                }
            } else if str_typed {
                // A *runtime* number, against a string whose type the compiler
                // knows: text, with the number via `CStr`. A declared `String`
                // behaves exactly as a literal does here -- the same split as
                // the strictness above, and it was missing on this branch for
                // the same reason, so every `CStr`/`StrReverse`/fold ordered
                // against a runtime number instead of comparing. Measured
                // (§27), with `a` a variable so the number is never static:
                //
                //   a = 5  : a < "10"           False  -- "5" sorts above "1"
                //   a = 5  : a < CStr(10)       False  -- declared String, same
                //   a = 5  : a < (CStr(1) & "0") False -- a fold of them, same
                //   b = 10 : a = 5 : a < CStr(b) False -- the value may be runtime
                //   a = 5  : a < Trim("10")     True   -- a Variant orders
                //   a = -2 : a < CStr("")       False  -- and it never errors
                //   a = -2 : a < CStr("abc")    True
                text.as_str().cmp(other.to_vba_string()?.as_str())
            } else {
                // Both runtime: the number sorts first, whatever it is.
                Ordering::Greater
            };
            // `ord` was computed with the string on the left; flip it if the
            // string was actually the right operand.
            Ok(Some(if str_on_left { ord } else { ord.reverse() }))
        }
        _ => Ok(Some(cmp_f64(numeric(lhs)?, numeric(rhs)?))),
    }
}

/// A `Boolean` as the number it *is*, which is what ordering compares.
///
/// `True` is -1, so it sorts below `False`. Rust's own `bool: Ord` has it the
/// other way round, and using that here reversed every `<`/`>` between a
/// string and a Boolean.
fn bool_as_number(b: bool) -> f64 {
    if b { -1.0 } else { 0.0 }
}

fn cmp_f64(a: f64, b: f64) -> std::cmp::Ordering {
    a.partial_cmp(&b).unwrap_or(std::cmp::Ordering::Equal)
}

/// Comparison between two runtime values, for callers with no constant
/// information (`Select Case`, and the interpreter's internal uses).
pub fn compare(lhs: &Variant, rhs: &Variant) -> VResult<Option<std::cmp::Ordering>> {
    compare_ctx(lhs, rhs, Operand::Runtime, Operand::Runtime)
}

#[cfg(test)]
mod tests {
    use super::*;

    // Every assertion below cites the numbered case in
    // `fuzz/vba_variant_probe.bas` that produced it against Excel 16.112.
    // The probe prints `TypeName(v) & "|" & CStr(v)`, so both halves of each
    // assertion -- type and value -- are measured, not assumed.

    fn shows(v: &Variant) -> String {
        format!(
            "{}|{}",
            v.type_name(),
            v.to_vba_string().unwrap_or_default()
        )
    }

    #[test]
    fn literal_typing_matches_excel() {
        // 1, 2, 15-20
        assert_eq!(shows(&Variant::from_literal(1.0, false)), "Integer|1");
        assert_eq!(
            shows(&Variant::from_literal(32767.0, false)),
            "Integer|32767"
        );
        assert_eq!(shows(&Variant::from_literal(32768.0, false)), "Long|32768");
        assert_eq!(
            shows(&Variant::from_literal(2147483647.0, false)),
            "Long|2147483647"
        );
        assert_eq!(
            shows(&Variant::from_literal(2147483648.0, false)),
            "Double|2147483648"
        );
        assert_eq!(shows(&Variant::from_literal(1.5, true)), "Double|1.5");
        // Case 19: `1E5` is Double even though 100000 alone would be Long.
        assert_eq!(
            shows(&Variant::from_literal(100000.0, true)),
            "Double|100000"
        );
        assert_eq!(
            shows(&Variant::from_literal(100000.0, false)),
            "Long|100000"
        );
    }

    #[test]
    fn constant_arithmetic_overflows_but_runtime_arithmetic_promotes() {
        // The correction fuzz/fuzz_vba.py forced. Between two literals,
        // Excel really does raise error 6; with a variable involved it widens
        // Integer -> Long -> Double instead.
        assert_eq!(
            shows(
                &add(
                    &Variant::Integer(32767),
                    &Variant::Integer(1),
                    ArithMode::Promote
                )
                .unwrap()
            ),
            "Long|32768"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Long(2147483647),
                    &Variant::Integer(1),
                    ArithMode::Promote
                )
                .unwrap()
            ),
            "Double|2147483648"
        );
        assert_eq!(
            shows(
                &mul(
                    &Variant::Long(100000),
                    &Variant::Long(100000),
                    ArithMode::Promote
                )
                .unwrap()
            ),
            "Double|10000000000"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Integer(1),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Integer|2"
        );
        assert_eq!(
            add(
                &Variant::Integer(32767),
                &Variant::Integer(1),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            6
        );
        assert_eq!(
            add(
                &Variant::Long(2147483647),
                &Variant::Integer(1),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            6
        );
        assert_eq!(
            mul(
                &Variant::Long(100000),
                &Variant::Long(100000),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            6
        );
    }

    #[test]
    fn plus_concatenates_only_when_both_sides_are_strings() {
        // Cases 35-37.
        assert_eq!(
            shows(
                &add(
                    &Variant::Str("1".into()),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Double|2"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Str("1".into()),
                    &Variant::Str("2".into()),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "String|12"
        );
        assert_eq!(
            add(
                &Variant::Str("abc".into()),
                &Variant::Integer(1),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            13
        );
        // Case 40: surrounding whitespace is ignored.
        assert_eq!(
            shows(
                &add(
                    &Variant::Str("  3  ".into()),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Double|4"
        );
    }

    #[test]
    fn division_is_always_double() {
        // Cases 24, 25.
        assert_eq!(
            shows(&div(&Variant::Integer(1), &Variant::Integer(2)).unwrap()),
            "Double|0.5"
        );
        assert_eq!(
            shows(&div(&Variant::Integer(4), &Variant::Integer(2)).unwrap()),
            "Double|2"
        );
        assert_eq!(
            div(&Variant::Integer(1), &Variant::Integer(0))
                .unwrap_err()
                .number,
            11
        );
    }

    #[test]
    fn int_div_and_mod_round_their_operands_first() {
        // Cases 26-31. `7.6 \ 2` is 4, not 3 -- and Long, not Integer.
        assert_eq!(
            shows(&int_div(&Variant::Integer(7), &Variant::Integer(2)).unwrap()),
            "Integer|3"
        );
        assert_eq!(
            shows(&int_div(&Variant::Integer(-7), &Variant::Integer(2)).unwrap()),
            "Integer|-3"
        );
        assert_eq!(
            shows(&int_div(&Variant::Double(7.6), &Variant::Integer(2)).unwrap()),
            "Long|4"
        );
        assert_eq!(
            shows(&modulo(&Variant::Integer(7), &Variant::Integer(2)).unwrap()),
            "Integer|1"
        );
        assert_eq!(
            shows(&modulo(&Variant::Integer(-7), &Variant::Integer(2)).unwrap()),
            "Integer|-1"
        );
        assert_eq!(
            shows(&modulo(&Variant::Double(7.6), &Variant::Integer(2)).unwrap()),
            "Long|0"
        );
    }

    #[test]
    fn pow_is_always_double() {
        // Case 32.
        assert_eq!(
            shows(
                &pow(
                    &Variant::Integer(2),
                    &Variant::Integer(2),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Double|4"
        );
    }

    #[test]
    fn booleans_are_minus_one() {
        // Cases 41, 42, 47.
        assert_eq!(
            shows(
                &add(
                    &Variant::Boolean(true),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Integer|0"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Boolean(true),
                    &Variant::Boolean(true),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Integer|-2"
        );
    }

    #[test]
    fn logical_operators_are_bitwise_unless_both_sides_are_boolean() {
        // Cases 43, 45, 46.
        assert_eq!(
            shows(
                &logical(
                    &Variant::Boolean(true),
                    &Variant::Boolean(false),
                    (Operand::Literal, Operand::Literal),
                    |a, b| a & b
                )
                .unwrap()
            ),
            "Boolean|False"
        );
        assert_eq!(
            shows(
                &logical(
                    &Variant::Integer(5),
                    &Variant::Integer(3),
                    (Operand::Literal, Operand::Literal),
                    |a, b| a & b
                )
                .unwrap()
            ),
            "Integer|1"
        );
        assert_eq!(shows(&not(&Variant::Integer(5)).unwrap()), "Integer|-6");
        assert_eq!(
            shows(&not(&Variant::Boolean(true)).unwrap()),
            "Boolean|False"
        );
    }

    #[test]
    fn null_propagates_through_arithmetic_but_not_concatenation() {
        // Cases 50-52.
        assert!(
            add(&Variant::Null, &Variant::Integer(1), ArithMode::Constant)
                .unwrap()
                .is_null()
        );
        assert_eq!(
            shows(&concat(&Variant::Null, &Variant::Str("a".into())).unwrap()),
            "String|a"
        );
    }

    #[test]
    fn empty_is_both_zero_and_the_empty_string() {
        // Cases 48, 49, 53, 54.
        assert_eq!(
            shows(&add(&Variant::Empty, &Variant::Integer(1), ArithMode::Constant).unwrap()),
            "Integer|1"
        );
        assert_eq!(
            shows(&concat(&Variant::Empty, &Variant::Str("a".into())).unwrap()),
            "String|a"
        );
        assert_eq!(
            compare(&Variant::Empty, &Variant::Integer(0)).unwrap(),
            Some(std::cmp::Ordering::Equal)
        );
        assert_eq!(
            compare(&Variant::Empty, &Variant::Str(String::new())).unwrap(),
            Some(std::cmp::Ordering::Equal)
        );
        // `+` against a String concatenates rather than coercing it, which is
        // what makes `Empty + "a"` a value at all instead of error 13, and
        // keeps `Empty + "1"` a String. `Empty + Empty` is the Integer 0.
        for (l, r) in [
            (Variant::Empty, Variant::Str("a".into())),
            (Variant::Str("a".into()), Variant::Empty),
        ] {
            assert_eq!(
                shows(&add(&l, &r, ArithMode::Constant).unwrap()),
                "String|a"
            );
        }
        assert_eq!(
            shows(
                &add(
                    &Variant::Empty,
                    &Variant::Str("1".into()),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "String|1"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Empty,
                    &Variant::Str(String::new()),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "String|"
        );
        assert_eq!(
            shows(&add(&Variant::Empty, &Variant::Empty, ArithMode::Constant).unwrap()),
            "Integer|0"
        );
        // Only `+`. The other operators still make the string a number.
        assert!(
            sub(
                &Variant::Empty,
                &Variant::Str("a".into()),
                ArithMode::Constant
            )
            .is_err()
        );
        assert!(
            mul(
                &Variant::Empty,
                &Variant::Str("a".into()),
                ArithMode::Constant
            )
            .is_err()
        );
    }

    #[test]
    fn rounding_is_bankers_not_half_away_from_zero() {
        // Cases 55-59. f64::round would give 1, 2, 3, -1, -2 here.
        assert_eq!(bankers_round(0.5), 0.0);
        assert_eq!(bankers_round(1.5), 2.0);
        assert_eq!(bankers_round(2.5), 2.0);
        assert_eq!(bankers_round(-0.5), -0.0);
        assert_eq!(bankers_round(-1.5), -2.0);
        assert_eq!(bankers_round(3.5), 4.0);
    }

    #[test]
    fn concat_stringifies_numbers() {
        // Case 38.
        assert_eq!(
            shows(&concat(&Variant::Integer(1), &Variant::Integer(2)).unwrap()),
            "String|12"
        );
    }

    #[test]
    fn negative_zero_keeps_its_sign() {
        // Excel's CStr renders -0.0 as "-0". Normalising it to "0" was a real
        // mismatch fuzz/fuzz_vba.py found on its first working run.
        assert_eq!(format_number(-0.0), "-0");
        // Positive zero stays unsigned.
        assert_eq!(format_number(0.0), "0");
        assert_eq!(format_number(-0.4 + 0.4), "0");
        assert_eq!(
            shows(&neg(&Variant::Double(0.0), ArithMode::Constant).unwrap()),
            "Double|-0"
        );
    }

    #[test]
    fn cstr_shows_exactly_fifteen_significant_digits() {
        // Excel is consistently 15 s.f. The previous implementation derived
        // the decimal count from log10 and got 16 digits in exponent form and
        // 14 in fixed form; the fuzzer caught it on every seed.
        assert_eq!(format_number(9.156321295314252e-5), "9.15632129531425E-05");
        assert_eq!(
            format_number(-0.000985221674876847),
            "-0.000985221674876847"
        );
        assert_eq!(format_number(0.000457247370827618), "0.000457247370827618");
        // The fixed/exponent boundary: 1e-4 stays fixed, below it switches.
        assert_eq!(format_number(0.0001), "0.0001");
        assert_eq!(format_number(0.00001), "1E-05");
        assert_eq!(format_number(1e15), "1000000000000000");
        assert_eq!(format_number(1e16), "1E+16");
    }

    #[test]
    fn number_formatting_matches_cstr() {
        // Case 63, plus the shapes CStr has to get right for the fuzzer to
        // compare strings at all.
        assert_eq!(format_number(1.5), "1.5");
        assert_eq!(format_number(2.0), "2");
        assert_eq!(format_number(0.0), "0");
        assert_eq!(format_number(-1.25), "-1.25");
        assert_eq!(format_number(1000.0), "1000");
        assert_eq!(format_number(0.1), "0.1");
    }

    #[test]
    fn string_to_number_accepts_what_vba_accepts() {
        // Case 64, plus the coercion cases.
        assert_eq!(parse_vba_number("1e3").unwrap(), 1000.0);
        assert_eq!(parse_vba_number("  3  ").unwrap(), 3.0);
        assert_eq!(parse_vba_number("1.5").unwrap(), 1.5);
        assert_eq!(parse_vba_number("&HFF").unwrap(), 255.0);
        // An empty string is a type mismatch, not a zero: measured via
        // `"" = 0`, `"" < 0` and `Not ""`, all of which are error 13.
        assert_eq!(parse_vba_number("").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("abc").unwrap_err().number, 13);
        // A trailing sign, which VBA reads as the number's sign. Measured:
        // CDbl("1-") is -1, CInt("1-") is -1, CDbl("1E2-") is -100,
        // CDbl("1 -") is -1, and IsNumeric("1-") is True.
        assert_eq!(parse_vba_number("1-").unwrap(), -1.0);
        assert_eq!(parse_vba_number("1+").unwrap(), 1.0);
        assert_eq!(parse_vba_number("2.5-").unwrap(), -2.5);
        assert_eq!(parse_vba_number("1 -").unwrap(), -1.0);
        assert_eq!(parse_vba_number("1E2-").unwrap(), -100.0);
        // It replaces the sign rather than compounding one.
        assert_eq!(parse_vba_number("-1-").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("1--").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("-").unwrap_err().number, 13);
    }
}
