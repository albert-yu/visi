Here’s a practical plan for adding type checking to Excel formula/spreadsheet syntax in `visi-core`.

## 1. Define the goal narrowly first

Excel is not statically typed in the usual sense. It has dynamic values, implicit coercions, array/range behavior, and function-specific exceptions.

So the first version should be:

> A formula analyzer that reports likely Excel errors and invalid reference/function usage, without rejecting formulas or changing evaluation behavior.

Treat it as diagnostics, not enforcement.

## 2. Add a new internal module

Add something like:

```text
visi-core/src/core/typecheck.rs
```

Keep it `pub(crate)` initially.

Core types:

```rust
enum FormulaType {
    Any,
    Blank,
    Number,
    Text,
    Boolean,
    Error,
    Scalar(Box<FormulaType>),
    Reference(ReferenceType),
    Array {
        item: Box<FormulaType>,
        rows: Option<usize>,
        cols: Option<usize>,
    },
    Union(Vec<FormulaType>),
}

enum ReferenceType {
    Cell,
    Range {
        rows: Option<usize>,
        cols: Option<usize>,
    },
    Column,
    Row,
    Structured,
}

enum TypeSeverity {
    Error,
    Warning,
    Note,
}

struct TypeDiagnostic {
    message: String,
    severity: TypeSeverity,
}
```

Do not add source comments, per repo guidance.

## 3. Build a type-checking context

The checker needs similar context to evaluation:

```rust
struct TypeContext<'a> {
    sheet: &'a Sheet,
    workbook_context: Option<&'a Context<'a>>,
    row: Option<usize>,
    col: Option<usize>,
    let_bindings: HashMap<String, FormulaType>,
}
```

It should know:

- Current sheet
- Other sheets
- Table/structured reference metadata
- Current row/column for `[@Column]`
- `LET`/`LAMBDA` bindings
- Existing cell result types when available

## 4. Traverse `parser::Expr`

Implement:

```rust
fn infer_expr_type(expr: &Expr, ctx: &mut TypeContext) -> (FormulaType, Vec<TypeDiagnostic>)
```

Initial rules:

### Literals

- `Number` → `Number`
- `String` → `Text`
- `Boolean` → `Boolean`
- `Error` → `Error`

### References

- `CellRef` → `Reference(Cell)`
- `RangeRef` → `Reference(Range { rows, cols })`
- `StructuredRef` → `Reference(Structured)` or `Array`
- Unknown sheet/table/column → diagnostic

### Operators

Map operators to expected input capabilities:

| Operator | Expected | Result |
|---|---|---|
| `+ - * / ^ %` | numeric coercible scalar | number or error |
| `&` | text coercible scalar | text |
| comparisons | scalar comparable | boolean |
| `:` / intersection / union | references | reference/union |
| implicit intersection | range/array/reference | scalar-ish |
| spill | dynamic array/reference | array/scalar |

The checker should understand “coercible to number” instead of only “is number”, because Excel accepts things like numeric text in many contexts.

## 5. Add coercion capability helpers

Model Excel coercion separately from concrete types:

```rust
enum CoercionTarget {
    Number,
    Text,
    Boolean,
    Reference,
    Array,
}
```

Then implement:

```rust
fn can_coerce(from: &FormulaType, to: CoercionTarget, mode: CoercionMode) -> CoercionResult
```

Important modes:

- Direct argument vs computed argument
- Literal vs cell/range reference
- Strict numeric
- Stat-family numeric
- Text-family coercion
- Blank handling: zero, empty string, skip, reject

This matters because the evaluator already has lots of function-specific coercion behavior in:

```text
visi-core/src/core/engine/sheet/mod.rs
visi-core/src/core/engine/sheet/functions/*.rs
```

The type checker should mirror that behavior gradually.

## 6. Add function signature metadata

Create an internal registry:

```rust
struct FunctionSig {
    name: &'static str,
    min_args: usize,
    max_args: Option<usize>,
    args: &'static [ArgSig],
    return_type: ReturnTypeRule,
    eval_policy: EvalPolicy,
}

enum ArgSig {
    Any,
    Number,
    Text,
    Boolean,
    Reference,
    Range,
    Array,
    Repeating(&'static ArgSig),
}
```

Start with high-value functions:

- Arithmetic/stat basics: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`
- Logical: `IF`, `IFS`, `AND`, `OR`
- Lookup/reference: `INDEX`, `MATCH`, `XLOOKUP`, `VLOOKUP`
- Text: `LEFT`, `RIGHT`, `MID`, `CONCAT`, `TEXT`
- Dynamic arrays: `FILTER`, `SORT`, `UNIQUE`, `TAKE`, `DROP`

Special-case lazy functions first:

- `IF`
- `IFERROR`
- `IFNA`
- `CHOOSE`
- `LET`
- `LAMBDA`

These cannot be checked as simple eager signatures.

## 7. Wire into formula parsing/evaluation

Current parsing entry point:

```rust
parse_excel_formula(input: &str) -> Result<Expr, String>
```

Add a sibling API internally:

```rust
pub(crate) fn check_excel_formula(
    input: &str,
    sheet: &Sheet,
    context: Option<&Context>,
    row: Option<usize>,
    col: Option<usize>,
) -> Vec<TypeDiagnostic>
```

This should:

1. Parse formula
2. If parse fails, return syntax diagnostic
3. If parse succeeds, infer/check the AST
4. Return diagnostics without evaluating

Later, expose a public API only if needed. If exposed from `core/mod.rs`, treat it as a semver commitment.

## 8. Add tests without Excel dependency

Add tests in a new file, for example:

```text
visi-core/src/core/typecheck.rs
```

or:

```text
visi-core/src/core/engine/tests/typecheck.rs
```

Test cases:

- Valid arithmetic: `=A1+1`
- Invalid arithmetic: `="abc"+1`
- Valid numeric text warning: `="1"+1`
- Bad arity: `=SUM()`, `=LEFT("abc")`
- Unknown function: `=NO_SUCH_FN(1)`
- Unknown sheet/table/column
- Range passed to scalar-only function
- Array shape mismatch where known
- Lazy checks: `=IF(TRUE, 1, "x"+A1)`

## 9. Use fuzzing to align behavior

For any uncertain rule, use the existing Excel differential fuzzing workflow.

Good targets:

- Numeric text coercion
- Blank coercion
- Error propagation order
- Boolean coercion
- Range vs scalar handling
- Dynamic array spill behavior
- Function arity and `#VALUE!`/`#N/A` distinctions

Every discovered mismatch should become a normal Rust regression test.

## 10. Rollout order

Recommended milestones:

1. **MVP diagnostics**
   - parse errors
   - unknown function
   - function arity
   - unknown sheet/table/column

2. **Operator checking**
   - arithmetic
   - concatenation
   - comparison
   - reference operators

3. **Basic function signatures**
   - common math/text/logical functions

4. **Reference/range/array shape checking**
   - `INDEX`, `FILTER`, `SORT`, `TAKE`, `DROP`, `HSTACK`, `VSTACK`

5. **LET/LAMBDA scope checking**

6. **Optional CLI/Python exposure**
   - Only after the internal API is stable enough.

## Key design principle

Do not make type checking affect evaluation at first. Excel’s semantics are too permissive and function-specific. Start with advisory diagnostics, keep evaluator behavior unchanged, and tighten the checker only when backed by tests or Excel probes.