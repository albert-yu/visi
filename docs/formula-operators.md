_The following content is LLM-generated._

# Formula operator support

This document tracks worksheet-formula unary and binary operators against Excel's operator families.

## Implemented

| Operator | Excel meaning | Notes |
| --- | --- | --- |
| `+` | addition / unary plus | Unary plus is parsed as a no-op. |
| `-` | subtraction / unary negation | Unary negation is implemented for numeric values. |
| `*` | multiplication |  |
| `/` | division | Division by zero returns `#DIV/0!`. |
| `^` | exponentiation | `**` is also accepted by the lexer. |
| `%` | postfix percent | Converts the preceding operand to a number and divides by 100. |
| `&` | string concatenation | Coerces operands with the same text rendering used by `CONCAT`/`CONCATENATE`; error operands propagate. |
| `=`, `<>`, `<`, `>`, `<=`, `>=` | comparisons | `==` and `!=` are also accepted by the lexer. |
| `:` | range construction | Implemented for A1 cell ranges, whole-row ranges, whole-column ranges, and sheet-qualified forms. It is parsed as part of references rather than as a general binary AST operator. |
| space | reference intersection | Implemented for cell/range references, including whole-row and whole-column ranges. Empty intersections return `#NULL!`. |
| `,` | reference union outside argument separation | Implemented inside parenthesized reference expressions such as `SUM((A1:A2,C1:C2))`. Function argument commas remain argument separators. |

## Missing or partial

No known worksheet-formula operator family is intentionally missing from the parser/evaluator. The reference operators still use the engine's current flat list representation for multi-cell results, so shape-sensitive behavior may be narrower than Excel in contexts outside aggregation functions.
