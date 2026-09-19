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
| `&` | string concatenation | Coerces operands with the same text rendering used by `CONCAT`/`CONCATENATE`; error operands propagate. |
| `=`, `<>`, `<`, `>`, `<=`, `>=` | comparisons | `==` and `!=` are also accepted by the lexer. |
| `:` | range construction | Implemented for A1 cell ranges, whole-row ranges, whole-column ranges, and sheet-qualified forms. It is not a general binary AST operator. |

## Missing or partial

| Operator | Excel meaning | Current behavior |
| --- | --- | --- |
| `%` | postfix percent, e.g. `50%` -> `0.5` | Not lexed; formulas containing `%` fail with an unexpected-character parse error. |
| space | reference intersection, e.g. `SUM(A1:C3 B2:D4)` | Not represented; formula lexing discards whitespace, so intersection cannot be expressed. |
| `,` | reference union outside argument separation, e.g. `SUM((A1:A2,C1:C2))` | Comma is only parsed as a function/list argument separator. There is no union-reference AST or evaluator support. |

The parser's internal `Op` enum currently has evaluator support for every variant it defines. The gaps above are Excel operators that do not yet have parser/AST/evaluator representation, or only have specialized range parsing rather than a first-class binary operator.
