# LET/LAMBDA and Higher-Order Array Functions

This document describes the **current, implemented** semantics of Excel-style `LET` / `LAMBDA`
and the higher-order array helpers (`MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`).
The goal is to keep Excel-compat behavior durable across refactors: if you change behavior,
update the tests and *then* update this doc.

> Scope: this doc focuses on evaluation semantics and error behavior. For parser/tokenization
> details, see `docs/01-formula-engine.md`.

---

## Supported syntax

### `LET`

```
LET(name1, value1, [name2, value2, ...], calculation)
```

- Must have **at least 3 arguments** and an **odd** number of arguments.
- `nameN` must be a **bare identifier** (a name reference like `x`, not `"x"`, not `A1`).
- Bindings are evaluated **left-to-right** (sequential binding).

Examples:

```excel
=LET(a, 2, b, a*3, b+1)          // 7
=LET(x, 1/0, IFERROR(x, 0))      // 0   (errors can be bound and recovered from)
```

### `LAMBDA`

```
LAMBDA(param1, [param2, ...], body)
```

- Parameter names must be **bare identifiers**.
- Parameter names are **case-insensitive** (Unicode uppercasing / casefolding).
- Parameter names must be **unique** (after casefolding).
- Evaluating `LAMBDA(...)` produces a **lambda value**; if you return it directly from a cell
  without invoking it, it formats as `#CALC!` (matching Excel’s behavior for uninvoked lambdas).

Examples:

```excel
=LAMBDA(x, x+1)(2)                     // 3
=LAMBDA(x, LAMBDA(y, x+y))(1)(2)       // 3
```

### Invocation (calling a lambda)

The engine supports two ways to invoke lambdas:

1. **Postfix call syntax**: `expr(args...)`

   Typically used for inline lambdas and parenthesized expressions:

   ```excel
   =LAMBDA(x, x+1)(2)
   =LET(f, LAMBDA(x, x+1), f)(2)
   =LET(FACT, LAMBDA(n, IF(n<=1, 1, n*FACT(n-1))), (FACT)(5))   // recursion preserved
   ```

2. **Named call syntax**: `Name(args...)`

   `Name` can refer to:
   - a `LET` binding,
   - a workbook defined name whose definition evaluates to a lambda,
   - a captured lexical binding.

   If `Name` also matches a built-in function, **a lexical lambda binding wins**.

   ```excel
   =LET(SUM, LAMBDA(x, x+1), SUM(1))    // 2 (calls the lambda, not the builtin SUM)
   ```

Lambda calls may also return **reference values** (e.g. a range). Those are preserved as
references (not coerced into value arrays) so they can be consumed by reference-aware functions:

```excel
=LET(f, LAMBDA(r, r), OFFSET(f(A1:A3), 0, 0, 1, 1))
```

---

## Scoping model

### Name normalization

`LET` variable names and `LAMBDA` parameter names are normalized by:

- `trim()` (leading/trailing whitespace ignored)
- Unicode casefolding (uppercasing, so `Ü` and `ü` match)

This matches how the engine performs lexical lookup.

### `LET` sequential binding

`LET` creates a new **local scope frame** for the duration of the `LET(...)` evaluation.

- Each `valueN` is evaluated in a scope that includes all earlier bindings
  (`name1..name{N-1}`).
- The final `calculation` expression is evaluated in a scope that includes **all** bindings.
- Reusing the same name in a single `LET` overwrites the previous binding (“last write wins”).
- `LET` bindings shadow outer lexical bindings and workbook names.

### Lambda parameter binding + shadowing

When a lambda is invoked:

- Each parameter is bound to the corresponding argument value.
- **Too many arguments** (`args.len() > params.len()`) returns `#VALUE!`.
- **Too few arguments** is allowed:
  - missing parameters are bound to `BLANK`
  - missing parameters are marked as “omitted” (see `ISOMITTED` below)
- Parameter bindings shadow captured variables of the same name.

### Closure capture

When `LAMBDA(...)` is evaluated, it captures a snapshot of the current lexical environment:

- All currently-bound `LET` bindings and lambda call frames are captured.
- Values are stored as engine `Value`s:
  - scalars are captured by value,
  - references are captured as reference objects (sheet + coordinates) and are dereferenced
    when used (so cell edits are seen by the lambda),
  - lambdas can be captured (higher-order usage).
- Internal bookkeeping bindings used to implement `ISOMITTED` are *not* captured.

### Recursion + recursion guard

To support recursion, the evaluator binds the **call name** to the invoked lambda inside the
lambda’s call scope.

- When invoked as `FACT(5)`, the name `FACT` is bound to the lambda value, so the body can
  call itself as `FACT(...)`.
- When invoked as `(FACT)(5)`, the call name is still preserved.
- When invoked as `LAMBDA(...)(...)`, the engine uses an internal sentinel name that cannot
  be referenced from formulas. For recursion, introduce a name via `LET`.

The engine enforces a recursion depth limit of **64** nested lambda calls. Exceeding this limit
returns `#CALC!`.

---

## Higher-order array helpers (MAP/REDUCE/SCAN/BYROW/BYCOL/MAKEARRAY)

Unless otherwise noted, “array arguments” are coerced as follows:

- Scalar → 1×1 array
- Range reference → rectangular array of cell values
- Multi-area reference (union) → `#VALUE!`

All iteration over arrays is in **row-major order** (row 1 left-to-right, then row 2, etc).

### `MAP`

```
MAP(array1, [array2, ...], lambda)
```

- Requires at least 2 arguments (≥1 array + lambda).
- `lambda` must be a lambda value.
- The lambda must accept exactly **N parameters**, where N is the number of input arrays.
- When the `lambda` argument is a bare identifier (e.g. a `LET` binding like `FACT`), the engine
  invokes it using that identifier as the **call name** so recursion works and LET-bound lambdas
  can shadow built-ins.
- **Broadcasting**:
  - The output shape is taken from the first non-1×1 input array.
  - Every other input must be either 1×1 or exactly the same shape.
  - Otherwise: `#VALUE!`.
- The lambda is invoked once per output cell.
- Lambda results are **scalarized**:
  - scalar / 1×1 array → ok
  - larger array or spill marker → `#VALUE!` (for that element)

### `BYROW`

```
BYROW(array, lambda)
```

- `lambda` must take **exactly 1** parameter.
- The lambda receives a 1×N array containing the current row.
- Returns an R×1 column vector (one result per input row).
- Lambda results are scalarized (same rule as `MAP`).

### `BYCOL`

```
BYCOL(array, lambda)
```

- `lambda` must take **exactly 1** parameter.
- The lambda receives an R×1 array containing the current column.
- Returns a 1×C row vector (one result per input column).
- Lambda results are scalarized.

### `MAKEARRAY`

```
MAKEARRAY(rows, cols, lambda)
```

- `rows` / `cols` are coerced to integers.
- `rows <= 0` or `cols <= 0` → `#VALUE!`.
- If the requested size overflows internal limits (conversion to `usize` or `rows*cols`),
  the result is `#NUM!`.
- `lambda` must take **exactly 2** parameters: `(row_index, col_index)`.
  - Indices are passed as **1-based** numbers.
- Lambda results are scalarized.

### `REDUCE`

Supported forms:

```
REDUCE(array, lambda)                       // initial omitted
REDUCE(initial_value, array, lambda)
```

- `lambda` must take **exactly 2** parameters: `(accumulator, value)`.
- Iteration order is row-major over the input array.
- If `initial_value` is omitted:
  - the first element becomes the initial accumulator,
  - if the array is empty → `#CALC!`.
- The accumulator is **not scalarized**:
  - the lambda can return dynamic arrays and continue accumulating over them
    (enables patterns like `VSTACK` / `HSTACK` accumulation).

### `SCAN`

Supported forms:

```
SCAN(array, lambda)                         // initial omitted
SCAN(initial_value, array, lambda)
```

- `lambda` must take **exactly 2** parameters: `(accumulator, value)`.
- Returns an array with the **same shape** as the input.
- If `initial_value` is omitted:
  - the first element (scalarized) becomes the initial accumulator and the first output element,
  - if the array is empty → `#CALC!`.
- `initial_value` (when provided) is scalarized before the scan begins.
- Each lambda result is scalarized before being used as the next accumulator / output element.

### `ISOMITTED` (supporting `LAMBDA`)

```
ISOMITTED(paramName)
```

- `paramName` must be a bare identifier.
- Returns `TRUE` only if the parameter was omitted because the caller passed fewer arguments
  than parameters.
- A blank placeholder argument (e.g. `f(1,)`) is **not** considered omitted.

---

## Error behavior matrix

`#VALUE!` is used for most type/shape/arity errors. `#CALC!` is reserved for calc-engine
conditions like recursion depth and empty-array reductions, and `#NUM!` is used for size
overflows.

| Function / operation | Arity errors | Identifier errors | Type/shape errors | Other notes |
|---|---|---|---|---|
| `LET` | `<3` args, or even arg count → `#VALUE!` | `nameN` not a bare identifier → `#VALUE!` | N/A | Bindings are sequential; bound errors can be recovered with `IFERROR` |
| `LAMBDA` | 0 args → `#VALUE!` | params not bare identifiers, or duplicate params → `#VALUE!` | N/A | Returning a lambda to the grid formats as `#CALC!` |
| Lambda call (`Name(...)` or `expr(...)`) | too many args → `#VALUE!` | N/A | calling a non-lambda value → `#VALUE!` | too much recursion → `#CALC!`; missing args become blank + “omitted” |
| `ISOMITTED` | not 1 arg → `#VALUE!` | arg not a bare identifier → `#VALUE!` | N/A | Detects omission, not explicit blank |
| `MAP` | `<2` args → `#VALUE!` | N/A | non-lambda → `#VALUE!`; param-count mismatch → `#VALUE!`; shape mismatch → `#VALUE!`; non-scalar lambda result → `#VALUE!` | Broadcast supports only 1×1 or exact-shape inputs |
| `BYROW` / `BYCOL` | not 2 args → `#VALUE!` | N/A | non-lambda → `#VALUE!`; param-count ≠ 1 → `#VALUE!`; non-scalar lambda result → `#VALUE!` | Each row/col is passed as an array value |
| `MAKEARRAY` | not 3 args → `#VALUE!` | N/A | non-lambda → `#VALUE!`; param-count ≠ 2 → `#VALUE!` | `rows/cols <= 0` → `#VALUE!`; size overflow → `#NUM!` |
| `REDUCE` | not 2 or 3 args → `#VALUE!` | N/A | non-lambda → `#VALUE!`; param-count ≠ 2 → `#VALUE!` | empty array with omitted initial → `#CALC!` |
| `SCAN` | not 2 or 3 args → `#VALUE!` | N/A | non-lambda → `#VALUE!`; param-count ≠ 2 → `#VALUE!`; non-scalar accumulator/lambda result → `#VALUE!` | empty array → `#CALC!` |

---

## Known gaps vs Excel (documented behavior differences)

- `MAP` broadcasting is currently limited to **1×1 scalars** or **exact shape** matches. Excel’s
  full broadcasting behavior (e.g. vector-style expansion) is not implemented.

If additional incompatibilities are discovered, they should be captured here along with a
repro formula and (ideally) a regression test.

---

## Performance notes

- `MAP` and `MAKEARRAY` invoke the lambda once per output cell (`O(rows*cols)` calls) and are
  currently evaluated in row-major order.
- `BYROW` / `BYCOL` allocate a fresh row/column array value for each invocation.
- `REDUCE` and `SCAN` flatten the input array in row-major order before iterating.
- Lambda recursion is capped (64) to avoid runaway evaluation.
