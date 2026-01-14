# Formula Engine

## Overview

The formula engine is the computational heart of the spreadsheet. It must parse, evaluate, and maintain dependencies for all formulas with **100% Excel behavioral compatibility**—including edge cases, error handling, and performance characteristics.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  FORMULA INPUT                                                   │
│  "=IF(A1>0, VLOOKUP(B1, Data!A:C, 3, FALSE), "")"              │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  LEXER (Tokenization)                                           │
│  ├── Identify token types (operators, refs, functions, etc.)   │
│  ├── Handle locale-specific separators (, vs ;)                │
│  └── Preserve whitespace for reconstruction                     │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  PARSER (AST Construction)                                      │
│  ├── Recursive descent / Chevrotain-style LL(k)                │
│  ├── Operator precedence handling                               │
│  ├── Reference resolution (A1, R1C1, structured, named)        │
│  └── Validation and error location                              │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  AST (Abstract Syntax Tree)                                     │
│  ├── Normalized representation with relative addressing         │
│  ├── Function call nodes with argument lists                    │
│  ├── Binary/unary operation nodes                               │
│  └── Reference nodes (cell, range, named, structured)          │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  DEPENDENCY GRAPH                                               │
│  ├── Register formula dependencies                              │
│  ├── Range node optimization for large ranges                  │
│  ├── Volatile function tracking                                 │
│  └── Circular reference detection                               │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  EVALUATOR                                                      │
│  ├── Stack-based evaluation (RPN-style internally)             │
│  ├── Multi-threaded independent branch execution               │
│  ├── Array formula expansion and spilling                      │
│  └── Error propagation and handling                             │
└─────────────────────────────────────────────────────────────────┘
```

---

## Lexer Specification

### Token Types

| Token Type | Examples | Notes |
|------------|----------|-------|
| NUMBER | `123`, `3.14`, `1E10`, `.5` | Scientific notation, leading decimal |
| STRING | `"Hello"`, `""` | Double-quote escaped by doubling |
| BOOLEAN | `TRUE`, `FALSE` | Case insensitive |
| ERROR | `#VALUE!`, `#REF!`, `#N/A` | All Excel error types (classic + dynamic array/data errors like `#SPILL!`, `#CALC!`, `#GETTING_DATA`, `#FIELD!`, etc.) |
| CELL_REF | `A1`, `$A$1`, `A$1`, `$A1` | Absolute/relative markers |
| RANGE_REF | `A1:B10`, `A:A`, `1:1` | Full column/row references |
| SHEET_REF | `Sheet1!A1`, `Sheet1:Sheet3!A1`, `'Sheet Name'!A1`, `'Sheet 1:Sheet 3'!A1` | Includes 3D sheet spans; quoted for special chars |
| EXTERNAL_REF | `[Book.xlsx]Sheet!A1` | External workbook references |
| STRUCTURED_REF | `[@Column]`, `Table1[Column]` | Table references |
| NAMED_RANGE | `MyRange`, `_xlnm.Print_Area` | Named references |
| FUNCTION | `SUM`, `VLOOKUP`, `_xlfn.XLOOKUP` | Including prefixed new functions |
| OPERATOR | `+`, `-`, `*`, `/`, `^`, `&`, `=`, `<>`, `<`, `>`, `<=`, `>=` | |
| PAREN | `(`, `)` | |
| ARRAY_OPEN | `{` | Array constants |
| ARRAY_CLOSE | `}` | |
| ARRAY_ROW_SEP | `;` | Row separator in arrays |
| ARRAY_COL_SEP | `,` | Column separator (locale-dependent) |
| ARG_SEP | `,` or `;` | Locale-dependent argument separator |
| WHITESPACE | ` `, `\t` | Preserved for reconstruction |
| INTERSECT | ` ` (space between refs) | Implicit intersection operator |
| UNION | `,` (between refs) | Reference union |

### Locale Handling

| Locale | Decimal | Argument Sep | Array Col | Array Row |
|--------|---------|--------------|-----------|-----------|
| US/UK | `.` | `,` | `,` | `;` |
| Germany | `,` | `;` | `\` | `;` |
| France | `,` | `;` | `;` | `!` |

The lexer must detect locale from file metadata or user settings and tokenize accordingly.

### R1C1 Mode

When R1C1 mode is enabled:
- `R1C1` = absolute row 1, column 1
- `RC` = current cell
- `R[1]C[-1]` = relative offset (1 row down, 1 column left)
- `R1C[-1]` = mixed (absolute row, relative column)

---

## Parser Specification

### Grammar (Simplified EBNF)

```ebnf
formula     = "=" expression ;
expression  = term (("+"|"-"|"&") term)* ;
term        = factor (("*"|"/") factor)* ;
factor      = base ("^" base)* ;
base        = unary | primary ;
unary       = ("-"|"+") base ;
primary     = number | string | boolean | error | reference | function_call | "(" expression ")" | array_literal ;
function_call = function_name "(" [arg_list] ")" ;
arg_list    = expression ("," expression)* ;
  // Sheet prefixes may qualify any reference (cells, ranges, structured refs, names):
  //   Sheet1!A1
  //   Sheet1:Sheet3!A1   // 3D sheet span (workbook sheet order)
  //   [Book.xlsx]Sheet1!A1
  sheet_prefix = (sheet_name | sheet_name ":" sheet_name | "[" workbook_name "]" sheet_name) "!" ;
  reference   = [sheet_prefix] (cell_ref | range_ref | named_ref | structured_ref) ;
  array_literal = "{" array_row (";" array_row)* "}" ;
  array_row   = array_element ("," array_element)* ;
  array_element = number | string | boolean | error ;
```

### External workbook references (`ExternalValueProvider`)

External workbook references like `=[Book.xlsx]Sheet1!A1` are resolved through the host-provided
[`ExternalValueProvider`](../crates/formula-engine/src/engine.rs) trait.

The engine passes a **sheet key** string to `ExternalValueProvider::get(sheet, addr)`:

* **Internal sheet access** (same workbook): `sheet` is the plain worksheet name, e.g. `"Sheet1"`.
* **External workbook access**: `sheet` is a **path-qualified, workbook-prefixed** key:
  * **Canonical external sheet key:** `"[workbook]sheet"`
    * Example: `"[Book.xlsx]Sheet1"`
    * Example (path-qualified): `"[C:\\path\\Book.xlsx]Sheet1"`

#### External 3D sheet spans (workbook sheet order)

Excel 3D spans inside an external workbook (e.g. `Sheet1:Sheet3`) are represented by the engine as a
single **span key**:

* **External 3D span key format:** `"[workbook]Sheet1:Sheet3"`

This key is **not** looked up directly via `ExternalValueProvider::get`. Instead, during evaluation the
engine expands the span into per-sheet keys using workbook sheet order returned by:

* `ExternalValueProvider::sheet_order(workbook) -> Option<Vec<String>>`

Expansion rules:

* `workbook` is the raw string inside the brackets (e.g. `"Book.xlsx"` or `"C:\\path\\Book.xlsx"`).
* The returned sheet names must be **plain sheet names** (no `[workbook]` prefix).
* Endpoint matching (`Sheet1` / `Sheet3`) is **case-insensitive**.
* If `sheet_order(...)` returns `None` **or** either endpoint is missing from the returned order, the
  3D span evaluates to `#REF!`.
* Degenerate spans where start and end are the same sheet (case-insensitive) are canonicalized to a
  single-sheet key (e.g. `"[Book.xlsx]Sheet1"`), so `sheet_order` is not required.

Example:

```txt
Formula: =SUM([Book.xlsx]Sheet1:Sheet3!A1)

sheet_order("Book.xlsx") -> ["Sheet1", "Sheet2", "Sheet3", ...]

Expanded lookups via `get(sheet, addr)` (conceptually):
  get("[Book.xlsx]Sheet1", A1)
  get("[Book.xlsx]Sheet2", A1)
  get("[Book.xlsx]Sheet3", A1)
```

#### Path-qualified external workbook canonicalization

Excel allows quoting a full path + workbook + sheet, e.g.:

```txt
'C:\path\[Book.xlsx]Sheet1'!A1
```

The engine canonicalizes the workbook identifier by folding the path into the `[workbook]` portion of
the external sheet key:

```txt
'C:\path\[Book.xlsx]Sheet1'!A1  =>  sheet key "[C:\path\Book.xlsx]Sheet1"
```

#### Current limitations / behavior notes

* **Bytecode backend:** formulas that contain external workbook references currently do **not** compile
  to bytecode (they fall back to the AST evaluator).
* **External structured references:** structured refs cannot be workbook/sheet-qualified today
  (e.g. `[Book.xlsx]Table1[Col]` evaluates to `#REF!`).
* **Volatility / invalidation:** external workbook references are treated as **volatile** (they are
  reevaluated on every `Engine::recalculate()` pass). There is not yet a fine-grained “external link
  invalidation” mechanism—hosts should call `recalculate()` when external values may have changed.
* **INDIRECT + external 3D spans:** `INDIRECT("[Book.xlsx]Sheet1:Sheet3!A1")` currently evaluates to
  `#REF!` (span expansion is only supported for direct references).

#### Minimal provider sketch (including `sheet_order`)

```rust
use formula_engine::{ExternalValueProvider, Value};
use formula_engine::eval::CellAddr;
use std::collections::HashMap;

struct Provider {
    // Keyed by the engine's canonical sheet key + cell address.
    cells: HashMap<(String, CellAddr), Value>,
    // Keyed by workbook string inside `[...]`, e.g. "Book.xlsx" or "C:\\path\\Book.xlsx".
    orders: HashMap<String, Vec<String>>,
}

impl ExternalValueProvider for Provider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        self.cells.get(&(sheet.to_string(), addr)).cloned()
    }

    fn sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        self.orders.get(workbook).cloned()
    }
}
```

### Operator Precedence

| Precedence | Operators | Associativity |
|------------|-----------|---------------|
| 1 (highest) | `:` (range), ` ` (intersect), `,` (union) | Left |
| 2 | `-` (negation) | Right |
| 3 | `%` (percent) | Left |
| 4 | `^` (exponent) | Right |
| 5 | `*`, `/` | Left |
| 6 | `+`, `-` | Left |
| 7 | `&` (concatenate) | Left |
| 8 | `=`, `<>`, `<`, `>`, `<=`, `>=` | Left |

### AST Node Types

```typescript
type ASTNode = 
  | NumberNode
  | StringNode
  | BooleanNode
  | ErrorNode
  | CellRefNode
  | RangeRefNode
  | NamedRefNode
  | StructuredRefNode
  | BinaryOpNode
  | UnaryOpNode
  | FunctionCallNode
  | ArrayLiteralNode;

interface CellRefNode {
  type: "cell_ref";
  row: number;        // 0-indexed
  col: number;        // 0-indexed  
  rowAbsolute: boolean;
  colAbsolute: boolean;
  // undefined = current sheet.
  // When present, may be a single sheet name or a 3D span (`Sheet1:Sheet3!A1`).
  sheet?: string | { start: string; end: string };
  workbook?: string;  // undefined = current workbook
}

interface FunctionCallNode {
  type: "function_call";
  name: string;       // normalized uppercase
  originalName: string; // preserve case for display
  args: ASTNode[];
  prefixed: boolean;  // true if _xlfn. prefix present
}
```

### Relative Addressing for AST Sharing

Formulas like `=A1+B1` in C1 and `=A2+B2` in C2 share the same normalized AST:

```typescript
// Both formulas compile to:
{
  type: "binary_op",
  op: "+",
  left: { type: "cell_ref", rowOffset: 0, colOffset: -2, rowAbsolute: false, colAbsolute: false },
  right: { type: "cell_ref", rowOffset: 0, colOffset: -1, rowAbsolute: false, colAbsolute: false }
}
```

This reduces memory usage significantly when formulas are dragged/filled.

---

## Dependency Graph

### Structure

```typescript
interface DependencyGraph {
  // Forward edges: cell -> cells that depend on it
  dependents: Map<CellId, Set<CellId>>;
  
  // Backward edges: cell -> cells it depends on
  precedents: Map<CellId, Set<CellId>>;
  
  // Range nodes for optimization
  rangeNodes: Map<RangeId, RangeNode>;
  
  // Volatile cells (must recalc always)
  volatileCells: Set<CellId>;
  
  // Calculation chain (topologically sorted)
  calcChain: CellId[];
  
  // Dirty cells needing recalculation
  dirtySet: Set<CellId>;
}

interface RangeNode {
  range: Range;
  dependents: Set<CellId>;  // Cells with formulas using this range
}
```

### Range Node Optimization

Without optimization, `SUM(A1:A1000)` creates 1000 edges. With range nodes:

```
RangeNode(A1:A1000) ← Cell(B1)
```

When cell A500 changes:
1. Find all RangeNodes containing A500
2. Mark their dependents dirty
3. No need to check 1000 individual edges

For cumulative patterns like:
```
A2: =SUM(A$1:A1)
A3: =SUM(A$1:A2)
A4: =SUM(A$1:A3)
```

Decompose ranges: `SUM(A$1:A3)` = `RangeNode(A1:A2) + Cell(A3)`

### Topological Sort (Kahn's Algorithm)

```typescript
function buildCalcChain(graph: DependencyGraph): CellId[] {
  const inDegree = new Map<CellId, number>();
  const queue: CellId[] = [];
  const result: CellId[] = [];
  
  // Initialize in-degrees
  for (const cell of graph.allCells()) {
    const degree = graph.precedents.get(cell)?.size ?? 0;
    inDegree.set(cell, degree);
    if (degree === 0) queue.push(cell);
  }
  
  // Process queue
  while (queue.length > 0) {
    const cell = queue.shift()!;
    result.push(cell);
    
    for (const dependent of graph.dependents.get(cell) ?? []) {
      const newDegree = inDegree.get(dependent)! - 1;
      inDegree.set(dependent, newDegree);
      if (newDegree === 0) queue.push(dependent);
    }
  }
  
  // Check for cycles
  if (result.length !== graph.cellCount()) {
    throw new CircularReferenceError(findCycle(graph));
  }
  
  return result;
}
```

### Dirty Marking

```typescript
function markDirty(cell: CellId, graph: DependencyGraph): void {
  if (graph.dirtySet.has(cell)) return;  // Already dirty
  
  graph.dirtySet.add(cell);
  
  // Mark all dependents dirty (transitive)
  const dependents = graph.dependents.get(cell);
  if (dependents) {
    for (const dep of dependents) {
      markDirty(dep, graph);
    }
  }
  
  // Mark dependents of containing range nodes
  for (const [rangeId, rangeNode] of graph.rangeNodes) {
    if (rangeNode.range.contains(cell)) {
      for (const dep of rangeNode.dependents) {
        markDirty(dep, graph);
      }
    }
  }
}
```

### Volatile Functions

These functions must recalculate on every workbook recalculation:

| Function | Behavior |
|----------|----------|
| `NOW()` | Returns current date and time |
| `TODAY()` | Returns current date |
| `RAND()` | Returns random number |
| `RANDBETWEEN()` | Returns random integer in range |
| `OFFSET()` | Returns reference offset from base (volatile because range can change) |
| `INDIRECT()` | Returns reference from string (volatile because target unknown at parse time) |
| `INFO()` | Returns workbook/system information (see [`docs/21-info-cell-metadata.md`](./21-info-cell-metadata.md) for host-provided metadata requirements) |
| `CELL()` | Returns cell/workbook information (see [`docs/21-info-cell-metadata.md`](./21-info-cell-metadata.md) for host-provided metadata requirements) |

Volatility propagates: if A1 contains `=NOW()`, any cell depending on A1 is also effectively volatile.

---

## Calculation Engine

### Multi-Threaded Recalculation

Excel uses Multi-Threaded Recalculation (MTR) for independent branches:

```
        A1 (dirty)
       /  \
      B1   B2     <- Can calculate in parallel
     /      \
    C1      C2    <- Must wait for B1, B2
     \      /
        D1        <- Must wait for C1, C2
```

**Thread-safe functions**: Most built-in functions (including many volatile functions like `NOW()`, `TODAY()`, `RAND()`)
**Non-thread-safe**: VBA UDFs, functions that depend on external system state (e.g. `RTD()`)

Note: Volatility and thread-safety are independent: volatility affects *when* a formula must be recalculated, while
thread-safety affects whether it can be evaluated in parallel.

Implementation approach:
1. Partition calc chain into independent subgraphs
2. Assign subgraphs to worker threads
3. Synchronize at merge points

### Evaluation Strategy

Stack-based evaluation (internally RPN):

```typescript
function evaluate(ast: ASTNode, context: EvalContext): CellValue {
  const stack: CellValue[] = [];
  
  // Convert AST to RPN instruction sequence
  const instructions = astToRPN(ast);
  
  for (const inst of instructions) {
    switch (inst.type) {
      case "push_number":
        stack.push(inst.value);
        break;
        
      case "push_ref":
        stack.push(resolveReference(inst.ref, context));
        break;
        
      case "binary_op":
        const right = stack.pop()!;
        const left = stack.pop()!;
        stack.push(applyBinaryOp(inst.op, left, right));
        break;
        
      case "call_function":
        const args = stack.splice(-inst.argCount);
        stack.push(callFunction(inst.name, args, context));
        break;
    }
  }
  
  return stack[0];
}
```

### Dynamic Array Spilling

When a formula returns an array:

```typescript
function handleArrayResult(
  origin: CellId, 
  result: CellValue[][], 
  sheet: Sheet
): void {
  const rows = result.length;
  const cols = result[0].length;
  
  // Check for spill blocking
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (r === 0 && c === 0) continue; // Origin cell is fine
      
      const targetCell = { row: origin.row + r, col: origin.col + c };
      if (!sheet.isEmpty(targetCell)) {
        // Spill blocked - set #SPILL! error on origin
        sheet.setValue(origin, ErrorValue.SPILL);
        return;
      }
    }
  }
  
  // Write spilled values
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const targetCell = { row: origin.row + r, col: origin.col + c };
      sheet.setSpilledValue(targetCell, result[r][c], origin);
    }
  }
}
```

### Implicit Intersection (@) Operator

In non-array context, multi-cell references implicitly intersect with the formula's row/column:

```
// In cell C5:
=A1:A10 * 2    // Implicitly becomes =@A1:A10 * 2 = A5 * 2
```

With dynamic arrays, explicit `@` is required for this behavior.

---

## Function Library

### Implementation Guidelines

Each function must specify:

```typescript
interface FunctionSpec {
  name: string;
  minArgs: number;
  maxArgs: number;
  returnType: ValueType | "any";
  argTypes: ArgSpec[];
  isVolatile: boolean;
  isThreadSafe: boolean;
  supportsArrays: boolean;
  implementation: (...args: CellValue[]) => CellValue;
}

interface ArgSpec {
  name: string;
  type: ValueType | ValueType[];
  optional: boolean;
  repeating: boolean;  // For varargs
  description: string;
}
```

### Function Categories and Counts

| Category | Count | Examples |
|----------|-------|----------|
| Math & Trig | 60+ | SUM, AVERAGE, ROUND, SIN, LOG |
| Statistical | 80+ | STDEV, CORREL, LINEST, NORM.DIST |
| Lookup & Reference | 20+ | VLOOKUP, XLOOKUP, INDEX, MATCH, INDIRECT |
| Text | 40+ | CONCATENATE, LEFT, FIND, SUBSTITUTE |
| Logical | 10+ | IF, AND, OR, IFS, SWITCH, XOR |
| Date & Time | 25+ | DATE, DATEVALUE, NETWORKDAYS, WORKDAY |
| Financial | 50+ | NPV, IRR, PMT, FV, XNPV, XIRR |
| Information | 20+ | ISBLANK, ISERROR, TYPE, CELL |
| Engineering | 40+ | CONVERT, COMPLEX, IMSUM, BIN2DEC |
| Database | 12 | DSUM, DCOUNT, DGET, DAVERAGE |
| Cube | 7 | CUBEVALUE, CUBEMEMBER, CUBERANKEDMEMBER |
| Web | 3 | WEBSERVICE, ENCODEURL, FILTERXML |
| Dynamic Array | 8 | FILTER, SORT, SORTBY, UNIQUE, SEQUENCE |
| Lambda | 9 | LAMBDA, LET, ISOMITTED, MAP, REDUCE, SCAN, MAKEARRAY, BYROW, BYCOL |

**Total: ~500 functions**

The parser also supports **invoking lambdas** using both postfix call syntax (`expr(args)`, e.g.
`LAMBDA(x, x+1)(2)`) and calling a name bound to a lambda (`LET(f, LAMBDA(...), f(2))`). See
[docs/19-lambda-functions.md](./19-lambda-functions.md) for the full semantics and error behavior.

### Critical Edge Cases

**SUM with mixed types:**
```
SUM("5", TRUE, 3) = 9  // "5" → 5, TRUE → 1
SUM(A1:A3) where A1="5" = 3  // Text in ranges ignored!
```

**VLOOKUP approximate match:**
```
// Data must be sorted ascending for approximate match
// Returns largest value ≤ lookup_value
// If lookup_value < smallest, returns #N/A
```

**DATE out of range:**
```
DATE(1899, 12, 31) = #NUM!  // Before Excel epoch
DATE(10000, 1, 1) = #NUM!   // After max date
DATE(2024, 0, 15) = DATE(2023, 12, 15)  // Month 0 = previous December
DATE(2024, 13, 1) = DATE(2025, 1, 1)    // Month 13 = next January
```

**Division by zero:**
```
=1/0   → #DIV/0!
=0/0   → #DIV/0!
=MOD(5, 0) → #DIV/0!
```

---

## Error Handling

### Error Types

| Error | Code | Cause |
|-------|------|-------|
| `#NULL!` | 1 | Invalid range intersection |
| `#DIV/0!` | 2 | Division by zero |
| `#VALUE!` | 3 | Wrong argument type |
| `#REF!` | 4 | Invalid cell reference |
| `#NAME?` | 5 | Unrecognized function/name |
| `#NUM!` | 6 | Invalid numeric value |
| `#N/A` | 7 | Value not available |
| `#GETTING_DATA` | 8 | External data loading |
| `#SPILL!` | 9 | Spill range blocked |
| `#CALC!` | 10 | Calculation engine error |
| `#FIELD!` | 11 | Field not found in record |
| `#CONNECT!` | 12 | Connection error |
| `#BLOCKED!` | 13 | Feature blocked |
| `#UNKNOWN!` | 14 | Unknown error |

### Error Propagation

Errors propagate through most operations:
```
= #VALUE! + 5  → #VALUE!
= IF(#N/A, 1, 2)  → #N/A
```

Exceptions:
```
= IFERROR(#VALUE!, "Error") → "Error"
= ISERROR(#VALUE!) → TRUE
= IF(TRUE, 1, #VALUE!) → 1  // Short-circuit
```

---

## Performance Optimizations

### Compiled Formula Cache

For frequently recalculated formulas, compile AST to optimized bytecode or native code via JIT.

### SIMD Vectorization

For bulk operations on ranges:
```rust
// Instead of:
for i in 0..len {
    result[i] = a[i] + b[i];
}

// Use SIMD:
use std::simd::*;
for chunk in data.chunks_exact(8) {
    let va: f64x8 = f64x8::from_slice(a_chunk);
    let vb: f64x8 = f64x8::from_slice(b_chunk);
    (va + vb).copy_to_slice(result_chunk);
}
```

### Lazy Evaluation

Don't evaluate branches that won't be used:
```
=IF(A1=0, 0, B1/A1)
// Don't evaluate B1/A1 if A1=0
```

### Incremental Recalculation

Only recalculate dirty cells, not entire workbook:
```typescript
function recalculate(graph: DependencyGraph): void {
  const dirtyCells = Array.from(graph.dirtySet)
    .sort((a, b) => graph.calcChainIndex(a) - graph.calcChainIndex(b));
  
  for (const cell of dirtyCells) {
    evaluateCell(cell);
    graph.dirtySet.delete(cell);
  }
}
```

---

## Testing Strategy

### Compatibility Test Suite

1. **Function behavior tests**: Test each function against Excel with various inputs
2. **Edge case coverage**: Document and test all known edge cases
3. **Round-trip formula tests**: Parse → serialize → parse should be identical
4. **Performance benchmarks**: Compare calculation speed with Excel

### Test Case Structure

```typescript
interface FormulaTestCase {
  description: string;
  formula: string;
  inputs: Record<string, CellValue>;  // A1 -> value
  expected: CellValue | ErrorValue;
  excelVersion?: string;  // If behavior differs by version
  notes?: string;
}

const testCases: FormulaTestCase[] = [
  {
    description: "SUM ignores text in ranges",
    formula: "=SUM(A1:A3)",
    inputs: { A1: "text", A2: 5, A3: 10 },
    expected: 15
  },
  {
    description: "VLOOKUP exact match not found",
    formula: "=VLOOKUP(5, A1:B3, 2, FALSE)",
    inputs: { A1: 1, B1: "a", A2: 2, B2: "b", A3: 3, B3: "c" },
    expected: ErrorValue.NA
  }
];
```

---

## Open Questions and Future Work

1. **Custom function extensibility**: How do user-defined functions integrate with the dependency graph?
2. **External data functions**: How to handle async data fetching in synchronous calc engine?
3. **Circular reference handling**: Support iterative calculation mode like Excel?
4. **GPU acceleration**: For very large datasets, offload to GPU compute shaders?
5. **Distributed calculation**: For enterprise, distribute calc across multiple machines?
