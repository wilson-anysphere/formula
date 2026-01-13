# Workstream A: Core Engine (Rust)

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Build the computational heart of Formula: the formula parser, dependency graph, calculation engine, function library, and data model. This is Rust code compiled to WASM for near-native performance in the browser/Tauri context.

**The goal:** 100% Excel formula behavioral compatibility—including edge cases, error handling, and performance characteristics.

---

## Scope

### Your Crates

| Crate | Purpose |
|-------|---------|
| `crates/formula-engine` | Formula parser, evaluator, dependency graph, 500+ functions |
| `crates/formula-model` | Workbook/sheet/cell data structures |
| `crates/formula-columnar` | VertiPaq-style columnar compression for large datasets |
| `crates/formula-dax` | DAX-like calculated columns and measures |
| `crates/formula-wasm` | WASM bindings exposing engine to JavaScript |
| `crates/formula-format` | Number formatting, date formatting |

### Your Documentation

- **Primary:** [`docs/01-formula-engine.md`](../docs/01-formula-engine.md) — formula parsing, evaluation, dependency tracking
- **Storage:** [`docs/04-data-model-storage.md`](../docs/04-data-model-storage.md) — cell storage, columnar compression
- **Pivots:** [`docs/adr/ADR-0005-pivot-tables-ownership-and-data-flow.md`](../docs/adr/ADR-0005-pivot-tables-ownership-and-data-flow.md) — PivotTables ownership boundaries + data flow (worksheet pivots vs Data Model pivots)

---

## Key Requirements

### Formula Engine

1. **Parse all Excel formula syntax:**
   - A1, R1C1, structured references, dynamic arrays
   - Locale-specific separators (`,` vs `;`)
   - External workbook references `[Book.xlsx]Sheet!A1`
   - 3D sheet references `Sheet1:Sheet3!A1`

2. **Implement 500+ functions with identical Excel behavior:**
   - Including edge cases and error propagation
   - `_xlfn.` prefixed newer functions (XLOOKUP, FILTER, etc.)
   - Volatile functions (NOW, RAND, OFFSET, INDIRECT)

3. **Dependency graph:**
   - Incremental dirty marking
   - Range node optimization for large ranges
   - Circular reference detection
   - 65,536 dependency threshold before full recalc mode

4. **Multi-threaded calculation:**
   - Parallel evaluation across independent branches
   - SIMD optimization where beneficial
   - Rayon for parallelism

5. **Dynamic array spilling:**
   - `#SPILL!` error handling
   - Implicit intersection (`@`) operator

### Data Model

1. **Sparse HashMap storage** — most spreadsheets are sparse
2. **Columnar compression** — VertiPaq-style for large datasets
3. **No arbitrary row limits** — scale to 100M+ rows via streaming
4. **Rich data types** — images, JSON, attachments in cells

### Critical Specifications

```
Formula display limit:     8,192 characters (16,384 bytes tokenized)
Nested function levels:    64
Arguments per function:    255
Numeric precision:         15 significant digits (IEEE 754 double)
Dependency threshold:      65,536 before full recalc mode
```

---

## Build Commands

**ALWAYS use the wrapper script for cargo commands:**

```bash
# CORRECT:
bash scripts/cargo_agent.sh build --release -p formula-engine
bash scripts/cargo_agent.sh test -p formula-engine
bash scripts/cargo_agent.sh check -p formula-wasm
bash scripts/cargo_agent.sh bench --bench perf_regressions

# WRONG - can exhaust RAM:
cargo build
cargo test
```

**Build WASM:**

```bash
pnpm build:wasm
```

---

## Performance Targets

| Metric | Target |
|--------|--------|
| Recalculation | <100ms for 100K cell dependency chain |
| Formula parse | <1ms for typical formula |
| Memory | <500MB for 100MB xlsx loaded |

---

## Coordination Points

- **UI Team:** Cell rendering, formula bar, error display — you provide the API they consume
- **File I/O Team:** Serialization format, formula text preservation, cached values
- **AI Team:** Context extraction for AI features (schema, types, samples)

---

## Testing

```bash
# Run engine tests
bash scripts/cargo_agent.sh test -p formula-engine

# Run specific test
bash scripts/cargo_agent.sh test -p formula-engine -- test_name

# Run with verbose output
bash scripts/cargo_agent.sh test -p formula-engine -- --nocapture
```

**Test categories:**
- Function behavior tests (match Excel exactly)
- Parser edge cases
- Dependency graph correctness
- Performance regression tests

---

## Reference

- Excel function reference: https://support.microsoft.com/en-us/excel
- ECMA-376 formula grammar (Part 1, §18.17)
- Function catalog: `shared/functionCatalog.json`
