# Workstream D: AI Integration

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Weave AI into the fabric of Formula—not as a chatbot sidebar, but as a co-pilot that enhances every interaction. Following Cursor's proven paradigm: an autonomy slider from passive assistance to fully autonomous agents.

**The goal:** AI that supercharges experts while enabling novices.

---

## ⚠️ CRITICAL: This is a Cursor Product

> **All AI goes through Cursor servers.**
>
> - **No local models** — all inference via Cursor backend
> - **No user API keys** — Cursor manages authentication
> - **No provider selection** — Cursor controls model routing
> - **Cursor controls harness and prompts** — consistent experience across all users

This is non-negotiable. Do not implement local model support, API key configuration, or provider selection UI. All AI features call Cursor's backend.

---

## Scope

### Your Code

| Location | Purpose |
|----------|---------|
| `packages/ai-completion` | Tab completion engine |
| `packages/ai-context` | Context extraction and sampling |
| `packages/ai-tools` | Tool definitions for AI agents |
| `packages/ai-rag` | RAG over cells (embeddings, retrieval) |
| `packages/ai-audit` | AI action audit trail |
| `packages/llm` | LLM client abstraction (Cursor backend) |
| `apps/desktop/src/ai/` | AI integration in desktop app |

### Your Documentation

- **Primary:** [`docs/05-ai-integration.md`](../docs/05-ai-integration.md) — AI modes, context management, tool calling
- **Pivot tool semantics:** [`docs/adr/ADR-0005-pivot-tables-ownership-and-data-flow.md`](../docs/adr/ADR-0005-pivot-tables-ownership-and-data-flow.md) — PivotTables ownership + data flow (avoid duplicating pivot engines inside AI tooling)

---

## The Autonomy Spectrum

```
LOW AUTONOMY                                              HIGH AUTONOMY
◄──────────────────────────────────────────────────────────────────────►

Tab Complete    Inline Edit    Chat Panel    Composer    Full Agent
     │              │              │            │            │
User types,    User selects   User asks    User        User sets
AI suggests    range, asks    questions,   describes   goal, AI
completions    for transform  AI answers   multi-step  executes
                                           operation   autonomously

Latency:       Latency:       Latency:     Latency:    Latency:
<100ms         <2s            <5s          <30s        minutes
```

---

## Key Requirements

### Mode 1: Tab Completion (<100ms)

- Trigger: User typing in formula bar
- Suggest formulas, values, completions based on context
- Aggressive caching for speed

### Mode 2: Inline Edit (Cmd/Ctrl+K, <2s)

- Trigger: User selects range, presses Cmd/Ctrl+K
- Natural language → data transformation
- Preview before apply

### Mode 3: Chat Panel (<5s)

- Toggle AI sidebar (Cmd+I on macOS / Ctrl+Shift+A on Windows/Linux)
- Ask questions about data
- Request analysis, explain formulas

### Mode 4: Cell Functions

- `=AI("prompt", range)` — AI as a formula
- Recalculates with data changes
- Results cached until invalidated

### Mode 5: Agent Mode (minutes)

- Full-height separate view
- Multi-step autonomous workflows
- Data gathering, cleaning, analysis
- Tool calling with approval gates

---

## Context Management

A 1M-row spreadsheet could exceed 100M tokens. Required strategies:

1. **Schema-First:** Send structure (headers, types, samples), not raw data
2. **Selective Sampling:** Random/stratified sample for statistical questions
3. **RAG Over Cells:** Retrieve relevant context via deterministic, offline hash embeddings (no user API keys, no local model setup)
4. **Hierarchical Summarization:** Pre-computed aggregates at sheet/region level
5. **Query-Aware Loading:** Dynamically load only ranges relevant to current query

### Context Budget (128K token model)

```
System prompt:          ~2K tokens
Schema:                 ~1K tokens
Sample data (100 rows): ~5K tokens
Conversation history:   ~10K tokens
Retrieved context:      ~20K tokens
Reserved for output:    ~20K tokens
───────────────────────────────────
Available buffer:       ~70K tokens
```

---

## Tool Calling

AI agents use tools to interact with the spreadsheet:

```typescript
const tools = [
  {
    name: "read_range",
    description: "Read cell values from a range",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range in A1 notation (e.g. 'Sheet1!A1:C10')" },
        include_formulas: { type: "boolean", default: false }
      },
      required: ["range"]
    }
  },
  {
    name: "write_cell",
    description: "Write a value or formula to a cell",
    parameters: {
      type: "object",
      properties: {
        cell: { type: "string", description: "Cell reference (e.g. 'Sheet1!A1')" },
        value: {
          description: "Scalar value or formula string",
          anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }, { type: "null" }]
        },
        is_formula: { type: "boolean", description: "Treat value as formula even if it does not start with '='." }
      },
      required: ["cell", "value"]
    }
  },
  {
    name: "create_pivot_table",
    description: "Create a pivot table from a source range.",
    parameters: {
      type: "object",
      properties: {
        source_range: { type: "string" },
        rows: { type: "array", items: { type: "string" } },
        columns: { type: "array", items: { type: "string" } },
        values: {
          type: "array",
          items: {
            type: "object",
            properties: {
              field: { type: "string" },
              aggregation: {
                type: "string",
                enum: [
                  "sum",
                  "count",
                  "average",
                  "min",
                  "max",
                  "product",
                  "countNumbers",
                  "stdDev",
                  "stdDevP",
                  "var",
                  "varP"
                ]
              }
            },
            required: ["field", "aggregation"]
          }
        },
        destination: { type: "string" }
      },
      required: ["source_range", "rows", "values"]
    }
  },
  {
    name: "detect_anomalies",
    description: "Find outliers in data",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string" },
        method: { type: "string", enum: ["zscore", "iqr", "isolation_forest"], default: "zscore" },
        threshold: { type: "number" }
      },
      required: ["range"]
    }
  }
];
```

Tool calling should be implemented using the provider-agnostic `ToolDefinition` and `LLMMessage` models in `packages/llm` (spreadsheet tool schemas live in `packages/ai-tools`). Cursor's backend is responsible for translating these structures into whatever underlying model API it routes to.

---

## Safety and Verification

1. **Computation over hallucination:** Use formulas to compute answers, don't guess
2. **Source attribution:** External data includes citation metadata
3. **Preview before apply:** Major changes show diff for user approval
4. **Confidence indicators:** Flag uncertain outputs
5. **Audit trail:** Log all AI actions for review

---

## Build & Run

```bash
# Install dependencies
pnpm install

# Run tests
pnpm test

# Run AI-specific tests (skip WASM build; set `FORMULA_SKIP_WASM_BUILD=1` or `true`)
FORMULA_SKIP_WASM_BUILD=1 pnpm vitest run packages/ai-*/src/*.test.ts
FORMULA_SKIP_WASM_BUILD=1 pnpm vitest run packages/llm/src/*.test.ts
```

---

## Coordination Points

- **UI Team:** AI sidebar display, inline edit UI, suggestion rendering
- **Core Engine Team:** Context extraction (cell values, types, formulas)
- **Collaboration Team:** AI actions in collaborative context

---

## Anti-Patterns to Avoid

- ❌ Local model support
- ❌ API key configuration UI
- ❌ Provider selection dropdowns
- ❌ Temperature/token limit controls for users
- ❌ "AI slop" design (giant icons, gradient backgrounds, chat-bot aesthetics)

---

## Reference

- Cursor's AI implementation patterns
- Tool calling: JSON-schema tool definitions (Cursor backend handles provider-specific encoding)
- Context: RAG best practices
