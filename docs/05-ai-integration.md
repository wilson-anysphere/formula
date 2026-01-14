# AI Integration

## Overview

AI integration is not a feature bolted on—it's woven into the fabric of the application. Following Cursor's proven paradigm, we implement an **autonomy slider** from passive assistance through active co-piloting to fully autonomous agents.

> **⚠️ This is a Cursor product. All AI goes through Cursor servers.**
>
> - **No local models** - all inference via Cursor backend
> - **No user API keys** - Cursor manages authentication
> - **No provider selection** - Cursor controls model routing
> - **Cursor controls harness and prompts** - consistent experience across all users

---

## Integration Modes

### The Autonomy Spectrum

```
┌────────────────────────────────────────────────────────────────────────────────┐
│  AUTONOMY SPECTRUM                                                             │
├────────────────────────────────────────────────────────────────────────────────┤
│                                                                                │
│  LOW AUTONOMY                                             HIGH AUTONOMY        │
│  ◄─────────────────────────────────────────────────────────────────────────►   │
│                                                                                │
│  Tab Complete    Inline Edit    Chat Panel    AI Cell Fn    Agent Mode         │
│       │              │              │            │             │               │
│  User types,    User selects   User asks    User writes   User sets goal,      │
│  AI suggests    range + prompt questions,   =AI(...)     AI uses tools         │
│  completions    AI applies     AI answers   async eval   iteratively           │
│                via tools                               (approval gated)       │
│                                                                                │
│  Latency:       Latency:       Latency:     Latency:    Latency:            │
│  <100ms         <2s            <5s          async       minutes             │
│                                                                                │
│  (All AI via Cursor servers - no local models, no user API keys, no provider selection) │
│                                                                                │
└────────────────────────────────────────────────────────────────────────────────┘
```

### Mode 1: Tab Completion

**Trigger:** User typing in formula bar or cell
**Latency requirement:** <100ms
**Backend:** Primarily local heuristics; optional Cursor server completion for formula drafts (hard timeout + caching)

**Code entrypoints:**
- Core completion engine: [`packages/ai-completion/src/tabCompletionEngine.js`](../packages/ai-completion/src/tabCompletionEngine.js)
- Desktop formula-bar glue (wires suggestions + previews into the UI): [`apps/desktop/src/ai/completion/formulaBarTabCompletion.ts`](../apps/desktop/src/ai/completion/formulaBarTabCompletion.ts)
 
Implementation details and extension guidance: see [AI Tab Completion](ai-tab-completion.md).

```typescript
interface TabCompletion {
  trigger: "typing" | "tab_key";
  context: {
    currentInput: string;
    cursorPosition: number;
    cellRef: CellRef;
    surroundingCells: CellContext;
  };
  suggestions: Suggestion[];
}

interface Suggestion {
  text: string;
  displayText: string;
  type: "formula" | "value" | "function_arg" | "range";
  confidence: number;
  preview?: CellValue;  // Show what it would calculate to
}
```

**Example scenarios:**

```
User types: =SUM(A
Suggestions:
  - =SUM(A1:A10)      [Completes range based on data]
  - =SUM(A:A)         [Entire column]
  - =SUM(Amount)      [Table column]

User types: =VLO
Suggestions:
  - =VLOOKUP(          [Standard completion]
  - =XLOOKUP(          [Suggest modern alternative]

User types: Tot
Suggestions:
  - Total              [Based on nearby text patterns]
```

> **Note (current behavior):** the desktop formula bar only renders tab completion as **pure insertion** ghost text at
> the caret (it cannot rewrite characters *before* the cursor). Because of this, tab completion does not currently
> “rewrite” non-formula text into formulas—formula suggestions (including Cursor backend calls) only run when the user is
> already editing a formula (input starts with `=`).

**Implementation notes (actual):**
- For a deep-dive (algorithms, schema integration, tests), see [AI Tab Completion](ai-tab-completion.md).
- `TabCompletionEngine.getSuggestions()` (`packages/ai-completion/...`) parses the partial draft and merges three sources (in parallel), caching the base results:
  - rule-based suggestions (functions, ranges, argument hints) **for formulas**
  - pattern suggestions (nearby repeated values) **for non-formula input** (`type: "value"`)
  - optional Cursor backend completion via `CursorTabCompletionClient` **for formulas** (strict timeout for UI responsiveness)
    - `CursorTabCompletionClient` can optionally use `getAuthHeaders()` for Cursor-managed auth headers when cookie auth is unavailable, and supports `signal?: AbortSignal` for cancellation.
- Desktop attaches formula previews by evaluating the suggested formula locally (see `createPreviewEvaluator` in
  [`apps/desktop/src/ai/completion/formulaBarTabCompletion.ts`](../apps/desktop/src/ai/completion/formulaBarTabCompletion.ts), which uses
  [`apps/desktop/src/spreadsheet/evaluateFormula.ts`](../apps/desktop/src/spreadsheet/evaluateFormula.ts)).

#### Locale-aware partial parsing (WASM engine)

Tab completion needs a fast “best-effort” parse of *in-progress* formulas to know whether the user is typing a function
name, which argument they are in, and what kinds of suggestions are appropriate. Formula exposes this as a locale-aware
entrypoint in the WASM engine:

```typescript
const partial = await engine.parseFormulaPartial(formula, cursor, { localeId });
```

**Code entrypoints:**
- Engine client API: [`packages/engine/src/client.ts`](../packages/engine/src/client.ts) (`parseFormulaPartial`, `setLocale`)
- Types for parse options/results: [`packages/engine/src/protocol.ts`](../packages/engine/src/protocol.ts) (`FormulaParseOptions`, `FormulaPartialParseResult`)
- WASM implementation: [`crates/formula-wasm/src/lib.rs`](../crates/formula-wasm/src/lib.rs)

- **Why it matters:** Excel-style formulas are locale-sensitive.
  - **Argument/list separators** vary (e.g. `,` vs `;`), which changes how `argIndex` is computed.
  - **Localized function names** may be accepted/shown (e.g. `SUM` vs a localized equivalent), which impacts function-name
    completion and signature help. `engine.parseFormulaPartial(...)` returns a **canonicalized** (English) function name in
    its returned context even when parsing localized formulas (and strips the `_xlfn.` prefix).
- **How the UI should use it:** call `engine.parseFormulaPartial(...)` on each keystroke (or on a small debounce) and treat
  the returned context/error as hints for completion, not as a hard validation pass.
- **How it composes with `packages/ai-completion`:** the rule-based tab completion engine in `packages/ai-completion`
  remains responsible for generating/ranking suggestions; its `TabCompletionEngine` supports injecting a
  `parsePartialFormula` implementation (sync or async). The UI can provide an adapter that delegates to
  `engine.parseFormulaPartial(formula, cursor, { localeId })` (with a fallback to the existing lightweight JS parser when
  WASM is unavailable) so completion logic stays consistent with the engine and the current locale.
- **Localized function-name completion:** the UI may also inject a locale-aware `FunctionRegistry` that registers localized
  aliases (e.g. de-DE `SUMME`) alongside canonical names, so function-name completion suggests localized spellings by
  default. Desktop does this in `apps/desktop/src/ai/completion/parsePartialFormula.ts` (`createLocaleAwareFunctionRegistry()`).
  `@formula/ai-completion` supports an optional `FunctionSpec.completionBoost` to bias name-completion ranking when needed.

### Mode 2: Inline Edit (Cmd/Ctrl+K)

**Trigger:** User selects range and presses Cmd/Ctrl+K
**Latency requirement:** <2s for small operations
**Backend:** Cursor servers

**Code entrypoints (desktop):**
- Inline edit UI + orchestration: [`apps/desktop/src/ai/inline-edit/`](../apps/desktop/src/ai/inline-edit/)
  - Controller entrypoint: [`inlineEditController.ts`](../apps/desktop/src/ai/inline-edit/inlineEditController.ts)
  - Overlay UI: [`inlineEditOverlay.ts`](../apps/desktop/src/ai/inline-edit/inlineEditOverlay.ts)

```typescript
interface InlineEditRequest {
  selection: Range;
  prompt: string;
  selectionContent: CellValue[][];
  surroundingContext: SheetContext;
}

interface InlineEditResponse {
  type: "transform" | "generate" | "format";
  changes: CellChange[];
  explanation?: string;
  confidence: number;
}
```

**Example scenarios:**

```
Selection: A1:A10 (list of names)
Prompt: "extract first names"
Result: New column with first names

Selection: B2:B100 (dates in various formats)
Prompt: "standardize to YYYY-MM-DD"
Result: All dates reformatted

Selection: D1:D50 (product descriptions)
Prompt: "extract price"
Result: New column with prices extracted

Selection: Empty range E1:E10
Prompt: "months of the year"
Result: January, February, March, ...
```

**Implementation notes (actual):**
- Inline edit runs a **tool-calling loop** against spreadsheet tools (no LLM codegen/sandboxing):
  - Build bounded workbook context for the selection (schema-first + optional RAG): [`apps/desktop/src/ai/context/WorkbookContextBuilder.ts`](../apps/desktop/src/ai/context/WorkbookContextBuilder.ts)
  - Execute the tool loop + audit: [`packages/ai-tools/src/llm/audited-run.ts`](../packages/ai-tools/src/llm/audited-run.ts) (`runChatWithToolsAudited`)
  - Execute spreadsheet tools: [`packages/ai-tools/src/llm/integration.ts`](../packages/ai-tools/src/llm/integration.ts) (`SpreadsheetLLMToolExecutor`)
  - Preview + approval gating before mutations: [`packages/ai-tools/src/preview/preview-engine.ts`](../packages/ai-tools/src/preview/preview-engine.ts)
- Desktop glue enforces DLP for the selected range *before* reading sample data or calling the LLM (blocked selections show an error and do not call the model).

### Mode 3: Chat Panel

**Trigger:** User opens AI panel, asks questions
**Latency requirement:** <5s for most queries
**Backend:** Cursor servers

**Code entrypoints (desktop):**
- Chat orchestration (context building + tool loop + audit wiring): [`apps/desktop/src/ai/chat/orchestrator.ts`](../apps/desktop/src/ai/chat/orchestrator.ts)

```typescript
interface ChatMessage {
  role: "user" | "assistant" | "system" | "tool";
  content: string;
  toolCallId?: string;
  attachments?: Attachment[];
}

interface Attachment {
  type: "range" | "chart" | "table" | "formula";
  data: any;
  reference: string;  // e.g., "Sheet1!A1:D10"
}
```

**Capabilities:**

1. **Answer questions about data**
   - "What's the trend in column C?"
   - "Which product has the highest sales?"
   - "Is there a correlation between columns B and D?"

2. **Explain formulas**
   - "Explain the formula in E5"
   - "Why is this formula giving an error?"
   - "How can I simplify this nested IF?"

3. **Generate insights**
   - "Summarize this data"
   - "What anomalies do you see?"
   - "What would you recommend based on this data?"

4. **Execute actions**
   - "Create a pivot table showing sales by region"
   - "Format column B as currency"
   - "Add a chart of monthly revenue"

   Implementation note: pivot creation/refresh is a cross-crate workflow (model schema + engine compute + optional Data Model).
   See [ADR-0005: PivotTables ownership and data flow across crates](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md).

**Implementation notes (actual):**
- The desktop chat panel uses a single React-agnostic orchestrator: [`apps/desktop/src/ai/chat/orchestrator.ts`](../apps/desktop/src/ai/chat/orchestrator.ts) (`createAiChatOrchestrator` / `sendMessage`).
- Each `sendMessage()`:
  1. Builds bounded workbook context (schema-first + RAG) and applies DLP gating.
  2. Runs the provider-agnostic tool loop with audit + (optional) claim verification via
     [`packages/ai-tools/src/llm/audited-run.ts`](../packages/ai-tools/src/llm/audited-run.ts).
  3. Feeds tool results back to the model as **bounded** `role:"tool"` messages using
     `serializeToolResultForModel` (re-exported from [`packages/llm/src/index.js`](../packages/llm/src/index.js) and implemented in
     [`packages/llm/src/toolResultSerialization.js`](../packages/llm/src/toolResultSerialization.js)) to avoid huge `read_range` matrices.

### Mode 4: Cell Functions (=AI())

**Trigger:** User enters AI formula
**Latency requirement:** Async with loading state
**Backend:** Cursor servers

**Code entrypoints (desktop):**
- Formula parsing + provenance + range sampling: [`apps/desktop/src/spreadsheet/evaluateFormula.ts`](../apps/desktop/src/spreadsheet/evaluateFormula.ts)
- Async evaluator + cache + DLP + audit: [`apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts`](../apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts)

```
=AI("Summarize this feedback", A1:A100)
=AI.EXTRACT("email address", B1)
=AI.CLASSIFY("positive/negative/neutral", C1)
=AI.TRANSLATE("Spanish", D1)
```

**Implementation (desktop):**

Cell functions are implemented as an async cell-evaluator + cache:
- Evaluator: [`apps/desktop/src/spreadsheet/evaluateFormula.ts`](../apps/desktop/src/spreadsheet/evaluateFormula.ts)
  - Parses `AI()`, `AI.EXTRACT`, `AI.CLASSIFY`, `AI.TRANSLATE`
  - Preserves provenance for direct cell/range references passed to AI functions:
    - Cell ref: `{ __cellRef: "Sheet1!A1", value: ... }`
    - Range ref: array of provenance cell values, tagged with `__rangeRef` + `__totalCells`
  - Direct range references are sampled (default 200 cells) to avoid materializing unbounded arrays:
    - include a small deterministic prefix (top-left cells)
    - plus a deterministic seeded random sample from across the remainder of the range
- Engine: [`apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts`](../apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts)
  - Returns `#GETTING_DATA` while an LLM request is pending, and reuses cached results.
  - **DLP enforcement**: evaluates `ai.cloudProcessing` policy using per-cell/range classification metadata.
    - BLOCK: returns `#DLP!` without calling the LLM.
    - REDACT: replaces disallowed cell values with `[REDACTED]` before calling the LLM.
  - **Input budgeting**: prompts are built from bounded JSON “compactions” (previews + deterministic samples),
    and include `{ total_cells, sampled_cells, truncated:true }` when ranges are sampled.
  - **Bounded audit logs**: audit entries store hashes + compacted metadata (never full range payloads);
    blocked runs are sanitized.

```typescript
// Pseudocode (high-level):
//
// - `evaluateFormula()` parses the formula and calls the AI engine synchronously:
//   - If the request is new: returns `#GETTING_DATA` and starts async work.
//   - If the request is cached: returns the cached value.
//
// - `AiCellFunctionEngine` handles:
//   - DLP classification from provenance (cell/range refs) + policy enforcement.
//   - Range sampling / prompt compaction (bounded previews + deterministic samples).
//   - Audit logging (hashes + small previews only).
//   - Caching keyed by hashes (prompt + inputs) to keep storage bounded.
```

### Mode 5: Agent Mode

**Trigger:** User enables agent mode, describes goal
**Latency requirement:** Minutes (complex workflows)
**Backend:** Cursor servers with tool calling

**Status:** Implemented in the desktop app.

**Code entrypoints (desktop):**
- Orchestrator: [`apps/desktop/src/ai/agent/agentOrchestrator.ts`](../apps/desktop/src/ai/agent/agentOrchestrator.ts) (`runAgentTask`)
- UI surface: [`apps/desktop/src/panels/ai-chat/AIChatPanelContainer.tsx`](../apps/desktop/src/panels/ai-chat/AIChatPanelContainer.tsx)

Implementation notes (actual):
- Builds bounded workbook context (schema-first + RAG) via [`apps/desktop/src/ai/context/WorkbookContextBuilder.ts`](../apps/desktop/src/ai/context/WorkbookContextBuilder.ts).
- Runs the tool-calling loop via [`packages/ai-tools/src/llm/audited-run.ts`](../packages/ai-tools/src/llm/audited-run.ts) (`runChatWithToolsAudited`, `mode: "agent"`).
- Preview-based approval gating via [`packages/ai-tools/src/preview/preview-engine.ts`](../packages/ai-tools/src/preview/preview-engine.ts)
  - Default configuration requires explicit approval for any non-noop mutation (`approval_cell_threshold: 0`).
  - Optional "continue after deny" allows the model to re-plan when an approval is denied.
- Emits progress events (`planning`, `tool_call`, `tool_result`, `assistant_message`, `complete` / `cancelled` / `error`).

**Capabilities:**
- Multi-step data gathering
- Complex analysis workflows
- Autonomous error correction
- Progress reporting

**Implementation sketch (aligned with code):**
- Entry function: `runAgentTask(params: RunAgentTaskParams)` in
  [`apps/desktop/src/ai/agent/agentOrchestrator.ts`](../apps/desktop/src/ai/agent/agentOrchestrator.ts).
- Agent loop:
  - Builds bounded workbook context via `WorkbookContextBuilder.build(...)` (schema-first + RAG + DLP).
  - Runs the tool-calling loop via `runChatWithToolsAudited` with `SpreadsheetLLMToolExecutor`.
  - Generates previews via `PreviewEngine` and requests approval for any non-noop mutation (default `approval_cell_threshold: 0`).
  - Emits `AgentProgressEvent` updates to the UI.

---

## Context Management

**Code entrypoints:**
- Core context builder (schema + sampling + RAG hooks) + DLP redaction/blocking: [`packages/ai-context/src/contextManager.js`](../packages/ai-context/src/contextManager.js)
- Package usage + examples (token budgeting + DLP notes): [`packages/ai-context/README.md`](../packages/ai-context/README.md)
- Query-aware table/region scoring + selection primitives (reusable; not yet wired into `ContextManager` policy): [`packages/ai-context/src/queryAware.js`](../packages/ai-context/src/queryAware.js)
- Desktop wrapper that builds per-message workbook context + budgets tokens: [`apps/desktop/src/ai/context/WorkbookContextBuilder.ts`](../apps/desktop/src/ai/context/WorkbookContextBuilder.ts)
- Desktop RAG service (persistent local index; deterministic hash embeddings): [`apps/desktop/src/ai/rag/ragService.ts`](../apps/desktop/src/ai/rag/ragService.ts)

Note: The code blocks in the subsections below are **illustrative**. The source-of-truth implementation is in the files
linked above; prefer those over the pseudocode when making product/architecture decisions.

### Prompt context shape (actual)

Context is ultimately fed to the model as a **bounded** markdown-ish string composed of multiple `## <section>` blocks
(packed to a token budget). The exact keys are defined in `packages/ai-context/src/contextManager.js` and currently include:
- `workbook_summary` (sheet list, tables, named ranges)
- `workbook_schema` (schema-first view of tables/columns/types)
- `attachments` (explicit user-provided ranges/charts/etc)
- `retrieved` (RAG chunks pulled from the workbook index)
- optional `dlp` notes when retrieved chunks are redacted

### The Context Problem

A 1M-row spreadsheet could exceed 100M tokens if naively serialized. We need intelligent context selection.

### Context Budget Allocation

```typescript
interface ContextBudget {
  totalTokens: number;       // e.g., 128K for a long-context model
  systemPrompt: number;      // ~2K
  schema: number;            // ~1K
  sampleData: number;        // ~5K
  conversationHistory: number; // ~10K
  retrievedContext: number;  // ~20K
  outputReserve: number;     // ~20K
  available: number;         // ~70K buffer
}

class ContextManager {
  private budget: ContextBudget;
  
  buildContext(request: AIRequest): AIContext {
    const context: AIContext = {
      schema: this.extractSchema(),
      samples: this.getSamples(request),
      relevantRanges: this.findRelevantRanges(request),
      recentHistory: this.getRecentHistory(),
      statistics: this.computeStatistics()
    };
    
    // Ensure within budget
    return this.trimToFit(context);
  }
}
```

### Schema Extraction

```typescript
interface SheetSchema {
  name: string;
  tables: TableSchema[];
  namedRanges: NamedRangeSchema[];
  dataRegions: DataRegion[];
}

interface TableSchema {
  name: string;
  range: string;
  columns: ColumnSchema[];
  rowCount: number;
}

interface ColumnSchema {
  name: string;
  type: InferredType;
  sampleValues: string[];
  statistics?: ColumnStats;
}

function extractSchema(sheet: Sheet): SheetSchema {
  const tables: TableSchema[] = [];
  
  // Find structured tables
  for (const table of sheet.tables) {
    tables.push({
      name: table.name,
      range: rangeToString(table.range),
      columns: table.columns.map(col => ({
        name: col.name,
        type: inferType(col),
        sampleValues: getSampleValues(col, 3),
        statistics: col.type === "number" ? computeStats(col) : undefined
      })),
      rowCount: table.rowCount
    });
  }
  
  // Find implicit data regions
  const dataRegions = detectDataRegions(sheet);
  
  return { name: sheet.name, tables, namedRanges: sheet.namedRanges, dataRegions };
}
```

### Intelligent Sampling

```typescript
class DataSampler {
  // Get representative sample for statistical questions
  getStratifiedSample(range: Range, sampleSize: number): CellValue[][] {
    const data = this.getRangeData(range);
    
    // If small enough, return all
    if (data.length <= sampleSize) return data;
    
    // Stratified sampling for better representation
    const strata = this.identifyStrata(data);
    const samplesPerStratum = Math.ceil(sampleSize / strata.length);
    
    const samples: CellValue[][] = [];
    for (const stratum of strata) {
      const stratumSamples = this.randomSample(stratum, samplesPerStratum);
      samples.push(...stratumSamples);
    }
    
    return samples.slice(0, sampleSize);
  }
  
  // Get edge cases for transformation questions
  getEdgeCaseSample(range: Range, sampleSize: number): CellValue[][] {
    const data = this.getRangeData(range);
    const samples: CellValue[][] = [];
    
    // Include first and last rows
    samples.push(data[0], data[data.length - 1]);
    
    // Include outliers
    const outliers = this.findOutliers(data);
    samples.push(...outliers.slice(0, 3));
    
    // Include nulls/errors if present
    const problemRows = data.filter(row => 
      row.some(cell => cell === null || cell instanceof Error)
    );
    samples.push(...problemRows.slice(0, 2));
    
    // Fill rest with random
    const remaining = sampleSize - samples.length;
    if (remaining > 0) {
      samples.push(...this.randomSample(data, remaining));
    }
    
    return samples.slice(0, sampleSize);
  }
}
```

### RAG Over Cells

For semantic search within large datasets:

> Note: Workbook RAG embeddings are **not user-configurable**. Formula does not accept user API keys or local model
> setup for embeddings. The current implementation uses deterministic hash embeddings (`HashEmbedder`) as a
> privacy/compliance-friendly baseline; a future Cursor-managed embedding service can replace this to improve retrieval
> quality. Hash embeddings are lower quality than modern ML embeddings, but work well enough for basic semantic-ish
> retrieval.

**Code entrypoints:**
- Desktop RAG service wrapper: [`apps/desktop/src/ai/rag/ragService.ts`](../apps/desktop/src/ai/rag/ragService.ts) (`createDesktopRagService`)
- Embeddings: [`packages/ai-rag/src/embedding/hashEmbedder.js`](../packages/ai-rag/src/embedding/hashEmbedder.js) (`HashEmbedder`)
- Workbook chunking/text extraction: [`packages/ai-rag/src/workbook/chunkWorkbook.js`](../packages/ai-rag/src/workbook/chunkWorkbook.js),
  [`packages/ai-rag/src/workbook/chunkToText.js`](../packages/ai-rag/src/workbook/chunkToText.js)
- Indexing pipeline: [`packages/ai-rag/src/pipeline/indexWorkbook.js`](../packages/ai-rag/src/pipeline/indexWorkbook.js)
- Retrieval: [`packages/ai-rag/src/retrieval/searchWorkbookRag.js`](../packages/ai-rag/src/retrieval/searchWorkbookRag.js)
- Vector stores: [`packages/ai-rag/src/store/sqliteVectorStore.js`](../packages/ai-rag/src/store/sqliteVectorStore.js) (browser) and
  [`packages/ai-rag/src/store/inMemoryVectorStore.js`](../packages/ai-rag/src/store/inMemoryVectorStore.js) (tests)
  - Tip: If you delete many chunks (e.g. the workbook structure changes drastically), call
    `await vectorStore.compact()` (alias: `vacuum()`) to run SQLite `VACUUM` and reclaim persisted storage space.

```typescript
// Desktop: build workbook RAG context for a user query (schema + retrieved chunks).
const rag = createDesktopRagService({ documentController, workbookId });
const workbookContext = await rag.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet,
  workbookId,
  query,
  dlp,
});
```

---

## Tool Calling

Spreadsheet tool calling is **provider-agnostic** and shared across Chat, Inline Edit, and Agent modes.

### Canonical tool schemas (source of truth)

- **Canonical tool schemas live in** [`packages/ai-tools/src/tool-schema.ts`](../packages/ai-tools/src/tool-schema.ts).
  - Defines tool names, JSON schemas, and Zod validators (via `TOOL_REGISTRY` / `validateToolCall`).
  - Parameter names are **snake_case** (e.g. `include_formulas`, `formula_template`, `start_row`), matching what the model sees.

### Tool execution

- Spreadsheet tool execution happens in [`packages/ai-tools/src/executor/tool-executor.ts`](../packages/ai-tools/src/executor/tool-executor.ts).
  - Enforces range limits (e.g. `max_read_range_cells`) and (when configured) DLP policy at execution time.
  - **Formula values (opt-in):** by default, formula cells are treated as having “no value” (returned as `null` by `read_range`, ignored by numeric ops, compared by formula text for sort/filter).
    - To support real spreadsheet backends that store both `{ formula, value }`, set `ToolExecutorOptions.include_formula_values = true`.
    - DLP-safe default: when DLP is configured, formula values are only surfaced/used when the range-level DLP decision is **ALLOW** (no redaction). Under **REDACT**, formula-derived values are treated as `null` to avoid inference/exfiltration.
    - Note: ToolExecutor evaluates DLP over the selected range only (it does not trace formula dependencies). Hosts that compute formula values should ensure `cell.value` does not incorporate restricted data outside the evaluated selection.
- The LLM-facing adapter is [`packages/ai-tools/src/llm/integration.ts`](../packages/ai-tools/src/llm/integration.ts) (`SpreadsheetLLMToolExecutor`),
  which connects the ToolExecutor to a host `SpreadsheetApi` (desktop uses `DocumentControllerSpreadsheetApi`).
- Desktop `SpreadsheetApi` implementation: [`apps/desktop/src/ai/tools/documentControllerSpreadsheetApi.ts`](../apps/desktop/src/ai/tools/documentControllerSpreadsheetApi.ts)

### Tool loop orchestration + approval gating

- Provider-agnostic tool-calling loop (streaming): [`packages/llm/src/toolCallingStreaming.js`](../packages/llm/src/toolCallingStreaming.js)
  (`runChatWithToolsStreaming`, re-exported from [`packages/llm/src/index.js`](../packages/llm/src/index.js))
  - Responsible for executing tool calls, appending `role:"tool"` messages, and continuing until the model stops calling tools.
- Audited wrapper used by desktop surfaces: [`packages/ai-tools/src/llm/audited-run.ts`](../packages/ai-tools/src/llm/audited-run.ts)
  - Records tool calls/results + token usage and supports optional post-response claim verification.
- Preview + approval gating helper (used by chat/agent surfaces): [`packages/ai-tools/src/llm/integration.ts`](../packages/ai-tools/src/llm/integration.ts)
  - `createPreviewApprovalHandler` runs `PreviewEngine` on a cloned workbook and denies risky tool calls unless the UI approves.

### Tool result bounding + audit compaction

- **Bounded tool results (for model context)**: [`packages/llm/src/toolResultSerialization.js`](../packages/llm/src/toolResultSerialization.js)
  (`serializeToolResultForModel`, re-exported from [`packages/llm/src/index.js`](../packages/llm/src/index.js)) summarizes high-volume tool results
  (notably `read_range`) before they are appended as
  `role: "tool"` messages.
- **Audit compaction**: [`packages/ai-tools/src/llm/audited-run.ts`](../packages/ai-tools/src/llm/audited-run.ts)
  (`runChatWithToolsAudited*`) stores bounded tool parameters and (by default) stores only `audit_result_summary` rather than full tool results,
  keeping audit logs safe and size-bounded.

> Reminder (Cursor constraints): there are **no local models**, **no user API keys**, and **no provider selection**. All AI requests go through
> Cursor-managed servers and routing.

## DLP Enforcement Surfaces

DLP for cloud AI processing is enforced in multiple layers; **do not rely on a single redaction step**:

- **Context building (prompt construction):** [`packages/ai-context/src/contextManager.js`](../packages/ai-context/src/contextManager.js)
- **Tool execution (tool results are fed back into the model context):** [`packages/ai-tools/src/executor/tool-executor.ts`](../packages/ai-tools/src/executor/tool-executor.ts)
- **AI cell functions:** [`apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts`](../apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts)

This redundancy matters because tool results become part of the conversation history (`role:"tool"`) and will be sent back to the model on
subsequent turns.

---

## Safety and Verification

### Preview Before Apply

(Illustrative interfaces below; see the real preview implementation in
[`packages/ai-tools/src/preview/preview-engine.ts`](../packages/ai-tools/src/preview/preview-engine.ts).)

```typescript
interface ActionPreview {
  description: string;
  affectedCells: CellRef[];
  changes: ChangePreview[];
  warnings: Warning[];
  requiresApproval: boolean;
}

interface ChangePreview {
  cell: CellRef;
  currentValue: CellValue;
  newValue: CellValue;
  type: "modify" | "create" | "delete";
}

class PreviewEngine {
  async generatePreview(action: AIAction): Promise<ActionPreview> {
    // Simulate action without applying
    const simulation = await this.simulate(action);
    
    // Identify affected cells
    const affected = this.getAffectedCells(simulation);
    
    // Generate change preview
    const changes = affected.map(cell => ({
      cell,
      currentValue: this.getCurrentValue(cell),
      newValue: simulation.getNewValue(cell),
      type: this.getChangeType(cell, simulation)
    }));
    
    // Check for warnings
    const warnings = this.checkWarnings(action, simulation);
    
    // Determine if approval needed
    const requiresApproval = 
      affected.length > 100 ||
      warnings.some(w => w.severity === "high") ||
      action.type === "delete";
    
    return {
      description: this.generateDescription(action),
      affectedCells: affected,
      changes: changes.slice(0, 20),  // Preview first 20
      warnings,
      requiresApproval
    };
  }
}
```

### Verification Layer

(Illustrative; real verification hooks live in
[`packages/ai-tools/src/llm/verification.ts`](../packages/ai-tools/src/llm/verification.ts) and are wired via
[`packages/ai-tools/src/llm/audited-run.ts`](../packages/ai-tools/src/llm/audited-run.ts).)

```typescript
class AIVerifier {
  // For factual claims about data, verify against actual values
  async verifyResponse(response: AIResponse, context: AIContext): Promise<VerificationResult> {
    const claims = this.extractClaims(response);
    const results: ClaimVerification[] = [];
    
    for (const claim of claims) {
      const verified = await this.verifyClaim(claim, context);
      results.push(verified);
    }
    
    return {
      allVerified: results.every(r => r.verified),
      claims: results,
      confidence: this.calculateConfidence(results)
    };
  }
  
  private async verifyClaim(claim: Claim, context: AIContext): Promise<ClaimVerification> {
    switch (claim.type) {
      case "numeric":
        // Verify numeric claims by computing
        const computed = await this.computeValue(claim.formula, context);
        return {
          claim,
          verified: Math.abs(computed - claim.value) < 0.01,
          actual: computed
        };
        
      case "existence":
        // Verify cell references exist
        const exists = this.cellExists(claim.reference);
        return { claim, verified: exists };
        
      case "comparison":
        // Verify comparisons
        const values = await this.getValues(claim.references);
        const holds = this.evaluateComparison(values, claim.operator);
        return { claim, verified: holds, actual: values };
        
      default:
        return { claim, verified: false, reason: "Cannot verify" };
    }
  }
}
```

### Audit Trail

(Illustrative; audit store interfaces + recorder live in `packages/ai-audit` (see
[`packages/ai-audit/src/store.ts`](../packages/ai-audit/src/store.ts) and
[`packages/ai-audit/src/recorder.ts`](../packages/ai-audit/src/recorder.ts)).)

```typescript
interface AIAuditEntry {
  timestamp: Date;
  userId: string;
  sessionId: string;
  
  mode: AIMode;
  input: string;
  context: string;  // Summarized
  
  model: string;
  promptTokens: number;
  completionTokens: number;
  latency: number;
  
  actions: ActionLog[];
  verification?: VerificationResult;
  
  userFeedback?: "accepted" | "rejected" | "modified";
}

class AIAuditLog {
  async log(entry: AIAuditEntry): Promise<void> {
    await this.db.insert("ai_audit_log", entry);
  }
  
  async getHistory(filters: AuditFilters): Promise<AIAuditEntry[]> {
    return this.db.query("ai_audit_log", filters);
  }
  
  async exportForCompliance(dateRange: DateRange): Promise<ComplianceReport> {
    const entries = await this.getHistory({ dateRange });
    return this.generateComplianceReport(entries);
  }
}
```

---

## Backend Configuration

### Cursor AI Integration

> **All AI is managed by Cursor.** There are no user-configurable API keys, no provider selection, and no local models. Cursor controls the harness, prompts, and model routing.

**Code entrypoints (desktop + shared):**
- Desktop LLM client wrapper (Cursor-only; no provider selection / API keys): [`apps/desktop/src/ai/llm/desktopLLMClient.ts`](../apps/desktop/src/ai/llm/desktopLLMClient.ts) (`getDesktopLLMClient`, `getDesktopModel`)
- Shared Cursor-only LLM client (exported from [`packages/llm/src/index.js`](../packages/llm/src/index.js)):
  - [`packages/llm/src/createLLMClient.js`](../packages/llm/src/createLLMClient.js) (throws if provider selection is attempted)
  - [`packages/llm/src/cursor.js`](../packages/llm/src/cursor.js) (`CursorLLMClient`; does not read user API keys)

```typescript
// Configuration is Cursor-managed, not user-configurable
interface CursorAIConfig {
  // These are set by Cursor, not by users
  mode: AIMode;
  maxTokens: number;
  timeout: number;
}

// Latency targets by mode (Cursor optimizes routing to meet these)
const LATENCY_TARGETS: Record<AIMode, number> = {
  tabCompletion: 100,    // <100ms
  inlineEdit: 2000,      // <2s
  chat: 5000,            // <5s
  cellFunction: 10000,   // async with loading state
  agent: 60000           // minutes for complex workflows
};
```

### Client integration (how requests reach Cursor)

- Desktop code calls `getDesktopLLMClient()` (see link above), which internally uses `createLLMClient()` to construct a Cursor-backed `LLMClient`.
- `createLLMClient()` is intentionally **Cursor-only**:
  - rejects provider selection at construction time
  - does **not** read user API keys from env/local storage
- All tool-calling requests (chat / inline edit / agent) are sent with tool definitions and return tool calls in the provider-agnostic format used by `packages/llm`.

---

## Testing Strategy

Note: The snippets below are illustrative. For real tests, see:
- `packages/ai-completion/test/*` (Node `node:test`)
- `packages/ai-tools/src/**/*.test.ts` / `*.vitest.ts`
- `apps/desktop/src/ai/**/__tests__` and `apps/desktop/src/spreadsheet/*test.ts`

### AI Response Testing

```typescript
describe("AI Integration", () => {
  describe("Tab Completion", () => {
    it("suggests formula completions", async () => {
      const engine = new TabCompletionEngine();
      
      const suggestions = await engine.getSuggestions({
        currentInput: "=SUM(",
        cursorPosition: 5,
        cellRef: { row: 0, col: 0 },
        surroundingCells: mockSurroundingCells
      });
      
      expect(suggestions.length).toBeGreaterThan(0);
      expect(suggestions[0].text).toMatch(/^=SUM\(/);
    });
    
    it("completes range based on data", async () => {
      const engine = new TabCompletionEngine();
      
      // Setup: column A has numbers in A1:A10
      const suggestions = await engine.getSuggestions({
        currentInput: "=SUM(A",
        cursorPosition: 6,
        cellRef: { row: 11, col: 1 },
        surroundingCells: { A1_A10: "numbers" }
      });
      
      expect(suggestions).toContainEqual(
        expect.objectContaining({ text: "=SUM(A1:A10)" })
      );
    });
  });
  
  describe("Tool Execution", () => {
    it("applies formula to column", async () => {
      const executor = new ToolExecutor(mockSheet);
      
      const result = await executor.execute({
        name: "apply_formula_column",
        parameters: {
          column: "C",
          formulaTemplate: "=A{row}+B{row}",
          startRow: 1,
          endRow: 10
        }
      });
      
      expect(result.success).toBe(true);
      expect(mockSheet.getCell(1, 2)).toBe("=A1+B1");
      expect(mockSheet.getCell(10, 2)).toBe("=A10+B10");
    });
  });
});
```

### Evaluation Metrics

```typescript
interface AIEvaluationMetrics {
  // Accuracy
  formulaCorrectness: number;      // % of generated formulas that calculate correctly
  taskCompletionRate: number;      // % of tasks completed successfully
  
  // User acceptance
  suggestionAcceptRate: number;    // % of suggestions accepted
  editAfterAcceptRate: number;     // % of accepted suggestions later edited
  
  // Performance
  averageLatency: Record<AIMode, number>;
  p99Latency: Record<AIMode, number>;
  
  // Safety
  hallucinationRate: number;       // % of responses with unverifiable claims
  undoRate: number;                // % of AI actions undone by user
}
```
