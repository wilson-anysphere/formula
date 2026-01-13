# `@formula/ai-context`

Utilities for turning large spreadsheets into **LLM-friendly context** safely and predictably.

This package focuses on:

- **Schema-first extraction**: infer table/region structure, headers, and column types without sending full raw data.
- **Sampling**: random and stratified row sampling for “show me examples” / statistical questions.
- **RAG over cells**: lightweight similarity search for relevant regions (single-sheet) and workbook-scale retrieval via `packages/ai-rag`.
- **Token budgeting**: build context strings that fit a budget (with deterministic trimming).
- **Message trimming**: trim/summarize conversation history to fit a model context window.
- **DLP hooks**: integrate structured DLP policy enforcement (block/redact) plus heuristic redaction helpers.

> Note: This repo uses ESM. The public entrypoint is `packages/ai-context/src/index.js`.
>
> Since this README lives in `packages/ai-context/`, the examples use the equivalent relative import: `./src/index.js`.

---

## Exports (high level)

```js
import {
  // Schema / sampling
  extractSheetSchema,
  randomSampleRows,
  stratifiedSampleRows,

  // Context builders
  ContextManager,

  // Token budgeting
  createHeuristicTokenEstimator,
  estimateTokens,
  estimateToolDefinitionTokens,
  packSectionsToTokenBudget,
  stableJsonStringify,

  // Message trimming
  trimMessagesToBudget,

  // Heuristic DLP helpers (defense-in-depth)
  classifyText,
  redactText,
} from "./src/index.js";
```

---

## Example: extract sheet schema + summarize

`extractSheetSchema()` produces a compact representation of what’s in a sheet: detected data regions, headers, inferred types, and a few sample values per column.

```js
import { extractSheetSchema, redactText, stableJsonStringify } from "./src/index.js";

const sheet = {
  name: "Transactions",
  values: [
    ["Date", "Merchant", "Category", "Amount"],
    ["2025-01-01", "Coffee Shop", "Food", 4.25],
    ["2025-01-02", "Grocery Store", "Food", 82.1],
    ["2025-01-03", "Ride Share", "Transport", 14.5],
  ],
};

const schema = extractSheetSchema(sheet);

// ⚠️ DLP NOTE:
// - `schema` may contain snippets of real cell content (e.g., sample values).
// - Do NOT send the raw object to a cloud model without redaction/policy checks.
//
// For quick demos you can use heuristic redaction on the *serialized* form:
const schemaForPrompt = redactText(stableJsonStringify(schema));

const prompt = [
  "You are helping a user understand a spreadsheet.",
  "Summarize what this sheet contains and what questions it can answer.",
  "",
  "Sheet schema:",
  schemaForPrompt,
].join("\n");

// await llm.chat([{ role: "user", content: prompt }])
```

If you need compliance-grade guarantees, prefer **structured DLP policy enforcement** (see [DLP safety](#dlp-safety)).

---

## Example: build single-sheet context with `ContextManager.buildContext`

`ContextManager.buildContext()` is the “easy button” for building a **token-budgeted prompt context string** for a single sheet:

- extracts a schema
- samples rows
- indexes the sheet and retrieves relevant regions (“RAG”)
- redacts/filters content via a pluggable `redactor`
- packs everything into a single string within `tokenBudgetTokens`

```js
import { ContextManager } from "./src/index.js";

const sheet = {
  name: "Transactions",
  values: [
    ["Date", "Merchant", "Category", "Amount"],
    ["2025-01-01", "Coffee Shop", "Food", 4.25],
    ["2025-01-02", "Grocery Store", "Food", 82.1],
    ["2025-01-03", "Ride Share", "Transport", 14.5],
    // ...
  ],
};

const cm = new ContextManager({
  // Total budget for the context string returned as `promptContext`.
  tokenBudgetTokens: 12_000,
});

const query = "What categories are driving the highest spend?";

const ctx = await cm.buildContext({
  sheet,
  query,
  // Optional knobs:
  sampleRows: 25,
  samplingStrategy: "stratified",
  stratifyByColumn: 2, // stratify by the "Category" column
  attachments: [{ type: "range", reference: "Transactions!A1:D100" }],
});

// ✅ Send THIS to the model.
console.log(ctx.promptContext);

// ⚠️ Footgun:
// Avoid serializing `ctx.schema`, `ctx.sampledRows`, or `ctx.retrieved` directly into prompts.
// Treat them as structured/diagnostic outputs unless you have redaction + policy checks.
```

### What `buildContext()` returns

- `promptContext` (string): **token-budgeted, redacted** context intended to be included in an LLM prompt.
- `schema` (object): extracted sheet schema (may include sample values).
- `sampledRows` (array): sampled rows (may include raw values).
- `retrieved` (array): retrieved region previews (may include raw values).

In most integrations, only `promptContext` should ever reach a cloud model.

---

## Example: build workbook context with `ContextManager.buildWorkbookContextFromSpreadsheetApi`

For workbook-scale retrieval (multiple sheets, persistent indexing, incremental updates), `ContextManager` integrates with `packages/ai-rag`.

### Minimal in-memory setup

```js
import { ContextManager } from "./src/index.js";
import { HashEmbedder, InMemoryVectorStore } from "../ai-rag/src/index.js";

// A SpreadsheetApi-like adapter. `workbookFromSpreadsheetApi` (ai-rag) only needs:
// - listSheets(): string[]
// - listNonEmptyCells(sheetName?): Array<{ address: { sheet, row, col }, cell: { value?, formula? } }>
//
// Note: coordinates are 1-based by default (A1 => row=1,col=1).
const spreadsheet = {
  listSheets() {
    return ["Sheet1"];
  },
  listNonEmptyCells(sheetName) {
    return [
      { address: { sheet: sheetName, row: 1, col: 1 }, cell: { value: "Header" } },
      { address: { sheet: sheetName, row: 2, col: 1 }, cell: { value: "Value" } },
    ];
  },
};

const embedder = new HashEmbedder({ dimension: 384 });
const vectorStore = new InMemoryVectorStore({ dimension: embedder.dimension });

const cm = new ContextManager({
  tokenBudgetTokens: 24_000,
  workbookRag: {
    vectorStore,
    embedder,
    topK: 8,
  },
});

const workbookCtx = await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet,
  workbookId: "workbook-123",
  query: "What does this workbook track?",
});

console.log(workbookCtx.promptContext);
```

### What `buildWorkbookContextFromSpreadsheetApi()` returns

- `promptContext` (string): token-budgeted context intended for an LLM prompt.
- `retrieved` (array): retrieved chunks with prompt-safe `text` plus metadata.
- `indexStats` (object): indexing stats from `packages/ai-rag` (helpful for performance/debug).

---

## Example: trim conversation history with `trimMessagesToBudget`

Most LLM APIs enforce a *single* context window limit for:

- system prompt
- tool definitions (if any)
- conversation history
- your spreadsheet context (`promptContext`)
- the model’s output

`trimMessagesToBudget()` helps keep the **message list** within a budget by:

- preserving non-generated system messages
- keeping the most recent N messages
- replacing older history with a single summary message (optional)

```js
import { trimMessagesToBudget, createHeuristicTokenEstimator } from "./src/index.js";

const estimator = createHeuristicTokenEstimator();

const messages = [
  { role: "system", content: "You are a helpful assistant." },
  { role: "user", content: "Here is my sheet..." },
  // ... lots of messages ...
  { role: "user", content: "Now answer my question." },
];

const trimmed = await trimMessagesToBudget({
  messages,
  maxTokens: 128_000, // model context window
  reserveForOutputTokens: 4_000,
  estimator,

  // Optional knobs:
  keepLastMessages: 40,
  summaryMaxTokens: 256,
  // summarize: async (msgsToSummarize) => "(your model summary here)",
});

// Use `trimmed` for the LLM call.
```

---

## DLP safety

This package supports two different (complementary) safety mechanisms:

### 1) Structured DLP (policy engine) — **compliance-grade**

Structured DLP lives in `packages/security/dlp` and is enforced in `ContextManager` when you pass the `dlp` option to:

- `ContextManager.buildContext({ dlp: ... })`
- `ContextManager.buildWorkbookContext({ dlp: ... })`
- `ContextManager.buildWorkbookContextFromSpreadsheetApi({ dlp: ... })`

The policy engine can return decisions like **ALLOW**, **REDACT**, or **BLOCK** for the `"AI_CLOUD_PROCESSING"` action. `ContextManager` will:

- throw on **BLOCK** (callers should catch `DlpViolationError`)
- redact content on **REDACT** (sheet cells or retrieved chunks depending on the call)
- emit audit events when an `auditLogger` is provided

Use this when you have real classification data (document/sheet/range/cell scopes) and need deterministic enforcement.

### 2) Heuristic redaction — **defense-in-depth only**

`classifyText()` / `redactText()` are small regex-based helpers (e.g., emails/SSNs/credit cards). They:

- reduce accidental leakage
- help keep retrieval/indexing “safe by default”

They are **not** a substitute for structured DLP.

### Critical guidance (avoid common footguns)

- Prefer sending only `promptContext` to cloud models.
- Treat `schema`, `sampledRows`, and `retrieved` as **structured internal outputs**.
  - If you must serialize them into prompts, **redact first** and ensure structured DLP policy allows it.
- Heuristic redaction is incomplete by design. Do not rely on it for compliance.

### Example: enforcing structured DLP when building context

```js
import { ContextManager } from "./src/index.js";
import { createDefaultOrgPolicy, DlpViolationError } from "../security/dlp/index.js";

const policy = createDefaultOrgPolicy();
const cm = new ContextManager();

try {
  const ctx = await cm.buildContext({
    sheet,
    query,
    dlp: {
      documentId: "doc-123",
      // sheetId is optional; defaults to `sheet.name` when omitted.
      policy,
      // Provide either:
      // - classificationRecords: [{ selector, classification }, ...]
      // - classificationStore: { list(documentId) => [{ selector, classification }, ...] }
      classificationRecords: [],
      auditLogger: { log: (event) => console.debug("DLP audit", event) },
    },
  });

  // Send only `ctx.promptContext` to the model.
  console.log(ctx.promptContext);
} catch (err) {
  if (err instanceof DlpViolationError) {
    // Show a user-facing message and do not call cloud LLMs with blocked content.
    console.error(err.message);
    return;
  }
  throw err;
}
```

---

## Token budgeting (how it works, how to tune it)

### How context token budgeting works (`ContextManager`)

`ContextManager` builds context as **sections** and then packs them into a budget with `packSectionsToTokenBudget()`:

- sections are sorted by `priority` (higher first)
- each section is trimmed to fit the *remaining* token budget
- low-priority sections may be fully omitted when budgets are tight

For single-sheet context (`buildContext()`), the default section priorities are:

1. `dlp` (if present) — priority 5
2. `retrieved` — priority 4
3. `schema` — priority 3
4. `attachments` — priority 2
5. `samples` — priority 1

This means that when budgets shrink, **samples drop first**, then attachments, then schema, etc.

### How message token budgeting works (`trimMessagesToBudget`)

`trimMessagesToBudget()` enforces:

```
allowed_prompt_tokens = maxTokens - reserveForOutputTokens
```

It then fits your message list into `allowed_prompt_tokens` by keeping system messages, keeping the newest messages, and (optionally) inserting a compact summary message.

### Tuning tips

- **Start with a top-level budget** based on your model’s context window. Keep a safety buffer because token estimates are heuristic.
  - Example: for a 128k model, target ~110k “prompt tokens” and reserve the rest.
- **Allocate budgets explicitly**:
  - `reserveForOutputTokens` (LLM output)
  - system prompt (often larger than expected)
  - tool schemas (if any)
  - conversation history (via `trimMessagesToBudget`)
  - spreadsheet context (via `ContextManager.tokenBudgetTokens`)
- If you are using tool calling, include tool schema tokens in your budgeting:
  ```js
  import { estimateToolDefinitionTokens, createHeuristicTokenEstimator } from "./src/index.js";
  const estimator = createHeuristicTokenEstimator();
  const toolTokens = estimateToolDefinitionTokens(tools, estimator);
  ```
- **Tune what drives size**:
  - `ContextManager.tokenBudgetTokens`: total context size.
  - `buildContext({ sampleRows })`: fewer rows → smaller prompts.
  - `buildWorkbookContext({ topK })`: fewer retrieved chunks → smaller prompts.
  - `createHeuristicTokenEstimator({ charsPerToken })`: adjust if your text is far from English-like (still approximate).
