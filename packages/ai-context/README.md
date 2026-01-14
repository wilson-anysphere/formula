# `@formula/ai-context`

Utilities for turning large spreadsheets into **LLM-friendly context** safely and predictably.

This package focuses on:

- **Schema-first extraction**: infer table/region structure, headers, and column types without sending full raw data.
- **Sampling**: random, stratified, head, tail, and systematic row sampling for “show me examples” / statistical questions.
- **RAG over cells**: lightweight similarity search for relevant regions (single-sheet) and workbook-scale retrieval via `packages/ai-rag`.
- **Token budgeting**: build context strings that fit a budget (with deterministic trimming).
- **Message trimming**: trim/summarize conversation history to fit a model context window.
- **DLP hooks**: integrate structured DLP policy enforcement (block/redact) plus heuristic redaction helpers.

> Note: This repo uses ESM. The public entrypoint is `packages/ai-context/src/index.js`.
>
> Since this README lives in `packages/ai-context/`, the examples use the equivalent relative import: `./src/index.js`.

---

## Where this is used (in this repo)

- Desktop chat orchestration: [`apps/desktop/src/ai/chat/orchestrator.ts`](../../apps/desktop/src/ai/chat/orchestrator.ts)
- Desktop workbook context wrapper: [`apps/desktop/src/ai/context/WorkbookContextBuilder.ts`](../../apps/desktop/src/ai/context/WorkbookContextBuilder.ts)

---

## Contents

- [Exports](#exports-high-level)
- [`ContextManager` options](#contextmanager-constructor-options)
- Examples
  - [Extract schema + summarize](#example-extract-sheet-schema--summarize)
  - [Extract workbook schema](#example-extract-workbook-schema-tables--named-ranges)
  - [Single-sheet context](#example-build-single-sheet-context-with-contextmanagerbuildcontext)
  - [Workbook context](#example-build-workbook-context-with-contextmanagerbuildworkbookcontextfromspreadsheetapi)
  - [Trim conversation history](#example-trim-conversation-history-with-trimmessagestobudget)
- [DLP safety](#dlp-safety)
- [Token budgeting](#token-budgeting-how-it-works-how-to-tune-it)
- [Troubleshooting](#troubleshooting--gotchas)

---

## Exports (high level)

```js
import {
  // Schema / sampling
  extractSheetSchema,
  extractWorkbookSchema,
  summarizeSheetSchema,
  summarizeWorkbookSchema,
  summarizeRegion,
  randomSampleRows,
  stratifiedSampleRows,
  headSampleRows,
  tailSampleRows,
  systematicSampleRows,

  // Query-aware selection
  scoreRegionForQuery,
  pickBestRegionForQuery,

  // TSV helpers (streaming-ish range formatting)
  valuesRangeToTsv,

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

## `ContextManager` (constructor options)

`ContextManager` is the main “batteries included” API.

```js
import { ContextManager } from "./src/index.js";

const cm = new ContextManager({
  tokenBudgetTokens: 16_000, // default
  // tokenEstimator: createHeuristicTokenEstimator(), // optional, but recommended to keep budgeting consistent
  // redactor: (text) => text, // optional (defaults to `redactText`)
  // ragIndex: new RagIndex(), // optional single-sheet RAG index
  // cacheSheetIndex: true, // default: cache single-sheet indexing by content signature
  // sheetIndexCacheLimit: 32, // default: max cached sheets per ContextManager instance
  // workbookRag: { vectorStore, embedder, topK }, // required for workbook context builders
});
```

Key options:

- `tokenBudgetTokens`: max tokens for `promptContext` produced by context builders.
- `tokenEstimator`: used for *all* token budgeting inside `ContextManager`. If you’re also using `trimMessagesToBudget`, pass the **same estimator** there too so budgets line up.
- `redactor(text)`: last-mile redaction for prompt-facing strings. This is **not** a replacement for structured DLP.
- `cacheSheetIndex`: when `true` (default), `buildContext()` will only call `ragIndex.indexSheet()` when the sheet’s
  content signature changes. This avoids repeated re-embedding on every call for unchanged sheets.
- `sheetIndexCacheLimit`: LRU limit for cached sheet signatures (defaults to 32). When an active sheet entry is evicted,
  its in-memory RAG chunks are also removed from the underlying store to keep memory bounded.
- `workbookRag`: enables workbook retrieval (`buildWorkbookContext*`). Requires:
  - `vectorStore` implementing the `packages/ai-rag` store interface used by indexing + retrieval:
    - `list(...)` (to load existing chunk hashes)
    - `upsert(...)` (to persist new/updated embeddings)
    - `delete(...)` (to remove stale chunks)
    - `query(...)` (to retrieve top-K chunks for a query)
    - (optionally) `close()`
    - e.g. `InMemoryVectorStore`, `SqliteVectorStore`
  - `embedder` implementing `embedTexts([...])` (e.g. `HashEmbedder`)
  - Optional tuning:
    - `topK`: default number of retrieved chunks per query (can be overridden per call).
    - `sampleRows`: how many rows of each workbook chunk to include when generating chunk text during indexing (lower keeps the index smaller/faster).

Many APIs in this package also accept `signal?: AbortSignal` to allow cancellation from UI surfaces.

### Clearing the single-sheet cache

If you reuse a `ContextManager` instance across many sheets and want to drop cached state (or free memory),
use:

```js
await cm.clearSheetIndexCache({ clearStore: true });
```

This clears both:

- the internal sheet signature cache used by `buildContext()`
- the in-memory sheet-level RAG chunks in `ragIndex.store` (optional via `clearStore: true`)

Cancellation behavior:

- If `signal.aborted` is set (or becomes set during work), helpers will throw an `Error` whose `name` is `"AbortError"`.
- This does not necessarily cancel underlying work (e.g. an embedder call), but it lets callers stop awaiting promptly.

```js
try {
  const ctx = await cm.buildContext({ sheet, query, signal });
  // ...
} catch (err) {
  if (err && typeof err === "object" && err.name === "AbortError") return;
  throw err;
}
```

---

## Example: extract sheet schema + summarize

`extractSheetSchema()` produces a compact representation of what’s in a sheet: detected data regions, headers, inferred types, and a few sample values per column.

```js
import { extractSheetSchema, summarizeSheetSchema, redactText } from "./src/index.js";

const sheet = {
  name: "Transactions",
  values: [
    ["Date", "Merchant", "Category", "Amount"],
    ["2025-01-01", "Coffee Shop", "Food", 4.25],
    ["2025-01-02", "Grocery Store", "Food", 82.1],
    ["2025-01-03", "Ride Share", "Transport", 14.5],
  ],
};

// Optional: when `values` is a cropped window of a larger sheet (0-based coordinates),
// provide `origin` so schema + retrieved ranges are emitted as absolute A1 ranges.
//
// Example: this `values` matrix corresponds to K11:L12 in the full sheet:
// const sheet = {
//   name: "Transactions",
//   origin: { row: 10, col: 10 },
//   values: [
//     ["X", "Y"],
//     ["Z", "W"],
//   ],
// };

const schema = extractSheetSchema(sheet);

// ⚠️ DLP NOTE:
// - `schema` may contain snippets of real cell content (e.g., sample values).
// - Do NOT send the raw object to a cloud model without redaction/policy checks.
//
// For prompt context, prefer `summarizeSheetSchema(schema)` which is:
// - deterministic + compact
// - schema-first (headers/types/row counts)
// - does NOT include sample cell values from `TableSchema.columns[*].sampleValues`
const schemaForPrompt = redactText(summarizeSheetSchema(schema));

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

## Example: extract workbook schema (tables + named ranges)

When you have workbook metadata (sheet list, table rects, named ranges) but don’t want to run full RAG indexing,
`extractWorkbookSchema()` produces a compact, deterministic summary:

```js
import { extractWorkbookSchema, summarizeWorkbookSchema } from "./src/index.js";

const workbook = {
  id: "workbook-123",
  sheets: [{ name: "Sheet1", cells: [["Product", "Sales"], ["Alpha", 10]] }],
  tables: [{ name: "SalesTable", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
  namedRanges: [{ name: "SalesData", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
};

const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 50 });
console.log(schema.tables[0].rangeA1); // "Sheet1!A1:B2"

// For prompt-friendly context you can format it as a compact deterministic string:
const summary = summarizeWorkbookSchema(schema);
console.log(summary);
```

Like `extractSheetSchema()`, this may include snippets of real cell content (headers). Apply redaction/policy checks
before sending any serialized schema to a cloud model.

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
  // Limits (optional):
  // - These are safety rails; overriding them can increase memory/CPU.
  // - `maxChunkRows` controls how many TSV rows are included in each retrieved chunk preview.
  // - `splitRegions` can improve retrieval quality for very tall tables by indexing multiple
  //   row windows per region (at the cost of a larger in-memory index).
  //
  // limits: {
  //   maxContextRows: 1000,
  //   maxContextCells: 200_000,
  //   maxChunkRows: 30,
  //   splitRegions: false,
  //   chunkRowOverlap: 3,
  //   maxChunksPerRegion: 50,
  // },
  // Sampling strategies:
  // - "random" (default)
  // - "stratified"
  // - "head"
  // - "tail"
  // - "systematic"
  samplingStrategy: "stratified",
  stratifyByColumn: 2, // stratify by the "Category" column
  attachments: [{ type: "range", reference: "Transactions!A1:D100" }],
});

// ✅ Send THIS to the model.
console.log(ctx.promptContext);

// ⚠️ Footgun:
// Avoid serializing `ctx.schema`, `ctx.sampledRows`, or `ctx.retrieved` directly into prompts.
// Treat them as structured/diagnostic outputs unless you have redaction + policy checks.
// When DLP is enabled and policy requires redaction, ContextManager will also redact these
// structured fields before returning (defense-in-depth).
```

### What `buildContext()` returns

- `promptContext` (string): **token-budgeted, redacted** context intended to be included in an LLM prompt.
- `schema` (object): extracted sheet schema (may include sample values; redacted under DLP REDACT).
- `sampledRows` (array): sampled rows (may include raw values; redacted under DLP REDACT).
- `retrieved` (array): retrieved region previews (may include raw values; redacted under DLP REDACT).

In most integrations, only `promptContext` should ever reach a cloud model.

#### Quick safety checklist (what can be sent to cloud models?)

| Value | Safe to send to a cloud model by default? | Notes |
|---|---:|---|
| `promptContext` | ✅ Yes | Designed to be prompt-facing and budgeted; still subject to your org’s DLP policy. |
| `schema` | ❌ No | May include sample values/snippets; treat as sensitive unless redacted + policy-allowed. |
| `sampledRows` | ❌ No | Contains raw cell values by design. |
| `retrieved` | ❌ No | Easy footgun: callers may serialize it directly; prefer the already-packed `promptContext`. |

### Performance & safety limits (single-sheet)

`buildContext()` expects `sheet.values` to be a 2D array, but real spreadsheets can accidentally materialize huge matrices
(e.g. full-row/column selections). To avoid OOMs and runaway indexing, `ContextManager.buildContext()` **caps** the matrix
it uses for schema extraction + sampling + single-sheet RAG:

- scans at most **1,000 rows**
- caps included columns to **500** (protects against “short but extremely wide” selections)
- caps total scanned cells to ~**200,000** by shrinking the column count accordingly

These caps can be configured:

- globally via `new ContextManager({ maxContextRows, maxContextCols, maxContextCells })`
- per-call via `buildContext({ limits: { maxContextRows, maxContextCols, maxContextCells } })`

If you pass user selections (especially whole-row selections / large pasted ranges), setting a conservative `maxContextCols`
is recommended to prevent prompt bloat and O(width) work across thousands of columns.

If you need larger coverage, prefer workbook RAG (`buildWorkbookContext*`) where the source of truth is a **sparse** list of
non-empty cells (via `SpreadsheetApi.listNonEmptyCells`) rather than a dense matrix.

### Improving retrieval quality for tall tables (single-sheet)

By default, single-sheet RAG indexes **one chunk per detected region**, and the chunk text includes only the first `maxChunkRows`
rows (plus an ellipsis line when truncated). This keeps the in-memory index small and stable.

For very tall regions (e.g. 200k rows × 1 column), queries that reference values near the bottom of the table may not retrieve well,
since the indexed preview only contains the top rows.

To improve this, you can opt into **row-window chunking**:

```js
const ctx = await cm.buildContext({
  sheet,
  query: "specialtoken",
  limits: {
    maxChunkRows: 30,
    splitRegions: true,
    chunkRowOverlap: 3,
    maxChunksPerRegion: 20,
  },
});
```

When enabled, each tall region is split into multiple fixed-size row windows, each with its own chunk id and A1 range.
This increases index size but typically improves recall for “bottom of table” queries.

---

## Example: build workbook context with `ContextManager.buildWorkbookContextFromSpreadsheetApi`

For workbook-scale retrieval (multiple sheets, persistent indexing, incremental updates), `ContextManager` integrates with `packages/ai-rag`.

### Minimal in-memory setup

```js
import { ContextManager } from "./src/index.js";
import { HashEmbedder } from "../ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../ai-rag/src/store/inMemoryVectorStore.js";

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

#### Quick safety checklist (workbook)

| Value | Safe to send to a cloud model by default? | Notes |
|---|---:|---|
| `promptContext` | ✅ Yes | Designed to be prompt-facing and budgeted; still subject to your org’s DLP policy. |
| `retrieved` | ⚠️ Usually | Chunk `text` is redacted/prompt-safe and `metadata.text` is stripped to avoid footguns, but prefer `promptContext` unless you need custom formatting. |
| `indexStats` | ❌ No | Diagnostic/perf info; not useful in prompts and may bloat context. |

### Common options for `buildWorkbookContextFromSpreadsheetApi()`

```js
const workbookCtx = await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet,
  workbookId,
  query,

  // Retrieval:
  topK: 8,

  // Indexing:
  // - If you maintain an incremental/persistent index elsewhere, you can skip indexing here.
  // - Safety: when `dlp` is enabled, indexing normally still runs (so chunk redaction can be applied
  //   before embedding/persistence). Use `skipIndexingWithDlp: true` only if you already enforced DLP
  //   during indexing.
  skipIndexing: false,
  skipIndexingWithDlp: false,

  // Output shaping:
  // - Set false if you only want structured `retrieved` results and will format your own prompt.
  includePromptContext: true,

  // Cancellation:
  signal,

  // Structured DLP (optional):
  dlp,
});
```

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
- preserving tool-call coherence (assistant messages with `toolCalls` + their `role: "tool"` results)
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
  // Tool-call coherence options:
  // - preserveToolCallPairs: true (default) prevents orphan tool messages / tool calls.
  // - dropToolMessagesFirst: false (default) can be set true to prefer summarizing completed
  //   tool-call groups (assistant(toolCalls)+tool*) before dropping other history.
  // summarize: async (msgsToSummarize) => "(your model summary here)",
});

// Use `trimmed` for the LLM call.
```

Notes:

- By default, `trimMessagesToBudget()` uses a **deterministic stub summarizer** (no model call). If you provide a
  `summarize()` callback that calls an LLM, treat that as another cloud AI request and apply the same DLP/policy
  checks you use for your main chat request.
- Generated summaries are marked with `[CONTEXT_SUMMARY]` so repeated trimming can replace/merge older summaries
  instead of accumulating them.

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
- redact *sensitive queries* before embedding when policy would not allow cloud AI processing for that query (defense-in-depth)
- emit audit events when an `auditLogger` is provided

Use this when you have real classification data (document/sheet/range/cell scopes) and need deterministic enforcement.

#### AI policy knobs (`ai.cloudProcessing`)

For cloud AI processing, the DLP policy rule (`DLP_ACTION.AI_CLOUD_PROCESSING` / `"ai.cloudProcessing"`) supports:

- `maxAllowed`:
  - the maximum classification level allowed to be sent to cloud AI (e.g. `"Confidential"`)
  - anything **over** this threshold becomes **REDACT** or **BLOCK**
- `redactDisallowed`:
  - when `true`, over-threshold content produces a **REDACT** decision (call proceeds, but content must be redacted)
  - when `false`, over-threshold content produces a **BLOCK** decision (do not call cloud AI)
- `allowRestrictedContent`:
  - when `true`, callers may opt into sending `"Restricted"` content by passing `dlp.includeRestrictedContent: true`
  - when `false`, `"Restricted"` content is **always blocked** for cloud AI, even if the caller sets `includeRestrictedContent`

`dlp.includeRestrictedContent` is therefore **not a bypass**; it is only honored when policy allows it.

### 2) Heuristic redaction — **defense-in-depth only**

`classifyText()` / `redactText()` are small regex-based helpers (e.g., emails/SSNs/credit cards). They:

- reduce accidental leakage
- help keep retrieval/indexing “safe by default”

When **structured DLP** is enabled (`buildContext({ dlp })` / `buildWorkbookContext({ dlp })`), `ContextManager` also uses
`classifyText()` as a **conservative heuristic input** to the policy engine:

- any heuristic “sensitive” finding is treated as a `Restricted` classification (`heuristic:*` labels)
- heuristic scanning covers both:
  - the bounded sheet/workbook text windows being prepared for prompt inclusion (including sheet names + schema metadata like tables/named ranges)
  - user-provided `attachments` (deep traversal, bounded)
- as defense-in-depth, heuristic checks also consider percent-encoded strings (e.g. `alice%40example.com`) to avoid leaks via encoded identifiers
- policy is evaluated on `max(structured, heuristic)` so rules can **BLOCK** or **REDACT** even when there are no
  structured classification records

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

### Structured classifications: record + selector shape (important!)

When you pass structured classifications via `dlp.classificationRecords` (or a `classificationStore`), each record is:

```ts
type ClassificationRecord = {
  selector: {
    scope: "document" | "sheet" | "range" | "column" | "cell";
    documentId: string;
    sheetId?: string;
    // range:
    range?: { start: { row: number; col: number }; end: { row: number; col: number } };
    // cell:
    row?: number;
    col?: number;
    // column:
    columnIndex?: number;
  };
  classification: { level: "Public" | "Internal" | "Confidential" | "Restricted"; labels?: string[] };
};
```

Key gotchas:

- **Row/col are 0-based**.
  - A1 `A1` is `{ row: 0, col: 0 }`.
  - Ranges are inclusive: `{ start: {row:0,col:0}, end: {row:9,col:3} }` covers the first 10 rows and 4 columns.
- `sheetId` should match whatever your app uses as the structured-DLP sheet identifier.
  - In many hosts this is the user-facing sheet name.
  - In some hosts (including parts of Formula desktop), there is a stable internal sheet id that can differ from the display name; in that case, provide a `sheetNameResolver` so DLP selectors still match.

Example `classificationRecords`:

```js
const classificationRecords = [
  {
    selector: { scope: "document", documentId: "doc-123" },
    classification: { level: "Confidential", labels: ["finance"] },
  },
  {
    selector: {
      scope: "range",
      documentId: "doc-123",
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 99, col: 3 } }, // A1:D100
    },
    classification: { level: "Restricted", labels: ["pii"] },
  },
  {
    selector: { scope: "cell", documentId: "doc-123", sheetId: "Sheet1", row: 1, col: 1 }, // B2
    classification: { level: "Restricted", labels: ["pii:ssn"] },
  },
];
```

If your environment has **stable sheet ids** that differ from display names, provide a resolver so DLP selectors still match:

```js
const sheetNameResolver = {
  // Called by ContextManager when it needs to map a user-facing sheet name back to a stable id.
  getSheetIdByName: (sheetName) => {
    return sheetNameToStableIdMap.get(sheetName) ?? sheetName;
  },
};

await cm.buildWorkbookContextFromSpreadsheetApi({
  spreadsheet,
  workbookId,
  query,
  dlp: { documentId: "doc-123", policy, classificationRecords, sheetNameResolver },
});
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
2. `attachment_data` (if present) — priority 4.5
3. `retrieved` — priority 4
4. `schema_summary` — priority 3.5
5. `schema` — priority 3
6. `attachments` — priority 2
7. `samples` — priority 1

Notes:

- `schema_summary` is a compact, deterministic, schema-first text summary (headers/types/counts) intended to survive tight budgets.
- `schema` is compact JSON intended for additional detail, but it excludes per-column sample values and is capped to a bounded
  number of tables/regions/columns to keep serialization + prompts predictable.

This means that when budgets shrink, **samples drop first**, then attachments, then schema JSON, etc.

For workbook context (`buildWorkbookContext*()`), the default section priorities are:

1. `dlp` (if present) — priority 5
2. `retrieved` — priority 4
3. `workbook_schema` — priority 3.5
4. `workbook_summary` — priority 3
5. `attachments` — priority 2

`workbook_schema` is a compact, schema-first summary (tables + inferred headers/types + named ranges) intended to
improve chat quality even when retrieval is sparse. When a workbook does not provide explicit tables/named ranges
(common for `buildWorkbookContextFromSpreadsheetApi`), `ContextManager` falls back to the already-indexed chunk
metadata in the vector store (data regions), extracting only the `COLUMNS:` line so no raw sample rows are included.

### Example: pack your own sections into a prompt budget

If you’re not using `ContextManager`, you can still use the token-budgeting primitives directly:

```js
import {
  createHeuristicTokenEstimator,
  packSectionsToTokenBudget,
  stableJsonStringify,
} from "./src/index.js";

const estimator = createHeuristicTokenEstimator();
const maxTokens = 8_000;

const sections = [
  { key: "schema", priority: 3, text: `Schema:\n${stableJsonStringify(schema)}` },
  { key: "retrieved", priority: 4, text: `Retrieved:\n${stableJsonStringify(retrieved)}` },
  { key: "samples", priority: 1, text: sampledRows.map((r) => JSON.stringify(r)).join("\n") },
].filter((s) => s.text);

const packed = packSectionsToTokenBudget(sections, maxTokens, estimator);
const promptContext = packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n");
```

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

### End-to-end example: budgeting a chat request

This is a common pattern:

1. pick a model context window (`contextWindowTokens`)
2. reserve space for output (`reserveForOutputTokens`)
3. subtract tool schema size (`estimateToolDefinitionTokens`)
4. build spreadsheet context within a fixed budget (`ContextManager.tokenBudgetTokens`)
5. trim messages to what remains (`trimMessagesToBudget`)

```js
import {
  ContextManager,
  createHeuristicTokenEstimator,
  estimateToolDefinitionTokens,
  trimMessagesToBudget,
} from "./src/index.js";

const estimator = createHeuristicTokenEstimator();

const contextWindowTokens = 128_000;
const reserveForOutputTokens = 4_000;

// Tool schemas can be surprisingly large; subtract them first.
const toolTokens = estimateToolDefinitionTokens(tools, estimator);
const maxMessageTokens = Math.max(0, contextWindowTokens - toolTokens);

// Use the *same* estimator for context + message trimming for consistent budgeting.
const cm = new ContextManager({
  tokenBudgetTokens: 20_000,
  tokenEstimator: estimator,
});

const { promptContext } = await cm.buildContext({ sheet, query });

const messages = [
  { role: "system", content: `${baseSystemPrompt}\n\n${promptContext}`.trim() },
  ...history,
  { role: "user", content: query },
];

const trimmedMessages = await trimMessagesToBudget({
  messages,
  maxTokens: maxMessageTokens,
  reserveForOutputTokens,
  estimator,
});

// Send `trimmedMessages` (+ `tools`) to your LLM client.
```

---

## Troubleshooting / gotchas

### `ContextManager.buildWorkbookContext requires workbookRag`

Workbook context builders (`buildWorkbookContext*`) need a `workbookRag` configuration in the `ContextManager`
constructor:

- `vectorStore` (implements `list`/`upsert`/`delete`/`query`)
- `embedder` (implements `embedTexts`)

See the workbook example above.

### “Vector dimension mismatch” errors

The embedder and vector store must agree on dimension:

```js
const embedder = new HashEmbedder({ dimension: 384 });
const vectorStore = new InMemoryVectorStore({ dimension: embedder.dimension });
```

### `buildWorkbookContextFromSpreadsheetApi` coordinate base

`buildWorkbookContextFromSpreadsheetApi()` assumes `SpreadsheetApi`-style **1-based** coordinates (`A1 => row=1, col=1`).

If your adapter uses **0-based** coordinates, build the workbook yourself with `packages/ai-rag` and call
`buildWorkbookContext()` instead:

```js
import { workbookFromSpreadsheetApi } from "../ai-rag/src/workbook/fromSpreadsheetApi.js";

const workbook = workbookFromSpreadsheetApi({
  spreadsheet,
  workbookId,
  coordinateBase: "zero",
});

const ctx = await cm.buildWorkbookContext({ workbook, query });
```

### DLP blocks (`DlpViolationError`)

If DLP blocks cloud AI processing, `ContextManager` will throw `DlpViolationError`. Catch it and **do not**
call a cloud model with blocked content.

Also note:

- `dlp.includeRestrictedContent` is only honored if policy allows it (`allowRestrictedContent: true`).
- If your sheet ids differ from display names, provide a `sheetNameResolver` so structured selectors match.
