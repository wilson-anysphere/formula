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
│  LOW AUTONOMY                                          HIGH AUTONOMY           │
│  ◄─────────────────────────────────────────────────────────────────────────►  │
│                                                                                │
│  Tab Complete    Inline Edit    Chat Panel    Composer    Full Agent          │
│       │              │              │            │            │               │
│  User types,    User selects   User asks    User        User sets           │
│  AI suggests    range, asks    questions,   describes   goal, AI           │
│  completions    for transform  AI answers   multi-step  executes           │
│                                             operation   autonomously        │
│                                                                                │
│  Latency:       Latency:       Latency:     Latency:    Latency:            │
│  <100ms         <2s            <5s          <30s        minutes             │
│                                                                                │
│  (All AI via Cursor servers - model selection managed by Cursor)            │
│                                                                                │
└────────────────────────────────────────────────────────────────────────────────┘
```

### Mode 1: Tab Completion

**Trigger:** User typing in formula bar or cell
**Latency requirement:** <100ms
**Backend:** Cursor servers with aggressive caching

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

User types: "Tot in B15
Suggestions:
  - "Total"            [Based on nearby text patterns]
  - =SUM(B1:B14)       [Detect "total" intent, suggest formula]
```

**Implementation:**

```typescript
class TabCompletionEngine {
  private cursorBackend: CursorAIClient;
  private cache: LRUCache<string, Suggestion[]>;
  
  async getSuggestions(context: CompletionContext): Promise<Suggestion[]> {
    // Check cache first
    const cacheKey = this.buildCacheKey(context);
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey)!;
    }
    
    // Parallel strategies
    const [formulaSuggestions, valueSuggestions, patternSuggestions] = await Promise.all([
      this.getFormulaSuggestions(context),
      this.getValueSuggestions(context),
      this.getPatternSuggestions(context)
    ]);
    
    // Rank and deduplicate
    const suggestions = this.rankSuggestions([
      ...formulaSuggestions,
      ...valueSuggestions,
      ...patternSuggestions
    ]);
    
    this.cache.set(cacheKey, suggestions);
    return suggestions.slice(0, 5);  // Top 5
  }
  
  private async getFormulaSuggestions(context: CompletionContext): Promise<Suggestion[]> {
    if (!context.currentInput.startsWith("=")) return [];
    
    // Parse partial formula
    const partialAST = this.parsePartial(context.currentInput);
    
    if (partialAST.inFunctionCall) {
      // Suggest arguments based on function signature
      return this.suggestFunctionArgs(partialAST.functionName, partialAST.argIndex, context);
    }
    
    if (partialAST.expectingRange) {
      // Suggest ranges based on data layout
      return this.suggestRanges(context);
    }
    
    // General formula completion via Cursor backend
    return this.cursorBackend.complete(context.currentInput, {
      maxTokens: 50,
      temperature: 0.1,
      stop: [")", ",", "\n"]
    });
  }
}
```

### Mode 2: Inline Edit (Cmd/Ctrl+K)

**Trigger:** User selects range and presses Cmd/Ctrl+K
**Latency requirement:** <2s for small operations
**Backend:** Cursor servers

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

**Implementation:**

```typescript
class InlineEditEngine {
  async processEdit(request: InlineEditRequest): Promise<InlineEditResponse> {
    // Classify intent
    const intent = await this.classifyIntent(request.prompt);
    
    switch (intent.type) {
      case "extract":
        return this.handleExtraction(request, intent);
      case "transform":
        return this.handleTransform(request, intent);
      case "generate":
        return this.handleGeneration(request, intent);
      case "format":
        return this.handleFormat(request, intent);
      case "formula":
        return this.handleFormulaGeneration(request, intent);
    }
  }
  
  private async handleExtraction(
    request: InlineEditRequest,
    intent: ExtractIntent
  ): Promise<InlineEditResponse> {
    // Sample data for LLM
    const sample = request.selectionContent.slice(0, 10);
    
    // Ask LLM to generate extraction pattern
    const prompt = `
Given these values:
${sample.map((row, i) => `${i + 1}. ${row[0]}`).join('\n')}

The user wants to: "${request.prompt}"

Generate a JavaScript function that extracts the desired value:
function extract(value) {
`;
    
    const code = await this.llm.complete(prompt);
    
    // Validate and sandbox the code
    const fn = this.sandboxedEval(code);
    
    // Apply to all values
    const changes: CellChange[] = [];
    for (let row = 0; row < request.selectionContent.length; row++) {
      const value = request.selectionContent[row][0];
      const extracted = fn(value);
      changes.push({
        row: request.selection.startRow + row,
        col: request.selection.endCol + 1,  // New column
        value: extracted
      });
    }
    
    return { type: "transform", changes, confidence: 0.9 };
  }
}
```

### Mode 3: Chat Panel

**Trigger:** User opens AI panel, asks questions
**Latency requirement:** <5s for most queries
**Backend:** Cursor servers

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

**Implementation:**

```typescript
// Helper that emits bounded, per-tool summaries (avoids huge `read_range` matrices).
// See: `packages/llm/src/toolResultSerialization.js`
import { serializeToolResultForModel } from "...";

class ChatPanelEngine {
  private conversation: ChatMessage[] = [];
  private tools: SpreadsheetTools;

  async processMessage(userMessage: string, attachments?: Attachment[]): Promise<ChatResponse> {
    // Add user message to conversation
    this.conversation.push({ role: "user", content: userMessage, attachments });

    // Build context
    const context = await this.buildContext(attachments);

    // Call LLM with tool use
    const response = await this.llm.chat({
      messages: [{ role: "system", content: this.buildSystemPrompt(context) }, ...this.conversation],
      tools: this.getToolDefinitions(),
      toolChoice: "auto",
    });

    // Process tool calls
    if (response.toolCalls) {
      const results = await this.executeToolCalls(response.toolCalls);

      // IMPORTANT: Tool results are part of the model's context on the next turn.
      // They should be appended as `role: "tool"` messages (matching the provider
      // tool-calling protocol), and must be DLP-redacted/blocked *at tool execution*
      // time (not only during prompt context construction).
      //
      // ALSO IMPORTANT: Never append full JSON tool results directly.
      // Tools like `read_range` can return huge matrices, which can blow up
      // model context windows and any persisted audit logs. Feed back a bounded,
      // per-tool summary instead (see `packages/llm/src/toolResultSerialization.js`).
      for (const { call, result } of results) {
        this.conversation.push({
          role: "tool",
          toolCallId: call.id,
          content: serializeToolResultForModel({ toolCall: call, result, maxChars: 20_000 }),
        });
      }

      // Get final response
      return this.processMessage("Please summarize what you did.");
    }

    // Add assistant response to conversation
    this.conversation.push({ role: "assistant", content: response.content });

    return {
      message: response.content,
      actions: response.suggestedActions,
    };
  }
}
```

### Mode 4: Cell Functions (=AI())

**Trigger:** User enters AI formula
**Latency requirement:** Async with loading state
**Backend:** Cursor servers

```
=AI("Summarize this feedback", A1:A100)
=AI.EXTRACT("email address", B1)
=AI.CLASSIFY("positive/negative/neutral", C1)
=AI.TRANSLATE("Spanish", D1)
```

**Implementation (desktop):**

Cell functions are implemented as an async cell-evaluator + cache:
- Evaluator: `apps/desktop/src/spreadsheet/evaluateFormula.ts`
  - Parses `AI()`, `AI.EXTRACT`, `AI.CLASSIFY`, `AI.TRANSLATE`
  - Preserves provenance for direct cell/range references passed to AI functions:
    - Cell ref: `{ __cellRef: "Sheet1!A1", value: ... }`
    - Range ref: array of provenance cell values, tagged with `__rangeRef` + `__totalCells`
  - Direct range references are sampled (default 200 cells) to avoid materializing unbounded arrays:
    - include a small deterministic prefix (top-left cells)
    - plus a deterministic seeded random sample from across the remainder of the range
- Engine: `apps/desktop/src/spreadsheet/AiCellFunctionEngine.ts`
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

Implementation notes:
- Orchestrator: `apps/desktop/src/ai/agent/agentOrchestrator.ts` (`runAgentTask`)
  - Builds workbook RAG context via `ContextManager.buildWorkbookContextFromSpreadsheetApi`
  - Runs the tool-calling loop via `runChatWithToolsAudited` with `mode: "agent"`
  - Emits progress events (`planning`, `tool_call`, `tool_result`, `assistant_message`, `complete` / `cancelled` / `error`)
- UI surface: `apps/desktop/src/panels/ai-chat/AIChatPanelContainer.tsx`
  - Agent tab (goal + constraints + run/cancel + live step log)
  - Preview-based approval gating via `PreviewEngine` + the shared approval modal
    - Default configuration requires explicit approval for any non-noop mutation (`approval_cell_threshold: 0`)
    - Optional "continue after deny" allows the model to re-plan when an approval is denied
  - Audit trail viewable in the AI Audit panel

**Capabilities:**
- Multi-step data gathering
- Complex analysis workflows
- Autonomous error correction
- Progress reporting

```typescript
interface AgentTask {
  goal: string;
  constraints?: string[];
  maxIterations?: number;
  requireApproval?: boolean;
}

interface AgentStep {
  thought: string;
  action: ToolCall;
  observation: string;
  status: "success" | "error" | "needs_approval";
}

class AgentEngine {
  async executeTask(task: AgentTask): Promise<AgentResult> {
    const steps: AgentStep[] = [];
    let iteration = 0;
    const maxIterations = task.maxIterations || 20;
    
    while (iteration < maxIterations) {
      iteration++;
      
      // Plan next action
      const plan = await this.planNextAction(task.goal, steps);
      
      if (plan.complete) {
        return { success: true, steps, finalResult: plan.result };
      }
      
      // Check if approval needed
      if (task.requireApproval && plan.action.requiresApproval) {
        const approved = await this.requestApproval(plan);
        if (!approved) {
          return { success: false, steps, reason: "User cancelled" };
        }
      }
      
      // Execute action
      const observation = await this.executeAction(plan.action);
      
      steps.push({
        thought: plan.thought,
        action: plan.action,
        observation: observation.result,
        status: observation.status
      });
      
      // Report progress
      this.reportProgress(steps);
      
      if (observation.status === "error") {
        // Try to recover
        const recovery = await this.planRecovery(steps, observation.error);
        if (!recovery.canRecover) {
          return { success: false, steps, reason: observation.error };
        }
      }
    }
    
    return { success: false, steps, reason: "Max iterations reached" };
  }
}
```

---

## Context Management

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

```typescript
class CellRAG {
  private vectorStore: VectorStore;
  private embedder: Embedder;
  
  async indexSheet(sheet: Sheet): Promise<void> {
    const chunks = this.chunkSheet(sheet);
    
    for (const chunk of chunks) {
      const text = this.chunkToText(chunk);
      const embedding = await this.embedder.embed(text);
      
      await this.vectorStore.add({
        id: `${sheet.id}-${chunk.range}`,
        embedding,
        metadata: {
          sheetId: sheet.id,
          range: chunk.range,
          preview: text.slice(0, 200)
        }
      });
    }
  }
  
  async findRelevant(query: string, topK: number = 5): Promise<RetrievedChunk[]> {
    const queryEmbedding = await this.embedder.embed(query);
    const results = await this.vectorStore.search(queryEmbedding, topK);
    
    return results.map(r => ({
      range: r.metadata.range,
      preview: r.metadata.preview,
      score: r.score
    }));
  }
  
  private chunkSheet(sheet: Sheet): Chunk[] {
    const chunks: Chunk[] = [];
    
    // Chunk by tables
    for (const table of sheet.tables) {
      chunks.push({
        type: "table",
        range: table.range,
        data: this.getTableData(table)
      });
    }
    
    // Chunk remaining data regions
    const regions = this.findDataRegions(sheet);
    for (const region of regions) {
      chunks.push({
        type: "region",
        range: region.range,
        data: this.getRegionData(region)
      });
    }
    
    return chunks;
  }
}
```

---

## Tool Calling

### Tool Definitions

```typescript
const SPREADSHEET_TOOLS: ToolDefinition[] = [
  {
    name: "read_range",
    description: "Read cell values from a range",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string", description: "Range in A1 notation (e.g., 'A1:D10')" },
        includeFormulas: { type: "boolean", default: false }
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
        cell: { type: "string", description: "Cell reference (e.g., 'A1')" },
        value: { type: "string", description: "Value or formula to write" }
      },
      required: ["cell", "value"]
    }
  },
  {
    name: "apply_formula_column",
    description: "Apply a formula pattern to an entire column",
    parameters: {
      type: "object",
      properties: {
        column: { type: "string", description: "Column letter" },
        formulaTemplate: { type: "string", description: "Formula with {row} placeholder" },
        startRow: { type: "number" },
        endRow: { type: "number", description: "-1 for last row with data" }
      },
      required: ["column", "formulaTemplate", "startRow"]
    }
  },
  {
    name: "create_chart",
    description: "Create a chart from data",
    parameters: {
      type: "object",
      properties: {
        chartType: { type: "string", enum: ["bar", "line", "pie", "scatter", "area"] },
        dataRange: { type: "string" },
        title: { type: "string" },
        position: { type: "string", description: "Where to place chart" }
      },
      required: ["chartType", "dataRange"]
    }
  },
  {
    name: "create_pivot_table",
    description: "Create a pivot table",
    parameters: {
      type: "object",
      properties: {
        sourceRange: { type: "string" },
        rows: { type: "array", items: { type: "string" } },
        columns: { type: "array", items: { type: "string" } },
        values: { 
          type: "array", 
          items: { 
            type: "object",
            properties: {
              field: { type: "string" },
              aggregation: { type: "string", enum: ["sum", "count", "average", "max", "min"] }
            }
          }
        },
        destination: { type: "string" }
      },
      required: ["sourceRange", "rows", "values"]
    }
  },
  {
    name: "sort_range",
    description: "Sort a range by one or more columns",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string" },
        sortBy: {
          type: "array",
          items: {
            type: "object",
            properties: {
              column: { type: "string" },
              order: { type: "string", enum: ["asc", "desc"] }
            }
          }
        }
      },
      required: ["range", "sortBy"]
    }
  },
  {
    name: "filter_range",
    description: "Filter a range based on criteria",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string" },
        criteria: {
          type: "array",
          items: {
            type: "object",
            properties: {
              column: { type: "string" },
              operator: { type: "string", enum: ["equals", "contains", "greater", "less", "between"] },
              value: { type: "string" },
              value2: { type: "string" }
            }
          }
        }
      },
      required: ["range", "criteria"]
    }
  },
  {
    name: "apply_formatting",
    description: "Apply formatting to a range",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string" },
        format: {
          type: "object",
          properties: {
            bold: { type: "boolean" },
            italic: { type: "boolean" },
            fontSize: { type: "number" },
            fontColor: { type: "string" },
            backgroundColor: { type: "string" },
            numberFormat: { type: "string" },
            horizontalAlign: { type: "string", enum: ["left", "center", "right"] }
          }
        }
      },
      required: ["range", "format"]
    }
  },
  {
    name: "detect_anomalies",
    description: "Find outliers and anomalies in data",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string" },
        method: { type: "string", enum: ["zscore", "iqr", "isolation_forest"] },
        threshold: { type: "number" }
      },
      required: ["range"]
    }
  },
  {
    name: "compute_statistics",
    description: "Compute statistical measures for a range",
    parameters: {
      type: "object",
      properties: {
        range: { type: "string" },
        measures: { 
          type: "array", 
          items: { 
            type: "string", 
            enum: ["mean", "median", "mode", "stdev", "variance", "min", "max", "quartiles", "correlation"] 
          }
        }
      },
      required: ["range"]
    }
  }
];
```

### Tool Executor

```typescript
class ToolExecutor {
  async execute(tool: ToolCall): Promise<ToolResult> {
    const startTime = performance.now();
    
    try {
      let result: any;
      
      switch (tool.name) {
        case "read_range":
          result = await this.readRange(tool.parameters);
          break;
        case "write_cell":
          result = await this.writeCell(tool.parameters);
          break;
        case "apply_formula_column":
          result = await this.applyFormulaColumn(tool.parameters);
          break;
        case "create_chart":
          result = await this.createChart(tool.parameters);
          break;
        case "create_pivot_table":
          result = await this.createPivotTable(tool.parameters);
          break;
        default:
          throw new Error(`Unknown tool: ${tool.name}`);
      }
      
      return {
        success: true,
        result,
        duration: performance.now() - startTime
      };
      
    } catch (error) {
      return {
        success: false,
        error: error.message,
        duration: performance.now() - startTime
      };
    }
  }
  
  private async applyFormulaColumn(params: ApplyFormulaParams): Promise<string> {
    const { column, formulaTemplate, startRow, endRow: endRowParam } = params;
    
    // Determine end row
    const endRow = endRowParam === -1 
      ? this.sheet.getLastRowWithData() 
      : endRowParam;
    
    // Apply formula to each row
    const colIndex = columnLetterToIndex(column);
    let count = 0;
    
    for (let row = startRow; row <= endRow; row++) {
      const formula = formulaTemplate.replace(/{row}/g, String(row));
      this.sheet.setCell(row, colIndex, formula);
      count++;
    }
    
    return `Applied formula to ${count} cells in column ${column}`;
  }
}
```

---

## Safety and Verification

### Preview Before Apply

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

### Client Integration

```typescript
class CursorAIClient {
  // All requests go through Cursor's authenticated backend
  // No API keys needed - Cursor handles authentication
  
  async complete(prompt: string, options: CompletionOptions): Promise<string> {
    const response = await this.cursorBackend.request({
      type: "completion",
      prompt,
      options
    });
    return response.completion;
  }
  
  async chat(messages: Message[], tools?: Tool[]): Promise<ChatResponse> {
    const response = await this.cursorBackend.request({
      type: "chat",
      messages,
      tools
    });
    return response;
  }
}
```

---

## Testing Strategy

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
