# Power Features

## Overview

Power features are what make Excel indispensable to professionals: Pivot Tables, Power Query, Data Model, What-If Analysis, and advanced statistical tools. We must match Excel's capabilities while adding AI-assisted workflows that make these features more accessible.

---

## Pivot Tables

Implementation ownership + cross-crate data flow is defined in
[ADR-0005: PivotTables ownership and data flow across crates](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md).
The TypeScript shapes below are conceptual (product/UX-oriented) and are not the
canonical persisted schema.

Implementation (actual):
- persisted schema + IDs: `formula-model`
- worksheet/range pivots + cache building: `formula-engine`
- Data Model pivots (measures/relationships): `formula-dax`
- XLSX import/export of pivot parts: `formula-xlsx`
- JS API surface: `formula-wasm`

See also:
- [`docs/21-xlsx-pivots.md`](./21-xlsx-pivots.md) — OpenXML pivot parts (pivotTables/pivotCache/slicers/timelines) compatibility + roadmap
- [`docs/21-dax-engine.md`](./21-dax-engine.md) — Data Model/DAX pivot execution details

### Architecture

```typescript
interface PivotTable {
  id: string;
  name: string;
  sourceRange: Range | TableRef | DataModelTable;
  destination: CellRef;
  
  // Field configuration
  rowFields: PivotField[];
  columnFields: PivotField[];
  valueFields: ValueField[];
  filterFields: FilterField[];
  
  // Layout options
  layout: "compact" | "outline" | "tabular";
  subtotals: "top" | "bottom" | "none";
  grandTotals: { rows: boolean; columns: boolean };
  
  // Calculated fields
  calculatedFields: CalculatedField[];
  calculatedItems: CalculatedItem[];
  
  // Caching
  cache: PivotCache;
}

interface PivotField {
  sourceField: string;
  name: string;
  sortOrder: "ascending" | "descending" | "manual";
  manualSort?: string[];
  grouping?: FieldGrouping;
  subtotal: AggregationType;
}

interface ValueField {
  sourceField: string;
  name: string;
  aggregation: AggregationType;
  numberFormat?: string;
  showAs?: ShowAsType;
  baseField?: string;
  baseItem?: string;
}

type AggregationType = 
  | "sum" | "count" | "average" | "max" | "min"
  | "product" | "countNumbers" | "stdDev" | "stdDevP"
  | "var" | "varP";

type ShowAsType =
  | "normal"
  | "percentOfGrandTotal"
  | "percentOfRowTotal"
  | "percentOfColumnTotal"
  | "percentOf"
  | "percentDifferenceFrom"
  | "runningTotal"
  | "rankAscending"
  | "rankDescending";
```

### Pivot Cache

```typescript
interface PivotCache {
  id: string;
  sourceType: "range" | "table" | "dataModel" | "external";
  source: any;
  
  // Cached data
  records: any[][];
  fields: CacheField[];
  
  // Unique values per field (for filtering UI)
  uniqueValues: Map<string, Set<any>>;
  
  // Last refresh
  lastRefreshed: Date;
  recordCount: number;
}

class PivotCacheManager {
  private caches: Map<string, PivotCache> = new Map();
  
  async createCache(source: PivotSource): Promise<PivotCache> {
    const cache: PivotCache = {
      id: crypto.randomUUID(),
      sourceType: source.type,
      source: source.reference,
      records: [],
      fields: [],
      uniqueValues: new Map(),
      lastRefreshed: new Date(),
      recordCount: 0
    };
    
    await this.refreshCache(cache);
    this.caches.set(cache.id, cache);
    
    return cache;
  }
  
  async refreshCache(cache: PivotCache): Promise<void> {
    const data = await this.loadSourceData(cache);
    
    // Extract fields from first row (headers)
    cache.fields = data.headers.map((name, index) => ({
      name,
      index,
      type: this.inferType(data.rows, index)
    }));
    
    // Store records
    cache.records = data.rows;
    cache.recordCount = data.rows.length;
    
    // Build unique values index
    for (const field of cache.fields) {
      const values = new Set<any>();
      for (const row of data.rows) {
        values.add(row[field.index]);
      }
      cache.uniqueValues.set(field.name, values);
    }
    
    cache.lastRefreshed = new Date();
  }
}
```

### Pivot Table Engine

```typescript
class PivotTableEngine {
  calculate(pivot: PivotTable): PivotResult {
    const cache = this.getCache(pivot.cache.id);
    
    // Build dimension keys
    const rowKeys = this.buildDimensionKeys(cache, pivot.rowFields);
    const colKeys = this.buildDimensionKeys(cache, pivot.columnFields);
    
    // Aggregate values
    const aggregations = new Map<string, Map<string, AggregateState>>();
    
    for (const record of cache.records) {
      // Apply filters
      if (!this.passesFilters(record, pivot.filterFields)) continue;
      
      const rowKey = this.getKey(record, pivot.rowFields);
      const colKey = this.getKey(record, pivot.columnFields);
      
      if (!aggregations.has(rowKey)) {
        aggregations.set(rowKey, new Map());
      }
      
      const rowAggs = aggregations.get(rowKey)!;
      if (!rowAggs.has(colKey)) {
        rowAggs.set(colKey, this.initAggregates(pivot.valueFields));
      }
      
      const state = rowAggs.get(colKey)!;
      this.updateAggregates(state, record, pivot.valueFields);
    }
    
    // Finalize aggregations
    const results = this.finalizeAggregates(aggregations, pivot);
    
    // Apply "Show As" calculations
    const transformed = this.applyShowAs(results, pivot);
    
    // Sort and generate output
    return this.generateOutput(transformed, rowKeys, colKeys, pivot);
  }
  
  private buildDimensionKeys(cache: PivotCache, fields: PivotField[]): string[][] {
    if (fields.length === 0) return [[""]];
    
    // Cartesian product of unique values
    let keys: string[][] = [[]];
    
    for (const field of fields) {
      const values = Array.from(cache.uniqueValues.get(field.sourceField) || []);
      const sorted = this.sortValues(values, field);
      
      const newKeys: string[][] = [];
      for (const key of keys) {
        for (const value of sorted) {
          newKeys.push([...key, String(value)]);
        }
      }
      keys = newKeys;
    }
    
    return keys;
  }
  
  private initAggregates(valueFields: ValueField[]): AggregateState {
    const state: AggregateState = {};
    
    for (const field of valueFields) {
      state[field.name] = {
        sum: 0,
        count: 0,
        min: Infinity,
        max: -Infinity,
        sumSquares: 0,
        values: []  // For median, mode, etc.
      };
    }
    
    return state;
  }
}
```

### AI-Assisted Pivot Creation

```typescript
class AIPivotAssistant {
  async suggestPivot(data: TableData, question: string): Promise<PivotSuggestion> {
    const schema = this.extractSchema(data);
    
    const prompt = `
Given this data schema:
${JSON.stringify(schema, null, 2)}

User question: "${question}"

Suggest a pivot table configuration that answers this question.
Return JSON with:
- rowFields: array of field names for rows
- columnFields: array of field names for columns
- valueFields: array of {field, aggregation} for values
- explanation: why this configuration answers the question
`;
    
    const response = await this.llm.complete(prompt);
    return JSON.parse(response);
  }
  
  async createPivotFromNaturalLanguage(
    data: TableData,
    request: string
  ): Promise<PivotTable> {
    const suggestion = await this.suggestPivot(data, request);
    
    // Validate suggestion against actual data
    const validated = this.validateSuggestion(suggestion, data);
    
    // Create pivot table
    return this.pivotManager.create({
      source: data.range,
      rowFields: validated.rowFields.map(f => ({ sourceField: f })),
      columnFields: validated.columnFields.map(f => ({ sourceField: f })),
      valueFields: validated.valueFields.map(v => ({
        sourceField: v.field,
        aggregation: v.aggregation,
        name: `${v.aggregation} of ${v.field}`
      }))
    });
  }
}
```

---

## Power Query Equivalent

### Query Model

```typescript
interface Query {
  id: string;
  name: string;
  source: QuerySource;
  steps: QueryStep[];
  destination?: Range | TableRef;
  refreshPolicy: RefreshPolicy;
}

type QuerySource =
  | { type: "range"; range: Range }
  | { type: "table"; table: string }
  | { type: "csv"; path: string; options: CSVOptions }
  | { type: "json"; path: string; jsonPath?: string }
  | { type: "database"; connection: DBConnection; query: string }
  | { type: "api"; url: string; method: string; headers?: Record<string, string> }
  | { type: "query"; queryId: string };  // Reference another query

interface QueryStep {
  id: string;
  name: string;
  operation: QueryOperation;
  resultSchema?: Schema;
}

type QueryOperation =
  | { type: "selectColumns"; columns: string[] }
  | { type: "removeColumns"; columns: string[] }
  | { type: "filterRows"; predicate: FilterPredicate }
  | { type: "sortRows"; sortBy: SortSpec[] }
  | { type: "groupBy"; groupColumns: string[]; aggregations: Aggregation[] }
  | { type: "addColumn"; name: string; formula: string }
  | { type: "renameColumn"; oldName: string; newName: string }
  | { type: "changeType"; column: string; newType: DataType }
  | { type: "pivot"; rowColumn: string; valueColumn: string; aggregation: string }
  | { type: "unpivot"; columns: string[]; nameColumn: string; valueColumn: string }
  | { type: "merge"; rightQuery: string; joinType: JoinType; leftKey: string; rightKey: string }
  | { type: "append"; queries: string[] }
  | { type: "fillDown"; columns: string[] }
  | { type: "replaceValues"; column: string; find: any; replace: any }
  | { type: "splitColumn"; column: string; delimiter: string };
```

### Refresh Orchestration ("Refresh All")

Excel's "Refresh All" respects query dependencies: referenced queries must refresh before dependents, and shared upstream dependencies should run at most once.

In Formula, this is handled by a dependency-aware refresh orchestrator that:

- Extracts a query dependency graph (`source.type === "query"`, plus `merge` + `append` dependencies)
- Executes dependencies before dependents (topological ordering)
- Shares a single execution session across the whole refresh so credentials/permissions are cached and prompts are minimized
- Supports concurrency for independent subgraphs and cancellation/progress reporting per query
- Cancels only downstream dependents on failure, while allowing independent subgraphs to continue (Excel-like Refresh All UX)

```js
import { QueryEngine, RefreshOrchestrator } from "@formula/power-query";

const engine = new QueryEngine({
  onCredentialRequest: async (connectorId, details) => {
    // prompt once per session per unique request
    return { token: "..." };
  },
  onPermissionRequest: async (kind, details) => true,
});

const orchestrator = new RefreshOrchestrator({
  engine,
  getContext: () => ({ queries: /* map of id -> Query */, tables: /* optional */ }),
  concurrency: 2,
});

// Register queries that can be refreshed.
for (const query of queries) orchestrator.registerQuery(query);

// Optional: subscribe to queued/started/progress/completed/error/cancelled events.
const unsubscribe = orchestrator.onEvent((evt) => {
  // evt includes { sessionId, phase: "dependency" | "target" }
});

const handle = orchestrator.refreshAll(); // refresh all registered queries
await handle.promise;
unsubscribe();
```

Notes:
- Event `job.id` values are namespaced with the refresh session (`${sessionId}:...`) so they are safe to treat as globally unique across concurrent refreshAll runs.
- `handle.cancel()` aborts the entire session. Some hosts may also use `handle.cancelQuery(queryId)` to abort a single query branch (cancelling downstream dependents while allowing independent targets to continue).
- Hosts can also use `orchestrator.triggerOnOpen()` to refresh all registered queries with `refreshPolicy: { type: "on-open" }` using the same dependency-aware semantics.
- For "refresh this one query (plus its dependencies)" UX, `orchestrator.refresh(queryId)` is a convenience wrapper around `refreshAll([queryId])` that returns a single-query promise.

### Query Result Caching

Power Query refreshes are often repeatable (same source + same steps). Formula's `QueryEngine` supports caching results to make repeated refreshes fast, including for large columnar datasets (Parquet/Arrow pipelines).

Key points:

- Cache keys incorporate the full query signature:
  - source request (including stable credential identity)
  - executed step list
  - execution options (e.g. `limit`)
- Cache entries are validated by default using **source state** (mtime/ETag) when the connector supports it, preventing stale results.
- Cache payloads support both row-backed and columnar results:
  - **DataTable**: JSON-friendly rows, with tagged cells to preserve Dates, BigInt, NaN/±Infinity, and Uint8Array values.
  - **ArrowTableAdapter**: stored as **Arrow IPC stream bytes** (plus Power Query column metadata) so Parquet-backed refreshes can be restored without re-reading or re-decoding the Parquet file.

Cache stores:

- `MemoryCacheStore`: in-memory (fastest; non-persistent)
- `IndexedDBCacheStore`: browser persistence (stores `Uint8Array` IPC bytes via structured clone)
- `FileSystemCacheStore`: Node persistence (`<hash>.json` + `<hash>.bin` for Arrow IPC bytes)
- `EncryptedCacheStore`: wrapper that encrypts cached values at rest using a host-provided AES-256-GCM provider (supports wrapping IndexedDB or filesystem stores; keeps ciphertext as `Uint8Array` without base64)
- `EncryptedFileSystemCacheStore`: Node persistence with encryption-at-rest (AES-256-GCM) and the same `.bin` blob strategy for Arrow IPC bytes

Provider helpers:

- `createWebCryptoCacheProvider({ keyVersion, keyBytes })`: create an AES-256-GCM provider backed by `crypto.subtle` (browser / WebView contexts)
- `createNodeCryptoCacheProvider({ keyVersion, keyBytes })`: create an AES-256-GCM provider backed by `node:crypto` (Node contexts; exported from `@formula/power-query/node`)

Desktop integration notes:

- Formula Desktop (Tauri) uses an encrypted IndexedDB Power Query cache by default.
  - Cached query results are wrapped in `EncryptedCacheStore`.
  - The AES-256-GCM key is generated once and stored in the OS keychain (via a Tauri command), so cached results remain decryptable across restarts.
  - Older plaintext cache databases are best-effort cleared/deleted on startup to avoid leaving behind readable cache data.

Note: Node-only helpers are available from the `@formula/power-query/node` entrypoint (for example `EncryptedFileSystemCacheStore` and `createNodeCryptoCacheProvider`).

Maintenance:

- `CacheManager.get()` evicts expired entries opportunistically on access.
- For long-lived persistent stores (IndexedDB / filesystem), hosts may also call
  `cache.pruneExpired()` periodically to proactively reclaim space. This is
  best-effort and never throws.

Example (in-memory caching):

```js
import { readFile, stat } from "node:fs/promises";

import {
  QueryEngine,
  CacheManager,
  MemoryCacheStore,
} from "@formula/power-query";

const cache = new CacheManager({ store: new MemoryCacheStore() });

const engine = new QueryEngine({
  cache,
  defaultCacheTtlMs: 60_000,
  fileAdapter: {
    // Provide stat() to enable source-state validation for file sources.
    stat: async (path) => ({ mtimeMs: (await stat(path)).mtimeMs }),
    readText: async (path) => readFile(path, "utf8"),
    readBinary: async (path) => {
      const bytes = await readFile(path);
      return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    },
  },
});

const { table, meta } = await engine.executeQueryWithMeta(query, context, {
  cache: { mode: "use", validation: "source-state" },
});
```

### Query Folding

Push operations to the data source when possible:

```typescript
class QueryFoldingEngine {
  // Operations that can be folded to SQL
  private foldableToSQL: Set<string> = new Set([
    "selectColumns", "removeColumns", "filterRows", "sortRows",
    "groupBy", "addColumn", "renameColumn", "merge"
  ]);
  
  // Operations that break folding
  private foldingBreakers: Set<string> = new Set([
    "addIndexColumn", "fillDown", "customFunction"
  ]);
  
  compile(query: Query): CompiledQuery {
    if (query.source.type !== "database") {
      // Non-SQL source - execute all steps locally
      return { type: "local", steps: query.steps };
    }
    
    const sqlParts: SQLPart[] = [];
    const localSteps: QueryStep[] = [];
    let foldingBroken = false;
    
    for (const step of query.steps) {
      if (foldingBroken || this.foldingBreakers.has(step.operation.type)) {
        foldingBroken = true;
        localSteps.push(step);
      } else if (this.foldableToSQL.has(step.operation.type)) {
        const sql = this.operationToSQL(step.operation);
        if (sql) {
          sqlParts.push(sql);
        } else {
          foldingBroken = true;
          localSteps.push(step);
        }
      } else {
        foldingBroken = true;
        localSteps.push(step);
      }
    }
    
    return {
      type: "hybrid",
      sql: this.combineSQLParts(sqlParts, query.source),
      localSteps
    };
  }
  
  private operationToSQL(op: QueryOperation): SQLPart | null {
    switch (op.type) {
      case "selectColumns":
        return { type: "select", columns: op.columns };
        
      case "filterRows":
        return { type: "where", predicate: this.predicateToSQL(op.predicate) };
        
      case "sortRows":
        return { type: "orderBy", specs: op.sortBy };
        
      case "groupBy":
        return {
          type: "groupBy",
          columns: op.groupColumns,
          aggregations: op.aggregations.map(a => this.aggregationToSQL(a))
        };
        
      case "merge":
        return {
          type: "join",
          joinType: op.joinType,
          rightTable: this.getQueryTable(op.rightQuery),
          leftKey: op.leftKey,
          rightKey: op.rightKey
        };
        
      default:
        return null;
    }
  }
}
```

### Privacy Levels + Formula Firewall

Excel/Power Query has a privacy model ("Public" / "Organizational" / "Private") enforced by the **formula firewall**. The goal is to prevent *implicit* data leakage when queries combine sources, especially when query folding/pushdown would send data from one privacy domain into another external system.

In Formula, the Power Query engine supports the same concept:

- **Stable source identities (`sourceId`)** are derived per source:
  - file: normalized path
  - web: origin (`scheme://host:port`)
  - database: `sql:<connectionId>` (from `source.connectionId` / `SqlConnector.getConnectionIdentity`)
  - workbook: `workbook:range`, `workbook:table:<name>`
- Hosts can classify sources by supplying privacy levels in the execution context:
  ```js
  const context = {
    privacy: {
      levelsBySourceId: {
        "sql:db-prod": "organizational",
        "https://api.example.com:443": "public",
        "workbook:table:Sales": "private",
      },
    },
  };
  ```
- The engine can be configured with a privacy enforcement mode:
  ```js
  const engine = new QueryEngine({
    privacyMode: "ignore" | "warn" | "enforce", // default: "ignore"
  });
  ```

Behavior summary:
- **Folding restrictions (safe by default)**: when `privacyMode !== "ignore"`, folding of `merge`/`append` across different privacy levels is prevented (downgrades to hybrid/local) and a structured diagnostic is emitted.
- **Local combination firewall**: in `"enforce"` mode, obvious high→low combinations (e.g. Private/Unknown + Public) are blocked with a `Formula.Firewall` error; in `"warn"` mode they run but emit a warning diagnostic.

Diagnostics are surfaced via the engine `onProgress` callback as `type: "privacy:firewall"` events with the involved source ids + levels.

### Query Editor UI

```typescript
interface QueryEditorState {
  query: Query;
  preview: PreviewData;
  selectedStep: number;
  schema: Schema;
}

class QueryEditor {
  private state: QueryEditorState;
  
  addStep(operation: QueryOperation): void {
    const step: QueryStep = {
      id: crypto.randomUUID(),
      name: this.generateStepName(operation),
      operation
    };
    
    // Insert after selected step
    this.state.query.steps.splice(this.state.selectedStep + 1, 0, step);
    this.state.selectedStep++;
    
    // Update preview
    this.refreshPreview();
  }
  
  async refreshPreview(): Promise<void> {
    // Execute query up to selected step
    const stepsToExecute = this.state.query.steps.slice(0, this.state.selectedStep + 1);
    
    const result = await this.executeSteps(
      this.state.query.source,
      stepsToExecute,
      { limit: 100 }  // Preview limit
    );
    
    this.state.preview = result;
    this.state.schema = this.inferSchema(result);
  }
  
  // AI assistance
  async suggestNextStep(intent: string): Promise<QueryOperation[]> {
    const context = {
      currentSchema: this.state.schema,
      sampleData: this.state.preview.rows.slice(0, 10),
      existingSteps: this.state.query.steps
    };
    
    return this.aiAssistant.suggestQuerySteps(context, intent);
  }
}
```

---

## Data Model (DAX)

### Table Relationships

```typescript
interface DataModel {
  tables: DataModelTable[];
  relationships: Relationship[];
  measures: Measure[];
  calculatedColumns: CalculatedColumn[];
}

interface DataModelTable {
  id: string;
  name: string;
  source: TableSource;
  columns: Column[];
  isHidden: boolean;
}

interface Relationship {
  id: string;
  name: string;
  fromTable: string;
  fromColumn: string;
  toTable: string;
  toColumn: string;
  cardinality: "oneToOne" | "oneToMany" | "manyToMany";
  crossFilterDirection: "single" | "both";
  isActive: boolean;
}

interface Measure {
  id: string;
  name: string;
  table: string;  // Home table
  expression: string;  // DAX expression
  formatString?: string;
  description?: string;
}

interface CalculatedColumn {
  id: string;
  name: string;
  table: string;
  expression: string;  // DAX expression
  dataType: DataType;
}
```

### DAX Evaluation Engine

```typescript
class DAXEngine {
  // Row context: the "current row" during iteration
  // Filter context: the filters applied before evaluation
  
  evaluate(
    expression: string,
    filterContext: FilterContext,
    rowContext?: RowContext
  ): any {
    const ast = this.parseDAX(expression);
    return this.evaluateAST(ast, filterContext, rowContext);
  }
  
  private evaluateAST(
    node: DAXNode,
    filterContext: FilterContext,
    rowContext?: RowContext
  ): any {
    switch (node.type) {
      case "measure":
        return this.evaluateMeasure(node.name, filterContext);
        
      case "column":
        if (rowContext) {
          return rowContext.getValue(node.table, node.column);
        } else {
          // Without row context, return column values as table
          return this.getColumnValues(node.table, node.column, filterContext);
        }
        
      case "aggregation":
        return this.evaluateAggregation(node, filterContext);
        
      case "calculate":
        // CALCULATE: modify filter context, then evaluate
        const newFilterContext = this.modifyFilterContext(
          filterContext,
          node.filterArgs
        );
        return this.evaluateAST(node.expression, newFilterContext, rowContext);
        
      case "iterator":
        // Iterators like SUMX, AVERAGEX create row context
        return this.evaluateIterator(node, filterContext);
        
      case "related":
        // RELATED: navigate relationship to get related value
        return this.evaluateRelated(node, rowContext);
        
      case "relatedTable":
        // RELATEDTABLE: get related rows from other side of relationship
        return this.evaluateRelatedTable(node, filterContext, rowContext);
    }
  }
  
  private evaluateIterator(
    node: IteratorNode,
    filterContext: FilterContext
  ): number {
    const table = this.evaluateTable(node.table, filterContext);
    
    let accumulator: number;
    switch (node.iteratorType) {
      case "SUMX": accumulator = 0; break;
      case "AVERAGEX": accumulator = 0; break;
      case "MAXX": accumulator = -Infinity; break;
      case "MINX": accumulator = Infinity; break;
      case "COUNTX": accumulator = 0; break;
    }
    
    let count = 0;
    for (const row of table.rows) {
      // Create row context for this row
      const rowContext = new RowContext(table, row);
      
      // Evaluate expression in row context
      const value = this.evaluateAST(node.expression, filterContext, rowContext);
      
      // Accumulate
      switch (node.iteratorType) {
        case "SUMX": accumulator += value || 0; break;
        case "AVERAGEX": accumulator += value || 0; count++; break;
        case "MAXX": accumulator = Math.max(accumulator, value); break;
        case "MINX": accumulator = Math.min(accumulator, value); break;
        case "COUNTX": if (value !== null && value !== undefined) count++; break;
      }
    }
    
    if (node.iteratorType === "AVERAGEX") {
      return count > 0 ? accumulator / count : 0;
    }
    if (node.iteratorType === "COUNTX") {
      return count;
    }
    
    return accumulator;
  }
  
  // Context transition: CALCULATE converts row context to filter context
  private contextTransition(rowContext: RowContext): FilterContext {
    const filters: Filter[] = [];
    
    for (const [table, columns] of rowContext.tables) {
      for (const [column, value] of columns) {
        filters.push({
          table,
          column,
          operator: "equals",
          value
        });
      }
    }
    
    return new FilterContext(filters);
  }
}
```

---

## What-If Analysis

The Rust implementations live in:

- Goal Seek: [`what_if/goal_seek.rs`](../crates/formula-engine/src/what_if/goal_seek.rs)
- Scenario Manager: [`what_if/scenario_manager.rs`](../crates/formula-engine/src/what_if/scenario_manager.rs)
- Monte Carlo: [`what_if/monte_carlo.rs`](../crates/formula-engine/src/what_if/monte_carlo.rs)

and are exposed as `formula_engine::what_if`.

All What‑If tools are written against the `what_if::WhatIfModel` trait (get/set cells + `recalculate()`), with an adapter for the real calc engine: `what_if::EngineWhatIfModel` (`what_if/engine_model.rs`).

### Common types (used by Goal Seek / Scenarios / Monte Carlo)

Rust:

- `what_if::CellRef` — a `#[serde(transparent)]` A1-style cell reference string. `EngineWhatIfModel` accepts `A1` (uses `default_sheet`) or `Sheet1!A1` / `\'My Sheet\'!A1`.
- `what_if::CellValue` — scalar-only values: `Number(f64)`, `Text(String)`, `Bool(bool)`, `Blank`.
- `what_if::WhatIfError<E>` — returned for invalid parameters, non-numeric cells, or underlying model failures.

Proposed JS/WASM DTOs (directly compatible with the current serde shapes):

```ts
export type CellRef = string;

// Mirrors: what_if::CellValue  (tagged enum; snake_case tags)
export type CellValue =
  | { type: "number"; value: number }
  | { type: "text"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "blank" };
```

Error surface (host contract):

- Rust functions return `Result<T, WhatIfError<_>>`. WASM bindings should throw a JS `Error` whose message is `WhatIfError::to_string()` (e.g. `"invalid parameters: iterations must be > 0"`).

### Goal Seek

Rust API ([`what_if/goal_seek.rs`](../crates/formula-engine/src/what_if/goal_seek.rs)):

- `GoalSeekParams` (`#[serde(rename_all = "camelCase")]`) — target/changing cells plus numeric tuning knobs.
- `GoalSeekResult` (`#[serde(rename_all = "camelCase")]`) — final solution + diagnostics.
- `GoalSeekProgress` (`#[serde(rename_all = "camelCase")]`) — progress event emitted once at iteration 0 and after every Newton/bisection step.
- `GoalSeekStatus` — `{ Converged, MaxIterationsReached, NoBracketFound, NumericalFailure }`.
- `GoalSeek::{solve, solve_with_progress}` — synchronous solver (Newton step + finite-difference derivative; falls back to bisection if derivative is too small or non-finite).

Proposed JS/WASM DTOs:

```ts
export interface GoalSeekParams {
  targetCell: CellRef;
  targetValue: number;
  changingCell: CellRef;

  // Optional in JS; bindings should fill Rust defaults from GoalSeekParams::new()
  maxIterations?: number; // default 100
  tolerance?: number; // default 0.001
  derivativeStep?: number | null; // null/undefined => auto (abs(x)*0.001 or 0.001)
  minDerivative?: number; // default 1e-10
  maxBracketExpansions?: number; // default 50
}

// Mirrors Rust GoalSeekStatus serialization today (PascalCase variant strings).
export type GoalSeekStatus =
  | "Converged"
  | "MaxIterationsReached"
  | "NoBracketFound"
  | "NumericalFailure";

export interface GoalSeekResult {
  status: GoalSeekStatus;
  solution: number;
  iterations: number;
  finalOutput: number;
  finalError: number; // finalOutput - targetValue
}

export interface GoalSeekProgress {
  iteration: number;
  input: number;
  output: number;
  error: number;
}

// Suggested binding shape (sync in Rust; host decides whether to run in a worker):
//   workbook.goalSeek(params, { defaultSheet?: string, onProgress?: (p: GoalSeekProgress) => void }): GoalSeekResult
```

Validation + edge cases (Rust behavior):

- `maxIterations` must be `> 0` → otherwise `WhatIfError::InvalidParams("max_iterations must be > 0")`.
- `tolerance` must be `> 0` → otherwise `WhatIfError::InvalidParams("tolerance must be > 0")`.
- `minDerivative` must be `> 0` → otherwise `WhatIfError::InvalidParams("min_derivative must be > 0")`.
- `changingCell` and `targetCell` must evaluate to `CellValue::Number(..)` → otherwise `WhatIfError::NonNumericCell { cell, value }`.
- If the starting state already satisfies the target (`|output-targetValue| < tolerance`) the solver returns `status: "Converged"` with `iterations: 0`.
- If Newton’s method produces a tiny/non-finite derivative, the solver switches to bisection:
  - If a sign-changing bracket cannot be found within `maxBracketExpansions`, the call succeeds but returns `status: "NoBracketFound"`.
- If a Newton step would produce a non-finite next input (NaN/±Inf), the call succeeds but returns `status: "NumericalFailure"`.

### Scenario Manager

Rust API ([`what_if/scenario_manager.rs`](../crates/formula-engine/src/what_if/scenario_manager.rs)):

- `ScenarioManager` — in-memory store of named scenarios + a “base” snapshot used for restore.
- `ScenarioId(u64)` (`#[serde(transparent)]`) — scenario identifier.
- `Scenario` (`#[serde(rename_all = "camelCase")]`) — scenario metadata + `{ CellRef -> CellValue }` map.
- `SummaryReport` (`#[serde(rename_all = "camelCase")]`) — `"Base"` row + per-scenario outputs.

Key methods:

- `create_scenario(name, changing_cells, values, created_by, comment) -> Result<ScenarioId, WhatIfError<_>>`
- `delete_scenario(id) -> bool`
- `apply_scenario(model, id) -> Result<(), WhatIfError<_>>`
- `restore_base(model) -> Result<(), WhatIfError<_>>`
- `generate_summary_report(model, result_cells, scenario_ids) -> Result<SummaryReport, WhatIfError<_>>`

Proposed JS/WASM DTOs:

```ts
// Note: Rust uses u64; WASM bindings should require these to be safe integers.
export type ScenarioId = number;

// Matches ScenarioManager::create_scenario inputs (parallel arrays).
export interface CreateScenarioParams {
  name: string;
  changingCells: CellRef[];
  values: CellValue[]; // must match changingCells length
  createdBy: string;
  comment?: string | null;
}

export interface Scenario {
  id: ScenarioId;
  name: string;
  changingCells: CellRef[];
  values: Record<CellRef, CellValue>;
  createdMs: number; // ms since Unix epoch
  createdBy: string;
  comment?: string | null;
}

export interface SummaryReport {
  changingCells: CellRef[];
  resultCells: CellRef[];
  // scenarioName -> (cell -> value). Always includes "Base".
  results: Record<string, Record<CellRef, CellValue>>;
}

// Suggested binding shape:
//   const mgr = workbook.createScenarioManager();
//   mgr.createScenario(params: CreateScenarioParams) -> ScenarioId
//   mgr.applyScenario(id) / mgr.restoreBase()
//   mgr.generateSummaryReport({ resultCells, scenarioIds }) -> SummaryReport
```

Validation + edge cases (Rust behavior):

- `create_scenario`: `changing_cells.len()` must equal `values.len()` → `WhatIfError::InvalidParams("changing_cells and values must have equal length")`.
- Scenario names are not required to be unique. **If two scenarios share a name**, `generate_summary_report` will overwrite the earlier entry in `results` (it’s a `HashMap<String, ...>` keyed by name).
- `apply_scenario` / `generate_summary_report` with an unknown `ScenarioId` → `WhatIfError::InvalidParams("scenario not found")`.
- `restore_base` is a no-op if no scenario has been applied yet (`base_values` empty).
- Base snapshot semantics: applying multiple scenarios captures the **union** of their `changing_cells` in `base_values` so `restore_base` can fully return to the original state even if scenarios touch different inputs.
- `restore_base` does **not** clear `base_values`; hosts can call `clear_base_values()` to reset the snapshot explicitly.

### Monte Carlo Simulation

Rust API ([`what_if/monte_carlo.rs`](../crates/formula-engine/src/what_if/monte_carlo.rs)):

- `SimulationConfig` (`#[serde(rename_all = "camelCase")]`), `InputDistribution`, `Distribution`, `CorrelationMatrix`
- `MonteCarloEngine::{run_simulation, run_simulation_with_progress}`
- `SimulationProgress` (`completedIterations`, `totalIterations`)
- `SimulationResult` (`outputStats`, `outputSamples`)
- `OutputStatistics` (`mean`, `median`, `stdDev`, `percentiles`, `histogram`, …)

Proposed JS/WASM DTOs (field names match Rust’s serde output):

```ts
export interface SimulationConfig {
  iterations: number;
  inputDistributions: InputDistribution[];
  outputCells: CellRef[];

  // Optional in JS; Rust defaults: seed=0, histogramBins=50
  seed?: number; // u64; require a safe integer to preserve determinism
  correlations?: CorrelationMatrix | null;
  histogramBins?: number;
}

export interface InputDistribution {
  cell: CellRef;
  distribution: Distribution;
}

// Mirrors Rust: #[serde(tag = "type", rename_all = "snake_case")]
export type Distribution =
  | { type: "normal"; mean: number; stdDev: number }
  | { type: "uniform"; min: number; max: number }
  | { type: "triangular"; min: number; mode: number; max: number }
  | { type: "lognormal"; mean: number; stdDev: number }
  | { type: "discrete"; values: number[]; probabilities: number[] }
  | { type: "beta"; alpha: number; beta: number; min?: number | null; max?: number | null }
  | { type: "exponential"; rate: number }
  | { type: "poisson"; lambda: number };

export interface CorrelationMatrix {
  matrix: number[][];
}

export interface SimulationProgress {
  completedIterations: number;
  totalIterations: number;
}

export interface HistogramBin {
  start: number;
  end: number;
  count: number;
}

export interface Histogram {
  bins: HistogramBin[];
}

export interface OutputStatistics {
  mean: number;
  median: number;
  stdDev: number; // sample std dev (n-1); 0 when n<=1
  min: number;
  max: number;
  // Keys are the fixed percentiles computed by Rust: 5,10,25,75,90,95.
  // (Serialized as an object; keys will be strings in JS.)
  percentiles: Record<string, number>;
  histogram: Histogram;
}

export interface SimulationResult {
  iterations: number;
  outputStats: Record<CellRef, OutputStatistics>;
  // Raw samples for charting: output cell -> length=iterations values.
  outputSamples: Record<CellRef, number[]>;
}

// Suggested binding shape:
//   workbook.runMonteCarloSimulation(config, { defaultSheet?: string, onProgress?: (p: SimulationProgress) => void }): SimulationResult
```

Validation + edge cases (Rust behavior):

- `iterations` must be `> 0` → `WhatIfError::InvalidParams("iterations must be > 0")`.
- `histogramBins` must be `> 0` → `WhatIfError::InvalidParams("histogram_bins must be > 0")`.
- `outputCells` must be non-empty → `WhatIfError::InvalidParams("output_cells must not be empty")`.
- Every `Distribution` is validated up-front:
  - normal/lognormal: `stdDev >= 0`
  - uniform: `min <= max`
  - triangular: `min <= mode <= max`
  - discrete: `values.len() > 0`, equal lengths, all `probabilities >= 0`, and `sum(probabilities) > 0` (does not need to equal 1)
  - beta: `alpha > 0 && beta > 0`, and if both `min/max` set then `min <= max`
  - exponential: `rate > 0`
  - poisson: `lambda >= 0`
- Output cells must evaluate to numbers each iteration; otherwise `WhatIfError::NonNumericCell { cell, value }`.
- Correlations:
  - `correlations.matrix` must be square, symmetric, diagonal=1, entries in `[-1, 1]`, and **positive definite** (Cholesky decomposition).
  - Correlated sampling is currently supported **only** when *all* input distributions are `{ type: "normal", ... }` → otherwise `InvalidParams("correlated sampling is currently supported only for normal distributions")`.
- Histogram edge cases:
  - If all samples are identical (`min == max`), Rust returns a single bin with `count = iterations`.
  - If samples are empty or min/max are non-finite (shouldn’t happen for valid runs), histogram bins are empty.

---

## Solver (Optimization)

The Rust implementation lives in [`solver/mod.rs`](../crates/formula-engine/src/solver/mod.rs) (and supporting files under `crates/formula-engine/src/solver/*`) and is exposed as `formula_engine::solver`.

This is a small-but-functional Excel-like Solver with three methods:

- **Simplex** (`SolveMethod::Simplex`) — linear programming (LP) with optional integer/binary variables (branch-and-bound).
- **GRG Nonlinear** (`SolveMethod::GrgNonlinear`) — penalty-based gradient method (continuous variables only).
- **Evolutionary** (`SolveMethod::Evolutionary`) — genetic algorithm (supports integer/binary; suitable for non-smooth problems).

### Rust API surface

Core types ([`solver/mod.rs`](../crates/formula-engine/src/solver/mod.rs)):

- `SolverModel` trait — model abstraction (`get_vars`/`set_vars`/`recalc`/`objective`/`constraints`).
- `EngineSolverModel` ([`solver/engine_model.rs`](../crates/formula-engine/src/solver/engine_model.rs)) — adapter that binds the solver to `formula_engine::Engine` cell references.
- `SolverProblem` — `{ objective: Objective, variables: Vec<VarSpec>, constraints: Vec<Constraint> }`
- `Objective` / `ObjectiveKind` — maximize/minimize/target (with `targetValue` + `targetTolerance`)
- `VarSpec` / `VarType` — bounds + variable domain (`Continuous | Integer | Binary`)
- `Constraint` / `Relation` — constraint index + relation + RHS (+ tolerance)
- `SolveOptions` — method selection, iteration limit, numeric tolerance, method-specific options, optional progress callback.
- `SolveOutcome` / `SolveStatus` — solution + status (`Optimal | Feasible | Infeasible | Unbounded | IterationLimit | Cancelled`)

### Proposed JS/WASM DTOs

Solver types are not currently `serde` DTOs in Rust, so WASM bindings will need small wrapper DTOs that map into the Rust structs.

```ts
export type SolveMethod = "simplex" | "grgNonlinear" | "evolutionary";
export type ObjectiveKind = "maximize" | "minimize" | "target";
export type Relation = "lessEqual" | "greaterEqual" | "equal";
export type VarType = "continuous" | "integer" | "binary";

export interface Objective {
  kind: ObjectiveKind;
  targetValue?: number; // required for kind="target"
  targetTolerance?: number; // defaults to 0 (Rust clamps to >=0)
}

export interface VarSpec {
  // Use -Infinity/Infinity for unbounded (maps directly to f64 bounds in Rust).
  lower: number;
  upper: number;
  varType: VarType;
}

export interface Constraint {
  // Index into the model's constraint vector / constraintCells list.
  index: number;
  relation: Relation;
  rhs: number;
  tolerance?: number; // default 1e-8 (Rust clamps to >=0)
}

export interface SolverProblem {
  objective: Objective;
  variables: VarSpec[];
  constraints: Constraint[];
}

export interface SimplexOptions {
  maxPivots?: number; // default 10_000
  maxBnbNodes?: number; // default 1_000
  integerTolerance?: number; // default 1e-6
}

export interface GrgOptions {
  initialStep?: number; // default 1.0
  diffStep?: number; // default 1e-5
  penaltyWeight?: number; // default 10.0
  penaltyGrowth?: number; // default 2.0
  lineSearchShrink?: number; // default 0.5
  lineSearchMaxSteps?: number; // default 20
}

export interface EvolutionaryOptions {
  populationSize?: number; // default 40
  eliteCount?: number; // default 4
  mutationRate?: number; // default 0.2
  crossoverRate?: number; // default 0.7
  penaltyWeight?: number; // default 50.0
  seed?: number; // default 0x5EED_5EED_1234_5678 (u64; require safe integer)
}

export interface SolveOptions {
  method: SolveMethod; // default "grgNonlinear"
  maxIterations?: number; // default 500
  tolerance?: number; // default 1e-8
  applySolution?: boolean; // default true
  simplex?: SimplexOptions;
  grg?: GrgOptions;
  evolutionary?: EvolutionaryOptions;
}

export interface SolveProgress {
  iteration: number;
  bestObjective: number;
  currentObjective: number;
  maxConstraintViolation: number;
}

export type SolveStatus =
  | "Optimal"
  | "Feasible"
  | "Infeasible"
  | "Unbounded"
  | "IterationLimit"
  | "Cancelled";

export interface SolveOutcome {
  status: SolveStatus;
  iterations: number;
  originalVars: number[];
  bestVars: number[];
  bestObjective: number;
  maxConstraintViolation: number;
}

// Suggested binding shape (Engine-backed):
//   workbook.solve(
//     {
//       defaultSheet?: string,
//       objectiveCell: CellRef,
//       variableCells: CellRef[],
//       constraintCells: CellRef[],
//       problem: SolverProblem,
//       options?: SolveOptions,
//       onProgress?: (p: SolveProgress) => boolean, // return false to cancel
//     }
//   ): SolveOutcome
```

Validation + edge cases (Rust behavior):

- `Solver::solve` validates:
  - `problem.variables.len()` must equal `model.num_vars()` → otherwise a `SolverError` with message like `"variable spec count (X) does not match model vars (Y)"`.
  - Every `Constraint.index` must be `< model.num_constraints()` → otherwise `SolverError("constraint index ... out of range ...")`.
- Integer/binary normalization:
  - `VarType::Integer`: bounds are normalized to `ceil(lower)` / `floor(upper)`. If `lower > upper` after normalization → `SolverError("integer var {idx} has empty bounds [...]")`.
  - `VarType::Binary`: bounds are forced to `[0, 1]` regardless of input.
- Simplex-specific:
  - If a variable has a non-finite lower bound, simplex treats it as `0.0` (Excel-like “Assume Non-Negative” default).
  - Simplex infers a linear model by finite differences at the starting point; if the objective or any constraint is non-finite during inference it returns a `SolverError` (not a partial outcome).
- GRG-specific: only `Continuous` variables participate in the gradient; integer/binary vars are effectively held fixed (use Simplex or Evolutionary for mixed-integer problems).
- Progress + cancellation:
  - `SolveOptions.progress` returns `false` to cancel; solver returns `SolveStatus::Cancelled` with the best solution found so far.
- Engine integration (`EngineSolverModel`):
  - Decision variables must be coercible to numbers at construction time; otherwise `EngineSolverModel::new` fails with a `SolverError` like `"cell Sheet1!A1 is not numeric (...)"`.
  - Objective and constraint cells are coerced per-iteration; non-numeric values become `NaN` (methods treat non-finite values as very bad via a large penalty).

---

## Statistical Tools

### Regression Analysis

```typescript
interface RegressionResult {
  coefficients: number[];
  intercept: number;
  rSquared: number;
  adjustedRSquared: number;
  standardError: number;
  fStatistic: number;
  pValue: number;
  coefficientStats: CoefficientStats[];
  residuals: number[];
  predictions: number[];
}

interface CoefficientStats {
  coefficient: number;
  standardError: number;
  tStatistic: number;
  pValue: number;
  confidenceInterval: [number, number];
}

class RegressionEngine {
  linearRegression(
    y: number[],
    X: number[][],
    options?: { intercept?: boolean; confidenceLevel?: number }
  ): RegressionResult {
    const n = y.length;
    const k = X[0].length;
    const intercept = options?.intercept !== false;
    
    // Add intercept column if needed
    let designMatrix = X;
    if (intercept) {
      designMatrix = X.map(row => [1, ...row]);
    }
    
    // Solve normal equations: (X'X)^-1 X'y
    const XtX = this.matrixMultiply(this.transpose(designMatrix), designMatrix);
    const XtXinv = this.matrixInverse(XtX);
    const Xty = this.matrixMultiply(this.transpose(designMatrix), y.map(v => [v]));
    const beta = this.matrixMultiply(XtXinv, Xty).map(row => row[0]);
    
    // Predictions and residuals
    const predictions = designMatrix.map(row => 
      row.reduce((sum, x, i) => sum + x * beta[i], 0)
    );
    const residuals = y.map((yi, i) => yi - predictions[i]);
    
    // Statistics
    const yMean = this.mean(y);
    const SST = y.reduce((sum, yi) => sum + Math.pow(yi - yMean, 2), 0);
    const SSR = residuals.reduce((sum, r) => sum + r * r, 0);
    const SSE = SST - SSR;
    
    const rSquared = 1 - SSR / SST;
    const adjustedRSquared = 1 - (SSR / (n - k - 1)) / (SST / (n - 1));
    const standardError = Math.sqrt(SSR / (n - k - 1));
    
    // F-statistic
    const fStatistic = (SSE / k) / (SSR / (n - k - 1));
    const pValue = 1 - this.fCDF(fStatistic, k, n - k - 1);
    
    // Coefficient statistics
    const coefficientStats = beta.map((b, i) => {
      const se = standardError * Math.sqrt(XtXinv[i][i]);
      const t = b / se;
      const p = 2 * (1 - this.tCDF(Math.abs(t), n - k - 1));
      const alpha = 1 - (options?.confidenceLevel || 0.95);
      const tCrit = this.tInverse(1 - alpha / 2, n - k - 1);
      
      return {
        coefficient: b,
        standardError: se,
        tStatistic: t,
        pValue: p,
        confidenceInterval: [b - tCrit * se, b + tCrit * se] as [number, number]
      };
    });
    
    return {
      coefficients: intercept ? beta.slice(1) : beta,
      intercept: intercept ? beta[0] : 0,
      rSquared,
      adjustedRSquared,
      standardError,
      fStatistic,
      pValue,
      coefficientStats: intercept ? coefficientStats.slice(1) : coefficientStats,
      residuals,
      predictions
    };
  }
}
```

---

## Testing Strategy

### Power Feature Tests

```typescript
describe("Pivot Tables", () => {
  it("calculates sum aggregation correctly", () => {
    const data = [
      ["Region", "Product", "Sales"],
      ["East", "A", 100],
      ["East", "B", 150],
      ["West", "A", 200],
      ["West", "B", 250]
    ];
    
    const pivot = createPivot({
      rowFields: ["Region"],
      valueFields: [{ field: "Sales", aggregation: "sum" }]
    });
    
    const result = pivot.calculate(data);
    
    expect(result.data).toEqual([
      ["Region", "Sum of Sales"],
      ["East", 250],
      ["West", 450],
      ["Grand Total", 700]
    ]);
  });
  
  it("handles multiple value fields", () => {
    // Test multiple aggregations
  });
  
  it("respects filter context", () => {
    // Test with filters applied
  });
});

describe("Monte Carlo Simulation", () => {
  it("produces expected distribution", async () => {
    const result = await simulation.run({
      iterations: 10000,
      inputDistributions: [
        { cell: A1, distribution: { type: "normal", mean: 100, stdDev: 10 } }
      ],
      outputCells: [A1]
    });
    
    // Mean should be close to 100
    expect(result.outputStats["A1"].mean).toBeCloseTo(100, 0);
    
    // StdDev should be close to 10
    expect(result.outputStats["A1"].stdDev).toBeCloseTo(10, 0);
  });
});
```
