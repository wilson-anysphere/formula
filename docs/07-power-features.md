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
// Note: the canonical persisted schema (formula-model) uses an untagged Rust enum
// `PivotFieldRef` for `sourceField`. Cache-backed fields serialize as a plain string
// for backward compatibility; Data Model pivots serialize as structured objects.
//
// Important: when ingesting existing XLSX PivotCaches, Excel may encode Data Model
// field captions using DAX-like strings such as:
// - `Table[Column]`
// - `'Table Name'[Column]` (quoted table identifiers)
// - bracket escapes inside `[...]` (e.g. `A]]B` to represent `A]B`)
// - measures sometimes stored as `Total Sales` (no brackets) or `[Total Sales]`
//
// The engine normalizes UI captions (e.g. Pivot header rows) but resolves multiple
// legacy encodings when binding `PivotFieldRef` values to cache fields.
type PivotFieldRef =
  | string                               // cache field (worksheet header text)
  | { table: string; column: string }    // Data Model column (Table[Column])
  | { measure: string };                 // Data Model measure name (without brackets)

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
  
  // Caching (runtime-only; derived from the source). Do not treat this as persisted workbook state.
  cache?: PivotCache;
}

interface PivotField {
  sourceField: PivotFieldRef;
  name: string;
  sortOrder: "ascending" | "descending" | "manual";
  manualSort?: string[];
  grouping?: FieldGrouping;
  subtotal: AggregationType;
}

interface ValueField {
  sourceField: PivotFieldRef;
  name: string;
  aggregation: AggregationType;
  numberFormat?: string;
  showAs?: ShowAsType;
  baseField?: PivotFieldRef;
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
      const fieldKey =
        typeof field.sourceField === "string"
          ? field.sourceField
          : "measure" in field.sourceField
            ? `[${field.sourceField.measure}]`
            : `${field.sourceField.table}[${field.sourceField.column}]`;
      const values = Array.from(cache.uniqueValues.get(fieldKey) || []);
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

- `what_if::CellRef` — a `#[serde(transparent)]` A1-style cell reference string (defined in [`what_if/types.rs`](../crates/formula-engine/src/what_if/types.rs)).
  - For engine-backed usage (`EngineWhatIfModel`), accepted forms include:
    - `"A1"` (uses the adapter’s `default_sheet`)
    - `"Sheet1!A1"`
    - `"'My Sheet'!A1"` (Excel-style quoting for spaces)
    - `"'O''Brien'!A1"` (escaped `'` via doubled apostrophe)
- `what_if::CellValue` — scalar-only values: `Number(f64)`, `Text(String)`, `Bool(bool)`, `Blank` (defined in [`what_if/types.rs`](../crates/formula-engine/src/what_if/types.rs)).
- `what_if::WhatIfError<E>` — returned for invalid parameters, non-numeric cells, or underlying model failures (defined in [`what_if/types.rs`](../crates/formula-engine/src/what_if/types.rs)).

Engine adapter notes (in-tree today):

- `what_if::EngineWhatIfModel` (defined in [`what_if/engine_model.rs`](../crates/formula-engine/src/what_if/engine_model.rs)) defaults to **single-threaded** recalculation (`RecalcMode::SingleThreaded`) to reduce per-iteration overhead; hosts can opt into `RecalcMode::MultiThreaded` via `EngineWhatIfModel::with_recalc_mode(...)`.
- `EngineWhatIfModel` does **not** validate that referenced sheets exist:
  - `Engine::get_cell_value` treats missing sheets as blank.
    - Because What‑If tools require numeric `CellValue::Number(..)` values for targets/outputs, missing-sheet reads will typically surface as `WhatIfError::NonNumericCell` (blank is not numeric).
  - `Engine::set_cell_value` creates missing sheets on demand.
  - Hosts (and any future WASM bindings) should generally reject missing sheets to avoid accidentally creating new ones due to typos.
- `EngineWhatIfModel` intentionally exposes a *scalar-only* view of engine values:
  - arrays/spills are degraded to their top-left value
  - errors are degraded to their error code string (e.g. `"#DIV/0!"`)
  - rich values (entities/records, references, lambdas, etc.) are degraded to the engine display string
  - implication: if a target/output cell evaluates to an error or a non-scalar, What‑If tools will typically treat it as **non-numeric** and return `WhatIfError::NonNumericCell`.
- Calculation mode:
  - `formula_engine::Engine` supports “automatic” calculation mode where `set_cell_value` triggers an immediate recalc.
  - Goal Seek / Scenario apply / Monte Carlo / Solver all call `recalculate()` explicitly as part of their algorithms.
  - To avoid redundant recalculation (and large slowdowns), hosts should keep the engine in **manual** calculation mode while running these tools (this is also the engine’s default in-tree today).

Proposed JS/WASM DTOs (directly compatible with the current serde shapes):

```ts
export type CellRef = string;

// Mirrors: what_if::CellValue  (tagged enum; snake_case tags)
export type CellValue =
  | { type: "number"; value: number }
  | { type: "text"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "blank" };

// Mirrors `formula_engine::RecalcMode` / the in-tree `formula-wasm` string enums.
export type RecalcMode = "singleThreaded" | "multiThreaded";
```

WASM binding implementation note (important):

- Several Rust request structs provide defaults via constructors (e.g. `GoalSeekParams::new`, `SimulationConfig::new`) rather than via `#[serde(default)]`.
  - If a WASM binding deserializes directly into the Rust type (e.g. `serde_wasm_bindgen::from_value::<GoalSeekParams>(...)`), you must provide all **non-`Option`** fields explicitly (constructor defaults are not applied).
  - `Option<T>` fields behave like normal serde optionals: they may be omitted or set to `null`.
  - To provide a JS-friendly API with optional fields like `maxIterations`, define a JS-facing DTO with optional fields and map it into the Rust types using constructors:
    - `GoalSeekParams::new(...)` + overwrite tuning fields when provided
    - `SimulationConfig::new(iterations)` + overwrite `seed`, `correlations`, `histogram_bins`, etc.

Error surface (host contract):

- Rust functions return `Result<T, WhatIfError<_>>`. WASM bindings should throw a JS `Error` whose message is `WhatIfError::to_string()` (e.g. `"invalid parameters: iterations must be > 0"`).
  - When using `EngineWhatIfModel`, cell-ref parse failures and engine recalculation failures surface as `WhatIfError::Model(EngineError)` and format like `"model error: ..."` (exact message depends on the engine error kind).

### Goal Seek

Rust API ([`what_if/goal_seek.rs`](../crates/formula-engine/src/what_if/goal_seek.rs)):

- `GoalSeekParams` (`#[serde(rename_all = "camelCase")]`) — target/changing cells plus numeric tuning knobs.
- `GoalSeekResult` (`#[serde(rename_all = "camelCase")]`) — final solution + diagnostics.
- `GoalSeekProgress` (`#[serde(rename_all = "camelCase")]`) — progress event emitted once at iteration 0 and after every Newton/bisection step.
- `GoalSeekStatus` — `{ Converged, MaxIterationsReached, NoBracketFound, NumericalFailure }`.
- `GoalSeek::{solve, solve_with_progress}` — synchronous solver (Newton step + finite-difference derivative; falls back to bisection if derivative is too small or non-finite).

JS/WASM DTOs:

```ts
/**
 * Rust serde DTO shape (if you expose `formula_engine::what_if::goal_seek::GoalSeekParams`
 * directly through wasm and deserialize it via serde).
 *
 * Note: `GoalSeekParams` itself has non-optional tuning fields; JS/WASM bindings typically expose
 * these as optionals and then map into `GoalSeekParams::new(...)` so defaults are applied (this is
 * how the in-tree `formula-wasm` `workbook.goalSeek(...)` binding works).
 */
export interface GoalSeekParams {
  targetCell: CellRef;
  targetValue: number;
  changingCell: CellRef;

  // Default: 100 (from `GoalSeekParams::new`).
  maxIterations: number;
  // Default: 1e-7 (from `GoalSeekParams::new`).
  tolerance: number;
  /**
   * Finite-difference derivative step.
   *
   * - Default: `null` (auto), in which case Rust uses `(abs(input)*0.001).max(0.001)` per-iteration.
   */
  derivativeStep?: number | null;
  // Default: 1e-10 (from `GoalSeekParams::new`).
  minDerivative: number;
  // Default: 50 (from `GoalSeekParams::new`).
  maxBracketExpansions: number;
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

// Matches the `formula-wasm` `CellChange` DTO (also returned by `workbook.recalculate()`).
export interface CellChange {
  sheet: string;
  address: string;
  value: any; // JSON scalar: null | boolean | number | string
}

/**
 * Current `formula-wasm` API:
 *   `workbook.goalSeek(params)` → `{ result, changes }`
 *   (see [`crates/formula-wasm/src/lib.rs`](../crates/formula-wasm/src/lib.rs)).
 */
export interface GoalSeekRequest {
  // Defaults to "Sheet1" when omitted/empty.
  sheet?: string;

  // A1 addresses WITHOUT a `Sheet!` prefix (sheet is provided separately).
  targetCell: string;
  targetValue: number;
  changingCell: string;

  // Optional tuning; when omitted (or set to `null`), defaults match `GoalSeekParams::new(...)`:
  //   maxIterations=100, tolerance=1e-7, derivativeStep=null (auto), minDerivative=1e-10,
  //   maxBracketExpansions=50.
  maxIterations?: number;
  tolerance?: number;
  derivativeStep?: number | null;
  minDerivative?: number;
  maxBracketExpansions?: number;
}

export interface GoalSeekResponse {
  result: GoalSeekResult;
  // A `recalculate()`-compatible cell change list covering the full goal seek run.
  changes: CellChange[];
}
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
- Side effects: Goal Seek mutates spreadsheet state as it searches:
  - `changingCell` is overwritten on every step.
    - For `status: "Converged"` and `status: "MaxIterationsReached"`, the model ends with `changingCell == solution`.
    - For `status: "NoBracketFound"` and `status: "NumericalFailure"`, the returned `solution` is best-effort, but the raw Rust algorithm may be left at the **last evaluated probe/bracketing value** rather than at `solution`. Callers who need a strict end state should explicitly write `solution` back into the cell after the call (the `formula-wasm` binding does this automatically and includes the delta in `changes`).
  - The model is recalculated after every update, so `targetCell` and dependents reflect the latest candidate.
  - Engine-backed note: `GoalSeek` sets the changing cell via `WhatIfModel::set_cell_value(CellValue::Number(...))`.
    - For `EngineWhatIfModel` / `formula-wasm`, this overwrites the cell as a literal value and clears any existing formula in `changingCell`.

WASM binding validation rules (current `formula-wasm` implementation):

- `sheet`:
  - must be a string when provided; empty/whitespace is treated as `"Sheet1"`.
  - must refer to an existing sheet (otherwise the call throws).
  - lookup is case-insensitive for ASCII letters (matches current `formula-wasm` sheet resolution semantics).
- `targetCell` / `changingCell`:
  - must be non-empty strings
  - must be valid A1 addresses
  - must **not** contain `!` (no sheet prefix)
  - `$` absolute markers are allowed (e.g. `"$A$1"`); the binding normalizes them to `A1` internally.
  - A1 parsing is case-insensitive (e.g. `"a1"` is accepted).
- `targetValue` must be a finite number.
- `maxIterations` (optional) must be an integer `> 0`.
  - when omitted or `null`, the binding uses the Rust default `100` (from `GoalSeekParams::new`).
- `tolerance` (optional) must be a finite number and `> 0`.
  - when omitted or `null`, the binding uses the Rust default `1e-7` (from `GoalSeekParams::new`).
- `derivativeStep` (optional) must be a finite number and `> 0`.
  - when omitted or `null`, the solver uses an auto step size per-iteration (Rust `GoalSeekParams.derivative_step = None`).
- `minDerivative` (optional) must be a finite number and `> 0`.
  - when omitted or `null`, the binding uses the Rust default `1e-10` (from `GoalSeekParams::new`).
- `maxBracketExpansions` (optional) must be an integer `> 0`.
  - when omitted or `null`, the binding uses the Rust default `50` (from `GoalSeekParams::new`).

WASM binding side effects / integration notes:

- `goalSeek` mutates the workbook state and updates the `changingCell` input to `result.solution` (as a number).
  - If the cell previously had a “rich” input (`setCellRich`), it is cleared so `getCell` reflects the scalar input.
  - Note: the raw Rust `GoalSeek` algorithm may leave the model at the **last evaluated probe value**
    rather than at `result.solution` for `NoBracketFound` and some `NumericalFailure` paths. The
    `formula-wasm` binding re-applies `result.solution` to the workbook state and performs a final
    recalc so dependents are consistent.
- The wasm `goalSeek` API runs recalculation internally and returns `{ result, changes }`, where:
  - `result` is the `GoalSeekResult`.
  - `changes` is a `recalculate()`-compatible `CellChange[]` so hosts can update UI/caches without a separate `recalculate()` call.

Example (apply `changes` to a UI cell store):

```ts
const { result, changes } = workbook.goalSeek({
  sheet: "Sheet1",
  targetCell: "B1",
  targetValue: 100,
  changingCell: "A1",
});

for (const ch of changes) {
  cellStore.set(`${ch.sheet}!${ch.address}`, ch.value);
}

console.log(result.status, result.solution);
```

Progress events:

- Rust exposes `GoalSeek::solve_with_progress` and the `GoalSeekProgress` struct.
- The current wasm binding does **not** expose progress events and does not support cancellation.

### Scenario Manager

Rust API ([`what_if/scenario_manager.rs`](../crates/formula-engine/src/what_if/scenario_manager.rs)):

- WASM status: Scenario Manager is **not currently exposed** via `formula-wasm` (unlike `goalSeek`). The DTOs below describe the intended binding surface if/when it is added.

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

/**
 * Proposed `formula-wasm` request shape (mirrors the existing `goalSeek` style).
 *
 * For simplicity and to match `goalSeek`, the request uses A1 addresses without `Sheet!`
 * prefixes and supplies the sheet separately. The binding can convert these into fully
 * qualified `CellRef` strings (`Sheet1!A1` / `'My Sheet'!A1`) before calling Rust so
 * scenarios remain stable even if the host’s “active sheet” changes.
 */
export interface CreateScenarioRequest {
  // Defaults to "Sheet1" when omitted/empty.
  sheet?: string;

  name: string;
  changingCells: string[]; // A1 addresses (no sheet prefix)
  values: CellValue[]; // must match changingCells length
  createdBy: string;
  comment?: string | null;
}

export interface GenerateScenarioSummaryReportRequest {
  // Defaults to "Sheet1" when omitted/empty.
  sheet?: string;
  resultCells: string[]; // A1 addresses (no sheet prefix)
  scenarioIds: ScenarioId[];
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
//   // Likely implemented as `WasmWorkbook` methods (consistent with existing formula-wasm APIs),
//   // but could also be exposed as a separate manager object. Either approach is viable as long
//   // as scenario state persists across calls.
//   workbook.createScenario(request: CreateScenarioRequest) -> ScenarioId
//   workbook.listScenarios() -> Scenario[]
//   workbook.getScenario(id: ScenarioId) -> Scenario | null
//   workbook.deleteScenario(id: ScenarioId) -> boolean
//   workbook.applyScenario(id: ScenarioId) / workbook.restoreScenarioBase()
//   workbook.generateScenarioSummaryReport(request: GenerateScenarioSummaryReportRequest) -> SummaryReport
```

Validation + edge cases (Rust behavior):

- `create_scenario`: `changing_cells.len()` must equal `values.len()` → `WhatIfError::InvalidParams("changing_cells and values must have equal length")`.
- `changing_cells` should not contain duplicates:
  - The scenario’s `values` are stored in a `HashMap<CellRef, CellValue>`, so duplicates will be de-duplicated (last value wins).
  - The `Scenario.changing_cells` vector will still contain the duplicates, which can be surprising in reports/UX. This is not currently validated by Rust.
- Scenario names are not required to be unique. **If two scenarios share a name**, `generate_summary_report` will overwrite the earlier entry in `results` (it’s a `HashMap<String, ...>` keyed by name).
  - The summary report always includes a `"Base"` row. If a scenario is named `"Base"`, it will overwrite that row.
- Map ordering:
  - `Scenario.values` and `SummaryReport.results` are `HashMap`s; iteration/serialization order is **not stable**. Hosts should treat these as maps keyed by `CellRef`/scenario name, not as ordered lists.
- `apply_scenario` / `generate_summary_report` with an unknown `ScenarioId` → `WhatIfError::InvalidParams("scenario not found")`.
- `restore_base` is a no-op if no scenario has been applied yet (`base_values` empty).
- Base snapshot semantics: applying multiple scenarios captures the **union** of their `changing_cells` in `base_values` so `restore_base` can fully return to the original state even if scenarios touch different inputs.
- `restore_base` does **not** clear `base_values`; hosts can call `clear_base_values()` to reset the snapshot explicitly.
- `ScenarioId` values are generated by `ScenarioManager` as an incrementing `u64` counter (starting at 0). They are local to the manager instance (not globally unique like UUIDs).
  - Implementation detail: the counter uses wrapping arithmetic (`wrapping_add(1)`), so after `u64::MAX` scenario creations it will wrap back to 0 and IDs may repeat.
- `SummaryReport.changingCells` is taken from the **first** scenario in `scenarioIds` (or empty when `scenarioIds` is empty). If a host allows scenarios with different `changingCells`, the summary report may not reflect a union; hosts should enforce a consistent changing-cell set for Excel-like behavior.
- Side effects:
  - `applyScenario` mutates the workbook state (writes scenario values and recalculates) and leaves that scenario active until `restoreBase()` is called.
  - `generateSummaryReport(...)` restores the base state before returning (it calls `restore_base` internally at the start and end).
  - Engine-backed note: Scenario Manager writes changing cells via `WhatIfModel::set_cell_value(...)`.
    - When using `EngineWhatIfModel` / a future wasm binding, this clears any existing formulas in those changing cells.
    - Base snapshots store **values**, not formulas; `restoreBase()` restores the captured values and does not reinstate original formulas.
    - Hosts should restrict scenario “changing cells” to literal input/value cells (Excel-like behavior).
- Additional in-tree behavior (useful for bindings / UX):
  - `ScenarioManager::current_scenario()` returns the currently-applied scenario id (or `None` if the base state is active).
  - Deleting the current scenario (`delete_scenario`) clears `current_scenario`.
  - `ScenarioManager` itself is in-memory and does not implement `Serialize`.
    - If a host wants persistence, it should persist enough data to recreate scenarios via `create_scenario`.
    - Note: `create_scenario` allocates new ids and timestamps (`createdMs`), so ids/timestamps will not be stable across reload unless the Rust API is extended.
    - If exposed via `formula-wasm`, the binding should also keep the JS-facing workbook input state consistent with scenario apply/restore mutations (similar to how `goalSeek` updates the `changingCell` input). Otherwise `getCell`/`toJson` may drift from engine state.

Proposed WASM binding validation rules (if/when implemented):

- `sheet`:
  - must be a string when provided; empty/whitespace is treated as `"Sheet1"`.
  - must refer to an existing sheet (otherwise throw `"missing sheet: ..."` like other `formula-wasm` APIs).
  - lookup should be case-insensitive for ASCII letters for consistency with other `formula-wasm` APIs.
- `changingCells` / `resultCells`:
  - each cell must be a non-empty string
  - must be a valid A1 address
  - must **not** contain `!` (no sheet prefix); the binding should apply the `sheet` context itself
  - A1 parsing should be case-insensitive; `$` absolute markers are allowed and can be normalized to plain `A1`.
- `values.length` must equal `changingCells.length` (mirror Rust error message if possible).

### Monte Carlo Simulation

Rust API ([`what_if/monte_carlo.rs`](../crates/formula-engine/src/what_if/monte_carlo.rs)):

- WASM status: Monte Carlo is **not currently exposed** via `formula-wasm`. The DTOs below describe the intended binding surface if/when it is added.

- `SimulationConfig` (`#[serde(rename_all = "camelCase")]`), `InputDistribution`, `Distribution`, `CorrelationMatrix`
- `MonteCarloEngine::{run_simulation, run_simulation_with_progress}`
- `SimulationProgress` (`completedIterations`, `totalIterations`)
- `SimulationResult` (`outputStats`, `outputSamples`)
- `OutputStatistics` (`mean`, `median`, `stdDev`, `percentiles`, `histogram`, …)

Proposed JS/WASM DTOs (field names match Rust’s serde output):

```ts
/**
 * Rust serde DTO shape (if you expose `formula_engine::what_if::monte_carlo::SimulationConfig`
 * directly through wasm and deserialize it via serde).
 */
export interface SimulationConfig {
  iterations: number;
  inputDistributions: InputDistribution[];
  outputCells: CellRef[];

  // Required by the Rust struct (u64; require a safe integer for JS determinism).
  seed: number;
  // Optional in Rust (Option<CorrelationMatrix>) and may be omitted or set to `null`.
  correlations?: CorrelationMatrix | null;
  // Required by the Rust struct.
  histogramBins: number;
}

export interface InputDistribution {
  cell: CellRef;
  distribution: Distribution;
}

// Mirrors Rust: #[serde(tag = "type", rename_all = "snake_case")]
//
// Notes:
// - beta: when `min`/`max` are omitted/null, Rust defaults to the unit interval [0, 1].
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

// Note: `outputStats` / `outputSamples` come from Rust `HashMap`s; JSON key order is not stable.
// Treat them as maps keyed by output `CellRef`, not as ordered lists.

/**
 * Proposed `formula-wasm` request shape (mirrors the existing `goalSeek` style).
 *
 * Note: cells are expressed as A1 addresses WITHOUT sheet prefixes; the sheet is
 * provided separately and used as the default sheet for the engine adapter.
 */
export interface MonteCarloRequest {
  // Defaults to "Sheet1" when omitted/empty.
  sheet?: string;

  iterations: number;
  inputDistributions: { cell: string; distribution: Distribution }[];
  outputCells: string[];

  // Optional; Rust defaults: seed=0, histogramBins=50
  seed?: number;
  correlations?: CorrelationMatrix | null;
  histogramBins?: number;

  // Optional; on wasm builds "multiThreaded" falls back to single-threaded recalc.
  recalcMode?: RecalcMode;

  /**
   * Optional progress callback; report-only (cannot cancel).
   *
   * The Rust engine reports progress roughly every 1% (and always on the last iteration).
   */
  onProgress?: (p: SimulationProgress) => void;
}

// Proposed future entrypoint:
//   workbook.runMonteCarloSimulation(request: MonteCarloRequest): SimulationResult
```

Validation + edge cases (Rust behavior):

- `iterations` must be `> 0` → `WhatIfError::InvalidParams("iterations must be > 0")`.
- `histogramBins` must be `> 0` → `WhatIfError::InvalidParams("histogram_bins must be > 0")`.
- `outputCells` must be non-empty → `WhatIfError::InvalidParams("output_cells must not be empty")`.
- `inputDistributions` may be empty (in which case the simulation simply re-evaluates outputs each iteration without changing any inputs).
- `outputCells` should not contain duplicates:
  - The implementation uses a `HashMap<CellRef, Vec<f64>>` for `outputSamples`, but iterates `outputCells` when pushing values, so duplicates would push multiple samples per iteration into the same vector.
  - This is not currently validated by Rust; hosts/bindings should enforce uniqueness.
- `inputDistributions[*].cell` should not contain duplicates (not currently validated):
  - duplicates will cause the same cell to be set multiple times per iteration (last write wins), and correlated sampling becomes ill-defined.
- Every `Distribution` is validated up-front. Invalid distributions return `WhatIfError::InvalidParams(<msg>)`, where `<msg>` is one of:
  - normal:
    - `"normal mean must be finite"`
    - `"normal std_dev must be finite"`
    - `"normal std_dev must be >= 0"`
  - uniform:
    - `"uniform min and max must be finite"`
    - `"uniform min must be <= max"`
  - triangular:
    - `"triangular min, mode, and max must be finite"`
    - `"triangular requires min <= mode <= max"`
  - lognormal:
    - `"lognormal mean must be finite"`
    - `"lognormal std_dev must be finite"`
    - `"lognormal std_dev must be >= 0"`
  - discrete:
    - `"discrete distribution requires at least one value"`
    - `"discrete values and probabilities must have equal length"`
    - `"discrete values must be finite"`
    - `"discrete probabilities must be finite"`
    - `"discrete probabilities must be >= 0"`
    - `"discrete probabilities must sum to > 0"` (does not need to equal 1)
  - beta:
    - `"beta alpha and beta must be finite"`
    - `"beta alpha and beta must be > 0"`
    - `"beta min and max must be finite"` (when both are provided)
    - `"beta min must be <= max"` (when both are provided)
    - When `min`/`max` are omitted/null, Rust defaults to `[0, 1]`.
  - exponential:
    - `"exponential rate must be finite"`
    - `"exponential rate must be > 0"`
  - poisson:
    - `"poisson lambda must be finite"`
    - `"poisson lambda must be >= 0"`
- Output cells must evaluate to numbers each iteration; otherwise `WhatIfError::NonNumericCell { cell, value }`.
- Correlations:
  - `correlations.matrix` is validated up-front. Invalid matrices return `WhatIfError::InvalidParams(<msg>)`, where `<msg>` is one of:
    - `"correlation matrix must not be empty"`
    - `"correlation matrix size must match input_distributions length"`
    - `"correlation matrix must be square"`
    - `"correlation matrix contains non-finite value"`
    - `"correlation matrix diagonal entries must be 1"`
    - `"correlation matrix entries must be within [-1, 1]"`
    - `"correlation matrix must be symmetric"`
    - `"correlation matrix is not positive definite"` (Cholesky decomposition failure)
  - When `correlations` is provided, correlated sampling uses a **Gaussian copula**:
    1. Generate correlated standard normals using the provided correlation matrix.
    2. Convert each normal sample `z` into a uniform sample `u = Φ(z)` (standard normal CDF).
    3. Transform `u` through the inverse CDF of each input distribution.
  - Correlations are supported for the following input distributions:
    - `normal`, `uniform`, `triangular`, `lognormal`, `exponential`, `beta`
  - Correlations are **not** supported for:
    - `discrete` → `InvalidParams("correlated sampling is not supported for discrete distributions")`
    - `poisson` → `InvalidParams("correlated sampling is not supported for poisson distributions")`
  - Matrix row/column order is the same as `inputDistributions` order (`matrix[i][j]` is corr(input i, input j)).
- Histogram edge cases:
  - If all samples are identical (`min == max`), Rust returns a single bin with `count = iterations`.
  - If samples are empty or min/max are non-finite (shouldn’t happen for valid runs), histogram bins are empty.
- Progress callback behavior:
  - `run_simulation_with_progress` reports progress roughly every 1% (and always on the final iteration).
- Progress callback:
  - Rust `run_simulation_with_progress` cannot cancel the run (progress is report-only); WASM bindings should document that cancellation is not supported for Monte Carlo (other than host-level task cancellation).
- Side effects: Monte Carlo also mutates spreadsheet state:
  - Each `inputDistributions[*].cell` is overwritten every iteration and is left set to the *last iteration’s* sampled value.
  - The model is recalculated after each sample batch; outputs reflect the last iteration at the end of the run.
  - If exposed via `formula-wasm`, the binding must keep the JS-facing workbook input state consistent with these mutations (similar to how `goalSeek` updates the `changingCell` input). Otherwise `getCell`/`toJson` may drift from engine state.
  - Engine-backed note: like Goal Seek, Monte Carlo writes input cells via `set_cell_value(CellValue::Number(...))`, which will clear any existing formulas in those input cells when using `EngineWhatIfModel`.

Proposed WASM binding validation rules (if/when implemented):

- `sheet`:
  - must be a string when provided; empty/whitespace is treated as `"Sheet1"`.
  - must refer to an existing sheet (otherwise throw).
  - lookup should be case-insensitive for ASCII letters for consistency with other `formula-wasm` APIs.
- `iterations` / `histogramBins`:
  - must be finite numbers, integer-valued, and `> 0`.
- `inputDistributions[*].cell` / `outputCells[*]`:
  - must be non-empty strings
  - must be valid A1 addresses
  - must **not** contain `!` (no sheet prefix); the binding should apply the `sheet` context itself
  - A1 parsing should be case-insensitive; `$` absolute markers are allowed and can be normalized to plain `A1`.
- `seed` (optional): must be a non-negative safe integer (to preserve deterministic seeding).
- `onProgress` (optional): must be a function.

WASM binding return shape (suggested):

- A future `runMonteCarloSimulation` binding can return only the `SimulationResult`.
- It should not attempt to return a full `recalculate()` cell-delta list for each iteration; hosts can fetch cells/ranges after the run if they need to refresh UI state.

---

## Solver (Optimization)

The Rust implementation lives in [`solver/mod.rs`](../crates/formula-engine/src/solver/mod.rs) (and supporting files under `crates/formula-engine/src/solver/*`) and is exposed as `formula_engine::solver`.

WASM status: Solver is **not currently exposed** via `formula-wasm`. The DTOs below describe the intended binding surface if/when it is added.

This is a small-but-functional Excel-like Solver with three methods:

- **Simplex** (`SolveMethod::Simplex`) — linear programming (LP) with optional integer/binary variables (branch-and-bound).
- **GRG Nonlinear** (`SolveMethod::GrgNonlinear`) — penalty-based gradient method (continuous variables only).
- **Evolutionary** (`SolveMethod::Evolutionary`) — genetic algorithm (supports integer/binary; suitable for non-smooth problems).

### Rust API surface

Core types ([`solver/mod.rs`](../crates/formula-engine/src/solver/mod.rs)):

- `SolverModel` trait — model abstraction (`num_vars`/`num_constraints` + `get_vars`/`set_vars`/`recalc`/`objective`/`constraints`).
- `EngineSolverModel` ([`solver/engine_model.rs`](../crates/formula-engine/src/solver/engine_model.rs)) — adapter that binds the solver to `formula_engine::Engine` cell references.
- `SolverProblem` — `{ objective: Objective, variables: Vec<VarSpec>, constraints: Vec<Constraint> }`
- `Objective` / `ObjectiveKind` — maximize/minimize/target (with `targetValue` + `targetTolerance`)
- `VarSpec` / `VarType` — bounds + variable domain (`Continuous | Integer | Binary`)
- `Constraint` / `Relation` — constraint index + relation + RHS (+ tolerance)
- `SolveMethod` — method selection (`Simplex | GrgNonlinear | Evolutionary`)
- `SolveOptions` — method selection, iteration limit, numeric tolerance, method-specific options, optional progress callback.
- `Progress` — progress snapshot sent to the callback (`iteration`, objectives, constraint violation).
- `SolveOutcome` / `SolveStatus` — solution + status (`Optimal | Feasible | Infeasible | Unbounded | IterationLimit | Cancelled`)
- `SolverError` — error type returned for invalid problems/models or engine/model failures.
- Method-specific option structs (re-exported from `solver/mod.rs`):
  - `SimplexOptions` ([`solver/simplex.rs`](../crates/formula-engine/src/solver/simplex.rs))
  - `GrgOptions` ([`solver/grg.rs`](../crates/formula-engine/src/solver/grg.rs))
  - `EvolutionaryOptions` ([`solver/evolutionary.rs`](../crates/formula-engine/src/solver/evolutionary.rs))

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
  /**
   * Optional RNG seed.
   *
   * - If provided, it should be a non-negative safe integer (<= `Number.MAX_SAFE_INTEGER`)
   *   so the binding can losslessly convert it to `u64`.
   * - If omitted, Rust uses a fixed internal `u64` seed (`0x5EED_5EED_1234_5678`).
   *   Note: that constant is **not** safely representable as a JS number, but callers
   *   typically don't need to pass it explicitly.
   * - Rust treats a seed of `0` as `1` internally (to avoid the Xorshift zero-state).
   */
  seed?: number;
}

export interface SolveOptions {
  method?: SolveMethod; // default "grgNonlinear"
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
//       // Defaults to "Sheet1" when omitted/empty.
//       sheet?: string,
//
//       // A1 addresses without `Sheet!` prefixes (mirrors `goalSeek`).
//       objectiveCell: string,
//       variableCells: string[],
//       constraintCells: string[],
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
- Iteration limits:
  - `SolveOptions.max_iterations` is used by **GRG** and **Evolutionary**.
  - **Simplex** uses method-specific limits (`SimplexOptions.max_pivots`, `SimplexOptions.max_bnb_nodes`) rather than `SolveOptions.max_iterations`.
  - `SolveOutcome.iterations` is method-dependent:
    - `Simplex`: branch-and-bound nodes searched (includes the root LP relaxation).
    - `GrgNonlinear`: optimization iterations executed.
    - `Evolutionary`: generations executed.
- Integer/binary normalization:
  - `VarType::Integer`: bounds are normalized to `ceil(lower)` / `floor(upper)`. If `lower > upper` after normalization → `SolverError("integer var {idx} has empty bounds [...]")`.
  - `VarType::Binary`: bounds are forced to `[0, 1]` regardless of input.
- Simplex-specific:
  - If a variable has a non-finite lower bound, simplex treats it as `0.0` (Excel-like “Assume Non-Negative” default).
    - If `upper` is finite and `< 0` after this normalization, simplex returns `SolveStatus::Infeasible` with `bestVars: []` (not a `SolverError`).
  - Simplex infers a linear model by finite differences at the starting point. Non-finite objective/constraint values during inference fail fast with a `SolverError` message like:
    - `"objective is not finite at the starting point (...); simplex requires a valid linear model"`
    - `"constraint {idx} is not finite at the starting point (...); simplex requires a valid linear model"`
    - `"objective is not finite while inferring coefficient for var {j} (...)"`
    - `"constraint {idx} is not finite while inferring coefficient for var {j} (...)"`
- GRG-specific: only `Continuous` variables participate in the gradient; integer/binary vars are effectively held fixed (use Simplex or Evolutionary for mixed-integer problems).
- Evolutionary-specific:
  - For variables with unbounded/infinite limits (`lower=-Infinity` or `upper=Infinity`), the evolutionary method uses a heuristic finite search window around the starting value (`center ± 10`) when generating random candidates and mutations. For meaningful global search, hosts should provide finite bounds.
- Progress + cancellation:
  - `SolveOptions.progress` returns `false` to cancel; solver returns `SolveStatus::Cancelled`.
    - `bestVars` contains the best solution found so far when available; it may be empty if the solver cancels before finding *any* feasible solution (notably for `Simplex`).
- Side effects:
  - The solver overwrites `variableCells` repeatedly while searching.
  - On return, `applySolution` controls final state:
    - `true`: the best solution is applied to the model.
    - `false`: the original variable values are restored (the solve still recalculates during the search).
  - If the solver fails to find *any* candidate solution (`outcome.bestVars` is empty), the model is restored to the original variables regardless of `applySolution`.
  - Variable domain projection:
    - All methods clamp candidate variables into `VarSpec` bounds.
    - Integer variables are snapped via `round()`; binary variables are snapped to `0`/`1` using a `0.5` threshold.
    - Implication: hosts should treat solver variables as *owned* by the solver during the run (don’t assume they remain exactly the user-entered floating-point values).
  - Engine-backed note: `EngineSolverModel` updates decision variables via `engine.set_cell_value(...)`, which clears any existing formulas in variable cells.
- Progress callback frequency (implementation detail, but useful for hosts):
  - `Simplex`: callback is invoked during branch-and-bound node search and once at the end (can be frequent for mixed-integer problems).
  - `GrgNonlinear`: callback is invoked once per iteration.
  - `Evolutionary`: callback is invoked once per generation.
- Engine integration (`EngineSolverModel`):
  - Cell-ref parsing:
    - Accepted forms include `A1` (default sheet), `Sheet1!A1`, and `'My Sheet'!A1` (quoted with Excel escaping `''`).
    - Invalid refs fail fast during `EngineSolverModel::new` (e.g. empty refs, missing sheet name when `!` is present).
      - `EngineSolverModel::new` returns `SolverError` messages like:
        - `"cell reference cannot be empty"`
        - `"invalid cell reference '<input>': missing address"` (e.g. `"Sheet1!"`)
        - `"invalid cell reference '<input>': missing sheet name"` (e.g. `"!A1"`)
    - Note: `EngineSolverModel` does **not** validate that the address portion is a syntactically-valid A1 reference (it stores it as a string).
      - The engine will treat invalid addresses as `#REF!`.
      - Because decision variables are read **strictly** at construction time, an invalid variable address will typically surface as a `SolverError("cell Sheet!<addr> is not numeric (value: #REF!)")` during `EngineSolverModel::new`.
      - Invalid objective/constraint addresses are read as `NaN` and then handled by the solver’s non-finite penalties.
    - Note: `EngineSolverModel` does **not** validate that a sheet exists at construction time:
      - For reads, the engine treats a missing sheet as blank; this is coerced to `0` for numeric reads.
      - When the solver writes decision variables via `engine.set_cell_value(...)`, the engine will create missing sheets on demand.
      - A WASM binding should generally treat missing sheets as an error for consistency with other `formula-wasm` APIs.
    - A future `formula-wasm` binding will likely mirror `goalSeek` and accept A1 addresses without `Sheet!` prefixes, using a separate `sheet` field as the default sheet.
  - Numeric coercion:
    - Decision variables are read **strictly** at construction time:
      - accepted: numbers, booleans (`TRUE`→1, `FALSE`→0), blanks (`0`), and numeric text (parsed using the engine’s `ValueLocaleConfig`, so thousands/decimal separators and common adornments are accepted)
      - rejected: arrays/spills and any value whose display string can’t be parsed as a finite number
      - failure returns a `SolverError` like: `"cell Sheet1!A1 is not numeric (value: ...)"`.
    - Objective + constraint cells are re-read after every `recalc()` using the same coercion rules, but **non-coercible values become `NaN`** rather than throwing.
      - Downstream solver code treats non-finite objective/constraint values as very bad via a fixed penalty (`NON_FINITE_PENALTY = 1e30`).

Proposed WASM binding validation rules (if/when implemented):

- `sheet`:
  - must be a string when provided; empty/whitespace is treated as `"Sheet1"`.
  - must refer to an existing sheet (otherwise throw).
  - lookup should be case-insensitive for ASCII letters for consistency with other `formula-wasm` APIs.
- `objectiveCell`, `variableCells[*]`, `constraintCells[*]`:
  - must be non-empty strings
  - must be valid A1 addresses
  - should **not** contain `!` (no sheet prefix) if the binding follows the `goalSeek` style
  - A1 parsing should be case-insensitive; `$` absolute markers are allowed and can be normalized to plain `A1`.
- `problem.variables.length` must equal `variableCells.length`.
- `variableCells` should not contain duplicates (not currently validated by Rust); duplicates would cause multiple decision variables to write to the same cell (last write wins).
- `constraintCells` should not contain duplicates (not currently validated); duplicates are allowed but can make constraints ambiguous in UI.
- `problem.constraints[*].index` must be within `[0, constraintCells.length)`.
- Numeric fields in `problem` / `options` should be finite (no NaN); integer-valued fields should be integers.
  - Variable specs: `lower`/`upper` may be `±Infinity` for unbounded, but should not be `NaN`. Prefer finite bounds for integer/binary vars.
  - Constraint specs: `rhs` should be finite; `tolerance` should be `>= 0`.
  - Objective: if `kind === "target"`, require finite `targetValue` and `targetTolerance >= 0`.
  - `options.maxIterations` should be an integer `> 0`; `options.tolerance` should be `>= 0`.
  - Simplex options: `maxPivots`/`maxBnbNodes` should be integers `> 0`; `integerTolerance >= 0`.
  - GRG options: `diffStep > 0`, `penaltyWeight >= 0`, `penaltyGrowth >= 1`, `lineSearchShrink` in `(0, 1)`, `lineSearchMaxSteps` integer `> 0`.
  - Evolutionary options: `populationSize` integer `> 0`, `eliteCount` integer in `[0, populationSize]`,
    `mutationRate`/`crossoverRate` in `[0, 1]`, `penaltyWeight >= 0`, `seed` is a non-negative safe integer.

WASM binding return shape (suggested):

- Return only the `SolveOutcome` (mirrors the Rust return type).
- If the host needs UI deltas, a binding can also return a `recalculate()`-compatible `CellChange[]`
  alongside the outcome (as `goalSeek` does); otherwise hosts can refresh via `getCell`/`getRange`.

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
    // Note: Monte Carlo is not yet exposed via `formula-wasm`; this is an example
    // of the *intended* JS/WASM-facing API.
    const result = workbook.runMonteCarloSimulation({
      iterations: 10000,
      inputDistributions: [
        { cell: "A1", distribution: { type: "normal", mean: 100, stdDev: 10 } }
      ],
      outputCells: ["A1"]
    });
    
    // Mean should be close to 100
    expect(result.outputStats["A1"].mean).toBeCloseTo(100, 0);
    
    // StdDev should be close to 10
    expect(result.outputStats["A1"].stdDev).toBeCloseTo(10, 0);
  });
});
```
