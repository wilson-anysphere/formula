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

### Goal Seek

```typescript
interface GoalSeekParams {
  targetCell: CellRef;      // Cell containing formula
  targetValue: number;      // Value we want
  changingCell: CellRef;    // Cell to adjust
  maxIterations?: number;
  tolerance?: number;
}

class GoalSeek {
  async solve(params: GoalSeekParams): Promise<GoalSeekResult> {
    const { targetCell, targetValue, changingCell } = params;
    const maxIterations = params.maxIterations || 100;
    const tolerance = params.tolerance || 0.001;
    
    // Get current values
    let currentInput = this.getCellValue(changingCell) as number;
    let currentOutput = this.getCellValue(targetCell) as number;
    
    // Newton-Raphson method with fallback to bisection
    let iteration = 0;
    let prevInput = currentInput;
    let prevOutput = currentOutput;
    
    // Initial perturbation to estimate derivative
    const delta = Math.abs(currentInput) * 0.001 || 0.001;
    
    while (iteration < maxIterations) {
      const error = currentOutput - targetValue;
      
      if (Math.abs(error) < tolerance) {
        return {
          success: true,
          solution: currentInput,
          iterations: iteration,
          finalError: error
        };
      }
      
      // Estimate derivative
      this.setCellValue(changingCell, currentInput + delta);
      await this.recalculate();
      const perturbedOutput = this.getCellValue(targetCell) as number;
      
      const derivative = (perturbedOutput - currentOutput) / delta;
      
      if (Math.abs(derivative) < 1e-10) {
        // Derivative too small, try bisection
        return this.bisectionFallback(params, currentInput);
      }
      
      // Newton-Raphson step
      prevInput = currentInput;
      prevOutput = currentOutput;
      currentInput = currentInput - error / derivative;
      
      // Apply new input
      this.setCellValue(changingCell, currentInput);
      await this.recalculate();
      currentOutput = this.getCellValue(targetCell) as number;
      
      iteration++;
    }
    
    return {
      success: false,
      solution: currentInput,
      iterations: iteration,
      finalError: currentOutput - targetValue
    };
  }
}
```

### Scenario Manager

```typescript
interface Scenario {
  id: string;
  name: string;
  changingCells: CellRef[];
  values: Record<string, CellValue>;  // Cell address -> value
  created: Date;
  createdBy: string;
  comment?: string;
}

class ScenarioManager {
  private scenarios: Map<string, Scenario> = new Map();
  private currentScenario?: string;
  private baseValues: Map<string, CellValue> = new Map();
  
  createScenario(name: string, cells: CellRef[], values: CellValue[]): Scenario {
    const scenario: Scenario = {
      id: crypto.randomUUID(),
      name,
      changingCells: cells,
      values: Object.fromEntries(cells.map((c, i) => [cellToAddress(c), values[i]])),
      created: new Date(),
      createdBy: this.currentUser
    };
    
    this.scenarios.set(scenario.id, scenario);
    return scenario;
  }
  
  async applyScenario(scenarioId: string): Promise<void> {
    const scenario = this.scenarios.get(scenarioId);
    if (!scenario) throw new Error("Scenario not found");
    
    // Save base values if not saved
    if (this.baseValues.size === 0) {
      for (const cell of scenario.changingCells) {
        this.baseValues.set(cellToAddress(cell), this.getCellValue(cell));
      }
    }
    
    // Apply scenario values
    for (const [address, value] of Object.entries(scenario.values)) {
      this.setCellValue(addressToCell(address), value);
    }
    
    this.currentScenario = scenarioId;
    await this.recalculate();
  }
  
  async generateSummaryReport(
    resultCells: CellRef[],
    scenarioIds: string[]
  ): Promise<SummaryReport> {
    const results: Record<string, Record<string, CellValue>> = {};
    
    // Get base results
    await this.restoreBase();
    results["Base"] = {};
    for (const cell of resultCells) {
      results["Base"][cellToAddress(cell)] = this.getCellValue(cell);
    }
    
    // Get scenario results
    for (const id of scenarioIds) {
      await this.applyScenario(id);
      const scenario = this.scenarios.get(id)!;
      results[scenario.name] = {};
      for (const cell of resultCells) {
        results[scenario.name][cellToAddress(cell)] = this.getCellValue(cell);
      }
    }
    
    await this.restoreBase();
    
    return { changingCells: scenarioIds[0] ? this.scenarios.get(scenarioIds[0])!.changingCells : [], resultCells, results };
  }
}
```

### Solver

```typescript
interface SolverProblem {
  objective: CellRef;
  objectiveType: "maximize" | "minimize" | "targetValue";
  targetValue?: number;
  changingCells: CellRef[];
  constraints: Constraint[];
  options: SolverOptions;
}

interface Constraint {
  cell: CellRef;
  relation: "<=" | ">=" | "=" | "int" | "bin";
  value?: number;
}

interface SolverOptions {
  maxTime: number;
  precision: number;
  assumeNonNegative: boolean;
  method: "simplex" | "grg" | "evolutionary";
}

class Solver {
  async solve(problem: SolverProblem): Promise<SolverResult> {
    switch (problem.options.method) {
      case "simplex":
        return this.solveSimplex(problem);
      case "grg":
        return this.solveGRG(problem);
      case "evolutionary":
        return this.solveEvolutionary(problem);
    }
  }
  
  private async solveGRG(problem: SolverProblem): Promise<SolverResult> {
    // Generalized Reduced Gradient method for nonlinear problems
    const vars = problem.changingCells;
    const n = vars.length;
    
    // Initialize
    let x = vars.map(v => this.getCellValue(v) as number);
    let bestObjective = await this.evaluateObjective(problem, x);
    
    const maxIterations = 1000;
    const tolerance = problem.options.precision;
    
    for (let iter = 0; iter < maxIterations; iter++) {
      // Compute gradient
      const gradient = await this.computeGradient(problem, x);
      
      // Compute reduced gradient (accounting for constraints)
      const reducedGradient = this.computeReducedGradient(gradient, x, problem.constraints);
      
      // Check convergence
      if (this.norm(reducedGradient) < tolerance) {
        return {
          success: true,
          solution: x,
          objectiveValue: bestObjective,
          iterations: iter
        };
      }
      
      // Line search
      const direction = problem.objectiveType === "maximize" 
        ? reducedGradient 
        : reducedGradient.map(g => -g);
      
      const stepSize = await this.lineSearch(problem, x, direction);
      
      // Update
      x = x.map((xi, i) => xi + stepSize * direction[i]);
      
      // Apply constraints
      x = this.projectToFeasible(x, problem.constraints);
      
      // Update cells and recalculate
      for (let i = 0; i < n; i++) {
        this.setCellValue(vars[i], x[i]);
      }
      await this.recalculate();
      
      bestObjective = await this.evaluateObjective(problem, x);
    }
    
    return {
      success: false,
      solution: x,
      objectiveValue: bestObjective,
      iterations: maxIterations
    };
  }
}
```

---

## Monte Carlo Simulation

```typescript
interface SimulationConfig {
  iterations: number;
  inputDistributions: InputDistribution[];
  outputCells: CellRef[];
  seed?: number;
  correlations?: CorrelationMatrix;
}

interface InputDistribution {
  cell: CellRef;
  distribution: Distribution;
}

type Distribution =
  | { type: "normal"; mean: number; stdDev: number }
  | { type: "uniform"; min: number; max: number }
  | { type: "triangular"; min: number; mode: number; max: number }
  | { type: "lognormal"; mean: number; stdDev: number }
  | { type: "discrete"; values: number[]; probabilities: number[] }
  | { type: "beta"; alpha: number; beta: number; min?: number; max?: number }
  | { type: "exponential"; rate: number }
  | { type: "poisson"; lambda: number };

class MonteCarloEngine {
  async runSimulation(config: SimulationConfig): Promise<SimulationResult> {
    const results: SimulationIteration[] = [];
    const rng = new SeededRandom(config.seed);
    
    // Generate correlated random numbers if needed
    const correlatedSamples = config.correlations
      ? this.generateCorrelatedSamples(config, rng)
      : null;
    
    for (let i = 0; i < config.iterations; i++) {
      // Generate input values
      const inputs: Record<string, number> = {};
      
      for (let j = 0; j < config.inputDistributions.length; j++) {
        const { cell, distribution } = config.inputDistributions[j];
        const address = cellToAddress(cell);
        
        if (correlatedSamples) {
          inputs[address] = correlatedSamples[i][j];
        } else {
          inputs[address] = this.sampleDistribution(distribution, rng);
        }
        
        this.setCellValue(cell, inputs[address]);
      }
      
      // Recalculate
      await this.recalculate();
      
      // Collect outputs
      const outputs: Record<string, number> = {};
      for (const cell of config.outputCells) {
        outputs[cellToAddress(cell)] = this.getCellValue(cell) as number;
      }
      
      results.push({ iteration: i, inputs, outputs });
      
      // Progress callback
      if (i % 100 === 0) {
        this.reportProgress(i / config.iterations);
      }
    }
    
    return this.analyzeResults(results, config);
  }
  
  private generateCorrelatedSamples(
    config: SimulationConfig,
    rng: SeededRandom
  ): number[][] {
    const n = config.iterations;
    const k = config.inputDistributions.length;
    const corr = config.correlations!;
    
    // Cholesky decomposition of correlation matrix
    const L = this.choleskyDecomposition(corr.matrix);
    
    // Generate independent standard normal samples
    const Z: number[][] = [];
    for (let i = 0; i < n; i++) {
      Z.push(Array(k).fill(0).map(() => this.standardNormal(rng)));
    }
    
    // Apply correlation
    const correlatedZ: number[][] = [];
    for (let i = 0; i < n; i++) {
      const row = Array(k).fill(0);
      for (let j = 0; j < k; j++) {
        for (let m = 0; m <= j; m++) {
          row[j] += L[j][m] * Z[i][m];
        }
      }
      correlatedZ.push(row);
    }
    
    // Transform to target distributions
    const samples: number[][] = [];
    for (let i = 0; i < n; i++) {
      const row: number[] = [];
      for (let j = 0; j < k; j++) {
        const u = this.normalCDF(correlatedZ[i][j]);
        row.push(this.inverseDistribution(config.inputDistributions[j].distribution, u));
      }
      samples.push(row);
    }
    
    return samples;
  }
  
  private analyzeResults(
    results: SimulationIteration[],
    config: SimulationConfig
  ): SimulationResult {
    const outputStats: Record<string, OutputStatistics> = {};
    
    for (const cell of config.outputCells) {
      const address = cellToAddress(cell);
      const values = results.map(r => r.outputs[address]);
      
      outputStats[address] = {
        mean: this.mean(values),
        median: this.median(values),
        stdDev: this.stdDev(values),
        min: Math.min(...values),
        max: Math.max(...values),
        percentiles: {
          5: this.percentile(values, 5),
          10: this.percentile(values, 10),
          25: this.percentile(values, 25),
          75: this.percentile(values, 75),
          90: this.percentile(values, 90),
          95: this.percentile(values, 95)
        },
        histogram: this.buildHistogram(values, 50)
      };
    }
    
    return {
      iterations: results.length,
      outputStats,
      rawResults: results
    };
  }
}
```

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
