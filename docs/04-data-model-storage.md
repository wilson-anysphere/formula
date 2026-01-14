# Data Model & Storage

## Overview

The data model must support both the traditional sparse spreadsheet pattern (most cells empty) and massive datasets (millions of rows). This requires a hybrid approach: sparse storage for typical spreadsheet use and columnar compression for analytical workloads.

---

## Cell Storage Architecture

### Sparse HashMap Storage

Most spreadsheets are sparse—users don't fill every cell. Efficient storage:

```typescript
interface CellStore {
  // Primary storage: only non-empty cells
  cells: Map<CellId, Cell>;
  
  // Metadata
  usedRange: Range;  // Bounding box of non-empty cells
  rowCount: number;  // Explicitly set row count (may exceed used range)
  colCount: number;  // Explicitly set column count
}

type CellId = string;  // "row,col" or encoded number

interface Cell {
  value: CellValue;
  formula?: string;
  ast?: ASTNode;      // Parsed formula (lazy)
  style?: number;     // Index into style table
  comment?: Comment;
  hyperlink?: Hyperlink;
  validation?: DataValidation;
}

type CellValue = 
  | number 
  | string 
  | boolean 
  | ErrorValue 
  | null           // Empty
  | RichText       // Formatted text
  | ArrayValue;    // Spilled array marker
```

### Cell ID Encoding

For faster hashing and comparison:

```typescript
// Option 1: String key (simple, readable)
function cellIdString(row: number, col: number): string {
  return `${row},${col}`;
}

// Option 2: Number encoding (faster comparison)
// Supports up to 2^20 rows (~1M) and 2^12 cols (~4K)
function cellIdNumber(row: number, col: number): number {
  return (row << 12) | col;
}

// Option 3: BigInt for unlimited range
function cellIdBigInt(row: number, col: number): bigint {
  return (BigInt(row) << 32n) | BigInt(col);
}
```

### Row-Based Storage Alternative

For sheets with mostly filled rows:

```typescript
interface RowStore {
  rows: Map<number, Row>;
}

interface Row {
  cells: Map<number, Cell>;  // col -> cell
  height?: number;
  style?: number;
  hidden?: boolean;
}
```

---

## Columnar Storage (Data Model / Power Pivot)

For large datasets and analytical queries, use VertiPaq-style columnar compression:

PivotTables that target the Data Model (PowerPivot) are computed via `formula-dax`; see
[ADR-0005: PivotTables ownership and data flow across crates](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md)
for the end-to-end ownership boundaries and refresh flows.

### Column Types

```typescript
type ColumnType = 
  | "number"
  | "string"
  | "boolean"
  | "datetime"
  | "currency"
  | "percentage";

interface Column {
  name: string;
  type: ColumnType;
  storage: ColumnStorage;
  statistics: ColumnStats;
}

interface ColumnStats {
  distinctCount: number;
  nullCount: number;
  min?: number | string;
  max?: number | string;
  sum?: number;
  avgLength?: number;  // For strings
}
```

### Calculated Columns (Materialized)

Calculated columns in the Data Model follow Power Pivot semantics: the expression is evaluated **once per row** and the results are **materialized** into the table as a new physical column. After materialization, the column behaves like any other stored column for filtering, grouping, joins, and storage accounting.

Key properties:

- **Row-wise evaluation, stored results:** `formula-dax` evaluates the expression in a row context; the resulting values are stored in `formula-columnar` alongside imported columns.
- **Immutable snapshots:** `formula-columnar::ColumnarTable` is immutable. Adding a calculated column produces a new table snapshot by **appending a newly encoded column** while **reusing the existing encoded columns unchanged** (no re-encoding of prior columns).
- **Single logical type (excluding blanks):** to support efficient encoding, a columnar calculated column must resolve to a single logical type across all *non-blank* rows: **number**, **string**, or **boolean**. Blanks are allowed and are encoded as nulls/validity bits, but mixed non-blank types are rejected.
- **(Optional) streaming materialization:** for very large tables, the encoder can operate in a streaming/batched mode (page-sized chunks): evaluate the expression for a batch of rows, encode that batch immediately, and emit chunks incrementally to avoid holding the entire computed column in memory.

### Encoding Strategies

VertiPaq uses three encoding algorithms in order of preference:

#### 1. Value Encoding (for integers)

```typescript
// If column has discoverable mathematical relationship
// Example: [10, 15, 30, 45] -> store offsets from min
// min=10, values stored as [0, 5, 20, 35]

interface ValueEncoding {
  type: "value";
  minValue: number;
  bitWidth: number;  // Bits needed per offset
  data: Uint8Array | Uint16Array | Uint32Array;
}

function encodeValues(values: number[]): ValueEncoding {
  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = max - min;
  const bitWidth = Math.ceil(Math.log2(range + 1));
  
  // Pack into appropriate typed array
  const ArrayType = bitWidth <= 8 ? Uint8Array : 
                   bitWidth <= 16 ? Uint16Array : Uint32Array;
  const data = new ArrayType(values.length);
  
  for (let i = 0; i < values.length; i++) {
    data[i] = values[i] - min;
  }
  
  return { type: "value", minValue: min, bitWidth, data };
}
```

#### 2. Dictionary Encoding (for strings and low-cardinality)

```typescript
// Create lookup table of distinct values
// Store indices instead of actual values

interface DictionaryEncoding {
  type: "dictionary";
  dictionary: string[];          // Distinct values
  indices: Uint8Array | Uint16Array | Uint32Array;  // Index per row
}

function encodeDictionary(values: string[]): DictionaryEncoding {
  const uniqueValues = [...new Set(values)];
  const valueToIndex = new Map(uniqueValues.map((v, i) => [v, i]));
  
  const bitWidth = Math.ceil(Math.log2(uniqueValues.length));
  const ArrayType = bitWidth <= 8 ? Uint8Array : 
                   bitWidth <= 16 ? Uint16Array : Uint32Array;
  const indices = new ArrayType(values.length);
  
  for (let i = 0; i < values.length; i++) {
    indices[i] = valueToIndex.get(values[i])!;
  }
  
  return { type: "dictionary", dictionary: uniqueValues, indices };
}
```

#### 3. Run-Length Encoding (for consecutive duplicates)

```typescript
// Applied after other encoding if beneficial
// Store (value, startRow, count) for runs

interface RLEEncoding {
  type: "rle";
  runs: Array<{ value: any; start: number; count: number }>;
}

function encodeRLE(values: any[]): RLEEncoding | null {
  const runs: Array<{ value: any; start: number; count: number }> = [];
  let currentRun = { value: values[0], start: 0, count: 1 };
  
  for (let i = 1; i < values.length; i++) {
    if (values[i] === currentRun.value) {
      currentRun.count++;
    } else {
      runs.push(currentRun);
      currentRun = { value: values[i], start: i, count: 1 };
    }
  }
  runs.push(currentRun);
  
  // Only use RLE if it actually compresses
  const rleSize = runs.length * 16;  // Rough estimate
  const rawSize = values.length * 8;
  
  return rleSize < rawSize * 0.7 ? { type: "rle", runs } : null;
}
```

### Compression Results

Typical compression ratios: **7-10x**

| Data Type | Raw Size | Compressed | Ratio |
|-----------|----------|------------|-------|
| Dates (1 year) | 8 bytes | 2 bits | 32x |
| Categories (10 values) | 8 bytes | 4 bits | 16x |
| Currency ($0-$1M) | 8 bytes | 20 bits | 3.2x |
| Free text | 50 bytes | 50 bytes | 1x |

---

## Relational Tables

### Table Definition

```typescript
interface Table {
  name: string;
  displayName: string;
  range: Range;
  columns: TableColumn[];
  style?: TableStyle;
  
  // Relational features
  primaryKey?: string;           // Column name
  relationships: Relationship[];
  
  // Features
  headerRow: boolean;
  totalsRow: boolean;
  autoFilter: boolean;
}

interface TableColumn {
  name: string;
  type: ColumnType;
  formula?: string;      // Calculated column
  totalsFormula?: string;
  format?: string;
}

interface Relationship {
  name: string;
  fromTable: string;
  fromColumn: string;
  toTable: string;
  toColumn: string;
  cardinality: "one-to-one" | "one-to-many" | "many-to-many";
  crossFilterDirection: "single" | "both";
}
```

#### Many-to-many relationships (`formula-dax` semantics)

`formula-dax` treats `"many-to-many"` relationships as **distinct-key**
relationships for filter propagation: when a relationship propagates a filter from one table to the
other, it does so by taking the distinct set of visible key values on the source side and applying
that set to the target side (conceptually similar to `TREATAS(VALUES(source[key]), target[key])`).

Row-context navigation follows the relationship orientation (`fromTable` → `toTable`):

- `RELATED(toTable[column])` is only defined when the current `fromTable[fromColumn]` value matches
  **at most one** row in `toTable[toColumn]`. If there are multiple matches (a common case for
  many-to-many), the engine treats it as an ambiguity error.
- `RELATEDTABLE(otherTable)` returns the set of matching rows, and can naturally return multiple
  rows for many-to-many relationships.

Pivot/group-by note: grouping/pivoting by columns across a many-to-many relationship **expands** a
base row into multiple related rows. When a key matches multiple rows on the `toTable` side, the
engine treats the group key as a set of possible values and produces groups for all combinations
(Cartesian product across group-by columns). This can duplicate measure contributions (a single fact
row can contribute to multiple groups), so summing group totals can “double count” compared to an
ungrouped total.

### Referential Integrity

```typescript
class RelationshipManager {
  private tables: Map<string, Table>;
  private relationships: Relationship[];
  
  validateRelationship(rel: Relationship): ValidationResult {
    const fromTable = this.tables.get(rel.fromTable);
    const toTable = this.tables.get(rel.toTable);
    
    if (!fromTable || !toTable) {
      return { valid: false, error: "Table not found" };
    }
    
    // Check column exists and types match
    const fromCol = fromTable.columns.find(c => c.name === rel.fromColumn);
    const toCol = toTable.columns.find(c => c.name === rel.toColumn);
    
    if (!fromCol || !toCol) {
      return { valid: false, error: "Column not found" };
    }
    
    if (fromCol.type !== toCol.type) {
      return { valid: false, error: "Column type mismatch" };
    }
    
    // For one-to-many, check uniqueness on "one" side
    if (rel.cardinality === "one-to-many") {
      const toValues = this.getColumnValues(toTable, rel.toColumn);
      const uniqueValues = new Set(toValues);
      if (uniqueValues.size !== toValues.length) {
        return { valid: false, error: "Target column must have unique values" };
      }
    }
    
    return { valid: true };
  }
  
  // RELATED function implementation
  getRelatedValue(
    fromRow: number,
    fromTable: string,
    relationship: Relationship
  ): CellValue {
    const fromValue = this.getCellValue(fromTable, fromRow, relationship.fromColumn);
    const toTable = this.tables.get(relationship.toTable)!;
    
    // Find matching row in target table
    const toColIndex = toTable.columns.findIndex(c => c.name === relationship.toColumn);
    for (let row = 0; row < this.getRowCount(toTable); row++) {
      if (this.getCellValue(relationship.toTable, row, relationship.toColumn) === fromValue) {
        return row;  // Return row index for further lookup
      }
    }
    
    return null;  // No match
  }
}
```

---

## Rich Data Types

### Cell Types Beyond Numbers and Text

```typescript
type RichCellValue =
  | PrimitiveValue
  | ImageValue
  | LinkedEntityValue
  | RecordValue
  | ArrayValue;

interface ImageValue {
  type: "image";
  imageId: string;       // Reference to image store
  altText?: string;
  width?: number;
  height?: number;
}

interface LinkedEntityValue {
  type: "entity";
  entityType: string;    // "stock", "geography", "currency", etc.
  entityId: string;
  displayValue: string;
  properties: Record<string, CellValue>;
}

interface RecordValue {
  type: "record";
  fields: Record<string, CellValue>;
  displayField: string;
}

interface ArrayValue {
  type: "array";
  originCell: CellRef;   // Cell containing the formula
  data: CellValue[][];
}
```

### Linked Entity Example (Stock Data)

```typescript
const stockEntity: LinkedEntityValue = {
  type: "entity",
  entityType: "stock",
  entityId: "AAPL",
  displayValue: "Apple Inc.",
  properties: {
    "Price": 178.50,
    "Change": 2.35,
    "Change%": 0.0133,
    "Volume": 52436789,
    "MarketCap": 2850000000000,
    "52WeekHigh": 199.62,
    "52WeekLow": 124.17
  }
};

// Access via formulas:
// =A1.Price       -> 178.50
// =A1.["Change%"] -> 0.0133
```

---

## Storage Layer

### Workbook & Sheet Metadata

Workbooks contain an **ordered list of sheets**. Each sheet has a **stable ID** (never changes on rename/reorder) and user-facing metadata used by the UI and XLSX round-tripping.

```typescript
type SheetVisibility = "visible" | "hidden" | "veryHidden";

interface SheetMeta {
  id: string;              // Stable internal ID (UUID)
  name: string;            // Display name (unique, case-insensitive)
  position: number;        // Ordering in workbook (0-based)
  visibility: SheetVisibility;
  tabColor?: string;       // Optional ARGB hex (e.g. "FFFF0000") or theme-based color in metadata
  xlsxSheetId?: number;    // Preserve Excel's sheetId on round-trip
  xlsxRelId?: string;      // Preserve workbook.xml relationship ID (r:id)
}

interface WorkbookMeta {
  id: string;
  name: string;
  sheets: SheetMeta[];     // Ordered by position
}
```

Key behaviors:
- Reorder updates `position` (and must persist).
- Rename updates `name` and rewrites formulas that reference the sheet.
- `veryHidden` must be preserved from XLSX even if the UI doesn’t expose it directly.

### SQLite Schema

```sql
-- Core tables
CREATE TABLE workbooks (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  metadata JSON
);

CREATE TABLE sheets (
  id TEXT PRIMARY KEY,
  workbook_id TEXT REFERENCES workbooks(id),
  name TEXT NOT NULL,
  position INTEGER,
  visibility TEXT NOT NULL DEFAULT 'visible' CHECK (visibility IN ('visible','hidden','veryHidden')),
  tab_color TEXT, -- Optional ARGB hex (e.g. 'FFFF0000'); theme/indexed colors stored in metadata JSON if needed
  xlsx_sheet_id INTEGER, -- Preserve the <sheet sheetId="..."> value for round-trip
  xlsx_rel_id TEXT,      -- Preserve the <sheet r:id="..."> relationship ID for round-trip
  frozen_rows INTEGER DEFAULT 0,
  frozen_cols INTEGER DEFAULT 0,
  zoom REAL DEFAULT 1.0,
  metadata JSON
);

CREATE TABLE cells (
  sheet_id TEXT REFERENCES sheets(id),
  row INTEGER,
  col INTEGER,
  value_type TEXT,  -- 'number', 'string', 'boolean', 'error', 'formula'
  value_number REAL,
  value_string TEXT,
  formula TEXT,
  style_id INTEGER,
  PRIMARY KEY (sheet_id, row, col)
);

-- Sparse storage: only non-empty cells stored
CREATE INDEX idx_cells_sheet ON cells(sheet_id);

-- Styles (deduplicated)
CREATE TABLE styles (
  id INTEGER PRIMARY KEY,
  font_id INTEGER REFERENCES fonts(id),
  fill_id INTEGER REFERENCES fills(id),
  border_id INTEGER REFERENCES borders(id),
  number_format TEXT,
  alignment JSON,
  protection JSON
);

-- Named ranges
CREATE TABLE named_ranges (
  workbook_id TEXT REFERENCES workbooks(id),
  name TEXT,
  scope TEXT,  -- 'workbook' or sheet name
  reference TEXT,
  PRIMARY KEY (workbook_id, name, scope)
);

-- Version history (CRDT-compatible)
CREATE TABLE change_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sheet_id TEXT REFERENCES sheets(id),
  timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  user_id TEXT,
  operation TEXT,  -- 'set_cell', 'delete_cell', 'insert_row', etc.
  target JSON,     -- Row/col/range affected
  old_value JSON,
  new_value JSON
);
```

For the Power Pivot-style Data Model, `formula-storage` persists columnar data separately using `data_model_*` tables. Calculated columns are persisted in two parts:

- **Values:** the computed column values are stored in `data_model_columns` / `data_model_chunks` like any other column (no special-case storage format).
- **Definitions:** the calculated column’s `(table, name, expression)` is stored in `data_model_calculated_columns` and loaded back into `formula-dax` as metadata via `add_calculated_column_definition` (so values are *not* recomputed during load).

### Auto-Save Strategy (Debounced Dirty-Page Flush)

```typescript
class AutoSaveManager {
  // Debounced save
  private readonly SAVE_DELAY = 1000;  // 1 second
  private readonly MAX_DELAY = 5000;   // 5 seconds max

  constructor(private readonly memory: MemoryManager) {}

  recordChange(change: Change): void {
    // Update the in-memory page cache immediately (for snappy UX).
    this.memory.recordChange(change);

    // Then debounce a flush of dirty pages to SQLite.
    this.scheduleFlush();
  }

  private async flush(): Promise<void> {
    // Transactional, ordered writeback of dirty pages.
    await this.memory.flushDirtyPages();
  }
}
```

---

## Data Validation

### Validation Types

```typescript
interface DataValidation {
  type: ValidationRuleType;
  rule: ValidationRule;
  errorStyle: "stop" | "warning" | "information";
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
  showDropdown?: boolean;
}

type ValidationRuleType = 
  | "any"
  | "whole_number"
  | "decimal"
  | "list"
  | "date"
  | "time"
  | "text_length"
  | "custom";

type ValidationRule =
  | { type: "any" }
  | { type: "whole_number" | "decimal"; operator: ComparisonOp; value1: number; value2?: number }
  | { type: "list"; values: string[] | CellRef }
  | { type: "date"; operator: ComparisonOp; value1: Date; value2?: Date }
  | { type: "text_length"; operator: ComparisonOp; value1: number; value2?: number }
  | { type: "custom"; formula: string };

type ComparisonOp = 
  | "between" | "not_between"
  | "equal" | "not_equal"
  | "greater" | "less"
  | "greater_equal" | "less_equal";
```

### Validation Execution

```typescript
class ValidationEngine {
  validate(cell: CellRef, value: CellValue, validation: DataValidation): ValidationResult {
    const rule = validation.rule;
    
    switch (rule.type) {
      case "any":
        return { valid: true };
        
      case "list":
        const allowedValues = this.resolveListValues(rule.values);
        return {
          valid: allowedValues.includes(String(value)),
          error: `Value must be one of: ${allowedValues.join(", ")}`
        };
        
      case "whole_number":
        if (typeof value !== "number" || !Number.isInteger(value)) {
          return { valid: false, error: "Value must be a whole number" };
        }
        return this.checkComparison(value, rule.operator, rule.value1, rule.value2);
        
      case "custom":
        // Evaluate custom formula with cell as context
        const result = this.evaluateFormula(rule.formula, cell, value);
        return { valid: result === true };
        
      // ... other types
    }
  }
  
  private checkComparison(
    value: number,
    op: ComparisonOp,
    v1: number,
    v2?: number
  ): ValidationResult {
    switch (op) {
      case "between":
        return { valid: value >= v1 && value <= v2! };
      case "not_between":
        return { valid: value < v1 || value > v2! };
      case "equal":
        return { valid: value === v1 };
      case "greater":
        return { valid: value > v1 };
      // ... other operators
    }
  }
}
```

---

## Memory Management

### Large Dataset Strategies

```typescript
class MemoryManager {
  // Fixed tile size; paging works on (sheet_id, page_row, page_col) keys.
  readonly rowsPerPage = 256;
  readonly colsPerPage = 256;

  // Byte-bounded LRU across pages (not whole sheets).
  readonly maxMemoryBytes = 500 * 1024 * 1024;
  readonly evictionWatermark = 0.8;
  private readonly pages = new LruCache<PageKey, PageData>();

  // Viewport-driven loading: callers request the visible range; the memory
  // manager loads any missing pages from SQLite and returns a sparse snapshot.
  async loadViewport(sheetId: string, viewport: CellRange): Promise<ViewportData> {
    const keys = pageKeysForRange(sheetId, viewport, this.rowsPerPage, this.colsPerPage);
    for (const key of keys) {
      if (!this.pages.has(key)) {
        const range = pageRange(key, this.rowsPerPage, this.colsPerPage);
        const cells = await this.db.loadCellsInRange(sheetId, range);
        this.pages.set(key, new PageData(cells));
      }
    }
    this.evictIfNeeded();
    return buildViewportData(viewport, this.pages);
  }

  // Edits mark pages dirty; dirty pages flush on eviction and via autosave.
  recordChange(change: Change): void {
    const key = pageKeyForCell(change.sheetId, change.row, change.col);
    const page = this.pages.getOrLoad(key);
    page.applyChange(change);
    page.dirty = true;
  }

  async flushDirtyPages(): Promise<void> {
    // Transactional writeback, preserving edit ordering.
    const changes = collectDirtyChangesInOrder(this.pages);
    await this.db.transaction(async (tx) => tx.applyCellChanges(changes));
    markPagesClean(this.pages);
  }

  private evictIfNeeded(): void {
    while (this.estimatedUsageBytes() > this.maxMemoryBytes * this.evictionWatermark) {
      const lru = this.pages.popLru();
      if (lru?.dirty) {
        // Flush before eviction to avoid losing edits.
        this.flushDirtyPages();
      }
    }
  }
}
```

### Streaming Large Datasets

```typescript
class StreamingDataLoader {
  async *loadRows(
    table: string,
    startRow: number,
    batchSize: number = 1000
  ): AsyncGenerator<Row[]> {
    let offset = startRow;
    
    while (true) {
      const rows = await this.db.query(
        `SELECT * FROM cells 
         WHERE sheet_id = ? AND row >= ? AND row < ?
         ORDER BY row, col`,
        [table, offset, offset + batchSize]
      );
      
      if (rows.length === 0) break;
      
      yield this.rowsToData(rows);
      offset += batchSize;
    }
  }
}
```

---

## Import/Export

### CSV Import with Type Inference

```typescript
class CSVImporter {
  async import(file: File, options: CSVOptions): Promise<SheetData> {
    const text = await file.text();
    const lines = this.parseCSV(text, options.delimiter);
    
    // Infer types from first N rows
    const sampleSize = Math.min(100, lines.length);
    const columnTypes = this.inferTypes(lines.slice(0, sampleSize));
    
    // Parse values with inferred types
    const cells = new Map<CellId, Cell>();
    
    for (let row = 0; row < lines.length; row++) {
      for (let col = 0; col < lines[row].length; col++) {
        const rawValue = lines[row][col];
        const value = this.parseValue(rawValue, columnTypes[col]);
        
        if (value !== null) {
          cells.set(cellId(row, col), { value });
        }
      }
    }
    
    return { cells, usedRange: { startRow: 0, endRow: lines.length - 1, startCol: 0, endCol: columnTypes.length - 1 } };
  }
  
  private inferTypes(rows: string[][]): ColumnType[] {
    const columnCount = Math.max(...rows.map(r => r.length));
    const types: ColumnType[] = new Array(columnCount).fill("string");
    
    for (let col = 0; col < columnCount; col++) {
      const values = rows.map(r => r[col]).filter(v => v !== "");
      
      if (values.every(v => this.isNumber(v))) {
        types[col] = "number";
      } else if (values.every(v => this.isDate(v))) {
        types[col] = "datetime";
      } else if (values.every(v => this.isBoolean(v))) {
        types[col] = "boolean";
      }
    }
    
    return types;
  }
}
```

### Parquet Support

```typescript
class ParquetLoader {
  async load(file: File): Promise<SheetData> {
    // Use arrow-js or similar library
    const buffer = await file.arrayBuffer();
    const table = await parquet.read(buffer);
    
    // Convert to cell storage
    const cells = new Map<CellId, Cell>();
    const schema = table.schema;
    
    // Header row
    for (let col = 0; col < schema.fields.length; col++) {
      cells.set(cellId(0, col), { value: schema.fields[col].name });
    }
    
    // Data rows
    for (let row = 0; row < table.numRows; row++) {
      for (let col = 0; col < schema.fields.length; col++) {
        const value = table.getColumn(col).get(row);
        if (value !== null) {
          cells.set(cellId(row + 1, col), { value: this.convertArrowValue(value) });
        }
      }
    }
    
    return { cells, usedRange: { startRow: 0, endRow: table.numRows, startCol: 0, endCol: schema.fields.length - 1 } };
  }
}
```

---

## Testing Strategy

### Data Integrity Tests

```typescript
describe("Data Model", () => {
  it("preserves data through save/load cycle", async () => {
    const original = createTestWorkbook();
    
    await storage.save(original);
    const loaded = await storage.load(original.id);
    
    expect(loaded).toDeepEqual(original);
  });
  
  it("handles sparse data efficiently", () => {
    const cells = new Map<CellId, Cell>();
    cells.set(cellId(0, 0), { value: 1 });
    cells.set(cellId(999999, 999), { value: 2 });
    
    const memoryBefore = process.memoryUsage().heapUsed;
    const store = new CellStore(cells);
    const memoryAfter = process.memoryUsage().heapUsed;
    
    // Memory should be proportional to cell count, not grid size
    expect(memoryAfter - memoryBefore).toBeLessThan(10000);  // <10KB for 2 cells
  });
});
```
