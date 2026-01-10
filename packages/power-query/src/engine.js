import { DataTable } from "./table.js";
import { applyOperation } from "./steps.js";

/**
 * @typedef {import("./model.js").Query} Query
 * @typedef {import("./model.js").QuerySource} QuerySource
 * @typedef {import("./model.js").QueryStep} QueryStep
 * @typedef {import("./model.js").QueryOperation} QueryOperation
 */

/**
 * @typedef {{
 *   tables?: Record<string, DataTable>;
 *   queries?: Record<string, Query>;
 * }} QueryExecutionContext
 */

/**
 * @typedef {{
 *   limit?: number;
 *   // Execute up to and including this step index.
 *   maxStepIndex?: number;
 * }} ExecuteOptions
 */

/**
 * Minimal CSV parser (RFC4180-ish) with support for quoted values.
 * @param {string} text
 * @param {{ delimiter?: string }} [options]
 * @returns {string[][]}
 */
export function parseCsv(text, options = {}) {
  const delimiter = options.delimiter ?? ",";

  /** @type {string[][]} */
  const rows = [];
  /** @type {string[]} */
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i <= text.length; i++) {
    const char = i === text.length ? "\n" : text[i];

    if (inQuotes) {
      if (char === '"') {
        const next = text[i + 1];
        if (next === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
        continue;
      }
      field += char;
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      continue;
    }

    if (char === delimiter) {
      row.push(field);
      field = "";
      continue;
    }

    if (char === "\r") {
      // Ignore CR; LF will handle row endings.
      continue;
    }

    if (char === "\n") {
      row.push(field);
      field = "";
      // Ignore empty trailing row when the file ends with a newline.
      if (!(row.length === 1 && row[0] === "" && i === text.length)) {
        rows.push(row);
      }
      row = [];
      continue;
    }

    field += char;
  }

  return rows;
}

/**
 * Convert a CSV cell to a primitive.
 * @param {string} value
 * @returns {unknown}
 */
export function parseCsvCell(value) {
  const trimmed = value.trim();
  if (trimmed === "") return null;
  if (/^(?:true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === "true";
  if (/^-?\d+(?:\.\d+)?$/.test(trimmed)) return Number(trimmed);
  return trimmed;
}

/**
 * @param {unknown} input
 * @param {string} path
 * @returns {unknown}
 */
function jsonPathSelect(input, path) {
  if (!path) return input;
  const parts = path.split(".").filter(Boolean);
  let current = input;
  for (const part of parts) {
    if (current == null) return undefined;
    const bracketMatch = part.match(/^(.+)\[(\d+)\]$/);
    if (bracketMatch) {
      const prop = bracketMatch[1];
      const index = Number(bracketMatch[2]);
      // @ts-ignore - runtime traversal
      current = current[prop];
      if (!Array.isArray(current)) return undefined;
      current = current[index];
    } else {
      // @ts-ignore - runtime traversal
      current = current[part];
    }
  }
  return current;
}

/**
 * @param {unknown} json
 * @returns {DataTable}
 */
function tableFromJson(json) {
  if (Array.isArray(json)) {
    if (json.length === 0) return new DataTable([], []);
    if (Array.isArray(json[0])) {
      // Treat as grid: [[header...], [row...]]
      return DataTable.fromGrid(/** @type {unknown[][]} */ (json), { hasHeaders: true, inferTypes: true });
    }

    // Array of objects.
    /** @type {Set<string>} */
    const keySet = new Set();
    for (const row of json) {
      if (row && typeof row === "object" && !Array.isArray(row)) {
        Object.keys(row).forEach((k) => keySet.add(k));
      }
    }
    const keys = Array.from(keySet);
    const columns = keys.map((name) => ({ name, type: "any" }));
    const rows = json.map((row) => {
      if (!row || typeof row !== "object" || Array.isArray(row)) {
        return keys.map(() => null);
      }
      // @ts-ignore - runtime access
      return keys.map((k) => row[k] ?? null);
    });
    return new DataTable(columns, rows);
  }

  if (json && typeof json === "object") {
    // Single object -> one row.
    const keys = Object.keys(json);
    const columns = keys.map((name) => ({ name, type: "any" }));
    // @ts-ignore - runtime access
    const row = keys.map((k) => json[k] ?? null);
    return new DataTable(columns, [row]);
  }

  return new DataTable([{ name: "Value", type: "any" }], [[json]]);
}

/**
 * @typedef {{
 *   querySql: (connection: unknown, sql: string) => Promise<DataTable>;
 * }} DatabaseAdapter
 */

/**
 * @typedef {{
 *   fetchTable: (url: string, options: { method: string; headers?: Record<string, string> }) => Promise<DataTable>;
 * }} ApiAdapter
 */

export class QueryEngine {
  /**
   * @param {{
   *   databaseAdapter?: DatabaseAdapter,
   *   apiAdapter?: ApiAdapter,
   *   fileAdapter?: { readText: (path: string) => Promise<string> }
   * }} [options]
   */
  constructor(options = {}) {
    this.databaseAdapter = options.databaseAdapter ?? null;
    this.apiAdapter = options.apiAdapter ?? null;
    this.fileAdapter = options.fileAdapter ?? null;
  }

  /**
   * Execute a full query.
   * @param {Query} query
   * @param {QueryExecutionContext} [context]
   * @param {ExecuteOptions} [options]
   * @returns {Promise<DataTable>}
   */
  async executeQuery(query, context = {}, options = {}) {
    const sourceTable = await this.loadSource(query.source, context, new Set([query.id]));
    const maxStepIndex = options.maxStepIndex ?? query.steps.length - 1;
    const steps = query.steps.slice(0, maxStepIndex + 1);
    const table = await this.executeSteps(sourceTable, steps, context);
    const limited = options.limit != null ? table.head(options.limit) : table;
    return limited;
  }

  /**
   * Execute a list of steps starting from an already-materialized table.
   * @param {DataTable} table
   * @param {QueryStep[]} steps
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async executeSteps(table, steps, context) {
    let current = table;
    for (const step of steps) {
      current = await this.applyStep(current, step.operation, context);
    }
    return current;
  }

  /**
   * @param {DataTable} table
   * @param {QueryOperation} operation
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async applyStep(table, operation, context) {
    switch (operation.type) {
      case "merge":
        return this.mergeTables(table, operation, context);
      case "append":
        return this.appendTables(table, operation, context);
      default:
        return applyOperation(table, operation);
    }
  }

  /**
   * @param {QuerySource} source
   * @param {QueryExecutionContext} context
   * @param {Set<string>} callStack
   * @returns {Promise<DataTable>}
   */
  async loadSource(source, context, callStack) {
    switch (source.type) {
      case "range": {
        const hasHeaders = source.range.hasHeaders ?? true;
        return DataTable.fromGrid(source.range.values, { hasHeaders, inferTypes: true });
      }
      case "table": {
        if (!context.tables?.[source.table]) {
          throw new Error(`Unknown table '${source.table}'`);
        }
        return context.tables[source.table];
      }
      case "csv": {
        if (!this.fileAdapter) {
          throw new Error("CSV source requires a QueryEngine fileAdapter");
        }
        const text = await this.fileAdapter.readText(source.path);
        const rows = parseCsv(text, { delimiter: source.options?.delimiter });
        const grid = rows.map((r) => r.map(parseCsvCell));
        const hasHeaders = source.options?.hasHeaders ?? true;
        return DataTable.fromGrid(grid, { hasHeaders, inferTypes: true });
      }
      case "json": {
        if (!this.fileAdapter) {
          throw new Error("JSON source requires a QueryEngine fileAdapter");
        }
        const text = await this.fileAdapter.readText(source.path);
        const parsed = JSON.parse(text);
        const selected = jsonPathSelect(parsed, source.jsonPath ?? "");
        return tableFromJson(selected);
      }
      case "query": {
        const target = context.queries?.[source.queryId];
        if (!target) throw new Error(`Unknown query '${source.queryId}'`);
        if (callStack.has(target.id)) {
          throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${target.id}`);
        }
        const nextStack = new Set(callStack);
        nextStack.add(target.id);
        return this.executeQuery(target, context, {});
      }
      case "database": {
        if (!this.databaseAdapter) {
          throw new Error("Database source requires a QueryEngine databaseAdapter");
        }
        return this.databaseAdapter.querySql(source.connection, source.query);
      }
      case "api": {
        if (!this.apiAdapter) {
          throw new Error("API source requires a QueryEngine apiAdapter");
        }
        return this.apiAdapter.fetchTable(source.url, { method: source.method, headers: source.headers });
      }
      default: {
        /** @type {never} */
        const exhausted = source;
        throw new Error(`Unsupported source type '${exhausted.type}'`);
      }
    }
  }

  /**
   * @param {DataTable} left
   * @param {import("./model.js").MergeOp} op
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async mergeTables(left, op, context) {
    const query = context.queries?.[op.rightQuery];
    if (!query) throw new Error(`Unknown query '${op.rightQuery}'`);
    const right = await this.executeQuery(query, context, {});

    const leftKeyIdx = left.getColumnIndex(op.leftKey);
    const rightKeyIdx = right.getColumnIndex(op.rightKey);

    /** @type {Map<unknown, number[]>} */
    const rightIndex = new Map();
    right.rows.forEach((row, idx) => {
      const key = row[rightKeyIdx];
      const bucket = rightIndex.get(key);
      if (bucket) bucket.push(idx);
      else rightIndex.set(key, [idx]);
    });

    const rightColumnsToInclude = right.columns
      .map((col, idx) => ({ col, idx }))
      .filter(({ col }) => col.name !== op.rightKey || op.rightKey !== op.leftKey);

    const leftNames = new Set(left.columns.map((c) => c.name));
    const rightColumns = rightColumnsToInclude.map(({ col }) => {
      if (!leftNames.has(col.name)) return col;
      return { ...col, name: `${col.name}.right` };
    });

    const outColumns = [...left.columns, ...rightColumns];

    /** @type {unknown[][]} */
    const outRows = [];

    const emit = (leftRow, rightRow) => {
      const leftValues = leftRow ?? Array.from({ length: left.columns.length }, () => null);
      const rightValues = rightColumnsToInclude.map(({ idx }) => rightRow?.[idx] ?? null);
      outRows.push([...leftValues, ...rightValues]);
    };

    if (op.joinType === "inner" || op.joinType === "left" || op.joinType === "full") {
      /** @type {Set<number>} */
      const matchedRight = new Set();

      for (const leftRow of left.rows) {
        const matchIndices = rightIndex.get(leftRow[leftKeyIdx]) ?? [];
        if (matchIndices.length === 0) {
          if (op.joinType !== "inner") emit(leftRow, null);
          continue;
        }

        for (const rightIdx of matchIndices) {
          matchedRight.add(rightIdx);
          emit(leftRow, right.rows[rightIdx]);
        }
      }

      if (op.joinType === "full") {
        right.rows.forEach((rightRow, idx) => {
          if (!matchedRight.has(idx)) {
            emit(null, rightRow);
          }
        });
      }

      return new DataTable(outColumns, outRows);
    }

    if (op.joinType === "right") {
      /** @type {Map<unknown, number[]>} */
      const leftIndex = new Map();
      left.rows.forEach((row, idx) => {
        const key = row[leftKeyIdx];
        const bucket = leftIndex.get(key);
        if (bucket) bucket.push(idx);
        else leftIndex.set(key, [idx]);
      });

      right.rows.forEach((rightRow) => {
        const matchIndices = leftIndex.get(rightRow[rightKeyIdx]) ?? [];
        if (matchIndices.length === 0) {
          emit(null, rightRow);
          return;
        }
        for (const leftIdx of matchIndices) {
          emit(left.rows[leftIdx], rightRow);
        }
      });

      return new DataTable(outColumns, outRows);
    }

    throw new Error(`Unsupported joinType '${op.joinType}'`);
  }

  /**
   * @param {DataTable} current
   * @param {import("./model.js").AppendOp} op
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async appendTables(current, op, context) {
    const tables = [current];
    for (const id of op.queries) {
      const query = context.queries?.[id];
      if (!query) throw new Error(`Unknown query '${id}'`);
      tables.push(await this.executeQuery(query, context, {}));
    }

    /** @type {string[]} */
    const columns = [];
    /** @type {Map<string, { type: string }>} */
    const columnMeta = new Map();

    for (const table of tables) {
      for (const col of table.columns) {
        if (columnMeta.has(col.name)) continue;
        columnMeta.set(col.name, { type: col.type });
        columns.push(col.name);
      }
    }

    const outColumns = columns.map((name) => ({ name, type: columnMeta.get(name)?.type ?? "any" }));

    const outRows = [];
    for (const table of tables) {
      const index = new Map(table.columns.map((c, idx) => [c.name, idx]));
      for (const row of table.rows) {
        outRows.push(columns.map((name) => row[index.get(name) ?? -1] ?? null));
      }
    }

    return new DataTable(outColumns, outRows);
  }
}
