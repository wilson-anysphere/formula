import { DataTable } from "../table.js";

/**
 * @typedef {import("./types.js").Connector} Connector
 * @typedef {import("./types.js").ConnectorExecuteOptions} ConnectorExecuteOptions
 * @typedef {import("./types.js").ConnectorResult} ConnectorResult
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
 *   format: "csv" | "json" | "parquet";
 *   path: string;
 *   csv?: { delimiter?: string; hasHeaders?: boolean };
 *   json?: { jsonPath?: string };
 * }} FileConnectorRequest
 */

/**
 * @typedef {{
 *   readText?: (path: string) => Promise<string>;
 *   readParquetTable?: (path: string, options?: { signal?: AbortSignal }) => Promise<DataTable>;
 *   stat?: (path: string) => Promise<{ mtimeMs: number }>;
 * }} FileConnectorOptions
 */

export class FileConnector {
  /**
   * @param {FileConnectorOptions} [options]
   */
  constructor(options = {}) {
    this.id = "file";
    this.permissionKind = "file:read";
    this.readText = options.readText ?? null;
    this.readParquetTable = options.readParquetTable ?? null;
    this.stat = options.stat ?? null;
  }

  /**
   * Lightweight source-state probe for cache validation.
   *
   * @param {FileConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<import("./types.js").SourceState>}
   */
  async getSourceState(request, options = {}) {
    const signal = options.signal;
    if (signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      throw err;
    }

    if (!this.stat) return {};
    const result = await this.stat(request.path);
    if (signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      throw err;
    }

    const mtimeMs = result?.mtimeMs;
    if (typeof mtimeMs !== "number" || !Number.isFinite(mtimeMs)) return {};
    return { sourceTimestamp: new Date(mtimeMs) };
  }

  /**
   * @param {FileConnectorRequest} request
   * @returns {unknown}
   */
  getCacheKey(request) {
    if (request.format === "csv") {
      return { connector: "file", format: "csv", path: request.path, csv: request.csv ?? {} };
    }
    if (request.format === "json") {
      return { connector: "file", format: "json", path: request.path, json: request.json ?? {} };
    }
    if (request.format === "parquet") {
      return { connector: "file", format: "parquet", path: request.path };
    }
    /** @type {never} */
    const exhausted = request.format;
    throw new Error(`Unsupported file format '${exhausted}'`);
  }

  /**
   * @param {FileConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<ConnectorResult>}
   */
  async execute(request, options = {}) {
    const now = options.now ?? (() => Date.now());
    const signal = options.signal;
    if (signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      throw err;
    }

    if (request.format === "csv") {
      if (!this.readText) {
        throw new Error("CSV source requires a FileConnector readText adapter");
      }
      const text = await this.readText(request.path);
      if (signal?.aborted) {
        const err = new Error("Aborted");
        err.name = "AbortError";
        throw err;
      }
      const rows = parseCsv(text, { delimiter: request.csv?.delimiter });
      const grid = rows.map((r) => r.map(parseCsvCell));
      const hasHeaders = request.csv?.hasHeaders ?? true;
      const table = DataTable.fromGrid(grid, { hasHeaders, inferTypes: true });
      return {
        table,
        meta: {
          refreshedAt: new Date(now()),
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rows.length,
          rowCountEstimate: table.rows.length,
          provenance: { kind: "file", path: request.path, format: "csv" },
        },
      };
    }

    if (request.format === "json") {
      if (!this.readText) {
        throw new Error("JSON source requires a FileConnector readText adapter");
      }
      const text = await this.readText(request.path);
      if (signal?.aborted) {
        const err = new Error("Aborted");
        err.name = "AbortError";
        throw err;
      }
      const parsed = JSON.parse(text);
      const selected = jsonPathSelect(parsed, request.json?.jsonPath ?? "");
      const table = tableFromJson(selected);
      return {
        table,
        meta: {
          refreshedAt: new Date(now()),
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rows.length,
          rowCountEstimate: table.rows.length,
          provenance: { kind: "file", path: request.path, format: "json", jsonPath: request.json?.jsonPath ?? "" },
        },
      };
    }

    if (request.format === "parquet") {
      if (!this.readParquetTable) {
        throw new Error("Parquet source requires a FileConnector readParquetTable adapter");
      }
      const table = await this.readParquetTable(request.path, { signal });
      return {
        table,
        meta: {
          refreshedAt: new Date(now()),
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rows.length,
          rowCountEstimate: table.rows.length,
          provenance: { kind: "file", path: request.path, format: "parquet" },
        },
      };
    }

    /** @type {never} */
    const exhausted = request;
    throw new Error(`Unsupported file request '${String(exhausted)}'`);
  }
}
