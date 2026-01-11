import { DataTable } from "../table.js";
import { parseCsv, parseCsvCell } from "./file.js";

/**
 * @typedef {import("./types.js").ConnectorExecuteOptions} ConnectorExecuteOptions
 * @typedef {import("./types.js").ConnectorResult} ConnectorResult
 */

/**
 * @typedef {Object} HttpConnectorRequest
 * @property {string} url
 * @property {string | undefined} [method]
 * @property {Record<string, string> | undefined} [headers]
 * @property {"auto" | "json" | "csv" | "text" | undefined} [responseType]
 *   How to interpret the response body. Defaults to auto-detect.
 * @property {string | undefined} [jsonPath] Optional JSON path to select an array/object from a larger payload.
 */

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
      return DataTable.fromGrid(/** @type {unknown[][]} */ (json), { hasHeaders: true, inferTypes: true });
    }

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
    const keys = Object.keys(json);
    const columns = keys.map((name) => ({ name, type: "any" }));
    // @ts-ignore - runtime access
    const row = keys.map((k) => json[k] ?? null);
    return new DataTable(columns, [row]);
  }

  return new DataTable([{ name: "Value", type: "any" }], [[json]]);
}

/**
 * @typedef {Object} HttpConnectorOptions
 * @property {typeof fetch | undefined} [fetch]
 * @property {((url: string, options: { method: string; headers?: Record<string, string>; signal?: AbortSignal; credentials?: unknown }) => Promise<DataTable>) | undefined} [fetchTable]
 *   Backwards compatible adapter used by the early prototype.
 *   If provided, it is used instead of `fetch` and is expected to return a DataTable directly.
 */

export class HttpConnector {
  /**
   * @param {HttpConnectorOptions} [options]
   */
  constructor(options = {}) {
    this.id = "http";
    this.permissionKind = "http:request";
    this.fetchFn = options.fetch ?? (typeof fetch === "function" ? fetch.bind(globalThis) : null);
    this.fetchTable = options.fetchTable ?? null;
  }

  /**
   * @param {HttpConnectorRequest} request
   * @returns {unknown}
   */
  getCacheKey(request) {
    return {
      connector: "http",
      url: request.url,
      method: (request.method ?? "GET").toUpperCase(),
      headers: request.headers ?? {},
      responseType: request.responseType ?? "auto",
      jsonPath: request.jsonPath ?? "",
    };
  }

  /**
   * @param {HttpConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<ConnectorResult>}
   */
  async execute(request, options = {}) {
    const now = options.now ?? (() => Date.now());
    const signal = options.signal;

    /** @type {Record<string, string>} */
    const headers = { ...(request.headers ?? {}) };

    let credentials = options.credentials;
    if (
      credentials &&
      typeof credentials === "object" &&
      !Array.isArray(credentials) &&
      // @ts-ignore - runtime access
      typeof credentials.getSecret === "function"
    ) {
      // Credential handle convention: hosts can return an object with a stable
      // credentialId plus a `getSecret()` method. This keeps secret retrieval
      // inside the connector execution path (and out of cache key hashing).
      // @ts-ignore - runtime call
      credentials = await credentials.getSecret();
    }

    if (credentials && typeof credentials === "object" && !Array.isArray(credentials)) {
      // Generic convention: host apps can return `{ headers }` as credentials for HTTP APIs.
      // @ts-ignore - runtime merge
      const extraHeaders = credentials.headers;
      if (extraHeaders && typeof extraHeaders === "object") {
        Object.assign(headers, extraHeaders);
      }
    }

    const method = (request.method ?? "GET").toUpperCase();

    let table;
    /** @type {Date | undefined} */
    let sourceTimestamp;

    if (this.fetchTable) {
      table = await this.fetchTable(request.url, { method, headers, signal, credentials });
    } else {
      if (!this.fetchFn) {
        throw new Error("HTTP source requires either a global fetch implementation or an HttpConnector fetch adapter");
      }

      const response = await this.fetchFn(request.url, { method, headers, signal });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status} for ${request.url}`);
      }

      const lastModified = response.headers.get("last-modified");
      if (lastModified) {
        const parsed = new Date(lastModified);
        if (!Number.isNaN(parsed.getTime())) sourceTimestamp = parsed;
      }

      const responseType = request.responseType ?? "auto";
      const contentType = response.headers.get("content-type") ?? "";
      const effectiveType =
        responseType !== "auto"
          ? responseType
          : contentType.includes("text/csv")
            ? "csv"
            : contentType.includes("application/json") || contentType.includes("+json")
              ? "json"
              : "text";

      if (effectiveType === "csv") {
        const text = await response.text();
        const rows = parseCsv(text, {});
        const grid = rows.map((r) => r.map(parseCsvCell));
        table = DataTable.fromGrid(grid, { hasHeaders: true, inferTypes: true });
      } else if (effectiveType === "json") {
        const json = await response.json();
        const selected = jsonPathSelect(json, request.jsonPath ?? "");
        table = tableFromJson(selected);
      } else {
        const text = await response.text();
        table = DataTable.fromGrid([["Value"], [text]], { hasHeaders: true, inferTypes: false });
      }
    }

    return {
      table,
      meta: {
        refreshedAt: new Date(now()),
        sourceTimestamp,
        schema: { columns: table.columns, inferred: true },
        rowCount: table.rows.length,
        rowCountEstimate: table.rows.length,
        provenance: { kind: "http", url: request.url, method },
      },
    };
  }
}
