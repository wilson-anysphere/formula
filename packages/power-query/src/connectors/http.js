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
 * @property {{ type: "oauth2"; providerId: string; scopes?: string[] } | undefined} [auth]
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
 * @property {import("../oauth2/manager.js").OAuth2Manager | undefined} [oauth2Manager]
 *   Optional OAuth2 manager used when requests specify `request.auth.type === "oauth2"`
 *   or when credential hooks return `{ oauth2: { providerId, scopes? } }`.
 * @property {number[] | undefined} [oauth2RetryStatusCodes]
 *   HTTP status codes that should trigger a single refresh + retry when using OAuth2 auth.
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
    this.oauth2Manager = options.oauth2Manager ?? null;
    this.oauth2RetryStatusCodes = Array.isArray(options.oauth2RetryStatusCodes) ? options.oauth2RetryStatusCodes : [401, 403];
  }

  /**
   * Lightweight source-state probe for cache validation.
   *
   * Uses an HTTP HEAD request (when `fetch` is available) to capture `etag` and
   * `last-modified` headers.
   *
   * @param {HttpConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<import("./types.js").SourceState>}
   */
  async getSourceState(request, options = {}) {
    const now = options.now ?? (() => Date.now());
    const signal = options.signal;
    if (signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      throw err;
    }

    if (!this.fetchFn) return {};

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

    /** @type {{ providerId: string; scopes?: string[] } | null} */
    let credentialOAuth2 = null;
    if (credentials && typeof credentials === "object" && !Array.isArray(credentials)) {
      // @ts-ignore - runtime merge
      const extraHeaders = credentials.headers;
      if (extraHeaders && typeof extraHeaders === "object") {
        Object.assign(headers, extraHeaders);
      }

      // @ts-ignore - runtime access
      const oauth2 = credentials.oauth2;
      if (oauth2 && typeof oauth2 === "object") {
        // @ts-ignore - runtime access
        const providerId = oauth2.providerId;
        // @ts-ignore - runtime access
        const scopes = oauth2.scopes;
        if (typeof providerId === "string" && providerId) {
          credentialOAuth2 = { providerId, scopes: Array.isArray(scopes) ? scopes : undefined };
        }
      }
    }

    /** @type {{ type: "oauth2"; providerId: string; scopes?: string[] } | null} */
    let oauth2Auth = null;
    if (request.auth?.type === "oauth2") {
      oauth2Auth = request.auth;
    } else if (credentialOAuth2) {
      oauth2Auth = { type: "oauth2", ...credentialOAuth2 };
    }

    const applyOAuthHeader = async (forceRefresh = false) => {
      if (!oauth2Auth) return;
      if (!this.oauth2Manager) {
        throw new Error("HTTP OAuth2 requests require configuring HttpConnector with an OAuth2Manager");
      }
      const token = await this.oauth2Manager.getAccessToken({
        providerId: oauth2Auth.providerId,
        scopes: oauth2Auth.scopes,
        signal,
        now,
        forceRefresh,
      });
      headers.Authorization = `Bearer ${token.accessToken}`;
    };

    await applyOAuthHeader(false);

    let response;
    try {
      response = await this.fetchFn(request.url, { method: "HEAD", headers, signal });
    } catch {
      return {};
    }

    if (!response.ok && oauth2Auth && this.oauth2RetryStatusCodes.includes(response.status)) {
      await applyOAuthHeader(true);
      try {
        response = await this.fetchFn(request.url, { method: "HEAD", headers, signal });
      } catch {
        return {};
      }
    }

    if (!response.ok) return {};

    const etag = response.headers.get("etag") ?? undefined;

    /** @type {Date | undefined} */
    let sourceTimestamp;
    const lastModified = response.headers.get("last-modified");
    if (lastModified) {
      const parsed = new Date(lastModified);
      if (!Number.isNaN(parsed.getTime())) sourceTimestamp = parsed;
    }

    return { etag, sourceTimestamp };
  }

  /**
   * @param {HttpConnectorRequest} request
   * @returns {unknown}
   */
  getCacheKey(request) {
    const normalizedScopes = (scopes) => {
      if (!Array.isArray(scopes)) return [];
      const cleaned = scopes
        .filter((s) => typeof s === "string")
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
      cleaned.sort();
      return Array.from(new Set(cleaned));
    };
    const auth =
      request.auth?.type === "oauth2"
        ? { type: "oauth2", providerId: request.auth.providerId, scopes: normalizedScopes(request.auth.scopes) }
        : null;
    const key = {
      connector: "http",
      url: request.url,
      method: (request.method ?? "GET").toUpperCase(),
      headers: request.headers ?? {},
      responseType: request.responseType ?? "auto",
      jsonPath: request.jsonPath ?? "",
    };
    // Avoid changing cache keys for unauthenticated callers by only including
    // the new `auth` field when present.
    if (auth) {
      // @ts-ignore - stable JSON shape
      key.auth = auth;
    }
    return key;
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

    /** @type {{ providerId: string; scopes?: string[] } | null} */
    let credentialOAuth2 = null;
    if (credentials && typeof credentials === "object" && !Array.isArray(credentials)) {
      // Generic convention: host apps can return `{ headers }` as credentials for HTTP APIs.
      // @ts-ignore - runtime merge
      const extraHeaders = credentials.headers;
      if (extraHeaders && typeof extraHeaders === "object") {
        Object.assign(headers, extraHeaders);
      }

      // Standard convention: host apps can return `{ oauth2: { providerId, scopes? } }`.
      // @ts-ignore - runtime access
      const oauth2 = credentials.oauth2;
      if (oauth2 && typeof oauth2 === "object") {
        // @ts-ignore - runtime access
        const providerId = oauth2.providerId;
        // @ts-ignore - runtime access
        const scopes = oauth2.scopes;
        if (typeof providerId === "string" && providerId) {
          credentialOAuth2 = { providerId, scopes: Array.isArray(scopes) ? scopes : undefined };
        }
      }
    }

    const method = (request.method ?? "GET").toUpperCase();

    /** @type {{ type: "oauth2"; providerId: string; scopes?: string[] } | null} */
    let oauth2Auth = null;
    if (request.auth?.type === "oauth2") {
      oauth2Auth = request.auth;
    } else if (credentialOAuth2) {
      oauth2Auth = { type: "oauth2", ...credentialOAuth2 };
    }

    const applyOAuthHeader = async (forceRefresh = false) => {
      if (!oauth2Auth) return;
      if (!this.oauth2Manager) {
        throw new Error("HTTP OAuth2 requests require configuring HttpConnector with an OAuth2Manager");
      }
      const token = await this.oauth2Manager.getAccessToken({
        providerId: oauth2Auth.providerId,
        scopes: oauth2Auth.scopes,
        signal,
        now,
        forceRefresh,
      });
      headers.Authorization = `Bearer ${token.accessToken}`;
    };

    await applyOAuthHeader(false);

    let table;
    /** @type {Date | undefined} */
    let sourceTimestamp;

    if (this.fetchTable) {
      const shouldRetry = (err) => {
        if (!oauth2Auth) return false;
        if (!err || (typeof err !== "object" && typeof err !== "function")) return false;
        // @ts-ignore - best-effort status extraction for host adapters
        const status = err.status ?? err.response?.status ?? null;
        return typeof status === "number" && this.oauth2RetryStatusCodes.includes(status);
      };

      try {
        table = await this.fetchTable(request.url, { method, headers, signal, credentials });
      } catch (err) {
        if (shouldRetry(err)) {
          await applyOAuthHeader(true);
          table = await this.fetchTable(request.url, { method, headers, signal, credentials });
        } else {
          throw err;
        }
      }
    } else {
      if (!this.fetchFn) {
        throw new Error("HTTP source requires either a global fetch implementation or an HttpConnector fetch adapter");
      }

      const fetchOnce = async () => {
        const response = await this.fetchFn(request.url, { method, headers, signal });
        return response;
      };

      let response = await fetchOnce();
      if (!response.ok && oauth2Auth && this.oauth2RetryStatusCodes.includes(response.status)) {
        await applyOAuthHeader(true);
        response = await fetchOnce();
      }

      if (!response.ok) throw new Error(`HTTP ${response.status} for ${request.url}`);

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
