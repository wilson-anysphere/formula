import { hashValue } from "../cache/key.js";

/**
 * @param {string} input
 * @returns {string}
 */
export function normalizeFilePath(input) {
  // Best-effort cross-platform normalization that does not rely on Node's `path`
  // module (so it can run in browsers/workers).
  let out = String(input);

  // Normalize slashes.
  out = out.replaceAll("\\", "/");
  // Preserve UNC / network path prefix (`//server/share`) by only collapsing
  // duplicate slashes after the leading `//` when present.
  if (out.startsWith("//")) {
    const rest = out.slice(2).replace(/\/{2,}/g, "/").replace(/^\/+/, "");
    out = `//${rest}`;
  } else {
    out = out.replace(/\/{2,}/g, "/");
  }

  // Lowercase Windows drive letters (C:\ -> c:/)
  out = out.replace(/^([A-Za-z]):(?=\/|$)/, (_, drive) => `${drive.toLowerCase()}:`);

  // Resolve "." / ".." segments.
  const isUnc = out.startsWith("//");
  const isAbs = isUnc || out.startsWith("/") || /^[a-z]:\//.test(out);
  const parts = out.split("/").filter((p) => p.length > 0);
  /** @type {string[]} */
  const resolved = [];
  for (const part of parts) {
    if (part === ".") continue;
    if (part === "..") {
      if (resolved.length > 0 && resolved[resolved.length - 1] !== "..") {
        resolved.pop();
      } else if (!isAbs) {
        resolved.push("..");
      }
      continue;
    }
    resolved.push(part);
  }

  const prefix = isAbs && !/^[a-z]:\//.test(out) ? (isUnc ? "//" : "/") : "";
  return prefix + resolved.join("/");
}

/**
 * Stable source id for a file source.
 * @param {string} path
 */
export function getFileSourceId(path) {
  return normalizeFilePath(path);
}

/**
 * Stable source id for an HTTP source. Returns `scheme://host:port` with an
 * explicit port even when the URL uses a default.
 *
 * @param {string} url
 */
export function getHttpSourceId(url) {
  const parsed = new URL(url);
  const scheme = parsed.protocol.toLowerCase();
  const hostname = parsed.hostname.toLowerCase();
  const defaultPort = scheme === "https:" ? "443" : scheme === "http:" ? "80" : "";
  const port = parsed.port || defaultPort;
  const host = hostname.includes(":")
    ? hostname.startsWith("[") && hostname.endsWith("]")
      ? hostname
      : `[${hostname}]`
    : hostname;
  if (port) {
    // Always include a port (explicitly) for http/https to keep the identifier stable.
    return `${scheme}//${host}:${port}`;
  }
  // Non-http(s) schemes may not have a port; omit the trailing ":".
  return `${scheme}//${host}`;
}

/**
 * Stable source id for a SQL connection. This is intentionally a *connection*
 * identity (not a per-query identity) so multiple SQL queries against the same
 * connection share a single privacy classification.
 *
 * @param {unknown} connection
 */
export function getSqlSourceId(connection) {
  if (typeof connection === "string") return connection.startsWith("sql:") ? connection : `sql:${connection}`;
  if (connection && typeof connection === "object" && !Array.isArray(connection)) {
    // Prefer a host-provided stable identifier when present.
    // @ts-ignore - runtime indexing
    const id = connection.id;
    if (typeof id === "string" && id.length > 0) return id.startsWith("sql:") ? id : `sql:${id}`;
    // @ts-ignore - runtime indexing
    const name = connection.name;
    if (typeof name === "string" && name.length > 0) return name.startsWith("sql:") ? name : `sql:${name}`;
  }
  return `sql:${hashValue(connection)}`;
}

/**
 * @param {import("../model.js").QuerySource} source
 * @returns {string | null}
 */
export function getSourceIdForQuerySource(source) {
  switch (source.type) {
    case "csv":
    case "json":
    case "parquet":
      return getFileSourceId(source.path);
    case "api":
      return getHttpSourceId(source.url);
    case "database": {
      const connectionId = source.connectionId;
      if (typeof connectionId === "string" && connectionId.length > 0) return getSqlSourceId(connectionId);
      return getSqlSourceId(source.connection);
    }
    case "range":
      return "workbook:range";
    case "table":
      return `workbook:table:${source.table}`;
    case "query":
      return null;
    default: {
      /** @type {never} */
      const exhausted = source;
      throw new Error(`Unsupported source type '${exhausted.type}'`);
    }
  }
}

/**
 * Extract a stable source id from connector provenance metadata.
 *
 * This is primarily used for cache keys and for tracking a query's "source set"
 * as it flows through transformations.
 *
 * @param {Record<string, unknown> | null | undefined} provenance
 * @returns {string | null}
 */
export function getSourceIdForProvenance(provenance) {
  if (!provenance || typeof provenance !== "object") return null;
  // @ts-ignore - runtime indexing
  const kind = provenance.kind;
  switch (kind) {
    case "file": {
      // @ts-ignore - runtime indexing
      const path = provenance.path;
      return typeof path === "string" ? getFileSourceId(path) : null;
    }
    case "http": {
      // @ts-ignore - runtime indexing
      const url = provenance.url;
      return typeof url === "string" ? getHttpSourceId(url) : null;
    }
    case "sql": {
      // SqlConnector provides `sourceId` (see `getSqlSourceId`).
      // @ts-ignore - runtime indexing
      const sourceId = provenance.sourceId;
      if (typeof sourceId === "string") return sourceId;
      // Backwards/forwards compatibility: if only a raw connection id is provided,
      // normalize it into the same `sql:<id>` namespace used elsewhere.
      // @ts-ignore - runtime indexing
      const connectionId = provenance.connectionId;
      return typeof connectionId === "string" && connectionId.length > 0 ? getSqlSourceId(connectionId) : null;
    }
    case "range":
      return "workbook:range";
    case "table": {
      // @ts-ignore - runtime indexing
      const table = provenance.table;
      return typeof table === "string" ? `workbook:table:${table}` : "workbook:table";
    }
    default:
      return null;
  }
}
