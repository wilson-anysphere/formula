import { parseA1Range } from "./a1.js";

/**
 * @typedef {import("./schema.js").SheetSchema} SheetSchema
 * @typedef {import("./schema.js").TableSchema} TableSchema
 * @typedef {import("./schema.js").DataRegionSchema} DataRegionSchema
 * @typedef {import("./schema.js").InferredType} InferredType
 *
 * @typedef {{
 *   /**
 *    * Maximum number of tables to include in the summary.
 *    *\/
 *   maxTables?: number,
 *   /**
 *    * Maximum number of data regions to include in the summary.
 *    *\/
 *   maxRegions?: number,
 *   /**
 *    * Maximum number of headers to include per table.
 *    *\/
 *   maxHeadersPerTable?: number,
 *   /**
 *    * Maximum number of types to include per table (defaults to `maxHeadersPerTable`).
 *    *\/
 *   maxTypesPerTable?: number,
 *   /**
 *    * Maximum number of headers to include per region.
 *    *\/
 *   maxHeadersPerRegion?: number,
 *   /**
 *    * Maximum number of types to include per region (defaults to `maxHeadersPerRegion`).
 *    *\/
 *   maxTypesPerRegion?: number,
 *   /**
 *    * Include table summaries (default true).
 *    *\/
 *   includeTables?: boolean,
 *   /**
 *    * Include region summaries (default true).
 *    *\/
 *   includeRegions?: boolean,
 *   /**
 *    * Maximum number of named ranges to include in the summary.
 *    *\/
 *   maxNamedRanges?: number,
 *   /**
 *    * Include named range summaries (default true).
 *    *\/
 *   includeNamedRanges?: boolean,
 * }} SummarizeSheetOptions
 */

/**
 * @param {unknown} value
 * @returns {number | null}
 */
function normalizeLimit(value) {
  if (value === undefined || value === null) return null;
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return null;
  return Math.max(0, Math.floor(n));
}

/**
 * @param {SummarizeSheetOptions | undefined} options
 * @returns {Required<SummarizeSheetOptions>}
 */
function normalizeOptions(options) {
  const maxTables = normalizeLimit(options?.maxTables);
  const maxRegions = normalizeLimit(options?.maxRegions);
  const maxNamedRanges = normalizeLimit(options?.maxNamedRanges);

  const maxHeadersPerTable = normalizeLimit(options?.maxHeadersPerTable);
  const maxTypesPerTable = normalizeLimit(options?.maxTypesPerTable);
  const maxHeadersPerRegion = normalizeLimit(options?.maxHeadersPerRegion);
  const maxTypesPerRegion = normalizeLimit(options?.maxTypesPerRegion);

  return {
    maxTables: maxTables ?? 20,
    maxRegions: maxRegions ?? 20,
    maxNamedRanges: maxNamedRanges ?? 20,
    maxHeadersPerTable: maxHeadersPerTable ?? 8,
    maxTypesPerTable: maxTypesPerTable ?? (maxHeadersPerTable ?? 8),
    maxHeadersPerRegion: maxHeadersPerRegion ?? 8,
    maxTypesPerRegion: maxTypesPerRegion ?? (maxHeadersPerRegion ?? 8),
    includeTables: options?.includeTables ?? true,
    includeRegions: options?.includeRegions ?? true,
    includeNamedRanges: options?.includeNamedRanges ?? true,
  };
}

/**
 * @param {unknown} value
 */
function cleanInline(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Escapes `]` and `|` so bracketed lists remain parseable.
 * @param {unknown} value
 */
function escapeInline(value) {
  return cleanInline(value)
    .replace(/\\/g, "\\\\")
    .replace(/\]/g, "\\]")
    .replace(/\|/g, "\\|");
}

/**
 * @param {unknown} value
 */
function bracket(value) {
  return `[${escapeInline(value)}]`;
}

/**
 * @param {unknown[]} items
 * @param {number} maxItems
 */
function formatList(items, maxItems) {
  const safeItems = Array.isArray(items) ? items.map(escapeInline) : [];
  if (safeItems.length === 0) return "[]";
  if (maxItems <= 0) return "[…]";

  const shown = safeItems.slice(0, maxItems);
  const omitted = Math.max(safeItems.length - shown.length, 0);
  if (omitted > 0) shown.push(`…+${omitted}`);
  return `[${shown.join("|")}]`;
}

/**
 * @param {string} a
 * @param {string} b
 */
function compareA1Ranges(a, b) {
  if (a === b) return 0;
  let pa = null;
  let pb = null;
  try {
    pa = parseA1Range(a);
  } catch {
    pa = null;
  }
  try {
    pb = parseA1Range(b);
  } catch {
    pb = null;
  }

  if (pa && pb) {
    const sa = pa.sheetName ?? "";
    const sb = pb.sheetName ?? "";
    if (sa !== sb) return sa.localeCompare(sb);
    return (
      (pa.startRow - pb.startRow) ||
      (pa.startCol - pb.startCol) ||
      (pa.endRow - pb.endRow) ||
      (pa.endCol - pb.endCol)
    );
  }
  if (pa && !pb) return -1;
  if (!pa && pb) return 1;
  return String(a).localeCompare(String(b));
}

/**
 * @param {TableSchema} table
 * @param {Map<string, DataRegionSchema>} regionByRange
 */
function resolveTableRegion(table, regionByRange) {
  const range = typeof table?.range === "string" ? table.range : "";
  return regionByRange.get(range) ?? null;
}

/**
 * @param {TableSchema} table
 */
function inferHasHeaderFromTable(table) {
  const cols = Array.isArray(table?.columns) ? table.columns : [];
  if (cols.length === 0) return false;
  for (let i = 0; i < cols.length; i++) {
    const expected = `Column${i + 1}`;
    const actual = typeof cols[i]?.name === "string" ? cols[i].name : "";
    if (actual !== expected) return true;
  }
  return false;
}

/**
 * @param {TableSchema} table
 * @param {Map<string, DataRegionSchema>} regionByRange
 * @param {Required<SummarizeSheetOptions>} opts
 */
function summarizeTable(table, regionByRange, opts) {
  const name = typeof table?.name === "string" ? table.name : "";
  const range = typeof table?.range === "string" ? table.range : "";
  const rowCount = Number.isFinite(table?.rowCount) ? Math.max(0, Math.floor(table.rowCount)) : 0;
  const columns = Array.isArray(table?.columns) ? table.columns : [];

  const region = resolveTableRegion(table, regionByRange);
  const hasHeader = region ? Boolean(region.hasHeader) : inferHasHeaderFromTable(table);
  const headers = region?.headers ?? columns.map((c) => c?.name ?? "");
  const types = region?.inferredColumnTypes ?? columns.map((c) => c?.type ?? "mixed");
  const colCountRaw =
    region && Number.isFinite(region.columnCount) ? Math.max(0, Math.floor(region.columnCount)) : headers.length;

  const hdrFlag = hasHeader ? 1 : 0;
  return `${bracket(name)} r=${bracket(range)} rows=${rowCount} cols=${colCountRaw} hdr=${hdrFlag} h=${formatList(
    headers,
    opts.maxHeadersPerTable,
  )} t=${formatList(types, opts.maxTypesPerTable)}`;
}

/**
 * @param {DataRegionSchema} region
 * @param {Required<SummarizeSheetOptions>} opts
 */
function summarizeDataRegion(region, opts) {
  const range = typeof region?.range === "string" ? region.range : "";
  const rowCount = Number.isFinite(region?.rowCount) ? Math.max(0, Math.floor(region.rowCount)) : 0;
  const colCount = Number.isFinite(region?.columnCount) ? Math.max(0, Math.floor(region.columnCount)) : 0;
  const hdrFlag = region?.hasHeader ? 1 : 0;
  const headers = Array.isArray(region?.headers) ? region.headers : [];
  const types = Array.isArray(region?.inferredColumnTypes) ? region.inferredColumnTypes : [];
  return `r=${bracket(range)} rows=${rowCount} cols=${colCount} hdr=${hdrFlag} h=${formatList(
    headers,
    opts.maxHeadersPerRegion,
  )} t=${formatList(types, opts.maxTypesPerRegion)}`;
}

/**
 * @param {{ name: string, range: string }} namedRange
 */
function summarizeNamedRange(namedRange) {
  const name = typeof namedRange?.name === "string" ? namedRange.name : "";
  const range = typeof namedRange?.range === "string" ? namedRange.range : "";
  return `${bracket(name)} r=${bracket(range)}`;
}

/**
 * Summarize either a `TableSchema` or `DataRegionSchema`.
 *
 * @param {TableSchema | DataRegionSchema} schemaRegionOrTable
 * @param {SummarizeSheetOptions} [options]
 * @returns {string}
 */
export function summarizeRegion(schemaRegionOrTable, options = {}) {
  const opts = normalizeOptions(options);
  const input = schemaRegionOrTable;
  if (!input || typeof input !== "object") return "";
  if (Array.isArray(/** @type {any} */ (input).columns)) {
    return `T ${summarizeTable(/** @type {TableSchema} */ (input), new Map(), opts)}`;
  }
  if (Array.isArray(/** @type {any} */ (input).inferredColumnTypes) || Array.isArray(/** @type {any} */ (input).headers)) {
    return `R ${summarizeDataRegion(/** @type {DataRegionSchema} */ (input), opts)}`;
  }
  return "";
}

/**
 * Produce a compact, deterministic, schema-first summary of a sheet.
 *
 * This is intended for LLM context where sampling raw cell values is too expensive.
 *
 * @param {SheetSchema} schema
 * @param {SummarizeSheetOptions} [options]
 * @returns {string}
 */
export function summarizeSheetSchema(schema, options = {}) {
  const opts = normalizeOptions(options);
  const sheetName = typeof schema?.name === "string" ? schema.name : "";
  const tables = Array.isArray(schema?.tables) ? schema.tables : [];
  const regions = Array.isArray(schema?.dataRegions) ? schema.dataRegions : [];
  const namedRanges = Array.isArray(schema?.namedRanges) ? schema.namedRanges : [];

  /** @type {Map<string, DataRegionSchema>} */
  const regionByRange = new Map();
  for (const region of regions) {
    if (!region || typeof region !== "object") continue;
    if (typeof region.range !== "string") continue;
    if (!regionByRange.has(region.range)) regionByRange.set(region.range, region);
  }

  const lines = [`sheet=${bracket(sheetName)} tables=${tables.length} regions=${regions.length} named=${namedRanges.length}`];

  if (opts.includeTables) {
    const ordered = tables
      .slice()
      .sort((a, b) => compareA1Ranges(String(a?.range ?? ""), String(b?.range ?? "")) || String(a?.name ?? "").localeCompare(String(b?.name ?? "")));
    const shown = ordered.slice(0, opts.maxTables);
    for (let i = 0; i < shown.length; i++) {
      lines.push(`T${i + 1} ${summarizeTable(shown[i], regionByRange, opts)}`);
    }
    const omitted = Math.max(ordered.length - shown.length, 0);
    if (omitted > 0) lines.push(`T…+${omitted}`);
  }

  if (opts.includeRegions) {
    const ordered = regions
      .slice()
      .sort((a, b) => compareA1Ranges(String(a?.range ?? ""), String(b?.range ?? "")));
    const shown = ordered.slice(0, opts.maxRegions);
    for (let i = 0; i < shown.length; i++) {
      lines.push(`R${i + 1} ${summarizeDataRegion(shown[i], opts)}`);
    }
    const omitted = Math.max(ordered.length - shown.length, 0);
    if (omitted > 0) lines.push(`R…+${omitted}`);
  }

  if (opts.includeNamedRanges) {
    const ordered = namedRanges
      .slice()
      .sort(
        (a, b) =>
          compareA1Ranges(String(a?.range ?? ""), String(b?.range ?? "")) ||
          String(a?.name ?? "").localeCompare(String(b?.name ?? "")),
      );
    const shown = ordered.slice(0, opts.maxNamedRanges);
    for (let i = 0; i < shown.length; i++) {
      lines.push(`N${i + 1} ${summarizeNamedRange(shown[i])}`);
    }
    const omitted = Math.max(ordered.length - shown.length, 0);
    if (omitted > 0) lines.push(`N…+${omitted}`);
  }

  return lines.join("\n");
}
