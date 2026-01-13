/**
 * Compact, deterministic summarization helpers for workbook-level schema output from
 * {@link extractWorkbookSchema}.
 *
 * This mirrors the style of `summarizeSheetSchema()` but operates on workbook metadata
 * (sheets + tables + named ranges).
 *
 * @typedef {import("./workbookSchema.js").WorkbookSchemaSummary} WorkbookSchemaSummary
 * @typedef {import("./schema.js").InferredType} InferredType
 *
 * @typedef {{
 *   /** Maximum number of sheet names to include in the sheet-name list. *\/
 *   maxSheets?: number,
 *   /** Maximum number of tables to include. *\/
 *   maxTables?: number,
 *   /** Maximum number of named ranges to include. *\/
 *   maxNamedRanges?: number,
 *   /** Maximum number of headers to include per table. *\/
 *   maxHeadersPerTable?: number,
 *   /** Maximum number of types to include per table (defaults to `maxHeadersPerTable`). *\/
 *   maxTypesPerTable?: number,
 *   /** Include the sheet list line (default true). *\/
 *   includeSheets?: boolean,
 *   /** Include table lines (default true). *\/
 *   includeTables?: boolean,
 *   /** Include named range lines (default true). *\/
 *   includeNamedRanges?: boolean,
 * }} SummarizeWorkbookOptions
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
 * @param {SummarizeWorkbookOptions | undefined} options
 * @returns {Required<SummarizeWorkbookOptions>}
 */
function normalizeOptions(options) {
  const maxSheets = normalizeLimit(options?.maxSheets);
  const maxTables = normalizeLimit(options?.maxTables);
  const maxNamedRanges = normalizeLimit(options?.maxNamedRanges);
  const maxHeadersPerTable = normalizeLimit(options?.maxHeadersPerTable);
  const maxTypesPerTable = normalizeLimit(options?.maxTypesPerTable);

  return {
    maxSheets: maxSheets ?? 20,
    maxTables: maxTables ?? 20,
    maxNamedRanges: maxNamedRanges ?? 20,
    maxHeadersPerTable: maxHeadersPerTable ?? 8,
    maxTypesPerTable: maxTypesPerTable ?? (maxHeadersPerTable ?? 8),
    includeSheets: options?.includeSheets ?? true,
    includeTables: options?.includeTables ?? true,
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
 * Produce a compact, deterministic, schema-first summary of a workbook.
 *
 * The intent is to include stable, high-signal schema details (table headers/types,
 * named ranges) without needing to serialize full JSON into prompts.
 *
 * @param {WorkbookSchemaSummary} schema
 * @param {SummarizeWorkbookOptions} [options]
 * @returns {string}
 */
export function summarizeWorkbookSchema(schema, options = {}) {
  const opts = normalizeOptions(options);
  const workbookId = typeof schema?.id === "string" ? schema.id : "";

  const sheets = Array.isArray(schema?.sheets) ? schema.sheets : [];
  const tables = Array.isArray(schema?.tables) ? schema.tables : [];
  const namedRanges = Array.isArray(schema?.namedRanges) ? schema.namedRanges : [];

  const sheetNames = sheets
    .map((s) => (typeof s?.name === "string" ? s.name : ""))
    .filter((s) => s !== "")
    .slice()
    .sort((a, b) => a.localeCompare(b));

  const out = [];
  out.push(`workbook=${bracket(workbookId)} sheets=${sheetNames.length} tables=${tables.length} named=${namedRanges.length}`);

  if (opts.includeSheets) {
    out.push(`s=${formatList(sheetNames, opts.maxSheets)}`);
  }

  if (opts.includeTables) {
    const sortedTables = tables
      .slice()
      .sort(
        (a, b) =>
          String(a?.sheetName ?? "").localeCompare(String(b?.sheetName ?? "")) ||
          String(a?.rangeA1 ?? "").localeCompare(String(b?.rangeA1 ?? "")) ||
          String(a?.name ?? "").localeCompare(String(b?.name ?? "")),
      );
    const limit = Math.min(sortedTables.length, opts.maxTables);
    for (let i = 0; i < limit; i++) {
      const t = sortedTables[i];
      const name = typeof t?.name === "string" ? t.name : "";
      const range = typeof t?.rangeA1 === "string" ? t.rangeA1 : "";
      const rowCount = Number.isFinite(t?.rowCount) ? Math.max(0, Math.floor(t.rowCount)) : 0;
      const colCount = Number.isFinite(t?.columnCount) ? Math.max(0, Math.floor(t.columnCount)) : 0;
      const headers = Array.isArray(t?.headers) ? t.headers : [];
      const types = Array.isArray(t?.inferredColumnTypes) ? t.inferredColumnTypes : /** @type {InferredType[]} */ ([]);
      out.push(
        `T${i + 1} ${bracket(name)} r=${bracket(range)} rows=${rowCount} cols=${colCount} h=${formatList(
          headers,
          opts.maxHeadersPerTable,
        )} t=${formatList(types, opts.maxTypesPerTable)}`,
      );
    }
  }

  if (opts.includeNamedRanges) {
    const sortedRanges = namedRanges
      .slice()
      .sort(
        (a, b) =>
          String(a?.sheetName ?? "").localeCompare(String(b?.sheetName ?? "")) ||
          String(a?.rangeA1 ?? "").localeCompare(String(b?.rangeA1 ?? "")) ||
          String(a?.name ?? "").localeCompare(String(b?.name ?? "")),
      );
    const limit = Math.min(sortedRanges.length, opts.maxNamedRanges);
    for (let i = 0; i < limit; i++) {
      const nr = sortedRanges[i];
      const name = typeof nr?.name === "string" ? nr.name : "";
      const range = typeof nr?.rangeA1 === "string" ? nr.rangeA1 : "";
      out.push(`N${i + 1} ${bracket(name)} r=${bracket(range)}`);
    }
  }

  return out.join("\n");
}

