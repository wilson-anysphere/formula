/**
 * Query-aware scoring helpers for selecting the most relevant table / region for a natural
 * language question.
 *
 * These helpers are intentionally policy-light and deterministic so `ContextManager` (or
 * downstream callers) can opt-in without locking in a specific retrieval strategy.
 */

/**
 * @typedef {"table" | "dataRegion"} RegionType
 * @typedef {{ type: RegionType, index: number }} RegionRef
 *
 * @typedef {import("./schema.js").SheetSchema} SheetSchema
 * @typedef {import("./schema.js").TableSchema} TableSchema
 * @typedef {import("./schema.js").DataRegionSchema} DataRegionSchema
 */

/**
 * Very small stopword list to keep token overlap focused on schema vocabulary.
 * This must remain deterministic and offline.
 */
const STOPWORDS = new Set([
  "a",
  "an",
  "and",
  "are",
  "as",
  "at",
  "by",
  "for",
  "from",
  "in",
  "is",
  "of",
  "on",
  "or",
  "per",
  "the",
  "to",
  "vs",
  "with",
]);

/**
 * @param {string} token
 */
function normalizeToken(token) {
  token = token.toLowerCase();
  // Simple plural handling ("costs" -> "cost") without pulling in a stemmer.
  if (
    token.length > 3 &&
    token.endsWith("s") &&
    !token.endsWith("ss") &&
    !token.endsWith("us") &&
    !token.endsWith("is")
  ) {
    token = token.slice(0, -1);
  }
  return token;
}

/**
 * @param {string} text
 * @param {{ dropStopwords?: boolean }} [options]
 */
function tokenize(text, options = {}) {
  if (typeof text !== "string") return [];
  // Split common schema identifier styles to improve overlap:
  // - "RevenueByRegion" -> "Revenue By Region"
  // - "revenue_by_region" -> "revenue by region"
  // - "Revenue2024" -> "Revenue 2024"
  const expanded = text
    .replace(/[_-]+/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    // "ABCDef" -> "ABC Def" (acronym boundary)
    .replace(/([A-Z])([A-Z][a-z])/g, "$1 $2")
    .replace(/([a-zA-Z])(\d)/g, "$1 $2")
    .replace(/(\d)([a-zA-Z])/g, "$1 $2");

  const raw = expanded.toLowerCase().match(/[a-z0-9]+/g) ?? [];
  const out = [];
  const seen = new Set();
  for (const t of raw) {
    const token = normalizeToken(t);
    if (!token) continue;
    if (options.dropStopwords && STOPWORDS.has(token)) continue;
    if (seen.has(token)) continue;
    seen.add(token);
    out.push(token);
  }
  return out;
}

/**
 * @param {string} header
 */
function isFallbackHeaderName(header) {
  const h = String(header ?? "").trim().toLowerCase();
  return /^column\d+$/.test(h);
}

/**
 * @param {RegionRef | TableSchema | DataRegionSchema | null | undefined} region
 * @param {SheetSchema | null | undefined} schema
 * @returns {{ type: RegionType, region: TableSchema | DataRegionSchema } | null}
 */
function resolveRegion(region, schema) {
  if (!region || typeof region !== "object") return null;

  // RegionRef form: `{ type, index }`
  if (
    "type" in region &&
    "index" in region &&
    (region.type === "table" || region.type === "dataRegion") &&
    Number.isInteger(region.index)
  ) {
    if (!schema || typeof schema !== "object") return null;
    if (region.type === "table") {
      const t = schema.tables?.[region.index];
      if (!t) return null;
      return { type: "table", region: t };
    }
    const r = schema.dataRegions?.[region.index];
    if (!r) return null;
    return { type: "dataRegion", region: r };
  }

  // Heuristic: TableSchema has `columns`; DataRegionSchema has `headers`.
  if ("columns" in region) {
    return { type: "table", region: /** @type {TableSchema} */ (region) };
  }
  if ("headers" in region) {
    return { type: "dataRegion", region: /** @type {DataRegionSchema} */ (region) };
  }

  return null;
}

/**
 * Score a table / data region for a given query.
 *
 * Higher is better. `0` means no match.
 *
 * Heuristics (deliberately simple and deterministic):
 *  - token overlap between query words and table names / headers
 *  - bonus for exact header matches (all header tokens appear in the query)
 *  - optional penalty for very small / header-only regions
 *
 * @param {RegionRef | TableSchema | DataRegionSchema | null | undefined} region
 * @param {SheetSchema | null | undefined} schema
 * @param {string} query
 * @returns {number}
 */
export function scoreRegionForQuery(region, schema, query) {
  if (typeof query !== "string") return 0;
  const queryTokens = tokenize(query, { dropStopwords: true });
  if (queryTokens.length === 0) return 0;
  const querySet = new Set(queryTokens);

  const resolved = resolveRegion(region, schema);
  if (!resolved) return 0;

  const { type, region: resolvedRegion } = resolved;

  /** @type {string[]} */
  const headerStrings = [];
  /** @type {string} */
  let name = "";
  /** @type {number} */
  let rowCount = 0;
  /** @type {number} */
  let columnCount = 0;
  /** @type {boolean} */
  let hasHeader = true;

  if (type === "table") {
    const table = /** @type {TableSchema} */ (resolvedRegion);
    name = table.name ?? "";
    rowCount = Number.isFinite(table.rowCount) ? table.rowCount : 0;
    columnCount = Array.isArray(table.columns) ? table.columns.length : 0;
    for (const col of table.columns ?? []) {
      headerStrings.push(col?.name ?? "");
    }
    // TableSchema always exposes column names; treat it as header-like for scoring.
    hasHeader = true;
  } else {
    const dataRegion = /** @type {DataRegionSchema} */ (resolvedRegion);
    rowCount = Number.isFinite(dataRegion.rowCount) ? dataRegion.rowCount : 0;
    columnCount = Number.isFinite(dataRegion.columnCount) ? dataRegion.columnCount : 0;
    hasHeader = Boolean(dataRegion.hasHeader);
    headerStrings.push(...(dataRegion.headers ?? []).map((h) => String(h ?? "")));
  }

  let score = 0;

  // --- Table name overlap ---
  if (name) {
    for (const t of tokenize(name, { dropStopwords: true })) {
      if (querySet.has(t)) score += 1;
    }
  }

  // --- Header token overlap + exact header bonuses ---
  const seenHeaderTokens = new Set();
  for (const header of headerStrings) {
    if (!header) continue;
    if (isFallbackHeaderName(header)) continue;

    const headerTokens = tokenize(header, { dropStopwords: true });
    if (headerTokens.length === 0) continue;

    for (const t of headerTokens) {
      if (seenHeaderTokens.has(t)) continue;
      seenHeaderTokens.add(t);
      if (querySet.has(t)) score += 2;
    }

    // Exact header match bonus: all header tokens appear in the query (order-independent).
    // This is particularly helpful for single-word headers like "Cost".
    let allTokensPresent = true;
    for (const t of headerTokens) {
      if (!querySet.has(t)) {
        allTokensPresent = false;
        break;
      }
    }
    if (allTokensPresent) score += 3;
  }

  // --- Small-region penalty ---
  const effectiveColumnCount = Math.max(1, columnCount);
  const estimatedCells = (Math.max(0, rowCount) + (hasHeader ? 1 : 0)) * effectiveColumnCount;

  if (rowCount <= 0) {
    // Header-only regions are common for titles and should not dominate.
    score -= 5;
  } else if (estimatedCells <= 4) {
    score -= 1;
  }

  if (!Number.isFinite(score)) return 0;
  // Ensure `0` consistently means "no match".
  return Math.max(0, score);
}

/**
 * Pick the best matching table / data region for a query.
 *
 * Returns null when no candidate receives a positive score.
 *
 * @param {SheetSchema | null | undefined} sheetSchema
 * @param {string} query
 * @returns {{ type: RegionType, index: number, range: string } | null}
 */
export function pickBestRegionForQuery(sheetSchema, query) {
  if (!sheetSchema || typeof sheetSchema !== "object") return null;

  /** @type {{ type: RegionType, index: number, range: string, score: number } | null} */
  let best = null;

  /**
   * @param {{ type: RegionType, index: number, range: string }} candidate
   * @param {number} score
   */
  function consider(candidate, score) {
    if (!Number.isFinite(score)) return;
    if (best === null) {
      best = { ...candidate, score };
      return;
    }

    if (score > best.score) {
      best = { ...candidate, score };
      return;
    }
    if (score < best.score) return;

    // Deterministic tie-break:
    // 1) prefer tables (they carry richer schema metadata)
    // 2) prefer earlier indices (stable ordering from schema extraction)
    // 3) prefer lexicographically smallest range (final deterministic fallback)
    if (candidate.type !== best.type) {
      const candidateRank = candidate.type === "table" ? 0 : 1;
      const bestRank = best.type === "table" ? 0 : 1;
      if (candidateRank < bestRank) best = { ...candidate, score };
      return;
    }
    if (candidate.index !== best.index) {
      if (candidate.index < best.index) best = { ...candidate, score };
      return;
    }
    if (candidate.range < best.range) best = { ...candidate, score };
  }

  for (let i = 0; i < (sheetSchema.tables?.length ?? 0); i++) {
    const table = sheetSchema.tables[i];
    const range = table?.range ?? "";
    const score = scoreRegionForQuery({ type: "table", index: i }, sheetSchema, query);
    consider({ type: "table", index: i, range }, score);
  }

  for (let i = 0; i < (sheetSchema.dataRegions?.length ?? 0); i++) {
    const region = sheetSchema.dataRegions[i];
    const range = region?.range ?? "";
    const score = scoreRegionForQuery({ type: "dataRegion", index: i }, sheetSchema, query);
    consider({ type: "dataRegion", index: i, range }, score);
  }

  if (!best || best.score <= 0) return null;
  return { type: best.type, index: best.index, range: best.range };
}
