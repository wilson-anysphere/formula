import { rectIntersectionArea, rectSize } from "../workbook/rect.js";

const DEFAULT_KIND_BOOST = Object.freeze({
  table: 0.08,
  namedRange: 0.06,
  dataRegion: 0.03,
  formulaRegion: 0.0,
});

/**
 * Tokenize user input into a stable list of lowercase terms.
 * @param {string} text
 * @returns {string[]}
 */
function tokenize(text) {
  const raw = String(text ?? "");
  // Insert separators so lexical matching behaves similarly to HashEmbedder tokenization:
  // - treat underscores as separators
  // - split camelCase/PascalCase and digit boundaries
  const separated = raw
    .replace(/_/g, " ")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    .replace(/([A-Za-z])([0-9])/g, "$1 $2")
    .replace(/([0-9])([A-Za-z])/g, "$1 $2");

  const tokens = separated.toLowerCase().split(/[^a-z0-9]+/g).filter(Boolean);
  if (tokens.length === 0) return [];
  // De-dupe while preserving first-seen order.
  const seen = new Set();
  const out = [];
  for (const t of tokens) {
    if (t.length < 2) continue;
    if (seen.has(t)) continue;
    seen.add(t);
    out.push(t);
  }
  return out;
}

/**
 * @param {string[]} tokens
 * @param {string} haystack
 */
function countTokenMatches(tokens, haystack) {
  if (!tokens.length) return 0;
  const s = String(haystack ?? "").toLowerCase();
  let count = 0;
  for (const t of tokens) {
    if (!t) continue;
    if (s.includes(t)) count += 1;
  }
  return count;
}

/**
 * Penalize extremely large chunks so we prefer concise context when scores are close.
 * @param {number} tokenCount
 * @param {{
 *   threshold: number,
 *   maxPenalty: number,
 *   scale: number
 * }} opts
 */
function tokenCountPenalty(tokenCount, opts) {
  if (!Number.isFinite(tokenCount)) return 0;
  if (tokenCount <= opts.threshold) return 0;
  const over = tokenCount - opts.threshold;
  if (over <= 0) return 0;
  // Linear ramp to `maxPenalty` at `threshold + scale`.
  return Math.min(opts.maxPenalty, (over / opts.scale) * opts.maxPenalty);
}

/**
 * Heuristic re-ranking for workbook chunk search results.
 *
 * This is intentionally deterministic and "explainable": the goal is to improve
 * retrieval quality for hash-based embeddings (which can produce clustered
 * similarity scores) by applying small, stable adjustments based on metadata.
 *
 * Note: the returned results overwrite `score` with the adjusted score so callers
 * can treat it as a single relevance value (including downstream dedupe logic).
 *
 * @template {{ id: string, score: number, metadata?: any }} T
 * @param {string} query
 * @param {T[]} results
 * @param {{
 *   kindBoost?: Partial<Record<string, number>>,
 *   titleTokenBoost?: number,
 *   sheetTokenBoost?: number,
 *   tokenPenaltyThreshold?: number,
 *   tokenPenaltyScale?: number,
 *   tokenPenaltyMax?: number,
 * }} [opts]
 * @returns {T[]}
 */
export function rerankWorkbookResults(query, results, opts = {}) {
  const tokens = tokenize(query);
  const kindBoost = { ...DEFAULT_KIND_BOOST, ...(opts.kindBoost ?? {}) };
  const titleTokenBoost = opts.titleTokenBoost ?? 0.04;
  const sheetTokenBoost = opts.sheetTokenBoost ?? 0.02;
  const tokenPenaltyThreshold = opts.tokenPenaltyThreshold ?? 400;
  const tokenPenaltyScale = opts.tokenPenaltyScale ?? 1200;
  const tokenPenaltyMax = opts.tokenPenaltyMax ?? 0.15;

  const decorated = results.map((r, index) => {
    const baseScore = Number.isFinite(r?.score) ? r.score : 0;
    const meta = r?.metadata ?? {};
    const kind = String(meta.kind ?? "");
    const title = String(meta.title ?? "");
    const sheetName = String(meta.sheetName ?? "");
    const tokenCount = Number(meta.tokenCount ?? 0);

    let adjustedScore = baseScore;
    adjustedScore += kindBoost[kind] ?? 0;
    adjustedScore += countTokenMatches(tokens, title) * titleTokenBoost;
    adjustedScore += countTokenMatches(tokens, sheetName) * sheetTokenBoost;
    adjustedScore -= tokenCountPenalty(tokenCount, {
      threshold: tokenPenaltyThreshold,
      scale: tokenPenaltyScale,
      maxPenalty: tokenPenaltyMax,
    });

    // Normalize any NaN/Infinity so sorting is stable.
    if (!Number.isFinite(adjustedScore)) adjustedScore = -Infinity;

    return { index, id: String(r?.id ?? ""), baseScore, adjustedScore, value: r };
  });

  decorated.sort((a, b) => {
    if (b.adjustedScore !== a.adjustedScore) return b.adjustedScore - a.adjustedScore;
    if (b.baseScore !== a.baseScore) return b.baseScore - a.baseScore;
    if (a.id < b.id) return -1;
    if (a.id > b.id) return 1;
    return a.index - b.index;
  });

  return decorated.map((d) => ({ ...d.value, score: d.adjustedScore }));
}

/**
 * @param {any} rect
 * @returns {rect is { r0: number, c0: number, r1: number, c1: number }}
 */
function isValidRect(rect) {
  if (!rect || typeof rect !== "object") return false;
  const { r0, c0, r1, c1 } = rect;
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return false;
  if (r1 < r0 || c1 < c0) return false;
  return true;
}

/**
 * Drop overlapping workbook chunks (same workbook + sheet) that are near-duplicates.
 *
 * This is intended to run after `rerankWorkbookResults` so `result.score` reflects
 * the adjusted score, but it will work with any numeric `score` field.
 *
 * Notes:
 * - Near-duplicate detection uses rect intersection / min-area overlap ratio.
 * - Deduping is scoped to (workbookId, sheetName) when workbook ids are present.
 * - Duplicate ids are always dropped (highest score wins).
 *
 * @template {{ id: string, score: number, metadata?: any }} T
 * @param {T[]} results
 * @param {{
 *   overlapRatioThreshold?: number,
 * }} [opts]
 * @returns {T[]}
 */
export function dedupeOverlappingResults(results, opts = {}) {
  const overlapRatioThreshold = opts.overlapRatioThreshold ?? 0.8;

  const decorated = results.map((r, index) => {
    const score = Number.isFinite(r?.score) ? r.score : 0;
    return { index, id: String(r?.id ?? ""), score, value: r };
  });

  // Ensure we keep the highest-scoring result when near-duplicates overlap.
  decorated.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (a.id < b.id) return -1;
    if (a.id > b.id) return 1;
    return a.index - b.index;
  });

  /** @type {typeof decorated} */
  const kept = [];
  const seenIds = new Set();

  for (const cand of decorated) {
    if (seenIds.has(cand.id)) continue;
    seenIds.add(cand.id);

    const meta = cand.value?.metadata ?? {};
    // `workbookId` was introduced later in the metadata schema. Preserve backwards
    // compatibility: if it is missing/empty, treat it as a single implicit group so
    // callers can still dedupe within a sheet.
    const workbookId = typeof meta.workbookId === "string" && meta.workbookId ? meta.workbookId : null;
    const sheetName = meta.sheetName;
    const rect = meta.rect;

    let isDup = false;
    if (typeof sheetName === "string" && sheetName && isValidRect(rect)) {
      for (const prev of kept) {
        const prevMeta = prev.value?.metadata ?? {};
        const prevWorkbookId =
          typeof prevMeta.workbookId === "string" && prevMeta.workbookId ? prevMeta.workbookId : null;
        if (prevWorkbookId !== workbookId) continue;
        if (prevMeta.sheetName !== sheetName) continue;
        const prevRect = prevMeta.rect;
        if (!isValidRect(prevRect)) continue;
        const inter = rectIntersectionArea(rect, prevRect);
        if (inter === 0) continue;
        const ratio = inter / Math.min(rectSize(rect), rectSize(prevRect));
        if (ratio > overlapRatioThreshold) {
          isDup = true;
          break;
        }
      }
    }

    if (!isDup) kept.push(cand);
  }

  return kept.map((k) => k.value);
}
