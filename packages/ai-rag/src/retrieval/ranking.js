import { rectIntersectionArea, rectSize } from "../workbook/rect.js";

/**
 * @typedef {import("../store/inMemoryVectorStore.js").VectorSearchResult} VectorSearchResult
 */

/**
 * Rerank workbook search results.
 *
 * This is a lightweight (and intentionally dependency-free) hook for applying
 * workbook-specific heuristics on top of raw vector similarity.
 *
 * NOTE: The default implementation is currently a no-op and preserves the input
 * order (already sorted by similarity score by the vector store). More advanced
 * reranking logic can be layered here without requiring callers to reimplement
 * the full retrieval pipeline.
 *
 * @param {{ queryText: string, results: VectorSearchResult[] }} params
 * @returns {VectorSearchResult[]}
 */
export function rerankWorkbookResults(params) {
  const results = params?.results ?? [];
  // Preserve ordering for now.
  return results;
}

/**
 * Dedupe search results that refer to overlapping workbook regions.
 *
 * Vector stores can return multiple chunks that cover the same or mostly the
 * same cells (e.g. a detected data region + a named table). For prompting,
 * including many near-duplicates wastes context budget.
 *
 * @param {{
 *   results: VectorSearchResult[],
 *   overlapRatio?: number,
 * }} params
 * @returns {VectorSearchResult[]}
 */
export function dedupeOverlappingResults(params) {
  const results = params?.results ?? [];
  const overlapRatio = params?.overlapRatio ?? 0.8;

  /** @type {VectorSearchResult[]} */
  const kept = [];
  /** @type {Set<string>} */
  const seenIds = new Set();

  for (const r of results) {
    if (!r || typeof r.id !== "string") continue;
    if (seenIds.has(r.id)) continue;
    seenIds.add(r.id);

    const meta = r.metadata ?? {};
    const rect = meta.rect;
    const sheetName = meta.sheetName;
    const workbookId = meta.workbookId;
    if (!rect || sheetName == null || workbookId == null) {
      kept.push(r);
      continue;
    }

    let overlaps = false;
    for (const ex of kept) {
      const exMeta = ex.metadata ?? {};
      if (exMeta.workbookId !== workbookId) continue;
      if (exMeta.sheetName !== sheetName) continue;
      const exRect = exMeta.rect;
      if (!exRect) continue;
      const inter = rectIntersectionArea(rect, exRect);
      if (inter === 0) continue;
      const ratio = inter / Math.min(rectSize(rect), rectSize(exRect));
      if (ratio > overlapRatio) {
        overlaps = true;
        break;
      }
    }

    if (!overlaps) kept.push(r);
  }

  return kept;
}

