import {
  dedupeOverlappingResults as dedupeOverlappingResultsImpl,
  rerankWorkbookResults as rerankWorkbookResultsImpl,
} from "./rankResults.js";

/**
 * @typedef {import("../store/inMemoryVectorStore.js").VectorSearchResult} VectorSearchResult
 */

/**
 * Backwards-compatible wrapper around `retrieval/rankResults.js`.
 *
 * `searchWorkbookRag` calls this with an object parameter, but we keep the
 * underlying implementation in a standalone helper that accepts
 * `(queryText, results, opts)`.
 *
 * @param {{ queryText: string, results: VectorSearchResult[] }} params
 * @returns {VectorSearchResult[]}
 */
export function rerankWorkbookResults(params) {
  const queryText = String(params?.queryText ?? "");
  const results = Array.isArray(params?.results) ? params.results : [];
  return rerankWorkbookResultsImpl(queryText, results);
}

/**
 * Backwards-compatible wrapper around `retrieval/rankResults.js`.
 *
 * @param {{ results: VectorSearchResult[], overlapRatio?: number }} params
 * @returns {VectorSearchResult[]}
 */
export function dedupeOverlappingResults(params) {
  const results = Array.isArray(params?.results) ? params.results : [];
  const overlapRatio = params?.overlapRatio;
  return dedupeOverlappingResultsImpl(results, {
    overlapRatioThreshold: overlapRatio ?? 0.8,
  });
}
