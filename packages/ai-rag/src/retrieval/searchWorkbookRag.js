import { dedupeOverlappingResults, rerankWorkbookResults } from "./ranking.js";
import { awaitWithAbort, throwIfAborted } from "../utils/abort.js";

/**
 * High-level helper for workbook RAG retrieval:
 * - Embed query
 * - Query vector store
 * - Optional rerank + dedupe
 *
 * @param {{
 *   queryText: string,
 *   workbookId: string,
 *   topK?: number,
 *   vectorStore: { query(vector: ArrayLike<number>, topK: number, opts?: { workbookId?: string, signal?: AbortSignal }): Promise<any[]> },
 *   embedder: { embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<ArrayLike<number>[]> },
 *   rerank?: boolean,
 *   dedupe?: boolean,
 *   signal?: AbortSignal,
 * }} params
 */
export async function searchWorkbookRag(params) {
  const signal = params?.signal;
  throwIfAborted(signal);

  const queryText = String(params?.queryText ?? "");
  const workbookIdRaw = params?.workbookId;
  const topKInput = params?.topK ?? 8;
  const vectorStore = params?.vectorStore;
  const embedder = params?.embedder;
  const rerank = params?.rerank ?? true;
  const dedupe = params?.dedupe ?? true;

  if (!queryText.trim()) return [];

  if (!Number.isFinite(topKInput)) {
    throw new Error(`searchWorkbookRag requires a finite topK (got ${String(topKInput)})`);
  }
  // Align with vector store semantics: floor floats and treat non-positive values as
  // "no retrieval" (return an empty list without embedding/querying).
  const topK = Math.floor(topKInput);
  if (topK <= 0) return [];

  const workbookId = typeof workbookIdRaw === "string" ? workbookIdRaw.trim() : "";
  if (!workbookId) {
    throw new Error("searchWorkbookRag requires a non-empty workbookId");
  }
  if (!vectorStore || typeof vectorStore.query !== "function") {
    throw new Error("searchWorkbookRag requires a vectorStore with a query() method");
  }
  if (!embedder || typeof embedder.embedTexts !== "function") {
    throw new Error("searchWorkbookRag requires an embedder with an embedTexts() method");
  }

  // When we rerank or dedupe we want a larger candidate set to work with.
  const oversample = rerank || dedupe ? 4 : 1;
  const queryK = Math.max(topK, Math.ceil(topK * oversample));

  const vectors = await awaitWithAbort(embedder.embedTexts([queryText], { signal }), signal);
  throwIfAborted(signal);
  if (!Array.isArray(vectors)) {
    throw new Error(
      "searchWorkbookRag embedder.embedTexts returned a non-array result; expected an array with a single vector"
    );
  }
  if (vectors.length !== 1) {
    throw new Error(
      `searchWorkbookRag embedder.embedTexts returned ${vectors.length} vector(s); expected 1`
    );
  }
  const qVec = vectors[0];
  if (!qVec) return [];
  const qLen = qVec?.length;
  if (!Number.isFinite(qLen) || qLen <= 0) {
    throw new Error(
      "searchWorkbookRag embedder.embedTexts returned an invalid query vector (expected an array-like vector with a finite length)"
    );
  }
  const expectedDim = vectorStore?.dimension;
  if (Number.isFinite(expectedDim) && qLen !== expectedDim) {
    throw new Error(
      `searchWorkbookRag query vector dimension mismatch: expected ${expectedDim}, got ${qLen}`
    );
  }
  for (let i = 0; i < qLen; i += 1) {
    const value = qVec[i];
    if (!Number.isFinite(value)) {
      throw new Error(
        `searchWorkbookRag embedder.embedTexts returned an invalid query vector value at index=${i}: expected a finite number`
      );
    }
  }

  /** @type {any[]} */
  let results = await awaitWithAbort(vectorStore.query(qVec, queryK, { workbookId, signal }), signal);
  throwIfAborted(signal);
  if (!Array.isArray(results)) results = [];
  // Defense in depth: even though we pass `workbookId` into the query options, filter the
  // returned results to avoid accidental cross-workbook leakage if a store implementation
  // ignores the option.
  results = results.filter((r) => {
    const id = r?.metadata?.workbookId;
    // If the store includes workbookId metadata, enforce it. Otherwise, assume the store
    // respected the query option and keep the result.
    return id == null || id === workbookId;
  });

  if (rerank) {
    results = rerankWorkbookResults({ queryText, results });
  }
  if (dedupe) {
    results = dedupeOverlappingResults({ results });
  }

  return results.slice(0, topK);
}
