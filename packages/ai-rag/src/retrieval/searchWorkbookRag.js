import { dedupeOverlappingResults, rerankWorkbookResults } from "./ranking.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * Await a promise but reject early if the AbortSignal is triggered.
 *
 * This cannot cancel underlying work, but it ensures callers can stop waiting
 * promptly when a request is canceled.
 *
 * @template T
 * @param {Promise<T> | T} promise
 * @param {AbortSignal | undefined} signal
 * @returns {Promise<T>}
 */
function awaitWithAbort(promise, signal) {
  if (!signal) return Promise.resolve(promise);
  if (signal.aborted) return Promise.reject(createAbortError());

  return new Promise((resolve, reject) => {
    const onAbort = () => reject(createAbortError());
    signal.addEventListener("abort", onAbort, { once: true });

    Promise.resolve(promise).then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      }
    );
  });
}

/**
 * High-level helper for workbook RAG retrieval:
 * - Embed query
 * - Query vector store
 * - Optional rerank + dedupe
 *
 * @param {{
 *   queryText: string,
 *   workbookId?: string,
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
  const workbookId = params?.workbookId;
  const topK = params?.topK ?? 8;
  const vectorStore = params?.vectorStore;
  const embedder = params?.embedder;
  const rerank = params?.rerank ?? true;
  const dedupe = params?.dedupe ?? true;

  if (!queryText.trim()) return [];
  if (!vectorStore || typeof vectorStore.query !== "function") {
    throw new Error("searchWorkbookRag requires a vectorStore with a query() method");
  }
  if (!embedder || typeof embedder.embedTexts !== "function") {
    throw new Error("searchWorkbookRag requires an embedder with an embedTexts() method");
  }

  if (!Number.isFinite(topK) || topK <= 0) {
    throw new Error(`searchWorkbookRag requires a positive topK (got ${topK})`);
  }

  // When we rerank or dedupe we want a larger candidate set to work with.
  const oversample = rerank || dedupe ? 4 : 1;
  const queryK = Math.max(topK, Math.ceil(topK * oversample));

  const [qVec] = await awaitWithAbort(embedder.embedTexts([queryText], { signal }), signal);
  throwIfAborted(signal);
  if (!qVec) return [];

  /** @type {any[]} */
  let results = await awaitWithAbort(vectorStore.query(qVec, queryK, { workbookId, signal }), signal);
  throwIfAborted(signal);
  if (!Array.isArray(results)) results = [];

  if (rerank) {
    results = rerankWorkbookResults({ queryText, results });
  }
  if (dedupe) {
    results = dedupeOverlappingResults({ results });
  }

  return results.slice(0, topK);
}

