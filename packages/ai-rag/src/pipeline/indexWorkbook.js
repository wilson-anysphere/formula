import { contentHash } from "../utils/hash.js";
import { chunkWorkbook } from "../workbook/chunkWorkbook.js";
import { chunkToText } from "../workbook/chunkToText.js";

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
 * This cannot cancel underlying work (e.g. an embedder call), but it ensures callers can
 * stop waiting promptly when a request is canceled.
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
 * @param {string} text
 */
export function approximateTokenCount(text) {
  // Heuristic: English token ~= 4 chars.
  return Math.ceil(text.length / 4);
}

/**
 * Index a workbook into a vector store, incrementally updating embeddings when
 * chunks change.
 *
 * Note: Formula's desktop workbook RAG uses deterministic, offline hash embeddings
 * (`HashEmbedder`) by default. Embeddings are not user-configurable (no API keys /
 * no local model setup). A future Cursor-managed embedding service can
 * replace this to improve retrieval quality.
 *
 * @param {{
 *   workbook: import('../workbook/workbookTypes').Workbook,
 *   vectorStore: any,
 *   embedder: { embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<ArrayLike<number>[]> },
 *   sampleRows?: number,
 *   transform?: (record: { id: string, text: string, metadata: any }) => ({ text?: string, metadata?: any } | null | Promise<{ text?: string, metadata?: any } | null>)
 *   signal?: AbortSignal,
 * }} params
 */
export async function indexWorkbook(params) {
  const signal = params.signal;
  throwIfAborted(signal);
  const { workbook, vectorStore, embedder } = params;
  const sampleRows = params.sampleRows ?? 5;
  const chunks = chunkWorkbook(workbook, { signal });
  throwIfAborted(signal);

  const existingForWorkbook = await awaitWithAbort(
    vectorStore.list({
      workbookId: workbook.id,
      includeVector: false,
      signal,
    }),
    signal
  );
  throwIfAborted(signal);
  const existingHashes = new Map(
    existingForWorkbook.map((r) => [r.id, r.metadata?.contentHash])
  );

  const currentIds = new Set();
  /** @type {{ id: string, text: string, metadata: any }[]} */
  const toUpsert = [];

  for (const chunk of chunks) {
    throwIfAborted(signal);
    const originalText = chunkToText(chunk, { sampleRows });

    /** @type {{ id: string, text: string, metadata: any }} */
    let record = {
      id: chunk.id,
      text: originalText,
      metadata: {
        workbookId: chunk.workbookId,
        sheetName: chunk.sheetName,
        kind: chunk.kind,
        title: chunk.title,
        rect: chunk.rect,
        text: originalText,
      },
    };

    if (typeof params.transform === "function") {
      throwIfAborted(signal);
      const transformed = await awaitWithAbort(params.transform(record), signal);
      throwIfAborted(signal);
      if (!transformed) continue;

      if (Object.prototype.hasOwnProperty.call(transformed, "text")) {
        record.text = transformed.text ?? "";
      }
      if (Object.prototype.hasOwnProperty.call(transformed, "metadata")) {
        record.metadata = transformed.metadata ?? record.metadata;
      }

      if (!Object.prototype.hasOwnProperty.call(record.metadata ?? {}, "text")) {
        record.metadata = { ...(record.metadata ?? {}), text: record.text };
      }
    }

    throwIfAborted(signal);
    const chunkHash = await awaitWithAbort(contentHash(record.text), signal);
    throwIfAborted(signal);
    currentIds.add(record.id);

    if (existingHashes.get(record.id) === chunkHash) continue;

    toUpsert.push({
      id: record.id,
      text: record.text,
      metadata: {
        ...(record.metadata ?? {}),
        contentHash: chunkHash,
        tokenCount: approximateTokenCount(record.text),
      },
    });
  }

  throwIfAborted(signal);
  const vectors =
    toUpsert.length > 0
      ? await awaitWithAbort(embedder.embedTexts(toUpsert.map((r) => r.text), { signal }), signal)
      : [];
  throwIfAborted(signal);

  if (toUpsert.length) {
    // Avoid aborting while awaiting persistence. Upserts are stateful; if we were to
    // reject early here, callers could start a new indexing run while the underlying
    // store is still writing.
    throwIfAborted(signal);
    await vectorStore.upsert(
      toUpsert.map((r, i) => ({
        id: r.id,
        vector: vectors[i],
        metadata: r.metadata,
      }))
    );
  }
  throwIfAborted(signal);

  // Delete stale records for this workbook.
  const staleIds = existingForWorkbook
    .map((r) => r.id)
    .filter((id) => !currentIds.has(id));
  if (staleIds.length) {
    // Avoid aborting while awaiting persistence. Deletes are stateful; if we were to
    // reject early here, callers could start a new indexing run while the underlying
    // store is still writing.
    throwIfAborted(signal);
    await vectorStore.delete(staleIds);
  }
  throwIfAborted(signal);

  return {
    totalChunks: chunks.length,
    upserted: toUpsert.length,
    skipped: chunks.length - toUpsert.length,
    deleted: staleIds.length,
  };
}
