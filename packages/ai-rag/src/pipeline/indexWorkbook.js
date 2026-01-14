import { contentHash } from "../utils/hash.js";
import { stableJsonStringify } from "../utils/stableJsonStringify.js";
import { awaitWithAbort, throwIfAborted } from "../utils/abort.js";
import { chunkWorkbook } from "../workbook/chunkWorkbook.js";
import { chunkToText } from "../workbook/chunkToText.js";

/**
 * Build a stable hash for record metadata.
 *
 * This intentionally excludes fields derived from the chunk text to avoid needless churn
 * and large hash payloads.
 *
 * @param {any} metadata
 */
async function computeMetadataHash(metadata) {
  const meta = metadata && typeof metadata === "object" ? metadata : {};
  /** @type {Record<string, any>} */
  const filtered = {};
  for (const [key, value] of Object.entries(meta)) {
    if (key === "contentHash" || key === "tokenCount" || key === "metadataHash" || key === "text") {
      continue;
    }
    filtered[key] = value;
  }
  return contentHash(stableJsonStringify(filtered));
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
 *   maxColumnsForSchema?: number,
 *   maxColumnsForRows?: number,
 *   tokenCount?: (text: string) => number,
 *   embedBatchSize?: number,
 *   onProgress?: (info: { phase: 'chunk'|'hash'|'embed'|'upsert'|'delete', processed: number, total?: number }) => void,
 *   transform?: (record: { id: string, text: string, metadata: any }) => ({ text?: string, metadata?: any } | null | Promise<{ text?: string, metadata?: any } | null>)
 *   signal?: AbortSignal,
 * }} params
 */
export async function indexWorkbook(params) {
  const signal = params.signal;
  const onProgress = typeof params.onProgress === "function" ? params.onProgress : undefined;
  const embedBatchSize = (() => {
    const raw = params.embedBatchSize;
    // Default (and any invalid values) preserve current behavior: one embed call for all texts.
    if (raw === Infinity) return Infinity;
    if (typeof raw !== "number" || !Number.isFinite(raw) || raw <= 0) return Infinity;
    return Math.max(1, Math.floor(raw));
  })();
  throwIfAborted(signal);
  const { workbook, vectorStore, embedder } = params;
  const rawEmbedderName = embedder?.name;
  const embedderName =
    typeof rawEmbedderName === "string" && rawEmbedderName.trim() !== ""
      ? rawEmbedderName.trim()
      : "unknown-embedder";
  const sampleRows = params.sampleRows ?? 5;
  const maxColumnsForSchema = params.maxColumnsForSchema;
  const maxColumnsForRows = params.maxColumnsForRows;
  const tokenCount = params.tokenCount ?? approximateTokenCount;
  const chunks = chunkWorkbook(workbook, { signal });
  throwIfAborted(signal);
  onProgress?.({ phase: "chunk", processed: 0, total: chunks.length });

  const supportsUpdateMetadata = typeof vectorStore.updateMetadata === "function";

  /** @type {string[]} */
  let existingIds = [];
  /** @type {Map<string, { contentHash?: string, metadataHash?: string }>} */
  let existingState = new Map();

  if (typeof vectorStore.listContentHashes === "function") {
    const existing = await awaitWithAbort(
      vectorStore.listContentHashes({ workbookId: workbook.id, signal }),
      signal
    );
    throwIfAborted(signal);
    existingIds = existing.map((r) => r.id);
    existingState = new Map(
      existing.map((r) => [
        r.id,
        { contentHash: r.contentHash ?? undefined, metadataHash: r.metadataHash ?? undefined },
      ])
    );
  } else {
    const existingForWorkbook = await awaitWithAbort(
      vectorStore.list({
        workbookId: workbook.id,
        includeVector: false,
        signal,
      }),
      signal
    );
    throwIfAborted(signal);
    existingIds = existingForWorkbook.map((r) => r.id);
    existingState = new Map(
      existingForWorkbook.map((r) => [
        r.id,
        { contentHash: r.metadata?.contentHash, metadataHash: r.metadata?.metadataHash },
      ])
    );
  }

  const currentIds = new Set();
  /** @type {{ id: string, text: string, metadata: any }[]} */
  const toUpsert = [];
  /** @type {{ id: string, metadata: any }[]} */
  const toUpdateMetadata = [];

  let processedChunks = 0;
  for (const chunk of chunks) {
    throwIfAborted(signal);
    const originalText = chunkToText(chunk, { sampleRows, maxColumnsForSchema, maxColumnsForRows });

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
        embedder: embedderName,
      },
    };

    if (typeof params.transform === "function") {
      throwIfAborted(signal);
      const transformed = await awaitWithAbort(params.transform(record), signal);
      throwIfAborted(signal);
      if (!transformed) {
        processedChunks += 1;
        onProgress?.({ phase: "chunk", processed: processedChunks, total: chunks.length });
        continue;
      }

      if (Object.prototype.hasOwnProperty.call(transformed, "text")) {
        record.text = transformed.text ?? "";
      }
      if (Object.prototype.hasOwnProperty.call(transformed, "metadata")) {
        record.metadata = transformed.metadata ?? record.metadata;
      }

      const hasText = Object.prototype.hasOwnProperty.call(record.metadata ?? {}, "text");
      const hasEmbedder = Object.prototype.hasOwnProperty.call(record.metadata ?? {}, "embedder");
      if (!hasText || !hasEmbedder) {
        record.metadata = {
          ...(record.metadata ?? {}),
          ...(hasText ? null : { text: record.text }),
          ...(hasEmbedder ? null : { embedder: embedderName }),
        };
      }
    }

    throwIfAborted(signal);
    const chunkHash = await awaitWithAbort(contentHash(`${embedderName}\n${record.text}`), signal);
    throwIfAborted(signal);
    const metadataHash = await awaitWithAbort(computeMetadataHash(record.metadata), signal);
    throwIfAborted(signal);

    processedChunks += 1;
    onProgress?.({ phase: "hash", processed: processedChunks, total: chunks.length });
    onProgress?.({ phase: "chunk", processed: processedChunks, total: chunks.length });
    currentIds.add(record.id);

    const existing = existingState.get(record.id);
    if (existing?.contentHash === chunkHash) {
      if (existing?.metadataHash === metadataHash) continue;

      if (supportsUpdateMetadata) {
        toUpdateMetadata.push({
          id: record.id,
          metadata: {
            ...(record.metadata ?? {}),
            // Re-apply so every record stores the embedder identity unless a transform
            // explicitly overrides it.
            ...(Object.prototype.hasOwnProperty.call(record.metadata ?? {}, "embedder")
              ? null
              : { embedder: embedderName }),
            contentHash: chunkHash,
            metadataHash,
            tokenCount: tokenCount(record.text),
          },
        });
        continue;
      }
      // Backward compatible fallback: store doesn't support `updateMetadata`, so
      // re-upsert (and re-embed) the record.
    }

    toUpsert.push({
      id: record.id,
      text: record.text,
      metadata: {
        ...(record.metadata ?? {}),
        // Re-apply so every record stores the embedder identity unless a transform
        // explicitly overrides it.
        ...(Object.prototype.hasOwnProperty.call(record.metadata ?? {}, "embedder")
          ? null
          : { embedder: embedderName }),
        contentHash: chunkHash,
        metadataHash,
        tokenCount: tokenCount(record.text),
      },
    });
  }

  throwIfAborted(signal);
  /** @type {ArrayLike<number>[]} */
  let vectors = [];
  if (toUpsert.length > 0) {
    const texts = toUpsert.map((r) => r.text);
    onProgress?.({ phase: "embed", processed: 0, total: texts.length });
    // Allow callers to cancel from within onProgress before kicking off embedding work.
    throwIfAborted(signal);

    if (texts.length > embedBatchSize) {
      for (let i = 0; i < texts.length; i += embedBatchSize) {
        throwIfAborted(signal);
        const batchTexts = texts.slice(i, i + embedBatchSize);
        const batchVectors = await awaitWithAbort(
          embedder.embedTexts(batchTexts, { signal }),
          signal
        );
        // Preserve AbortSignal semantics: if callers cancel while/after the embedder resolves,
        // surface AbortError before validating embedder output.
        throwIfAborted(signal);
        if (!Array.isArray(batchVectors)) {
          throw new Error(
            `embedder.embedTexts returned a non-array result; expected an array of length ${batchTexts.length}`
          );
        }
        if (batchVectors.length !== batchTexts.length) {
          throw new Error(
            `embedder.embedTexts returned ${batchVectors.length} vector(s); expected ${batchTexts.length}`
          );
        }
        vectors.push(...batchVectors);
        onProgress?.({
          phase: "embed",
          processed: Math.min(i + batchTexts.length, texts.length),
          total: texts.length,
        });
      }
    } else {
      vectors = await awaitWithAbort(embedder.embedTexts(texts, { signal }), signal);
      onProgress?.({ phase: "embed", processed: texts.length, total: texts.length });
    }
  } else {
    onProgress?.({ phase: "embed", processed: 0, total: 0 });
  }
  throwIfAborted(signal);

  if (toUpsert.length > 0) {
    if (!Array.isArray(vectors)) {
      throw new Error(
        `embedder.embedTexts returned a non-array result; expected an array of length ${toUpsert.length}`
      );
    }
    if (vectors.length !== toUpsert.length) {
      throw new Error(
        `embedder.embedTexts returned ${vectors.length} vector(s); expected ${toUpsert.length}`
      );
    }

    const storeDim = vectorStore?.dimension;
    /** @type {number | undefined} */
    let expectedLen = Number.isFinite(storeDim) ? storeDim : undefined;
    for (let i = 0; i < vectors.length; i += 1) {
      // Preserve AbortSignal semantics: validation happens before persistence, so aborts should
      // be able to stop work promptly even when indexing large corpora.
      throwIfAborted(signal);
      const vec = vectors[i];
      const len = vec?.length;
      if (!Number.isFinite(len) || len <= 0) {
        throw new Error(
          `embedder.embedTexts returned an invalid vector for id=${toUpsert[i].id}: expected an array-like vector with a finite length`
        );
      }
      if (expectedLen === undefined) {
        expectedLen = len;
      } else if (len !== expectedLen) {
        throw new Error(
          `Vector dimension mismatch for id=${toUpsert[i].id}: expected ${expectedLen}, got ${len}`
        );
      }
      for (let j = 0; j < len; j += 1) {
        // Checking every ~256 values keeps abort responsiveness without adding too much overhead.
        if ((j & 0xff) === 0) throwIfAborted(signal);
        const value = vec[j];
        if (!Number.isFinite(value)) {
          throw new Error(
            `embedder.embedTexts returned an invalid vector value for id=${toUpsert[i].id} at index=${j}: expected a finite number`
          );
        }
      }
    }
  }

  // Delete stale records for this workbook.
  const staleIds = existingIds.filter((id) => !currentIds.has(id));

  const upsertRecords = toUpsert.map((r, i) => ({
    id: r.id,
    vector: vectors[i],
    metadata: r.metadata,
  }));
  const upsertTotal = upsertRecords.length + toUpdateMetadata.length;

  const hasMutations =
    upsertRecords.length > 0 || (supportsUpdateMetadata && toUpdateMetadata.length > 0) || staleIds.length > 0;
  if (hasMutations && typeof vectorStore.batch === "function") {
    // Avoid aborting while awaiting persistence. Batches are stateful; if we were to
    // reject early here, callers could start a new indexing run while the underlying
    // store is still writing.
    throwIfAborted(signal);
    await vectorStore.batch(async () => {
      if (signal?.aborted) return;
      if (upsertTotal > 0) {
        onProgress?.({ phase: "upsert", processed: 0, total: upsertTotal });
        // Allow callers to cancel from within onProgress before starting persistence.
        if (signal?.aborted) return;
      }
      let upsertProcessed = 0;
      if (upsertRecords.length) {
        await vectorStore.upsert(upsertRecords);
        upsertProcessed += upsertRecords.length;
        onProgress?.({ phase: "upsert", processed: upsertProcessed, total: upsertTotal });
      }
      // Preserve the existing behavior: if an abort happens during the upsert, skip
      // subsequent mutations, but do not throw inside the batch (throwing would skip
      // the final persistence snapshot).
      if (signal?.aborted) return;
      if (supportsUpdateMetadata && toUpdateMetadata.length) {
        await vectorStore.updateMetadata(toUpdateMetadata);
        upsertProcessed += toUpdateMetadata.length;
        onProgress?.({ phase: "upsert", processed: upsertProcessed, total: upsertTotal });
      }
      if (signal?.aborted) return;
      if (staleIds.length) {
        onProgress?.({ phase: "delete", processed: 0, total: staleIds.length });
        // Allow callers to cancel from within onProgress before starting persistence.
        if (signal?.aborted) return;
        await vectorStore.delete(staleIds);
        onProgress?.({ phase: "delete", processed: staleIds.length, total: staleIds.length });
      }
    });
  } else {
    if (upsertTotal > 0) {
      // Avoid aborting while awaiting persistence. Upserts are stateful; if we were to
      // reject early here, callers could start a new indexing run while the underlying
      // store is still writing.
      throwIfAborted(signal);
      onProgress?.({ phase: "upsert", processed: 0, total: upsertTotal });
      // Allow callers to cancel from within onProgress before starting persistence.
      throwIfAborted(signal);
    }
    let upsertProcessed = 0;
    if (upsertRecords.length) {
      await vectorStore.upsert(upsertRecords);
      upsertProcessed += upsertRecords.length;
      onProgress?.({ phase: "upsert", processed: upsertProcessed, total: upsertTotal });
    }
    throwIfAborted(signal);

    if (supportsUpdateMetadata && toUpdateMetadata.length) {
      // Avoid aborting while awaiting persistence. Metadata updates are stateful; if we
      // were to reject early here, callers could start a new indexing run while the
      // underlying store is still writing.
      throwIfAborted(signal);
      await vectorStore.updateMetadata(toUpdateMetadata);
      upsertProcessed += toUpdateMetadata.length;
      onProgress?.({ phase: "upsert", processed: upsertProcessed, total: upsertTotal });
    }
    throwIfAborted(signal);

    if (staleIds.length) {
      // Avoid aborting while awaiting persistence. Deletes are stateful; if we were to
      // reject early here, callers could start a new indexing run while the underlying
      // store is still writing.
      throwIfAborted(signal);
      onProgress?.({ phase: "delete", processed: 0, total: staleIds.length });
      // Allow callers to cancel from within onProgress before starting persistence.
      throwIfAborted(signal);
      await vectorStore.delete(staleIds);
      onProgress?.({ phase: "delete", processed: staleIds.length, total: staleIds.length });
    }
  }
  throwIfAborted(signal);

  return {
    totalChunks: chunks.length,
    upserted: toUpsert.length,
    skipped: chunks.length - toUpsert.length,
    deleted: staleIds.length,
  };
}
