import { sha256Hex } from "../utils/hash.js";
import { chunkWorkbook } from "../workbook/chunkWorkbook.js";
import { chunkToText } from "../workbook/chunkToText.js";

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
 * @param {{
 *   workbook: import('../workbook/workbookTypes').Workbook,
 *   vectorStore: any,
 *   embedder: { embedTexts(texts: string[]): Promise<ArrayLike<number>[]> },
 *   sampleRows?: number,
 * }} params
 */
export async function indexWorkbook(params) {
  const { workbook, vectorStore, embedder } = params;
  const sampleRows = params.sampleRows ?? 5;
  const chunks = chunkWorkbook(workbook);

  const currentIds = new Set();
  /** @type {{ id: string, text: string, metadata: any }[]} */
  const toUpsert = [];

  for (const chunk of chunks) {
    const text = chunkToText(chunk, { sampleRows });
    const contentHash = sha256Hex(text);
    currentIds.add(chunk.id);

    const existing = await vectorStore.get(chunk.id);
    if (existing?.metadata?.contentHash === contentHash) continue;

    toUpsert.push({
      id: chunk.id,
      text,
      metadata: {
        workbookId: chunk.workbookId,
        sheetName: chunk.sheetName,
        kind: chunk.kind,
        title: chunk.title,
        rect: chunk.rect,
        text,
        contentHash,
        tokenCount: approximateTokenCount(text),
      },
    });
  }

  const vectors =
    toUpsert.length > 0 ? await embedder.embedTexts(toUpsert.map((r) => r.text)) : [];

  await vectorStore.upsert(
    toUpsert.map((r, i) => ({
      id: r.id,
      vector: vectors[i],
      metadata: r.metadata,
    }))
  );

  // Delete stale records for this workbook.
  const existing = await vectorStore.list({
    filter: (metadata) => metadata && metadata.workbookId === workbook.id,
  });
  const staleIds = existing.map((r) => r.id).filter((id) => !currentIds.has(id));
  if (staleIds.length) await vectorStore.delete(staleIds);

  return {
    totalChunks: chunks.length,
    upserted: toUpsert.length,
    skipped: chunks.length - toUpsert.length,
    deleted: staleIds.length,
  };
}
