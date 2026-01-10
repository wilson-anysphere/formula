import { extractSheetSchema } from "./schema.js";
import { RagIndex } from "./rag.js";
import { packSectionsToTokenBudget } from "./tokenBudget.js";
import { randomSampleRows, stratifiedSampleRows } from "./sampling.js";
import { classifyText, redactText } from "./dlp.js";

import { indexWorkbook } from "../../ai-rag/src/pipeline/indexWorkbook.js";

/**
 * @typedef {{ type: "range"|"formula"|"table"|"chart", reference: string, data?: any }} Attachment
 */

/**
 * @typedef {{
 *   vectorStore: any,
 *   embedder: { embedTexts(texts: string[]): Promise<ArrayLike<number>[]> },
 *   topK?: number,
 *   sampleRows?: number
 * }} WorkbookRagOptions
 */

export class ContextManager {
  /**
   * @param {{
   *   tokenBudgetTokens?: number,
   *   ragIndex?: RagIndex,
   *   workbookRag?: WorkbookRagOptions,
   *   redactor?: (text: string) => string
   * }} [options]
   */
  constructor(options = {}) {
    this.tokenBudgetTokens = options.tokenBudgetTokens ?? 16_000;
    this.ragIndex = options.ragIndex ?? new RagIndex();
    this.workbookRag = options.workbookRag;
    this.redactor = options.redactor ?? redactText;
  }

  /**
   * Build a compact context payload for chat prompts for a single sheet.
   *
   * @param {{
   *   sheet: { name: string, values: unknown[][], namedRanges?: any[] },
   *   query: string,
   *   attachments?: Attachment[],
   *   sampleRows?: number,
   *   samplingStrategy?: "random" | "stratified",
   *   stratifyByColumn?: number
   * }} params
   */
  async buildContext(params) {
    const schema = extractSheetSchema(params.sheet);

    // Index once per build for now; in the app this should be cached per sheet.
    await this.ragIndex.indexSheet(params.sheet);
    const retrieved = await this.ragIndex.search(params.query, 5);

    const sampleRows = params.sampleRows ?? 20;
    const dataForSampling = params.sheet.values.slice(0, 1_000); // safety cap
    const sampled =
      params.samplingStrategy === "stratified"
        ? stratifiedSampleRows(dataForSampling, sampleRows, {
            getStratum: (row) => String(row[params.stratifyByColumn ?? 0] ?? ""),
            seed: 1,
          })
        : randomSampleRows(dataForSampling, sampleRows, { seed: 1 });

    const sections = [
      {
        key: "schema",
        priority: 3,
        text: this.redactor(`Sheet schema (schema-first):\n${JSON.stringify(schema, null, 2)}`),
      },
      {
        key: "attachments",
        priority: 4,
        text: params.attachments?.length
          ? this.redactor(`User-provided attachments:\n${JSON.stringify(params.attachments, null, 2)}`)
          : "",
      },
      {
        key: "samples",
        priority: 2,
        text: sampled.length
          ? this.redactor(`Sample rows:\n${sampled.map((r) => JSON.stringify(r)).join("\n")}`)
          : "",
      },
      {
        key: "retrieved",
        priority: 1,
        text: retrieved.length ? this.redactor(`Retrieved context:\n${JSON.stringify(retrieved, null, 2)}`) : "",
      },
    ].filter((s) => s.text);

    const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens);
    return {
      schema,
      retrieved,
      sampledRows: sampled,
      promptContext: packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n"),
    };
  }

  /**
   * Build context for an entire workbook using the persistent vector store from
   * `packages/ai-rag`.
   *
   * Callers are expected to provide `workbookRag` (vectorStore + embedder). The
   * embeddings are cached incrementally via content hashes.
   *
   * @param {{
   *   workbook: any,
   *   query: string,
   *   attachments?: Attachment[],
   *   topK?: number
   * }} params
   */
  async buildWorkbookContext(params) {
    if (!this.workbookRag) throw new Error("ContextManager.buildWorkbookContext requires workbookRag");
    const { vectorStore, embedder } = this.workbookRag;
    const topK = params.topK ?? this.workbookRag.topK ?? 8;

    const indexStats = await indexWorkbook({
      workbook: params.workbook,
      vectorStore,
      embedder,
      sampleRows: this.workbookRag.sampleRows,
    });

    const [queryVector] = await embedder.embedTexts([params.query]);
    const hits = await vectorStore.query(queryVector, topK, {
      filter: (metadata) => metadata && metadata.workbookId === params.workbook.id,
    });

    const retrievedChunks = hits.map((hit, idx) => {
      const meta = hit.metadata ?? {};
      const title = meta.title ?? hit.id;
      const kind = meta.kind ?? "chunk";
      const header = `#${idx + 1} score=${hit.score.toFixed(3)} kind=${kind} sheet=${meta.sheetName ?? ""} title="${title}"`;
      const text = meta.text ?? "";
      const redacted = this.redactor(`${header}\n${text}`);
      return {
        id: hit.id,
        score: hit.score,
        metadata: meta,
        text: redacted,
        dlp: classifyText(redacted),
      };
    });

    const sections = [
      {
        key: "workbook_summary",
        priority: 3,
        text: this.redactor(
          `Workbook summary:\n${JSON.stringify(
            {
              id: params.workbook.id,
              sheets: (params.workbook.sheets ?? []).map((s) => s.name),
              tables: (params.workbook.tables ?? []).map((t) => ({
                name: t.name,
                sheetName: t.sheetName,
                rect: t.rect,
              })),
              namedRanges: (params.workbook.namedRanges ?? []).map((r) => ({
                name: r.name,
                sheetName: r.sheetName,
                rect: r.rect,
              })),
            },
            null,
            2
          )}`
        ),
      },
      {
        key: "attachments",
        priority: 4,
        text: params.attachments?.length
          ? this.redactor(`User-provided attachments:\n${JSON.stringify(params.attachments, null, 2)}`)
          : "",
      },
      {
        key: "retrieved",
        priority: 1,
        text: retrievedChunks.length ? `Retrieved workbook context:\n${retrievedChunks.map((c) => c.text).join("\n\n")}` : "",
      },
    ].filter((s) => s.text);

    const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens);
    return {
      indexStats,
      retrieved: retrievedChunks,
      promptContext: packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n"),
    };
  }
}
