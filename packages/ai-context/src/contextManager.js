import { extractSheetSchema } from "./schema.js";
import { chunkSheetByRegions, RagIndex } from "./rag.js";
import { packSectionsToTokenBudget } from "./tokenBudget.js";
import { randomSampleRows, stratifiedSampleRows } from "./sampling.js";

/**
 * @typedef {{ type: "range"|"formula"|"table"|"chart", reference: string, data?: any }} Attachment
 */

export class ContextManager {
  /**
   * @param {{
   *   tokenBudgetTokens?: number,
   *   ragIndex?: RagIndex
   * }} [options]
   */
  constructor(options = {}) {
    this.tokenBudgetTokens = options.tokenBudgetTokens ?? 16_000;
    this.ragIndex = options.ragIndex ?? new RagIndex();
  }

  /**
   * Build a compact context payload for chat prompts.
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
        text: `Sheet schema (schema-first):\n${JSON.stringify(schema, null, 2)}`,
      },
      {
        key: "attachments",
        priority: 4,
        text: params.attachments?.length
          ? `User-provided attachments:\n${JSON.stringify(params.attachments, null, 2)}`
          : "",
      },
      {
        key: "samples",
        priority: 2,
        text: sampled.length ? `Sample rows:\n${sampled.map((r) => JSON.stringify(r)).join("\n")}` : "",
      },
      {
        key: "retrieved",
        priority: 1,
        text: retrieved.length ? `Retrieved context:\n${JSON.stringify(retrieved, null, 2)}` : "",
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
}
