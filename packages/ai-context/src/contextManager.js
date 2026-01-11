import { extractSheetSchema } from "./schema.js";
import { RagIndex } from "./rag.js";
import { packSectionsToTokenBudget } from "./tokenBudget.js";
import { randomSampleRows, stratifiedSampleRows } from "./sampling.js";
import { classifyText, redactText } from "./dlp.js";

import { indexWorkbook } from "../../ai-rag/src/pipeline/indexWorkbook.js";
import { workbookFromSpreadsheetApi } from "../../ai-rag/src/workbook/fromSpreadsheetApi.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { evaluatePolicy, DLP_DECISION } from "../../security/dlp/src/policyEngine.js";
import { CLASSIFICATION_LEVEL, maxClassification } from "../../security/dlp/src/classification.js";
import { effectiveCellClassification, effectiveRangeClassification } from "../../security/dlp/src/selectors.js";
import { DlpViolationError } from "../../security/dlp/src/errors.js";

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
   *   stratifyByColumn?: number,
   *   dlp?: {
   *     documentId: string,
   *     sheetId?: string,
   *     policy: any,
   *     classificationRecords?: Array<{ selector: any, classification: any }>,
   *     classificationStore?: { list(documentId: string): Array<{ selector: any, classification: any }> },
   *     includeRestrictedContent?: boolean,
   *     auditLogger?: { log(event: any): void }
   *   }
   * }} params
   */
  async buildContext(params) {
    const dlp = params.dlp;
    const rawSheet = params.sheet;

    const safeRowCap = 1_000;
    const valuesForContext = rawSheet.values.slice(0, safeRowCap);
    let sheetForContext = { ...rawSheet, values: valuesForContext };

    let dlpRedactedCells = 0;
    let dlpSelectionClassification = null;
    let dlpDecision = null;

    if (dlp) {
      const records = dlp.classificationRecords ?? dlp.classificationStore?.list(dlp.documentId) ?? [];
      const sheetId = dlp.sheetId ?? rawSheet.name;
      const includeRestrictedContent = dlp.includeRestrictedContent ?? false;

      const maxCols = valuesForContext.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);
      const rangeRef = {
        documentId: dlp.documentId,
        sheetId,
        range: {
          start: { row: 0, col: 0 },
          end: { row: Math.max(valuesForContext.length - 1, 0), col: Math.max(maxCols - 1, 0) },
        },
      };

      dlpSelectionClassification = effectiveRangeClassification(rangeRef, records);
      dlpDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: dlpSelectionClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });

      if (dlpDecision.decision === DLP_DECISION.BLOCK) {
        dlp.auditLogger?.log({
          type: "ai.context",
          documentId: dlp.documentId,
          sheetId,
          decision: dlpDecision,
          selectionClassification: dlpSelectionClassification,
          redactedCellCount: 0,
        });
        throw new DlpViolationError(dlpDecision);
      }

      // Redact at cell level (deterministic placeholder).
      const redactedValues = valuesForContext.map((row, r) =>
        (row ?? []).map((value, c) => {
          const classification = effectiveCellClassification(
            { documentId: dlp.documentId, sheetId, row: r, col: c },
            records,
          );
          const cellDecision = evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification,
            policy: dlp.policy,
            options: { includeRestrictedContent },
          });
          if (cellDecision.decision === DLP_DECISION.ALLOW) return value;
          dlpRedactedCells++;
          return "[REDACTED]";
        }),
      );

      sheetForContext = { ...rawSheet, name: sheetId, values: redactedValues };
    }

    const schema = extractSheetSchema(sheetForContext);

    // Index once per build for now; in the app this should be cached per sheet.
    await this.ragIndex.indexSheet(sheetForContext);
    const retrieved = await this.ragIndex.search(params.query, 5);

    const sampleRows = params.sampleRows ?? 20;
    const dataForSampling = sheetForContext.values; // already capped
    const sampled =
      params.samplingStrategy === "stratified"
        ? stratifiedSampleRows(dataForSampling, sampleRows, {
            getStratum: (row) => String(row[params.stratifyByColumn ?? 0] ?? ""),
            seed: 1,
          })
        : randomSampleRows(dataForSampling, sampleRows, { seed: 1 });

    const sections = [
      ...(dlpRedactedCells > 0
        ? [
            {
              key: "dlp",
              priority: 5,
              text: `DLP: ${dlpRedactedCells} cells were redacted due to policy.`,
            },
          ]
        : []),
      {
        key: "schema",
        priority: 3,
        text: this.redactor(`Sheet schema (schema-first):\n${JSON.stringify(schema, null, 2)}`),
      },
      {
        key: "attachments",
        priority: 2,
        text: params.attachments?.length
          ? this.redactor(`User-provided attachments:\n${JSON.stringify(params.attachments, null, 2)}`)
          : "",
      },
      {
        key: "samples",
        priority: 1,
        text: sampled.length
          ? this.redactor(`Sample rows:\n${sampled.map((r) => JSON.stringify(r)).join("\n")}`)
          : "",
      },
      {
        key: "retrieved",
        priority: 4,
        text: retrieved.length ? this.redactor(`Retrieved context:\n${JSON.stringify(retrieved, null, 2)}`) : "",
      },
    ].filter((s) => s.text);

    const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens);

    if (dlp) {
      dlp.auditLogger?.log({
        type: "ai.context",
        documentId: dlp.documentId,
        sheetId: dlp.sheetId ?? rawSheet.name,
        decision: dlpDecision,
        selectionClassification: dlpSelectionClassification,
        redactedCellCount: dlpRedactedCells,
      });
    }

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
   *   topK?: number,
   *   dlp?: {
   *     documentId: string,
   *     policy: any,
   *     classificationRecords?: Array<{ selector: any, classification: any }>,
   *     classificationStore?: { list(documentId: string): Array<{ selector: any, classification: any }> },
   *     includeRestrictedContent?: boolean,
   *     auditLogger?: { log(event: any): void }
   *   }
   * }} params
   */
  async buildWorkbookContext(params) {
    if (!this.workbookRag) throw new Error("ContextManager.buildWorkbookContext requires workbookRag");
    const { vectorStore, embedder } = this.workbookRag;
    const topK = params.topK ?? this.workbookRag.topK ?? 8;
    const dlp = params.dlp;
    const includeRestrictedContent = dlp?.includeRestrictedContent ?? false;
    const classificationRecords =
      dlp?.classificationRecords ?? dlp?.classificationStore?.list(dlp.documentId) ?? [];

    /**
     * @param {any} rect
     */
    function rectToRange(rect) {
      if (!rect || typeof rect !== "object") return null;
      const { r0, c0, r1, c1 } = rect;
      if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return null;
      return { start: { row: r0, col: c0 }, end: { row: r1, col: c1 } };
    }

    /**
     * @param {{level: string, findings: string[]}} heuristic
     */
    function heuristicToPolicyClassification(heuristic) {
      if (heuristic?.level === "sensitive") {
        // Conservatively map any heuristic "sensitive" findings to Restricted so policies can
        // block or redact the chunk before it is sent to a cloud model.
        const labels = (heuristic.findings || []).map((f) => `heuristic:${f}`);
        return { level: CLASSIFICATION_LEVEL.RESTRICTED, labels };
      }
      return { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
    }

    /**
     * @param {string} text
     */
    function firstLine(text) {
      const s = String(text ?? "");
      const idx = s.indexOf("\n");
      return idx === -1 ? s : s.slice(0, idx);
    }

    const indexStats = await indexWorkbook({
      workbook: params.workbook,
      vectorStore,
      embedder,
      sampleRows: this.workbookRag.sampleRows,
      transform: dlp
        ? (record) => {
            const rawText = record.text ?? "";
            const heuristic = classifyText(rawText);
            const heuristicClassification = heuristicToPolicyClassification(heuristic);

            // Fold in structured DLP classifications for the chunk's sheet + rect metadata.
            let recordClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
            const range = rectToRange(record.metadata?.rect);
            const sheetName = record.metadata?.sheetName;
            if (range && sheetName) {
              recordClassification = effectiveRangeClassification(
                { documentId: dlp.documentId, sheetId: sheetName, range },
                classificationRecords,
              );
            }

            const classification = maxClassification(recordClassification, heuristicClassification);

            const recordDecision = evaluatePolicy({
              action: DLP_ACTION.AI_CLOUD_PROCESSING,
              classification: recordClassification,
              policy: dlp.policy,
              options: { includeRestrictedContent },
            });

            const decision = evaluatePolicy({
              action: DLP_ACTION.AI_CLOUD_PROCESSING,
              classification,
              policy: dlp.policy,
              options: { includeRestrictedContent },
            });

            let safeText = rawText;
            if (decision.decision !== DLP_DECISION.ALLOW) {
              // If DLP redaction is required due to explicit document/sheet/range classification,
              // redact the entire content; pattern-based redaction isn't sufficient in that case.
              if (recordDecision.decision !== DLP_DECISION.ALLOW) {
                safeText = this.redactor(`${firstLine(rawText)}\n[REDACTED]`);
              } else {
                safeText = this.redactor(rawText);
              }
              if (!includeRestrictedContent && classifyText(safeText).level === "sensitive") {
                safeText = this.redactor(`${firstLine(rawText)}\n[REDACTED]`);
              }
            }

            return {
              text: safeText,
              metadata: {
                ...(record.metadata ?? {}),
                // Store the heuristic classification computed on the *raw* chunk text so policy
                // enforcement can still detect sensitive chunks even if `text` is redacted before
                // embedding / persistence.
                dlpHeuristic: heuristic,
                text: safeText,
              },
            };
          }
        : undefined,
    });

    const [queryVector] = await embedder.embedTexts([params.query]);
    const hits = await vectorStore.query(queryVector, topK, {
      workbookId: params.workbook.id,
      filter: (metadata) => metadata && metadata.workbookId === params.workbook.id,
    });

    /** @type {{level: string, labels: string[]} } */
    let overallClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
    /** @type {ReturnType<typeof evaluatePolicy> | null} */
    let overallDecision = null;
    let redactedChunkCount = 0;

    /** @type {any[]} */
    const chunkAudits = [];

    // Evaluate policy for all retrieved chunks before returning any prompt context.
    for (const [idx, hit] of hits.entries()) {
      const meta = hit.metadata ?? {};
      const title = meta.title ?? hit.id;
      const kind = meta.kind ?? "chunk";
      const header = `#${idx + 1} score=${hit.score.toFixed(3)} kind=${kind} sheet=${meta.sheetName ?? ""} title="${title}"`;
      const text = meta.text ?? "";
      const raw = `${header}\n${text}`;

      const heuristic = meta.dlpHeuristic ?? classifyText(raw);
      const heuristicClassification = heuristicToPolicyClassification(heuristic);

      // If the caller provided structured cell/range classifications, fold those in using the
      // chunk's sheet + rect metadata.
      let recordClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
      if (dlp) {
        const range = rectToRange(meta.rect);
        const sheetName = meta.sheetName;
        if (range && sheetName) {
          recordClassification = effectiveRangeClassification(
            { documentId: dlp.documentId, sheetId: sheetName, range },
            classificationRecords,
          );
        }
      }

      const classification = maxClassification(recordClassification, heuristicClassification);
      overallClassification = maxClassification(overallClassification, classification);

      const recordDecision = dlp
        ? evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification: recordClassification,
            policy: dlp.policy,
            options: { includeRestrictedContent },
          })
        : null;

      const decision = dlp
        ? evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification,
            policy: dlp.policy,
            options: { includeRestrictedContent },
          })
        : null;

      chunkAudits.push({
        id: hit.id,
        kind,
        sheetName: meta.sheetName,
        title,
        rect: meta.rect,
        recordClassification,
        recordDecision,
        classification,
        decision,
        heuristic,
      });

      if (decision?.decision === DLP_DECISION.BLOCK) {
        overallDecision = decision;
        dlp.auditLogger?.log({
          type: "ai.workbook_context",
          documentId: dlp.documentId,
          workbookId: params.workbook.id,
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          decision: overallDecision,
          classification: overallClassification,
          redactedChunkCount: 0,
          blockedChunkId: hit.id,
          chunks: chunkAudits,
        });
        throw new DlpViolationError(decision);
      }
    }

    if (dlp) {
      overallDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: overallClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });
      if (overallDecision.decision === DLP_DECISION.BLOCK) {
        dlp.auditLogger?.log({
          type: "ai.workbook_context",
          documentId: dlp.documentId,
          workbookId: params.workbook.id,
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          decision: overallDecision,
          classification: overallClassification,
          redactedChunkCount: 0,
          blockedChunkId: null,
          chunks: chunkAudits,
        });
        throw new DlpViolationError(overallDecision);
      }
    }

    const retrievedChunks = hits.map((hit, idx) => {
      const meta = hit.metadata ?? {};
      const title = meta.title ?? hit.id;
      const kind = meta.kind ?? "chunk";
      const header = `#${idx + 1} score=${hit.score.toFixed(3)} kind=${kind} sheet=${meta.sheetName ?? ""} title="${title}"`;
      const text = meta.text ?? "";
      const raw = `${header}\n${text}`;

      const audit = chunkAudits[idx];
      const decision = audit?.decision ?? null;
      const recordDecision = audit?.recordDecision ?? null;

      let outText = this.redactor(raw);
      let redacted = false;

      if (dlp && decision?.decision === DLP_DECISION.REDACT) {
        // If the chunk is disallowed due to explicit document/sheet/range classification,
        // redact the entire content. Pattern-based redaction is only safe when the
        // classification is derived solely from those patterns.
        if (recordDecision && recordDecision.decision !== DLP_DECISION.ALLOW) {
          outText = this.redactor(`${header}\n[REDACTED]`);
        }
        redacted = true;
      }

      // Defense-in-depth: if we're not explicitly including restricted content, never send
      // text that still matches the heuristic sensitive detectors.
      if (dlp && !includeRestrictedContent && classifyText(outText).level === "sensitive") {
        outText = this.redactor(`${header}\n[REDACTED]`);
        redacted = true;
      }

      if (redacted) redactedChunkCount += 1;

      // Do not return the raw chunk text stored in vector-store metadata. The prompt-safe
      // `text` field is already provided separately, and returning unredacted metadata
      // creates an easy footgun for callers that might serialize metadata into cloud LLM
      // prompts.
      const { text: _metaText, ...safeMeta } = meta;

      return {
        id: hit.id,
        score: hit.score,
        metadata: safeMeta,
        text: outText,
        dlp: classifyText(outText),
      };
    });

    const sections = [
      ...(dlp && redactedChunkCount > 0
        ? [
            {
              key: "dlp",
              priority: 5,
              text: `DLP: ${redactedChunkCount} retrieved chunks were redacted due to policy.`,
            },
          ]
        : []),
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
        priority: 2,
        text: params.attachments?.length
          ? this.redactor(`User-provided attachments:\n${JSON.stringify(params.attachments, null, 2)}`)
          : "",
      },
      {
        key: "retrieved",
        priority: 4,
        text: retrievedChunks.length ? `Retrieved workbook context:\n${retrievedChunks.map((c) => c.text).join("\n\n")}` : "",
      },
    ].filter((s) => s.text);

    const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens);

    if (dlp) {
      dlp.auditLogger?.log({
        type: "ai.workbook_context",
        documentId: dlp.documentId,
        workbookId: params.workbook.id,
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision: overallDecision,
        classification: overallClassification,
        redactedChunkCount,
        totalChunkCount: retrievedChunks.length,
        chunks: chunkAudits,
      });
    }
    return {
      indexStats,
      retrieved: retrievedChunks,
      promptContext: packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n"),
    };
  }

  /**
   * Convenience: build workbook RAG context from a `@formula/ai-tools`-style SpreadsheetApi.
   *
   * Note: `SpreadsheetApi` cell addresses are 1-based (A1 => row=1,col=1), while
   * `packages/ai-rag` uses 0-based coordinates internally.
   *
   * @param {{
   *   spreadsheet: any,
   *   workbookId: string,
   *   query: string,
   *   attachments?: Attachment[],
   *   topK?: number,
   *   dlp?: {
   *     documentId: string,
   *     policy: any,
   *     classificationRecords?: Array<{ selector: any, classification: any }>,
   *     classificationStore?: { list(documentId: string): Array<{ selector: any, classification: any }> },
   *     includeRestrictedContent?: boolean,
   *     auditLogger?: { log(event: any): void }
   *   }
   * }} params
   */
  async buildWorkbookContextFromSpreadsheetApi(params) {
    const workbook = workbookFromSpreadsheetApi({
      spreadsheet: params.spreadsheet,
      workbookId: params.workbookId,
      coordinateBase: "one",
    });
    return this.buildWorkbookContext({
      workbook,
      query: params.query,
      attachments: params.attachments,
      topK: params.topK,
      dlp: params.dlp,
    });
  }
}
