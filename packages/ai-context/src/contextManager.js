import { extractSheetSchema } from "./schema.js";
import { RagIndex } from "./rag.js";
import { DEFAULT_TOKEN_ESTIMATOR, packSectionsToTokenBudget, stableJsonStringify } from "./tokenBudget.js";
import { randomSampleRows, stratifiedSampleRows } from "./sampling.js";
import { classifyText, redactText } from "./dlp.js";

import { indexWorkbook } from "../../ai-rag/src/pipeline/indexWorkbook.js";
import { workbookFromSpreadsheetApi } from "../../ai-rag/src/workbook/fromSpreadsheetApi.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { evaluatePolicy, DLP_DECISION } from "../../security/dlp/src/policyEngine.js";
import {
  CLASSIFICATION_LEVEL,
  DEFAULT_CLASSIFICATION,
  classificationRank,
  maxClassification,
} from "../../security/dlp/src/classification.js";
import { effectiveRangeClassification, normalizeRange } from "../../security/dlp/src/selectors.js";
import { DlpViolationError } from "../../security/dlp/src/errors.js";

const DEFAULT_CLASSIFICATION_RANK = classificationRank(CLASSIFICATION_LEVEL.PUBLIC);
const RESTRICTED_CLASSIFICATION_RANK = classificationRank(CLASSIFICATION_LEVEL.RESTRICTED);

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
 * This cannot cancel underlying work (e.g. embedding requests), but it ensures callers can
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
 * @typedef {{ type: "range"|"formula"|"table"|"chart", reference: string, data?: any }} Attachment
 */

/**
 * @typedef {{
 *   vectorStore: any,
 *   embedder: { embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<ArrayLike<number>[]> },
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
   *   redactor?: (text: string) => string,
   *   tokenEstimator?: import("./tokenBudget.js").TokenEstimator
   * }} [options]
   */
  constructor(options = {}) {
    this.tokenBudgetTokens = options.tokenBudgetTokens ?? 16_000;
    this.ragIndex = options.ragIndex ?? new RagIndex();
    this.workbookRag = options.workbookRag;
    this.redactor = options.redactor ?? redactText;
    this.estimator = options.tokenEstimator ?? DEFAULT_TOKEN_ESTIMATOR;
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
   *   signal?: AbortSignal,
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
    const signal = params.signal;
    throwIfAborted(signal);
    const dlp = params.dlp;
    const rawSheet = params.sheet;

    const safeRowCap = 1_000;
    // `values` is a 2D JS array. With Excel-scale sheets, full-row/column selections can
    // explode into multi-million-cell matrices. Keep the context payload bounded so schema
    // extraction / RAG chunking can't OOM the worker.
    const safeCellCap = 200_000;
    const rawValues = Array.isArray(rawSheet?.values) ? rawSheet.values : [];
    const rowCount = Math.min(rawValues.length, safeRowCap);
    const safeColCap = rowCount > 0 ? Math.max(1, Math.floor(safeCellCap / rowCount)) : 0;
    const valuesForContext = rawValues.slice(0, rowCount).map((row) => {
      if (!Array.isArray(row) || safeColCap === 0) return [];
      return row.length <= safeColCap ? row.slice() : row.slice(0, safeColCap);
    });
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

      const normalizedRange = normalizeRange(rangeRef.range);
      dlpSelectionClassification = effectiveRangeClassification({ ...rangeRef, range: normalizedRange }, records);
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

      // Only do per-cell enforcement under REDACT decisions; in ALLOW cases the range max
      // classification is within the threshold so every in-range cell must be allowed.
      let nextValues;
      if (dlpDecision.decision === DLP_DECISION.REDACT) {
        const index = buildDlpRangeIndex({ documentId: dlp.documentId, sheetId, range: normalizedRange }, records);
        const maxAllowedRank = dlpDecision.maxAllowed === null ? null : classificationRank(dlpDecision.maxAllowed);
        const policyAllowsRestrictedContent = Boolean(dlp.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.allowRestrictedContent);
        const cellCheck = { index, maxAllowedRank, includeRestrictedContent, policyAllowsRestrictedContent };

        // Redact at cell level (deterministic placeholder).
        nextValues = [];
        for (let r = 0; r < valuesForContext.length; r++) {
          const row = valuesForContext[r] ?? [];
          const nextRow = [];
          for (let c = 0; c < row.length; c++) {
            if (isDlpCellAllowedFromIndex(cellCheck, r, c)) {
              nextRow.push(row[c]);
              continue;
            }
            dlpRedactedCells++;
            nextRow.push("[REDACTED]");
          }
          nextValues.push(nextRow);
        }
      } else {
        // Preserve the previous behavior of returning fresh row arrays (but skip DLP scans).
        nextValues = valuesForContext.map((row) => (row ?? []).slice());
      }

      sheetForContext = { ...rawSheet, name: sheetId, values: nextValues };
    }

    throwIfAborted(signal);
    const schema = extractSheetSchema(sheetForContext, { signal });

    // Index once per build for now; in the app this should be cached per sheet.
    throwIfAborted(signal);
    await this.ragIndex.indexSheet(sheetForContext, { signal });
    throwIfAborted(signal);
    const retrieved = await this.ragIndex.search(params.query, 5, { signal });
    throwIfAborted(signal);

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
        text: this.redactor(`Sheet schema (schema-first):\n${stableJsonStringify(schema)}`),
      },
      {
        key: "attachments",
        priority: 2,
        text: params.attachments?.length
          ? this.redactor(`User-provided attachments:\n${stableJsonStringify(params.attachments)}`)
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
        text: retrieved.length ? this.redactor(`Retrieved context:\n${stableJsonStringify(retrieved)}`) : "",
      },
    ].filter((s) => s.text);

    throwIfAborted(signal);
    const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens, this.estimator);
    throwIfAborted(signal);

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
   * Note: In Formula's desktop app, the embedder is the deterministic, offline
   * `HashEmbedder` by default (not user-configurable). A future Cursor-managed
   * embedding service can replace it to improve retrieval quality.
   *
   * @param {{
   *   workbook: any,
   *   query: string,
   *   attachments?: Attachment[],
   *   topK?: number,
   *   skipIndexing?: boolean,
   *   skipIndexingWithDlp?: boolean,
   *   includePromptContext?: boolean,
   *   signal?: AbortSignal,
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
    const signal = params.signal;
    throwIfAborted(signal);
    if (!this.workbookRag) throw new Error("ContextManager.buildWorkbookContext requires workbookRag");
    const { vectorStore, embedder } = this.workbookRag;
    const topK = params.topK ?? this.workbookRag.topK ?? 8;
    const includePromptContext = params.includePromptContext ?? true;
    const dlp = params.dlp;
    const includeRestrictedContent = dlp?.includeRestrictedContent ?? false;
    const classificationRecords =
      dlp?.classificationRecords ?? dlp?.classificationStore?.list(dlp.documentId) ?? [];

    // Some hosts (notably the desktop DocumentController) keep a stable internal sheet id
    // even after a user renames the sheet. In those cases:
    // - RAG chunk metadata uses the user-facing display name (better retrieval quality)
    // - Structured DLP classifications are recorded against the stable sheet id
    //
    // When a resolver is provided, map chunk `metadata.sheetName` back to the stable id
    // before applying structured DLP classification.
    const dlpSheetNameResolver =
      (dlp && (dlp.sheetNameResolver ?? dlp.sheet_name_resolver)) || null;
    const resolveDlpSheetId = (sheetNameOrId) => {
      const raw = typeof sheetNameOrId === "string" ? sheetNameOrId.trim() : "";
      if (!raw) return "";
      if (dlpSheetNameResolver && typeof dlpSheetNameResolver.getSheetIdByName === "function") {
        try {
          return dlpSheetNameResolver.getSheetIdByName(raw) ?? raw;
        } catch {
          return raw;
        }
      }
      return raw;
    };

    // Large enterprise classification record sets can make per-chunk range classification
    // expensive if we linearly scan `classificationRecords` for every chunk. Build a
    // document-level selector index once and reuse it for all range lookups in this call.
    /** @type {ReturnType<typeof buildDlpDocumentIndex> | null} */
    let dlpDocumentIndex = null;
    const getDlpDocumentIndex = () => {
      if (!dlp) return null;
      if (!classificationRecords.length) return null;
      if (!dlpDocumentIndex) {
        dlpDocumentIndex = buildDlpDocumentIndex({ documentId: dlp.documentId, records: classificationRecords });
      }
      return dlpDocumentIndex;
    };

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

    /** @type {Map<string, ReturnType<typeof classifyText>>} */
    const heuristicByChunkId = new Map();

    // In the desktop app, indexing can be expensive (it enumerates all non-empty cells).
    // Allow callers that manage their own incremental indexing to skip this step.
    //
    // Safety: if DLP enforcement is enabled for this build, indexing must still run so
    // chunk redaction can be applied before embedding and persistence.
    const skipIndexing = (params.skipIndexing ?? false) === true;
    const skipIndexingWithDlp = (params.skipIndexingWithDlp ?? false) === true;
    const shouldIndex = !skipIndexing || (Boolean(dlp) && !skipIndexingWithDlp);

    throwIfAborted(signal);
    const indexStats = shouldIndex
      ? await indexWorkbook({
          workbook: params.workbook,
          vectorStore,
          embedder,
          sampleRows: this.workbookRag.sampleRows,
          signal,
          transform: dlp
            ? (record) => {
                const rawText = record.text ?? "";
                const heuristic = classifyText(rawText);
                heuristicByChunkId.set(record.id, heuristic);
                const heuristicClassification = heuristicToPolicyClassification(heuristic);

                // Fold in structured DLP classifications for the chunk's sheet + rect metadata.
                let recordClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
                const range = rectToRange(record.metadata?.rect);
                 const sheetName = record.metadata?.sheetName;
                 const sheetId = sheetName ? resolveDlpSheetId(sheetName) : "";
                 if (range && sheetId) {
                   const index = getDlpDocumentIndex();
                   recordClassification = index
                     ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range })
                     : effectiveRangeClassification({ documentId: dlp.documentId, sheetId, range }, classificationRecords);
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
                  if (decision.decision === DLP_DECISION.BLOCK) {
                    // If the policy blocks cloud AI processing for this chunk, do not send any
                    // workbook content to the embedder. Persist only a minimal placeholder so
                    // the vector store cannot contain raw restricted data.
                    safeText = this.redactor(`${firstLine(rawText)}\n[REDACTED]`);
                  } else {
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
        })
      : null;

    // If DLP is enabled, redact the query before sending it to the embedder when policy
    // would not allow that sensitive content to be processed by cloud AI.
    //
    // Today, Formula's workbook RAG uses deterministic hash embeddings (offline), but we
    // still redact here so:
    // - retrieval stays consistent when indexed chunk text has been replaced with deterministic placeholders
    // - this remains safe if a future Cursor-managed embedding service is introduced
    const queryHeuristic = dlp ? classifyText(params.query) : null;
    const queryClassification = queryHeuristic ? heuristicToPolicyClassification(queryHeuristic) : null;
    const queryDecision =
      dlp && queryClassification
        ? evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification: queryClassification,
            policy: dlp.policy,
            options: { includeRestrictedContent },
          })
        : null;
    const queryForEmbedding =
      dlp && queryDecision && queryDecision.decision !== DLP_DECISION.ALLOW ? this.redactor(params.query) : params.query;
    throwIfAborted(signal);
    const [queryVector] = await awaitWithAbort(embedder.embedTexts([queryForEmbedding], { signal }), signal);
    throwIfAborted(signal);
    const hits = await awaitWithAbort(
      vectorStore.query(queryVector, topK, {
        workbookId: params.workbook.id,
        filter: (metadata) => metadata && metadata.workbookId === params.workbook.id,
        signal,
      }),
      signal
    );
    throwIfAborted(signal);

    /** @type {{level: string, labels: string[]} } */
    let overallClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
    // Structured DLP classifications are scoped to specific selectors (cell/range/etc).
    // When cloud AI processing is fully blocked (redactDisallowed=false), we need to
    // short-circuit even if no chunks are retrieved (e.g. a workbook with a single
    // classified cell that doesn't form a multi-cell RAG chunk). Treat the document's
    // overall classification as the max of all structured records so we can enforce
    // policies deterministically before sending anything to a cloud model.
    if (dlp && classificationRecords.length) {
      for (const record of classificationRecords) {
        throwIfAborted(signal);
        const classification = record?.classification;
        if (classification && typeof classification === "object") {
          overallClassification = maxClassification(overallClassification, classification);
        }
      }
    }
    /** @type {ReturnType<typeof evaluatePolicy> | null} */
    let overallDecision = null;
    let redactedChunkCount = 0;

    /** @type {any[]} */
    const chunkAudits = [];

    // Evaluate policy for all retrieved chunks before returning any prompt context.
    for (const [idx, hit] of hits.entries()) {
      throwIfAborted(signal);
      const meta = hit.metadata ?? {};
      const title = meta.title ?? hit.id;
      const kind = meta.kind ?? "chunk";
      const header = `#${idx + 1} score=${hit.score.toFixed(3)} kind=${kind} sheet=${meta.sheetName ?? ""} title="${title}"`;
      const text = meta.text ?? "";
      const raw = `${header}\n${text}`;

      const heuristic = heuristicByChunkId.get(hit.id) ?? meta.dlpHeuristic ?? classifyText(raw);
      const heuristicClassification = heuristicToPolicyClassification(heuristic);

      // If the caller provided structured cell/range classifications, fold those in using the
      // chunk's sheet + rect metadata.
      let recordClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
      if (dlp) {
        const range = rectToRange(meta.rect);
         const sheetName = meta.sheetName;
         const sheetId = sheetName ? resolveDlpSheetId(sheetName) : "";
         if (range && sheetId) {
           const index = getDlpDocumentIndex();
           recordClassification = index
             ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range })
             : effectiveRangeClassification({ documentId: dlp.documentId, sheetId, range }, classificationRecords);
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

    throwIfAborted(signal);
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
        dlp: heuristicByChunkId.get(hit.id) ?? meta.dlpHeuristic ?? classifyText(outText),
      };
    });

    let promptContext = "";
    if (includePromptContext) {
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
            `Workbook summary:\n${stableJsonStringify({
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
            })}`
          ),
        },
        {
          key: "attachments",
          priority: 2,
          text: params.attachments?.length
            ? this.redactor(`User-provided attachments:\n${stableJsonStringify(params.attachments)}`)
            : "",
        },
        {
          key: "retrieved",
          priority: 4,
          text: retrievedChunks.length
            ? `Retrieved workbook context:\n${retrievedChunks.map((c) => c.text).join("\n\n")}`
            : "",
        },
      ].filter((s) => s.text);

      throwIfAborted(signal);
      const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens, this.estimator);
      promptContext = packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n");
    }

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
      promptContext,
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
   *   skipIndexing?: boolean,
   *   skipIndexingWithDlp?: boolean,
   *   includePromptContext?: boolean,
   *   signal?: AbortSignal,
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
    const signal = params.signal;
    throwIfAborted(signal);
    const skipIndexing = (params.skipIndexing ?? false) === true;
    const skipIndexingWithDlp = (params.skipIndexingWithDlp ?? false) === true;

    // Some SpreadsheetApi hosts (desktop DocumentController adapter) can provide a
    // sheet-name resolver. Thread it through to DLP enforcement so structured
    // classifications keyed by stable sheet ids can still match RAG chunk metadata
    // (which uses user-facing sheet names).
    const spreadsheetResolver =
      params?.spreadsheet?.sheetNameResolver ?? params?.spreadsheet?.sheet_name_resolver ?? null;
    const dlp =
      params.dlp && spreadsheetResolver && !(params.dlp.sheetNameResolver || params.dlp.sheet_name_resolver)
        ? { ...params.dlp, sheetNameResolver: spreadsheetResolver, sheet_name_resolver: spreadsheetResolver }
        : params.dlp;

    const workbook =
      skipIndexing && (!dlp || skipIndexingWithDlp)
        ? {
            id: params.workbookId,
            sheets: (params.spreadsheet?.listSheets?.() ?? []).map((name) => ({ name, cells: new Map() })),
          }
        : workbookFromSpreadsheetApi({
            spreadsheet: params.spreadsheet,
            workbookId: params.workbookId,
            coordinateBase: "one",
            signal,
          });
    return this.buildWorkbookContext({
      workbook,
      query: params.query,
      attachments: params.attachments,
      topK: params.topK,
      dlp,
      skipIndexing: params.skipIndexing,
      skipIndexingWithDlp: params.skipIndexingWithDlp,
      includePromptContext: params.includePromptContext,
      signal,
    });
  }
}

function cellInNormalizedRange(cell, range) {
  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.col >= range.start.col &&
    cell.col <= range.end.col
  );
}

function rangesIntersectNormalized(a, b) {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

function buildDlpRangeIndex(ref, records) {
  const selectionRange = ref.range;
  const startRow = selectionRange.start.row;
  const startCol = selectionRange.start.col;
  const rowCount = selectionRange.end.row - selectionRange.start.row + 1;
  const colCount = selectionRange.end.col - selectionRange.start.col + 1;

  const rankFromClassification = (classification) => {
    if (!classification) return DEFAULT_CLASSIFICATION_RANK;
    if (typeof classification !== "object") {
      throw new Error("Classification must be an object");
    }
    return classificationRank(classification.level);
  };

  let docRankMax = DEFAULT_CLASSIFICATION_RANK;
  let sheetRankMax = DEFAULT_CLASSIFICATION_RANK;
  const columnRankByOffset = new Uint8Array(colCount);
  let cellRankByOffset = null;
  const rangeRecords = [];
  let rangeRankMax = DEFAULT_CLASSIFICATION_RANK;

  for (const record of records || []) {
    if (!record || !record.selector || typeof record.selector !== "object") continue;
    const selector = record.selector;
    if (selector.documentId !== ref.documentId) continue;

    // The per-cell allow/redact decision depends only on the max classification level.
    // Public records cannot increase the effective rank and are ignored for performance.
    const recordRank = rankFromClassification(record.classification);
    if (recordRank <= DEFAULT_CLASSIFICATION_RANK) continue;

    switch (selector.scope) {
      case "document": {
        docRankMax = Math.max(docRankMax, recordRank);
        break;
      }
      case "sheet": {
        if (selector.sheetId === ref.sheetId) {
          sheetRankMax = Math.max(sheetRankMax, recordRank);
        }
        break;
      }
      case "column": {
        if (selector.sheetId !== ref.sheetId) break;
        if (typeof selector.columnIndex === "number") {
          const colIndex = selector.columnIndex;
          if (colIndex < selectionRange.start.col || colIndex > selectionRange.end.col) break;
          const offset = colIndex - startCol;
          if (recordRank > columnRankByOffset[offset]) columnRankByOffset[offset] = recordRank;
        } else {
          // Table/columnId selectors require table metadata to evaluate; ContextManager's cell refs
          // do not include table context, so these selectors cannot apply and are ignored.
        }
        break;
      }
      case "cell": {
        if (selector.sheetId !== ref.sheetId) break;
        if (typeof selector.row !== "number" || typeof selector.col !== "number") break;
        if (
          selector.row < selectionRange.start.row ||
          selector.row > selectionRange.end.row ||
          selector.col < selectionRange.start.col ||
          selector.col > selectionRange.end.col
        ) {
          break;
        }
        const rowOffset = selector.row - startRow;
        const colOffset = selector.col - startCol;
        if (rowOffset < 0 || colOffset < 0 || rowOffset >= rowCount || colOffset >= colCount) break;
        if (cellRankByOffset === null) {
          cellRankByOffset = new Uint8Array(rowCount * colCount);
        }
        const offset = rowOffset * colCount + colOffset;
        if (recordRank > cellRankByOffset[offset]) cellRankByOffset[offset] = recordRank;
        break;
      }
      case "range": {
        if (selector.sheetId !== ref.sheetId) break;
        if (!selector.range) break;
        const normalized = normalizeRange(selector.range);
        if (!rangesIntersectNormalized(normalized, selectionRange)) break;
        if (recordRank > rangeRankMax) rangeRankMax = recordRank;
        rangeRecords.push({
          startRow: normalized.start.row,
          endRow: normalized.end.row,
          startCol: normalized.start.col,
          endCol: normalized.end.col,
          rank: recordRank,
        });
        break;
      }
      default: {
        // Unknown selector scope: ignore.
        break;
      }
    }
  }

  if (rangeRecords.length > 1) {
    rangeRecords.sort((a, b) => b.rank - a.rank);
  }

  return {
    docRankMax,
    sheetRankMax,
    startRow,
    startCol,
    rowCount,
    colCount,
    columnRankByOffset,
    cellRankByOffset,
    rangeRecords,
    rangeRankMax,
  };
}

function isDlpCellAllowedFromIndex(params, row0, col0) {
  const { index, maxAllowedRank, includeRestrictedContent, policyAllowsRestrictedContent } = params;
  if (maxAllowedRank === null) return false;

  // If we're explicitly including restricted content and policy allows it, a cell can become
  // ALLOW even if its classification exceeds `maxAllowed` (evaluatePolicy short-circuits for
  // Restricted + includeRestrictedContent).
  const restrictedOverrideAllowed = includeRestrictedContent && policyAllowsRestrictedContent;
  const canShortCircuitOverThreshold = !restrictedOverrideAllowed;
  const restrictedAllowed = includeRestrictedContent ? policyAllowsRestrictedContent : maxAllowedRank >= RESTRICTED_CLASSIFICATION_RANK;

  let rank = index.docRankMax;
  if (index.sheetRankMax > rank) rank = index.sheetRankMax;

  if (rank === RESTRICTED_CLASSIFICATION_RANK) return restrictedAllowed;
  if (canShortCircuitOverThreshold && rank > maxAllowedRank) return false;

  const colOffset = col0 - index.startCol;
  const colRank = index.columnRankByOffset[colOffset] ?? DEFAULT_CLASSIFICATION_RANK;
  if (colRank > rank) rank = colRank;

  if (rank === RESTRICTED_CLASSIFICATION_RANK) return restrictedAllowed;
  if (canShortCircuitOverThreshold && rank > maxAllowedRank) return false;

  if (index.cellRankByOffset !== null) {
    const rowOffset = row0 - index.startRow;
    if (rowOffset >= 0 && rowOffset < index.rowCount && colOffset >= 0 && colOffset < index.colCount) {
      const offset = rowOffset * index.colCount + colOffset;
      const cellRank = index.cellRankByOffset[offset] ?? DEFAULT_CLASSIFICATION_RANK;
      if (cellRank > rank) rank = cellRank;
    }
  }

  if (rank === RESTRICTED_CLASSIFICATION_RANK) return restrictedAllowed;
  if (canShortCircuitOverThreshold && rank > maxAllowedRank) return false;

  const rangeCanAffectDecision =
    index.rangeRankMax > maxAllowedRank ||
    (!restrictedAllowed && index.rangeRankMax === RESTRICTED_CLASSIFICATION_RANK);
  if (rangeCanAffectDecision && index.rangeRankMax > rank) {
    for (const record of index.rangeRecords) {
      if (record.rank <= rank) break;
      if (row0 < record.startRow || row0 > record.endRow || col0 < record.startCol || col0 > record.endCol) continue;
      rank = record.rank;
      if (rank === RESTRICTED_CLASSIFICATION_RANK) return restrictedAllowed;
      if (canShortCircuitOverThreshold && rank > maxAllowedRank) return false;
      if (rank === index.rangeRankMax) break;
    }
  }

  if (rank === RESTRICTED_CLASSIFICATION_RANK) return restrictedAllowed;
  return rank <= maxAllowedRank;
}

function buildDlpDocumentIndex(params) {
  let docClassificationMax = { ...DEFAULT_CLASSIFICATION };
  const sheetClassificationMaxBySheetId = new Map();
  const columnClassificationBySheetId = new Map();
  const cellClassificationBySheetId = new Map();
  const rangeRecordsBySheetId = new Map();

  /**
   * @param {string} sheetId
   */
  function ensureSheetMax(sheetId) {
    const existing = sheetClassificationMaxBySheetId.get(sheetId);
    if (existing) return existing;
    const init = { ...DEFAULT_CLASSIFICATION };
    sheetClassificationMaxBySheetId.set(sheetId, init);
    return init;
  }

  /**
   * @param {string} sheetId
   */
  function ensureColMap(sheetId) {
    const existing = columnClassificationBySheetId.get(sheetId);
    if (existing) return existing;
    const init = new Map();
    columnClassificationBySheetId.set(sheetId, init);
    return init;
  }

  /**
   * @param {string} sheetId
   */
  function ensureCellMap(sheetId) {
    const existing = cellClassificationBySheetId.get(sheetId);
    if (existing) return existing;
    const init = new Map();
    cellClassificationBySheetId.set(sheetId, init);
    return init;
  }

  /**
   * @param {string} sheetId
   */
  function ensureRangeList(sheetId) {
    const existing = rangeRecordsBySheetId.get(sheetId);
    if (existing) return existing;
    const init = [];
    rangeRecordsBySheetId.set(sheetId, init);
    return init;
  }

  for (const record of params.records || []) {
    if (!record || !record.selector || typeof record.selector !== "object") continue;
    const selector = record.selector;
    if (selector.documentId !== params.documentId) continue;

    switch (selector.scope) {
      case "document": {
        docClassificationMax = maxClassification(docClassificationMax, record.classification);
        break;
      }
      case "sheet": {
        if (!selector.sheetId) break;
        const existing = ensureSheetMax(selector.sheetId);
        sheetClassificationMaxBySheetId.set(selector.sheetId, maxClassification(existing, record.classification));
        break;
      }
      case "column": {
        if (!selector.sheetId) break;
        if (typeof selector.columnIndex !== "number") break;
        const colMap = ensureColMap(selector.sheetId);
        const existing = colMap.get(selector.columnIndex);
        colMap.set(selector.columnIndex, existing ? maxClassification(existing, record.classification) : record.classification);
        break;
      }
      case "cell": {
        if (!selector.sheetId) break;
        if (typeof selector.row !== "number" || typeof selector.col !== "number") break;
        const key = `${selector.row},${selector.col}`;
        const cellMap = ensureCellMap(selector.sheetId);
        const existing = cellMap.get(key);
        cellMap.set(key, existing ? maxClassification(existing, record.classification) : record.classification);
        break;
      }
      case "range": {
        if (!selector.sheetId) break;
        if (!selector.range) break;
        let normalized;
        try {
          normalized = normalizeRange(selector.range);
        } catch {
          break;
        }
        ensureRangeList(selector.sheetId).push({ range: normalized, classification: record.classification });
        break;
      }
      default:
        break;
    }
  }

  return {
    documentId: params.documentId,
    docClassificationMax,
    sheetClassificationMaxBySheetId,
    columnClassificationBySheetId,
    cellClassificationBySheetId,
    rangeRecordsBySheetId,
  };
}

function effectiveRangeClassificationFromDocumentIndex(index, rangeRef) {
  if (!index || rangeRef.documentId !== index.documentId) return { ...DEFAULT_CLASSIFICATION };

  let classification = { ...DEFAULT_CLASSIFICATION };
  classification = maxClassification(classification, index.docClassificationMax);
  if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;

  const normalized = normalizeRange(rangeRef.range);

  const sheetMax = index.sheetClassificationMaxBySheetId.get(rangeRef.sheetId);
  if (sheetMax) classification = maxClassification(classification, sheetMax);
  if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;

  const colMap = index.columnClassificationBySheetId.get(rangeRef.sheetId);
  if (colMap) {
    for (let col = normalized.start.col; col <= normalized.end.col; col++) {
      const colClassification = colMap.get(col);
      if (!colClassification) continue;
      classification = maxClassification(classification, colClassification);
      if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;
    }
  }

  const rangeRecords = index.rangeRecordsBySheetId.get(rangeRef.sheetId) ?? [];
  for (const record of rangeRecords) {
    if (!rangesIntersectNormalized(record.range, normalized)) continue;
    classification = maxClassification(classification, record.classification);
    if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;
  }

  const cellMap = index.cellClassificationBySheetId.get(rangeRef.sheetId);
  if (cellMap) {
    const rows = normalized.end.row - normalized.start.row + 1;
    const cols = normalized.end.col - normalized.start.col + 1;
    const rangeCells = rows * cols;

    // The chunk ranges used in workbook RAG are bounded (50x50 by default). Prefer
    // scanning the range coordinates rather than all classified cells on the sheet.
    // Use the smaller side if the caller ever passes a very large range.
    if (rangeCells <= cellMap.size) {
      for (let row = normalized.start.row; row <= normalized.end.row; row++) {
        for (let col = normalized.start.col; col <= normalized.end.col; col++) {
          const cellClassification = cellMap.get(`${row},${col}`);
          if (!cellClassification) continue;
          classification = maxClassification(classification, cellClassification);
          if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;
        }
      }
    } else {
      for (const [key, cellClassification] of cellMap.entries()) {
        const [rowRaw, colRaw] = String(key).split(",");
        const row = Number(rowRaw);
        const col = Number(colRaw);
        if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
        if (row < normalized.start.row || row > normalized.end.row) continue;
        if (col < normalized.start.col || col > normalized.end.col) continue;
        classification = maxClassification(classification, cellClassification);
        if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;
      }
    }
  }

  return classification;
}
