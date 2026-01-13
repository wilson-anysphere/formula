import { RagIndex } from "./rag.js";
import { DEFAULT_TOKEN_ESTIMATOR, packSectionsToTokenBudget, stableJsonStringify } from "./tokenBudget.js";
import { headSampleRows, randomSampleRows, stratifiedSampleRows, systematicSampleRows, tailSampleRows } from "./sampling.js";
import { classifyText, redactText } from "./dlp.js";
import { isCellEmpty, parseA1Range, rangeToA1 } from "./a1.js";
import { awaitWithAbort, throwIfAborted } from "./abort.js";
import { inferColumnType, isLikelyHeaderRow } from "./schema.js";

import { indexWorkbook } from "../../ai-rag/src/pipeline/indexWorkbook.js";
import { searchWorkbookRag } from "../../ai-rag/src/retrieval/searchWorkbookRag.js";
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

const DEFAULT_SHEET_INDEX_CACHE_LIMIT = 32;
const DEFAULT_RAG_MAX_CHUNK_ROWS = 30;
const SHEET_INDEX_SIGNATURE_VERSION = 1;

const FNV_OFFSET_64 = 0xcbf29ce484222325n;
const FNV_PRIME_64 = 0x100000001b3n;
const FNV_MASK_64 = 0xffffffffffffffffn;

/**
 * @param {bigint} hash
 * @param {string} input
 */
function fnv1a64Update(hash, input) {
  let out = hash;
  for (let i = 0; i < input.length; i++) {
    out ^= BigInt(input.charCodeAt(i));
    out = (out * FNV_PRIME_64) & FNV_MASK_64;
  }
  return out;
}

/**
 * @param {unknown} origin
 */
function normalizeSheetOrigin(origin) {
  const row = origin && typeof origin === "object" && Number.isInteger(origin.row) && origin.row >= 0 ? origin.row : 0;
  const col = origin && typeof origin === "object" && Number.isInteger(origin.col) && origin.col >= 0 ? origin.col : 0;
  return { row, col };
}

/**
 * Cache key for a single-sheet RAG index, keyed by (sheet.name, sheet.origin?).
 * @param {{ name?: unknown, origin?: any }} sheet
 */
function sheetIndexCacheKey(sheet) {
  const name = sheet && typeof sheet === "object" && typeof sheet.name === "string" ? sheet.name : "";
  const origin = normalizeSheetOrigin(sheet?.origin);
  return `${name}::${origin.row},${origin.col}`;
}

/**
 * @param {unknown} value
 * @param {{ signal?: AbortSignal }} [options]
 */
function stableHashValue(value, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);

  /**
   * @param {bigint} hash
   * @param {unknown} v
   */
  function walk(hash, v) {
    throwIfAborted(signal);
    if (v === undefined || v === null) return fnv1a64Update(hash, "null");
    if (typeof v === "boolean") return fnv1a64Update(hash, v ? "true" : "false");

    if (typeof v === "number") {
      if (Number.isNaN(v)) return fnv1a64Update(hash, "num:NaN");
      if (v === Infinity) return fnv1a64Update(hash, "num:Infinity");
      if (v === -Infinity) return fnv1a64Update(hash, "num:-Infinity");
      // Preserve the sign of -0.
      if (Object.is(v, -0)) return fnv1a64Update(hash, "num:-0");
      return fnv1a64Update(hash, `num:${String(v)}`);
    }

    if (typeof v === "string") {
      hash = fnv1a64Update(hash, "str:");
      hash = fnv1a64Update(hash, String(v.length));
      hash = fnv1a64Update(hash, ":");
      return fnv1a64Update(hash, v);
    }

    if (typeof v === "bigint") return fnv1a64Update(hash, `bigint:${v.toString()}`);
    if (typeof v === "symbol") return fnv1a64Update(hash, `symbol:${v.toString()}`);
    if (typeof v === "function") return fnv1a64Update(hash, `fn:${v.name || "anonymous"}`);
    if (v instanceof Date) return fnv1a64Update(hash, `date:${v.toISOString()}`);

    if (v instanceof Map) {
      hash = fnv1a64Update(hash, "map{");
      for (const [k, val] of v.entries()) {
        hash = fnv1a64Update(hash, "k=");
        hash = walk(hash, k);
        hash = fnv1a64Update(hash, "v=");
        hash = walk(hash, val);
        hash = fnv1a64Update(hash, ";");
      }
      return fnv1a64Update(hash, "}");
    }

    if (v instanceof Set) {
      hash = fnv1a64Update(hash, "set[");
      for (const val of v.values()) {
        hash = walk(hash, val);
        hash = fnv1a64Update(hash, ";");
      }
      return fnv1a64Update(hash, "]");
    }

    if (Array.isArray(v)) {
      hash = fnv1a64Update(hash, "[");
      hash = fnv1a64Update(hash, String(v.length));
      hash = fnv1a64Update(hash, ":");
      for (const item of v) {
        hash = walk(hash, item);
        hash = fnv1a64Update(hash, ",");
      }
      return fnv1a64Update(hash, "]");
    }

    if (typeof v === "object") {
      const obj = /** @type {Record<string, unknown>} */ (v);
      const keys = Object.keys(obj).sort();
      hash = fnv1a64Update(hash, "{");
      for (const key of keys) {
        hash = fnv1a64Update(hash, "k:");
        hash = fnv1a64Update(hash, key);
        hash = fnv1a64Update(hash, "v:");
        hash = walk(hash, obj[key]);
        hash = fnv1a64Update(hash, ";");
      }
      return fnv1a64Update(hash, "}");
    }

    return fnv1a64Update(hash, `other:${String(v)}`);
  }

  const hashed = walk(FNV_OFFSET_64, value);
  return hashed.toString(16).padStart(16, "0");
}

/**
 * Deterministic signature of inputs that affect RAG chunking/indexing.
 *
 * @param {{ name?: string, origin?: any, values?: unknown[][] }} sheet
 * @param {{ maxChunkRows?: number, signal?: AbortSignal }} [options]
 */
function computeSheetIndexSignature(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const origin = normalizeSheetOrigin(sheet?.origin);
  const maxChunkRows = options.maxChunkRows ?? DEFAULT_RAG_MAX_CHUNK_ROWS;

  let hash = FNV_OFFSET_64;
  hash = fnv1a64Update(hash, `sig:v${SHEET_INDEX_SIGNATURE_VERSION}\n`);
  hash = fnv1a64Update(hash, `name:${sheet?.name ?? ""}\n`);
  hash = fnv1a64Update(hash, `origin:${origin.row},${origin.col}\n`);
  hash = fnv1a64Update(hash, `maxChunkRows:${String(maxChunkRows)}\n`);
  hash = fnv1a64Update(hash, "values:");
  hash = fnv1a64Update(hash, stableHashValue(sheet?.values ?? [], { signal }));
  return hash.toString(16).padStart(16, "0");
}

/**
 * @param {unknown} value
 * @param {number} fallback
 */
function normalizeNonNegativeInt(value, fallback) {
  if (value === undefined || value === null) return fallback;
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(0, Math.floor(n));
}

/**
 * Normalize a workbook cell (ai-rag / spreadsheet-like) to a scalar value for schema inference.
 *
 * Supported shapes:
 * - { v, f } (ai-rag normalized cell)
 * - { value, formula } (SpreadsheetApi cell)
 * - raw scalar (string/number/etc)
 *
 * @param {unknown} cell
 * @returns {unknown}
 */
function workbookCellToScalar(cell) {
  if (cell == null) return cell;
  if (typeof cell !== "object") return cell;
  const obj = /** @type {any} */ (cell);
  const f = obj.f ?? obj.formula;
  if (typeof f === "string" && f.trim() !== "") {
    // Ensure formula strings get classified as formulas by downstream type inference.
    const trimmed = f.trim();
    return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
  }
  if (Object.prototype.hasOwnProperty.call(obj, "v")) return obj.v;
  if (Object.prototype.hasOwnProperty.call(obj, "value")) return obj.value;
  return cell;
}

/**
 * @param {unknown} value
 */
function isEmptyScalar(value) {
  return value === null || value === undefined || value === "";
}

/**
 * @param {any} sheet
 * @param {number} row
 * @param {number} col
 * @returns {unknown}
 */
function getWorkbookSheetCell(sheet, row, col) {
  if (!sheet || typeof sheet !== "object") return undefined;
  const values = sheet.values;
  if (Array.isArray(values)) return workbookCellToScalar(values[row]?.[col]);

  const cells = sheet.cells;
  if (Array.isArray(cells)) return workbookCellToScalar(cells[row]?.[col]);

  if (cells && typeof cells.get === "function") {
    const keyComma = `${row},${col}`;
    const keyColon = `${row}:${col}`;
    return workbookCellToScalar(cells.get(keyComma) ?? cells.get(keyColon));
  }

  return undefined;
}

/**
 * Normalize DLP options so ContextManager methods can accept both camelCase and
 * snake_case field names (e.g. when options are deserialized from JSON in a
 * non-TS host).
 *
 * @param {any} dlp
 * @returns {null | {
 *   documentId: string,
 *   sheetId?: string,
 *   policy: any,
 *   classificationRecords?: Array<{ selector: any, classification: any }>,
 *   classificationStore?: { list(documentId: string): Array<{ selector: any, classification: any }> },
 *   includeRestrictedContent: boolean,
 *   auditLogger?: { log(event: any): void },
 *   sheetNameResolver?: any
 * }}
 */
function normalizeDlpOptions(dlp) {
  if (!dlp) return null;
  if (typeof dlp !== "object") {
    throw new Error("DLP options must be an object");
  }
  return {
    documentId: dlp.documentId ?? dlp.document_id,
    sheetId: dlp.sheetId ?? dlp.sheet_id,
    policy: dlp.policy,
    classificationRecords: dlp.classificationRecords ?? dlp.classification_records,
    classificationStore: dlp.classificationStore ?? dlp.classification_store,
    includeRestrictedContent: (dlp.includeRestrictedContent ?? dlp.include_restricted_content ?? false) === true,
    auditLogger: dlp.auditLogger,
    sheetNameResolver: dlp.sheetNameResolver ?? dlp.sheet_name_resolver,
  };
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
   *   cacheSheetIndex?: boolean,
   *   workbookRag?: WorkbookRagOptions,
   *   maxContextRows?: number,
   *   maxContextCells?: number,
   *   maxChunkRows?: number,
   *   sheetRagTopK?: number,
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
    // These caps are primarily safety rails to prevent accidental Excel-scale context
    // selections from blowing up into multi-million-cell matrices in memory.
    //
    // Hosts can tune these for their own memory/quality tradeoffs.
    this.maxContextRows = normalizeNonNegativeInt(options.maxContextRows, 1_000);
    this.maxContextCells = normalizeNonNegativeInt(options.maxContextCells, 200_000);
    // `maxChunkRows` controls how many TSV rows are included in each RAG chunk's text.
    this.maxChunkRows = normalizeNonNegativeInt(options.maxChunkRows, 30);
    // Top-K retrieved regions for sheet-level (non-workbook) RAG.
    this.sheetRagTopK = normalizeNonNegativeInt(options.sheetRagTopK, 5);

    this.cacheSheetIndex = options.cacheSheetIndex ?? true;
    /** @type {Map<string, { signature: string, schema: any, sheetName: string }>} */
    this._sheetIndexCache = new Map();
    this._sheetIndexCacheLimit = DEFAULT_SHEET_INDEX_CACHE_LIMIT;
    /** @type {Map<string, string>} */
    this._sheetNameToActiveCacheKey = new Map();
  }

  /**
   * Index a sheet into the in-memory RAG store, with incremental caching by sheet signature.
   *
   * Returns the extracted schema (reused from chunking/indexing when possible).
   *
   * @param {{ name: string, values: unknown[][], origin?: any }} sheet
   * @param {{ signal?: AbortSignal, maxChunkRows?: number }} [options]
   * @returns {Promise<{ schema: any }>}
   */
  async _ensureSheetIndexed(sheet, options = {}) {
    const signal = options.signal;
    const maxChunkRows = options.maxChunkRows;
    throwIfAborted(signal);

    if (!this.cacheSheetIndex) {
      const { schema } = await this.ragIndex.indexSheet(sheet, { signal, maxChunkRows });
      return { schema };
    }

    const cacheKey = sheetIndexCacheKey(sheet);
    const signature = computeSheetIndexSignature(sheet, { signal, maxChunkRows });

    const cached = this._sheetIndexCache.get(cacheKey);
    if (cached) {
      // Refresh LRU on access.
      this._sheetIndexCache.delete(cacheKey);
      this._sheetIndexCache.set(cacheKey, cached);
    }

    const activeKey = this._sheetNameToActiveCacheKey.get(sheet.name);
    const upToDate = cached?.signature === signature && activeKey === cacheKey;
    if (upToDate) return { schema: cached.schema };

    const { schema } = await this.ragIndex.indexSheet(sheet, { signal, maxChunkRows });

    // Update caches after successful indexing.
    this._sheetNameToActiveCacheKey.set(sheet.name, cacheKey);
    this._sheetIndexCache.delete(cacheKey);
    this._sheetIndexCache.set(cacheKey, { signature, schema, sheetName: sheet.name });
    while (this._sheetIndexCache.size > this._sheetIndexCacheLimit) {
      const oldestKey = this._sheetIndexCache.keys().next().value;
      if (oldestKey === undefined) break;
      const oldestEntry = this._sheetIndexCache.get(oldestKey);
      this._sheetIndexCache.delete(oldestKey);
      if (oldestEntry?.sheetName) {
        const activeKeyForSheet = this._sheetNameToActiveCacheKey.get(oldestEntry.sheetName);
        if (activeKeyForSheet === oldestKey) {
          this._sheetNameToActiveCacheKey.delete(oldestEntry.sheetName);
        }
      }
    }

    return { schema };
  }

  /**
   * Build a compact context payload for chat prompts for a single sheet.
   *
   * @param {{
   *   sheet: { name: string, values: unknown[][], namedRanges?: any[] },
   *   query: string,
   *   attachments?: Attachment[],
   *   sampleRows?: number,
   *   samplingStrategy?: "random" | "stratified" | "head" | "tail" | "systematic",
   *   stratifyByColumn?: number,
   *   limits?: { maxContextRows?: number, maxContextCells?: number, maxChunkRows?: number },
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
    const dlp = normalizeDlpOptions(params.dlp);
    const rawSheet = params.sheet;

    const safeRowCap = normalizeNonNegativeInt(params.limits?.maxContextRows, this.maxContextRows);
    // `values` is a 2D JS array. With Excel-scale sheets, full-row/column selections can
    // explode into multi-million-cell matrices. Keep the context payload bounded so schema
    // extraction / RAG chunking can't OOM the worker.
    const safeCellCap = normalizeNonNegativeInt(params.limits?.maxContextCells, this.maxContextCells);
    const maxChunkRows = normalizeNonNegativeInt(params.limits?.maxChunkRows, this.maxChunkRows);
    const rawValues = Array.isArray(rawSheet?.values) ? rawSheet.values : [];
    const rowCount = Math.min(rawValues.length, safeRowCap);
    const safeColCap = rowCount > 0 ? Math.max(1, Math.floor(safeCellCap / rowCount)) : 0;
    const valuesForContext = rawValues.slice(0, rowCount).map((row) => {
      if (!Array.isArray(row) || safeColCap === 0) return [];
      return row.length <= safeColCap ? row.slice() : row.slice(0, safeColCap);
    });
    const origin =
      rawSheet && typeof rawSheet === "object" && rawSheet.origin && typeof rawSheet.origin === "object"
        ? {
            row: Number.isInteger(rawSheet.origin.row) && rawSheet.origin.row >= 0 ? rawSheet.origin.row : 0,
            col: Number.isInteger(rawSheet.origin.col) && rawSheet.origin.col >= 0 ? rawSheet.origin.col : 0,
          }
        : { row: 0, col: 0 };
    let sheetForContext = { ...rawSheet, values: valuesForContext };

    let dlpRedactedCells = 0;
    let dlpSelectionClassification = null;
    let dlpDecision = null;
    let dlpHeuristic = null;
    let dlpHeuristicApplied = false;
    let dlpAuditDocumentId = null;
    let dlpAuditSheetId = null;

    if (dlp) {
      const documentId = dlp.documentId;
      const records = dlp.classificationRecords ?? dlp.classificationStore?.list?.(documentId) ?? [];
      const includeRestrictedContent = dlp.includeRestrictedContent ?? false;

      // Some hosts keep stable internal sheet ids even after a user renames the sheet
      // (id != display name). When a resolver is provided, map the user-facing name back
      // to the stable id before evaluating structured DLP selectors.
      const dlpSheetNameResolver = dlp.sheetNameResolver ?? null;
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

      const sheetId = resolveDlpSheetId(dlp.sheetId ?? rawSheet.name);
      dlpAuditDocumentId = documentId;
      dlpAuditSheetId = sheetId;

      const maxCols = valuesForContext.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);
      const rangeRef = {
        documentId,
        sheetId,
        range: {
          start: { row: origin.row, col: origin.col },
          end: {
            row: origin.row + Math.max(valuesForContext.length - 1, 0),
            col: origin.col + Math.max(maxCols - 1, 0),
          },
        },
      };

      const normalizedRange = normalizeRange(rangeRef.range);
      const structuredSelectionClassification = effectiveRangeClassification({ ...rangeRef, range: normalizedRange }, records);
      dlpSelectionClassification = structuredSelectionClassification;
      let structuredDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: structuredSelectionClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });
      dlpDecision = structuredDecision;

      if (structuredDecision.decision === DLP_DECISION.BLOCK) {
        dlp.auditLogger?.log({
          type: "ai.context",
          documentId,
          sheetId,
          sheetName: rawSheet.name,
          decision: structuredDecision,
          selectionClassification: structuredSelectionClassification,
          redactedCellCount: 0,
        });
        throw new DlpViolationError(structuredDecision);
      }

      // Only do per-cell enforcement under REDACT decisions; in ALLOW cases the range max
      // classification is within the threshold so every in-range cell must be allowed.
      let nextValues;
      if (structuredDecision.decision === DLP_DECISION.REDACT) {
        const maxAllowedRank =
          structuredDecision.maxAllowed === null ? null : classificationRank(structuredDecision.maxAllowed);
        const index = buildDlpRangeIndex({ documentId, sheetId, range: normalizedRange }, records, {
          maxAllowedRank: maxAllowedRank ?? DEFAULT_CLASSIFICATION_RANK,
          signal,
        });
        const policyAllowsRestrictedContent = Boolean(dlp.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.allowRestrictedContent);
        const cellCheck = { index, maxAllowedRank, includeRestrictedContent, policyAllowsRestrictedContent, signal };

        // Redact at cell level (deterministic placeholder).
        nextValues = [];
        const originRow = origin.row;
        const originCol = origin.col;
        for (let r = 0; r < valuesForContext.length; r++) {
          throwIfAborted(signal);
          const row = valuesForContext[r] ?? [];
          const nextRow = [];
          for (let c = 0; c < row.length; c++) {
            throwIfAborted(signal);
            if (isDlpCellAllowedFromIndex(cellCheck, originRow + r, originCol + c)) {
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

      sheetForContext = { ...rawSheet, values: nextValues };

      // Workbook DLP enforcement treats heuristic sensitive patterns as Restricted when evaluating
      // AI cloud processing policies. Mirror that behavior in the single-sheet context path so
      // callers can't accidentally leak sensitive content when no structured selectors are present.
      if (structuredDecision.decision === DLP_DECISION.ALLOW) {
        dlpHeuristic = classifyValuesForDlp(nextValues, { signal });
        const heuristicPolicyClassification = heuristicToPolicyClassification(dlpHeuristic);
        const combinedClassification = maxClassification(structuredSelectionClassification, heuristicPolicyClassification);
        // Preserve the structured classification for audit/debug, but evaluate policy on the max.
        if (heuristicPolicyClassification.level !== CLASSIFICATION_LEVEL.PUBLIC) {
          dlpHeuristicApplied = true;
          dlpSelectionClassification = combinedClassification;
          dlpDecision = evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification: combinedClassification,
            policy: dlp.policy,
            options: { includeRestrictedContent },
          });
        }
      }

      if (dlpDecision.decision === DLP_DECISION.BLOCK) {
        dlp.auditLogger?.log({
          type: "ai.context",
          documentId,
          sheetId,
          sheetName: rawSheet.name,
          decision: dlpDecision,
          selectionClassification: dlpSelectionClassification,
          redactedCellCount: 0,
        });
        throw new DlpViolationError(dlpDecision);
      }

      // Under REDACT decisions, defensively apply heuristic redaction to the context sheet so:
      //  - schema / sampling / retrieval don't contain raw sensitive strings in structured outputs
      //  - in-memory RAG text doesn't retain sensitive patterns (defense-in-depth)
      if (dlpDecision.decision === DLP_DECISION.REDACT) {
        sheetForContext = {
          ...sheetForContext,
          values: redactValuesForDlp(sheetForContext.values, this.redactor, { signal, includeRestrictedContent }),
        };
      }
    }

    throwIfAborted(signal);
    // Index sheet into the in-memory RAG store (cached by content signature).
    //
    // `RagIndex.indexSheet()` extracts the schema as part of chunking; `_ensureSheetIndexed()`
    // reuses that work (or cached results) so we don't run schema extraction twice.
    const { schema } = await this._ensureSheetIndexed(sheetForContext, { signal, maxChunkRows });
    throwIfAborted(signal);
    const retrieved = await this.ragIndex.search(params.query, this.sheetRagTopK, { signal });
    throwIfAborted(signal);

    const sampleRows = params.sampleRows ?? 20;
    const dataForSampling = sheetForContext.values; // already capped
    let sampled;
    switch (params.samplingStrategy) {
      case "stratified": {
        sampled = stratifiedSampleRows(dataForSampling, sampleRows, {
          getStratum: (row) => String(row[params.stratifyByColumn ?? 0] ?? ""),
          seed: 1,
        });
        break;
      }
      case "head": {
        sampled = headSampleRows(dataForSampling, sampleRows);
        break;
      }
      case "tail": {
        sampled = tailSampleRows(dataForSampling, sampleRows);
        break;
      }
      case "systematic": {
        sampled = systematicSampleRows(dataForSampling, sampleRows, { seed: 1 });
        break;
      }
      case "random":
      default: {
        sampled = randomSampleRows(dataForSampling, sampleRows, { seed: 1 });
        break;
      }
    }

    const attachmentData = buildRangeAttachmentSectionText(
      { sheet: sheetForContext, attachments: params.attachments },
      { maxRows: 30, maxAttachments: 3 },
    );

    const shouldReturnRedactedStructured = Boolean(dlp) && dlpDecision?.decision === DLP_DECISION.REDACT;
    const includeRestrictedContentForStructured =
      dlp?.includeRestrictedContent ?? dlp?.include_restricted_content ?? false;
    const schemaOut = shouldReturnRedactedStructured
      ? redactStructuredValue(schema, this.redactor, { signal, includeRestrictedContent: includeRestrictedContentForStructured })
      : schema;
    const sampledOut = shouldReturnRedactedStructured
      ? redactStructuredValue(sampled, this.redactor, { signal, includeRestrictedContent: includeRestrictedContentForStructured })
      : sampled;
    const retrievedOut = shouldReturnRedactedStructured
      ? redactStructuredValue(retrieved, this.redactor, { signal, includeRestrictedContent: includeRestrictedContentForStructured })
      : retrieved;

    const sections = [
      ...((dlpRedactedCells > 0 || (dlpDecision?.decision === DLP_DECISION.REDACT && dlpHeuristicApplied))
         ? [
             {
               key: "dlp",
               priority: 5,
               text:
                 dlpRedactedCells > 0
                   ? `DLP: ${dlpRedactedCells} cells were redacted due to policy.`
                   : `DLP: sensitive patterns were redacted due to policy.`,
             },
           ]
         : []),
      ...(attachmentData
        ? [
            {
              key: "attachment_data",
              // Slightly below DLP policy notes, but above retrieved/schema/samples.
              priority: 4.5,
              text: this.redactor(attachmentData),
            },
          ]
        : []),
      {
        key: "schema",
        priority: 3,
        text: this.redactor(`Sheet schema (schema-first):\n${stableJsonStringify(schemaOut)}`),
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
        text: sampledOut.length
          ? this.redactor(`Sample rows:\n${sampledOut.map((r) => JSON.stringify(r)).join("\n")}`)
          : "",
      },
      {
        key: "retrieved",
        priority: 4,
        text: retrievedOut.length ? this.redactor(`Retrieved context:\n${stableJsonStringify(retrievedOut)}`) : "",
      },
    ].filter((s) => s.text);

    throwIfAborted(signal);
    const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens, this.estimator, { signal });
    throwIfAborted(signal);

    if (dlp) {
      dlp.auditLogger?.log({
        type: "ai.context",
        documentId: dlpAuditDocumentId ?? dlp.documentId,
        sheetId: dlpAuditSheetId ?? dlp.sheetId ?? rawSheet.name,
        sheetName: rawSheet.name,
        decision: dlpDecision,
        selectionClassification: dlpSelectionClassification,
        redactedCellCount: dlpRedactedCells,
      });
    }

    return {
      schema: schemaOut,
      retrieved: retrievedOut,
      sampledRows: sampledOut,
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
    const dlp = normalizeDlpOptions(params.dlp);
    const includeRestrictedContent = dlp?.includeRestrictedContent ?? false;
    const classificationRecords =
      dlp?.classificationRecords ?? dlp?.classificationStore?.list?.(dlp.documentId) ?? [];

    // Some hosts (notably the desktop DocumentController) keep a stable internal sheet id
    // even after a user renames the sheet. In those cases:
    // - RAG chunk metadata uses the user-facing display name (better retrieval quality)
    // - Structured DLP classifications are recorded against the stable sheet id
    //
    // When a resolver is provided, map chunk `metadata.sheetName` back to the stable id
    // before applying structured DLP classification.
    const dlpSheetNameResolver = (dlp && dlp.sheetNameResolver) || null;
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
        dlpDocumentIndex = buildDlpDocumentIndex({ documentId: dlp.documentId, records: classificationRecords, signal });
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
                    ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range }, signal)
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
    const hits = await searchWorkbookRag({
      queryText: queryForEmbedding,
      workbookId: params.workbook.id,
      topK,
      vectorStore,
      embedder,
      signal,
    });

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
            ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range }, signal)
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
      const schemaLines = [];
      const tables = Array.isArray(params.workbook?.tables) ? params.workbook.tables : [];
      const namedRanges = Array.isArray(params.workbook?.namedRanges) ? params.workbook.namedRanges : [];
      const sheetByName = new Map((params.workbook?.sheets ?? []).map((s) => [s.name, s]));

      const maxTables = 25;
      const maxColumns = 25;
      const maxTypeSampleRows = 50;

      // Prefer stable output ordering to help with cacheability and tests.
      const sortedTables = tables
        .slice()
        .sort((a, b) =>
          String(a?.sheetName ?? "").localeCompare(String(b?.sheetName ?? "")) ||
          String(a?.name ?? "").localeCompare(String(b?.name ?? ""))
        );

      for (let i = 0; i < sortedTables.length && i < maxTables; i++) {
        throwIfAborted(signal);
        const table = sortedTables[i];
        if (!table || typeof table !== "object") continue;
        const name = typeof table.name === "string" ? table.name : "";
        const sheetName = typeof table.sheetName === "string" ? table.sheetName : "";
        const rect = table.rect;
        if (!name || !sheetName || !rect || typeof rect !== "object") continue;
        const r0 = rect.r0;
        const c0 = rect.c0;
        const r1 = rect.r1;
        const c1 = rect.c1;
        if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) continue;
        if (r1 < r0 || c1 < c0) continue;

        const sheet = sheetByName.get(sheetName);
        if (!sheet) continue;

        // Structured DLP classifications: if the table's range is disallowed due to explicit
        // selectors (document/sheet/range), omit column details entirely (match chunk redaction).
        if (dlp) {
          const range = rectToRange(rect);
          const sheetId = sheetName ? resolveDlpSheetId(sheetName) : "";
          if (range && sheetId) {
            const index = getDlpDocumentIndex();
            const recordClassification = index
              ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range }, signal)
              : effectiveRangeClassification({ documentId: dlp.documentId, sheetId, range }, classificationRecords);
            const recordDecision = evaluatePolicy({
              action: DLP_ACTION.AI_CLOUD_PROCESSING,
              classification: recordClassification,
              policy: dlp.policy,
              options: { includeRestrictedContent },
            });
            if (recordDecision.decision !== DLP_DECISION.ALLOW) {
              schemaLines.push(`- Table ${name} (sheet="${sheetName}"): [REDACTED]`);
              continue;
            }
          }
        }

        const colCount = Math.max(0, c1 - c0 + 1);
        const boundedColCount = Math.min(colCount, maxColumns);

        /** @type {unknown[]} */
        const headerRowValues = [];
        /** @type {unknown[]} */
        const nextRowValues = [];
        for (let offset = 0; offset < boundedColCount; offset++) {
          throwIfAborted(signal);
          const c = c0 + offset;
          headerRowValues.push(getWorkbookSheetCell(sheet, r0, c));
          nextRowValues.push(getWorkbookSheetCell(sheet, r0 + 1, c));
        }

        const hasHeader = isLikelyHeaderRow(headerRowValues, nextRowValues);
        const dataStartRow = hasHeader ? r0 + 1 : r0;

        const cols = [];
        for (let offset = 0; offset < boundedColCount; offset++) {
          throwIfAborted(signal);
          const rawHeader = headerRowValues[offset];
          const header =
            hasHeader &&
            typeof rawHeader === "string" &&
            rawHeader.trim() !== "" &&
            !rawHeader.trim().startsWith("=") &&
            !/^[+-]?\d+(?:\.\d+)?$/.test(rawHeader.trim())
              ? rawHeader.trim()
              : `Column${offset + 1}`;

          /** @type {unknown[]} */
          const samples = [];
          for (let r = dataStartRow; r <= r1 && samples.length < maxTypeSampleRows; r++) {
            throwIfAborted(signal);
            const v = getWorkbookSheetCell(sheet, r, c0 + offset);
            if (v === undefined) continue;
            if (isEmptyScalar(v)) continue;
            samples.push(v);
          }
          const type = inferColumnType(samples, { signal });
          cols.push(`${header} (${type})`);
        }

        const colSuffix = cols.length < colCount ? " | …" : "";
        schemaLines.push(`- Table ${name} (sheet="${sheetName}"): ${cols.join(" | ")}${colSuffix}`);
      }

      // Named ranges can be helpful anchors even without cell samples.
      const sortedNamedRanges = namedRanges
        .slice()
        .sort((a, b) =>
          String(a?.sheetName ?? "").localeCompare(String(b?.sheetName ?? "")) ||
          String(a?.name ?? "").localeCompare(String(b?.name ?? ""))
        );
      const maxNamedRanges = 25;
      for (let i = 0; i < sortedNamedRanges.length && i < maxNamedRanges; i++) {
        throwIfAborted(signal);
        const nr = sortedNamedRanges[i];
        if (!nr || typeof nr !== "object") continue;
        const name = typeof nr.name === "string" ? nr.name : "";
        const sheetName = typeof nr.sheetName === "string" ? nr.sheetName : "";
        const rect = nr.rect;
        if (!name || !sheetName || !rect || typeof rect !== "object") continue;
        schemaLines.push(`- Named range ${name} (sheet="${sheetName}")`);
      }

      const workbookSchemaText = schemaLines.length ? this.redactor(`Workbook schema (schema-first):\n${schemaLines.join("\n")}`) : "";

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
          key: "workbook_schema",
          // Keep between workbook_summary (3) and retrieved (4).
          priority: 3.5,
          text: workbookSchemaText,
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
      const packed = packSectionsToTokenBudget(sections, this.tokenBudgetTokens, this.estimator, { signal });
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
    let dlp = normalizeDlpOptions(params.dlp);
    if (dlp && spreadsheetResolver && !dlp.sheetNameResolver) {
      dlp = { ...dlp, sheetNameResolver: spreadsheetResolver };
    }

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

/**
 * @param {unknown[][]} values
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {ReturnType<typeof classifyText>}
 */
function classifyValuesForDlp(values, options = {}) {
  const signal = options.signal;
  /** @type {Set<string>} */
  const findings = new Set();
  for (const row of values || []) {
    throwIfAborted(signal);
    for (const cell of row || []) {
      throwIfAborted(signal);
      if (cell === null || cell === undefined) continue;
      const heuristic = classifyText(String(cell));
      if (heuristic.level !== "sensitive") continue;
      for (const f of heuristic.findings || []) findings.add(String(f));
      // Early exit once we've found at least one sensitive pattern; policy evaluation only
      // needs the max classification, not exhaustive findings.
      if (findings.size > 0) {
        return { level: "sensitive", findings: [...findings] };
      }
    }
  }
  return { level: "public", findings: [] };
}

/**
 * @param {{level: string, findings: string[]}} heuristic
 */
function heuristicToPolicyClassification(heuristic) {
  if (heuristic?.level === "sensitive") {
    const labels = (heuristic.findings || []).map((f) => `heuristic:${f}`);
    return { level: CLASSIFICATION_LEVEL.RESTRICTED, labels };
  }
  return { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
}

/**
 * Apply a redactor to every string cell in a sheet values matrix.
 * @param {unknown[][]} values
 * @param {(text: string) => string} redactor
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {unknown[][]}
 */
function redactValuesForDlp(values, redactor, options = {}) {
  const signal = options.signal;
  const includeRestrictedContent = options.includeRestrictedContent ?? false;
  /** @type {unknown[][]} */
  const out = [];
  for (const row of values || []) {
    throwIfAborted(signal);
    if (!Array.isArray(row)) {
      out.push([]);
      continue;
    }
    const nextRow = [];
    for (const cell of row) {
      throwIfAborted(signal);
      if (typeof cell !== "string") {
        nextRow.push(cell);
        continue;
      }
      const redacted = redactor(cell);
      // Defense-in-depth: if the configured redactor is a no-op (or incomplete),
      // ensure heuristic sensitive patterns never slip through under DLP redaction.
      if (!includeRestrictedContent && classifyText(redacted).level === "sensitive") {
        nextRow.push("[REDACTED]");
        continue;
      }
      nextRow.push(redacted);
    }
    out.push(nextRow);
  }
  return out;
}

/**
 * Deeply apply a redactor to string fields inside structured output objects.
 *
 * This is used to ensure `buildContext()` does not leak sensitive values through its
 * structured return payload when DLP requires redaction.
 *
 * @template T
 * @param {T} value
 * @param {(text: string) => string} redactor
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {T}
 */
function redactStructuredValue(value, redactor, options = {}) {
  const signal = options.signal;
  const includeRestrictedContent = options.includeRestrictedContent ?? false;
  throwIfAborted(signal);

  if (typeof value === "string") {
    const redacted = redactor(value);
    if (!includeRestrictedContent && classifyText(redacted).level === "sensitive") {
      return /** @type {T} */ ("[REDACTED]");
    }
    return /** @type {T} */ (redacted);
  }
  if (value === null || value === undefined) return value;
  if (typeof value !== "object") return value;
  if (value instanceof Date) return value;

  if (Array.isArray(value)) {
    return /** @type {T} */ (
      value.map((v) => redactStructuredValue(v, redactor, { signal, includeRestrictedContent }))
    );
  }

  const proto = Object.getPrototypeOf(value);
  if (proto !== Object.prototype && proto !== null) {
    // Avoid walking exotic objects (Map, Set, class instances).
    return value;
  }

  /** @type {any} */
  const out = {};
  for (const [key, v] of Object.entries(value)) {
    throwIfAborted(signal);
    out[key] = redactStructuredValue(v, redactor, { signal, includeRestrictedContent });
  }
  return out;
}

function normalizeSheetNameForComparison(name) {
  const raw = typeof name === "string" ? name.trim() : "";
  if (!raw) return "";
  if (raw.startsWith("'") && raw.endsWith("'") && raw.length >= 2) {
    return raw.slice(1, -1).replace(/''/g, "'");
  }
  return raw;
}

function getSheetOrigin(sheet) {
  if (!sheet || typeof sheet !== "object") return { row: 0, col: 0 };
  const origin = sheet.origin;
  if (!origin || typeof origin !== "object") return { row: 0, col: 0 };
  const row = Number.isInteger(origin.row) && origin.row >= 0 ? origin.row : 0;
  const col = Number.isInteger(origin.col) && origin.col >= 0 ? origin.col : 0;
  return { row, col };
}

/**
 * @param {{ sheet: { name: string, values: unknown[][], origin?: { row: number, col: number } }, attachments?: Attachment[] }} params
 * @param {{ maxRows?: number, maxAttachments?: number }} [options]
 */
function buildRangeAttachmentSectionText(params, options = {}) {
  const attachments = Array.isArray(params.attachments) ? params.attachments : [];
  if (attachments.length === 0) return "";
  const sheet = params.sheet;
  const sheetName = sheet?.name ?? "";
  const normalizedSheetName = normalizeSheetNameForComparison(sheetName);
  const maxRows = options.maxRows ?? 30;
  const maxAttachments = options.maxAttachments ?? 3;

  const values = Array.isArray(sheet?.values) ? sheet.values : [];
  const matrixRowCount = values.length;
  let matrixColCount = 0;
  for (const row of values) {
    if (!Array.isArray(row)) continue;
    if (row.length > matrixColCount) matrixColCount = row.length;
  }

  const origin = getSheetOrigin(sheet);
  const availableRange =
    matrixRowCount > 0 && matrixColCount > 0
      ? rangeToA1({
          sheetName,
          startRow: origin.row,
          startCol: origin.col,
          endRow: origin.row + matrixRowCount - 1,
          endCol: origin.col + matrixColCount - 1,
        })
      : "";

  const entries = [];
  let included = 0;

  for (const attachment of attachments) {
    if (included >= maxAttachments) break;
    if (!attachment || attachment.type !== "range" || typeof attachment.reference !== "string") continue;

    let parsed;
    try {
      parsed = parseA1Range(attachment.reference);
    } catch {
      continue;
    }

    const attachmentSheetName = normalizeSheetNameForComparison(parsed.sheetName);
    if (attachmentSheetName && attachmentSheetName !== normalizedSheetName) continue;

    const canonicalRange = rangeToA1({ ...parsed, sheetName });

    if (matrixRowCount === 0 || matrixColCount === 0) {
      entries.push(`${canonicalRange}: (no sheet values available to preview)`);
      included += 1;
      continue;
    }

    const local = {
      startRow: parsed.startRow - origin.row,
      startCol: parsed.startCol - origin.col,
      endRow: parsed.endRow - origin.row,
      endCol: parsed.endCol - origin.col,
    };

    const intersects =
      !(local.endRow < 0 || local.endCol < 0 || local.startRow >= matrixRowCount || local.startCol >= matrixColCount);

    if (!intersects) {
      const windowSuffix = availableRange ? ` (${availableRange})` : "";
      entries.push(`${canonicalRange}: (range is outside the available sheet window${windowSuffix})`);
      included += 1;
      continue;
    }

    const clamped = {
      startRow: Math.max(0, local.startRow),
      startCol: Math.max(0, local.startCol),
      endRow: Math.min(matrixRowCount - 1, local.endRow),
      endCol: Math.min(matrixColCount - 1, local.endCol),
    };

    const matrix = [];
    for (let r = clamped.startRow; r <= clamped.endRow; r++) {
      const row = values[r] ?? [];
      matrix.push(row.slice(clamped.startCol, clamped.endCol + 1));
    }

    entries.push(`${canonicalRange}:\n${matrixToTsv(matrix, { maxRows })}`);
    included += 1;
  }

  if (entries.length === 0) return "";
  return `Attached range data:\n${entries.join("\n\n")}`;
}

/**
 * @param {unknown[][]} matrix
 * @param {{ maxRows: number }} options
 */
function matrixToTsv(matrix, options) {
  const lines = [];
  const limit = Math.min(matrix.length, options.maxRows);
  for (let r = 0; r < limit; r++) {
    const row = matrix[r];
    lines.push((row ?? []).map((v) => (isCellEmpty(v) ? "" : String(v))).join("\t"));
  }
  if (matrix.length > limit) lines.push(`… (${matrix.length - limit} more rows)`);
  return lines.join("\n");
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

function buildDlpRangeIndex(ref, records, opts) {
  const signal = opts?.signal;
  throwIfAborted(signal);
  const selectionRange = ref.range;
  const startRow = selectionRange.start.row;
  const startCol = selectionRange.start.col;
  const rowCount = selectionRange.end.row - selectionRange.start.row + 1;
  const colCount = selectionRange.end.col - selectionRange.start.col + 1;
  const maxAllowedRank = opts?.maxAllowedRank ?? DEFAULT_CLASSIFICATION_RANK;

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
    throwIfAborted(signal);
    if (!record || !record.selector || typeof record.selector !== "object") continue;
    const selector = record.selector;
    if (selector.documentId !== ref.documentId) continue;

    // The per-cell allow/redact decision depends only on the max classification level.
    // Records at/below the policy `maxAllowed` threshold cannot change allow/redact decisions
    // and are ignored for performance (reduces index build and avoids allocating dense arrays
    // for allowed-only classifications).
    const recordRank = rankFromClassification(record.classification);
    if (recordRank <= maxAllowedRank) continue;

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
    throwIfAborted(signal);
    rangeRecords.sort((a, b) => b.rank - a.rank);
  }

  return {
    docRankMax,
    sheetRankMax,
    baseRank: Math.max(docRankMax, sheetRankMax),
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
  const { index, maxAllowedRank, includeRestrictedContent, policyAllowsRestrictedContent, signal } = params;
  throwIfAborted(signal);
  if (maxAllowedRank === null) return false;

  // If we're explicitly including restricted content and policy allows it, a cell can become
  // ALLOW even if its classification exceeds `maxAllowed` (evaluatePolicy short-circuits for
  // Restricted + includeRestrictedContent).
  const restrictedOverrideAllowed = includeRestrictedContent && policyAllowsRestrictedContent;
  const canShortCircuitOverThreshold = !restrictedOverrideAllowed;
  const restrictedAllowed = includeRestrictedContent ? policyAllowsRestrictedContent : maxAllowedRank >= RESTRICTED_CLASSIFICATION_RANK;

  let rank = index.baseRank;

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
      throwIfAborted(signal);
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
  const signal = params.signal;
  throwIfAborted(signal);
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
    throwIfAborted(signal);
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

function effectiveRangeClassificationFromDocumentIndex(index, rangeRef, signal) {
  throwIfAborted(signal);
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
      throwIfAborted(signal);
      const colClassification = colMap.get(col);
      if (!colClassification) continue;
      classification = maxClassification(classification, colClassification);
      if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;
    }
  }

  const rangeRecords = index.rangeRecordsBySheetId.get(rangeRef.sheetId) ?? [];
  for (const record of rangeRecords) {
    throwIfAborted(signal);
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
        throwIfAborted(signal);
        for (let col = normalized.start.col; col <= normalized.end.col; col++) {
          throwIfAborted(signal);
          const cellClassification = cellMap.get(`${row},${col}`);
          if (!cellClassification) continue;
          classification = maxClassification(classification, cellClassification);
          if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) return classification;
        }
      }
    } else {
      for (const [key, cellClassification] of cellMap.entries()) {
        throwIfAborted(signal);
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
