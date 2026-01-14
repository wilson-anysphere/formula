import { extractSheetSchema } from "./schema.js";
import { RagIndex } from "./rag.js";
import { deleteSheetRegionChunks } from "./ragIds.js";
import { valuesRangeToTsv } from "./tsv.js";
import { DEFAULT_TOKEN_ESTIMATOR, packSectionsToTokenBudget, stableJsonStringify } from "./tokenBudget.js";
import { headSampleRows, randomSampleRows, stratifiedSampleRows, systematicSampleRows, tailSampleRows } from "./sampling.js";
import { classifyText, redactText } from "./dlp.js";
import { parseA1Range, rangeToA1 } from "./a1.js";
import { awaitWithAbort, throwIfAborted } from "./abort.js";
import { extractWorkbookSchema } from "./workbookSchema.js";
import { summarizeSheetSchema } from "./summarizeSheet.js";

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
const DEFAULT_RAG_CHUNK_ROW_OVERLAP = 3;
const DEFAULT_RAG_MAX_CHUNKS_PER_REGION = 50;
const DEFAULT_MAX_CONTEXT_COLS = 500;
const SHEET_INDEX_SIGNATURE_VERSION = 1;
const SHEET_SCHEMA_SIGNATURE_VERSION = 1;

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
  const stack = new WeakSet();

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

    if (v && typeof v === "object") {
      if (stack.has(v)) return fnv1a64Update(hash, "[Circular]");
      stack.add(v);
      try {
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
      } finally {
        stack.delete(v);
      }
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
 * @param {{
 *   maxChunkRows?: number,
 *   splitRegions?: boolean,
 *   chunkRowOverlap?: number,
 *   maxChunksPerRegion?: number,
 *   valuesHash?: string,
 *   signal?: AbortSignal
 * }} [options]
 */
function computeSheetIndexSignature(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const origin = normalizeSheetOrigin(sheet?.origin);
  const maxChunkRows = options.maxChunkRows ?? DEFAULT_RAG_MAX_CHUNK_ROWS;
  const valuesHash = options.valuesHash ?? stableHashValue(sheet?.values ?? [], { signal });
  const splitRegions = options.splitRegions === true;
  const chunkRowOverlap = splitRegions
    ? normalizeOptionalNonNegativeInt(options.chunkRowOverlap) ?? DEFAULT_RAG_CHUNK_ROW_OVERLAP
    : undefined;
  const maxChunksPerRegion = splitRegions
    ? normalizeOptionalNonNegativeInt(options.maxChunksPerRegion) ?? DEFAULT_RAG_MAX_CHUNKS_PER_REGION
    : undefined;

  let hash = FNV_OFFSET_64;
  hash = fnv1a64Update(hash, `sig:v${SHEET_INDEX_SIGNATURE_VERSION}\n`);
  hash = fnv1a64Update(hash, `name:${sheet?.name ?? ""}\n`);
  hash = fnv1a64Update(hash, `origin:${origin.row},${origin.col}\n`);
  hash = fnv1a64Update(hash, `maxChunkRows:${String(maxChunkRows)}\n`);
  hash = fnv1a64Update(hash, `splitRegions:${splitRegions ? "1" : "0"}\n`);
  if (splitRegions) {
    hash = fnv1a64Update(hash, `chunkRowOverlap:${String(chunkRowOverlap)}\n`);
    hash = fnv1a64Update(hash, `maxChunksPerRegion:${String(maxChunksPerRegion)}\n`);
  }
  hash = fnv1a64Update(hash, "values:");
  hash = fnv1a64Update(hash, valuesHash);
  return hash.toString(16).padStart(16, "0");
}

/**
 * Deterministic signature of inputs that affect schema extraction.
 *
 * This is intentionally separate from the RAG index signature so callers can update
 * schema-only metadata (e.g. named ranges / tables) without re-embedding.
 *
 * @param {{ name?: string, origin?: any, values?: unknown[][], tables?: any, namedRanges?: any }} sheet
 * @param {{ valuesHash?: string, signal?: AbortSignal }} [options]
 */
function computeSheetSchemaSignature(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const origin = normalizeSheetOrigin(sheet?.origin);
  const valuesHash = options.valuesHash ?? stableHashValue(sheet?.values ?? [], { signal });
  const tablesHash = stableHashValue(sheet?.tables ?? [], { signal });
  const namedRangesHash = stableHashValue(sheet?.namedRanges ?? [], { signal });

  let hash = FNV_OFFSET_64;
  hash = fnv1a64Update(hash, `schema:v${SHEET_SCHEMA_SIGNATURE_VERSION}\n`);
  hash = fnv1a64Update(hash, `name:${sheet?.name ?? ""}\n`);
  hash = fnv1a64Update(hash, `origin:${origin.row},${origin.col}\n`);
  hash = fnv1a64Update(hash, "values:");
  hash = fnv1a64Update(hash, valuesHash);
  hash = fnv1a64Update(hash, "\ntables:");
  hash = fnv1a64Update(hash, tablesHash);
  hash = fnv1a64Update(hash, "\nnamedRanges:");
  hash = fnv1a64Update(hash, namedRangesHash);
  return hash.toString(16).padStart(16, "0");
}

/**
 * @param {unknown} value
 * @param {number} fallback
 */
function normalizeNonNegativeInt(value, fallback) {
  if (value === undefined || value === null) return fallback;
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return fallback;
  return Math.floor(n);
}

/**
 * @param {unknown} value
 * @returns {number | undefined}
 */
function normalizeOptionalNonNegativeInt(value) {
  if (value === undefined || value === null) return undefined;
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return undefined;
  return Math.floor(n);
}

// NOTE: workbook schema extraction helpers live in `workbookSchema.js`. ContextManager
// only formats the schema for prompt inclusion (and applies DLP/redaction).

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
 * Normalize a workbook/table rect to the canonical `{ r0, c0, r1, c1 }` shape.
 *
 * Some hosts may accidentally attach extra metadata fields to rect objects (or use alternative
 * shapes). Under DLP redaction we treat those extra fields as prompt-unsafe because they can
 * contain arbitrary non-heuristic sensitive strings (e.g. "TopSecret") that a no-op redactor
 * cannot detect.
 *
 * @param {any} rect
 * @returns {{ r0: number, c0: number, r1: number, c1: number } | null}
 */
function normalizeWorkbookRect(rect) {
  if (!rect || typeof rect !== "object") return null;
  let r0 = rect.r0;
  let c0 = rect.c0;
  let r1 = rect.r1;
  let c1 = rect.c1;

  // Range-like shapes.
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) {
    r0 = rect.startRow;
    c0 = rect.startCol;
    r1 = rect.endRow;
    c1 = rect.endCol;
  }

  // Nested `{ start: {row,col}, end: {row,col} }`.
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) {
    r0 = rect.start?.row;
    c0 = rect.start?.col;
    r1 = rect.end?.row;
    c1 = rect.end?.col;
  }

  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return null;
  const rr0 = /** @type {number} */ (r0);
  const rr1 = /** @type {number} */ (r1);
  const cc0 = /** @type {number} */ (c0);
  const cc1 = /** @type {number} */ (c1);
  return { r0: Math.min(rr0, rr1), c0: Math.min(cc0, cc1), r1: Math.max(rr0, rr1), c1: Math.max(cc0, cc1) };
}

/**
 * Normalize a workbook chunk kind to one of the known `ai-rag` kinds.
 *
 * Vector stores can contain arbitrary metadata fields (including `kind`). Under DLP enforcement we treat
 * unknown kind strings as prompt-unsafe because they can contain non-heuristic secrets (e.g. "TopSecret")
 * that a no-op redactor cannot detect.
 *
 * @param {unknown} kind
 * @returns {"chunk" | "table" | "namedRange" | "dataRegion" | "formulaRegion"}
 */
function normalizeWorkbookChunkKind(kind) {
  // Avoid calling `.toString()` on arbitrary user-provided objects (3p vector stores can
  // persist non-plain metadata). Treat non-string kinds as unknown.
  const raw = typeof kind === "string" ? kind.trim() : "";
  if (!raw) return "chunk";
  const lowered = raw.toLowerCase();
  if (lowered === "table") return "table";
  if (lowered === "namedrange") return "namedRange";
  if (lowered === "dataregion") return "dataRegion";
  if (lowered === "formularegion") return "formulaRegion";
  if (lowered === "chunk") return "chunk";
  return "chunk";
}

/**
 * Detect whether a workbook chunk kind is unknown/untrusted.
 *
 * @param {unknown} kind
 */
function isWorkbookChunkKindUnknown(kind) {
  if (kind === null || kind === undefined) return false;
  if (typeof kind !== "string") return true;
  const raw = kind.trim().toLowerCase();
  if (!raw) return false;
  return raw !== "chunk" && raw !== "table" && raw !== "namedrange" && raw !== "dataregion" && raw !== "formularegion";
}

/**
 * Filter vector-store chunk metadata before returning it to callers under structured DLP redaction.
 *
 * Vector stores (especially third-party ones) may persist additional metadata fields. Those fields can
 * contain arbitrary, user-controlled identifiers (e.g. workbook ids, file paths) that are not reliably
 * detectable by heuristic redaction. When structured DLP requires redaction anywhere in the workbook,
 * conservatively drop unknown metadata keys so non-heuristic secrets cannot leak even if the configured
 * redactor is a no-op.
 *
 * @param {any} metadata
 * @returns {any}
 */
function filterWorkbookChunkMetadataForOutput(metadata) {
  if (!metadata || typeof metadata !== "object" || Array.isArray(metadata)) return metadata;
  const allowedKeys = [
    "workbookId",
    "sheetName",
    "kind",
    "title",
    "rect",
    "embedder",
    "contentHash",
    "metadataHash",
    "tokenCount",
  ];
  /** @type {any} */
  const out = {};
  for (const key of allowedKeys) {
    if (!Object.prototype.hasOwnProperty.call(metadata, key)) continue;
    if (key === "kind") {
      out.kind = normalizeWorkbookChunkKind(metadata.kind);
      continue;
    }

    // Metadata tokens can be arbitrarily shaped in third-party stores. Under structured DLP redaction,
    // ensure we never return nested objects that could contain non-heuristic secrets.
    const value = metadata[key];
    if (key === "rect") {
      out.rect = value;
      continue;
    }
    if (key === "tokenCount") {
      const n = Number(value);
      if (Number.isFinite(n) && n >= 0) out.tokenCount = Math.floor(n);
      continue;
    }
    if (
      key === "workbookId" ||
      key === "sheetName" ||
      key === "title" ||
      key === "embedder" ||
      key === "contentHash" ||
      key === "metadataHash"
    ) {
      if (typeof value === "string") out[key] = value;
      else if (value !== null && value !== undefined) out[key] = "[REDACTED]";
      continue;
    }

    // Any other allowlisted keys are treated as safe primitives; drop otherwise.
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
      out[key] = value;
    }
  }
  return out;
}

/**
 * Prompt-sanitize attachment objects.
 *
 * Some callers (e.g. UI layers) may attach rich `data` payloads for selections or tables
 * (including sampled cell values). Those payloads are prompt-unsafe and can bypass
 * DLP/redaction if they are inlined directly into a model prompt.
 *
 * As a safe default, we never include raw `data` for `range` or `table` attachments in
 * prompt context. `formula`/`chart` attachments keep their bounded `data` payloads.
 *
 * @param {unknown[] | null | undefined} attachments
 * @param {{ dropAllData?: boolean }} [options]
 * @returns {unknown[] | null | undefined}
 */
function compactAttachmentsForPrompt(attachments, options = {}) {
  const dropAllData = options.dropAllData === true;
  if (!Array.isArray(attachments)) {
    if (!dropAllData) return attachments;
    if (attachments === null || attachments === undefined) return attachments;
    if (typeof attachments === "string" && attachments.trim() === "") return attachments;
    // Under structured DLP redaction, treat malformed non-array attachment payloads as
    // prompt-unsafe (they can contain arbitrary non-heuristic sensitive strings).
    return ["[REDACTED]"];
  }
  return attachments.map((item) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      // Under structured DLP redaction, treat any non-object attachment entries as
      // prompt-unsafe (they can contain arbitrary non-heuristic sensitive strings).
      if (dropAllData) return "[REDACTED]";
      return item;
    }
    const type = item.type;
    const reference = item.reference;
    if (dropAllData) {
      // Under structured DLP redaction, attachments may contain arbitrary nested payloads
      // (including raw cell values) that are not reliably detectable by heuristic redaction.
      // Keep only a minimal prompt-safe skeleton.
      const allowedTypes = new Set(["range", "formula", "table", "chart"]);
      const normalizedType = typeof type === "string" ? type.trim().toLowerCase() : "";
      const safeType = allowedTypes.has(normalizedType) ? normalizedType : "[REDACTED]";
      return {
        type: safeType,
        reference: typeof reference === "string" ? reference : "[REDACTED]",
      };
    }
    if (type !== "range" && type !== "table") return item;
    // Drop raw data (can contain copied workbook values).
    const { data: _data, ...rest } = /** @type {any} */ (item);
    return rest;
  });
}

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
   *   sheetIndexCacheLimit?: number,
   *   workbookRag?: WorkbookRagOptions,
   *   maxContextRows?: number,
   *   maxContextCols?: number,
   *   maxContextCells?: number,
   *   maxChunkRows?: number,
   *   /**
   *    * Split tall regions into multiple row windows for better retrieval quality.
   *    * Defaults to false (opt-in).
   *    *\/
   *   splitRegions?: boolean,
   *   /**
   *    * Row overlap between region windows (only when splitRegions is enabled).
   *    * Defaults to 3.
   *    *\/
   *   chunkRowOverlap?: number,
   *   /**
   *    * Maximum number of windows per region (only when splitRegions is enabled).
   *    * Defaults to 50.
   *    *\/
   *   maxChunksPerRegion?: number,
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
    // Explicit safety cap for extremely wide (but short) selections (e.g. 10 rows x 16k cols).
    this.maxContextCols = normalizeNonNegativeInt(options.maxContextCols, DEFAULT_MAX_CONTEXT_COLS);
    // `maxChunkRows` controls how many TSV rows are included in each RAG chunk's text.
    this.maxChunkRows = normalizeNonNegativeInt(options.maxChunkRows, 30);
    this.splitRegions = options.splitRegions === true;
    this.chunkRowOverlap = normalizeNonNegativeInt(options.chunkRowOverlap, DEFAULT_RAG_CHUNK_ROW_OVERLAP);
    this.maxChunksPerRegion = normalizeNonNegativeInt(options.maxChunksPerRegion, DEFAULT_RAG_MAX_CHUNKS_PER_REGION);
    // Top-K retrieved regions for sheet-level (non-workbook) RAG.
    this.sheetRagTopK = normalizeNonNegativeInt(options.sheetRagTopK, 5);

    this.cacheSheetIndex = options.cacheSheetIndex ?? true;
    /** @type {Map<string, { signature: string, schemaSignature: string, schema: any, sheetName: string }>} */
    this._sheetIndexCache = new Map();
    this._sheetIndexCacheLimit = Math.max(
      1,
      normalizeNonNegativeInt(options.sheetIndexCacheLimit, DEFAULT_SHEET_INDEX_CACHE_LIMIT),
    );
    /** @type {Map<string, string>} */
    this._sheetNameToActiveCacheKey = new Map();
    /** @type {Map<string, Promise<void>>} */
    this._sheetIndexLocks = new Map();
    /** @type {Promise<void>} */
    this._sheetIndexGlobalLock = Promise.resolve();
  }

  /**
   * Serialize operations that mutate the sheet-level RagIndex store across all sheets.
   *
   * This is used by `clearSheetIndexCache({ clearStore: true })` to ensure the store
   * is emptied deterministically (i.e. no concurrent index operation can re-add chunks
   * after the clear completes).
   *
   * @template T
   * @param {AbortSignal | undefined} signal
   * @param {() => Promise<T>} fn
   * @returns {Promise<T>}
   */
  async _withGlobalSheetIndexLock(signal, fn) {
    const prev = this._sheetIndexGlobalLock ?? Promise.resolve();
    /** @type {() => void} */
    let release = () => {};
    const current = new Promise((resolve) => {
      release = resolve;
    });
    const chain = prev.then(() => current);
    this._sheetIndexGlobalLock = chain;

    try {
      await awaitWithAbort(prev, signal);
      throwIfAborted(signal);
      return await fn();
    } finally {
      release();
    }
  }

  /**
   * Ensure only one sheet-level index operation runs at a time per sheet name.
   *
   * `RagIndex.indexSheet()` clears and re-adds chunks by a per-sheet prefix,
   * so concurrent indexing of the same sheet name can interleave deletes/adds and
   * leave a mixed chunk set in the in-memory store.
   *
   * @template T
   * @param {string} sheetName
   * @param {AbortSignal | undefined} signal
   * @param {() => Promise<T>} fn
   * @returns {Promise<T>}
   */
  async _withSheetIndexLock(sheetName, signal, fn) {
    const key = typeof sheetName === "string" ? sheetName : String(sheetName ?? "");
    // Ensure we don't start an index pass while a global sheet-store operation (clear) is running.
    await awaitWithAbort(this._sheetIndexGlobalLock, signal);
    throwIfAborted(signal);
    const prev = this._sheetIndexLocks.get(key) ?? Promise.resolve();
    /** @type {() => void} */
    let release = () => {};
    const current = new Promise((resolve) => {
      release = resolve;
    });
    const chain = prev.then(() => current);
    this._sheetIndexLocks.set(key, chain);
    let acquired = false;

    try {
      // Allow callers to abort while waiting for a previous indexing pass.
      await awaitWithAbort(prev, signal);
      acquired = true;
      throwIfAborted(signal);
      return await fn();
    } finally {
      release();
      if (acquired) {
        if (this._sheetIndexLocks.get(key) === chain) {
          this._sheetIndexLocks.delete(key);
        }
      } else {
        // If this call aborted while *waiting* for the lock, do not eagerly delete the
        // lock chain: doing so would allow subsequent calls to bypass the in-flight
        // operation they were waiting on.
        //
        // Instead, schedule cleanup after the chain settles (after the previous holder
        // completes). Only delete if we are still the tail.
        chain.finally(() => {
          if (this._sheetIndexLocks.get(key) === chain) {
            this._sheetIndexLocks.delete(key);
          }
        });
      }
    }
  }

  /**
   * Clear the single-sheet RAG indexing cache.
   *
   * Note: This only affects the in-memory, sheet-level RagIndex used by `buildContext()`.
   * It does not impact workbook-level RAG (`buildWorkbookContext()`), which uses a caller-
   * supplied persistent vector store.
   *
   * @param {{ clearStore?: boolean, signal?: AbortSignal }} [options]
   */
  async clearSheetIndexCache(options = {}) {
    const signal = options.signal;
    const clearStore = options.clearStore === true;
    throwIfAborted(signal);

    await this._withGlobalSheetIndexLock(signal, async () => {
      // Wait for all in-flight per-sheet indexing operations to finish before clearing the store.
      const locks = Array.from(this._sheetIndexLocks.values());
      if (locks.length) {
        await awaitWithAbort(Promise.allSettled(locks), signal);
        throwIfAborted(signal);
      }

      this._sheetIndexCache.clear();
      this._sheetNameToActiveCacheKey.clear();

      if (!clearStore) return;
      const store = this.ragIndex?.store;
      if (!store) return;

      // Prefer the store's deleteByPrefix API so callers can abort long clears.
      // Passing an empty prefix clears all ids (every string starts with "").
      if (typeof store.deleteByPrefix === "function") {
        await store.deleteByPrefix("", { signal });
        return;
      }

      // Fall back to common in-memory store shapes.
      if (typeof store.clear === "function") {
        await store.clear();
        return;
      }
      if (store.items && typeof store.items.clear === "function") {
        store.items.clear();
      }
    });
  }

  /**
   * Index a sheet into the in-memory RAG store, with incremental caching by sheet signature.
   *
   * Returns the extracted schema (reused from chunking/indexing when possible).
   *
   * @param {{ name: string, values: unknown[][], origin?: any }} sheet
   * @param {{
   *   signal?: AbortSignal,
   *   maxChunkRows?: number,
   *   splitRegions?: boolean,
   *   chunkRowOverlap?: number,
   *   maxChunksPerRegion?: number,
   * }} [options]
   * @returns {Promise<{ schema: any }>}
   */
  async _ensureSheetIndexed(sheet, options = {}) {
    const signal = options.signal;
    const maxChunkRows = options.maxChunkRows;
    const splitRegions = (options.splitRegions ?? this.splitRegions) === true;
    const chunkRowOverlap = splitRegions
      ? normalizeNonNegativeInt(options.chunkRowOverlap, this.chunkRowOverlap)
      : undefined;
    const maxChunksPerRegion = splitRegions
      ? normalizeNonNegativeInt(options.maxChunksPerRegion, this.maxChunksPerRegion)
      : undefined;
    throwIfAborted(signal);
    const sheetName = typeof sheet?.name === "string" ? sheet.name : String(sheet?.name ?? "");

    return await this._withSheetIndexLock(sheetName, signal, async () => {
      return await this._ensureSheetIndexedLocked(sheet, {
        signal,
        maxChunkRows,
        splitRegions,
        chunkRowOverlap,
        maxChunksPerRegion,
      });
    });
  }

  /**
   * Like `_ensureSheetIndexed()`, but assumes the per-sheet index lock is already held.
   *
   * This allows callers to keep the lock across multiple operations (e.g. indexing + search)
   * so other concurrent calls cannot swap out the underlying in-memory RAG store between
   * steps.
   *
   * @param {{ name: string, values: unknown[][], origin?: any }} sheet
   * @param {{
   *   signal?: AbortSignal,
   *   maxChunkRows?: number,
   *   splitRegions?: boolean,
   *   chunkRowOverlap?: number,
   *   maxChunksPerRegion?: number,
   * }} [options]
   * @returns {Promise<{ schema: any }>}
   */
  async _ensureSheetIndexedLocked(sheet, options = {}) {
    const signal = options.signal;
    const maxChunkRows = options.maxChunkRows;
    const splitRegions = options.splitRegions === true;
    const chunkRowOverlap = splitRegions ? options.chunkRowOverlap : undefined;
    const maxChunksPerRegion = splitRegions ? options.maxChunksPerRegion : undefined;
    throwIfAborted(signal);
    const sheetName = typeof sheet?.name === "string" ? sheet.name : String(sheet?.name ?? "");

    if (!this.cacheSheetIndex) {
      const indexStats = await this.ragIndex.indexSheet(sheet, {
        signal,
        maxChunkRows,
        splitRegions,
        chunkRowOverlap,
        maxChunksPerRegion,
      });
      const schema = indexStats?.schema ?? extractSheetSchema(sheet, { signal });
      return { schema };
    }

    const valuesHash = stableHashValue(sheet?.values ?? [], { signal });
    const cacheKey = sheetIndexCacheKey(sheet);
    const signature = computeSheetIndexSignature(sheet, {
      signal,
      maxChunkRows,
      splitRegions,
      chunkRowOverlap,
      maxChunksPerRegion,
      valuesHash,
    });
    const schemaSignature = computeSheetSchemaSignature(sheet, { signal, valuesHash });

    const cached = this._sheetIndexCache.get(cacheKey);
    if (cached) {
      // Refresh LRU on access.
      this._sheetIndexCache.delete(cacheKey);
      this._sheetIndexCache.set(cacheKey, cached);
    }

    const activeKey = this._sheetNameToActiveCacheKey.get(sheetName);
    const upToDate = cached?.signature === signature && activeKey === cacheKey;
    if (upToDate) {
      if (cached?.schemaSignature === schemaSignature) return { schema: cached.schema };
      // Schema-only metadata changed (e.g. named ranges / tables). Recompute schema without
      // re-indexing/embedding.
      const schema = extractSheetSchema(sheet, { signal });
      const nextCached = {
        signature,
        schemaSignature,
        schema,
        sheetName,
      };
      this._sheetIndexCache.delete(cacheKey);
      this._sheetIndexCache.set(cacheKey, nextCached);
      return { schema };
    }

    const indexStats = await this.ragIndex.indexSheet(sheet, {
      signal,
      maxChunkRows,
      splitRegions,
      chunkRowOverlap,
      maxChunksPerRegion,
    });
    const schema = indexStats?.schema ?? extractSheetSchema(sheet, { signal });

    // Update caches after successful indexing.
    this._sheetNameToActiveCacheKey.set(sheetName, cacheKey);
    this._sheetIndexCache.delete(cacheKey);
    this._sheetIndexCache.set(cacheKey, { signature, schemaSignature, schema, sheetName });
    while (this._sheetIndexCache.size > this._sheetIndexCacheLimit) {
      const oldestKey = this._sheetIndexCache.keys().next().value;
      if (oldestKey === undefined) break;
      const oldestEntry = this._sheetIndexCache.get(oldestKey);
      this._sheetIndexCache.delete(oldestKey);
      if (oldestEntry?.sheetName) {
        const evictedSheetName = oldestEntry.sheetName;
        const activeKeyForSheet = this._sheetNameToActiveCacheKey.get(evictedSheetName);
        if (activeKeyForSheet === oldestKey) {
          this._sheetNameToActiveCacheKey.delete(evictedSheetName);

          // Bound in-memory RAG storage as well as the signature cache. When a sheet's active
          // index entry is evicted from the LRU, delete the sheet's chunks from the vector
          // store so `RagIndex.search()` doesn't keep considering stale sheets forever.
          throwIfAborted(signal);
          await deleteSheetRegionChunks(this.ragIndex?.store, evictedSheetName, { signal });
        }
      }
    }

    return { schema };
  }

  /**
   * Build a compact context payload for chat prompts for a single sheet.
   *
   * @param {{
   *   sheet: {
   *     name: string,
   *     values: unknown[][],
   *     /**
   *      * Optional coordinate origin (0-based) for the provided `values` matrix.
   *      * When `values` is a cropped window of a larger sheet, `origin` lets schema
   *      * extraction, retrieval ranges, and DLP selectors refer to correct absolute
   *      * coordinates.
   *      *\/
   *     origin?: { row: number, col: number },
   *     namedRanges?: any[],
   *     tables?: any[],
   *   },
   *   query: string,
   *   attachments?: Attachment[],
   *   sampleRows?: number,
   *   samplingStrategy?: "random" | "stratified" | "head" | "tail" | "systematic",
   *   stratifyByColumn?: number,
   *   limits?: {
   *     maxContextRows?: number,
   *     maxContextCols?: number,
   *     maxContextCells?: number,
   *     maxChunkRows?: number,
   *     /**
   *      * Split tall regions into multiple row windows to improve retrieval quality.
   *      *\/
   *     splitRegions?: boolean,
   *     /**
   *      * Row overlap between region windows (only when splitRegions is enabled).
   *      *\/
   *     chunkRowOverlap?: number,
   *     /**
   *      * Maximum number of windows per region (only when splitRegions is enabled).
   *      *\/
   *     maxChunksPerRegion?: number,
   *   },
   *   signal?: AbortSignal,
   *   dlp?: {
   *     documentId: string,
   *     sheetId?: string,
   *     policy: any,
   *     classificationRecords?: Array<{ selector: any, classification: any }>,
   *     classificationStore?: { list(documentId: string): Array<{ selector: any, classification: any }> },
   *     includeRestrictedContent?: boolean,
   *     auditLogger?: { log(event: any): void },
   *     sheetNameResolver?: any,
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
    const safeExplicitColCap = normalizeNonNegativeInt(params.limits?.maxContextCols, this.maxContextCols);
    const safeCellCap = normalizeNonNegativeInt(params.limits?.maxContextCells, this.maxContextCells);
    const maxChunkRows = normalizeNonNegativeInt(params.limits?.maxChunkRows, this.maxChunkRows);
    const splitRegions = (params.limits?.splitRegions ?? this.splitRegions) === true;
    const chunkRowOverlap = splitRegions
      ? normalizeNonNegativeInt(params.limits?.chunkRowOverlap, this.chunkRowOverlap)
      : undefined;
    const maxChunksPerRegion = splitRegions
      ? normalizeNonNegativeInt(params.limits?.maxChunksPerRegion, this.maxChunksPerRegion)
      : undefined;
    const rawValues = Array.isArray(rawSheet?.values) ? rawSheet.values : [];
    // Respect both the row cap and the total cell cap.
    // If `maxContextRows` is larger than `maxContextCells`, we need to clamp the row count
    // further so we can still include at least one column per row without exceeding the
    // total cell budget.
    let rowCount = Math.min(rawValues.length, safeRowCap);
    if (safeCellCap === 0 || safeExplicitColCap === 0) rowCount = 0;
    else if (rowCount > safeCellCap) rowCount = safeCellCap;
    const safeColCap =
      rowCount > 0 ? Math.min(Math.floor(safeCellCap / rowCount), safeExplicitColCap) : 0;
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
    /** @type {ReturnType<typeof evaluatePolicy> | null} */
    let dlpStructuredDecision = null;
    /** @type {ReturnType<typeof evaluatePolicy> | null} */
    let dlpStructuredSheetDecision = null;
    let dlpHeuristic = null;
    let dlpHeuristicApplied = false;
    let dlpAuditDocumentId = null;
    let dlpAuditSheetId = null;
    /** @type {Array<{ selector: any, classification: any }>} */
    let dlpClassificationRecords = [];
    const policyAllowsRestrictedContent = Boolean(dlp?.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.allowRestrictedContent);

    if (dlp) {
      const documentId = dlp.documentId;
      const records = dlp.classificationRecords ?? dlp.classificationStore?.list?.(documentId) ?? [];
      const includeRestrictedContent = dlp.includeRestrictedContent ?? false;
      dlpClassificationRecords = records;

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

      // Structured sheet-level classification (document/sheet scopes only). This is used to
      // decide whether we should redact *metadata tokens* like the sheet name itself, which may
      // contain non-heuristic sensitive strings (e.g. "TopSecret") even when the configured
      // redactor is a no-op.
      //
      // Note: buildContext is window-scoped; range/cell selectors outside the provided origin
      // window should not force sheet-name redaction (that would make unrelated selectors affect
      // prompt ranges and break sheet display-name expectations).
      let structuredSheetClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
      for (const record of records) {
        throwIfAborted(signal);
        const selector = record?.selector;
        if (!selector || typeof selector !== "object") continue;
        if (selector.documentId !== documentId) continue;
        if (selector.scope === "document") {
          structuredSheetClassification = maxClassification(structuredSheetClassification, record.classification);
        } else if (selector.scope === "sheet" && selector.sheetId === sheetId) {
          structuredSheetClassification = maxClassification(structuredSheetClassification, record.classification);
        }
      }
      dlpStructuredSheetDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: structuredSheetClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });

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
       dlpStructuredDecision = structuredDecision;
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

      // Workbook DLP enforcement treats heuristic sensitive patterns as Restricted when evaluating
      // AI cloud processing policies. Mirror that behavior in the single-sheet context path so
      // callers can't accidentally leak sensitive content even when no structured selectors are present.
      dlpHeuristic = classifyValuesForDlp(valuesForContext, { signal });
      const heuristicPolicyClassification = heuristicToPolicyClassification(dlpHeuristic);
      const attachmentsHeuristic = classifyStructuredForDlp(params.attachments ?? [], { signal });
      const attachmentsPolicyClassification = heuristicToPolicyClassification(attachmentsHeuristic);
      const sheetNameHeuristic = classifyTextForDlp(String(rawSheet?.name ?? ""));
      const sheetNamePolicyClassification = heuristicToPolicyClassification(sheetNameHeuristic);
      const sheetMetaHeuristic = classifyStructuredForDlp(
        {
          tables: rawSheet?.tables ?? [],
          namedRanges: rawSheet?.namedRanges ?? [],
        },
        { signal },
      );
      const sheetMetaPolicyClassification = heuristicToPolicyClassification(sheetMetaHeuristic);
      const combinedClassification = maxClassification(
        maxClassification(maxClassification(structuredSelectionClassification, heuristicPolicyClassification), attachmentsPolicyClassification),
        maxClassification(sheetNamePolicyClassification, sheetMetaPolicyClassification),
      );
      dlpSelectionClassification = combinedClassification;
      if (
        heuristicPolicyClassification.level !== CLASSIFICATION_LEVEL.PUBLIC ||
        attachmentsPolicyClassification.level !== CLASSIFICATION_LEVEL.PUBLIC ||
        sheetNamePolicyClassification.level !== CLASSIFICATION_LEVEL.PUBLIC ||
        sheetMetaPolicyClassification.level !== CLASSIFICATION_LEVEL.PUBLIC
      ) {
        dlpHeuristicApplied = true;
      }
      // Always evaluate the combined policy decision so explicit (structured) DLP enforcement can
      // influence the overall decision even when heuristic classifications are all Public.
      dlpDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: combinedClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });

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

      // Under REDACT decisions, defensively apply heuristic redaction to the context sheet so:
      //  - schema / sampling / retrieval don't contain raw sensitive strings in structured outputs
      //  - in-memory RAG text doesn't retain sensitive patterns (defense-in-depth)
      if (dlpDecision.decision === DLP_DECISION.REDACT) {
        sheetForContext = {
          ...sheetForContext,
          values: redactValuesForDlp(sheetForContext.values, this.redactor, {
            signal,
            includeRestrictedContent,
            policyAllowsRestrictedContent,
          }),
        };
      }
    }

    throwIfAborted(signal);
    let queryForRag = params.query;
    if (dlp) {
      const queryHeuristic = classifyTextForDlp(params.query);
      const queryClassification = heuristicToPolicyClassification(queryHeuristic);
      const queryDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: queryClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent: dlp.includeRestrictedContent },
      });
      if (queryDecision.decision !== DLP_DECISION.ALLOW) {
        queryForRag = this.redactor(params.query);
        const restrictedAllowed = dlp.includeRestrictedContent && policyAllowsRestrictedContent;
        if (!restrictedAllowed && classifyTextForDlp(queryForRag).level === "sensitive") {
          queryForRag = "[REDACTED]";
        }
      }

      // Structured DLP can require redaction for non-heuristic metadata tokens (e.g. sheet/table/namedRange
      // names) that a no-op redactor cannot detect. If those disallowed identifiers appear in the
      // user's query, strip them before embedding to keep future cloud embedders safe.
      if (
        dlpStructuredDecision?.decision === DLP_DECISION.REDACT &&
        dlpClassificationRecords.length > 0 &&
        typeof queryForRag === "string"
      ) {
        let nextQuery = queryForRag;

        /**
         * @param {string} haystack
         * @param {string} needle
         * @param {string} replacement
         */
        const replaceAll = (haystack, needle, replacement) => {
          if (!needle) return haystack;
          if (!haystack.includes(needle)) return haystack;
          return haystack.split(needle).join(replacement);
        };

        /**
         * Replace a disallowed token and common encodings/quoting forms.
         * @param {string} token
         */
        const redactQueryToken = (token) => {
          const raw = String(token ?? "");
          if (!raw) return;
          nextQuery = replaceAll(nextQuery, raw, "[REDACTED]");
          const encoded = encodeURIComponent(raw);
          if (encoded && encoded !== raw) nextQuery = replaceAll(nextQuery, encoded, "[REDACTED]");
          // Excel-style quoted sheet names.
          const quoted = `'${raw.replace(/'/g, "''")}'`;
          if (quoted !== raw) nextQuery = replaceAll(nextQuery, quoted, "'[REDACTED]'");
        };

        // Treat the sheet name as a disallowed metadata token under structured DLP redaction.
        // (Note: this does not affect the returned A1 range strings; it only keeps embedding inputs safe.)
        const sheetNameToken = String(rawSheet?.name ?? "");
        if (
          sheetNameToken &&
          (nextQuery.includes(sheetNameToken) ||
            nextQuery.includes(encodeURIComponent(sheetNameToken)) ||
            nextQuery.includes(`'${sheetNameToken.replace(/'/g, "''")}'`))
        ) {
          redactQueryToken(sheetNameToken);
        }

        // Table/namedRange names can also contain non-heuristic sensitive identifiers.
        // Under structured DLP redaction, treat them as disallowed metadata tokens too. They can
        // contain non-heuristic secrets that a no-op redactor cannot detect, and should never be
        // sent to a future cloud embedder.
        if (Array.isArray(rawSheet?.tables)) {
          for (const t of rawSheet.tables) {
            throwIfAborted(signal);
            const name = String(t?.name ?? "");
            if (!name) continue;
            if (!nextQuery.includes(name) && !nextQuery.includes(encodeURIComponent(name))) continue;
            redactQueryToken(name);
          }
        }
        if (Array.isArray(rawSheet?.namedRanges)) {
          for (const r of rawSheet.namedRanges) {
            throwIfAborted(signal);
            const name = String(r?.name ?? "");
            if (!name) continue;
            if (!nextQuery.includes(name) && !nextQuery.includes(encodeURIComponent(name))) continue;
            redactQueryToken(name);
          }
        }

        queryForRag = nextQuery;
      }
    }

    // Index sheet into the in-memory RAG store (cached by content signature) and retrieve
    // relevant chunks. Both steps must run under the per-sheet lock so concurrent calls
    // cannot swap out the underlying store between indexing and retrieval (which could
    // otherwise leak unredacted content under DLP REDACT decisions).
    const sheetName = typeof sheetForContext?.name === "string" ? sheetForContext.name : String(sheetForContext?.name ?? "");
    const { schema, retrieved } = await this._withSheetIndexLock(sheetName, signal, async () => {
      const { schema } = await this._ensureSheetIndexedLocked(sheetForContext, {
        signal,
        maxChunkRows,
        splitRegions,
        chunkRowOverlap,
        maxChunksPerRegion,
      });
      throwIfAborted(signal);
      const retrieved = await this.ragIndex.search(queryForRag, this.sheetRagTopK, { signal });
      return { schema, retrieved };
    });
    throwIfAborted(signal);

    const shouldRedactStructuredSheetNameToken =
      Boolean(dlp) && dlpStructuredSheetDecision && dlpStructuredSheetDecision.decision !== DLP_DECISION.ALLOW;

    // Structured DLP redaction can be triggered by explicit range/cell classifications that are not
    // detectable by heuristic redactors. In those cases, treat table/namedRange names as disallowed
    // metadata tokens too (they can contain non-heuristic sensitive strings like "TopSecret").
    const shouldRedactStructuredSchemaTokens =
      Boolean(dlp) && dlpStructuredDecision?.decision === DLP_DECISION.REDACT && dlpClassificationRecords.length > 0;
    const schemaForDlp = (() => {
      if (!shouldRedactStructuredSchemaTokens) return schema;
      throwIfAborted(signal);
      const includeRestrictedContent = dlp?.includeRestrictedContent ?? false;
      const redactedSheetName = shouldRedactStructuredSheetNameToken ? "[REDACTED]" : null;
      const redactAllExplicitSchemaNames = true;

      /** @type {Set<string>} */
      const explicitTableRanges = new Set();
      if (Array.isArray(rawSheet?.tables)) {
        for (const t of rawSheet.tables) {
          throwIfAborted(signal);
          const rawRange = typeof t?.range === "string" ? t.range : "";
          if (!rawRange) continue;
          let parsed;
          try {
            parsed = parseA1Range(rawRange);
          } catch {
            continue;
          }
          // Canonicalize to match `extractSheetSchema` output.
          const canonical = rangeToA1({ ...parsed, sheetName: rawSheet.name });
          if (canonical) explicitTableRanges.add(canonical);
        }
      }

      /**
       * Redact the sheet-name component of an A1 range under structured sheet-level DLP.
       * @param {unknown} rangeA1
       */
      const redactA1SheetName = (rangeA1) => {
        const raw = String(rangeA1 ?? "");
        if (!raw || !redactedSheetName) return raw;
        try {
          const parsed = parseA1Range(raw);
          return rangeToA1({ ...parsed, sheetName: redactedSheetName });
        } catch {
          return "[REDACTED]";
        }
      };

      /**
       * Evaluate policy for an A1 range string (table/namedRange) using structured DLP selectors.
       *
       * @param {unknown} rangeA1
       */
      const recordDecisionForA1Range = (rangeA1) => {
        const raw = String(rangeA1 ?? "");
        if (!raw) return null;
        let parsed;
        try {
          parsed = parseA1Range(raw);
        } catch {
          // If we cannot parse the range, be conservative and treat it as disallowed metadata.
          return { decision: DLP_DECISION.REDACT };
        }
        const rangeRef = {
          documentId: dlpAuditDocumentId ?? dlp.documentId,
          sheetId: dlpAuditSheetId ?? dlp.sheetId ?? rawSheet.name,
          range: {
            start: { row: parsed.startRow, col: parsed.startCol },
            end: { row: parsed.endRow, col: parsed.endCol },
          },
        };
        const classification = effectiveRangeClassification(rangeRef, dlpClassificationRecords);
        return evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification,
          policy: dlp.policy,
          options: { includeRestrictedContent },
        });
      };

      const nextTables = Array.isArray(schema?.tables)
        ? schema.tables.map((t) => {
            const decision = recordDecisionForA1Range(t?.range);
            const safeRange = redactA1SheetName(t?.range);
            const isExplicit = explicitTableRanges.has(String(t?.range ?? ""));
            if (decision && decision.decision !== DLP_DECISION.ALLOW) return { ...t, name: "[REDACTED]", range: safeRange };
            if (redactAllExplicitSchemaNames && isExplicit) return { ...t, name: "[REDACTED]", range: safeRange };
            return { ...t, range: safeRange };
          })
        : schema?.tables;

      const nextNamedRanges = Array.isArray(schema?.namedRanges)
        ? schema.namedRanges.map((r) => {
            const decision = recordDecisionForA1Range(r?.range);
            const safeRange = redactA1SheetName(r?.range);
            if (decision && decision.decision !== DLP_DECISION.ALLOW) return { ...r, name: "[REDACTED]", range: safeRange };
            if (redactAllExplicitSchemaNames) return { ...r, name: "[REDACTED]", range: safeRange };
            return { ...r, range: safeRange };
          })
        : schema?.namedRanges;

      const nextDataRegions = Array.isArray(schema?.dataRegions)
        ? schema.dataRegions.map((region) => ({
            ...region,
            range: redactA1SheetName(region?.range),
          }))
        : schema?.dataRegions;

      return {
        ...(schema && typeof schema === "object" ? schema : {}),
        ...(redactedSheetName ? { name: redactedSheetName } : null),
        ...(Array.isArray(nextTables) ? { tables: nextTables } : null),
        ...(Array.isArray(nextNamedRanges) ? { namedRanges: nextNamedRanges } : null),
        ...(Array.isArray(nextDataRegions) ? { dataRegions: nextDataRegions } : null),
      };
    })();

    const retrievedForDlp = (() => {
      if (!shouldRedactStructuredSheetNameToken) return retrieved;
      return (retrieved ?? []).map((hit) => {
        if (!hit || typeof hit !== "object") return hit;
        const rawRange = /** @type {any} */ (hit).range;
        const raw = String(rawRange ?? "");
        if (!raw) return hit;
        try {
          const parsed = parseA1Range(raw);
          const safeRange = rangeToA1({ ...parsed, sheetName: "[REDACTED]" });
          return { ...hit, range: safeRange };
        } catch {
          return hit;
        }
      });
    })();

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

    const attachmentDataRaw = buildRangeAttachmentSectionText(
      { sheet: sheetForContext, attachments: params.attachments },
      {
        maxRows: 30,
        maxAttachments: 3,
        sheetNameForOutput: shouldRedactStructuredSheetNameToken ? "[REDACTED]" : undefined,
        signal,
      },
    );

    const shouldReturnRedactedStructured = Boolean(dlp) && dlpDecision?.decision === DLP_DECISION.REDACT;
    const includeRestrictedContentForStructured =
      dlp?.includeRestrictedContent ?? dlp?.include_restricted_content ?? false;
    const attachmentsForPromptUnsafe = shouldReturnRedactedStructured
      ? redactStructuredValue(params.attachments ?? [], this.redactor, {
          signal,
          includeRestrictedContent: includeRestrictedContentForStructured,
          policyAllowsRestrictedContent,
        })
      : params.attachments;
    const shouldDropAllAttachmentData =
      Boolean(dlp) && dlpStructuredDecision?.decision === DLP_DECISION.REDACT;
    const attachmentsForPromptRaw = compactAttachmentsForPrompt(attachmentsForPromptUnsafe, {
      dropAllData: shouldDropAllAttachmentData,
    });
    const attachmentsForPrompt = (() => {
      if (!Array.isArray(attachmentsForPromptRaw)) return attachmentsForPromptRaw;

      // Under structured DLP redaction, treat attachment reference strings as disallowed metadata
      // tokens too. They can contain non-heuristic secrets (e.g. "TopSecret") that a no-op redactor
      // cannot detect.
      if (shouldDropAllAttachmentData) {
        return attachmentsForPromptRaw.map((item) => {
          if (!item || typeof item !== "object" || Array.isArray(item)) return item;
          const type = /** @type {any} */ (item).type;
          const reference = /** @type {any} */ (item).reference;
          if (typeof reference !== "string") return { ...item, reference: "[REDACTED]" };

          if (type === "range") {
            let parsed;
            try {
              parsed = parseA1Range(reference);
            } catch {
              return { ...item, reference: "[REDACTED]" };
            }
            if (!parsed.sheetName) return item;
            return { ...item, reference: rangeToA1({ ...parsed, sheetName: "[REDACTED]" }) };
          }

          // Table/chart/formula/etc references are identifiers; redact them entirely.
          return { ...item, reference: "[REDACTED]" };
        });
      }

      // Otherwise, only rewrite explicit sheet-qualified range references when the sheet name itself
      // is structurally disallowed (sheet-level structured DLP).
      if (shouldRedactStructuredSheetNameToken) {
        return attachmentsForPromptRaw.map((item) => {
          if (!item || typeof item !== "object" || Array.isArray(item)) return item;
          const type = /** @type {any} */ (item).type;
          const reference = /** @type {any} */ (item).reference;
          if (type !== "range" || typeof reference !== "string") return item;
          let parsed;
          try {
            parsed = parseA1Range(reference);
          } catch {
            // Best-effort: if the reference includes the raw sheet name, strip it entirely.
            const sheetNameRaw = String(rawSheet?.name ?? "");
            if (sheetNameRaw && reference.includes(sheetNameRaw)) {
              return { ...item, reference: "[REDACTED]" };
            }
            return item;
          }

          // Only rewrite explicit sheet-qualified references that point at the current sheet.
          if (!parsed.sheetName) return item;
          if (normalizeSheetNameForComparison(parsed.sheetName) !== normalizeSheetNameForComparison(rawSheet?.name ?? "")) {
            return item;
          }

          return { ...item, reference: rangeToA1({ ...parsed, sheetName: "[REDACTED]" }) };
        });
      }

      return attachmentsForPromptRaw;
    })();
    const schemaOut = shouldReturnRedactedStructured
      ? redactStructuredValue(schemaForDlp, this.redactor, {
          signal,
          includeRestrictedContent: includeRestrictedContentForStructured,
          policyAllowsRestrictedContent,
        })
      : schemaForDlp;
    const sampledOut = shouldReturnRedactedStructured
      ? redactStructuredValue(sampled, this.redactor, {
          signal,
          includeRestrictedContent: includeRestrictedContentForStructured,
          policyAllowsRestrictedContent,
        })
      : sampled;
    const retrievedOut = shouldReturnRedactedStructured
      ? redactStructuredValue(retrievedForDlp, this.redactor, {
          signal,
          includeRestrictedContent: includeRestrictedContentForStructured,
          policyAllowsRestrictedContent,
        })
      : retrievedForDlp;
    const schemaForPrompt = compactSheetSchemaForPrompt(schemaOut, {
      maxTables: 10,
      maxRegions: 10,
      maxNamedRanges: 10,
      maxColumns: 25,
    });
    const attachmentData = shouldReturnRedactedStructured
      ? redactStructuredValue(attachmentDataRaw, this.redactor, {
          signal,
          includeRestrictedContent: includeRestrictedContentForStructured,
          policyAllowsRestrictedContent,
        })
      : attachmentDataRaw;

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
        key: "schema_summary",
        // Prefer the compact summary over raw JSON when budgets are tight.
        priority: 3.5,
        text: this.redactor(
          `Sheet schema summary:\n${summarizeSheetSchema(schemaOut, {
            maxTables: 10,
            maxRegions: 10,
            maxNamedRanges: 10,
            maxHeadersPerTable: 8,
            maxHeadersPerRegion: 8,
          })}`,
        ),
      },
      {
        key: "schema",
        priority: 3,
        text: this.redactor(`Sheet schema (schema-first):\n${stableJsonStringify(schemaForPrompt)}`),
      },
      {
        key: "attachments",
        priority: 2,
        text:
          Array.isArray(attachmentsForPrompt) && attachmentsForPrompt.length
            ? this.redactor(`User-provided attachments:\n${stableJsonStringify(attachmentsForPrompt)}`)
            : "",
      },
      {
        key: "samples",
        priority: 1,
        text: sampledOut.length
          ? this.redactor(`Sample rows:\n${sampledOut.map((r) => stableJsonStringify(r)).join("\n")}`)
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
   *     auditLogger?: { log(event: any): void },
   *     sheetNameResolver?: any,
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
    const policyAllowsRestrictedContent = Boolean(dlp?.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.allowRestrictedContent);
    const maxAllowed = dlp?.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.maxAllowed ?? null;
    const maxAllowedRank = maxAllowed === null ? null : classificationRank(maxAllowed);
    // In `evaluatePolicy()`, `includeRestrictedContent=true` triggers a special-case path for
    // Restricted classifications:
    // - policy must set `allowRestrictedContent=true`, otherwise the decision is BLOCK even if
    //   `maxAllowed` is already Restricted.
    //
    // Keep our defense-in-depth redaction gates aligned with those semantics so we never treat
    // restricted content as allowed when policy would block it.
    const restrictedAllowed =
      maxAllowedRank !== null &&
      (includeRestrictedContent ? policyAllowsRestrictedContent : maxAllowedRank >= RESTRICTED_CLASSIFICATION_RANK);
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
     * Redact a short metadata token (sheet name / title) under DLP redaction.
     *
     * This is a defense-in-depth helper so DLP flows remain safe even when the configured
     * `ContextManager.redactor` is a no-op.
     *
     * @param {unknown} value
     */
    const redactChunkToken = (value) => {
      // Avoid calling `.toString()` on arbitrary objects (3p vector stores can persist
      // untrusted metadata tokens). Under DLP enforcement, treat non-primitive tokens as
      // prompt-unsafe so non-heuristic secrets cannot leak even if the configured redactor
      // is a no-op.
      const raw = (() => {
        if (!dlp) return String(value ?? "");
        if (value === null || value === undefined) return "";
        if (typeof value === "string") return value;
        if (typeof value === "number" || typeof value === "boolean" || typeof value === "bigint") return String(value);
        return "[REDACTED]";
      })();
      if (!dlp) return raw;
      const redacted = this.redactor(raw);
      if (!restrictedAllowed && classifyTextForDlp(redacted).level === "sensitive") return "[REDACTED]";
      return redacted;
    };

    /** @type {Map<string, boolean>} */
    const sheetNameDisallowedCache = new Map();
    /**
     * Determine whether a sheet-name token should be treated as disallowed metadata under
     * structured DLP.
     *
     * Even if an individual chunk/range is allowed, the sheet name itself is a user-controlled
     * identifier that can contain non-heuristic sensitive strings (e.g. "TopSecret") that a
     * no-op redactor cannot detect. When any structured selector on the sheet would require
     * redaction, conservatively redact the sheet name everywhere it appears in prompt context
     * and structured outputs.
     *
     * Best-effort: if we cannot build the structured selector index, we fall back to heuristic
     * behavior (no extra sheet-name redaction).
     *
     * @param {unknown} sheetName
     */
    const sheetNameDisallowed = (sheetName) => {
      const raw = typeof sheetName === "string" ? sheetName.trim() : "";
      if (!dlp || !raw) return false;
      const sheetId = resolveDlpSheetId(raw);
      if (!sheetId) return false;
      const cached = sheetNameDisallowedCache.get(sheetId);
      if (cached !== undefined) return cached;

      const index = getDlpDocumentIndex();
      if (!index) {
        sheetNameDisallowedCache.set(sheetId, false);
        return false;
      }

      // Base classification applied to the entire sheet (document + sheet selectors).
      let baseClassification = maxClassification({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }, index.docClassificationMax);
      const sheetMax = index.sheetClassificationMaxBySheetId.get(sheetId);
      if (sheetMax) baseClassification = maxClassification(baseClassification, sheetMax);

      // If the sheet itself is disallowed due to doc/sheet scope, redact the name.
      const baseDecision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: baseClassification,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });
      if (baseDecision.decision !== DLP_DECISION.ALLOW) {
        sheetNameDisallowedCache.set(sheetId, true);
        return true;
      }

      // Otherwise, if any structured selector on the sheet would require redaction,
      // treat the sheet name as disallowed metadata.
      const colMap = index.columnClassificationBySheetId.get(sheetId);
      if (colMap) {
        for (const colClassification of colMap.values()) {
          throwIfAborted(signal);
          const decision = evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification: maxClassification(baseClassification, colClassification),
            policy: dlp.policy,
            options: { includeRestrictedContent },
          });
          if (decision.decision !== DLP_DECISION.ALLOW) {
            sheetNameDisallowedCache.set(sheetId, true);
            return true;
          }
        }
      }

      const rangeRecords = index.rangeRecordsBySheetId.get(sheetId) ?? [];
      for (const rec of rangeRecords) {
        throwIfAborted(signal);
        const decision = evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification: maxClassification(baseClassification, rec.classification),
          policy: dlp.policy,
          options: { includeRestrictedContent },
        });
        if (decision.decision !== DLP_DECISION.ALLOW) {
          sheetNameDisallowedCache.set(sheetId, true);
          return true;
        }
      }

      const cellMap = index.cellClassificationBySheetId.get(sheetId);
      if (cellMap) {
        for (const cellClassification of cellMap.values()) {
          throwIfAborted(signal);
          const decision = evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification: maxClassification(baseClassification, cellClassification),
            policy: dlp.policy,
            options: { includeRestrictedContent },
          });
          if (decision.decision !== DLP_DECISION.ALLOW) {
            sheetNameDisallowedCache.set(sheetId, true);
            return true;
          }
        }
      }

      sheetNameDisallowedCache.set(sheetId, false);
      return false;
    };

    /**
     * Structured DLP decision helper for a workbook rect (table/namedRange/chunk).
     *
     * @param {unknown} sheetName
     * @param {any} rect
     */
    const rectDisallowed = (sheetName, rect) => {
      const rawSheet = typeof sheetName === "string" ? sheetName.trim() : "";
      if (!dlp || !rawSheet) return false;
      const sheetId = resolveDlpSheetId(rawSheet);
      if (!sheetId) return false;
      const range = rectToRange(rect);
      if (!range) return false;
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
      return recordDecision.decision !== DLP_DECISION.ALLOW;
    };

    /**
     * @param {any} rect
     */
    function rectToA1WithoutSheet(rect) {
      if (!rect || typeof rect !== "object") return "";
      const { r0, c0, r1, c1 } = rect;
      if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return "";
      if (r1 < r0 || c1 < c0) return "";
      try {
        return rangeToA1({ startRow: r0, startCol: c0, endRow: r1, endCol: c1 });
      } catch {
        return "";
      }
    }

    /**
     * Safe first line for persisted redacted chunks (used to avoid leaking sheet/title metadata
     * when a chunk must be blocked/redacted).
     *
     * Note: when DLP redaction is required due to *structured* classification (document/sheet/range),
     * we must treat metadata tokens (sheet name, title) as disallowed too. Those tokens can contain
     * user-provided identifiers that are not detectable by heuristic redaction (e.g. "TopSecret").
     *
     * @param {any} metadata
     * @param {{ redactTokens?: boolean }} [options]
     */
    const safeChunkFirstLineFromMetadata = (metadata, options = {}) => {
      const meta = metadata && typeof metadata === "object" ? metadata : {};
      const kind = normalizeWorkbookChunkKind(meta.kind ?? "chunk").toUpperCase();
      const shouldRedactTokens = options.redactTokens === true;
      const optionToken = (value) => {
        if (value === null || value === undefined) return "";
        if (typeof value === "string") return value;
        // Avoid leaking via custom `toString()` implementations on persisted objects.
        if (!dlp && typeof value !== "object") return String(value);
        if (typeof value === "number" || typeof value === "boolean" || typeof value === "bigint") return String(value);
        return "[REDACTED]";
      };
      const title =
        shouldRedactTokens
          ? "[REDACTED]"
          : Object.prototype.hasOwnProperty.call(options, "titleForOutput")
            ? optionToken(options.titleForOutput)
            : redactChunkToken(meta.title ?? "");
      const sheetName =
        shouldRedactTokens
          ? "[REDACTED]"
          : Object.prototype.hasOwnProperty.call(options, "sheetNameForOutput")
            ? optionToken(options.sheetNameForOutput)
            : redactChunkToken(meta.sheetName ?? "");
      const rectA1 = rectToA1WithoutSheet(meta.rect);
      const sheetPart = `sheet="${sheetName}"`;
      const rangePart = rectA1 ? `, range="${rectA1}"` : "";
      return `${kind}: ${title} (${sheetPart}${rangePart})`;
    };

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
                const heuristic = classifyTextForDlp(rawText);
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

                const shouldRedactStructuredMetadataTokens =
                  decision.decision === DLP_DECISION.BLOCK || recordDecision.decision !== DLP_DECISION.ALLOW;

                let safeText = rawText;
                if (decision.decision !== DLP_DECISION.ALLOW) {
                  if (decision.decision === DLP_DECISION.BLOCK) {
                    // If the policy blocks cloud AI processing for this chunk, do not send any
                    // workbook content to the embedder. Persist only a minimal placeholder so
                    // the vector store cannot contain raw restricted data.
                    safeText = this.redactor(
                      `${safeChunkFirstLineFromMetadata(record.metadata, { redactTokens: true })}\n[REDACTED]`,
                    );
                  } else {
                    // If DLP redaction is required due to explicit document/sheet/range classification,
                    // redact the entire content; pattern-based redaction isn't sufficient in that case.
                    if (recordDecision.decision !== DLP_DECISION.ALLOW) {
                      safeText = this.redactor(
                        `${safeChunkFirstLineFromMetadata(record.metadata, { redactTokens: true })}\n[REDACTED]`,
                      );
                    } else {
                      safeText = this.redactor(rawText);
                    }
                    if (!restrictedAllowed && classifyTextForDlp(safeText).level === "sensitive") {
                      safeText = this.redactor(
                        `${safeChunkFirstLineFromMetadata(record.metadata, { redactTokens: true })}\n[REDACTED]`,
                      );
                    }
                  }
                }

                // Structured DLP classifications can require redaction for a sheet even when this
                // particular chunk range is allowed (e.g. a Restricted range elsewhere on the same
                // sheet). In those cases, treat user-controlled metadata tokens (sheet name, and
                // table/namedRange titles) as disallowed for embedding/persistence so a no-op
                // redactor cannot leak non-heuristic secrets like "TopSecret" to a cloud embedder.
                const sheetNameTokenDisallowed = sheetNameDisallowed(sheetName);
                const rawKind = record.metadata?.kind;
                const kind = normalizeWorkbookChunkKind(rawKind ?? "");
                const kindUnknown = isWorkbookChunkKindUnknown(rawKind);
                const titleTokenDisallowed =
                  sheetNameTokenDisallowed && (kind === "table" || kind === "namedRange" || kindUnknown);
                if (sheetNameTokenDisallowed && recordDecision.decision === DLP_DECISION.ALLOW) {
                  const safeFirstLine = safeChunkFirstLineFromMetadata(record.metadata, {
                    sheetNameForOutput: "[REDACTED]",
                    ...(titleTokenDisallowed ? { titleForOutput: "[REDACTED]" } : {}),
                  });
                  const nl = safeText.indexOf("\n");
                  const rest = nl === -1 ? "" : safeText.slice(nl);
                  safeText = `${safeFirstLine}${rest}`;
                }

                // Defense-in-depth: if the configured redactor is a no-op (or incomplete),
                // ensure heuristic sensitive patterns never slip through under DLP enforcement.
                if (!restrictedAllowed && classifyTextForDlp(safeText).level === "sensitive") {
                  safeText = "[REDACTED]";
                }

                return {
                  text: safeText,
                  metadata: {
                    ...(record.metadata ?? {}),
                    ...(sheetId ? { dlpSheetId: sheetId } : null),
                    ...(shouldRedactStructuredMetadataTokens
                      ? {
                          // Strip potentially sensitive, user-controlled metadata tokens under structured DLP.
                          sheetName: "[REDACTED]",
                          title: "[REDACTED]",
                        }
                      : null),
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
    const queryHeuristic = dlp ? classifyTextForDlp(params.query) : null;
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
    let queryForEmbedding = params.query;
    if (dlp && queryDecision && queryDecision.decision !== DLP_DECISION.ALLOW) {
      queryForEmbedding = this.redactor(params.query);
      if (!restrictedAllowed && classifyTextForDlp(queryForEmbedding).level === "sensitive") {
        queryForEmbedding = "[REDACTED]";
      }
    }

    // Structured DLP can require redaction for non-heuristic workbook metadata tokens (e.g. sheet names,
    // table names) that a no-op redactor cannot detect. If those disallowed identifiers appear in the
    // user's query, strip them before embedding to keep future cloud embedders safe.
    if (dlp && classificationRecords.length && typeof queryForEmbedding === "string") {
      let nextQuery = queryForEmbedding;

      /**
       * @param {string} haystack
       * @param {string} needle
       * @param {string} replacement
       */
      const replaceAll = (haystack, needle, replacement) => {
        if (!needle) return haystack;
        if (!haystack.includes(needle)) return haystack;
        return haystack.split(needle).join(replacement);
      };

      /**
       * Replace a disallowed token and common encodings/quoting forms.
       * @param {string} token
       */
      const redactQueryToken = (token) => {
        const raw = String(token ?? "");
        if (!raw) return;
        nextQuery = replaceAll(nextQuery, raw, "[REDACTED]");
        const encoded = encodeURIComponent(raw);
        if (encoded && encoded !== raw) nextQuery = replaceAll(nextQuery, encoded, "[REDACTED]");
        // Excel-style quoted sheet names.
        const quoted = `'${raw.replace(/'/g, "''")}'`;
        if (quoted !== raw) nextQuery = replaceAll(nextQuery, quoted, "'[REDACTED]'");
      };

      // Workbook id itself can be user-controlled metadata. If any structured selector would require
      // redaction for cloud AI processing, treat the workbook id as disallowed too.
      let structuredOverallClassificationForQuery = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
      for (const record of classificationRecords) {
        throwIfAborted(signal);
        const classification = record?.classification;
        if (classification && typeof classification === "object") {
          structuredOverallClassificationForQuery = maxClassification(structuredOverallClassificationForQuery, classification);
        }
      }
      const structuredOverallDecisionForQuery = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification: structuredOverallClassificationForQuery,
        policy: dlp.policy,
        options: { includeRestrictedContent },
      });
      if (structuredOverallDecisionForQuery.decision !== DLP_DECISION.ALLOW) {
        redactQueryToken(String(params.workbook.id ?? ""));
      }

      // Sheet names.
      for (const s of params.workbook.sheets ?? []) {
        throwIfAborted(signal);
        const sheetName = String(s?.name ?? "");
        if (!sheetName) continue;
        if (
          !nextQuery.includes(sheetName) &&
          !nextQuery.includes(encodeURIComponent(sheetName)) &&
          !nextQuery.includes(`'${sheetName.replace(/'/g, "''")}'`)
        ) {
          continue;
        }
        if (sheetNameDisallowed(sheetName)) {
          redactQueryToken(sheetName);
        }
      }

      // Table/namedRange names.
      for (const t of params.workbook.tables ?? []) {
        throwIfAborted(signal);
        const name = String(t?.name ?? "");
        if (!name) continue;
        if (!nextQuery.includes(name) && !nextQuery.includes(encodeURIComponent(name))) continue;
        const sheetName = String(t?.sheetName ?? "");
        if (sheetNameDisallowed(sheetName) || rectDisallowed(sheetName, t?.rect)) {
          redactQueryToken(name);
        }
      }
      for (const r of params.workbook.namedRanges ?? []) {
        throwIfAborted(signal);
        const name = String(r?.name ?? "");
        if (!name) continue;
        if (!nextQuery.includes(name) && !nextQuery.includes(encodeURIComponent(name))) continue;
        const sheetName = String(r?.sheetName ?? "");
        if (sheetNameDisallowed(sheetName) || rectDisallowed(sheetName, r?.rect)) {
          redactQueryToken(name);
        }
      }

      queryForEmbedding = nextQuery;
    }
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
    /** @type {{level: string, labels: string[]} } */
    let structuredOverallClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
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
          structuredOverallClassification = maxClassification(structuredOverallClassification, classification);
          overallClassification = maxClassification(overallClassification, classification);
        }
      }
    }
    const structuredOverallDecision = dlp
      ? evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification: structuredOverallClassification,
          policy: dlp.policy,
          options: { includeRestrictedContent },
        })
      : null;
    // Attachments can contain user-provided context (e.g. chart annotations) that should
    // influence the overall AI cloud processing decision when DLP is enabled.
    if (dlp && Array.isArray(params.attachments) && params.attachments.length) {
      const attachmentsHeuristic = classifyStructuredForDlp(params.attachments, { signal });
      const attachmentsClassification = heuristicToPolicyClassification(attachmentsHeuristic);
      overallClassification = maxClassification(overallClassification, attachmentsClassification);
    }

    // When structured DLP requires redaction anywhere in the workbook, treat the workbook id as
    // disallowed metadata too. Workbook ids can be user-controlled and may contain non-heuristic
    // sensitive strings that a no-op redactor cannot detect.
    const workbookIdTokenDisallowed =
      Boolean(dlp) && structuredOverallDecision && structuredOverallDecision.decision !== DLP_DECISION.ALLOW;
    /** @type {ReturnType<typeof evaluatePolicy> | null} */
    let overallDecision = null;
    let redactedChunkCount = 0;

    /** @type {any[]} */
    const chunkAudits = [];

    /**
     * Merge two heuristic classifications.
     *
     * We intentionally combine persisted (index-time) heuristics with runtime heuristics so:
     * - we still remember that the *underlying* chunk content was sensitive even if the stored
     *   `metadata.text` has since been replaced with a redacted placeholder
     * - we don't miss newly-detected patterns (e.g. percent-encoded tokens) if an older index
     *   stored an incomplete heuristic result
     *
     * @param {any} a
     * @param {any} b
     */
    const mergeHeuristics = (a, b) => {
      const aa = a && typeof a === "object" ? a : { level: "public", findings: [] };
      const bb = b && typeof b === "object" ? b : { level: "public", findings: [] };
      // Findings should always be a list of stable detector identifiers (strings). Vector stores
      // can be untrusted; avoid coercing arbitrary objects via `String()`/template literals
      // because custom `toString()` implementations can leak non-heuristic secrets even when
      // ContextManager.redactor is a no-op.
      const toSafeFinding = (f) => {
        if (typeof f === "string") return f;
        // Allow primitives (shouldn't happen today, but keep output stable).
        if (typeof f === "number" || typeof f === "boolean" || typeof f === "bigint") return String(f);
        return null;
      };
      const findings = new Set(
        [...(Array.isArray(aa.findings) ? aa.findings : []), ...(Array.isArray(bb.findings) ? bb.findings : [])]
          .map(toSafeFinding)
          .filter((v) => typeof v === "string" && v !== ""),
      );
      const level = aa.level === "sensitive" || bb.level === "sensitive" ? "sensitive" : "public";
      return { level, findings: [...findings] };
    };

    // Evaluate policy for all retrieved chunks before returning any prompt context.
    for (const [idx, hit] of hits.entries()) {
      throwIfAborted(signal);
      const meta = hit.metadata ?? {};
      const rawKind = meta.kind;
      const kind = dlp ? normalizeWorkbookChunkKind(rawKind ?? "chunk") : typeof rawKind === "string" ? rawKind : "chunk";
      const headerSheetName = typeof meta.sheetName === "string" ? meta.sheetName : "";
      const headerTitle = typeof meta.title === "string" ? meta.title : meta.title == null ? hit.id : "[REDACTED]";
      const header = `#${idx + 1} score=${hit.score.toFixed(3)} kind=${kind} sheet=${headerSheetName} title="${headerTitle}"`;
      const text = typeof meta.text === "string" ? meta.text : "";
      const raw = `${header}\n${text}`;

      const storedHeuristic = heuristicByChunkId.get(hit.id) ?? meta.dlpHeuristic;
      const heuristic = mergeHeuristics(storedHeuristic, classifyTextForDlp(raw));
      const heuristicClassification = heuristicToPolicyClassification(heuristic);

      // If the caller provided structured cell/range classifications, fold those in using the
      // chunk's sheet + rect metadata.
      let recordClassification = { level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] };
      if (dlp) {
        const range = rectToRange(meta.rect);
        const sheetName = meta.sheetName;
        const storedSheetId = typeof meta.dlpSheetId === "string" ? meta.dlpSheetId.trim() : "";
        const sheetId = storedSheetId || (typeof sheetName === "string" ? resolveDlpSheetId(sheetName) : "");
        if (range && sheetId) {
          const index = getDlpDocumentIndex();
          recordClassification = index
            ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range }, signal)
            : effectiveRangeClassification({ documentId: dlp.documentId, sheetId, range }, classificationRecords);
        } else if (workbookIdTokenDisallowed && classificationRecords.length) {
          // If structured DLP requires redaction anywhere in the workbook but this chunk
          // does not include enough metadata to compute its selector (missing rect/sheet),
          // be conservative and assume the worst-case structured classification so we
          // never leak non-heuristic restricted content.
          recordClassification = maxClassification(structuredOverallClassification, {
            level: structuredOverallClassification.level,
            labels: ["structured:missingChunkMetadata"],
          });
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
        sheetName: headerSheetName,
        title: headerTitle,
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
      const rawKind = meta.kind;
      const kind = dlp ? normalizeWorkbookChunkKind(rawKind ?? "chunk") : (meta.kind ?? "chunk");
      const kindUnknown = Boolean(dlp) && isWorkbookChunkKindUnknown(rawKind);
      const audit = chunkAudits[idx];
      const decision = audit?.decision ?? null;
      const recordDecision = audit?.recordDecision ?? null;
      const rangeDisallowed = Boolean(dlp) && recordDecision && recordDecision.decision !== DLP_DECISION.ALLOW;
      const rawSheetName =
        typeof meta.sheetName === "string" ? meta.sheetName : dlp ? "" : String(meta.sheetName ?? "");
      const sheetNameInvalid = Boolean(dlp) && meta.sheetName != null && typeof meta.sheetName !== "string";
      const sheetNameTokenDisallowed = Boolean(dlp) && (sheetNameInvalid || sheetNameDisallowed(rawSheetName));
      const titleInvalid = Boolean(dlp) && meta.title != null && typeof meta.title !== "string";
      const titleTokenDisallowed =
        titleInvalid ||
        (sheetNameTokenDisallowed && (kind === "table" || kind === "namedRange")) ||
        (workbookIdTokenDisallowed && kindUnknown);
      const shouldRedactSheetNameToken = Boolean(dlp) && (rangeDisallowed || sheetNameTokenDisallowed);
      const shouldRedactTitleToken = Boolean(dlp) && (rangeDisallowed || titleTokenDisallowed);
      const shouldRedactChunkId = shouldRedactSheetNameToken || shouldRedactTitleToken || workbookIdTokenDisallowed;
      // Legacy / third-party vector stores may omit `metadata.title`. Falling back to `hit.id`
      // can leak disallowed metadata tokens (workbook id, sheet name) under structured DLP,
      // because chunk ids embed those identifiers and heuristic redaction can't detect
      // non-heuristic secrets like "TopSecret".
      const title = dlp
        ? typeof meta.title === "string"
          ? meta.title
          : meta.title == null
            ? shouldRedactChunkId
              ? "[REDACTED]"
              : hit.id
            : "[REDACTED]"
        : meta.title ?? hit.id;

      const safeSheetName = shouldRedactSheetNameToken
        ? "[REDACTED]"
        : dlp
          ? redactChunkToken(rawSheetName)
          : rawSheetName;
      const safeTitle = shouldRedactTitleToken
        ? "[REDACTED]"
        : dlp
          ? redactChunkToken(title)
          : String(title ?? "");
      const header = `#${idx + 1} score=${hit.score.toFixed(3)} kind=${kind} sheet=${safeSheetName} title="${safeTitle}"`;
      const rawText = typeof meta.text === "string" ? meta.text : "";
      const text =
        (shouldRedactSheetNameToken || shouldRedactTitleToken) && rawText.includes("\n")
          ? (() => {
              // The stored chunk text (from `chunkToText`) repeats user-controlled metadata
              // tokens (title + sheet name) in its first line. If those tokens are disallowed
              // under structured DLP, rewrite the first line deterministically so non-heuristic
              // sensitive strings cannot leak even with a no-op redactor.
              const idxNl = rawText.indexOf("\n");
              const rest = idxNl === -1 ? "" : rawText.slice(idxNl);
              const safeFirstLine = safeChunkFirstLineFromMetadata(meta, {
                sheetNameForOutput: safeSheetName,
                titleForOutput: safeTitle,
              });
              return `${safeFirstLine}${rest}`;
            })()
          : rawText;
      const raw = `${header}\n${text}`;

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
      if (dlp && !restrictedAllowed && classifyTextForDlp(outText).level === "sensitive") {
        outText = this.redactor(`${header}\n[REDACTED]`);
        redacted = true;
      }

      if (dlp && !restrictedAllowed && classifyTextForDlp(outText).level === "sensitive") {
        outText = "[REDACTED]";
        redacted = true;
      }

      if (redacted) redactedChunkCount += 1;

      // Do not return the raw chunk text stored in vector-store metadata. The prompt-safe
      // `text` field is already provided separately, and returning unredacted metadata
      // creates an easy footgun for callers that might serialize metadata into cloud LLM
      // prompts.
      const { text: _metaText, dlpSheetId: _metaSheetId, ...safeMeta } = meta;
      /** @type {any} */
      const safeMetaOut = { ...safeMeta };
      if (dlp && Object.prototype.hasOwnProperty.call(safeMetaOut, "kind")) safeMetaOut.kind = kind;
      if (shouldRedactSheetNameToken) safeMetaOut.sheetName = "[REDACTED]";
      if (shouldRedactTitleToken) safeMetaOut.title = "[REDACTED]";
      if (workbookIdTokenDisallowed) safeMetaOut.workbookId = "[REDACTED]";
      // Under DLP redaction, strip any extra fields from rect objects to prevent non-heuristic
      // strings from leaking via `metadata.rect` (e.g. if a host attaches `{..., note:"TopSecret" }`).
      if (dlp && overallDecision?.decision === DLP_DECISION.REDACT && Object.prototype.hasOwnProperty.call(safeMetaOut, "rect")) {
        const normalizedRect = normalizeWorkbookRect(safeMetaOut.rect);
        if (normalizedRect) safeMetaOut.rect = normalizedRect;
        else delete safeMetaOut.rect;
      }
      // Under structured DLP redaction, treat any unknown metadata keys as prompt-unsafe (they can
      // contain non-heuristic identifiers). Drop them deterministically so a no-op redactor cannot leak.
      const finalMetaOut = dlp && workbookIdTokenDisallowed ? filterWorkbookChunkMetadataForOutput(safeMetaOut) : safeMetaOut;

      return {
        id: shouldRedactChunkId ? `redacted:${idx + 1}` : hit.id,
        score: hit.score,
        metadata: finalMetaOut,
        text: outText,
        dlp: mergeHeuristics(heuristicByChunkId.get(hit.id) ?? meta.dlpHeuristic, classifyTextForDlp(outText)),
      };
    });

    let promptContext = "";
    if (includePromptContext) {
      const shouldRedactPromptStruct = Boolean(dlp) && overallDecision?.decision === DLP_DECISION.REDACT;
      const attachmentsForPromptUnsafe = shouldRedactPromptStruct
        ? redactStructuredValue(params.attachments ?? [], this.redactor, {
            signal,
            includeRestrictedContent,
            policyAllowsRestrictedContent,
          })
        : params.attachments;
      const shouldDropAllAttachmentData =
        Boolean(dlp) && structuredOverallDecision?.decision === DLP_DECISION.REDACT;
      const attachmentsForPromptBase = compactAttachmentsForPrompt(attachmentsForPromptUnsafe, {
        dropAllData: shouldDropAllAttachmentData,
      });
      const attachmentsForPrompt = (() => {
        // When DLP redaction is required due to structured selectors (document/sheet/range/cell),
        // treat sheet-name tokens embedded in range attachment references as disallowed metadata too.
        // Those tokens can include non-heuristic sensitive strings (e.g. "TopSecret") that a no-op
        // redactor cannot detect.
        if (!dlp || !shouldDropAllAttachmentData) return attachmentsForPromptBase;
        if (!Array.isArray(attachmentsForPromptBase)) return attachmentsForPromptBase;

        // Range-level structured decisions should also be able to trigger metadata redaction for the
        // attachment reference. The reference (including sheet name) is an identifier for the disallowed
        // content, and may contain non-heuristic sensitive strings.
        const index = getDlpDocumentIndex();

        return attachmentsForPromptBase.map((item) => {
          if (!item || typeof item !== "object" || Array.isArray(item)) return item;
          const type = item.type;
          const reference = item.reference;
          if (typeof reference !== "string") return { ...item, reference: "[REDACTED]" };

          if (type === "range") {
            let parsed;
            try {
              parsed = parseA1Range(reference);
            } catch {
              return { ...item, reference: "[REDACTED]" };
            }
            if (!parsed.sheetName) return item;
            if (sheetNameDisallowed(parsed.sheetName)) {
              return { ...item, reference: rangeToA1({ ...parsed, sheetName: "[REDACTED]" }) };
            }
            const sheetId = resolveDlpSheetId(parsed.sheetName);
            const range = {
              start: { row: parsed.startRow, col: parsed.startCol },
              end: { row: parsed.endRow, col: parsed.endCol },
            };
            const recordClassification = index
              ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range }, signal)
              : effectiveRangeClassification({ documentId: dlp.documentId, sheetId, range }, classificationRecords);
            const recordDecision = evaluatePolicy({
              action: DLP_ACTION.AI_CLOUD_PROCESSING,
              classification: recordClassification,
              policy: dlp.policy,
              options: { includeRestrictedContent },
            });
            if (recordDecision.decision === DLP_DECISION.ALLOW) return item;
            return { ...item, reference: rangeToA1({ ...parsed, sheetName: "[REDACTED]" }) };
          }

          if (type === "table") {
            const target = reference;
            const tables = Array.isArray(params.workbook?.tables) ? params.workbook.tables : [];
            let matched = false;
            for (const t of tables) {
              throwIfAborted(signal);
              if (!t || typeof t !== "object") continue;
              if (t.name !== target) continue;
              matched = true;
              const sheetName = String(t.sheetName ?? "");
              if (sheetNameDisallowed(sheetName)) {
                return { ...item, reference: "[REDACTED]" };
              }
              const rect = t.rect;
              const range = rectToRange(rect);
              const sheetId = sheetName ? resolveDlpSheetId(sheetName) : "";
              if (!range || !sheetId) {
                // If we cannot evaluate the structured selector, be conservative and redact.
                return { ...item, reference: "[REDACTED]" };
              }
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
                return { ...item, reference: "[REDACTED]" };
              }
              break;
            }
            if (!matched) return { ...item, reference: "[REDACTED]" };
            return item;
          }

          // Other attachment types (chart/formula/etc) do not have a structured selector model today.
          // Under structured DLP redaction, treat their `reference` strings as disallowed metadata tokens.
          return { ...item, reference: "[REDACTED]" };
        });
      })();

      const schemaRestrictedDecision = dlp
        ? evaluatePolicy({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            classification: { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
            policy: dlp.policy,
            options: { includeRestrictedContent },
          })
        : null;
      const schemaRestrictedBlocks = schemaRestrictedDecision?.decision === DLP_DECISION.BLOCK;

      /**
       * Redact a schema token (table name / header / range title) when DLP requires it.
       *
       * This mirrors ContextManager's defense-in-depth behavior elsewhere: even when a caller
       * supplies a no-op redactor, heuristic-sensitive strings should not leak under DLP REDACT.
       *
       * @param {unknown} value
       * @param {string} [blockedReason]
       */
      const redactSchemaToken = (value, blockedReason = "workbook_schema") => {
        // Avoid calling `.toString()` on arbitrary objects (vector stores can persist
        // untrusted metadata). Under DLP enforcement, treat non-primitive schema tokens as
        // prompt-unsafe so non-heuristic secrets cannot leak even if the configured redactor
        // is a no-op.
        const raw = (() => {
          if (!dlp) return String(value ?? "");
          if (value === null || value === undefined) return "";
          if (typeof value === "string") return value;
          if (typeof value === "number" || typeof value === "boolean" || typeof value === "bigint") return String(value);
          return "[REDACTED]";
        })();
        if (!dlp) return raw;
        throwIfAborted(signal);

        const heuristic = classifyTextForDlp(raw);
        if (heuristic.level !== "sensitive") return raw;
        if (restrictedAllowed) return raw;

        if (schemaRestrictedBlocks) {
          const decision = schemaRestrictedDecision;
          if (decision) {
            dlp.auditLogger?.log({
              type: "ai.workbook_context",
              documentId: dlp.documentId,
              workbookId: params.workbook.id,
              action: DLP_ACTION.AI_CLOUD_PROCESSING,
              decision,
              classification: decision.classification ?? { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
              redactedChunkCount,
              blockedChunkId: null,
              blockedReason,
              chunks: chunkAudits,
            });
            throw new DlpViolationError(decision);
          }
          throw new DlpViolationError({
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            decision: DLP_DECISION.BLOCK,
            reasonCode: "dlp.blockedByPolicy",
            classification: { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: [] },
            maxAllowed: null,
          });
        }

        const redacted = this.redactor(raw);
        if (classifyTextForDlp(redacted).level === "sensitive") return "[REDACTED]";
        return redacted;
      };

      /**
       * Redact the sheet-name component of an A1 range string under DLP redaction.
       *
       * `extractWorkbookSchema()` produces absolute A1 ranges like `Sheet1!A1:B10`. If the sheet
       * name itself is heuristic-sensitive (e.g. contains an email), a no-op redactor would
       * otherwise leak it into the prompt.
       *
       * @param {unknown} a1
       * @param {string} [blockedReason]
       */
      const redactRangeA1 = (a1, blockedReason = "workbook_schema") => {
        const raw = typeof a1 === "string" ? a1 : dlp ? "" : String(a1 ?? "");
        if (!dlp) return raw;
        if (!raw) return raw;
        try {
          const parsed = parseA1Range(raw);
          const safeSheet = parsed.sheetName
            ? sheetNameDisallowed(parsed.sheetName)
              ? "[REDACTED]"
              : redactSchemaToken(parsed.sheetName, blockedReason)
            : "";
          return rangeToA1({ ...parsed, sheetName: safeSheet || undefined });
        } catch {
          const redacted = this.redactor(raw);
          if (!restrictedAllowed && classifyTextForDlp(redacted).level === "sensitive") return "[REDACTED]";
          return redacted;
        }
      };

      /**
       * Deterministically format an A1 range for a rect while forcing a redacted sheet name.
       * @param {any} rect
       */
      const redactedSheetRangeA1ForRect = (rect) => {
        if (!rect || typeof rect !== "object") return "";
        const { r0, c0, r1, c1 } = rect;
        if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return "";
        if (r1 < r0 || c1 < c0) return "";
        try {
          return rangeToA1({ sheetName: "[REDACTED]", startRow: r0, startCol: c0, endRow: r1, endCol: c1 });
        } catch {
          return "";
        }
      };

      /**
       * Redact a workbook RAG `COLUMNS:` detail string while preserving inferred types.
       * @param {string} columnsLine
       */
      const redactColumnsLine = (columnsLine) => {
        const raw = typeof columnsLine === "string" ? columnsLine : "";
        if (!raw) return "";
        const parts = raw.split(/\s*\|\s*/g);
        const out = [];
        for (const part of parts) {
          const seg = String(part ?? "");
          if (!seg) continue;
          const m = seg.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
          if (m) {
            const safeHeader = redactSchemaToken(m[1]);
            out.push(`${safeHeader} (${m[2]})`);
          } else {
            out.push(redactSchemaToken(seg));
          }
        }
        return out.join(" | ");
      };
      const schemaLines = [];
      const maxTables = 25;
      const maxColumns = 25;
      const maxNamedRanges = 25;
      const rawTables = Array.isArray(params.workbook?.tables) ? params.workbook.tables : [];
      const rawNamedRanges = Array.isArray(params.workbook?.namedRanges) ? params.workbook.namedRanges : [];
      const tablesForSchema = rawTables
        .slice()
        .sort(
          (a, b) =>
            String(a?.sheetName ?? "").localeCompare(String(b?.sheetName ?? "")) ||
            String(a?.name ?? "").localeCompare(String(b?.name ?? "")),
        )
        .slice(0, maxTables);
      const namedRangesForSchema = rawNamedRanges
        .slice()
        .sort(
          (a, b) =>
            String(a?.sheetName ?? "").localeCompare(String(b?.sheetName ?? "")) ||
            String(a?.name ?? "").localeCompare(String(b?.name ?? "")),
        )
        .slice(0, maxNamedRanges);
      const schema = extractWorkbookSchema(
        {
          id: params.workbook.id,
          sheets: params.workbook?.sheets ?? [],
          tables: tablesForSchema,
          namedRanges: namedRangesForSchema,
        },
        { maxAnalyzeRows: 50, maxAnalyzeCols: maxColumns, signal },
      );

      // Structured DLP classifications: if a table's range is disallowed due to explicit
      // selectors (document/sheet/range), omit column details entirely (match chunk redaction).
      for (let i = 0; i < schema.tables.length && i < maxTables; i++) {
        throwIfAborted(signal);
        const table = schema.tables[i];
        const name = table?.name ?? "";
        const sheetName = table?.sheetName ?? "";
        const safeName = sheetNameDisallowed(sheetName) ? "[REDACTED]" : redactSchemaToken(name);
        const rect = table?.rect;
        const rangeA1 = redactRangeA1(table?.rangeA1 ?? "");
        if (!name || !sheetName || !rect || typeof rect !== "object") continue;

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
              const redactedRange = redactedSheetRangeA1ForRect(rect) || "[REDACTED]";
              schemaLines.push(`- Table [REDACTED] (range="${redactedRange}"): [REDACTED]`);
              continue;
            }
          }
        }

        const headers = Array.isArray(table.headers) ? table.headers : [];
        const types = Array.isArray(table.inferredColumnTypes) ? table.inferredColumnTypes : [];
        const colCount = Math.max(headers.length, types.length, table.columnCount ?? 0);
        const boundedColCount = Math.min(colCount, maxColumns);
        const cols = [];

        for (let c = 0; c < boundedColCount; c++) {
          throwIfAborted(signal);
          const header = headers[c] ?? `Column${c + 1}`;
          const safeHeader = redactSchemaToken(header);
          const type = types[c] ?? "mixed";
          cols.push(`${safeHeader} (${type})`);
        }

        const colSuffix = boundedColCount < colCount ? " | …" : "";
        schemaLines.push(`- Table ${safeName} (range="${rangeA1}"): ${cols.join(" | ")}${colSuffix}`);
      }

      // Named ranges can be helpful anchors even without cell samples.
      for (let i = 0; i < schema.namedRanges.length && i < maxNamedRanges; i++) {
        throwIfAborted(signal);
        const nr = schema.namedRanges[i];
        const name = nr?.name ?? "";
        const safeName = sheetNameDisallowed(nr?.sheetName ?? "") ? "[REDACTED]" : redactSchemaToken(name);
        const rangeA1 = redactRangeA1(nr?.rangeA1 ?? "");
        if (!name) continue;
        if (dlp) {
          const rect = nr?.rect;
          const sheetName = nr?.sheetName ?? "";
          const sheetId = sheetName ? resolveDlpSheetId(sheetName) : "";
          const range = rectToRange(rect);
          if (sheetId && range) {
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
              const redactedRange = redactedSheetRangeA1ForRect(rect) || "[REDACTED]";
              schemaLines.push(`- Named range [REDACTED] (range="${redactedRange}")`);
              continue;
            }
          }
        }

        schemaLines.push(`- Named range ${safeName} (range="${rangeA1}")`);
      }

      // Fallback: if the workbook doesn't supply explicit tables/named ranges, reuse the
      // already-indexed (and DLP-redacted) chunk metadata stored in the vector store.
      //
      // This improves chat quality by still providing a schema-first outline even when
      // retrieval is sparse or empty (e.g. query doesn't match anything).
      if (schemaLines.length === 0) {
        try {
          throwIfAborted(signal);
          const stored = await awaitWithAbort(
            vectorStore.list({ workbookId: params.workbook.id, includeVector: false, signal }),
            signal
          );
          /** @type {Array<{ id: string, metadata?: any }>} */
          const records = Array.isArray(stored) ? stored : [];
 
           /**
            * @param {string} text
            */
           const extractColumnsLine = (text) => {
            const raw = typeof text === "string" ? text : "";
            const lines = raw.split("\n");
            for (const line of lines) {
              if (line.startsWith("COLUMNS:")) return line.replace(/^COLUMNS:\s*/, "");
            }
            return "";
          };
  
          const candidates = records
            .map((r) => r?.metadata ?? {})
            .filter((m) => m && typeof m === "object")
            .filter((m) => m.workbookId === params.workbook.id)
            .filter((m) => m.kind === "table" || m.kind === "namedRange" || m.kind === "dataRegion")
            .sort((a, b) => {
              const sheetCmp = (typeof a.sheetName === "string" ? a.sheetName : "").localeCompare(
                typeof b.sheetName === "string" ? b.sheetName : "",
              );
              if (sheetCmp) return sheetCmp;
              const kindCmp = (typeof a.kind === "string" ? a.kind : "").localeCompare(typeof b.kind === "string" ? b.kind : "");
              if (kindCmp) return kindCmp;
              const ar = a.rect ?? {};
              const br = b.rect ?? {};
              const coordCmp =
                (Number(ar.r0 ?? 0) - Number(br.r0 ?? 0)) ||
                (Number(ar.c0 ?? 0) - Number(br.c0 ?? 0)) ||
                (Number(ar.r1 ?? 0) - Number(br.r1 ?? 0)) ||
                (Number(ar.c1 ?? 0) - Number(br.c1 ?? 0));
              if (coordCmp) return coordCmp;
              return (typeof a.title === "string" ? a.title : "").localeCompare(typeof b.title === "string" ? b.title : "");
            })
            .slice(0, maxTables);
  
          for (const meta of candidates) {
            throwIfAborted(signal);
            const kind = typeof meta.kind === "string" ? meta.kind : "";
            const title = typeof meta.title === "string" ? meta.title : "";
            const sheetName = typeof meta.sheetName === "string" ? meta.sheetName : "";
            const sheetNameTokenDisallowed = sheetNameDisallowed(sheetName);
            const titleTokenDisallowed =
              sheetNameTokenDisallowed && (kind === "table" || kind === "namedRange");
            const safeTitle = titleTokenDisallowed ? "[REDACTED]" : title ? redactSchemaToken(title) : "";
            const safeSheetName = sheetNameTokenDisallowed ? "[REDACTED]" : sheetName ? redactSchemaToken(sheetName) : "";
            const rect = meta.rect ?? {};
            const r0 = rect.r0;
            const c0 = rect.c0;
            const r1 = rect.r1;
            const c1 = rect.c1;
            if (!sheetName || ![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) continue;
            if (r1 < r0 || c1 < c0) continue;
            const rangeA1 = rangeToA1({
              sheetName: safeSheetName || sheetName,
              startRow: r0,
              startCol: c0,
              endRow: r1,
              endCol: c1,
            });

            // Structured DLP classifications: if this chunk range is disallowed due to explicit
            // document/sheet/range selectors, do not include any derived header/type strings.
            if (dlp) {
              const range = rectToRange({ r0, c0, r1, c1 });
              const storedSheetId = typeof meta.dlpSheetId === "string" ? meta.dlpSheetId.trim() : "";
              const sheetId = storedSheetId || resolveDlpSheetId(sheetName);
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
                  const label = kind === "table" ? "Table" : kind === "namedRange" ? "Named range" : "Data region";
                  const redactedRange = redactedSheetRangeA1ForRect({ r0, c0, r1, c1 }) || "[REDACTED]";
                  const nameSuffix = kind === "dataRegion" ? "" : title ? ` [REDACTED]` : "";
                  schemaLines.push(`- ${label}${nameSuffix} (range="${redactedRange}"): [REDACTED]`);
                  continue;
                }
              }
            }

            const storedText = typeof meta.text === "string" ? meta.text : "";
            const columns = extractColumnsLine(storedText);
            const isRedacted = /\[REDACTED\]/.test(storedText);

            const label = kind === "table" ? "Table" : kind === "namedRange" ? "Named range" : "Data region";
            const nameSuffix = kind === "dataRegion" ? "" : title ? ` ${safeTitle}` : "";
            const detail = columns ? redactColumnsLine(columns) : isRedacted ? "[REDACTED]" : "";
            if (!detail) continue;
            schemaLines.push(`- ${label}${nameSuffix} (range="${rangeA1}"): ${detail}`);
          }
        } catch {
          // Best-effort; if the vector store cannot list records, continue without a schema section.
        }
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
          text: (() => {
            const summary = {
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
            };

            const safeSummary = (() => {
              if (!dlp) return summary;
              throwIfAborted(signal);

              const index = getDlpDocumentIndex();

              /**
               * @param {string} sheetName
               */
              const sheetNameDisallowed = (sheetName) => {
                const raw = typeof sheetName === "string" ? sheetName : "";
                if (!raw) return false;
                const sheetId = resolveDlpSheetId(raw);
                // Best-effort: if we cannot build the structured selector index, fall back to
                // not redacting non-heuristic sheet names (heuristic redaction still applies).
                if (!index) return false;

                // Base classification applied to the entire sheet (document + sheet selectors).
                let baseClassification = maxClassification({ level: CLASSIFICATION_LEVEL.PUBLIC, labels: [] }, index.docClassificationMax);
                const sheetMax = index.sheetClassificationMaxBySheetId.get(sheetId);
                if (sheetMax) baseClassification = maxClassification(baseClassification, sheetMax);

                // If the sheet itself is disallowed due to doc/sheet scope, redact the name.
                const baseDecision = evaluatePolicy({
                  action: DLP_ACTION.AI_CLOUD_PROCESSING,
                  classification: baseClassification,
                  policy: dlp.policy,
                  options: { includeRestrictedContent },
                });
                if (baseDecision.decision !== DLP_DECISION.ALLOW) return true;

                // Otherwise, if any structured selector on the sheet would require redaction,
                // treat the sheet name as disallowed metadata.
                const colMap = index.columnClassificationBySheetId.get(sheetId);
                if (colMap) {
                  for (const colClassification of colMap.values()) {
                    throwIfAborted(signal);
                    const decision = evaluatePolicy({
                      action: DLP_ACTION.AI_CLOUD_PROCESSING,
                      classification: maxClassification(baseClassification, colClassification),
                      policy: dlp.policy,
                      options: { includeRestrictedContent },
                    });
                    if (decision.decision !== DLP_DECISION.ALLOW) return true;
                  }
                }

                const rangeRecords = index.rangeRecordsBySheetId.get(sheetId) ?? [];
                for (const rec of rangeRecords) {
                  throwIfAborted(signal);
                  const decision = evaluatePolicy({
                    action: DLP_ACTION.AI_CLOUD_PROCESSING,
                    classification: maxClassification(baseClassification, rec.classification),
                    policy: dlp.policy,
                    options: { includeRestrictedContent },
                  });
                  if (decision.decision !== DLP_DECISION.ALLOW) return true;
                }

                const cellMap = index.cellClassificationBySheetId.get(sheetId);
                if (cellMap) {
                  for (const cellClassification of cellMap.values()) {
                    throwIfAborted(signal);
                    const decision = evaluatePolicy({
                      action: DLP_ACTION.AI_CLOUD_PROCESSING,
                      classification: maxClassification(baseClassification, cellClassification),
                      policy: dlp.policy,
                      options: { includeRestrictedContent },
                    });
                    if (decision.decision !== DLP_DECISION.ALLOW) return true;
                  }
                }

                return false;
              };

              /**
               * @param {string} sheetName
               * @param {any} rect
               */
              const rectDisallowed = (sheetName, rect) => {
                const rawSheet = typeof sheetName === "string" ? sheetName : "";
                if (!rawSheet) return false;
                const sheetId = resolveDlpSheetId(rawSheet);
                if (!sheetId) return false;
                const range = rectToRange(rect);
                if (!range) return false;
                const recordClassification = index
                  ? effectiveRangeClassificationFromDocumentIndex(index, { documentId: dlp.documentId, sheetId, range }, signal)
                  : effectiveRangeClassification({ documentId: dlp.documentId, sheetId, range }, classificationRecords);
                const recordDecision = evaluatePolicy({
                  action: DLP_ACTION.AI_CLOUD_PROCESSING,
                  classification: recordClassification,
                  policy: dlp.policy,
                  options: { includeRestrictedContent },
                });
                return recordDecision.decision !== DLP_DECISION.ALLOW;
              };

              const safeWorkbookId = (() => {
                // When structured DLP requires redaction anywhere in the workbook (document/sheet/range),
                // treat the workbook id as disallowed metadata too. Workbook ids can be user-controlled and
                // may contain non-heuristic sensitive strings (e.g. "TopSecret") that a no-op redactor cannot
                // detect.
                if (structuredOverallDecision?.decision && structuredOverallDecision.decision !== DLP_DECISION.ALLOW) {
                  return "[REDACTED]";
                }
                // Otherwise, keep the previous heuristic-only behavior.
                return redactSchemaToken(summary.id, "workbook_summary");
              })();

              return {
                id: safeWorkbookId,
                sheets: (summary.sheets ?? []).map((name) => {
                  if (name != null && typeof name !== "string") return "[REDACTED]";
                  const rawName = typeof name === "string" ? name : "";
                  return sheetNameDisallowed(rawName) ? "[REDACTED]" : redactSchemaToken(rawName, "workbook_summary");
                }),
                tables: (summary.tables ?? []).map((t) => {
                  const sheetName = typeof t?.sheetName === "string" ? t.sheetName : "";
                  const sheetNameInvalid = t?.sheetName != null && typeof t.sheetName !== "string";
                  const sheetTokenDisallowed = sheetNameInvalid || sheetNameDisallowed(sheetName);
                  const safeSheetName = sheetTokenDisallowed ? "[REDACTED]" : redactSchemaToken(sheetName, "workbook_summary");
                  const disallowed = sheetNameInvalid ? true : rectDisallowed(sheetName, t?.rect);
                  const safeName =
                    disallowed || sheetTokenDisallowed || (t?.name != null && typeof t.name !== "string")
                      ? "[REDACTED]"
                      : redactSchemaToken(t?.name ?? "", "workbook_summary");
                  const shouldSanitizeRect = overallDecision?.decision === DLP_DECISION.REDACT;
                  const safeRect = shouldSanitizeRect ? normalizeWorkbookRect(t?.rect) : t?.rect;
                  return {
                    ...t,
                    name: safeName,
                    sheetName: safeSheetName,
                    ...(shouldSanitizeRect ? { rect: safeRect } : null),
                  };
                }),
                namedRanges: (summary.namedRanges ?? []).map((r) => {
                  const sheetName = typeof r?.sheetName === "string" ? r.sheetName : "";
                  const sheetNameInvalid = r?.sheetName != null && typeof r.sheetName !== "string";
                  const sheetTokenDisallowed = sheetNameInvalid || sheetNameDisallowed(sheetName);
                  const safeSheetName = sheetTokenDisallowed ? "[REDACTED]" : redactSchemaToken(sheetName, "workbook_summary");
                  const disallowed = sheetNameInvalid ? true : rectDisallowed(sheetName, r?.rect);
                  const safeName =
                    disallowed || sheetTokenDisallowed || (r?.name != null && typeof r.name !== "string")
                      ? "[REDACTED]"
                      : redactSchemaToken(r?.name ?? "", "workbook_summary");
                  const shouldSanitizeRect = overallDecision?.decision === DLP_DECISION.REDACT;
                  const safeRect = shouldSanitizeRect ? normalizeWorkbookRect(r?.rect) : r?.rect;
                  return {
                    ...r,
                    name: safeName,
                    sheetName: safeSheetName,
                    ...(shouldSanitizeRect ? { rect: safeRect } : null),
                  };
                }),
              };
            })();

            return this.redactor(`Workbook summary:\n${stableJsonStringify(safeSummary)}`);
          })(),
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
          text:
            Array.isArray(attachmentsForPrompt) && attachmentsForPrompt.length
              ? this.redactor(`User-provided attachments:\n${stableJsonStringify(attachmentsForPrompt)}`)
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

    const retrievedOut =
      dlp && overallDecision?.decision === DLP_DECISION.REDACT
        ? redactStructuredValue(retrievedChunks, this.redactor, {
            signal,
            includeRestrictedContent,
            policyAllowsRestrictedContent,
          })
        : retrievedChunks;
    return {
      indexStats,
      retrieved: retrievedOut,
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
   *     auditLogger?: { log(event: any): void },
   *     sheetNameResolver?: any,
   *   }
   * }} params
   */
  async buildWorkbookContextFromSpreadsheetApi(params) {
    const signal = params.signal;
    throwIfAborted(signal);
    const skipIndexing = (params.skipIndexing ?? false) === true;
    const skipIndexingWithDlp = (params.skipIndexingWithDlp ?? false) === true;
    const includeFormulaValues = (params.includeFormulaValues ?? params.include_formula_values ?? false) === true;

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
            includeFormulaValues,
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
 * Compact a SheetSchema for prompt inclusion.
 *
 * `extractSheetSchema()` includes small `sampleValues` arrays per column which can be:
 * - token-expensive
 * - privacy-sensitive (raw cell values)
 *
 * For prompt context, we keep a schema-first representation (names, ranges, types, counts)
 * without embedding raw sample cell values.
 *
 * @param {any} schema
 */
function compactSheetSchemaForPrompt(schema, options = {}) {
  if (!schema || typeof schema !== "object") return schema;
  const maxTables = normalizeNonNegativeInt(options.maxTables, 25);
  const maxRegions = normalizeNonNegativeInt(options.maxRegions, 25);
  const maxNamedRanges = normalizeNonNegativeInt(options.maxNamedRanges, 25);
  const maxColumns = normalizeNonNegativeInt(options.maxColumns, 25);

  const tables = Array.isArray(schema.tables) ? schema.tables : [];
  const namedRanges = Array.isArray(schema.namedRanges) ? schema.namedRanges : [];
  const dataRegions = Array.isArray(schema.dataRegions) ? schema.dataRegions : [];

  return {
    name: typeof schema.name === "string" ? schema.name : "",
    tables: tables.slice(0, maxTables).map((t) => {
      const columns = Array.isArray(t?.columns) ? t.columns : [];
      return {
        name: typeof t?.name === "string" ? t.name : "",
        range: typeof t?.range === "string" ? t.range : "",
        rowCount: Number.isFinite(t?.rowCount) ? Math.max(0, Math.floor(t.rowCount)) : 0,
        columns: columns.slice(0, maxColumns).map((c) => ({
          name: typeof c?.name === "string" ? c.name : "",
          type: typeof c?.type === "string" ? c.type : "mixed",
        })),
      };
    }),
    namedRanges: namedRanges.slice(0, maxNamedRanges).map((r) => ({
      name: typeof r?.name === "string" ? r.name : "",
      range: typeof r?.range === "string" ? r.range : "",
    })),
    dataRegions: dataRegions.slice(0, maxRegions).map((r) => ({
      range: typeof r?.range === "string" ? r.range : "",
      hasHeader: Boolean(r?.hasHeader),
      headers: Array.isArray(r?.headers) ? r.headers.map((h) => String(h ?? "")) : [],
      inferredColumnTypes: Array.isArray(r?.inferredColumnTypes)
        ? r.inferredColumnTypes.map((t) => String(t ?? "mixed"))
        : [],
      rowCount: Number.isFinite(r?.rowCount) ? Math.max(0, Math.floor(r.rowCount)) : 0,
      columnCount: Number.isFinite(r?.columnCount) ? Math.max(0, Math.floor(r.columnCount)) : 0,
    })),
  };
}

/**
 * Like `classifyText()`, but also detects percent-encoded representations of sensitive tokens
 * (e.g. "alice%40example.com" -> "alice@example.com").
 *
 * This is primarily to prevent metadata leaks via encoded identifiers (chunk ids, etc) when
 * DLP requires redaction and callers use a no-op redactor.
 *
 * @param {string} text
 * @returns {ReturnType<typeof classifyText>}
 */
function classifyTextForDlp(text) {
  const raw = String(text ?? "");
  const direct = classifyText(raw);
  if (direct.level === "sensitive") return direct;
  if (!raw.includes("%")) return direct;
  if (!/%[0-9A-Fa-f]{2}/.test(raw)) return direct;
  try {
    const decoded = decodeURIComponent(raw);
    if (decoded && decoded !== raw) {
      const decodedHeuristic = classifyText(decoded);
      if (decodedHeuristic.level === "sensitive") return decodedHeuristic;
    }
  } catch {
    // ignore decode errors (invalid percent sequences)
  }
  return direct;
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
      const heuristic = (() => {
        const t = typeof cell;
        if (t === "string" || t === "number" || t === "bigint") {
          return classifyTextForDlp(String(cell));
        }
        if (t === "object") {
          // Sheet cells can contain rich values (objects) whose visible text lives in nested fields
          // (e.g. `{ text, runs }` for rich text, typed values, in-cell images, etc). Reuse the
          // structured traversal logic (bounded) so we still detect heuristic-sensitive strings.
          return classifyStructuredForDlp(cell, { signal, maxNodes: 200, maxDepth: 10, maxStringLength: 10_000 });
        }
        return { level: "public", findings: [] };
      })();
      if (heuristic.level !== "sensitive") continue;
      for (const f of heuristic.findings || []) {
        if (typeof f === "string") findings.add(f);
        else if (typeof f === "number" || typeof f === "boolean" || typeof f === "bigint") findings.add(String(f));
      }
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
 * Heuristically classify a structured value (attachments, arbitrary objects) for DLP.
 *
 * This mirrors `classifyValuesForDlp()` but for non-tabular values, with:
 *  - cycle detection
 *  - bounded traversal to avoid pathological user-provided payloads
 *
 * @param {unknown} value
 * @param {{ signal?: AbortSignal, maxNodes?: number, maxDepth?: number, maxStringLength?: number }} [options]
 * @returns {ReturnType<typeof classifyText>}
 */
function classifyStructuredForDlp(value, options = {}) {
  const signal = options.signal;
  const maxNodes = normalizeNonNegativeInt(options.maxNodes, 5_000);
  const maxDepth = normalizeNonNegativeInt(options.maxDepth, 20);
  const maxStringLength = normalizeNonNegativeInt(options.maxStringLength, 10_000);
  /** @type {Set<string>} */
  const findings = new Set();
  let found = false;
  let nodes = 0;
  const seen = new WeakSet();

  /**
   * @param {unknown} text
   */
  function considerText(text) {
    if (found) return;
    const raw = String(text ?? "");
    const truncated = raw.length > maxStringLength ? raw.slice(0, maxStringLength) : raw;
    const heuristic = classifyTextForDlp(truncated);
    if (heuristic.level !== "sensitive") return;
    found = true;
    for (const f of heuristic.findings || []) findings.add(String(f));
  }

  /**
   * @param {unknown} v
   * @param {number} depth
   */
  function visit(v, depth) {
    if (found) return;
    throwIfAborted(signal);
    nodes += 1;
    if (nodes > maxNodes) return;
    if (v === null || v === undefined) return;
    const t = typeof v;
    if (t === "string" || t === "number" || t === "bigint") {
      considerText(v);
      return;
    }
    if (t === "boolean" || t === "symbol" || t === "function") return;
    if (t !== "object") return;

    // Cycle detection: attachments can contain arbitrary nested objects (including Map/Set/arrays).
    // Use a shared WeakSet for the whole traversal so self-referential structures do not:
    //  - hang traversal
    //  - exceed maxNodes without ever reaching later sibling fields that might contain sensitive content
    //  - overflow the call stack
    if (seen.has(v)) return;
    seen.add(v);

    if (v instanceof Date) {
      considerText(v.toISOString());
      return;
    }
    // Avoid scanning large binary blobs.
    if (v instanceof ArrayBuffer || ArrayBuffer.isView(v)) return;
    if (depth >= maxDepth) return;

    if (v instanceof Map) {
      for (const [k, val] of v.entries()) {
        if (found) break;
        visit(k, depth + 1);
        if (found) break;
        visit(val, depth + 1);
      }
      return;
    }

    if (v instanceof Set) {
      for (const val of v.values()) {
        if (found) break;
        visit(val, depth + 1);
      }
      return;
    }

    if (Array.isArray(v)) {
      for (const item of v) {
        if (found) break;
        visit(item, depth + 1);
      }
      return;
    }

    if (t === "object") {
      const obj = /** @type {Record<string, unknown>} */ (v);
      for (const key of Object.keys(obj)) {
        if (found) break;
        considerText(key);
        if (found) break;
        visit(obj[key], depth + 1);
      }
    }
  }

  visit(value, 0);
  return findings.size > 0 ? { level: "sensitive", findings: [...findings] } : { level: "public", findings: [] };
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
 * Apply a redactor to sheet cell values and ensure heuristic-sensitive values do not leak
 * under DLP redaction.
 *
 * Note: We treat numbers/bigints as potential carriers of sensitive patterns (e.g. credit
 * card numbers) because prompt formatting stringifies them.
 *
 * @param {unknown[][]} values
 * @param {(text: string) => string} redactor
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {unknown[][]}
 */
function redactValuesForDlp(values, redactor, options = {}) {
  const signal = options.signal;
  const includeRestrictedContent = options.includeRestrictedContent ?? false;
  const policyAllowsRestrictedContent = options.policyAllowsRestrictedContent ?? false;
  const restrictedAllowed = includeRestrictedContent && policyAllowsRestrictedContent;
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
      const isTextLike = typeof cell === "string" || typeof cell === "number" || typeof cell === "bigint";
      if (isTextLike) {
        const raw = String(cell);
        const redacted = redactor(raw);
        // Defense-in-depth: if the configured redactor is a no-op (or incomplete),
        // ensure heuristic sensitive patterns never slip through under DLP redaction.
        if (!restrictedAllowed && classifyTextForDlp(redacted).level === "sensitive") {
          nextRow.push("[REDACTED]");
          continue;
        }

        // Preserve non-string primitives when redaction is a no-op (keeps API behavior stable),
        // but allow redactors to return strings when they transform the content.
        if (typeof cell !== "string" && redacted === raw) {
          nextRow.push(cell);
          continue;
        }

        nextRow.push(redacted);
        continue;
      }

      // Rich cell values (DocumentController rich text, typed values, etc) can be objects that
      // still render into prompt/embedding text via `valuesRangeToTsv`. Under DLP REDACT, deep-redact
      // those objects so heuristic-sensitive strings cannot leak to embeddings even when the
      // configured redactor is a no-op.
      if (cell && typeof cell === "object") {
        nextRow.push(redactStructuredValue(cell, redactor, { signal, includeRestrictedContent, policyAllowsRestrictedContent }));
        continue;
      }

      nextRow.push(cell);
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
  const policyAllowsRestrictedContent = options.policyAllowsRestrictedContent ?? false;
  const restrictedAllowed = includeRestrictedContent && policyAllowsRestrictedContent;
  // Cycle detection for arbitrarily nested user-provided attachment objects.
  // Use WeakSet so redaction does not retain references beyond this call.
  const seen = options._seen ?? new WeakSet();
  throwIfAborted(signal);

  if (typeof value === "string") {
    const redacted = redactor(value);
    if (!restrictedAllowed && classifyTextForDlp(redacted).level === "sensitive") {
      return /** @type {T} */ ("[REDACTED]");
    }
    return /** @type {T} */ (redacted);
  }
  if (typeof value === "number" || typeof value === "bigint") {
    const raw = String(value);
    const redacted = redactor(raw);
    if (!restrictedAllowed && classifyTextForDlp(redacted).level === "sensitive") {
      return /** @type {T} */ ("[REDACTED]");
    }
    if (redacted !== raw) {
      return /** @type {T} */ (redacted);
    }
    return value;
  }
  if (value === null || value === undefined) return value;
  if (typeof value !== "object") return value;
  if (value instanceof Date) return value;

  // Avoid infinite recursion for cyclic structures.
  if (seen.has(value)) {
    return /** @type {T} */ ("[REDACTED]");
  }
  seen.add(value);

  // Typed arrays / ArrayBuffers can contain arbitrary bytes and can be very large. Avoid
  // enumerating them into JSON (and avoid leaking raw bytes) under DLP redaction.
  if (typeof ArrayBuffer !== "undefined") {
    if (value instanceof ArrayBuffer) {
      return /** @type {T} */ ("[REDACTED]");
    }
    if (typeof ArrayBuffer.isView === "function" && ArrayBuffer.isView(value)) {
      return /** @type {T} */ ("[REDACTED]");
    }
  }

  if (value instanceof Map) {
    /** @type {Map<any, any>} */
    const out = new Map();
    for (const [k, v] of value.entries()) {
      throwIfAborted(signal);
      out.set(
        redactStructuredValue(k, redactor, { signal, includeRestrictedContent, policyAllowsRestrictedContent, _seen: seen }),
        redactStructuredValue(v, redactor, { signal, includeRestrictedContent, policyAllowsRestrictedContent, _seen: seen }),
      );
    }
    return /** @type {T} */ (out);
  }

  if (value instanceof Set) {
    /** @type {Set<any>} */
    const out = new Set();
    for (const v of value.values()) {
      throwIfAborted(signal);
      out.add(
        redactStructuredValue(v, redactor, { signal, includeRestrictedContent, policyAllowsRestrictedContent, _seen: seen }),
      );
    }
    return /** @type {T} */ (out);
  }

  if (Array.isArray(value)) {
    return /** @type {T} */ (
      value.map((v) =>
        redactStructuredValue(v, redactor, { signal, includeRestrictedContent, policyAllowsRestrictedContent, _seen: seen }),
      )
    );
  }

  const proto = Object.getPrototypeOf(value);
  const entries = Object.entries(value);
  if (entries.length === 0) return value;
  // For class instances and other non-plain objects, fall back to a shallow key/value
  // projection matching how JSON/stableJsonStringify will serialize them.
  if (proto !== Object.prototype && proto !== null && entries.length > 200) {
    return /** @type {T} */ ("[REDACTED]");
  }

  /** @type {any} */
  const out = {};
  /** @type {Set<string>} */
  const usedKeys = new Set();
  let redactedKeyIndex = 0;
  for (const [key, v] of entries) {
    throwIfAborted(signal);
    let outKey = key;
    if (!restrictedAllowed) {
      const redactedKey = redactor(key);
      if (classifyTextForDlp(redactedKey).level === "sensitive") {
        // If the key itself is sensitive and redaction did not remove the pattern,
        // replace it entirely with a stable placeholder.
        outKey = `[REDACTED_KEY_${redactedKeyIndex++}]`;
      } else if (redactedKey !== key) {
        // Prefer the redactor's placeholder when it successfully removes the sensitive
        // pattern (e.g. `[REDACTED_EMAIL]`).
        outKey = redactedKey;
      }
    }

    // Ensure we do not overwrite sibling keys if multiple keys redact to the same placeholder.
    if (usedKeys.has(outKey)) {
      const base = outKey;
      let suffix = 1;
      while (usedKeys.has(`${base}_${suffix}`)) suffix += 1;
      outKey = `${base}_${suffix}`;
    }
    usedKeys.add(outKey);

    out[outKey] = redactStructuredValue(v, redactor, {
      signal,
      includeRestrictedContent,
      policyAllowsRestrictedContent,
      _seen: seen,
    });
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
 * @param {{ maxRows?: number, maxAttachments?: number, sheetNameForOutput?: string, signal?: AbortSignal }} [options]
 */
function buildRangeAttachmentSectionText(params, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const attachments = Array.isArray(params.attachments) ? params.attachments : [];
  if (attachments.length === 0) return "";
  const sheet = params.sheet;
  const sheetName = sheet?.name ?? "";
  const sheetNameForOutput = typeof options.sheetNameForOutput === "string" ? options.sheetNameForOutput : sheetName;
  const normalizedSheetName = normalizeSheetNameForComparison(sheetName);
  const maxRows = options.maxRows ?? 30;
  const maxAttachments = options.maxAttachments ?? 3;

  const values = Array.isArray(sheet?.values) ? sheet.values : [];
  const matrixRowCount = values.length;
  let matrixColCount = 0;
  const LARGE_ROW_THRESHOLD = 10_000;
  if (matrixRowCount > LARGE_ROW_THRESHOLD) {
    // Avoid iterating over holes in sparse Excel-scale matrices.
    for (const key in values) {
      throwIfAborted(signal);
      const row = values[key];
      if (!Array.isArray(row)) continue;
      if (row.length > matrixColCount) matrixColCount = row.length;
    }
  } else {
    for (const row of values) {
      throwIfAborted(signal);
      if (!Array.isArray(row)) continue;
      if (row.length > matrixColCount) matrixColCount = row.length;
    }
  }

  const origin = getSheetOrigin(sheet);
  const availableRange =
    matrixRowCount > 0 && matrixColCount > 0
      ? rangeToA1({
          sheetName: sheetNameForOutput,
          startRow: origin.row,
          startCol: origin.col,
          endRow: origin.row + matrixRowCount - 1,
          endCol: origin.col + matrixColCount - 1,
        })
      : "";

  const entries = [];
  let included = 0;

  for (const attachment of attachments) {
    throwIfAborted(signal);
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

    const canonicalRange = rangeToA1({ ...parsed, sheetName: sheetNameForOutput || sheetName });

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

    entries.push(`${canonicalRange}:\n${valuesRangeToTsv(values, clamped, { maxRows, signal })}`);
    included += 1;
  }

  if (entries.length === 0) return "";
  return `Attached range data:\n${entries.join("\n\n")}`;
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
