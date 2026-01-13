import type { LLMClient, LLMMessage } from "../../../../packages/llm/src/index.js";

import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";
import { AIAuditRecorder } from "../../../../packages/ai-audit/src/recorder.js";

import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import {
  CLASSIFICATION_LEVEL,
  DEFAULT_CLASSIFICATION,
  classificationRank,
  maxClassification,
  normalizeClassification,
} from "../../../../packages/security/dlp/src/classification.js";
import { DLP_DECISION, evaluatePolicy } from "../../../../packages/security/dlp/src/policyEngine.js";
import { effectiveCellClassification, effectiveRangeClassification, normalizeRange, selectorKey } from "../../../../packages/security/dlp/src/selectors.js";

import { parseA1Range, rangeToA1, type RangeAddress } from "./a1.js";
import {
  PROVENANCE_REF_SEPARATOR,
  type AiFunctionArgumentProvenance,
  type AiFunctionEvaluator,
  type CellValue,
  type ProvenanceCellValue,
  isSpreadsheetErrorCode,
  type SpreadsheetValue,
} from "./evaluateFormula.js";

import type { SheetNameResolver } from "../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";

import { getDesktopAIAuditStore } from "../ai/audit/auditStore.js";
import { getAiCloudDlpOptions } from "../ai/dlp/aiDlp.js";
import { getDesktopLLMClient, getDesktopModel } from "../ai/llm/desktopLLMClient.js";

export const AI_CELL_PLACEHOLDER = "#GETTING_DATA";
export const AI_CELL_DLP_ERROR = "#DLP!";
export const AI_CELL_ERROR = "#AI!";

const DLP_REDACTION_PLACEHOLDER = "[REDACTED]";

const DEFAULT_RANGE_SAMPLE_LIMIT = 200;
const DEFAULT_ERROR_CACHE_TTL_MS = 60_000;
const MAX_PROMPT_CHARS = 2_000;
const MAX_SCALAR_CHARS = 500;
const MAX_RANGE_HEADER_VALUES = 20;
const MAX_RANGE_PREVIEW_VALUES = 30;
const MAX_RANGE_SAMPLE_VALUES = 30;
const MAX_USER_MESSAGE_CHARS = 16_000;
const DEFAULT_MAX_CONCURRENT_REQUESTS = 4;
const DEFAULT_REQUEST_TIMEOUT_MS = 30_000;
const MAX_OUTPUT_CHARS = 10_000;

const DLP_INDEX_CACHE_MAX_ENTRIES = 5;

// Cache persistence intentionally uses a short debounce to avoid repeatedly serializing
// the entire cache map when many AI cells resolve in a burst.
const CACHE_PERSIST_DEBOUNCE_MS = 50;

type DlpNormalizedRange = ReturnType<typeof normalizeRange>;

type DlpCellIndex = {
  documentId: string;
  docClassificationMax: any;
  sheetClassificationMaxBySheetId: Map<string, any>;
  columnClassificationBySheetId: Map<string, Map<number, any>>;
  cellClassificationBySheetId: Map<string, Map<string, any>>;
  rangeRecordsBySheetId: Map<string, Array<{ range: DlpNormalizedRange; classification: any }>>;
  fallbackRecordsBySheetId: Map<string, Array<{ selector: any; classification: any }>>;
};

export interface AiCellFunctionEngineOptions {
  llmClient?: LLMClient;
  model?: string;
  auditStore?: AIAuditStore;
  workbookId?: string;
  sessionId?: string;
  userId?: string;
  onUpdate?: () => void;
  /**
   * Optional resolver for mapping user-facing sheet display names (used in formula text)
   * back to stable sheet ids (used by DLP classifications and other internal metadata).
   */
  sheetNameResolver?: SheetNameResolver | null;
  cache?: {
    /**
     * When set, persists cache entries in localStorage under this key.
     */
    persistKey?: string;
    /**
     * Maximum number of cached entries to retain.
     */
    maxEntries?: number;
    /**
     * How long to keep `#AI!` error cache entries before retrying.
     * Successful values are cached indefinitely.
     */
    errorTtlMs?: number;
  };
  limits?: {
    /**
     * Maximum number of cells to materialize for direct range references passed into AI() calls.
     *
     * This feeds `AiFunctionEvaluator.rangeSampleLimit`, which `evaluateFormula` uses to avoid
     * materializing unbounded arrays.
     */
    maxInputCells?: number;
    /**
     * Hard cap on the user message sent to the model (character-based heuristic).
     */
    maxPromptChars?: number;
    /**
     * Hard cap on `prompt` stored in audit entries.
     */
    maxAuditPreviewChars?: number;
    /**
     * Hard cap on scalar cell values serialized into prompt compactions.
     */
    maxCellChars?: number;
    /**
     * Maximum number of parallel LLM requests kicked off by AI cell recalculation.
     *
     * When many AI() formulas are recalculated at once (e.g. pasting a column),
     * the engine will queue additional requests beyond this limit while still
     * returning `#GETTING_DATA` synchronously.
     */
    maxConcurrentRequests?: number;
    /**
     * Hard cap on the final AI cell output stored in cache/returned to the spreadsheet.
     *
     * Large string outputs can freeze the grid/UI, so we clamp deterministically.
     */
    maxOutputChars?: number;
    /**
     * Maximum time to wait for a single LLM request before failing.
     *
     * Prevents AI cell functions from getting stuck on `#GETTING_DATA` forever if the
     * backend hangs or the network stalls.
     */
    requestTimeoutMs?: number;
  };
}

type CacheEntry = { value: SpreadsheetValue; updatedAtMs: number };

class ConcurrencyLimiter {
  private readonly maxConcurrent: number;
  private active = 0;
  private readonly queue: Array<() => void> = [];

  constructor(maxConcurrent: number) {
    this.maxConcurrent = clampInt(maxConcurrent, { min: 1, max: 10_000 });
  }

  run<T>(start: (release: () => void) => Promise<T>): Promise<T> {
    if (this.active < this.maxConcurrent) {
      return this.start(start);
    }

    return new Promise<T>((resolve, reject) => {
      this.queue.push(() => {
        this.run(start).then(resolve, reject);
      });
    });
  }

  private start<T>(start: (release: () => void) => Promise<T>): Promise<T> {
    this.active += 1;

    let released = false;
    const release = () => {
      if (released) return;
      released = true;
      this.active -= 1;
      this.drain();
    };

    let promise: Promise<T>;
    try {
      promise = Promise.resolve(start(release));
    } catch (error) {
      release();
      return Promise.reject(error);
    }

    // Failsafe: ensure slots are returned even if the caller forgets to `release()`.
    // Avoid unhandled rejections by consuming the returned promise.
    promise.finally(release).catch(() => undefined);
    return promise;
  }

  private drain(): void {
    while (this.active < this.maxConcurrent && this.queue.length > 0) {
      const next = this.queue.shift();
      if (next) next();
    }
  }
}

/**
 * Manages async AI cell function evaluation for a workbook/session:
 * - returns `#GETTING_DATA` synchronously while an LLM request is pending
 * - caches resolved values keyed by function+prompt+inputsHash
 * - applies DLP policy enforcement / redaction before sending to cloud models
 * - writes audit entries to `@formula/ai-audit`
 */
export class AiCellFunctionEngine implements AiFunctionEvaluator {
  private readonly llmClient: LLMClient;
  private readonly model: string;
  private readonly auditStore: AIAuditStore;
  private readonly workbookId: string;
  private readonly sessionId: string;
  private readonly userId?: string;
  private readonly onUpdate?: () => void;
  private readonly sheetNameResolver: SheetNameResolver | null;

  private readonly cachePersistKey?: string;
  private readonly cacheMaxEntries: number;
  private readonly cacheErrorTtlMs: number;

  private readonly maxInputCells: number;
  private readonly maxUserMessageChars: number;
  private readonly maxAuditPreviewChars: number;
  private readonly maxCellChars: number;
  private readonly maxOutputChars: number;
  private readonly requestTimeoutMs: number;

  private readonly requestLimiter: ConcurrencyLimiter;

  private readonly cache = new Map<string, CacheEntry>();
  private readonly inFlightByKey = new Map<string, Promise<void>>();
  private readonly pendingAudits = new Set<Promise<void>>();
  private readonly dlpIndexCache = new Map<string, DlpCellIndex>();

  private cachePersistTimer: ReturnType<typeof setTimeout> | null = null;
  private cachePersistPromise: Promise<void> | null = null;
  private cachePersistPromiseResolve: (() => void) | null = null;

  constructor(options: AiCellFunctionEngineOptions = {}) {
    this.llmClient = options.llmClient ?? getDesktopLLMClient();
    this.model = options.model ?? (this.llmClient as any)?.model ?? getDesktopModel();
    this.auditStore = options.auditStore ?? getDesktopAIAuditStore();
    this.workbookId = options.workbookId ?? "local-workbook";
    this.sessionId = options.sessionId ?? createSessionId(this.workbookId);
    this.userId = options.userId;
    this.onUpdate = options.onUpdate;
    this.sheetNameResolver = options.sheetNameResolver ?? null;

    this.cachePersistKey = options.cache?.persistKey;
    this.cacheMaxEntries = options.cache?.maxEntries ?? 500;
    const errorTtlMs = options.cache?.errorTtlMs;
    this.cacheErrorTtlMs =
      typeof errorTtlMs === "number" && Number.isFinite(errorTtlMs) ? Math.max(0, Math.trunc(errorTtlMs)) : DEFAULT_ERROR_CACHE_TTL_MS;

    this.maxInputCells = clampInt(options.limits?.maxInputCells ?? DEFAULT_RANGE_SAMPLE_LIMIT, { min: 1, max: 10_000 });
    this.maxUserMessageChars = clampInt(options.limits?.maxPromptChars ?? MAX_USER_MESSAGE_CHARS, {
      min: 1_000,
      max: 1_000_000,
    });
    this.maxAuditPreviewChars = clampInt(options.limits?.maxAuditPreviewChars ?? 2_000, { min: 200, max: 100_000 });
    this.maxCellChars = clampInt(options.limits?.maxCellChars ?? MAX_SCALAR_CHARS, { min: 50, max: 100_000 });
    this.maxOutputChars = clampInt(options.limits?.maxOutputChars ?? MAX_OUTPUT_CHARS, { min: 1, max: 1_000_000 });
    this.requestTimeoutMs = clampInt(options.limits?.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS, { min: 1, max: 3_600_000 });

    this.requestLimiter = new ConcurrencyLimiter(options.limits?.maxConcurrentRequests ?? DEFAULT_MAX_CONCURRENT_REQUESTS);

    this.loadCacheFromStorage();
  }

  get rangeSampleLimit(): number {
    return this.maxInputCells;
  }

  evaluateAiFunction(params: {
    name: string;
    args: CellValue[];
    cellAddress?: string;
    argProvenance?: AiFunctionArgumentProvenance[];
  }): SpreadsheetValue {
    const fn = params.name.toUpperCase();
    const cellAddress = params.cellAddress;

    if (params.args.length === 0) return "#VALUE!";
    if ((fn === "AI.EXTRACT" || fn === "AI.CLASSIFY" || fn === "AI.TRANSLATE") && params.args.length < 2) return "#VALUE!";

    const argError = firstErrorCode(params.args);
    if (argError) return argError;

    const defaultSheetId = sheetIdFromCellAddress(cellAddress);
    const dlp = getAiCloudDlpOptions({
      documentId: this.workbookId,
      sheetId: defaultSheetId,
      sheetNameResolver: this.sheetNameResolver,
    });

    const alignedProvenance = alignArgProvenance({
      args: params.args,
      provenance: params.argProvenance,
      defaultSheetId,
      documentId: this.workbookId,
      functionName: fn,
      sheetNameResolver: this.sheetNameResolver,
    });

    const referencedSheetIds = collectReferencedSheetIds({
      provenance: alignedProvenance,
      defaultSheetId,
      sheetNameResolver: this.sheetNameResolver,
    });

    const hasCellRefs = alignedProvenance.some((prov) => (prov?.cells?.length ?? 0) > 0);
    const hasAnyRefs = alignedProvenance.some((prov) => (prov?.cells?.length ?? 0) > 0 || (prov?.ranges?.length ?? 0) > 0);
    let classificationIndex: DlpCellIndex | null =
      hasCellRefs && dlp.classificationRecords.length > 0
        ? this.getMemoizedDlpCellIndex({ sheetIds: referencedSheetIds, records: dlp.classificationRecords })
        : null;

    // AI prompts are authored directly in formulas, so we can safely inspect a bit more text
    // for heuristic DLP classification than we allow for arbitrary cell values.
    const literalMaxChars = Math.max(this.maxCellChars, MAX_PROMPT_CHARS);

    const { selectionClassification, references } = computeSelectionClassification({
      documentId: this.workbookId,
      args: params.args,
      provenance: alignedProvenance,
      defaultSheetId,
      maxCellChars: literalMaxChars,
      classificationRecords: dlp.classificationRecords,
      classificationIndex,
      sheetNameResolver: this.sheetNameResolver,
    });

    const decision = evaluatePolicy({
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      classification: selectionClassification,
      policy: dlp.policy,
      options: { includeRestrictedContent: false },
    });
    const maxAllowedRank = decision.maxAllowed === null ? null : classificationRank(decision.maxAllowed);

    if (!classificationIndex && decision.decision !== DLP_DECISION.ALLOW && hasAnyRefs && dlp.classificationRecords.length > 0) {
      classificationIndex = this.getMemoizedDlpCellIndex({ sheetIds: referencedSheetIds, records: dlp.classificationRecords });
    }

    const { prompt, inputs, inputsHash, inputsCompaction, redactedCount } = preparePromptAndInputs({
      functionName: fn,
      args: params.args,
      provenance: alignedProvenance,
      decision,
      maxAllowedRank,
      documentId: this.workbookId,
      defaultSheetId,
      policy: dlp.policy,
      classificationRecords: dlp.classificationRecords,
      classificationIndex,
      maxCellChars: this.maxCellChars,
      maxLiteralChars: literalMaxChars,
      sheetNameResolver: this.sheetNameResolver,
    });

    const promptHash = hashText(prompt);

    const cacheKey = `${this.model}\u0000${fn}\u0000${promptHash}\u0000${inputsHash}`;

    if (decision.decision === DLP_DECISION.BLOCK) {
      dlp.auditLogger?.log({
        type: "ai.cell_function",
        documentId: this.workbookId,
        sheetId: defaultSheetId,
        cell: cellAddress,
        function: fn,
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision,
        selectionClassification,
        redactedCellCount: redactedCount,
        prompt_hash: promptHash,
        inputs_hash: inputsHash,
        references,
      });

      const auditPromise = this.auditBlockedRun({
        functionName: fn,
        prompt,
        inputsHash,
        inputsCompaction,
        cellAddress,
        references,
        dlp: { decision, selectionClassification, redactedCount },
      });
      this.trackPendingAudit(auditPromise);

      return AI_CELL_DLP_ERROR;
    }

    if (decision.decision === DLP_DECISION.REDACT) {
      dlp.auditLogger?.log({
        type: "ai.cell_function",
        documentId: this.workbookId,
        sheetId: defaultSheetId,
        cell: cellAddress,
        function: fn,
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision,
        selectionClassification,
        redactedCellCount: redactedCount,
        prompt_hash: promptHash,
        inputs_hash: inputsHash,
        references,
      });
    }

    const cached = this.cache.get(cacheKey);
    if (cached) {
      if (cached.value === AI_CELL_ERROR) {
        const ageMs = Date.now() - cached.updatedAtMs;
        if (Number.isFinite(ageMs) && ageMs >= this.cacheErrorTtlMs) {
          this.cache.delete(cacheKey);
          this.saveCacheToStorage();
        } else {
          return cached.value;
        }
      } else {
        return cached.value;
      }
    }

    if (this.inFlightByKey.has(cacheKey)) return AI_CELL_PLACEHOLDER;

    this.startRequest({
      cacheKey,
      functionName: fn,
      prompt,
      inputs,
      inputsHash,
      inputsCompaction,
      cellAddress,
      references,
      dlp: { decision, selectionClassification, redactedCount },
    });

    return AI_CELL_PLACEHOLDER;
  }

  /**
   * Await all in-flight LLM requests and pending audit flushes (useful in tests).
   */
  async waitForIdle(): Promise<void> {
    // Keep draining in case awaiting a promise schedules more work.
    while (this.inFlightByKey.size > 0 || this.pendingAudits.size > 0) {
      const snapshot = [...Array.from(this.inFlightByKey.values()), ...Array.from(this.pendingAudits)];
      await Promise.all(snapshot.map((p) => p.catch(() => undefined)));
    }

    // Also ensure any debounced cache persistence has been flushed so tests can
    // deterministically inspect localStorage.
    await this.flushCachePersistenceForTests();
  }

  /**
   * Flush any pending cache persistence. Intended for tests; production code
   * should rely on the debounced persistence mechanism.
   */
  async flushCachePersistenceForTests(): Promise<void> {
    const pending = this.cachePersistPromise;
    this.flushCachePersistenceNow();
    if (pending) await pending;
  }

  private getMemoizedDlpCellIndex(params: {
    sheetIds: Set<string>;
    records: Array<{ selector: any; classification: any }>;
  }): DlpCellIndex {
    // NOTE: We key by workbookId + referenced sheet ids + a hash of the classification record set.
    // This allows reusing the expensive per-cell classification index across multiple AI() evaluations
    // in the same workbook/session without leaking across workbooks or missing updates.
    const sheetIdsKey = stableJsonStringify(Array.from(params.sheetIds).sort());
    const recordsHash = hashClassificationRecords(params.records);
    const cacheKey = `${this.workbookId}\u0000${sheetIdsKey}\u0000${recordsHash}`;

    const existing = this.dlpIndexCache.get(cacheKey);
    if (existing) {
      // Refresh LRU position.
      this.dlpIndexCache.delete(cacheKey);
      this.dlpIndexCache.set(cacheKey, existing);
      return existing;
    }

    const built = __dlpIndexBuilder.buildDlpCellIndex({
      documentId: this.workbookId,
      sheetIds: params.sheetIds,
      records: params.records,
    });
    this.dlpIndexCache.set(cacheKey, built);
    while (this.dlpIndexCache.size > DLP_INDEX_CACHE_MAX_ENTRIES) {
      const oldestKey = this.dlpIndexCache.keys().next().value as string | undefined;
      if (oldestKey === undefined) break;
      this.dlpIndexCache.delete(oldestKey);
    }
    return built;
  }

  private startRequest(params: {
    cacheKey: string;
    functionName: string;
    prompt: string;
    inputs: unknown;
    inputsHash: string;
    inputsCompaction: unknown;
    cellAddress?: string;
    references: unknown;
    dlp: { decision: any; selectionClassification: any; redactedCount: number };
  }): void {
    const request = this.runRequest(params).finally(() => {
      this.inFlightByKey.delete(params.cacheKey);
    });
    this.inFlightByKey.set(params.cacheKey, request);
  }

  private async runRequest(params: {
    cacheKey: string;
    functionName: string;
    prompt: string;
    inputs: unknown;
    inputsHash: string;
    inputsCompaction: unknown;
    cellAddress?: string;
    references: unknown;
    dlp: { decision: any; selectionClassification: any; redactedCount: number };
  }): Promise<void> {
    const auditInput: any = {
      function: params.functionName,
      prompt: truncateText(params.prompt, this.maxAuditPreviewChars),
      prompt_hash: hashText(params.prompt),
      inputs_hash: params.inputsHash,
      inputs_compaction: params.inputsCompaction,
      cell: params.cellAddress,
      workbookId: this.workbookId,
      references: params.references,
      dlp: {
        decision: params.dlp.decision,
        selectionClassification: params.dlp.selectionClassification,
        redactedCount: params.dlp.redactedCount,
      },
    };

    const recorder = new AIAuditRecorder({
      store: this.auditStore,
      session_id: this.sessionId,
      workbook_id: this.workbookId,
      user_id: this.userId,
      mode: "cell_function",
      input: auditInput,
      model: this.model,
    });

    const abortController = typeof AbortController !== "undefined" ? new AbortController() : null;
    const timeoutMs = this.requestTimeoutMs;
    const timeoutError = new Error(`AI cell function request timed out after ${timeoutMs}ms`);
    let didTimeout = false;
    let timeoutId: ReturnType<typeof setTimeout> | null = null;

    try {
      let started = 0;
      const response = await this.requestLimiter.run((release) => {
        started = nowMs();

        const chatPromise = this.llmClient.chat({
          model: this.model,
          messages: buildMessages({
            functionName: params.functionName,
            prompt: params.prompt,
            inputs: params.inputs,
            maxPromptChars: this.maxUserMessageChars,
          }),
          ...(abortController ? { signal: abortController.signal } : {}),
        });
        // Release concurrency slots as soon as the underlying chat promise settles so
        // queued requests can start in the same microtask flush.
        chatPromise.finally(release).catch(() => undefined);

        const timeoutPromise = new Promise<never>((_resolve, reject) => {
          timeoutId = setTimeout(() => {
            didTimeout = true;
            try {
              abortController?.abort();
            } catch {
              // ignore
            }
            release();
            reject(timeoutError);
          }, timeoutMs);
        });

        return Promise.race([chatPromise, timeoutPromise]) as any;
      });
      if (timeoutId != null) {
        clearTimeout(timeoutId);
        timeoutId = null;
      }
      recorder.recordModelLatency(nowMs() - started);

      const promptTokens = response.usage?.promptTokens;
      const completionTokens = response.usage?.completionTokens;
      if (typeof promptTokens === "number" || typeof completionTokens === "number") {
        recorder.recordTokenUsage({
          prompt_tokens: typeof promptTokens === "number" ? promptTokens : 0,
          completion_tokens: typeof completionTokens === "number" ? completionTokens : 0,
        });
      }

      const rawContent = String(response.message?.content ?? "");
      const sanitized = sanitizeCellText(rawContent);

      // Record bounded output metadata (never store the full output string in audit entries).
      auditInput.output_hash = hashText(sanitized);
      auditInput.output_preview = truncateText(sanitized, this.maxAuditPreviewChars);
      auditInput.output_chars = sanitized.length;

      const normalizedValue = normalizeModelOutputForCellValue(sanitized, { maxStringChars: this.maxOutputChars });
      auditInput.output_type = normalizedValue === null ? "null" : typeof normalizedValue;
      if (typeof normalizedValue !== "string") auditInput.output_value = normalizedValue;

      this.writeCache(params.cacheKey, normalizedValue);
    } catch (error) {
      const finalError = didTimeout ? timeoutError : error;
      auditInput.error = finalError instanceof Error ? finalError.message : String(finalError);
      if (didTimeout) recorder.setUserFeedback("rejected");
      this.writeCache(params.cacheKey, AI_CELL_ERROR);
    } finally {
      if (timeoutId != null) clearTimeout(timeoutId);
      await recorder.finalize();
      this.onUpdate?.();
    }
  }

  private trackPendingAudit(promise: Promise<void>): void {
    // Prevent unhandled rejection warnings in tests if an audit write fails.
    promise.catch(() => undefined);
    this.pendingAudits.add(promise);
    promise.finally(() => {
      this.pendingAudits.delete(promise);
    });
  }

  private async auditBlockedRun(params: {
    functionName: string;
    prompt: string;
    inputsHash: string;
    inputsCompaction: unknown;
    cellAddress?: string;
    references: unknown;
    dlp: { decision: any; selectionClassification: any; redactedCount: number };
  }): Promise<void> {
    const recorder = new AIAuditRecorder({
      store: this.auditStore,
      session_id: this.sessionId,
      workbook_id: this.workbookId,
      user_id: this.userId,
      mode: "cell_function",
      model: this.model,
      input: {
        function: params.functionName,
        prompt: truncateText(params.prompt, this.maxAuditPreviewChars),
        prompt_hash: hashText(params.prompt),
        inputs_hash: params.inputsHash,
        inputs_compaction: params.inputsCompaction,
        cell: params.cellAddress,
        workbookId: this.workbookId,
        references: params.references,
        dlp: params.dlp,
        blocked: true,
      },
    });
    await recorder.finalize();
  }

  private writeCache(cacheKey: string, value: SpreadsheetValue): void {
    this.cache.set(cacheKey, { value, updatedAtMs: Date.now() });
    while (this.cache.size > this.cacheMaxEntries) {
      const oldestKey = this.cache.keys().next().value as string | undefined;
      if (oldestKey === undefined) break;
      this.cache.delete(oldestKey);
    }
    this.scheduleCachePersistence();
  }

  private loadCacheFromStorage(): void {
    if (!this.cachePersistKey) return;
    const storage = getLocalStorageOrNull();
    if (!storage) return;
    try {
      const raw = storage.getItem(this.cachePersistKey);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return;
      for (const entry of parsed) {
        if (!entry || typeof entry !== "object") continue;
        const key = (entry as any).key;
        const value = (entry as any).value;
        const updatedAtMs = (entry as any).updatedAtMs;
        if (typeof key !== "string") continue;
        const normalizedValue = normalizePersistedCellValue(value, this.maxOutputChars);
        if (normalizedValue === undefined) continue;
        // Cache key format has evolved over time. We currently expect:
        //   `${model}\0${function}\0${promptHash}\0${inputsHash}`
        // where both hashes are FNV-1a hex digests (either 8-hex legacy or 16-hex current).
        //
        // Drop legacy keys that embed the raw prompt text (can be large / sensitive).
        if (!isSupportedCacheKey(key)) continue;
        this.cache.set(key, {
          value: normalizedValue,
          updatedAtMs: typeof updatedAtMs === "number" ? updatedAtMs : Date.now(),
        });
      }
    } catch {
      // Ignore corrupt caches.
    }
  }

  private saveCacheToStorage(): void {
    if (!this.cachePersistKey) return;
    const storage = getLocalStorageOrNull();
    if (!storage) return;
    try {
      const entries = Array.from(this.cache.entries()).map(([key, entry]) => ({ key, ...entry }));
      storage.setItem(this.cachePersistKey, JSON.stringify(entries));
    } catch {
      // Ignore persistence failures.
    }
  }

  private ensureCachePersistPromise(): Promise<void> {
    if (this.cachePersistPromise) return this.cachePersistPromise;
    this.cachePersistPromise = new Promise<void>((resolve) => {
      this.cachePersistPromiseResolve = resolve;
    });
    return this.cachePersistPromise;
  }

  private scheduleCachePersistence(): void {
    if (!this.cachePersistKey) return;
    if (!getLocalStorageOrNull()) return;

    this.ensureCachePersistPromise();

    if (this.cachePersistTimer) {
      clearTimeout(this.cachePersistTimer);
      this.cachePersistTimer = null;
    }

    this.cachePersistTimer = setTimeout(() => {
      this.flushCachePersistenceNow();
    }, CACHE_PERSIST_DEBOUNCE_MS);
  }

  private flushCachePersistenceNow(): void {
    if (this.cachePersistTimer) {
      clearTimeout(this.cachePersistTimer);
      this.cachePersistTimer = null;
    }

    const resolve = this.cachePersistPromiseResolve;
    const promise = this.cachePersistPromise;
    if (!promise) return;

    try {
      this.saveCacheToStorage();
    } finally {
      this.cachePersistPromise = null;
      this.cachePersistPromiseResolve = null;
      resolve?.();
    }
  }
}

function isSupportedCacheKey(key: string): boolean {
  const parts = key.split("\u0000");
  if (parts.length !== 4) return false;
  const promptHash = parts[2] ?? "";
  const inputsHash = parts[3] ?? "";
  const isSupportedHash = (value: string) => /^(?:[0-9a-f]{8}|[0-9a-f]{16})$/.test(value);
  return isSupportedHash(promptHash) && isSupportedHash(inputsHash);
}

type AiCellFunctionReferences = {
  args: Array<{ cells: string[]; ranges: string[] }>;
  cells: string[];
  ranges: string[];
};

function sheetIdFromCellAddress(cellAddress: string | undefined): string {
  if (!cellAddress) return "Sheet1";
  const bang = cellAddress.indexOf("!");
  if (bang === -1) return "Sheet1";
  const sheet = cellAddress.slice(0, bang).trim();
  return sheet || "Sheet1";
}

function splitSheetQualifier(input: string): { sheetId: string | null; ref: string } {
  const s = String(input).trim();

  const quoted = s.match(/^'((?:[^']|'')+)'!(.+)$/);
  if (quoted) {
    return { sheetId: quoted[1].replace(/''/g, "'"), ref: quoted[2] };
  }

  const unquoted = s.match(/^([^!]+)!(.+)$/);
  if (unquoted) return { sheetId: unquoted[1], ref: unquoted[2] };

  return { sheetId: null, ref: s };
}

function sheetDisplayName(sheetId: string, sheetNameResolver: SheetNameResolver | null): string {
  const id = String(sheetId ?? "").trim();
  if (!id) return "";
  return sheetNameResolver?.getSheetNameById(id) ?? id;
}

type ParsedProvenanceRef = { sheetId: string; range: RangeAddress; canonical: string; isCell: boolean };

function parseProvenanceRef(
  refText: string,
  defaultSheetId: string,
  sheetNameResolver: SheetNameResolver | null,
): ParsedProvenanceRef | null {
  const cleaned = String(refText).replaceAll("$", "").trim();
  if (!cleaned) return null;
  const { sheetId: sheetFromRef, ref } = splitSheetQualifier(cleaned);
  const rawSheet = (sheetFromRef ?? defaultSheetId).trim();
  // Only resolve sheet-qualified refs. The default sheet id (from the caller's `cellAddress`)
  // is expected to already be a stable id and should not be reinterpreted as a display name.
  const sheetId = sheetFromRef ? sheetNameResolver?.getSheetIdByName(rawSheet) ?? rawSheet : rawSheet;
  if (!sheetId) return null;

  const range = parseA1Range(ref);
  if (!range) return null;

  const a1 = rangeToA1(range);
  const canonical = `${sheetId}!${a1}`;
  const isCell = range.start.row === range.end.row && range.start.col === range.end.col;
  return { sheetId, range, canonical, isCell };
}

function mergeProvenance(a: AiFunctionArgumentProvenance, b: AiFunctionArgumentProvenance): AiFunctionArgumentProvenance {
  return {
    cells: Array.from(new Set([...(a.cells ?? []), ...(b.cells ?? [])])).sort(),
    ranges: Array.from(new Set([...(a.ranges ?? []), ...(b.ranges ?? [])])).sort(),
  };
}

function rangeRefFromArray(value: CellValue): string | null {
  if (!Array.isArray(value)) return null;
  const ref = (value as any).__rangeRef;
  return typeof ref === "string" && ref.trim() ? ref.trim() : null;
}

function inferProvenanceFromValue(params: {
  value: CellValue;
  documentId: string;
  defaultSheetId: string;
  functionName: string;
  argIndex: number;
}): AiFunctionArgumentProvenance {
  const value = params.value;
  const cells = new Set<string>();
  const ranges = new Set<string>();

  const addRef = (refText: string): void => {
    const cleaned = String(refText).replaceAll("$", "").trim();
    if (!cleaned) return;
    if (cleaned.includes(":")) ranges.add(cleaned);
    else cells.add(cleaned);
  };

  if (isProvenanceCellValue(value)) {
    for (const ref of String(value.__cellRef).split(PROVENANCE_REF_SEPARATOR)) addRef(ref);
    return { cells: Array.from(cells), ranges: Array.from(ranges) };
  }

  const rangeRef = rangeRefFromArray(value);
  if (rangeRef) {
    ranges.add(rangeRef);
    return { cells: [], ranges: Array.from(ranges) };
  }

  if (Array.isArray(value)) {
    const arr = value as any[];
    const total = arr.length;

    // Avoid scanning arbitrarily large arrays. For arrays without a `__rangeRef`, we
    // sample the same preview + sample indices we include in the prompt compaction
    // so DLP policy decisions cover exactly the values we might send to the model.
    const seedHex = hashText(`${params.documentId}:${params.defaultSheetId}:${params.functionName}:${params.argIndex}:array`);
    const rand = mulberry32(seedFromHashHex(seedHex));
    const previewCount = Math.min(MAX_RANGE_PREVIEW_VALUES, total);
    const previewIndices = new Set<number>();
    for (let i = 0; i < previewCount; i += 1) previewIndices.add(i);
    const availableForSample = Math.max(0, total - previewCount);
    const sampleCount = Math.min(MAX_RANGE_SAMPLE_VALUES, availableForSample);
    const sampleIndices = pickSampleIndices({ total, count: sampleCount, rand, exclude: previewIndices });

    const visit = (index: number): void => {
      const entry = arr[index];
      if (!isProvenanceCellValue(entry)) return;
      for (const ref of String(entry.__cellRef).split(PROVENANCE_REF_SEPARATOR)) addRef(ref);
    };

    for (let i = 0; i < previewCount; i += 1) visit(i);
    for (const idx of sampleIndices) visit(idx);
  }

  return { cells: Array.from(cells), ranges: Array.from(ranges) };
}

function normalizeProvenance(
  entry: AiFunctionArgumentProvenance | undefined,
  defaultSheetId: string,
  sheetNameResolver: SheetNameResolver | null,
): AiFunctionArgumentProvenance {
  const cells = new Set<string>();
  const ranges = new Set<string>();

  const add = (raw: string): void => {
    const parsed = parseProvenanceRef(raw, defaultSheetId, sheetNameResolver);
    if (!parsed) return;
    if (parsed.isCell) cells.add(parsed.canonical);
    else ranges.add(parsed.canonical);
  };

  for (const c of entry?.cells ?? []) add(c);
  for (const r of entry?.ranges ?? []) add(r);

  return { cells: Array.from(cells).sort(), ranges: Array.from(ranges).sort() };
}

function alignArgProvenance(params: {
  args: CellValue[];
  provenance: AiFunctionArgumentProvenance[] | undefined;
  defaultSheetId: string;
  documentId: string;
  functionName: string;
  sheetNameResolver: SheetNameResolver | null;
}): AiFunctionArgumentProvenance[] {
  const out: AiFunctionArgumentProvenance[] = [];
  for (let i = 0; i < params.args.length; i += 1) {
    const fromCaller = normalizeProvenance(params.provenance?.[i], params.defaultSheetId, params.sheetNameResolver);
    const fromValue = normalizeProvenance(
      inferProvenanceFromValue({
        value: params.args[i]!,
        documentId: params.documentId,
        defaultSheetId: params.defaultSheetId,
        functionName: params.functionName,
        argIndex: i,
      }),
      params.defaultSheetId,
      params.sheetNameResolver,
    );
    out.push(mergeProvenance(fromCaller, fromValue));
  }
  return out;
}

function summarizeReferences(provenance: AiFunctionArgumentProvenance[]): AiCellFunctionReferences {
  const cells = new Set<string>();
  const ranges = new Set<string>();
  const args = provenance.map((arg) => {
    const argCells = Array.from(new Set(arg.cells ?? [])).filter(Boolean);
    const argRanges = Array.from(new Set(arg.ranges ?? [])).filter(Boolean);
    for (const c of argCells) cells.add(c);
    for (const r of argRanges) ranges.add(r);
    return { cells: argCells, ranges: argRanges };
  });
  return { args, cells: Array.from(cells).sort(), ranges: Array.from(ranges).sort() };
}

function computeSelectionClassification(params: {
  documentId: string;
  defaultSheetId: string;
  args: CellValue[];
  provenance: AiFunctionArgumentProvenance[];
  maxCellChars: number;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationIndex: DlpCellIndex | null;
  sheetNameResolver: SheetNameResolver | null;
}): { selectionClassification: any; references: AiCellFunctionReferences } {
  const references = summarizeReferences(params.provenance);

  let selectionClassification = { ...DEFAULT_CLASSIFICATION };

  // 1) Heuristic classification for literal arguments (prompt/input strings typed directly into formulas).
  for (let i = 0; i < params.args.length; i += 1) {
    const prov = params.provenance[i];
    const hasRefs = Boolean((prov?.cells?.length ?? 0) > 0 || (prov?.ranges?.length ?? 0) > 0);
    if (hasRefs) continue;
    selectionClassification = maxClassification(selectionClassification, heuristicClassifyValue(params.args[i], params.maxCellChars));
    if (selectionClassification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
  }

  // 2) Store-backed provenance classification (enterprise DLP).
  for (const cellRef of references.cells) {
    const parsed = parseProvenanceRef(cellRef, params.defaultSheetId, params.sheetNameResolver);
    if (!parsed || !parsed.isCell) continue;
    const classification = params.classificationIndex
      ? effectiveCellClassificationFromIndex(params.classificationIndex, {
          documentId: params.documentId,
          sheetId: parsed.sheetId,
          row: parsed.range.start.row,
          col: parsed.range.start.col,
        })
      : effectiveCellClassification(
          { documentId: params.documentId, sheetId: parsed.sheetId, row: parsed.range.start.row, col: parsed.range.start.col } as any,
          params.classificationRecords,
        );
    selectionClassification = maxClassification(selectionClassification, classification);
    if (selectionClassification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
  }

  if (selectionClassification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    for (const rangeRef of references.ranges) {
      const parsed = parseProvenanceRef(rangeRef, params.defaultSheetId, params.sheetNameResolver);
      if (!parsed || parsed.isCell) continue;
      const classification = effectiveRangeClassification(
        { documentId: params.documentId, sheetId: parsed.sheetId, range: parsed.range } as any,
        params.classificationRecords,
      );
      selectionClassification = maxClassification(selectionClassification, classification);
      if (selectionClassification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
    }
  }

  return { selectionClassification: normalizeClassification(selectionClassification), references };
}

function heuristicClassifyValue(value: CellValue, maxChars: number): any {
  if (Array.isArray(value)) return { ...DEFAULT_CLASSIFICATION };
  const scalar = isProvenanceCellValue(value) ? value.value : (value as SpreadsheetValue);
  const text = formatScalar(scalar, { maxChars }).toLowerCase();
  if (!text) return { ...DEFAULT_CLASSIFICATION };

  const labels = ["heuristic"];
  if (text.includes("password") || text.includes("ssn")) return { level: CLASSIFICATION_LEVEL.RESTRICTED, labels };
  if (text.includes("top secret") || text.includes("secret")) return { level: CLASSIFICATION_LEVEL.RESTRICTED, labels };
  if (text.includes("confidential")) return { level: CLASSIFICATION_LEVEL.CONFIDENTIAL, labels };
  if (text.includes("internal")) return { level: CLASSIFICATION_LEVEL.INTERNAL, labels };
  return { ...DEFAULT_CLASSIFICATION };
}

function preparePromptAndInputs(params: {
  functionName: string;
  args: CellValue[];
  provenance: AiFunctionArgumentProvenance[];
  decision: any;
  maxAllowedRank: number | null;
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationIndex: DlpCellIndex | null;
  maxCellChars: number;
  maxLiteralChars: number;
  sheetNameResolver: SheetNameResolver | null;
}): {
  prompt: string;
  inputs: unknown;
  inputsHash: string;
  inputsCompaction: unknown;
  redactedCount: number;
} {
  const shouldRedact = params.decision?.decision !== DLP_DECISION.ALLOW;

  let redactedCount = 0;

  const promptResult = compactArgForPrompt({
    functionName: params.functionName,
    argIndex: 0,
    value: params.args[0] ?? null,
    provenance: params.provenance[0] ?? { cells: [], ranges: [] },
    shouldRedact,
    maxAllowedRank: params.maxAllowedRank,
    documentId: params.documentId,
    defaultSheetId: params.defaultSheetId,
    policy: params.policy,
    classificationRecords: params.classificationRecords,
    classificationIndex: params.classificationIndex,
    maxCellChars: params.maxCellChars,
    maxLiteralChars: params.maxLiteralChars,
    sheetNameResolver: params.sheetNameResolver,
  });
  redactedCount += promptResult.redactedCount;

  const promptRaw = typeof promptResult.value === "string" ? promptResult.value : stableJsonStringify(promptResult.value);
  const prompt = truncateText(promptRaw, MAX_PROMPT_CHARS);

  const inputResults = params.args.slice(1).map((arg, idx) =>
    compactArgForPrompt({
      functionName: params.functionName,
      argIndex: idx + 1,
      value: arg,
      provenance: params.provenance[idx + 1] ?? { cells: [], ranges: [] },
      shouldRedact,
      maxAllowedRank: params.maxAllowedRank,
      documentId: params.documentId,
      defaultSheetId: params.defaultSheetId,
      policy: params.policy,
      classificationRecords: params.classificationRecords,
      classificationIndex: params.classificationIndex,
      maxCellChars: params.maxCellChars,
      maxLiteralChars: params.maxLiteralChars,
      sheetNameResolver: params.sheetNameResolver,
    }),
  );

  for (const result of inputResults) redactedCount += result.redactedCount;

  const compactedInputs: unknown =
    inputResults.length === 0 ? null : inputResults.length === 1 ? inputResults[0].value : inputResults.map((r) => r.value);

  const inputsCompaction: unknown =
    inputResults.length === 0
      ? null
      : inputResults.length === 1
        ? inputResults[0].compaction
        : inputResults.map((r) => r.compaction);

  const inputsHash = hashText(stableJsonStringify(compactedInputs));
  return { prompt, inputs: compactedInputs, inputsHash, inputsCompaction, redactedCount };
}

function compactArgForPrompt(params: {
  functionName: string;
  argIndex: number;
  value: CellValue;
  provenance: AiFunctionArgumentProvenance;
  shouldRedact: boolean;
  maxAllowedRank: number | null;
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationIndex: DlpCellIndex | null;
  maxCellChars: number;
  maxLiteralChars: number;
  sheetNameResolver: SheetNameResolver | null;
}): { value: unknown; compaction: unknown; redactedCount: number } {
  if (Array.isArray(params.value)) {
    return compactArrayForPrompt({
      functionName: params.functionName,
      argIndex: params.argIndex,
      entries: params.value as Array<SpreadsheetValue | ProvenanceCellValue>,
      rangeRef: rangeRefFromArray(params.value),
      provenance: params.provenance,
      shouldRedact: params.shouldRedact,
      maxAllowedRank: params.maxAllowedRank,
      documentId: params.documentId,
      defaultSheetId: params.defaultSheetId,
      policy: params.policy,
      classificationRecords: params.classificationRecords,
      classificationIndex: params.classificationIndex,
      maxCellChars: params.maxCellChars,
      sheetNameResolver: params.sheetNameResolver,
    });
  }

  const scalarValue = isProvenanceCellValue(params.value) ? params.value.value : (params.value as SpreadsheetValue);
  return compactScalarForPrompt({
    value: scalarValue,
    provenance: params.provenance,
    shouldRedact: params.shouldRedact,
    maxAllowedRank: params.maxAllowedRank,
    documentId: params.documentId,
    defaultSheetId: params.defaultSheetId,
    policy: params.policy,
    classificationRecords: params.classificationRecords,
    classificationIndex: params.classificationIndex,
    maxCellChars: params.maxCellChars,
    maxLiteralChars: params.maxLiteralChars,
    sheetNameResolver: params.sheetNameResolver,
  });
}

function classificationForProvenance(params: {
  documentId: string;
  defaultSheetId: string;
  provenance: AiFunctionArgumentProvenance;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationIndex: DlpCellIndex | null;
  sheetNameResolver: SheetNameResolver | null;
}): any {
  let classification = { ...DEFAULT_CLASSIFICATION };

  const visit = (refText: string): void => {
    const parsed = parseProvenanceRef(refText, params.defaultSheetId, params.sheetNameResolver);
    if (!parsed) return;

    if (parsed.isCell) {
      classification = maxClassification(
        classification,
        params.classificationIndex
          ? effectiveCellClassificationFromIndex(params.classificationIndex, {
              documentId: params.documentId,
              sheetId: parsed.sheetId,
              row: parsed.range.start.row,
              col: parsed.range.start.col,
            })
          : effectiveCellClassification(
              { documentId: params.documentId, sheetId: parsed.sheetId, row: parsed.range.start.row, col: parsed.range.start.col } as any,
              params.classificationRecords,
            ),
      );
      return;
    }

    classification = maxClassification(
      classification,
      effectiveRangeClassification({ documentId: params.documentId, sheetId: parsed.sheetId, range: parsed.range } as any, params.classificationRecords),
    );
  };

  for (const cellRef of params.provenance.cells ?? []) visit(cellRef);
  for (const rangeRef of params.provenance.ranges ?? []) visit(rangeRef);
  return classification;
}

function compactScalarForPrompt(params: {
  value: SpreadsheetValue;
  provenance: AiFunctionArgumentProvenance;
  shouldRedact: boolean;
  maxAllowedRank: number | null;
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationIndex: DlpCellIndex | null;
  maxCellChars: number;
  maxLiteralChars: number;
  sheetNameResolver: SheetNameResolver | null;
}): { value: string; compaction: unknown; redactedCount: number } {
  const hasRefs = Boolean((params.provenance.cells?.length ?? 0) > 0 || (params.provenance.ranges?.length ?? 0) > 0);
  const maxChars = hasRefs ? params.maxCellChars : params.maxLiteralChars;

  // Fast path: when the overall AI request is allowed, we never redact individual scalar
  // values and don't need to compute per-argument DLP classifications.
  if (!params.shouldRedact) {
    return {
      value: formatScalar(params.value, { maxChars }),
      compaction: { kind: "scalar", redacted: false },
      redactedCount: 0,
    };
  }

  let classification = { ...DEFAULT_CLASSIFICATION };
  if (hasRefs) {
    classification = classificationForProvenance({
      documentId: params.documentId,
      defaultSheetId: params.defaultSheetId,
      provenance: params.provenance,
      classificationRecords: params.classificationRecords,
      classificationIndex: params.classificationIndex,
      sheetNameResolver: params.sheetNameResolver,
    });
  } else {
    classification = heuristicClassifyValue(params.value as any, params.maxLiteralChars);
  }

  const allowed = params.maxAllowedRank !== null && classificationRank(classification.level) <= params.maxAllowedRank;
  if (params.shouldRedact && !allowed) {
    return { value: DLP_REDACTION_PLACEHOLDER, compaction: { kind: "scalar", redacted: true }, redactedCount: 1 };
  }

  return {
    value: formatScalar(params.value, { maxChars }),
    compaction: { kind: "scalar", redacted: false },
    redactedCount: 0,
  };
}

function compactArrayForPrompt(params: {
  functionName: string;
  argIndex: number;
  entries: Array<SpreadsheetValue | ProvenanceCellValue>;
  rangeRef: string | null;
  provenance: AiFunctionArgumentProvenance;
  shouldRedact: boolean;
  maxAllowedRank: number | null;
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  classificationIndex: DlpCellIndex | null;
  maxCellChars: number;
  sheetNameResolver: SheetNameResolver | null;
}): { value: unknown; compaction: unknown; redactedCount: number } {
  const providedCount = params.entries.length;

  // Avoid materializing copies of large arrays; we only access the handful of indices
  // that we include in the prompt compaction.
  const getEntryAt = (index: number): SpreadsheetValue | ProvenanceCellValue | null => {
    if (index < 0 || index >= providedCount) return null;
    return (params.entries[index] ?? null) as any;
  };

  const getCellRefAt = (index: number): string | null => {
    const entry = getEntryAt(index);
    if (!entry || !isProvenanceCellValue(entry)) return null;
    const ref = String(entry.__cellRef ?? "").trim();
    return ref ? ref : null;
  };

  const hasPerCellRefs = (() => {
    const scanCount = Math.min(providedCount, 5);
    for (let i = 0; i < scanCount; i += 1) {
      if (getCellRefAt(i)) return true;
    }
    return false;
  })();

  const rangeCandidate = params.provenance.ranges?.length === 1 ? params.provenance.ranges[0]! : params.rangeRef;
  const parsedRange = rangeCandidate ? parseProvenanceRef(rangeCandidate, params.defaultSheetId, params.sheetNameResolver) : null;
  const isRange = Boolean(parsedRange && !parsedRange.isCell);

  const rangeSheetId = parsedRange?.sheetId ?? null;
  const range = isRange ? parsedRange!.range : null;
  const cols = range ? Math.max(1, range.end.col - range.start.col + 1) : null;
  const rows = range ? Math.max(1, range.end.row - range.start.row + 1) : null;
  const totalCells = range && rows && cols ? rows * cols : providedCount;
  const truncated = totalCells > providedCount;
  const countMeta = {
    total_cells: totalCells,
    sampled_cells: providedCount,
    ...(truncated ? { truncated: true } : {}),
  };

  const rangeA1 = range ? rangeToA1(range) : null;
  // Use stable ids for determinism (sampling) but display names in prompts/audit logs.
  // This avoids leaking internal sheet ids to the model while keeping sampling stable across renames.
  const rangeStableText = rangeA1 && rangeSheetId ? `${rangeSheetId}!${rangeA1}` : rangeA1;
  const rangeDisplayText =
    rangeA1 && rangeSheetId
      ? `${formatSheetNameForA1(sheetDisplayName(rangeSheetId, params.sheetNameResolver) || rangeSheetId)}!${rangeA1}`
      : rangeA1;

  const seedHex = hashText(
    `${params.documentId}:${rangeSheetId ?? params.defaultSheetId}:${params.functionName}:${params.argIndex}:${rangeStableText ?? "array"}`,
  );
  const rand = mulberry32(seedFromHashHex(seedHex));

  const previewCount = Math.min(MAX_RANGE_PREVIEW_VALUES, providedCount);
  const previewIndices = new Set<number>();
  for (let i = 0; i < previewCount; i += 1) previewIndices.add(i);

  const availableForSample = Math.max(0, providedCount - previewCount);
  const sampleCount = Math.min(MAX_RANGE_SAMPLE_VALUES, availableForSample);
  const sampleIndices = pickSampleIndices({ total: providedCount, count: sampleCount, rand, exclude: previewIndices });

  const rawHeaderCount = range && cols && cols > 1 ? Math.min(cols, MAX_RANGE_HEADER_VALUES) : 0;
  const headerCount = Math.min(rawHeaderCount, providedCount);
  const headerIndices = headerCount > 0 ? Array.from({ length: headerCount }, (_v, i) => i) : [];

  const stats = { numericCount: 0, min: 0, max: 0, sum: 0 };
  const numericSeen = new Set<number>();
  const redactedSeen = new Set<number>();
  let redactedCount = 0;

  const classificationForCellRef = (refText: string): any => {
    let classification = { ...DEFAULT_CLASSIFICATION };
    for (const ref of String(refText).split(PROVENANCE_REF_SEPARATOR)) {
      const parsed = parseProvenanceRef(ref, params.defaultSheetId, params.sheetNameResolver);
      if (!parsed) continue;

      if (parsed.isCell) {
        classification = maxClassification(
          classification,
          params.classificationIndex
            ? effectiveCellClassificationFromIndex(params.classificationIndex, {
                documentId: params.documentId,
                sheetId: parsed.sheetId,
                row: parsed.range.start.row,
                col: parsed.range.start.col,
              })
            : effectiveCellClassification(
                { documentId: params.documentId, sheetId: parsed.sheetId, row: parsed.range.start.row, col: parsed.range.start.col } as any,
                params.classificationRecords,
              ),
        );
      } else {
        classification = maxClassification(
          classification,
          effectiveRangeClassification(
            { documentId: params.documentId, sheetId: parsed.sheetId, range: parsed.range } as any,
            params.classificationRecords,
          ),
        );
      }

      if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
    }
    return classification;
  };

  const shouldRedactAll =
    params.shouldRedact &&
    !range &&
    !hasPerCellRefs &&
    ((params.provenance.cells?.length ?? 0) > 0 || (params.provenance.ranges?.length ?? 0) > 0) &&
    (() => {
      const classification = classificationForProvenance({
        documentId: params.documentId,
        defaultSheetId: params.defaultSheetId,
        provenance: params.provenance,
        classificationRecords: params.classificationRecords,
        classificationIndex: params.classificationIndex,
        sheetNameResolver: params.sheetNameResolver,
      });
      return params.maxAllowedRank === null || classificationRank(classification.level) > params.maxAllowedRank;
    })();

  const recordNumeric = (n: number): void => {
    if (stats.numericCount === 0) {
      stats.min = n;
      stats.max = n;
      stats.sum = n;
      stats.numericCount = 1;
      return;
    }
    stats.numericCount += 1;
    stats.min = Math.min(stats.min, n);
    stats.max = Math.max(stats.max, n);
    stats.sum += n;
  };

  const formatAt = (index: number): string => {
    const entry = getEntryAt(index);
    const raw = entry && isProvenanceCellValue(entry) ? entry.value : (entry as SpreadsheetValue);
    const cellRef = entry && isProvenanceCellValue(entry) ? String(entry.__cellRef ?? "").trim() : null;

    if (params.shouldRedact) {
      if (shouldRedactAll) {
        if (!redactedSeen.has(index)) {
          redactedSeen.add(index);
          redactedCount += 1;
        }
        return DLP_REDACTION_PLACEHOLDER;
      }

      if (cellRef) {
        const classification = classificationForCellRef(cellRef);
        const allowed = params.maxAllowedRank !== null && classificationRank(classification.level) <= params.maxAllowedRank;
        if (!allowed) {
          if (!redactedSeen.has(index)) {
            redactedSeen.add(index);
            redactedCount += 1;
          }
          return DLP_REDACTION_PLACEHOLDER;
        }
      } else
      // If we can map indices -> cells, enforce per-cell redaction.
      if (range && cols && rangeSheetId) {
        const rowOffset = Math.floor(index / cols);
        const colOffset = index % cols;
        const row = range.start.row + rowOffset;
        const col = range.start.col + colOffset;

        const classification = params.classificationIndex
          ? effectiveCellClassificationFromIndex(params.classificationIndex, { documentId: params.documentId, sheetId: rangeSheetId, row, col })
          : effectiveCellClassification(
              { documentId: params.documentId, sheetId: rangeSheetId, row, col } as any,
              params.classificationRecords,
            );
        const allowed = params.maxAllowedRank !== null && classificationRank(classification.level) <= params.maxAllowedRank;
        if (!allowed) {
          if (!redactedSeen.has(index)) {
            redactedSeen.add(index);
            redactedCount += 1;
          }
          return DLP_REDACTION_PLACEHOLDER;
        }
      }
    }

    const text = formatScalar(raw, { maxChars: params.maxCellChars });
    if (typeof raw === "number" && Number.isFinite(raw) && !numericSeen.has(index)) {
      numericSeen.add(index);
      recordNumeric(raw);
    }
    return text;
  };

  const header = headerIndices.length ? headerIndices.map((i) => formatAt(i)) : undefined;
  const preview = Array.from({ length: previewCount }, (_v, i) => formatAt(i));
  const sample = sampleIndices.length ? sampleIndices.map((i) => [i, formatAt(i)] as const) : undefined;

  const numericSummary =
    stats.numericCount > 0
      ? {
          count: stats.numericCount,
          min: stats.min,
          max: stats.max,
          mean: stats.sum / stats.numericCount,
        }
      : undefined;

  const shape = range
    ? {
        rows,
        cols,
        total_cells: totalCells,
        provided_cells: providedCount,
      }
    : { total_cells: providedCount };

  const value = {
    kind: rangeDisplayText ? "range" : "array",
    ...(rangeDisplayText ? { range: rangeDisplayText } : {}),
    ...countMeta,
    shape,
    ...(header ? { header } : {}),
    preview,
    ...(sample ? { sample: { seed: seedHex, values: sample } } : {}),
    ...(numericSummary ? { numeric_summary: numericSummary } : {}),
    ...(redactedCount ? { redacted_values: redactedCount } : {}),
  };

  const compaction = {
    kind: rangeDisplayText ? "range" : "array",
    ...(rangeDisplayText ? { range: rangeDisplayText } : {}),
    ...countMeta,
    shape,
    header_count: headerIndices.length,
    preview_count: previewCount,
    sample: sampleIndices.length ? { seed: seedHex, indices: sampleIndices } : null,
    redacted_values: redactedCount,
  };

  return { value, compaction, redactedCount };
}

function collectReferencedSheetIds(params: {
  provenance: AiFunctionArgumentProvenance[];
  defaultSheetId: string;
  sheetNameResolver: SheetNameResolver | null;
}): Set<string> {
  const sheetIds = new Set<string>();
  sheetIds.add(params.defaultSheetId);

  for (const prov of params.provenance ?? []) {
    for (const cellRef of prov?.cells ?? []) {
      for (const part of String(cellRef).split(PROVENANCE_REF_SEPARATOR)) {
        const parsed = parseProvenanceRef(part, params.defaultSheetId, params.sheetNameResolver);
        if (parsed?.sheetId) sheetIds.add(parsed.sheetId);
      }
    }
    for (const rangeRef of prov?.ranges ?? []) {
      for (const part of String(rangeRef).split(PROVENANCE_REF_SEPARATOR)) {
        const parsed = parseProvenanceRef(part, params.defaultSheetId, params.sheetNameResolver);
        if (parsed?.sheetId) sheetIds.add(parsed.sheetId);
      }
    }
  }

  return sheetIds;
}

/**
 * Test seam: allows spying on DLP index construction without relying on ESM export rewriting.
 *
 * `AiCellFunctionEngine` always calls this indirection when building indices.
 */
export const __dlpIndexBuilder = {
  buildDlpCellIndex,
};

function buildDlpCellIndex(params: {
  documentId: string;
  sheetIds: Set<string>;
  records: Array<{ selector: any; classification: any }>;
}): DlpCellIndex {
  let docClassificationMax: any = { ...DEFAULT_CLASSIFICATION };
  const sheetClassificationMaxBySheetId = new Map<string, any>();
  const columnClassificationBySheetId = new Map<string, Map<number, any>>();
  const cellClassificationBySheetId = new Map<string, Map<string, any>>();
  const rangeRecordsBySheetId = new Map<string, Array<{ range: DlpNormalizedRange; classification: any }>>();
  const fallbackRecordsBySheetId = new Map<string, Array<{ selector: any; classification: any }>>();

  function ensureSheetMax(sheetId: string): any {
    const existing = sheetClassificationMaxBySheetId.get(sheetId);
    if (existing) return existing;
    const init: any = { ...DEFAULT_CLASSIFICATION };
    sheetClassificationMaxBySheetId.set(sheetId, init);
    return init;
  }

  function ensureColMap(sheetId: string): Map<number, any> {
    const existing = columnClassificationBySheetId.get(sheetId);
    if (existing) return existing;
    const init = new Map<number, any>();
    columnClassificationBySheetId.set(sheetId, init);
    return init;
  }

  function ensureCellMap(sheetId: string): Map<string, any> {
    const existing = cellClassificationBySheetId.get(sheetId);
    if (existing) return existing;
    const init = new Map<string, any>();
    cellClassificationBySheetId.set(sheetId, init);
    return init;
  }

  function ensureRangeList(sheetId: string): Array<{ range: DlpNormalizedRange; classification: any }> {
    const existing = rangeRecordsBySheetId.get(sheetId);
    if (existing) return existing;
    const init: Array<{ range: DlpNormalizedRange; classification: any }> = [];
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
        if (typeof selector.sheetId !== "string" || !params.sheetIds.has(selector.sheetId)) break;
        const existing = ensureSheetMax(selector.sheetId);
        sheetClassificationMaxBySheetId.set(selector.sheetId, maxClassification(existing, record.classification));
        break;
      }
      case "column": {
        if (typeof selector.sheetId !== "string" || !params.sheetIds.has(selector.sheetId)) break;
        if (typeof selector.columnIndex === "number") {
          const colMap = ensureColMap(selector.sheetId);
          const existing = colMap.get(selector.columnIndex);
          colMap.set(selector.columnIndex, existing ? maxClassification(existing, record.classification) : record.classification);
        } else {
          // Table/columnId selectors require table metadata to evaluate; AiCellFunctionEngine
          // currently only operates on sheet coordinates (row/col) and has no table context,
          // so these selectors cannot apply and are ignored.
        }
        break;
      }
      case "cell": {
        if (typeof selector.sheetId !== "string" || !params.sheetIds.has(selector.sheetId)) break;
        if (typeof selector.row !== "number" || typeof selector.col !== "number") break;
        const key = `${selector.row},${selector.col}`;
        const cellMap = ensureCellMap(selector.sheetId);
        const existing = cellMap.get(key);
        cellMap.set(key, existing ? maxClassification(existing, record.classification) : record.classification);
        break;
      }
      case "range": {
        if (typeof selector.sheetId !== "string" || !params.sheetIds.has(selector.sheetId)) break;
        if (!selector.range) break;
        try {
          const normalized = normalizeRange(selector.range);
          ensureRangeList(selector.sheetId).push({ range: normalized, classification: record.classification });
        } catch {
          // Ignore invalid persisted ranges.
        }
        break;
      }
      default: {
        // Unknown selector scope: ignore.
        break;
      }
    }
  }

  return {
    documentId: params.documentId,
    docClassificationMax,
    sheetClassificationMaxBySheetId,
    columnClassificationBySheetId,
    cellClassificationBySheetId,
    rangeRecordsBySheetId,
    fallbackRecordsBySheetId,
  };
}

function cellInNormalizedRange(cell: { row: number; col: number }, range: DlpNormalizedRange): boolean {
  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.col >= range.start.col &&
    cell.col <= range.end.col
  );
}

function effectiveCellClassificationFromIndex(
  index: DlpCellIndex,
  cellRef: { documentId: string; sheetId: string; row: number; col: number },
): any {
  if (cellRef.documentId !== index.documentId) return { ...DEFAULT_CLASSIFICATION };

  let classification: any = { ...DEFAULT_CLASSIFICATION };

  classification = maxClassification(classification, index.docClassificationMax);
  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const sheetMax = index.sheetClassificationMaxBySheetId.get(cellRef.sheetId);
    if (sheetMax) classification = maxClassification(classification, sheetMax);
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const colMap = index.columnClassificationBySheetId.get(cellRef.sheetId);
    const colClassification = colMap?.get(cellRef.col);
    if (colClassification) classification = maxClassification(classification, colClassification);
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const cellMap = index.cellClassificationBySheetId.get(cellRef.sheetId);
    const cellClassification = cellMap?.get(`${cellRef.row},${cellRef.col}`);
    if (cellClassification) classification = maxClassification(classification, cellClassification);
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const rangeRecords = index.rangeRecordsBySheetId.get(cellRef.sheetId) ?? [];
    for (const record of rangeRecords) {
      if (!cellInNormalizedRange(cellRef, record.range)) continue;
      classification = maxClassification(classification, record.classification);
      if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
    }
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const fallbackRecords = index.fallbackRecordsBySheetId.get(cellRef.sheetId) ?? [];
    if (fallbackRecords.length > 0) {
      const fallback = effectiveCellClassification(cellRef as any, fallbackRecords as any);
      classification = maxClassification(classification, fallback);
    }
  }

  return classification;
}

function mulberry32(seed: number): () => number {
  let t = seed >>> 0;
  return () => {
    t += 0x6d2b79f5;
    let x = Math.imul(t ^ (t >>> 15), 1 | t);
    x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  };
}

function pickSampleIndices(params: { total: number; count: number; rand: () => number; exclude?: Set<number> }): number[] {
  const out = new Set<number>();
  const exclude = params.exclude ?? new Set<number>();
  const maxAttempts = Math.max(100, params.count * 20);
  let attempts = 0;

  while (out.size < params.count && attempts < maxAttempts) {
    attempts += 1;
    const idx = Math.floor(params.rand() * params.total);
    if (exclude.has(idx)) continue;
    out.add(idx);
  }

  return Array.from(out).sort((a, b) => a - b);
}

function buildMessages(params: {
  functionName: string;
  prompt: string;
  inputs: unknown;
  maxPromptChars: number;
}): LLMMessage[] {
  const system: LLMMessage = {
    role: "system",
    content: "You are an AI function inside a spreadsheet cell. Return ONLY the final cell value (no markdown, no extra explanation).",
  };

  const fn = params.functionName.toUpperCase();
  const prompt = params.prompt;
  const inputText = stableJsonStringify(params.inputs);

  let userContent = "";
  if (fn === "AI") {
    userContent = `Task: ${prompt}\n\nInput:\n${inputText}`.trim();
  } else if (fn === "AI.EXTRACT") {
    userContent = `Extract "${prompt}" from the following input. Return only the extracted value (or empty string).\n\n${inputText}`;
  } else if (fn === "AI.CLASSIFY") {
    userContent = `Classify the following input into one of these categories: ${prompt}. Return ONLY the category.\n\n${inputText}`;
  } else if (fn === "AI.TRANSLATE") {
    userContent = `Translate the following text into ${prompt}. Return ONLY the translated text.\n\n${inputText}`;
  } else {
    userContent = `${prompt}\n\n${inputText}`;
  }

  const user: LLMMessage = { role: "user", content: truncateText(userContent, params.maxPromptChars) };
  return [system, user];
}

function firstErrorCode(args: CellValue[]): string | null {
  for (const arg of args) {
    const err = firstErrorCodeInValue(arg);
    if (err) return err;
  }
  return null;
}

function firstErrorCodeInValue(value: CellValue): string | null {
  if (isErrorCode(value)) return value;
  if (isProvenanceCellValue(value)) {
    return isErrorCode(value.value) ? value.value : null;
  }
  if (Array.isArray(value)) {
    const arr = value as any[];
    const total = arr.length;
    const max = Math.min(total, DEFAULT_RANGE_SAMPLE_LIMIT);
    if (total <= max) {
      for (let i = 0; i < total; i += 1) {
        const err = firstErrorCodeInValue(arr[i] as any);
        if (err) return err;
      }
      return null;
    }

    // For very large arrays, scan a small deterministic prefix plus a seeded sample
    // from the remainder. This avoids O(N) scans on huge arrays.
    const prefixCount = Math.min(max, 50);
    for (let i = 0; i < prefixCount; i += 1) {
      const err = firstErrorCodeInValue(arr[i] as any);
      if (err) return err;
    }

    const remaining = max - prefixCount;
    if (remaining <= 0) return null;

    const exclude = new Set<number>();
    for (let i = 0; i < prefixCount; i += 1) exclude.add(i);
    const seedHex = hashText(`errors:${total}:${max}`);
    const rand = mulberry32(seedFromHashHex(seedHex));
    const sample = pickSampleIndices({ total, count: remaining, rand, exclude });
    for (const idx of sample) {
      const err = firstErrorCodeInValue(arr[idx] as any);
      if (err) return err;
    }
  }
  return null;
}

function isErrorCode(value: unknown): value is string {
  return isSpreadsheetErrorCode(value);
}

function isProvenanceCellValue(value: unknown): value is ProvenanceCellValue {
  if (!value || typeof value !== "object") return false;
  const v = value as any;
  return typeof v.__cellRef === "string" && "value" in v;
}

function sanitizeCellText(text: string): string {
  const trimmed = text.trim();
  if (!trimmed) return "";

  // Common "LLM helpfulness" wrappers we don't want in single-cell outputs.
  if (trimmed.startsWith("```") && trimmed.endsWith("```")) {
    return trimmed.slice(3, -3).trim();
  }

  return trimmed;
}

const NUMERIC_LITERAL_REGEX = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function normalizeModelOutputForCellValue(text: string, opts: { maxStringChars: number }): SpreadsheetValue {
  const sanitized = sanitizeCellText(text);
  const trimmed = sanitized.trim();
  if (!trimmed) return null;

  const upper = trimmed.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  if (upper === "NULL" || upper === "NONE") return null;

  if (NUMERIC_LITERAL_REGEX.test(trimmed)) {
    const num = Number(trimmed);
    if (Number.isFinite(num)) return num;
  }

  return truncateAiOutputText(trimmed, opts.maxStringChars);
}

function truncateAiOutputText(text: string, maxChars: number): string {
  const s = String(text);
  if (!Number.isFinite(maxChars) || maxChars <= 0) return "";
  if (s.length <= maxChars) return s;
  if (maxChars === 1) return "…";
  return `${s.slice(0, Math.max(0, maxChars - 1))}…`;
}

function normalizePersistedCellValue(value: unknown, maxStringChars: number): SpreadsheetValue | undefined {
  if (value === null) return null;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : undefined;
  if (typeof value === "string") return normalizeModelOutputForCellValue(value, { maxStringChars });
  return undefined;
}

function formatScalar(value: SpreadsheetValue, opts: { maxChars: number }): string {
  let text = "";
  if (value === null) text = "";
  else if (typeof value === "string") text = value;
  else if (typeof value === "number") text = Number.isFinite(value) ? String(value) : "";
  else if (typeof value === "boolean") text = value ? "TRUE" : "FALSE";
  else text = String(value);
  return truncateText(text, opts.maxChars);
}

function truncateText(text: string, maxChars: number): string {
  const s = String(text);
  if (!Number.isFinite(maxChars) || maxChars <= 0) return "";
  if (s.length <= maxChars) return s;
  // Keep the truncation marker within `maxChars` while still ending with an
  // ellipsis so callers can reliably detect truncation with `endsWith("…")`.
  const marker = "[TRUNCATED]…";
  if (maxChars <= marker.length) return marker.slice(0, Math.max(0, maxChars));
  return `${s.slice(0, Math.max(0, maxChars - marker.length))}${marker}`;
}

function clampInt(value: number, opts: { min: number; max: number }): number {
  const n = Number.isFinite(value) ? Math.trunc(value) : opts.min;
  return Math.max(opts.min, Math.min(opts.max, n));
}

function stableJsonStringify(value: unknown): string {
  try {
    return JSON.stringify(value, (_key, v) => (typeof v === "undefined" ? null : v));
  } catch {
    return String(value);
  }
}

const FNV1A_64_OFFSET_BASIS = 0xcbf29ce484222325n;
const FNV1A_64_PRIME = 0x100000001b3n;
const FNV1A_64_MASK = 0xffffffffffffffffn;

function fnv1a64Update(hash: bigint, text: string): bigint {
  let h = hash;
  const s = String(text);
  for (let i = 0; i < s.length; i += 1) {
    h ^= BigInt(s.charCodeAt(i));
    h = (h * FNV1A_64_PRIME) & FNV1A_64_MASK;
  }
  return h;
}

function hashText(text: string): string {
  // FNV-1a 64-bit for deterministic, dependency-free hashing.
  const hash = fnv1a64Update(FNV1A_64_OFFSET_BASIS, text) & FNV1A_64_MASK;
  return hash.toString(16).padStart(16, "0");
}

function seedFromHashHex(seedHex: string): number {
  // NOTE: `hashText` returns 16-hex characters (64-bit), which cannot be safely parsed
  // with `parseInt` into a JS number without losing bits > 2^53. We intentionally use
  // BigInt and take the low 32 bits so seeded sampling remains deterministic.
  const hex = String(seedHex ?? "").trim();
  if (!hex) return 0;
  try {
    return Number(BigInt(`0x${hex}`) & 0xffff_ffffn) >>> 0;
  } catch {
    return 0;
  }
}

function hashClassificationRecords(records: Array<{ selector: any; classification: any }>): string {
  // We intentionally hash only the fields that impact enforcement/indexing. This keeps the hash
  // stable across innocuous record shape changes while still invalidating on selector/classification
  // updates.
  let hash = FNV1A_64_OFFSET_BASIS;
  hash = fnv1a64Update(hash, String(records?.length ?? 0));

  for (const record of records || []) {
    if (!record || !record.selector || typeof record.selector !== "object") continue;

    let selectorStable = "";
    try {
      selectorStable = selectorKey(record.selector);
    } catch {
      selectorStable = stableJsonStringify(record.selector);
    }

    const classification = normalizeClassification(record.classification);
    const classificationStable = `${classification.level}:${(classification.labels ?? []).join(",")}`;

    hash = fnv1a64Update(hash, selectorStable);
    hash = fnv1a64Update(hash, "\u0000");
    hash = fnv1a64Update(hash, classificationStable);
    hash = fnv1a64Update(hash, "\u0000");
  }

  return (hash & FNV1A_64_MASK).toString(16).padStart(16, "0");
}

function createSessionId(prefix: string): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto && typeof crypto.randomUUID === "function") {
    return `${prefix}:${crypto.randomUUID()}`;
  }
  return `${prefix}:${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function getLocalStorageOrNull(): Storage | null {
  // Prefer `window.localStorage` when available.
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
      return storage;
    } catch {
      // ignore
    }
  }

  try {
    if (typeof globalThis === "undefined") return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}
