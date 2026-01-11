import type { LLMClient, LLMMessage } from "../../../../packages/llm/src/types.js";
import { createLLMClient } from "../../../../packages/llm/src/createLLMClient.js";

import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";
import { AIAuditRecorder } from "../../../../packages/ai-audit/src/recorder.js";

import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL, DEFAULT_CLASSIFICATION, maxClassification, normalizeClassification } from "../../../../packages/security/dlp/src/classification.js";
import { DLP_DECISION, evaluatePolicy } from "../../../../packages/security/dlp/src/policyEngine.js";
import { effectiveCellClassification, effectiveRangeClassification } from "../../../../packages/security/dlp/src/selectors.js";

import { parseA1Range, rangeToA1, type RangeAddress } from "./a1.js";
import {
  PROVENANCE_REF_SEPARATOR,
  type AiFunctionArgumentProvenance,
  type AiFunctionEvaluator,
  type CellValue,
  type ProvenanceCellValue,
  type SpreadsheetValue,
} from "./evaluateFormula.js";

import { getDesktopAIAuditStore } from "../ai/audit/auditStore.js";
import { getAiCloudDlpOptions } from "../ai/dlp/aiDlp.js";
import { loadDesktopLLMConfig } from "../ai/llm/settings.js";

export const AI_CELL_PLACEHOLDER = "#GETTING_DATA";
export const AI_CELL_DLP_ERROR = "#DLP!";
export const AI_CELL_ERROR = "#AI!";

const DLP_REDACTION_PLACEHOLDER = "[REDACTED]";

const DEFAULT_RANGE_SAMPLE_LIMIT = 200;
const MAX_PROMPT_CHARS = 2_000;
const MAX_SCALAR_CHARS = 500;
const MAX_RANGE_HEADER_VALUES = 20;
const MAX_RANGE_PREVIEW_VALUES = 30;
const MAX_RANGE_SAMPLE_VALUES = 30;
const MAX_USER_MESSAGE_CHARS = 16_000;

export interface AiCellFunctionEngineOptions {
  llmClient?: LLMClient;
  model?: string;
  auditStore?: AIAuditStore;
  workbookId?: string;
  sessionId?: string;
  userId?: string;
  onUpdate?: () => void;
  cache?: {
    /**
     * When set, persists cache entries in localStorage under this key.
     */
    persistKey?: string;
    /**
     * Maximum number of cached entries to retain.
     */
    maxEntries?: number;
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
  };
}

type CacheEntry = { value: string; updatedAtMs: number };

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

  private readonly cachePersistKey?: string;
  private readonly cacheMaxEntries: number;

  private readonly maxInputCells: number;
  private readonly maxUserMessageChars: number;
  private readonly maxAuditPreviewChars: number;
  private readonly maxCellChars: number;

  private readonly cache = new Map<string, CacheEntry>();
  private readonly inFlightByKey = new Map<string, Promise<void>>();
  private readonly pendingAudits = new Set<Promise<void>>();

  constructor(options: AiCellFunctionEngineOptions = {}) {
    this.llmClient = options.llmClient ?? createDefaultClient();
    this.model = options.model ?? (this.llmClient as any)?.model ?? "gpt-4o-mini";
    this.auditStore = options.auditStore ?? getDesktopAIAuditStore();
    this.workbookId = options.workbookId ?? "local-workbook";
    this.sessionId = options.sessionId ?? createSessionId(this.workbookId);
    this.userId = options.userId;
    this.onUpdate = options.onUpdate;

    this.cachePersistKey = options.cache?.persistKey;
    this.cacheMaxEntries = options.cache?.maxEntries ?? 500;

    this.maxInputCells = clampInt(options.limits?.maxInputCells ?? DEFAULT_RANGE_SAMPLE_LIMIT, { min: 1, max: 10_000 });
    this.maxUserMessageChars = clampInt(options.limits?.maxPromptChars ?? MAX_USER_MESSAGE_CHARS, {
      min: 1_000,
      max: 1_000_000,
    });
    this.maxAuditPreviewChars = clampInt(options.limits?.maxAuditPreviewChars ?? 2_000, { min: 200, max: 100_000 });
    this.maxCellChars = clampInt(options.limits?.maxCellChars ?? MAX_SCALAR_CHARS, { min: 50, max: 100_000 });

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
    const dlp = getAiCloudDlpOptions({ documentId: this.workbookId, sheetId: defaultSheetId });

    const alignedProvenance = alignArgProvenance({
      args: params.args,
      provenance: params.argProvenance,
      defaultSheetId,
    });

    const { selectionClassification, references } = computeSelectionClassification({
      documentId: this.workbookId,
      args: params.args,
      provenance: alignedProvenance,
      defaultSheetId,
      maxCellChars: this.maxCellChars,
      classificationRecords: dlp.classificationRecords,
    });

    const decision = evaluatePolicy({
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      classification: selectionClassification,
      policy: dlp.policy,
      options: { includeRestrictedContent: false },
    });

    const { prompt, inputs, inputsHash, inputsCompaction, redactedCount } = preparePromptAndInputs({
      functionName: fn,
      args: params.args,
      provenance: alignedProvenance,
      decision,
      documentId: this.workbookId,
      defaultSheetId,
      policy: dlp.policy,
      classificationRecords: dlp.classificationRecords,
      maxCellChars: this.maxCellChars,
    });

    const cacheKey = `${this.model}\u0000${fn}\u0000${prompt}\u0000${inputsHash}`;

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
        references,
      });
    }

    const cached = this.cache.get(cacheKey);
    if (cached) return cached.value;

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

    try {
      const started = nowMs();
      const response = await this.llmClient.chat({
        model: this.model,
        messages: buildMessages({
          functionName: params.functionName,
          prompt: params.prompt,
          inputs: params.inputs,
          maxPromptChars: this.maxUserMessageChars,
        }),
      });
      recorder.recordModelLatency(nowMs() - started);

      const promptTokens = response.usage?.promptTokens;
      const completionTokens = response.usage?.completionTokens;
      if (typeof promptTokens === "number" || typeof completionTokens === "number") {
        recorder.recordTokenUsage({
          prompt_tokens: typeof promptTokens === "number" ? promptTokens : 0,
          completion_tokens: typeof completionTokens === "number" ? completionTokens : 0,
        });
      }

      const content = String(response.message?.content ?? "").trim();
      const finalText = sanitizeCellText(content);
      this.writeCache(params.cacheKey, finalText);
    } catch (error) {
      auditInput.error = error instanceof Error ? error.message : String(error);
      this.writeCache(params.cacheKey, AI_CELL_ERROR);
    } finally {
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

  private writeCache(cacheKey: string, value: string): void {
    this.cache.set(cacheKey, { value, updatedAtMs: Date.now() });
    while (this.cache.size > this.cacheMaxEntries) {
      const oldestKey = this.cache.keys().next().value as string | undefined;
      if (oldestKey === undefined) break;
      this.cache.delete(oldestKey);
    }
    this.saveCacheToStorage();
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
        if (typeof key !== "string" || typeof value !== "string") continue;
        this.cache.set(key, {
          value,
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

type ParsedProvenanceRef = { sheetId: string; range: RangeAddress; canonical: string; isCell: boolean };

function parseProvenanceRef(refText: string, defaultSheetId: string): ParsedProvenanceRef | null {
  const cleaned = String(refText).replaceAll("$", "").trim();
  if (!cleaned) return null;
  const { sheetId: sheetFromRef, ref } = splitSheetQualifier(cleaned);
  const sheetId = (sheetFromRef ?? defaultSheetId).trim();
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

function inferProvenanceFromValue(value: CellValue): AiFunctionArgumentProvenance {
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
    for (const entry of value as any[]) {
      if (!isProvenanceCellValue(entry)) continue;
      for (const ref of String(entry.__cellRef).split(PROVENANCE_REF_SEPARATOR)) addRef(ref);
    }
  }

  return { cells: Array.from(cells), ranges: Array.from(ranges) };
}

function normalizeProvenance(entry: AiFunctionArgumentProvenance | undefined, defaultSheetId: string): AiFunctionArgumentProvenance {
  const cells = new Set<string>();
  const ranges = new Set<string>();

  const add = (raw: string): void => {
    const parsed = parseProvenanceRef(raw, defaultSheetId);
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
}): AiFunctionArgumentProvenance[] {
  const out: AiFunctionArgumentProvenance[] = [];
  for (let i = 0; i < params.args.length; i += 1) {
    const fromCaller = normalizeProvenance(params.provenance?.[i], params.defaultSheetId);
    const fromValue = normalizeProvenance(inferProvenanceFromValue(params.args[i]!), params.defaultSheetId);
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
    const parsed = parseProvenanceRef(cellRef, params.defaultSheetId);
    if (!parsed || !parsed.isCell) continue;
    const classification = effectiveCellClassification(
      { documentId: params.documentId, sheetId: parsed.sheetId, row: parsed.range.start.row, col: parsed.range.start.col } as any,
      params.classificationRecords,
    );
    selectionClassification = maxClassification(selectionClassification, classification);
    if (selectionClassification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
  }

  if (selectionClassification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    for (const rangeRef of references.ranges) {
      const parsed = parseProvenanceRef(rangeRef, params.defaultSheetId);
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
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  maxCellChars: number;
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
    documentId: params.documentId,
    defaultSheetId: params.defaultSheetId,
    policy: params.policy,
    classificationRecords: params.classificationRecords,
    maxCellChars: params.maxCellChars,
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
      documentId: params.documentId,
      defaultSheetId: params.defaultSheetId,
      policy: params.policy,
      classificationRecords: params.classificationRecords,
      maxCellChars: params.maxCellChars,
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
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  maxCellChars: number;
}): { value: unknown; compaction: unknown; redactedCount: number } {
  if (Array.isArray(params.value)) {
    return compactArrayForPrompt({
      functionName: params.functionName,
      argIndex: params.argIndex,
      values: unwrapArrayValues(params.value),
      rangeRef: rangeRefFromArray(params.value),
      provenance: params.provenance,
      shouldRedact: params.shouldRedact,
      documentId: params.documentId,
      defaultSheetId: params.defaultSheetId,
      policy: params.policy,
      classificationRecords: params.classificationRecords,
      maxCellChars: params.maxCellChars,
    });
  }

  const scalarValue = isProvenanceCellValue(params.value) ? params.value.value : (params.value as SpreadsheetValue);
  return compactScalarForPrompt({
    value: scalarValue,
    provenance: params.provenance,
    shouldRedact: params.shouldRedact,
    documentId: params.documentId,
    defaultSheetId: params.defaultSheetId,
    policy: params.policy,
    classificationRecords: params.classificationRecords,
    maxCellChars: params.maxCellChars,
  });
}

function classificationForProvenance(params: {
  documentId: string;
  defaultSheetId: string;
  provenance: AiFunctionArgumentProvenance;
  classificationRecords: Array<{ selector: any; classification: any }>;
}): any {
  let classification = { ...DEFAULT_CLASSIFICATION };

  const visit = (refText: string): void => {
    const parsed = parseProvenanceRef(refText, params.defaultSheetId);
    if (!parsed) return;

    if (parsed.isCell) {
      classification = maxClassification(
        classification,
        effectiveCellClassification(
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
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  maxCellChars: number;
}): { value: string; compaction: unknown; redactedCount: number } {
  let classification = { ...DEFAULT_CLASSIFICATION };

  const hasRefs = Boolean((params.provenance.cells?.length ?? 0) > 0 || (params.provenance.ranges?.length ?? 0) > 0);
  if (hasRefs) {
    classification = classificationForProvenance({
      documentId: params.documentId,
      defaultSheetId: params.defaultSheetId,
      provenance: params.provenance,
      classificationRecords: params.classificationRecords,
    });
  } else {
    classification = heuristicClassifyValue(params.value as any, params.maxCellChars);
  }

  const decision = evaluatePolicy({
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    classification,
    policy: params.policy,
    options: { includeRestrictedContent: false },
  });

  if (params.shouldRedact && decision.decision !== DLP_DECISION.ALLOW) {
    return { value: DLP_REDACTION_PLACEHOLDER, compaction: { kind: "scalar", redacted: true }, redactedCount: 1 };
  }

  return {
    value: formatScalar(params.value, { maxChars: params.maxCellChars }),
    compaction: { kind: "scalar", redacted: false },
    redactedCount: 0,
  };
}

function compactArrayForPrompt(params: {
  functionName: string;
  argIndex: number;
  values: SpreadsheetValue[];
  rangeRef: string | null;
  provenance: AiFunctionArgumentProvenance;
  shouldRedact: boolean;
  documentId: string;
  defaultSheetId: string;
  policy: any;
  classificationRecords: Array<{ selector: any; classification: any }>;
  maxCellChars: number;
}): { value: unknown; compaction: unknown; redactedCount: number } {
  const providedCount = params.values.length;

  const rangeCandidate = params.provenance.ranges?.length === 1 ? params.provenance.ranges[0]! : params.rangeRef;
  const parsedRange = rangeCandidate ? parseProvenanceRef(rangeCandidate, params.defaultSheetId) : null;
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

  const rangeText = range && rangeSheetId ? `${rangeSheetId}!${rangeToA1(range)}` : range ? rangeToA1(range) : null;

  const seedHex = hashText(
    `${params.documentId}:${rangeSheetId ?? params.defaultSheetId}:${params.functionName}:${params.argIndex}:${rangeText ?? "array"}`,
  );
  const seed = Number.parseInt(seedHex, 16) >>> 0;
  const rand = mulberry32(seed);

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

  const shouldRedactAll =
    params.shouldRedact &&
    !range &&
    ((params.provenance.cells?.length ?? 0) > 0 || (params.provenance.ranges?.length ?? 0) > 0) &&
    (() => {
      const classification = classificationForProvenance({
        documentId: params.documentId,
        defaultSheetId: params.defaultSheetId,
        provenance: params.provenance,
        classificationRecords: params.classificationRecords,
      });
      const decision = evaluatePolicy({
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        classification,
        policy: params.policy,
        options: { includeRestrictedContent: false },
      });
      return decision.decision !== DLP_DECISION.ALLOW;
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
    const raw = params.values[index] ?? null;

    if (params.shouldRedact) {
      if (shouldRedactAll) {
        if (!redactedSeen.has(index)) {
          redactedSeen.add(index);
          redactedCount += 1;
        }
        return DLP_REDACTION_PLACEHOLDER;
      }

      // If we can map indices -> cells, enforce per-cell redaction.
      if (range && cols && rangeSheetId) {
        const rowOffset = Math.floor(index / cols);
        const colOffset = index % cols;
        const row = range.start.row + rowOffset;
        const col = range.start.col + colOffset;

        const classification = effectiveCellClassification(
          { documentId: params.documentId, sheetId: rangeSheetId, row, col } as any,
          params.classificationRecords,
        );
        const cellDecision = evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification,
          policy: params.policy,
          options: { includeRestrictedContent: false },
        });
        if (cellDecision.decision !== DLP_DECISION.ALLOW) {
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
    kind: rangeText ? "range" : "array",
    ...(rangeText ? { range: rangeText } : {}),
    ...countMeta,
    shape,
    ...(header ? { header } : {}),
    preview,
    ...(sample ? { sample: { seed: seedHex, values: sample } } : {}),
    ...(numericSummary ? { numeric_summary: numericSummary } : {}),
    ...(redactedCount ? { redacted_values: redactedCount } : {}),
  };

  const compaction = {
    kind: rangeText ? "range" : "array",
    ...(rangeText ? { range: rangeText } : {}),
    ...countMeta,
    shape,
    header_count: headerIndices.length,
    preview_count: previewCount,
    sample: sampleIndices.length ? { seed: seedHex, indices: sampleIndices } : null,
    redacted_values: redactedCount,
  };

  return { value, compaction, redactedCount };
}

function unwrapArrayValues(values: Array<SpreadsheetValue | ProvenanceCellValue>): SpreadsheetValue[] {
  return values.map((entry) => (isProvenanceCellValue(entry) ? entry.value : (entry as SpreadsheetValue)));
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
    for (const entry of value) {
      const err = firstErrorCodeInValue(entry as any);
      if (err) return err;
    }
  }
  return null;
}

function isErrorCode(value: unknown): value is string {
  return typeof value === "string" && value.startsWith("#");
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

function hashText(text: string): string {
  // FNV-1a 32-bit for deterministic, dependency-free hashing.
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function createDefaultClient(): LLMClient {
  try {
    const config = loadDesktopLLMConfig();
    if (config) return createLLMClient(config as any) as any;
  } catch {
    // fall through to stub
  }
  return {
    async chat() {
      return {
        message: { role: "assistant", content: "AI is not configured." },
        usage: { promptTokens: 0, completionTokens: 0 },
      } as any;
    },
  };
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
