import type { LLMClient, LLMMessage } from "../../../../packages/llm/src/types.js";
import { createLLMClient } from "../../../../packages/llm/src/createLLMClient.js";

import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";
import { AIAuditRecorder } from "../../../../packages/ai-audit/src/recorder.js";

import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { DEFAULT_CLASSIFICATION, maxClassification, normalizeClassification } from "../../../../packages/security/dlp/src/classification.js";
import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { DLP_DECISION, evaluatePolicy } from "../../../../packages/security/dlp/src/policyEngine.js";
import { effectiveCellClassification, effectiveRangeClassification, a1ToCell } from "../../../../packages/security/dlp/src/selectors.js";

import { PROVENANCE_REF_SEPARATOR, type AiFunctionEvaluator, type CellValue, type ProvenanceCellValue, type SpreadsheetValue } from "./evaluateFormula.js";
import { getDesktopAIAuditStore } from "../ai/audit/auditStore.js";

import { loadDesktopLLMConfig } from "../ai/llm/settings.js";

export const AI_CELL_PLACEHOLDER = "#GETTING_DATA";
export const AI_CELL_DLP_ERROR = "#DLP!";
export const AI_CELL_ERROR = "#AI!";

const DLP_REDACTION_PLACEHOLDER = "[REDACTED]";

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
  dlp?: {
    policy?: any;
    /**
     * Classify inputs before cloud processing. If omitted, inputs default to Public.
     */
    classify?: (value: SpreadsheetValue) => { level: string; labels?: string[] };
    /**
     * Document id used to look up classification records.
     */
    documentId?: string;
    /**
     * Optional classification store used to resolve cell/range classifications.
     */
    classificationStore?: { list(documentId: string): Array<{ selector: any; classification: any }> };
    /**
     * Optional audit logger for DLP decisions (e.g. `InMemoryAuditLogger`).
     */
    auditLogger?: { log(event: any): void };
    includeRestrictedContent?: boolean;
  };
  limits?: {
    /**
     * Maximum number of cells to serialize from range/array arguments.
     */
    maxInputCells?: number;
    /**
     * Hard cap on the user prompt message size (character-based heuristic).
     */
    maxPromptChars?: number;
    /**
     * Hard cap on `inputs_preview` stored in the AI audit entry.
     */
    maxAuditPreviewChars?: number;
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
  private readonly workbookId?: string;
  private readonly sessionId: string;
  private readonly userId?: string;
  private readonly onUpdate?: () => void;

  private readonly cachePersistKey?: string;
  private readonly cacheMaxEntries: number;

  private readonly cache = new Map<string, CacheEntry>();
  private readonly inFlightByKey = new Map<string, Promise<void>>();

  private readonly dlpPolicy: any;
  private readonly classifyForDlp: (value: SpreadsheetValue) => { level: string; labels?: string[] };
  private readonly dlpDocumentId?: string;
  private readonly dlpClassificationStore?: { list(documentId: string): Array<{ selector: any; classification: any }> };
  private readonly dlpAuditLogger?: { log(event: any): void };
  private readonly includeRestrictedContent: boolean;
  private readonly maxInputCells: number;
  private readonly maxPromptChars: number;
  private readonly maxAuditPreviewChars: number;

  constructor(options: AiCellFunctionEngineOptions = {}) {
    this.llmClient = options.llmClient ?? createDefaultClient();
    this.model = options.model ?? (this.llmClient as any)?.model ?? "gpt-4o-mini";
    this.auditStore = options.auditStore ?? getDesktopAIAuditStore();
    this.workbookId = options.workbookId;
    this.sessionId = options.sessionId ?? createSessionId(options.workbookId ?? "workbook");
    this.userId = options.userId;
    this.onUpdate = options.onUpdate;

    this.cachePersistKey = options.cache?.persistKey;
    this.cacheMaxEntries = options.cache?.maxEntries ?? 500;

    this.dlpPolicy = options.dlp?.policy ?? createDefaultOrgPolicy();
    this.classifyForDlp = options.dlp?.classify ?? (() => ({ ...DEFAULT_CLASSIFICATION }));
    this.dlpDocumentId = options.dlp?.documentId;
    this.dlpClassificationStore = options.dlp?.classificationStore;
    this.dlpAuditLogger = options.dlp?.auditLogger;
    this.includeRestrictedContent = Boolean(options.dlp?.includeRestrictedContent);

    this.maxInputCells = clampInt(options.limits?.maxInputCells ?? 200, { min: 1, max: 10_000 });
    this.maxPromptChars = clampInt(options.limits?.maxPromptChars ?? 25_000, { min: 1_000, max: 1_000_000 });
    this.maxAuditPreviewChars = clampInt(options.limits?.maxAuditPreviewChars ?? 2_000, { min: 200, max: 100_000 });

    this.loadCacheFromStorage();
  }

  evaluateAiFunction(params: { name: string; args: CellValue[]; cellAddress?: string }): SpreadsheetValue {
    const fn = params.name.toUpperCase();
    const cellAddress = params.cellAddress;

    if (params.args.length === 0) return "#VALUE!";
    if ((fn === "AI.EXTRACT" || fn === "AI.CLASSIFY" || fn === "AI.TRANSLATE") && params.args.length < 2) return "#VALUE!";

    const argError = firstErrorCode(params.args);
    if (argError) return argError;

    const promptArg = params.args[0] ?? null;
    const rawInputs = params.args.slice(1);

    const { decision, selectionClassification, redactedCount, prompt, inputs, inputsPreview } = this.prepareRequest({
      promptArg,
      inputs: rawInputs,
      cellAddress,
    });

    const inputsHash = hashText(stableJsonStringify(inputs));
    const promptHash = hashText(prompt);
    const cacheKey = `${fn}\u0000${promptHash}\u0000${inputsHash}`;

    if (decision.decision === DLP_DECISION.BLOCK) {
      this.dlpAuditLogger?.log({
        type: "ai.cell_function",
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision,
        selectionClassification,
        redactedCount,
        inputs_hash: inputsHash,
        inputs_preview: inputsPreview,
        documentId: this.dlpDocumentId,
        cell: cellAddress,
        function: fn,
      });

      // Deterministic cell error for blocked content.
      void this.auditBlockedRun({
        functionName: fn,
        prompt,
        inputsHash,
        inputsPreview,
        cellAddress,
        dlp: { decision, selectionClassification, redactedCount },
      });

      return AI_CELL_DLP_ERROR;
    }

    if (decision.decision === DLP_DECISION.REDACT) {
      this.dlpAuditLogger?.log({
        type: "ai.cell_function",
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision,
        selectionClassification,
        redactedCount,
        inputs_hash: inputsHash,
        inputs_preview: inputsPreview,
        documentId: this.dlpDocumentId,
        cell: cellAddress,
        function: fn,
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
      inputsPreview,
      cellAddress,
      dlp: { decision, selectionClassification, redactedCount },
    });

    return AI_CELL_PLACEHOLDER;
  }

  /**
   * Await all in-flight LLM requests (useful in tests).
   */
  async waitForIdle(): Promise<void> {
    // Keep draining in case awaiting a promise schedules more work.
    while (this.inFlightByKey.size > 0) {
      const snapshot = Array.from(this.inFlightByKey.values());
      await Promise.all(snapshot.map((p) => p.catch(() => undefined)));
    }
  }

  private startRequest(params: {
    cacheKey: string;
    functionName: string;
      prompt: string;
      inputs: unknown;
      inputsHash: string;
      inputsPreview?: string;
      cellAddress?: string;
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
      inputsPreview?: string;
      cellAddress?: string;
      dlp: { decision: any; selectionClassification: any; redactedCount: number };
    }): Promise<void> {
    const auditInput: any = {
      function: params.functionName,
      prompt: params.prompt,
      prompt_hash: hashText(params.prompt),
      inputs_hash: params.inputsHash,
      inputs_preview: params.inputsPreview,
      cell: params.cellAddress,
      workbookId: this.workbookId,
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
          maxPromptChars: this.maxPromptChars,
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
        const key = entry.key;
        const value = entry.value;
        const updatedAtMs = entry.updatedAtMs;
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

  private sampleAndClassifyArg(
    arg: CellValue,
    records: Array<{ selector: any; classification: any }>,
    defaultSheetId: string | null,
  ): ClassifiedArg {
    if (Array.isArray(arg)) {
      const providedCells = arg.length;
      const metaTotalCells = (arg as any).__totalCells;
      const totalCells =
        typeof metaTotalCells === "number" && Number.isFinite(metaTotalCells) && metaTotalCells >= 0 ? metaTotalCells : providedCells;
      const sampledCells = Math.min(providedCells, this.maxInputCells);
      const rangeRef = (arg as any).__rangeRef;
      const rangeClassification =
        typeof rangeRef === "string" && rangeRef.trim()
          ? this.classificationForCellRef(rangeRef, null, records, defaultSheetId)
          : null;
      const items: ClassifiedItem[] = [];

      for (let i = 0; i < sampledCells; i += 1) {
        const entry = (arg as any[])[i];
        if (isProvenanceCellValue(entry)) {
          let classification = this.classificationForCellRef(entry.__cellRef, entry.value, records, defaultSheetId);
          if (rangeClassification) classification = maxClassification(classification, rangeClassification);
          items.push({
            value: entry.value,
            classification,
          });
        } else {
          const value = entry as SpreadsheetValue;
          let classification = this.classifyForDlp(value);
          if (rangeClassification) classification = maxClassification(classification, rangeClassification);
          items.push({ value, classification });
        }
      }

      return { kind: "range", items, totalCells, sampledCells, truncated: totalCells > sampledCells };
    }

    if (isProvenanceCellValue(arg)) {
      return {
        kind: "scalar",
        items: [
          {
            value: arg.value,
            classification: this.classificationForCellRef(arg.__cellRef, arg.value, records, defaultSheetId),
          },
        ],
        totalCells: 1,
        sampledCells: 1,
        truncated: false,
      };
    }

    const value = arg as SpreadsheetValue;
    return {
      kind: "scalar",
      items: [{ value, classification: this.classifyForDlp(value) }],
      totalCells: 1,
      sampledCells: 1,
      truncated: false,
    };
  }

  private classificationForCellRef(
    cellRef: string,
    value: SpreadsheetValue,
    records: Array<{ selector: any; classification: any }>,
    defaultSheetId: string | null,
  ): any {
    const valueClassification = this.classifyForDlp(value);
    if (!this.dlpDocumentId) return valueClassification;

    const cleaned = String(cellRef).replaceAll("$", "").trim();
    if (!cleaned) return valueClassification;

    const refs = cleaned
      .split(PROVENANCE_REF_SEPARATOR)
      .map((part) => part.trim())
      .filter(Boolean);

    let storeClassification = { ...DEFAULT_CLASSIFICATION };

    for (const ref of refs) {
      const bang = ref.indexOf("!");
      const rawSheet = bang === -1 ? null : ref.slice(0, bang).trim();
      const a1 = bang === -1 ? ref : ref.slice(bang + 1);

      const sheetIdRaw = rawSheet
        ? rawSheet.startsWith("'") && rawSheet.endsWith("'")
          ? rawSheet.slice(1, -1)
          : rawSheet
        : null;
      const sheetId = sheetIdRaw ?? defaultSheetId;
      if (!sheetId) continue;

      const colon = a1.indexOf(":");
      if (colon !== -1) {
        const [startA1, endA1] = a1.split(":", 2);
        if (!startA1 || !endA1) continue;
        try {
          const start = a1ToCell(startA1);
          const end = a1ToCell(endA1);
          storeClassification = maxClassification(
            storeClassification,
            effectiveRangeClassification(
              { documentId: this.dlpDocumentId, sheetId, range: { start, end } },
              records,
            ),
          );
        } catch {
          // ignore invalid ref fragments
        }
        continue;
      }

      try {
        const cell = a1ToCell(a1);
        storeClassification = maxClassification(
          storeClassification,
          effectiveCellClassification({ documentId: this.dlpDocumentId, sheetId, row: cell.row, col: cell.col }, records),
        );
      } catch {
        // ignore invalid ref fragments
      }
    }

    return maxClassification(storeClassification, valueClassification);
  }

  private prepareRequest(params: {
    promptArg: CellValue;
    inputs: CellValue[];
    cellAddress?: string;
  }): {
    decision: any;
    selectionClassification: any;
    redactedCount: number;
    prompt: string;
    inputs: unknown;
    inputsPreview?: string;
  } {
    const records = this.dlpClassificationStore && this.dlpDocumentId ? this.dlpClassificationStore.list(this.dlpDocumentId) : [];
    const defaultSheetId = sheetIdFromCellRef(params.cellAddress);

    const prompt = this.sampleAndClassifyArg(params.promptArg, records, defaultSheetId);
    const inputs = params.inputs.map((arg) => this.sampleAndClassifyArg(arg, records, defaultSheetId));

    let selectionClassification = { ...DEFAULT_CLASSIFICATION };
    for (const item of prompt.items) {
      selectionClassification = maxClassification(selectionClassification, item.classification);
    }
    for (const arg of inputs) {
      for (const item of arg.items) {
        selectionClassification = maxClassification(selectionClassification, item.classification);
      }
    }

    const decision = evaluatePolicy({
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      classification: selectionClassification,
      policy: this.dlpPolicy,
      options: { includeRestrictedContent: this.includeRestrictedContent },
    });

    let redactedCount = 0;
    if (decision.decision === DLP_DECISION.REDACT) {
      for (const item of prompt.items) {
        item.value = redactIfNeeded(item.value, item.classification, decision, this.dlpPolicy, () => {
          redactedCount += 1;
        }) as SpreadsheetValue;
      }
      for (const arg of inputs) {
        for (const item of arg.items) {
          item.value = redactIfNeeded(item.value, item.classification, decision, this.dlpPolicy, () => {
            redactedCount += 1;
          }) as SpreadsheetValue;
        }
      }
    }

    const safePrompt = renderPromptArg(prompt);
    const safeInputs = renderInputs(inputs);
    const inputsPreview =
      decision.decision === DLP_DECISION.BLOCK ? undefined : truncateText(stableJsonStringify(safeInputs), this.maxAuditPreviewChars);

    return { decision, selectionClassification, redactedCount, prompt: safePrompt, inputs: safeInputs, inputsPreview };
  }

  private async auditBlockedRun(params: {
    functionName: string;
    prompt: string;
    inputsHash: string;
    inputsPreview?: string;
    cellAddress?: string;
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
        prompt: params.prompt,
        inputs_hash: params.inputsHash,
        inputs_preview: params.inputsPreview,
        cell: params.cellAddress,
        workbookId: this.workbookId,
        dlp: params.dlp,
        blocked: true,
      },
    });
    await recorder.finalize();
  }
}

type ClassifiedItem = { value: SpreadsheetValue; classification: any };
type ClassifiedArg = {
  kind: "scalar" | "range";
  items: ClassifiedItem[];
  totalCells: number;
  sampledCells: number;
  truncated: boolean;
};

function buildMessages(params: { functionName: string; prompt: string; inputs: unknown; maxPromptChars?: number }): LLMMessage[] {
  const system: LLMMessage = {
    role: "system",
    content:
      "You are an AI function inside a spreadsheet cell. Return ONLY the final cell value (no markdown, no extra explanation).",
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

  const user: LLMMessage = {
    role: "user",
    content: typeof params.maxPromptChars === "number" ? truncateText(userContent, params.maxPromptChars) : userContent,
  };
  return [system, user];
}

function sheetIdFromCellRef(cellAddress?: string): string | null {
  if (!cellAddress) return null;
  const bang = cellAddress.indexOf("!");
  if (bang === -1) return null;
  return cellAddress.slice(0, bang);
}

function renderPromptArg(arg: ClassifiedArg): string {
  if (arg.kind === "range") {
    const text = arg.items.map((item) => normalizeScalar(item.value)).join(", ");
    return arg.truncated ? `${text} …` : text;
  }
  return normalizeScalar(arg.items[0]?.value ?? null);
}

function renderInputs(args: ClassifiedArg[]): unknown {
  if (args.length === 0) return null;
  if (args.length === 1) return renderInputArg(args[0]!);
  return args.map((arg) => renderInputArg(arg));
}

function renderInputArg(arg: ClassifiedArg): unknown {
  if (arg.kind === "range") {
    const sample = arg.items.map((item) => normalizeScalar(item.value));
    if (!arg.truncated) return sample;
    return {
      truncated: true,
      total_cells: arg.totalCells,
      sampled_cells: arg.sampledCells,
      sample,
    };
  }
  return normalizeScalar(arg.items[0]?.value ?? null);
}

function normalizeScalar(value: SpreadsheetValue): string {
  if (value === null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function firstErrorCode(args: CellValue[]): string | null {
  for (const arg of args) {
    const err = firstErrorCodeInValue(arg);
    if (err) return err;
  }
  return null;
}

function isErrorCode(value: unknown): value is string {
  return typeof value === "string" && value.startsWith("#");
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

function isProvenanceCellValue(value: unknown): value is ProvenanceCellValue {
  if (!value || typeof value !== "object") return false;
  const v = value as any;
  return typeof v.__cellRef === "string" && "value" in v;
}

function redactIfNeeded(
  value: SpreadsheetValue,
  classification: any,
  selectionDecision: any,
  policy: any,
  onRedact: () => void,
): SpreadsheetValue {
  const normalized = normalizeClassification(classification);
  const decision = evaluatePolicy({
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    classification: normalized,
    policy,
    options: { includeRestrictedContent: false },
  });
  if (selectionDecision.decision === DLP_DECISION.REDACT && decision.decision !== DLP_DECISION.ALLOW) {
    onRedact();
    return DLP_REDACTION_PLACEHOLDER;
  }
  return value;
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

function truncateText(text: string, maxChars: number): string {
  if (!Number.isFinite(maxChars) || maxChars <= 0) return "";
  if (text.length <= maxChars) return text;
  const suffix = "…[TRUNCATED]";
  if (maxChars <= suffix.length) return text.slice(0, maxChars);
  return `${text.slice(0, maxChars - suffix.length)}${suffix}`;
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
