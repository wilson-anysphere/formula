import type { LLMClient, LLMMessage } from "../../../../packages/llm/src/types.js";
import { OpenAIClient } from "../../../../packages/llm/src/openai.js";

import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";
import { LocalStorageAIAuditStore } from "../../../../packages/ai-audit/src/local-storage-store.js";
import { AIAuditRecorder } from "../../../../packages/ai-audit/src/recorder.js";

import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { DEFAULT_CLASSIFICATION, maxClassification, normalizeClassification } from "../../../../packages/security/dlp/src/classification.js";
import { createDefaultOrgPolicy } from "../../../../packages/security/dlp/src/policy.js";
import { DLP_DECISION, evaluatePolicy } from "../../../../packages/security/dlp/src/policyEngine.js";

import type { AiFunctionEvaluator, CellValue, SpreadsheetValue } from "./evaluateFormula.js";

export const AI_CELL_PLACEHOLDER = "#GETTING_DATA";
export const AI_CELL_DLP_ERROR = "#DLP!";
export const AI_CELL_ERROR = "#AI!";

const DLP_REDACTION_PLACEHOLDER = "[REDACTED]";

export interface AiCellFunctionEngineOptions {
  llmClient?: LLMClient;
  model?: string;
  auditStore?: AIAuditStore;
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
     * Optional audit logger for DLP decisions (e.g. `InMemoryAuditLogger`).
     */
    auditLogger?: { log(event: any): void };
    includeRestrictedContent?: boolean;
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
  private readonly sessionId: string;
  private readonly userId?: string;
  private readonly onUpdate?: () => void;

  private readonly cachePersistKey?: string;
  private readonly cacheMaxEntries: number;

  private readonly cache = new Map<string, CacheEntry>();
  private readonly inFlightByKey = new Map<string, Promise<void>>();

  private readonly dlpPolicy: any;
  private readonly classifyForDlp: (value: SpreadsheetValue) => { level: string; labels?: string[] };
  private readonly dlpAuditLogger?: { log(event: any): void };
  private readonly includeRestrictedContent: boolean;

  constructor(options: AiCellFunctionEngineOptions = {}) {
    this.llmClient = options.llmClient ?? createDefaultClient();
    this.model = options.model ?? "gpt-4o-mini";
    this.auditStore = options.auditStore ?? new LocalStorageAIAuditStore();
    this.sessionId = options.sessionId ?? createSessionId("workbook");
    this.userId = options.userId;
    this.onUpdate = options.onUpdate;

    this.cachePersistKey = options.cache?.persistKey;
    this.cacheMaxEntries = options.cache?.maxEntries ?? 500;

    this.dlpPolicy = options.dlp?.policy ?? createDefaultOrgPolicy();
    this.classifyForDlp = options.dlp?.classify ?? (() => ({ ...DEFAULT_CLASSIFICATION }));
    this.dlpAuditLogger = options.dlp?.auditLogger;
    this.includeRestrictedContent = Boolean(options.dlp?.includeRestrictedContent);

    this.loadCacheFromStorage();
  }

  evaluateAiFunction(params: { name: string; args: CellValue[]; cellAddress?: string }): SpreadsheetValue {
    const fn = params.name.toUpperCase();
    const cellAddress = params.cellAddress;

    if (params.args.length === 0) return "#VALUE!";
    if ((fn === "AI.EXTRACT" || fn === "AI.CLASSIFY" || fn === "AI.TRANSLATE") && params.args.length < 2) return "#VALUE!";

    const argError = firstErrorCode(params.args);
    if (argError) return argError;

    const rawPrompt = normalizePrompt(params.args[0] ?? null);
    const rawInputs = params.args.slice(1);

    const { decision, selectionClassification, redactedCount, prompt, inputs } = this.applyDlp(rawPrompt, rawInputs);

    const inputsHash = hashText(stableJsonStringify(inputs));
    const cacheKey = `${fn}\u0000${prompt}\u0000${inputsHash}`;

    if (decision.decision === DLP_DECISION.BLOCK) {
      this.dlpAuditLogger?.log({
        type: "ai.cell_function",
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision,
        selectionClassification,
        redactedCount,
        cell: cellAddress,
        function: fn,
      });

      // Deterministic cell error for blocked content.
      void this.auditBlockedRun({
        functionName: fn,
        prompt,
        inputsHash,
        cellAddress,
        dlp: { decision, selectionClassification, redactedCount },
      });

      return AI_CELL_DLP_ERROR;
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
    cellAddress?: string;
    dlp: { decision: any; selectionClassification: any; redactedCount: number };
  }): Promise<void> {
    const auditInput: any = {
      function: params.functionName,
      prompt: params.prompt,
      inputs_hash: params.inputsHash,
      cell: params.cellAddress,
      dlp: {
        decision: params.dlp.decision,
        selectionClassification: params.dlp.selectionClassification,
        redactedCount: params.dlp.redactedCount,
      },
    };

    const recorder = new AIAuditRecorder({
      store: this.auditStore,
      session_id: this.sessionId,
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

  private applyDlp(
    prompt: string,
    inputs: CellValue[],
  ): {
    decision: any;
    selectionClassification: any;
    redactedCount: number;
    prompt: string;
    inputs: unknown;
  } {
    const flattenedValues: SpreadsheetValue[] = [prompt];
    for (const arg of inputs) {
      if (Array.isArray(arg)) {
        flattenedValues.push(...(arg as SpreadsheetValue[]));
      } else {
        flattenedValues.push(arg as SpreadsheetValue);
      }
    }

    let selectionClassification = { ...DEFAULT_CLASSIFICATION };
    for (const value of flattenedValues) {
      selectionClassification = maxClassification(selectionClassification, this.classifyForDlp(value));
    }

    const decision = evaluatePolicy({
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      classification: selectionClassification,
      policy: this.dlpPolicy,
      options: { includeRestrictedContent: this.includeRestrictedContent },
    });

    if (decision.decision !== DLP_DECISION.REDACT) {
      return { decision, selectionClassification, redactedCount: 0, prompt, inputs: normalizeInputs(inputs) };
    }

    let redactedCount = 0;
    const safePrompt = normalizeScalar(
      redactIfNeeded(prompt, this.classifyForDlp(prompt), decision, this.dlpPolicy, () => {
        redactedCount += 1;
      }) as SpreadsheetValue,
    );

    const safeInputs = normalizeInputs(
      inputs.map((arg) => {
        if (Array.isArray(arg)) {
          return (arg as SpreadsheetValue[]).map((v) =>
            redactIfNeeded(v, this.classifyForDlp(v), decision, this.dlpPolicy, () => {
              redactedCount += 1;
            }),
          );
        }
        return redactIfNeeded(arg as SpreadsheetValue, this.classifyForDlp(arg as SpreadsheetValue), decision, this.dlpPolicy, () => {
          redactedCount += 1;
        });
      }) as any,
    );

    this.dlpAuditLogger?.log({
      type: "ai.cell_function",
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      decision,
      selectionClassification,
      redactedCount,
    });

    return { decision, selectionClassification, redactedCount, prompt: safePrompt, inputs: safeInputs };
  }

  private async auditBlockedRun(params: {
    functionName: string;
    prompt: string;
    inputsHash: string;
    cellAddress?: string;
    dlp: { decision: any; selectionClassification: any; redactedCount: number };
  }): Promise<void> {
    const recorder = new AIAuditRecorder({
      store: this.auditStore,
      session_id: this.sessionId,
      user_id: this.userId,
      mode: "cell_function",
      model: this.model,
      input: {
        function: params.functionName,
        prompt: params.prompt,
        inputs_hash: params.inputsHash,
        cell: params.cellAddress,
        dlp: params.dlp,
        blocked: true,
      },
    });
    await recorder.finalize();
  }
}

function buildMessages(params: { functionName: string; prompt: string; inputs: unknown }): LLMMessage[] {
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

  const user: LLMMessage = { role: "user", content: userContent };
  return [system, user];
}

function normalizePrompt(arg: CellValue): string {
  if (Array.isArray(arg)) {
    return (arg as SpreadsheetValue[]).map((v) => normalizeScalar(v)).join(", ");
  }
  return normalizeScalar(arg as SpreadsheetValue);
}

function normalizeInputs(args: CellValue[]): unknown {
  if (args.length === 0) return null;
  if (args.length === 1) {
    const only = args[0] as any;
    if (Array.isArray(only)) return (only as SpreadsheetValue[]).map((v) => normalizeScalar(v));
    return normalizeScalar(only as SpreadsheetValue);
  }
  return args.map((arg) => {
    if (Array.isArray(arg)) return (arg as SpreadsheetValue[]).map((v) => normalizeScalar(v));
    return normalizeScalar(arg as SpreadsheetValue);
  });
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
    if (isErrorCode(arg)) return arg;
    if (Array.isArray(arg)) {
      for (const value of arg) {
        if (isErrorCode(value)) return value;
      }
    }
  }
  return null;
}

function isErrorCode(value: unknown): value is string {
  return typeof value === "string" && value.startsWith("#");
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
    return new OpenAIClient();
  } catch {
    return {
      async chat() {
        return {
          message: { role: "assistant", content: "AI is not configured." },
          usage: { promptTokens: 0, completionTokens: 0 },
        } as any;
      },
    };
  }
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
