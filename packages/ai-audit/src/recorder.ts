import type { AIAuditStore } from "./store.ts";
import type { AIAuditEntry, AIMode, TokenUsage, ToolCallLog, UserFeedback, AIVerificationResult } from "./types.ts";

export interface AIAuditRecorderOptions {
  store: AIAuditStore;
  session_id: string;
  workbook_id?: string;
  mode: AIMode;
  input: unknown;
  model: string;
  user_id?: string;
  timestamp_ms?: number;
}

export interface ToolCallRecord {
  id?: string;
  name: string;
  parameters: unknown;
  requires_approval?: boolean;
}

export class AIAuditRecorder {
  private readonly store: AIAuditStore;
  private readonly startedAtMs: number;

  private readonly toolCallIndexById = new Map<string, number>();

  readonly entry: AIAuditEntry;
  /**
   * Captures any persistence error encountered during `finalize()`.
   *
   * `finalize()` is intentionally best-effort (never throws) because it's often
   * invoked from `finally` blocks in higher-level AI flows.
   */
  finalizeError: unknown | undefined;
  /**
   * Stringified message form of `finalizeError` (for backwards compatibility with
   * older callers that expected a string).
   */
  finalize_error?: string;

  constructor(options: AIAuditRecorderOptions) {
    this.store = options.store;
    this.startedAtMs = nowMs();

    this.entry = {
      id: createAuditId(),
      timestamp_ms: options.timestamp_ms ?? Date.now(),
      session_id: options.session_id,
      workbook_id: options.workbook_id,
      user_id: options.user_id,
      mode: options.mode,
      input: options.input,
      model: options.model,
      tool_calls: []
    };
  }

  recordTokenUsage(usage: TokenUsage): void {
    const existing = this.entry.token_usage ?? { prompt_tokens: 0, completion_tokens: 0 };
    const promptTokens = existing.prompt_tokens + (usage.prompt_tokens ?? 0);
    const completionTokens = existing.completion_tokens + (usage.completion_tokens ?? 0);
    const totalTokens =
      (existing.total_tokens ?? existing.prompt_tokens + existing.completion_tokens) +
      (usage.total_tokens ?? (usage.prompt_tokens ?? 0) + (usage.completion_tokens ?? 0));

    this.entry.token_usage = {
      prompt_tokens: promptTokens,
      completion_tokens: completionTokens,
      total_tokens: totalTokens
    };
  }

  recordModelLatency(duration_ms: number): void {
    this.entry.latency_ms = (this.entry.latency_ms ?? 0) + duration_ms;
  }

  recordToolCall(call: ToolCallRecord): number {
    const record: ToolCallLog = {
      name: call.name,
      parameters: call.parameters,
      requires_approval: call.requires_approval
    };

    const index = this.entry.tool_calls.push(record) - 1;
    if (call.id) {
      this.toolCallIndexById.set(call.id, index);
    }
    return index;
  }

  recordToolApproval(callIdOrIndex: string | number, approved: boolean): void {
    const index = this.resolveToolIndex(callIdOrIndex);
    const entry = this.entry.tool_calls[index];
    if (!entry) return;
    entry.approved = approved;
  }

  recordToolResult(
    callIdOrIndex: string | number,
    result: {
      ok?: boolean;
      duration_ms?: number;
      result?: unknown;
      audit_result_summary?: unknown;
      result_truncated?: boolean;
      error?: string;
    }
  ): void {
    const index = this.resolveToolIndex(callIdOrIndex);
    const entry = this.entry.tool_calls[index];
    if (!entry) return;
    if (result.ok !== undefined) entry.ok = result.ok;
    if (result.duration_ms !== undefined) entry.duration_ms = result.duration_ms;
    if (result.result !== undefined) entry.result = result.result;
    if (result.audit_result_summary !== undefined) entry.audit_result_summary = result.audit_result_summary;
    if (result.result_truncated !== undefined) entry.result_truncated = result.result_truncated;
    if (result.error !== undefined) entry.error = result.error;
  }

  setUserFeedback(feedback: UserFeedback): void {
    this.entry.user_feedback = feedback;
  }

  setVerification(verification: AIVerificationResult): void {
    this.entry.verification = verification;
  }

  async finalize(): Promise<void> {
    if (this.entry.latency_ms === undefined) {
      this.entry.latency_ms = nowMs() - this.startedAtMs;
    }

    try {
      await this.store.logEntry(this.entry);
      this.finalizeError = undefined;
      this.finalize_error = undefined;
    } catch (error) {
      this.finalizeError = error;
      this.finalize_error = error instanceof Error ? error.message : String(error);
    }
  }

  getFinalizeError(): unknown | undefined {
    return this.finalizeError;
  }

  private resolveToolIndex(callIdOrIndex: string | number): number {
    if (typeof callIdOrIndex === "number") return callIdOrIndex;
    return this.toolCallIndexById.get(callIdOrIndex) ?? -1;
  }
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function createAuditId(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `audit_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}
