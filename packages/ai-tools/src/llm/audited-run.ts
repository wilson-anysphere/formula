import { AIAuditRecorder } from "@formula/ai-audit";
import type { AIAuditStore, AIMode, TokenUsage } from "@formula/ai-audit";

import { runChatWithToolsStreaming, serializeToolResultForModel } from "../../../llm/src/index.js";
import type { ChatStreamEvent, ToolCall } from "../../../llm/src/types.js";

import { redactUrlSecrets } from "../utils/urlRedaction.ts";
import { classifyQueryNeedsTools, verifyAssistantClaims, verifyToolUsage, type VerificationResult } from "./verification.ts";

export interface AuditedRunOptions {
  audit_store: AIAuditStore;
  session_id: string;
  workbook_id?: string;
  user_id?: string;
  mode: AIMode;
  input: unknown;
  model: string;
  /**
   * When true, store the full tool result object in the audit entry.
   *
   * Default is false, which stores a bounded `audit_result_summary` instead to
   * avoid blowing up LocalStorage-backed audit stores.
   */
  store_full_tool_results?: boolean;
  /**
   * Max size of the stored tool result summary (in characters).
   */
  max_audit_result_chars?: number;
  /**
   * Max size of tool call parameters stored in audit entries (in characters).
   *
   * Tool call parameters can contain large payloads (e.g. `set_range.values`). We
   * cap them to keep LocalStorage-backed audit logs bounded.
   */
  max_audit_parameter_chars?: number;
}

export interface AuditedRunParams {
  client: { chat: (request: any) => Promise<any>; streamChat?: (request: any) => AsyncIterable<ChatStreamEvent> };
  tool_executor: { tools: any[]; execute: (call: any) => Promise<any> };
  messages: any[];
  audit: AuditedRunOptions;
  max_iterations?: number;
  require_approval?: (call: ToolCall) => Promise<boolean>;
  /**
   * When true, approval denials are surfaced to the model as tool results (ok:false)
   * and the loop continues, allowing the model to re-plan. Default behavior is to throw.
   */
  continue_on_approval_denied?: boolean;
  /**
   * Optional hooks for UI surfaces (e.g. chat panels) that still want to surface
   * tool call + result events while relying on this helper for audit logging.
   */
  on_tool_call?: (call: ToolCall, meta: { requiresApproval: boolean }) => void;
  on_tool_result?: (call: ToolCall, result: unknown) => void;
  /**
   * Optional stream hook for UI surfaces that want partial assistant output.
   */
  on_stream_event?: (event: ChatStreamEvent) => void;
  /**
   * Optional LLM request parameters forwarded to `runChatWithTools`.
   * If omitted, `audit.model` is used as the request model (providers may ignore it).
   */
  model?: string;
  temperature?: number;
  max_tokens?: number;
  /**
   * Optional abort signal forwarded to the underlying LLM client.
   */
  signal?: AbortSignal;
  /**
   * When enabled, if the query is classified as needing data tools but the model
   * produced an answer without any successful tool calls, the run is retried
   * once with a stricter system instruction to "use tools; do not guess".
   */
  strict_tool_verification?: boolean;
  /**
   * Optional attachment context used by the verifier (range/table/chart/formula references).
   */
  attachments?: unknown[] | null;
  /**
   * When enabled, run a post-response verification pass that extracts numeric
   * spreadsheet claims from the assistant message and validates them via
   * read-only spreadsheet tools (e.g. compute_statistics, read_range).
   */
  verify_claims?: boolean;
  /**
   * Optional tool executor used *only* for claim verification (to avoid side
   * effects like UI tool-result capture). Defaults to `tool_executor`.
   */
  verification_tool_executor?: { tools?: any[]; execute: (call: any) => Promise<any> };
  /**
   * Limit the number of extracted claims verified per response (safety/perf).
   */
  verification_max_claims?: number;
}

/**
 * Run the provider-agnostic tool-calling loop with audit logging.
 *
 * This is intended as a thin integration helper to wire `packages/llm` + `packages/ai-tools`
 * to `packages/ai-audit` without duplicating orchestration logic across UI surfaces.
 */
export async function runChatWithToolsAudited(params: AuditedRunParams): Promise<{ messages: any[]; final: string }> {
  const result = await runChatWithToolsAuditedVerified(params);
  return { messages: result.messages, final: result.final };
}

export async function runChatWithToolsAuditedVerified(
  params: AuditedRunParams
): Promise<{ messages: any[]; final: string; verification: VerificationResult }> {
  const recorder = new AIAuditRecorder({
    store: params.audit.audit_store,
    session_id: params.audit.session_id,
    workbook_id: params.audit.workbook_id,
    user_id: params.audit.user_id,
    mode: params.audit.mode,
    input: params.audit.input,
    model: params.audit.model
  });

  const auditedClient = {
    async chat(request: any) {
      const started = nowMs();
      const response = await params.client.chat(request);
      recorder.recordModelLatency(nowMs() - started);
      const usage = extractTokenUsage(response?.usage);
      if (usage) recorder.recordTokenUsage(usage);
      return response;
    },
    streamChat: params.client.streamChat
      ? async function* streamChat(request: any) {
          const started = nowMs();
          let recordedUsage = false;
          try {
            for await (const event of params.client.streamChat!(request)) {
              if (!recordedUsage && event && typeof event === "object" && event.type === "done" && event.usage) {
                const promptRaw = event.usage.promptTokens;
                const completionRaw = event.usage.completionTokens;
                const totalRaw = event.usage.totalTokens;
                const prompt = typeof promptRaw === "number" && Number.isFinite(promptRaw) ? promptRaw : null;
                const completion =
                  typeof completionRaw === "number" && Number.isFinite(completionRaw) ? completionRaw : null;
                const total = typeof totalRaw === "number" && Number.isFinite(totalRaw) ? totalRaw : null;
                if (prompt != null || completion != null) {
                  recorder.recordTokenUsage({
                    prompt_tokens: prompt ?? 0,
                    completion_tokens: completion ?? 0,
                    total_tokens: total ?? (prompt != null && completion != null ? prompt + completion : undefined)
                  });
                }
                recordedUsage = true;
              }
              yield event;
            }
          } finally {
            recorder.recordModelLatency(nowMs() - started);
          }
        }
      : undefined
  };

  try {
    const userText = extractLastUserText(params.messages);
    const needsTools = classifyQueryNeedsTools({ userText, attachments: params.attachments });

    const runOnce = async (messages: any[]) =>
      runChatWithToolsStreaming({
        client: auditedClient as any,
        toolExecutor: params.tool_executor as any,
        messages: messages as any,
        maxIterations: params.max_iterations,
        continueOnApprovalDenied: params.continue_on_approval_denied,
        onStreamEvent: params.on_stream_event,
        model: params.model ?? params.audit.model,
        temperature: params.temperature,
        maxTokens: params.max_tokens,
        signal: params.signal,
        onToolCall: (call: any, meta: any) => {
          const maxParamChars = params.audit.max_audit_parameter_chars ?? 20_000;
          const rawParams = sanitizeAuditToolParameters(call.name, call.arguments);
          recorder.recordToolCall({
            id: call.id,
            name: call.name,
            parameters: compactAuditValue(rawParams, maxParamChars),
            requires_approval: Boolean(meta?.requiresApproval)
          });
          params.on_tool_call?.(call, meta);
        },
        onToolResult: (call: any, toolResult: any) => {
          const storeFull = params.audit.store_full_tool_results ?? false;
          const maxSummaryChars = params.audit.max_audit_result_chars ?? 20_000;
          const summary = serializeToolResultForModel({ toolCall: call, result: toolResult, maxChars: maxSummaryChars });
          recorder.recordToolResult(call.id, {
            ok: typeof toolResult?.ok === "boolean" ? toolResult.ok : undefined,
            duration_ms: extractToolDuration(toolResult),
            ...(storeFull ? { result: toolResult } : {}),
            audit_result_summary: summary,
            result_truncated: !storeFull,
            error: toolResult?.error?.message ? String(toolResult.error.message) : undefined
          });
          params.on_tool_result?.(call, toolResult);
        },
        requireApproval: async (call: any) => {
          const approved = await (params.require_approval ?? (async () => true))(call);
          recorder.recordToolApproval(call.id, approved);
          return approved;
        }
      });

    let result = await runOnce(params.messages);
    const successfulToolInitially = recorder.entry.tool_calls.some((call) => call.ok === true);

    if (params.strict_tool_verification && needsTools && !successfulToolInitially) {
      const strictMessages = appendStrictToolInstruction(params.messages);
      result = await runOnce(strictMessages);
    }

    const baseVerification = verifyToolUsage({
      needsTools,
      toolCalls: recorder.entry.tool_calls.map((call) => ({ name: call.name, ok: call.ok }))
    });

    let verification = baseVerification;

    if (params.verify_claims) {
      const claimSummary = await verifyAssistantClaims({
        assistantText: result.final,
        userText,
        attachments: params.attachments,
        toolCalls: recorder.entry.tool_calls.map((call) => ({ name: call.name, parameters: call.parameters })),
        toolExecutor: (params.verification_tool_executor ?? params.tool_executor) as any,
        maxClaims: params.verification_max_claims
      });

      if (claimSummary) {
        const warnings = [...claimSummary.warnings];
        if (needsTools && !baseVerification.used_tools) {
          warnings.unshift("Model did not use tools for a data question.");
        }
        verification = {
          needs_tools: baseVerification.needs_tools,
          used_tools: baseVerification.used_tools,
          verified: claimSummary.verified,
          confidence: claimSummary.confidence,
          warnings,
          claims: claimSummary.claims
        };
      }
    }

    recorder.setVerification(verification);

    recorder.setUserFeedback("accepted");
    return { ...result, verification };
  } catch (error) {
    recorder.setUserFeedback("rejected");
    throw error;
  } finally {
    await recorder.finalize();
  }
}

function extractTokenUsage(usage: any): TokenUsage | null {
  if (!usage || typeof usage !== "object") return null;
  const prompt = Number(usage.promptTokens ?? usage.prompt_tokens ?? 0);
  const completion = Number(usage.completionTokens ?? usage.completion_tokens ?? 0);
  if (!Number.isFinite(prompt) && !Number.isFinite(completion)) return null;
  return {
    prompt_tokens: Number.isFinite(prompt) ? prompt : 0,
    completion_tokens: Number.isFinite(completion) ? completion : 0,
    total_tokens: Number.isFinite(prompt) && Number.isFinite(completion) ? prompt + completion : undefined
  };
}

function extractToolDuration(result: any): number | undefined {
  const duration = result?.timing?.duration_ms;
  if (typeof duration === "number" && Number.isFinite(duration)) return duration;
  return undefined;
}

function sanitizeAuditToolParameters(name: string, parameters: unknown): unknown {
  if (!parameters || typeof parameters !== "object" || Array.isArray(parameters)) return parameters;
  if (name !== "fetch_external_data") return parameters;

  const params = { ...(parameters as Record<string, unknown>) };
  if (typeof params.url === "string") {
    params.url = safeUrlForAudit(params.url);
  }
  if (params.headers && typeof params.headers === "object" && !Array.isArray(params.headers)) {
    params.headers = redactHeaders(params.headers as Record<string, unknown>);
  }
  return params;
}

function safeUrlForAudit(raw: string): string {
  return redactUrlSecrets(raw);
}

function redactHeaders(headers: Record<string, unknown>): Record<string, unknown> {
  const redacted: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(headers)) {
    if (isSensitiveHeader(key)) {
      redacted[key] = "REDACTED";
    } else {
      redacted[key] = value;
    }
  }
  return redacted;
}

function isSensitiveHeader(name: string): boolean {
  const normalized = name.toLowerCase();
  if (normalized === "authorization") return true;
  if (normalized === "proxy-authorization") return true;
  if (normalized === "cookie") return true;
  if (normalized === "set-cookie") return true;
  if (normalized.includes("token")) return true;
  if (normalized.includes("secret")) return true;
  if (normalized.includes("signature")) return true;
  if (normalized.includes("api-key") || normalized.includes("apikey")) return true;
  if (normalized.endsWith("key")) return true;
  return false;
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function extractLastUserText(messages: any[]): string {
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg && typeof msg === "object" && msg.role === "user" && typeof msg.content === "string") {
      return msg.content;
    }
  }
  return "";
}

function appendStrictToolInstruction(messages: any[]): any[] {
  const strictSystemMessage = {
    role: "system",
    content: "You MUST use tools to read/compute before answering; do not guess."
  };
  let insertionIndex = 0;
  while (insertionIndex < messages.length) {
    const msg = messages[insertionIndex];
    if (msg && typeof msg === "object" && msg.role === "system") {
      insertionIndex++;
      continue;
    }
    break;
  }
  return [...messages.slice(0, insertionIndex), strictSystemMessage, ...messages.slice(insertionIndex)];
}

function compactAuditValue(value: unknown, maxChars: number): unknown {
  const limit = typeof maxChars === "number" && Number.isFinite(maxChars) && maxChars > 0 ? Math.floor(maxChars) : 20_000;
  const estimated = estimateJsonLength(value, limit);
  if (estimated <= limit) {
    // Ensure the value is actually JSON-serializable and within the budget.
    // (Tool call parameters should already be JSON-safe, but we keep this defensive.)
    try {
      const json = JSON.stringify(value);
      if (typeof json === "string" && json.length <= limit) return value;
    } catch {
      // fall through to truncation / stringification below
    }
  }

  const attempts = [
    { maxDepth: 6, maxArrayLength: 100, maxObjectKeys: 100, maxStringLength: 2_000 },
    { maxDepth: 5, maxArrayLength: 50, maxObjectKeys: 50, maxStringLength: 1_000 },
    { maxDepth: 4, maxArrayLength: 25, maxObjectKeys: 25, maxStringLength: 500 },
    { maxDepth: 3, maxArrayLength: 10, maxObjectKeys: 10, maxStringLength: 200 }
  ];

  for (const attempt of attempts) {
    const truncated = truncateUnknown(value, attempt);
    const candidate = withAuditTruncationMetadata(truncated.value, {
      truncated: true,
      // The original value exceeded the audit cap; we intentionally avoid a full
      // JSON.stringify of the input (which can be very large) and store a lower-bound
      // estimate instead.
      original_chars: estimated
    });
    const candidateJson = safeJsonStringify(candidate);
    if (candidateJson.length <= limit) return candidate;
  }

  return {
    audit_truncated: true,
    audit_original_chars: estimated,
    audit_note: "Tool call parameters exceeded audit size budget and could not be fully summarized."
  };
}

function safeJsonStringify(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    const json = JSON.stringify(value);
    return typeof json === "string" ? json : String(value);
  } catch {
    try {
      const json = JSON.stringify(String(value));
      return typeof json === "string" ? json : String(value);
    } catch {
      return String(value);
    }
  }
}

function estimateJsonLength(value: unknown, limit: number): number {
  const max = typeof limit === "number" && Number.isFinite(limit) && limit > 0 ? Math.floor(limit) : 20_000;
  let chars = 0;
  const seen = new WeakSet<object>();
  const stop = Symbol("stop");

  const add = (count: number) => {
    chars += count;
    if (chars > max) throw stop;
  };

  const estimateString = (text: string) => {
    // JSON.stringify escapes quotes, backslashes, control characters, and a couple
    // of line separators. We compute the length without allocating the escaped string.
    // Includes surrounding quotes.
    add(1); // opening quote
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      // Control chars.
      if (code < 0x20) {
        // \u00XX
        add(6);
        continue;
      }
      // Quotes / backslash.
      if (code === 0x22 /* " */ || code === 0x5c /* \\ */) {
        add(2);
        continue;
      }
      // \b \t \n \f \r
      if (code === 0x08 || code === 0x09 || code === 0x0a || code === 0x0c || code === 0x0d) {
        add(2);
        continue;
      }
      // U+2028 / U+2029 are escaped by JSON.stringify.
      if (code === 0x2028 || code === 0x2029) {
        add(6);
        continue;
      }
      add(1);
    }
    add(1); // closing quote
  };

  const walk = (v: unknown) => {
    if (v === null) {
      add(4); // null
      return;
    }
    const t = typeof v;
    if (t === "string") {
      estimateString(v as string);
      return;
    }
    if (t === "number") {
      if (!Number.isFinite(v as number)) {
        add(4); // null
      } else {
        add(String(v).length);
      }
      return;
    }
    if (t === "boolean") {
      add(v ? 4 : 5);
      return;
    }
    if (t === "undefined" || t === "function" || t === "symbol" || t === "bigint") {
      // Not JSON-serializable (or omitted); treat as null-ish placeholder.
      add(4);
      return;
    }

    // object
    if (Array.isArray(v)) {
      add(1); // [
      for (let i = 0; i < v.length; i++) {
        if (i > 0) add(1); // comma
        walk(v[i]);
      }
      add(1); // ]
      return;
    }

    const obj = v as Record<string, unknown>;
    if (seen.has(obj)) {
      add(10); // "[Circular]" (approx)
      return;
    }
    seen.add(obj);

    add(1); // {
    const keys = Object.keys(obj);
    for (let i = 0; i < keys.length; i++) {
      if (i > 0) add(1); // comma
      const key = keys[i]!;
      estimateString(key);
      add(1); // :
      walk(obj[key]);
    }
    add(1); // }
  };

  try {
    walk(value);
    return chars;
  } catch (err) {
    if (err === stop) return max + 1;
    return max + 1;
  }
}

function withAuditTruncationMetadata(value: unknown, meta: { truncated: boolean; original_chars: number }): unknown {
  if (!meta.truncated) return value;
  if (!value || typeof value !== "object" || Array.isArray(value)) return value;
  const obj = value as Record<string, unknown>;
  return {
    ...obj,
    audit_truncated: true,
    audit_original_chars: meta.original_chars
  };
}

function truncateUnknown(
  value: unknown,
  options: { maxDepth: number; maxArrayLength: number; maxObjectKeys: number; maxStringLength: number }
): { value: unknown; truncated: boolean } {
  const seen = new WeakSet<object>();

  const walk = (v: unknown, depth: number): { value: unknown; truncated: boolean } => {
    if (typeof v === "string") {
      const truncated = v.length > options.maxStringLength;
      return { value: truncateString(v, options.maxStringLength), truncated };
    }
    if (typeof v === "number" || typeof v === "boolean" || v === null) return { value: v, truncated: false };
    if (v === undefined) return { value: null, truncated: true };

    if (depth >= options.maxDepth) {
      return { value: "[truncated: max depth]", truncated: true };
    }

    if (Array.isArray(v)) {
      const out: unknown[] = [];
      let truncated = false;
      const len = Math.min(v.length, options.maxArrayLength);
      for (let i = 0; i < len; i++) {
        const child = walk(v[i], depth + 1);
        out.push(child.value);
        truncated ||= child.truncated;
      }
      if (v.length > len) {
        out.push(`[truncated: ${v.length - len} more items]`);
        truncated = true;
      }
      return { value: out, truncated };
    }

    if (typeof v === "object") {
      const obj = v as Record<string, unknown>;
      if (seen.has(obj)) return { value: "[truncated: circular]", truncated: true };
      seen.add(obj);

      const out: Record<string, unknown> = {};
      let truncated = false;
      const keys = Object.keys(obj);
      const len = Math.min(keys.length, options.maxObjectKeys);
      for (let i = 0; i < len; i++) {
        const key = keys[i]!;
        const child = walk(obj[key], depth + 1);
        out[key] = child.value;
        truncated ||= child.truncated;
      }
      if (keys.length > len) {
        out.__truncated_keys__ = keys.length - len;
        truncated = true;
      }
      return { value: out, truncated };
    }

    return { value: truncateString(String(v), options.maxStringLength), truncated: true };
  };

  return walk(value, 0);
}

function truncateString(value: string, maxLength: number): string {
  if (value.length <= maxLength) return value;
  return `${value.slice(0, maxLength)}…[truncated ${value.length - maxLength} chars]`;
}
