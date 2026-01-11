import { AIAuditRecorder } from "@formula/ai-audit";
import type { AIAuditStore, AIMode, TokenUsage } from "@formula/ai-audit";

import { runChatWithTools } from "../../../llm/src/toolCalling.js";

import type { LLMToolCall } from "./integration.js";
import { classifyQueryNeedsTools, verifyToolUsage, type VerificationResult } from "./verification.js";

export interface AuditedRunOptions {
  audit_store: AIAuditStore;
  session_id: string;
  user_id?: string;
  mode: AIMode;
  input: unknown;
  model: string;
}

export interface AuditedRunParams {
  client: { chat: (request: any) => Promise<any> };
  tool_executor: { tools: any[]; execute: (call: any) => Promise<any> };
  messages: any[];
  audit: AuditedRunOptions;
  max_iterations?: number;
  require_approval?: (call: LLMToolCall) => Promise<boolean>;
  /**
   * When true, approval denials are surfaced to the model as tool results (ok:false)
   * and the loop continues, allowing the model to re-plan. Default behavior is to throw.
   */
  continue_on_approval_denied?: boolean;
  /**
   * Optional hooks for UI surfaces (e.g. chat panels) that still want to surface
   * tool call + result events while relying on this helper for audit logging.
   */
  on_tool_call?: (call: LLMToolCall, meta: { requiresApproval: boolean }) => void;
  on_tool_result?: (call: LLMToolCall, result: unknown) => void;
  /**
   * Optional LLM request parameters forwarded to `runChatWithTools`.
   * If omitted, `audit.model` is used as the request model (providers may ignore it).
   */
  model?: string;
  temperature?: number;
  max_tokens?: number;
  /**
   * When enabled, if the query is classified as needing data tools but the model
   * produced an answer without using any tools, the run is retried once with a
   * stricter system instruction to "use tools; do not guess".
   */
  strict_tool_verification?: boolean;
  /**
   * Optional attachment context used by the verifier (range/table/chart/formula references).
   */
  attachments?: unknown[] | null;
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
    }
  };

  try {
    const userText = extractLastUserText(params.messages);
    const needsTools = classifyQueryNeedsTools({ userText, attachments: params.attachments });

    const runOnce = async (messages: any[]) =>
      runChatWithTools({
        client: auditedClient as any,
        toolExecutor: params.tool_executor as any,
        messages: messages as any,
        maxIterations: params.max_iterations,
        continueOnApprovalDenied: params.continue_on_approval_denied,
        model: params.model ?? params.audit.model,
        temperature: params.temperature,
        maxTokens: params.max_tokens,
        onToolCall: (call: any, meta: any) => {
          recorder.recordToolCall({
            id: call.id,
            name: call.name,
            parameters: sanitizeAuditToolParameters(call.name, call.arguments),
            requires_approval: Boolean(meta?.requiresApproval)
          });
          params.on_tool_call?.(call, meta);
        },
        onToolResult: (call: any, toolResult: any) => {
          recorder.recordToolResult(call.id, {
            ok: typeof toolResult?.ok === "boolean" ? toolResult.ok : undefined,
            duration_ms: extractToolDuration(toolResult),
            result: toolResult,
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
    const usedToolsInitially = recorder.entry.tool_calls.length > 0;

    if (params.strict_tool_verification && needsTools && !usedToolsInitially) {
      const strictMessages = appendStrictToolInstruction(params.messages);
      result = await runOnce(strictMessages);
    }

    const verification = verifyToolUsage({
      needsTools,
      toolCalls: recorder.entry.tool_calls.map((call) => ({ name: call.name, ok: call.ok }))
    });

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
  try {
    const url = new URL(raw);
    url.username = "";
    url.password = "";
    url.hash = "";

    if (url.search) {
      const params = new URLSearchParams(url.search);
      const keys = Array.from(new Set(Array.from(params.keys())));
      for (const key of keys) {
        if (!isSensitiveQueryParam(key)) continue;
        const count = params.getAll(key).length;
        params.delete(key);
        for (let i = 0; i < count; i++) params.append(key, "REDACTED");
      }
      const next = params.toString();
      url.search = next ? `?${next}` : "";
    }

    return url.toString();
  } catch {
    return raw;
  }
}

function isSensitiveQueryParam(key: string): boolean {
  const normalized = key.toLowerCase();
  return (
    normalized === "key" ||
    normalized === "api_key" ||
    normalized === "apikey" ||
    normalized === "token" ||
    normalized === "access_token" ||
    normalized === "auth" ||
    normalized === "authorization" ||
    normalized === "signature" ||
    normalized === "sig" ||
    normalized === "password" ||
    normalized === "secret"
  );
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
