import { AIAuditRecorder } from "@formula/ai-audit/src/recorder.js";
import type { AIAuditStore } from "@formula/ai-audit/src/store.js";
import type { AIMode, TokenUsage } from "@formula/ai-audit/src/types.js";

import { runChatWithTools } from "../../../llm/src/toolCalling.js";

import type { LLMToolCall } from "./integration.js";

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
  on_tool_call?: (call: LLMToolCall, meta: { requiresApproval: boolean }) => void;
  on_tool_result?: (call: LLMToolCall, result: unknown) => void;
}

/**
 * Run the provider-agnostic tool-calling loop with audit logging.
 *
 * This is intended as a thin integration helper to wire `packages/llm` + `packages/ai-tools`
 * to `packages/ai-audit` without duplicating orchestration logic across UI surfaces.
 */
export async function runChatWithToolsAudited(params: AuditedRunParams): Promise<{ messages: any[]; final: string }> {
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
    const result = await runChatWithTools({
      client: auditedClient as any,
      toolExecutor: params.tool_executor as any,
      messages: params.messages as any,
      maxIterations: params.max_iterations,
      onToolCall: (call: any, meta: any) => {
        recorder.recordToolCall({
          id: call.id,
          name: call.name,
          parameters: call.arguments,
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

    recorder.setUserFeedback("accepted");
    return result;
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

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}
