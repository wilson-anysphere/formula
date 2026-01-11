import type { DocumentController } from "../../document/documentController.js";

import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";

import type { AIAuditStore } from "../../../../../packages/ai-audit/src/store.js";
import type { PreviewEngineOptions, ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import { PreviewEngine } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import { runChatWithToolsAudited } from "../../../../../packages/ai-tools/src/llm/audited-run.js";
import { SpreadsheetLLMToolExecutor } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../../../../packages/ai-rag/src/index.js";
import type { LLMClient, ToolCall } from "../../../../../packages/llm/src/types.js";

export interface AgentApprovalRequest {
  call: ToolCall;
  preview: ToolPlanPreview;
}

export type AgentProgressEvent =
  | {
      type: "planning";
      iteration: number;
    }
  | {
      type: "tool_call";
      iteration: number;
      call: ToolCall;
      requiresApproval: boolean;
    }
  | {
      type: "tool_result";
      iteration: number;
      call: ToolCall;
      result: unknown;
      ok?: boolean;
      error?: string;
    }
  | {
      type: "assistant_message";
      iteration: number;
      content: string;
    }
  | {
      type: "complete";
      iteration: number;
      result: AgentTaskResult;
    }
  | {
      type: "cancelled";
      iteration: number;
      reason: "abort" | "timeout" | "approval_denied";
      message?: string;
    }
  | {
      type: "error";
      iteration: number;
      message: string;
      error: unknown;
    };

export type AgentTaskStatus = "complete" | "needs_approval" | "cancelled" | "error";

export interface AgentTaskResult {
  status: AgentTaskStatus;
  session_id: string;
  final?: string;
  messages?: any[];
  error?: string;
  denied_call?: ToolCall;
}

export interface RunAgentTaskParams {
  goal: string;
  constraints?: string[];
  workbookId: string;
  documentController: DocumentController;
  llmClient: LLMClient;
  auditStore: AIAuditStore;
  /**
   * Default sheet used when tool calls omit a sheet prefix (e.g. "A1" instead of "Sheet2!A1").
   * Defaults to "Sheet1".
   */
  defaultSheetId?: string;
  /**
   * Optional host capability for chart creation (enables the `create_chart` tool).
   */
  createChart?: SpreadsheetApi["createChart"];
  /**
   * Preview engine configuration for approval gating.
   *
   * NOTE: Agent mode defaults to `approval_cell_threshold: 0` so any non-noop mutation
   * requires explicit user approval.
   */
  previewOptions?: PreviewEngineOptions;

  onProgress?: (event: AgentProgressEvent) => void;
  onApprovalRequired?: (request: AgentApprovalRequest) => Promise<boolean>;

  maxIterations?: number;
  maxDurationMs?: number;
  signal?: AbortSignal;
  model?: string;
}

class AgentCancelledError extends Error {
  override name = "AbortError";
}

class AgentTimeoutError extends Error {
  override name = "TimeoutError";
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function createSessionId(prefix: string): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto && typeof crypto.randomUUID === "function") {
    return `${prefix}_${crypto.randomUUID()}`;
  }
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function isAbortError(error: unknown): boolean {
  return (
    error instanceof AgentCancelledError ||
    (error instanceof Error && (error.name === "AbortError" || /aborted/i.test(error.message)))
  );
}

function isTimeoutError(error: unknown): boolean {
  return error instanceof AgentTimeoutError || (error instanceof Error && error.name === "TimeoutError");
}

function isApprovalDeniedError(error: unknown): boolean {
  return error instanceof Error && /requires approval and was denied/i.test(error.message);
}

function createContextManager(): ContextManager {
  const dimension = 128;
  const embedder = new HashEmbedder({ dimension });
  const vectorStore = new InMemoryVectorStore({ dimension });
  return new ContextManager({
    tokenBudgetTokens: 8_000,
    workbookRag: { vectorStore, embedder, topK: 6, sampleRows: 6 }
  });
}

export async function runAgentTask(params: RunAgentTaskParams): Promise<AgentTaskResult> {
  const goal = params.goal.trim();
  if (!goal) {
    return { status: "error", session_id: createSessionId("agent"), error: "Goal is required." };
  }

  const maxIterations = params.maxIterations ?? 20;
  const maxDurationMs = params.maxDurationMs ?? 5 * 60 * 1000;
  const startedAt = nowMs();
  const sessionId = createSessionId("agent");

  let iteration = 0;
  let deniedCall: ToolCall | undefined;

  const signal = params.signal;

  function emit(event: AgentProgressEvent) {
    try {
      params.onProgress?.(event);
    } catch {
      // Ignore progress handler failures; agent execution should not be blocked by UI issues.
    }
  }

  function remainingMs(): number {
    return Math.max(0, maxDurationMs - (nowMs() - startedAt));
  }

  function throwIfCancelled(): void {
    if (signal?.aborted) throw new AgentCancelledError("Agent run aborted");
    if (remainingMs() <= 0) throw new AgentTimeoutError(`Agent run exceeded maxDurationMs (${maxDurationMs})`);
  }

  async function guard<T>(promise: Promise<T>): Promise<T> {
    throwIfCancelled();

    const timeoutMs = remainingMs();
    let abortListener: (() => void) | null = null;
    let timeoutId: ReturnType<typeof setTimeout> | null = null;

    const abortPromise =
      signal == null
        ? null
        : new Promise<never>((_resolve, reject) => {
            abortListener = () => reject(new AgentCancelledError("Agent run aborted"));
            signal.addEventListener("abort", abortListener, { once: true });
          });

    const timeoutPromise =
      Number.isFinite(timeoutMs) && timeoutMs > 0
        ? new Promise<never>((_resolve, reject) => {
            timeoutId = setTimeout(
              () => reject(new AgentTimeoutError(`Agent run exceeded maxDurationMs (${maxDurationMs})`)),
              timeoutMs
            );
          })
        : Promise.reject(new AgentTimeoutError(`Agent run exceeded maxDurationMs (${maxDurationMs})`));

    try {
      return await Promise.race([promise, abortPromise, timeoutPromise].filter(Boolean) as Promise<T>[]);
    } finally {
      if (abortListener && signal) signal.removeEventListener("abort", abortListener);
      if (timeoutId != null) clearTimeout(timeoutId);
    }
  }

  try {
    throwIfCancelled();

    const defaultSheetId = params.defaultSheetId ?? "Sheet1";
    const spreadsheet = new DocumentControllerSpreadsheetApi(params.documentController, { createChart: params.createChart });
    const toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
      default_sheet: defaultSheetId,
      require_approval_for_mutations: true
    });
    const previewEngine = new PreviewEngine({ approval_cell_threshold: 0, ...(params.previewOptions ?? {}) });
    const contextManager = createContextManager();

    const userMessage = [
      `Goal: ${goal}`,
      ...(params.constraints?.length ? ["", "Constraints:", ...params.constraints.map((c) => `- ${c}`)] : []),
      "",
      "Work autonomously using the provided spreadsheet tools. Read data before making claims. At the end, summarize what you did."
    ].join("\n");

    const messages: any[] = [
      {
        role: "system",
        content: "You are an autonomous spreadsheet agent."
      },
      { role: "user", content: userMessage }
    ];

    function buildSystemPrompt(workbookContext: string): string {
      return [
        "You are an autonomous spreadsheet agent running inside a spreadsheet application.",
        "Use spreadsheet tools to inspect and modify the workbook to achieve the user's goal.",
        "Mutating actions are approval-gated; propose minimal, safe changes.",
        "",
        `Goal: ${goal}`,
        ...(params.constraints?.length ? ["", `Constraints:\n${params.constraints.map((c) => `- ${c}`).join("\n")}`] : []),
        "",
        "Workbook context (RAG):",
        workbookContext || "(no workbook context available)"
      ].join("\n");
    }

    async function refreshSystemMessage(targetMessages: any[]): Promise<void> {
      const ctx = await guard(
        contextManager.buildWorkbookContextFromSpreadsheetApi({
          spreadsheet,
          workbookId: params.workbookId,
          query: goal
        })
      );
      if (!targetMessages.length || targetMessages[0]?.role !== "system") {
        targetMessages.unshift({ role: "system", content: buildSystemPrompt(ctx.promptContext) });
      } else {
        targetMessages[0].content = buildSystemPrompt(ctx.promptContext);
      }
    }

    const wrappedClient = {
      async chat(request: any) {
        iteration += 1;
        emit({ type: "planning", iteration });
        await refreshSystemMessage(request.messages);
        const response = await guard(params.llmClient.chat({ ...request, model: request.model ?? params.model }));
        const content = response?.message?.content;
        if (typeof content === "string" && content.trim()) {
          emit({ type: "assistant_message", iteration, content });
        }
        return response;
      }
    };

    const wrappedToolExecutor = {
      tools: toolExecutor.tools,
      async execute(call: any) {
        throwIfCancelled();
        return guard(toolExecutor.execute(call));
      }
    };

    const requireApproval = async (call: ToolCall) => {
      throwIfCancelled();
      const preview = await guard(
        previewEngine.generatePreview([{ name: call.name, parameters: call.arguments }], spreadsheet, {
          default_sheet: defaultSheetId
        })
      );
      if (!preview.requires_approval) return true;
      if (!params.onApprovalRequired) {
        deniedCall = call;
        return false;
      }

      const approved = await guard(params.onApprovalRequired({ call, preview }));
      if (!approved) deniedCall = call;
      return approved;
    };

    const result = await guard(
      runChatWithToolsAudited({
        client: wrappedClient,
        tool_executor: wrappedToolExecutor,
        messages,
        audit: {
          audit_store: params.auditStore,
          session_id: sessionId,
          mode: "agent",
          input: { goal, constraints: params.constraints ?? [], workbookId: params.workbookId },
          model: params.model ?? "unknown"
        },
        max_iterations: maxIterations,
        require_approval: requireApproval,
        on_tool_call: (call: ToolCall, meta: { requiresApproval: boolean }) => {
          emit({ type: "tool_call", iteration, call, requiresApproval: meta.requiresApproval });
        },
        on_tool_result: (call: ToolCall, toolResult: any) => {
          emit({
            type: "tool_result",
            iteration,
            call,
            result: toolResult,
            ok: typeof toolResult?.ok === "boolean" ? toolResult.ok : undefined,
            error: toolResult?.error?.message ? String(toolResult.error.message) : undefined
          });
        }
      })
    );

    const finalResult: AgentTaskResult = {
      status: "complete",
      session_id: sessionId,
      final: result.final,
      messages: result.messages
    };
    emit({ type: "complete", iteration, result: finalResult });
    return finalResult;
  } catch (error) {
    const reason = isTimeoutError(error)
      ? ("timeout" as const)
      : deniedCall || isApprovalDeniedError(error)
        ? ("approval_denied" as const)
        : isAbortError(error)
          ? ("abort" as const)
          : null;

    if (reason) {
      emit({
        type: "cancelled",
        iteration,
        reason,
        message: error instanceof Error ? error.message : String(error)
      });
      return {
        status: deniedCall ? "needs_approval" : "cancelled",
        session_id: sessionId,
        denied_call: deniedCall,
        error: error instanceof Error ? error.message : String(error)
      };
    }

    emit({
      type: "error",
      iteration,
      message: error instanceof Error ? error.message : String(error),
      error
    });

    return {
      status: "error",
      session_id: sessionId,
      error: error instanceof Error ? error.message : String(error)
    };
  }
}
