import type { DocumentController } from "../../document/documentController.js";

import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";

import type { AIAuditStore } from "../../../../../packages/ai-audit/src/store.js";
import type { PreviewEngineOptions, ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import { PreviewEngine } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import { runChatWithToolsAudited } from "../../../../../packages/ai-tools/src/llm/audited-run.js";
import { SpreadsheetLLMToolExecutor } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import type { TokenEstimator } from "../../../../../packages/ai-context/src/tokenBudget.js";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens } from "../../../../../packages/ai-context/src/tokenBudget.js";
import { trimMessagesToBudget } from "../../../../../packages/ai-context/src/trimMessagesToBudget.js";
import type { LLMClient, ToolCall } from "../../../../../packages/llm/src/index.js";
import { DlpViolationError } from "../../../../../packages/security/dlp/src/errors.js";

import { maybeGetAiCloudDlpOptions } from "../dlp/aiDlp.js";
import { computeDlpCacheKey } from "../dlp/dlpCacheKey.js";
import { getDesktopToolPolicy } from "../toolPolicy.js";

import { createDesktopRagService, type DesktopRagService, type DesktopRagServiceOptions } from "../rag/ragService.js";
import { getDesktopAIAuditStore } from "../audit/auditStore.js";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../contextBudget.js";
import { WorkbookContextBuilder, type WorkbookContextBuildStats, type WorkbookSchemaProvider } from "../context/WorkbookContextBuilder.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";

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
  /**
   * Optional resolver that maps user-facing sheet names (display names) to stable
   * DocumentController sheet ids.
   */
  sheetNameResolver?: SheetNameResolver | null;
  llmClient: LLMClient;
  schemaProvider?: WorkbookSchemaProvider | null;
  auditStore?: AIAuditStore;
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
  /**
   * When true, approval denials are returned to the model as tool results and
   * the agent continues running, allowing the model to re-plan.
   *
   * Default is false (stop safely on denial).
   */
  continueOnApprovalDenied?: boolean;

  onProgress?: (event: AgentProgressEvent) => void;
  onApprovalRequired?: (request: AgentApprovalRequest) => Promise<boolean>;

  maxIterations?: number;
  maxDurationMs?: number;
  signal?: AbortSignal;
  model?: string;

  ragService?: DesktopRagService;
  /**
   * Options for the default desktop workbook RAG service (if `ragService` is not provided).
   *
   * Note: Desktop workbook RAG uses deterministic hash embeddings by design
   * (offline; no API keys / local model setup).
   */
  ragOptions?: Omit<DesktopRagServiceOptions, "documentController" | "workbookId">;
  /**
   * Optional override for the model context window used to budget prompts.
   * If omitted, a best-effort default is derived from `model`.
   */
  contextWindowTokens?: number;
  reserveForOutputTokens?: number;
  keepLastMessages?: number;
  tokenEstimator?: TokenEstimator;

  /**
   * Optional hook for workbook context build telemetry. When provided, this will
   * enable `WorkbookContextBuilder` instrumentation and invoke this callback for
   * each context build done by the agent loop.
   *
   * NOTE: By default, build stats are only logged in dev builds.
   */
  onWorkbookContextBuildStats?: (stats: WorkbookContextBuildStats) => void;
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

export async function runAgentTask(params: RunAgentTaskParams): Promise<AgentTaskResult> {
  const goal = params.goal.trim();
  if (!goal) {
    return { status: "error", session_id: createSessionId("agent"), error: "Goal is required." };
  }

  const auditStore = params.auditStore ?? getDesktopAIAuditStore();

  const maxIterations = params.maxIterations ?? 20;
  const maxDurationMs = params.maxDurationMs ?? 5 * 60 * 1000;
  const startedAt = nowMs();
  const sessionId = createSessionId("agent");

  let iteration = 0;
  let deniedCall: ToolCall | undefined;

  const signal = params.signal;
  const ownedRagService =
    params.ragService == null
      ? createDesktopRagService({
          documentController: params.documentController,
          workbookId: params.workbookId,
          ...(params.ragOptions ?? {}),
          ...(params.tokenEstimator ? { tokenEstimator: params.tokenEstimator } : {})
        })
      : null;
  const ragService = (params.ragService ?? ownedRagService)!;

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

  async function guard<T>(promise: Promise<T>, abortController?: AbortController): Promise<T> {
    throwIfCancelled();

    const timeoutMs = remainingMs();
    let abortListener: (() => void) | null = null;
    let timeoutId: ReturnType<typeof setTimeout> | null = null;

    const abortPromise =
      signal == null
        ? null
        : new Promise<never>((_resolve, reject) => {
            abortListener = () => {
              try {
                abortController?.abort();
              } catch {
                // ignore
              }
              reject(new AgentCancelledError("Agent run aborted"));
            };
            signal.addEventListener("abort", abortListener, { once: true });
            // Handle the race where the signal aborts before registering the listener.
            if (signal.aborted) abortListener();
          });

    const timeoutPromise =
      Number.isFinite(timeoutMs) && timeoutMs > 0
        ? new Promise<never>((_resolve, reject) => {
            timeoutId = setTimeout(() => {
              try {
                abortController?.abort();
              } catch {
                // ignore
              }
              reject(new AgentTimeoutError(`Agent run exceeded maxDurationMs (${maxDurationMs})`));
            }, timeoutMs);
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

    const estimator = params.tokenEstimator ?? createHeuristicTokenEstimator();
    const modelName = params.model ?? "unknown";
    const contextWindowTokens = params.contextWindowTokens ?? getModeContextWindowTokens("agent", modelName);
    const reserveForOutputTokens =
      params.reserveForOutputTokens ?? getDefaultReserveForOutputTokens("agent", contextWindowTokens);
    const keepLastMessages = params.keepLastMessages ?? 60;

    const defaultSheetId = params.defaultSheetId ?? "Sheet1";
    const spreadsheet = new DocumentControllerSpreadsheetApi(params.documentController, {
      createChart: params.createChart,
      sheetNameResolver: params.sheetNameResolver ?? null
    });
    const toolPolicy = getDesktopToolPolicy({ mode: "agent" });

    let dlp =
      maybeGetAiCloudDlpOptions({
        documentId: params.workbookId,
        sheetId: defaultSheetId,
        sheetNameResolver: params.sheetNameResolver,
      }) ?? undefined;
    let dlpKey = computeDlpCacheKey(dlp);
    const devOnBuildStats =
      import.meta.env.MODE === "development"
        ? (stats: WorkbookContextBuildStats) => {
            try {
              console.debug("[ai] WorkbookContextBuilder build stats (agent)", stats);
            } catch {
              // ignore
            }
          }
        : undefined;
    const onBuildStats =
      devOnBuildStats || params.onWorkbookContextBuildStats
        ? (stats: WorkbookContextBuildStats) => {
            devOnBuildStats?.(stats);
            params.onWorkbookContextBuildStats?.(stats);
          }
        : undefined;

    function refreshDlpIfNeeded(): void {
      const next =
        maybeGetAiCloudDlpOptions({
          documentId: params.workbookId,
          sheetId: defaultSheetId,
          sheetNameResolver: params.sheetNameResolver,
        }) ?? undefined;
      const nextKey = computeDlpCacheKey(next);
      if (nextKey === dlpKey) return;
      dlp = next;
      dlpKey = nextKey;
    }

    let toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
      default_sheet: defaultSheetId,
      sheet_name_resolver: params.sheetNameResolver ?? null,
      require_approval_for_mutations: true,
      dlp,
      toolPolicy
    });
    let toolExecutorDlpKey = dlpKey;
    const offeredTools = toolExecutor.tools.map((t) => t.name);
    const previewEngine = new PreviewEngine({ approval_cell_threshold: 0, ...(params.previewOptions ?? {}) });

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

    const contextBuilder = new WorkbookContextBuilder({
      workbookId: params.workbookId,
      documentController: params.documentController,
      spreadsheet,
      ragService,
      schemaProvider: params.schemaProvider ?? null,
      sheetNameResolver: params.sheetNameResolver ?? null,
      mode: "agent",
      model: modelName,
      contextWindowTokens,
      reserveForOutputTokens,
      tokenEstimator: estimator as any,
      onBuildStats
    });

    async function refreshSystemMessage(targetMessages: any[]): Promise<void> {
      throwIfCancelled();
      refreshDlpIfNeeded();
      if (toolExecutorDlpKey !== dlpKey) {
        toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
          default_sheet: defaultSheetId,
          sheet_name_resolver: params.sheetNameResolver ?? null,
          require_approval_for_mutations: true,
          dlp,
          toolPolicy
        });
        toolExecutorDlpKey = dlpKey;
      }
      const contextAbortController = new AbortController();
      const ctx = await guard(
        contextBuilder.build({
          activeSheetId: defaultSheetId,
          dlp,
          focusQuestion: goal,
          signal: contextAbortController.signal
        }),
        contextAbortController
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

        const toolTokens = estimateToolDefinitionTokens(request?.tools as any, estimator);
        const maxMessageTokens = Math.max(0, contextWindowTokens - toolTokens);
        const trimmed = await trimMessagesToBudget({
          messages: request.messages as any,
          maxTokens: maxMessageTokens,
          reserveForOutputTokens,
          estimator,
          keepLastMessages,
          signal
        });
        if (Array.isArray(request.messages)) {
          const next = trimmed === request.messages ? trimmed.slice() : trimmed;
          request.messages.length = 0;
          request.messages.push(...next);
        } else {
          request.messages = trimmed;
        }

        const response = await guard(params.llmClient.chat({ ...request, model: request.model ?? params.model }));
        const content = response?.message?.content;
        if (typeof content === "string" && content.trim()) {
          emit({ type: "assistant_message", iteration, content });
        }
        return response;
      },
      streamChat: params.llmClient.streamChat
        ? async function* streamChat(request: any) {
            iteration += 1;
            emit({ type: "planning", iteration });
            await refreshSystemMessage(request.messages);

            const toolTokens = estimateToolDefinitionTokens(request?.tools as any, estimator);
            const maxMessageTokens = Math.max(0, contextWindowTokens - toolTokens);
            const trimmed = await trimMessagesToBudget({
              messages: request.messages as any,
              maxTokens: maxMessageTokens,
              reserveForOutputTokens,
              estimator,
              keepLastMessages,
              signal
            });
            if (Array.isArray(request.messages)) {
              const next = trimmed === request.messages ? trimmed.slice() : trimmed;
              request.messages.length = 0;
              request.messages.push(...next);
            } else {
              request.messages = trimmed;
            }

            let content = "";
            for await (const event of params.llmClient.streamChat!({ ...request, model: request.model ?? params.model })) {
              if (event?.type === "text" && typeof event.delta === "string" && event.delta.length > 0) {
                content += event.delta;
                if (content.trim()) emit({ type: "assistant_message", iteration, content });
              }
              yield event;
            }
          }
        : undefined
    };

    const wrappedToolExecutor = {
      get tools() {
        return toolExecutor.tools;
      },
      async execute(call: any) {
        throwIfCancelled();
        // Refresh DLP before executing the tool call so we don't leak workbook data
        // if policies/classifications change during a long-running agent task.
        refreshDlpIfNeeded();
        if (toolExecutorDlpKey !== dlpKey) {
          toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
            default_sheet: defaultSheetId,
            sheet_name_resolver: params.sheetNameResolver ?? null,
            require_approval_for_mutations: true,
            dlp,
            toolPolicy
          });
          toolExecutorDlpKey = dlpKey;
        }
        return guard(toolExecutor.execute(call));
      }
    };

    const requireApproval = async (call: ToolCall) => {
      throwIfCancelled();
      const preview = await guard(
        previewEngine.generatePreview([{ name: call.name, parameters: call.arguments }], spreadsheet, {
          default_sheet: defaultSheetId,
          sheet_name_resolver: params.sheetNameResolver ?? null
        })
      );
      if (!preview.requires_approval) return true;
      if (!params.onApprovalRequired) {
        deniedCall = call;
        return false;
      }

      const approved = await guard(params.onApprovalRequired({ call, preview }));
      if (!approved && !params.continueOnApprovalDenied) deniedCall = call;
      return approved;
    };

    const result = await guard(
      runChatWithToolsAudited({
        client: wrappedClient,
        tool_executor: wrappedToolExecutor,
        messages,
        signal,
        audit: {
          audit_store: auditStore,
          session_id: sessionId,
          workbook_id: params.workbookId,
          mode: "agent",
          input: { goal, constraints: params.constraints ?? [], workbookId: params.workbookId, offered_tools: offeredTools },
          model: params.model ?? "unknown"
        },
        max_iterations: maxIterations,
        require_approval: requireApproval,
        continue_on_approval_denied: params.continueOnApprovalDenied,
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
    if (error instanceof DlpViolationError) {
      const message = error.message || "Operation blocked by data loss prevention policy.";
      emit({ type: "error", iteration, message, error });
      return { status: "error", session_id: sessionId, error: message };
    }

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
  } finally {
    try {
      await ownedRagService?.dispose();
    } catch {
      // ignore
    }
  }
}
