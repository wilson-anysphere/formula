import type { ChatStreamEvent, LLMClient, LLMMessage } from "../../../../../packages/llm/src/index.js";

import type { AIAuditStore } from "../../../../../packages/ai-audit/src/store.js";
import type { AIAuditEntry, AuditListFilters } from "../../../../../packages/ai-audit/src/types.js";
import { AIAuditRecorder } from "../../../../../packages/ai-audit/src/recorder.js";

import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import type { TokenEstimator } from "../../../../../packages/ai-context/src/tokenBudget.js";
import {
  createHeuristicTokenEstimator,
  estimateToolDefinitionTokens,
} from "../../../../../packages/ai-context/src/tokenBudget.js";
import { trimMessagesToBudget } from "../../../../../packages/ai-context/src/trimMessagesToBudget.js";

import { rectToA1 } from "../../../../../packages/ai-rag/src/workbook/rect.js";

import type { ToolExecutionResult } from "../../../../../packages/ai-tools/src/executor/tool-executor.js";
import type {
  LLMToolCall,
  PreviewApprovalRequest,
  SpreadsheetLLMToolExecutorOptions
} from "../../../../../packages/ai-tools/src/llm/integration.js";
import { SpreadsheetLLMToolExecutor, createPreviewApprovalHandler } from "../../../../../packages/ai-tools/src/llm/integration.js";
import { runChatWithToolsAuditedVerified } from "../../../../../packages/ai-tools/src/llm/audited-run.js";
import type { PreviewEngineOptions, ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import { parseA1Range } from "../../../../../packages/ai-tools/src/spreadsheet/a1.ts";

import { DlpViolationError } from "../../../../../packages/security/dlp/src/errors.js";

import type { DocumentController } from "../../document/documentController.js";
import type { Range } from "../../selection/types";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";

import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";
import { createDesktopRagService, type DesktopRagService, type DesktopRagServiceOptions } from "../rag/ragService.js";
import { getDesktopAIAuditStore } from "../audit/auditStore.js";
import { maybeGetAiCloudDlpOptions } from "../dlp/aiDlp.js";
import { computeDlpCacheKey } from "../dlp/dlpCacheKey.js";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../contextBudget.js";
import { getDesktopToolPolicy } from "../toolPolicy.js";
import { WorkbookContextBuilder, type WorkbookContextBuildStats, type WorkbookSchemaProvider } from "../context/WorkbookContextBuilder.js";

export type AiChatAttachment =
  | { type: "range"; reference: string; data?: unknown }
  | { type: "formula"; reference: string; data?: { formula: string } }
  | { type: "table"; reference: string; data?: unknown }
  | { type: "chart"; reference: string; data?: unknown };

const MAX_ATTACHMENT_DATA_CHARS_FOR_PROMPT = 2_000;
const MAX_ATTACHMENT_DATA_CHARS_FOR_AUDIT = 2_000;

function selectionFromAttachments(
  attachments: AiChatAttachment[],
  defaultSheetId: string,
  sheetNameResolver?: SheetNameResolver | null,
): { sheetId: string; range: Range } | undefined {
  const rangeAttachment = attachments.find((a) => a.type === "range" && typeof a.reference === "string");
  if (!rangeAttachment) return undefined;

  try {
    const parsed = parseA1Range(rangeAttachment.reference, defaultSheetId);
    const sheetId = sheetNameResolver?.getSheetIdByName(parsed.sheet) ?? parsed.sheet;
    return {
      sheetId,
      range: {
        startRow: parsed.startRow - 1,
        endRow: parsed.endRow - 1,
        startCol: parsed.startCol - 1,
        endCol: parsed.endCol - 1,
      },
    };
  } catch {
    return undefined;
  }
}

export interface SendAiChatMessageParams {
  text: string;
  attachments?: AiChatAttachment[];
  /**
   * Prior LLM messages (excluding the system prompt). Returned from a prior
   * `sendMessage` call.
   */
  history?: LLMMessage[];
  /**
   * Optional hook for UI surfaces that want to display tool calls as they happen.
   */
  onToolCall?: (call: LLMToolCall, meta: { requiresApproval: boolean }) => void;
  /**
   * Optional hook for displaying tool results as they return.
   */
  onToolResult?: (call: LLMToolCall, result: ToolExecutionResult) => void;
  /**
   * Optional hook for rendering incremental assistant output.
   */
  onStreamEvent?: (event: ChatStreamEvent) => void;
  /**
   * Optional abort signal (cancels streaming + tool calling).
   */
  signal?: AbortSignal;
}

export interface SendAiChatMessageResult {
  finalText: string;
  messages: LLMMessage[];
  toolResults: ToolExecutionResult[];
  verification?: unknown;
  context: {
    workbookId: string;
    promptContext: string;
    retrievedChunkIds: string[];
    retrievedRanges: string[];
    retrieved: unknown[];
    indexStats?: unknown;
    tokenBudgetTokens?: number;
  };
  auditEntryId?: string;
  sessionId: string;
}

export class AiChatOrchestratorError extends Error {
  readonly sessionId: string;
  readonly auditEntryId?: string;

  constructor(message: string, params: { sessionId: string; auditEntryId?: string; cause?: unknown }) {
    super(message, { cause: params.cause } as any);
    this.name = "AiChatOrchestratorError";
    this.sessionId = params.sessionId;
    this.auditEntryId = params.auditEntryId;
  }
}

export interface AiChatOrchestratorOptions {
  documentController: DocumentController;
  workbookId: string;
  llmClient: LLMClient;
  model: string;
  /**
   * When enabled (default), if the user's query is classified as needing tools
   * (e.g. it references cells/ranges) but the model answers without any tool
   * calls, we retry once with a stricter system instruction that forces tool
   * usage ("do not guess").
   *
   * This plugs a common failure mode where models respond confidently about
   * workbook data without reading it.
   */
  strictToolVerification?: boolean;
  /**
   * Optional resolver that maps user-facing sheet names (display names) to stable
   * DocumentController sheet ids.
   *
   * When provided, AI tool calls can safely reference renamed sheets without
   * accidentally creating phantom sheets in the DocumentController model.
   */
  sheetNameResolver?: SheetNameResolver | null;

  getActiveSheetId?: () => string;
  /**
   * Optional UI selection hook. When provided, the chat orchestrator will include a
   * selection data block in the workbook context (unless a range attachment is
   * explicitly provided, which takes precedence).
   *
   * Expected coordinates are 0-based (desktop `Range`).
   */
  getSelectedRange?: () => { sheetId: string; range: Range } | null;
  /**
   * Optional provider for workbook metadata not exposed via SpreadsheetApi (named
   * ranges, explicit tables, etc).
   */
  schemaProvider?: WorkbookSchemaProvider | null;
  /**
   * Optional chart host implementation. When provided, tool calls like
   * `create_chart` will add a chart to the desktop UI (via SpreadsheetApi
   * integration).
   */
  createChart?: SpreadsheetApi["createChart"];

  /**
   * If not provided, defaults to the desktop audit store (sqlite-backed with
   * LocalStorage fallback).
   */
  auditStore?: AIAuditStore;
  sessionId?: string;

  /**
   * Context builder used to produce schema-first + RAG workbook context per message.
   *
   * If omitted, the orchestrator will create a default desktop RAG service backed
   * by a persistent sqlite vector store (stored in LocalStorage) and
   * deterministic hash embeddings (offline; no API keys / local model setup).
   */
  contextManager?: ContextManager;
  ragService?: DesktopRagService;
  ragOptions?: Omit<DesktopRagServiceOptions, "documentController" | "workbookId">;

  systemPrompt?: string;

  /**
   * Approval callback when a tool preview flags risk. Safe default: if a tool
   * requires approval and this callback is not provided, the orchestrator will
   * deny the tool call.
   */
  onApprovalRequired?: (request: { call: LLMToolCall; preview: ToolPlanPreview }) => Promise<boolean>;

  previewOptions?: PreviewEngineOptions;

  /**
   * Additional tool executor configuration (e.g. external data permissions).
   * `default_sheet` is supplied automatically per message.
   */
  toolExecutorOptions?: Omit<SpreadsheetLLMToolExecutorOptions, "default_sheet" | "require_approval_for_mutations">;

  /**
   * Optional override for the model context window used to budget prompts.
   * If omitted, a best-effort default is derived from `model`.
   */
  contextWindowTokens?: number;
  /**
   * Tokens to reserve for the model's completion. Used when trimming messages to
   * avoid "prompt too long" errors from providers.
   */
  reserveForOutputTokens?: number;
  /**
   * Count-based cap: keep at most the most recent N non-system messages even if
   * they would fit under the token budget.
   */
  keepLastMessages?: number;
  /**
   * Token estimator used for context budgeting. Defaults to a lightweight
   * heuristic (4 chars/token) but can be overridden with provider-specific
   * tokenizers.
   */
  tokenEstimator?: TokenEstimator;

  /**
   * Optional hook for workbook context build telemetry. When provided, this will
   * enable `WorkbookContextBuilder` instrumentation and invoke this callback once
   * per `sendMessage()` call.
   *
   * NOTE: By default, build stats are only logged in dev builds.
   */
  onWorkbookContextBuildStats?: (stats: WorkbookContextBuildStats) => void;
}

/**
 * React-agnostic chat orchestrator for the desktop app:
 * - Builds workbook context (schema-first + RAG) for each user message
 * - Runs tool-calling loop with preview + approval gating
 * - Writes audited runs to an `AIAuditStore`
 */
export function createAiChatOrchestrator(options: AiChatOrchestratorOptions) {
  const auditStore = options.auditStore ?? getDesktopAIAuditStore();
  const sessionId = options.sessionId ?? createSessionId(options.workbookId);
  const estimator = options.tokenEstimator ?? createHeuristicTokenEstimator();
  const contextWindowTokens = options.contextWindowTokens ?? getModeContextWindowTokens("chat", options.model);
  const reserveForOutputTokens =
    options.reserveForOutputTokens ?? getDefaultReserveForOutputTokens("chat", contextWindowTokens);
  const keepLastMessages = options.keepLastMessages ?? 40;
  const strictToolVerification = options.strictToolVerification ?? true;

  const spreadsheet = new DocumentControllerSpreadsheetApi(options.documentController, {
    createChart: options.createChart,
    sheetNameResolver: options.sheetNameResolver ?? null
  });
  const devOnBuildStats =
    import.meta.env.MODE === "development"
      ? (stats: WorkbookContextBuildStats) => {
          try {
            console.debug("[ai] WorkbookContextBuilder build stats (chat)", stats);
          } catch {
            // ignore
          }
        }
      : undefined;
  const onBuildStats =
    devOnBuildStats || options.onWorkbookContextBuildStats
      ? (stats: WorkbookContextBuildStats) => {
          devOnBuildStats?.(stats);
          options.onWorkbookContextBuildStats?.(stats);
        }
      : undefined;

  // If no context provider is passed, we create a default DesktopRagService backed by
  // persistent local storage + DocumentController listeners. In that case the
  // orchestrator owns it and must dispose it when torn down.
  let ownedRagService: DesktopRagService | null = null;
  const providedContextProvider = options.contextManager ?? options.ragService;
  const contextProvider: ContextManager | DesktopRagService =
    providedContextProvider ??
    (ownedRagService = createDesktopRagService({
      documentController: options.documentController,
      workbookId: options.workbookId,
      ...(options.ragOptions ?? {}),
      ...(options.tokenEstimator ? { tokenEstimator: estimator } : {}),
    }));

  const baseSystemPrompt =
    options.systemPrompt ??
    "You are an AI assistant inside a spreadsheet app. Prefer using tools to read data before making claims.";

  // Cache WorkbookContextBuilder across sendMessage calls so its per-sheet schema/block caches
  // actually reduce latency for multi-turn chats.
  //
  // DLP safety: only reuse the builder when the DLP inputs (policy/classifications/includeRestrictedContent)
  // are unchanged, and when tool execution options that affect read results (e.g. include_formula_values)
  // are unchanged. If the user changes policy/classifications during a session we recreate the builder so
  // data cached under a less-restrictive setting can't leak under a stricter one.
  let cachedContextBuilder: WorkbookContextBuilder | null = null;
  let cachedContextBuilderDlpKey: string | null = null;

  function createContextBuilder(): WorkbookContextBuilder {
    return new WorkbookContextBuilder({
      workbookId: options.workbookId,
      documentController: options.documentController,
      spreadsheet,
      ragService: contextProvider as any,
      schemaProvider: options.schemaProvider ?? null,
      sheetNameResolver: options.sheetNameResolver ?? null,
      includeFormulaValues: Boolean(options.toolExecutorOptions?.include_formula_values),
      mode: "chat",
      model: options.model,
      contextWindowTokens,
      reserveForOutputTokens,
      tokenEstimator: estimator as any,
      onBuildStats,
    });
  }

  let disposePromise: Promise<void> | null = null;
  async function dispose(): Promise<void> {
    if (disposePromise) return disposePromise;
    disposePromise = (async () => {
      // Clear any long-lived workbook context caches kept by the orchestrator.
      cachedContextBuilder = null;
      cachedContextBuilderDlpKey = null;
      try {
        await ownedRagService?.dispose();
      } catch {
        // Ignore disposal errors; we never want teardown to crash a UI unmount path.
      } finally {
        ownedRagService = null;
      }
    })();
    return disposePromise;
  }

  function createAbortError(message = "Aborted"): Error {
    const err = new Error(message);
    err.name = "AbortError";
    return err;
  }

  function throwIfAborted(signal: AbortSignal | undefined): void {
    if (signal?.aborted) throw createAbortError();
  }

  function isAbortError(error: unknown): boolean {
    // AbortSignal cancellations often surface as a DOMException("AbortError") in browsers,
    // but many runtimes/libraries simply throw an Error with `name === "AbortError"`.
    // We treat both as cancellations and preserve them for callers.
    if (!error) return false;
    if (typeof error !== "object" && typeof error !== "function") return false;
    return (error as any).name === "AbortError";
  }

  async function withAbort<T>(signal: AbortSignal | undefined, promise: Promise<T>): Promise<T> {
    if (!signal) return promise;
    throwIfAborted(signal);

    let rejectAbort!: (reason?: unknown) => void;
    const abortPromise = new Promise<never>((_, reject) => {
      rejectAbort = reject;
    });
    const onAbort = () => rejectAbort(createAbortError());
    signal.addEventListener("abort", onAbort, { once: true });

    // Handle the race where the signal aborts between the initial `throwIfAborted`
    // check and registering the listener (AbortSignal does not re-fire past events
    // for late listeners).
    if (signal.aborted) onAbort();

    try {
      return await Promise.race([
        promise,
        abortPromise,
      ]);
    } finally {
      signal.removeEventListener("abort", onAbort);
    }
  }

  async function sendMessage(params: SendAiChatMessageParams): Promise<SendAiChatMessageResult> {
    const signal = params.signal;
    const text = params.text.trim();
    if (!text) throw new Error("sendMessage requires non-empty text");

    const activeSheetId = options.getActiveSheetId?.() ?? "Sheet1";
    const attachments = params.attachments ?? [];
    const promptAttachments = compactAttachmentsForPrompt(attachments);
    const auditAttachments = compactAttachmentsForAudit(attachments);
    const selectedRange =
      selectionFromAttachments(attachments, activeSheetId, options.sheetNameResolver) ?? options.getSelectedRange?.() ?? undefined;

    const dlp =
      maybeGetAiCloudDlpOptions({
        documentId: options.workbookId,
        sheetId: activeSheetId,
        sheetNameResolver: options.sheetNameResolver,
      }) ?? undefined;

    let workbookContext: any;
    try {
      throwIfAborted(signal);
      const dlpKey = computeDlpCacheKey(dlp);
      const contextBuilderKey = `${dlpKey}|formula_values:${options.toolExecutorOptions?.include_formula_values ? "1" : "0"}`;
      if (!cachedContextBuilder || cachedContextBuilderDlpKey !== contextBuilderKey) {
        cachedContextBuilder = createContextBuilder();
        cachedContextBuilderDlpKey = contextBuilderKey;
      }
      workbookContext = await withAbort(
        signal,
        cachedContextBuilder.build({
          activeSheetId,
          signal,
          dlp,
          ...(selectedRange ? { selectedRange } : {}),
          focusQuestion: text,
          // Keep attachment payloads bounded before they reach downstream context builders / RAG.
          attachments: promptAttachments,
        }),
      );
    } catch (error) {
      // Hard stop: DLP says we cannot send any workbook content to a cloud model.
      // IMPORTANT: do not call the LLM in this case.
      if (error instanceof DlpViolationError) {
        // Write an audit entry even though we never reach the tool/LLM loop.
        // IMPORTANT: keep this payload sanitized (no workbook context / sampled cell values).
        let offeredTools: string[] | undefined;
        try {
          const toolPolicy =
            options.toolExecutorOptions?.toolPolicy ??
            getDesktopToolPolicy({ mode: "chat", prompt: text, hasAttachments: attachments.length > 0 });
          const toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
            ...(options.toolExecutorOptions ?? {}),
            toolPolicy,
            default_sheet: activeSheetId,
            sheet_name_resolver: options.toolExecutorOptions?.sheet_name_resolver ?? options.sheetNameResolver ?? null,
            require_approval_for_mutations: true,
            dlp,
          });
          offeredTools = toolExecutor.tools.map((t) => t.name);
        } catch {
          // ignore
        }

        const recorder = new AIAuditRecorder({
          store: auditStore,
          session_id: sessionId,
          workbook_id: options.workbookId,
          mode: "chat",
          model: options.model,
          input: {
            blocked: true,
            text: truncateTextForAudit(text, MAX_AUDIT_TEXT_CHARS),
            attachments: sanitizeAttachmentsForAudit(attachments),
            workbookId: options.workbookId,
            sheetId: activeSheetId,
            ...(offeredTools ? { offered_tools: offeredTools } : {}),
            dlp: { decision: (error as any)?.decision ?? null },
          },
        });

        try {
          await recorder.finalize();
        } catch {
          // Best-effort: do not mask the DLP error if the audit store fails.
        }

        throw new AiChatOrchestratorError(error.message, { sessionId, auditEntryId: recorder.entry.id, cause: error });
      }
      throw error;
    }

    const promptContext = formatPromptContext(workbookContext.promptContext);

    const llmMessages: LLMMessage[] = [
      {
        role: "system",
        content: `${baseSystemPrompt}\n\n${promptContext}`.trim()
      },
      ...sanitizeHistory(params.history),
      {
        role: "user",
        content: formatUserMessage(text, promptAttachments)
      }
    ];

    const toolResults: ToolExecutionResult[] = [];
    const toolPolicy =
      options.toolExecutorOptions?.toolPolicy ??
      getDesktopToolPolicy({ mode: "chat", prompt: text, hasAttachments: attachments.length > 0 });
    const toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
      ...(options.toolExecutorOptions ?? {}),
      toolPolicy,
      default_sheet: activeSheetId,
      sheet_name_resolver: options.toolExecutorOptions?.sheet_name_resolver ?? options.sheetNameResolver ?? null,
      require_approval_for_mutations: true,
      dlp
    });
    const offeredTools = toolExecutor.tools.map((t) => t.name);

    const toolTokens = estimateToolDefinitionTokens(toolExecutor.tools as any, estimator);
    const maxMessageTokens = Math.max(0, contextWindowTokens - toolTokens);

    const budgetedInitialMessages = await trimMessagesToBudget({
      messages: llmMessages as any,
      maxTokens: maxMessageTokens,
      reserveForOutputTokens,
      estimator,
      keepLastMessages,
      signal
    });

    const requireApproval = createPreviewApprovalHandler({
      spreadsheet,
      preview_options: options.previewOptions,
      executor_options: {
        default_sheet: activeSheetId,
        sheet_name_resolver: options.toolExecutorOptions?.sheet_name_resolver ?? options.sheetNameResolver ?? null,
        include_formula_values: options.toolExecutorOptions?.include_formula_values ?? false
      },
      on_approval_required: async (request: PreviewApprovalRequest) => {
        return (
          options.onApprovalRequired?.({
            call: request.call,
            preview: request.preview
          }) ?? false
        );
      }
    });

    const capturingAuditStore = new CapturingAuditStore(auditStore);
    try {
      const result = await runChatWithToolsAuditedVerified({
        client: {
          chat: async (request: any) => {
            const requestToolTokens = estimateToolDefinitionTokens(request?.tools as any, estimator);
            const requestMaxMessageTokens = Math.max(0, contextWindowTokens - requestToolTokens);
            const requestSignal: AbortSignal | undefined = request?.signal ?? signal;
            const trimmed = await trimMessagesToBudget({
              messages: request.messages as any,
              maxTokens: requestMaxMessageTokens,
              reserveForOutputTokens,
              estimator,
              keepLastMessages,
              signal: requestSignal
            });

            if (Array.isArray(request.messages)) {
              const next = trimmed === request.messages ? trimmed.slice() : trimmed;
              request.messages.length = 0;
              request.messages.push(...next);
            } else {
              request.messages = trimmed;
            }
            return options.llmClient.chat({ ...request, model: request?.model ?? options.model } as any);
          },
          streamChat: options.llmClient.streamChat
              ? async function* streamChat(request: any) {
                  const requestToolTokens = estimateToolDefinitionTokens(request?.tools as any, estimator);
                  const requestMaxMessageTokens = Math.max(0, contextWindowTokens - requestToolTokens);
                  const requestSignal: AbortSignal | undefined = request?.signal ?? signal;
                  const trimmed = await trimMessagesToBudget({
                    messages: request.messages as any,
                    maxTokens: requestMaxMessageTokens,
                    reserveForOutputTokens,
                    estimator,
                    keepLastMessages,
                    signal: requestSignal
                  });
 
                if (Array.isArray(request.messages)) {
                  const next = trimmed === request.messages ? trimmed.slice() : trimmed;
                  request.messages.length = 0;
                  request.messages.push(...next);
                } else {
                  request.messages = trimmed;
                }

                for await (const event of options.llmClient.streamChat!({
                  ...request,
                  model: request?.model ?? options.model
                } as any)) {
                  yield event;
                }
              }
            : undefined
        } as any,
        tool_executor: {
          tools: toolExecutor.tools,
          execute: async (call: any) => {
            const out = await toolExecutor.execute(call);
            toolResults.push(out as ToolExecutionResult);
            return out;
          }
        },
        messages: budgetedInitialMessages as any,
        // The verifier only needs attachment references; keep this payload bounded.
        attachments: promptAttachments,
        require_approval: requireApproval as any,
        on_tool_call: params.onToolCall as any,
        on_tool_result: params.onToolResult as any,
        on_stream_event: params.onStreamEvent as any,
        signal: params.signal,
        strict_tool_verification: strictToolVerification,
        verify_claims: true,
        verification_tool_executor: toolExecutor as any,
        audit: {
          audit_store: capturingAuditStore,
          session_id: sessionId,
          workbook_id: options.workbookId,
          mode: "chat",
          model: options.model,
          input: {
            text,
            attachments: auditAttachments,
            workbookId: options.workbookId,
            sheetId: activeSheetId,
            context: summarizeContextForAudit(workbookContext, options.sheetNameResolver ?? null),
            offered_tools: offeredTools
          }
        }
      });

      return {
        finalText: result.final,
        messages: stripLeadingSystemMessages(result.messages as LLMMessage[]),
        toolResults,
        verification: result.verification,
        context: {
          workbookId: options.workbookId,
          promptContext: workbookContext.promptContext ?? "",
          retrievedChunkIds: (workbookContext.retrieved ?? []).map((c: any) => c.id).filter(Boolean),
          retrievedRanges: extractRetrievedRanges(workbookContext.retrieved ?? [], options.sheetNameResolver ?? null),
          retrieved: workbookContext.retrieved ?? [],
          indexStats: workbookContext.indexStats,
          tokenBudgetTokens:
            contextProvider instanceof ContextManager
              ? (contextProvider as any)?.tokenBudgetTokens
              : ((await (contextProvider as DesktopRagService).getContextManager()) as any)?.tokenBudgetTokens
        },
        auditEntryId: capturingAuditStore.lastEntry?.id,
        sessionId
      };
    } catch (error) {
      if (isAbortError(error)) throw error;
      const message = error instanceof Error ? error.message : String(error);
      throw new AiChatOrchestratorError(message, {
        sessionId,
        auditEntryId: capturingAuditStore.lastEntry?.id,
        cause: error
      });
    }
  }

  return {
    sessionId,
    sendMessage,
    dispose
  };
}

class CapturingAuditStore implements AIAuditStore {
  private readonly inner: AIAuditStore;
  lastEntry: AIAuditEntry | null = null;

  constructor(inner: AIAuditStore) {
    this.inner = inner;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.lastEntry = entry;
    await this.inner.logEntry(entry);
  }

  async listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]> {
    return this.inner.listEntries(filters);
  }
}

function formatUserMessage(text: string, attachments: AiChatAttachment[]): string {
  // Attachments passed to this helper are expected to be prompt-safe already
  // (see `compactAttachmentsForPrompt` in `sendMessage`). Avoid re-compacting here
  // so compaction remains idempotent (e.g. range/table attachments are force-truncated).
  if (!attachments.length) return text;
  return `${text}\n\nAttachments:\n${formatAttachmentsForPrompt(attachments)}`;
}

function formatAttachmentsForPrompt(attachments: AiChatAttachment[]) {
  return attachments
    .map((a) => `- ${a.type}: ${a.reference}${a.data !== undefined ? ` (${stableJson(a.data)})` : ""}`)
    .join("\n");
}

function stableJson(value: unknown): string {
  try {
    return JSON.stringify(stabilizeJson(value));
  } catch {
    try {
      return JSON.stringify(String(value));
    } catch {
      return String(value);
    }
  }
}

function stabilizeJson(value: unknown, stack: WeakSet<object> = new WeakSet()): unknown {
  if (value === undefined) return null;
  if (typeof value === "bigint") return value.toString();
  if (typeof value === "symbol") return value.toString();
  if (typeof value === "function") return `[Function ${value.name || "anonymous"}]`;

  if (!value || typeof value !== "object") return value;

  const obj = value as object;
  if (stack.has(obj)) return "[Circular]";
  stack.add(obj);
  try {
    if (value instanceof Date) return value.toISOString();
    if (value instanceof Map) {
      return Array.from(value.entries()).map(([k, v]) => [stabilizeJson(k, stack), stabilizeJson(v, stack)]);
    }
    if (value instanceof Set) return Array.from(value.values()).map((v) => stabilizeJson(v, stack));

    if (Array.isArray(value)) return value.map((v) => stabilizeJson(v, stack));

    const record = value as Record<string, unknown>;
    const keys = Object.keys(record).sort();
    const out: Record<string, unknown> = {};
    for (const key of keys) out[key] = stabilizeJson(record[key], stack);
    return out;
  } finally {
    stack.delete(obj);
  }
}

function fnv1a32(value: string): number {
  // 32-bit FNV-1a hash. (Stable across runs.)
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function hashString(value: string): string {
  return fnv1a32(value).toString(16);
}

function compactAttachmentData(
  data: unknown,
  options: { maxChars: number; forceTruncate: boolean },
): unknown {
  const { json, stabilized, ok } = safeStableJson(data);
  const originalChars = json.length;
  const shouldTruncate = options.forceTruncate || !ok || originalChars > options.maxChars;
  if (!shouldTruncate) return stabilized;
  return {
    truncated: true,
    hash: hashString(json),
    original_chars: originalChars,
  };
}

function safeStableJson(value: unknown): { json: string; stabilized: unknown; ok: boolean } {
  try {
    const stabilized = stabilizeJson(value);
    const json = JSON.stringify(stabilized);
    if (typeof json === "string") return { json, stabilized, ok: true };
    // Should be rare (e.g. JSON.stringify(undefined) => undefined), but keep the return type stable.
    const fallback = JSON.stringify(String(value)) ?? String(value);
    return { json: fallback, stabilized: String(value), ok: false };
  } catch {
    try {
      const fallback = JSON.stringify(String(value));
      return { json: typeof fallback === "string" ? fallback : String(value), stabilized: String(value), ok: false };
    } catch {
      return { json: String(value), stabilized: String(value), ok: false };
    }
  }
}

function compactAttachmentsForPrompt(attachments: AiChatAttachment[]): AiChatAttachment[] {
  if (!attachments.length) return [];
  return attachments.map((a) => {
    const out: AiChatAttachment = { type: a.type, reference: a.reference } as any;
    if (a.data !== undefined) {
      // Prompts may be sent to cloud models. Never forward raw selection/table payloads
      // (even when small) because they can contain copied workbook values and would bypass
      // DLP redaction paths.
      const forceTruncate = a.type === "range" || a.type === "table";
      out.data = compactAttachmentData(a.data, { maxChars: MAX_ATTACHMENT_DATA_CHARS_FOR_PROMPT, forceTruncate });
    }
    return out;
  });
}

function compactAttachmentsForAudit(attachments: AiChatAttachment[]): Array<{ type: string; reference: string; data?: unknown }> {
  if (!attachments.length) return [];
  return attachments.map((a) => {
    const out: { type: string; reference: string; data?: unknown } = { type: a.type, reference: a.reference };
    if (a.data !== undefined) {
      // Audit logs are persisted locally (often backed by LocalStorage). Never store raw
      // selection/table payloads (they can contain user data) and keep all attachment data bounded.
      const forceTruncate = a.type === "range" || a.type === "table";
      out.data = compactAttachmentData(a.data, { maxChars: MAX_ATTACHMENT_DATA_CHARS_FOR_AUDIT, forceTruncate });
    }
    return out;
  });
}

function formatPromptContext(promptContext: string): string {
  const trimmed = String(promptContext ?? "").trim();
  if (!trimmed) return "WORKBOOK_CONTEXT:\n(none)";
  return `WORKBOOK_CONTEXT:\n${trimmed}`;
}

function summarizeContextForAudit(workbookContext: any, sheetNameResolver?: SheetNameResolver | null) {
  const retrieved = workbookContext?.retrieved ?? [];
  return {
    retrieved_chunk_ids: retrieved.map((c: any) => c.id).filter(Boolean),
    retrieved_ranges: extractRetrievedRanges(retrieved, sheetNameResolver ?? null),
    retrieved_count: retrieved.length,
    index_stats: workbookContext?.indexStats
  };
}

function createSessionId(workbookId: string): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto && typeof crypto.randomUUID === "function") {
    return `${workbookId}:${crypto.randomUUID()}`;
  }
  return `${workbookId}:${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function stripLeadingSystemMessage(messages: LLMMessage[]): LLMMessage[] {
  return stripLeadingSystemMessages(messages);
}

function stripLeadingSystemMessages(messages: LLMMessage[]): LLMMessage[] {
  const out = messages.slice();
  while (out[0]?.role === "system") out.shift();
  return out;
}

function sanitizeHistory(history: LLMMessage[] | undefined): LLMMessage[] {
  if (!history) return [];
  // The orchestrator always injects its own system prompt (including context).
  // Drop any system messages that callers may have included.
  return history.filter((m) => m.role !== "system");
}

function extractRetrievedRanges(retrieved: any[], sheetNameResolver?: SheetNameResolver | null): string[] {
  const resolver = sheetNameResolver ?? null;
  const out: string[] = [];
  for (const chunk of retrieved) {
    const meta = chunk?.metadata;
    if (!meta) continue;
    const rawSheet = typeof meta.sheetName === "string" ? meta.sheetName.trim() : "";
    const rect = meta.rect;
    if (!rawSheet || !rect) continue;
    try {
      const range = rectToA1(rect);
      const sheetName = resolver?.getSheetNameById(rawSheet) ?? rawSheet;
      out.push(`${formatSheetNameForA1(sheetName)}!${range}`);
    } catch {
      // Ignore malformed rect metadata.
    }
  }
  return out;
}

const MAX_AUDIT_TEXT_CHARS = 8_000;

function truncateTextForAudit(text: string, maxChars: number): string {
  const s = String(text ?? "");
  if (!Number.isFinite(maxChars) || maxChars <= 0) return "";
  if (s.length <= maxChars) return s;
  const marker = "[TRUNCATED]…";
  if (maxChars <= marker.length) return marker.slice(0, Math.max(0, maxChars));
  return `${s.slice(0, Math.max(0, maxChars - marker.length))}${marker}`;
}

function sanitizeAttachmentsForAudit(attachments: AiChatAttachment[]): Array<{ type: string; reference: string }> {
  // Attachments may include sampled/rich payloads. For audit entries generated in the DLP-blocked
  // path we only persist the attachment *references* (no data blobs).
  return (attachments ?? []).map((a) => ({
    type: String((a as any)?.type ?? ""),
    reference: truncateTextForAudit(String((a as any)?.reference ?? ""), 1_000),
  }));
}
