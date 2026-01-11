import type { LLMClient, LLMMessage } from "../../../../../packages/llm/src/types.js";

import type { AIAuditStore } from "../../../../../packages/ai-audit/src/store.js";
import type { AIAuditEntry, AuditListFilters } from "../../../../../packages/ai-audit/src/types.js";

import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import type { TokenEstimator } from "../../../../../packages/ai-context/src/tokenBudget.js";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens } from "../../../../../packages/ai-context/src/tokenBudget.js";
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

import { DLP_ACTION } from "../../../../../packages/security/dlp/src/actions.js";
import { DlpViolationError } from "../../../../../packages/security/dlp/src/errors.js";

import type { DocumentController } from "../../document/documentController.js";

import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";
import { createDesktopRagService, type DesktopRagService, type DesktopRagServiceOptions } from "../rag/ragService.js";
import { getDesktopAIAuditStore } from "../audit/auditStore.js";
import { maybeGetAiCloudDlpOptions } from "../dlp/aiDlp.js";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../contextBudget.js";

export type AiChatAttachment =
  | { type: "range"; reference: string; data?: unknown }
  | { type: "formula"; reference: string; data?: { formula: string } }
  | { type: "table"; reference: string; data?: unknown }
  | { type: "chart"; reference: string; data?: unknown };

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

  getActiveSheetId?: () => string;
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
   * by a persistent sqlite vector store (stored in LocalStorage).
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

  const spreadsheet = new DocumentControllerSpreadsheetApi(options.documentController, { createChart: options.createChart });

  const contextProvider:
    | ContextManager
    | DesktopRagService = options.contextManager ??
    options.ragService ??
    createDesktopRagService({
      documentController: options.documentController,
      workbookId: options.workbookId,
      ...(options.ragOptions ?? {}),
    });

  const baseSystemPrompt =
    options.systemPrompt ??
    "You are an AI assistant inside a spreadsheet app. Prefer using tools to read data before making claims.";

  async function sendMessage(params: SendAiChatMessageParams): Promise<SendAiChatMessageResult> {
    const text = params.text.trim();
    if (!text) throw new Error("sendMessage requires non-empty text");

    const activeSheetId = options.getActiveSheetId?.() ?? "Sheet1";
    const attachments = params.attachments ?? [];

    const dlp = maybeGetAiCloudDlpOptions({ documentId: options.workbookId, sheetId: activeSheetId }) ?? undefined;
    // DLP context building triggers a full workbook scan for redaction before indexing.
    // Preserve the desktop RAG service's incremental indexing fast path when there are
    // no classifications and the policy doesn't outright forbid cloud processing.
    const aiRule = (dlp as any)?.policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING];
    const shouldApplyDlpToContext = dlp ? dlp.classificationRecords.length > 0 || aiRule?.maxAllowed == null : false;
    const dlpForContext = shouldApplyDlpToContext ? dlp : undefined;

    let workbookContext: any;
    try {
      workbookContext = await contextProvider.buildWorkbookContextFromSpreadsheetApi({
        spreadsheet,
        workbookId: options.workbookId,
        query: text,
        attachments,
        dlp: dlpForContext
      });
    } catch (error) {
      // Hard stop: DLP says we cannot send any workbook content to a cloud model.
      // IMPORTANT: do not call the LLM in this case.
      if (error instanceof DlpViolationError) {
        throw new AiChatOrchestratorError(error.message, { sessionId, cause: error });
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
        content: formatUserMessage(text, attachments)
      }
    ];

    const toolResults: ToolExecutionResult[] = [];
    const toolExecutor = new SpreadsheetLLMToolExecutor(spreadsheet, {
      ...(options.toolExecutorOptions ?? {}),
      default_sheet: activeSheetId,
      require_approval_for_mutations: true,
      dlp
    });

    const toolTokens = estimateToolDefinitionTokens(toolExecutor.tools as any, estimator);
    const maxMessageTokens = Math.max(0, contextWindowTokens - toolTokens);

    const budgetedInitialMessages = await trimMessagesToBudget({
      messages: llmMessages as any,
      maxTokens: maxMessageTokens,
      reserveForOutputTokens,
      estimator,
      keepLastMessages
    });

    const requireApproval = createPreviewApprovalHandler({
      spreadsheet,
      preview_options: options.previewOptions,
      executor_options: { default_sheet: activeSheetId },
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
            const trimmed = await trimMessagesToBudget({
              messages: request.messages as any,
              maxTokens: requestMaxMessageTokens,
              reserveForOutputTokens,
              estimator,
              keepLastMessages
            });

            if (Array.isArray(request.messages)) {
              const next = trimmed === request.messages ? trimmed.slice() : trimmed;
              request.messages.length = 0;
              request.messages.push(...next);
            } else {
              request.messages = trimmed;
            }
            return options.llmClient.chat({ ...request, model: request?.model ?? options.model } as any);
          }
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
        attachments,
        require_approval: requireApproval as any,
        on_tool_call: params.onToolCall as any,
        on_tool_result: params.onToolResult as any,
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
            attachments,
            workbookId: options.workbookId,
            sheetId: activeSheetId,
            context: summarizeContextForAudit(workbookContext)
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
          retrievedRanges: extractRetrievedRanges(workbookContext.retrieved ?? []),
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
    sendMessage
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
  if (!attachments.length) return text;
  return `${text}\n\nAttachments:\n${formatAttachmentsForPrompt(attachments)}`;
}

function formatAttachmentsForPrompt(attachments: AiChatAttachment[]) {
  return attachments
    .map((a) => `- ${a.type}: ${a.reference}${a.data ? ` (${stableJson(a.data)})` : ""}`)
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

function stabilizeJson(value: unknown): unknown {
  if (value === undefined) return null;
  if (typeof value === "bigint") return value.toString();
  if (typeof value === "symbol") return value.toString();
  if (typeof value === "function") return `[Function ${value.name || "anonymous"}]`;
  if (value instanceof Date) return value.toISOString();

  if (Array.isArray(value)) return value.map((v) => stabilizeJson(v));

  if (value && typeof value === "object") {
    const obj = value as Record<string, unknown>;
    const keys = Object.keys(obj).sort();
    const out: Record<string, unknown> = {};
    for (const key of keys) out[key] = stabilizeJson(obj[key]);
    return out;
  }

  return value;
}

function formatPromptContext(promptContext: string): string {
  const trimmed = String(promptContext ?? "").trim();
  if (!trimmed) return "WORKBOOK_CONTEXT:\n(none)";
  return `WORKBOOK_CONTEXT:\n${trimmed}`;
}

function summarizeContextForAudit(workbookContext: any) {
  const retrieved = workbookContext?.retrieved ?? [];
  return {
    retrieved_chunk_ids: retrieved.map((c: any) => c.id).filter(Boolean),
    retrieved_ranges: extractRetrievedRanges(retrieved),
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

function extractRetrievedRanges(retrieved: any[]): string[] {
  const out: string[] = [];
  for (const chunk of retrieved) {
    const meta = chunk?.metadata;
    if (!meta) continue;
    const sheetName = typeof meta.sheetName === "string" ? meta.sheetName : null;
    const rect = meta.rect;
    if (!sheetName || !rect) continue;
    try {
      const range = rectToA1(rect);
      out.push(`${formatSheetNameForA1(sheetName)}!${range}`);
    } catch {
      // Ignore malformed rect metadata.
    }
  }
  return out;
}

function formatSheetNameForA1(sheetName: string): string {
  // Quote when needed (Excel style): 'Sheet Name'!A1
  if (/^[A-Za-z0-9_]+$/.test(sheetName)) return sheetName;
  return `'${sheetName.replace(/'/g, "''")}'`;
}
