import { DocumentController } from "../../document/documentController.js";
import { rangeToA1 } from "../../selection/a1";
import type { Range } from "../../selection/types";

import { ToolExecutor, PreviewEngine, runChatWithToolsAudited } from "../../../../../packages/ai-tools/src/index.js";
import { SpreadsheetLLMToolExecutor } from "../../../../../packages/ai-tools/src/llm/integration.js";

import { OpenAIClient } from "../../../../../packages/llm/src/openai.js";

import type { AIAuditStore } from "../../../../../packages/ai-audit/src/store.js";

import { DLP_ACTION } from "../../../../../packages/security/dlp/src/actions.js";
import { formatDlpDecisionMessage } from "../../../../../packages/security/dlp/src/errors.js";
import { DLP_DECISION, evaluatePolicy } from "../../../../../packages/security/dlp/src/policyEngine.js";
import { effectiveRangeClassification } from "../../../../../packages/security/dlp/src/selectors.js";

import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";
import { getDesktopAIAuditStore } from "../audit/auditStore.js";
import { maybeGetAiCloudDlpOptions } from "../dlp/aiDlp.js";
import { InlineEditOverlay } from "./inlineEditOverlay";
import type { TokenEstimator } from "../../../../../packages/ai-context/src/tokenBudget.js";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens } from "../../../../../packages/ai-context/src/tokenBudget.js";
import { trimMessagesToBudget } from "../../../../../packages/ai-context/src/trimMessagesToBudget.js";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../contextBudget.js";

const OPENAI_API_KEY_STORAGE_KEY = "formula:openaiApiKey";

export interface InlineEditLLMClient {
  chat: (request: any) => Promise<any>;
}

export interface InlineEditControllerOptions {
  container: HTMLElement;
  document: DocumentController;
  /**
   * Identifier for the active workbook. Used to correlate audit entries with the
   * audit log viewer (which defaults to filtering by workbook id).
   */
  workbookId?: string;
  getSheetId: () => string;
  getSelectionRange: () => Range | null;
  onApplied?: () => void;
  onClosed?: () => void;

  llmClient?: InlineEditLLMClient;
  model?: string;
  auditStore?: AIAuditStore;

  /**
   * Optional override for the strict inline-edit prompt budget. The effective
   * budget is always capped to the inline-edit maximum (see `getModeContextWindowTokens`).
   */
  contextWindowTokens?: number;
  reserveForOutputTokens?: number;
  keepLastMessages?: number;
  tokenEstimator?: TokenEstimator;
}

export class InlineEditController {
  private readonly overlay: InlineEditOverlay;
  private readonly previewEngine = new PreviewEngine();
  private isRunning = false;
  private abortController: AbortController | null = null;

  constructor(private readonly options: InlineEditControllerOptions) {
    this.overlay = new InlineEditOverlay(options.container);
  }

  isOpen(): boolean {
    return this.overlay.isOpen();
  }

  open(): void {
    if (this.isRunning) return;
    if (this.overlay.isOpen()) return;

    const range = this.options.getSelectionRange();
    if (!range) return;

    const sheetId = this.options.getSheetId();
    const selectionLabel = `${sheetId}!${rangeToA1(range)}`;

    this.overlay.open(selectionLabel, {
      onCancel: () => {
        this.cancel();
      },
      onRun: (prompt) => {
        void this.runInlineEdit({ sheetId, range, prompt });
      }
    });
  }

  close(): void {
    this.cancel();
  }

  private async runInlineEdit(params: { sheetId: string; range: Range; prompt: string }): Promise<void> {
    if (this.isRunning) return;
    const client = this.options.llmClient ?? createDefaultInlineEditClient({ model: this.options.model });
    const model = this.options.model ?? (client as any)?.model ?? "gpt-4o-mini";
    if (!client) {
      this.overlay.showError(
        "AI client is not configured. Open the AI panel to set an OpenAI API key (stored in localStorage)."
      );
      return;
    }

    this.isRunning = true;
    const abortController = new AbortController();
    this.abortController = abortController;
    const signal = abortController.signal;
    let batchStarted = false;
    try {
      const workbookId = this.options.workbookId ?? "local-workbook";
      const dlp = maybeGetAiCloudDlpOptions({ documentId: workbookId, sheetId: params.sheetId }) ?? undefined;

      // If the selection itself is blocked for cloud processing, stop before reading any
      // sample data or calling the LLM.
      const selectionRangeRef = {
        documentId: workbookId,
        sheetId: params.sheetId,
        range: {
          start: { row: params.range.startRow, col: params.range.startCol },
          end: { row: params.range.endRow, col: params.range.endCol }
        }
      };
      if (dlp) {
        const selectionClassification = effectiveRangeClassification(selectionRangeRef as any, dlp.classificationRecords);
        const selectionDecision = evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification: selectionClassification,
          policy: dlp.policy,
          options: { includeRestrictedContent: false }
        });
        if (selectionDecision.decision === DLP_DECISION.BLOCK) {
          dlp.auditLogger?.log({
            type: "ai.inline_edit",
            documentId: workbookId,
            sheetId: params.sheetId,
            range: selectionRangeRef.range,
            action: DLP_ACTION.AI_CLOUD_PROCESSING,
            decision: selectionDecision,
            selectionClassification,
            redactedCellCount: 0
          });
          this.overlay.showError(formatDlpDecisionMessage(selectionDecision));
          return;
        }
      }
      const estimator = this.options.tokenEstimator ?? createHeuristicTokenEstimator();
      const strictContextWindowTokens = getModeContextWindowTokens("inline_edit", model);
      const contextWindowTokens = Math.min(this.options.contextWindowTokens ?? strictContextWindowTokens, strictContextWindowTokens);
      const reserveForOutputTokens =
        this.options.reserveForOutputTokens ?? getDefaultReserveForOutputTokens("inline_edit", contextWindowTokens);
      const keepLastMessages = this.options.keepLastMessages ?? 20;

      const baseApi = new DocumentControllerSpreadsheetApi(this.options.document);
      const api = createAbortableSpreadsheetApi(baseApi, signal);
      const executor = new ToolExecutor(api, { default_sheet: params.sheetId, dlp });

      const selectionRef = `${params.sheetId}!${rangeToA1(params.range)}`;
      const sampleRef = buildSampleRange(params.sheetId, params.range, { maxRows: 10, maxCols: 10 });

      this.overlay.setRunning("Reading selection…");
      throwIfAborted(signal);
      const sampleResult = await executor.execute({
        name: "read_range",
        parameters: { range: sampleRef, include_formulas: false }
      });
      throwIfAborted(signal);

      const sampleValues =
        sampleResult.ok && sampleResult.data && "values" in sampleResult.data ? (sampleResult.data as any).values : null;

      const messages = buildMessages({
        sheetId: params.sheetId,
        selection: selectionRef,
        sampleRange: sampleRef,
        sampleValues,
        prompt: params.prompt
      });

      const toolExecutor = new SpreadsheetLLMToolExecutor(api, {
        default_sheet: params.sheetId,
        require_approval_for_mutations: true,
        dlp
      });
      const abortableToolExecutor = {
        tools: toolExecutor.tools,
        execute: async (call: any) => {
          throwIfAborted(signal);
          const result = await toolExecutor.execute(call);
          throwIfAborted(signal);
          return result;
        }
      };

      const auditStore = this.options.auditStore ?? getDesktopAIAuditStore();
      const sessionId = createSessionId();

      try {
        this.overlay.setRunning("Running AI tools…");
        await runChatWithToolsAudited({
          client: {
            chat: async (request: any) => {
              throwIfAborted(signal);
              const toolTokens = estimateToolDefinitionTokens(request?.tools as any, estimator);
              const maxMessageTokens = Math.max(0, contextWindowTokens - toolTokens);
              const trimmed = await trimMessagesToBudget({
                messages: request.messages as any,
                maxTokens: maxMessageTokens,
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
              const response = await client.chat({ ...request, signal });
              throwIfAborted(signal);
              return response;
            }
          },
          tool_executor: abortableToolExecutor as any,
          messages,
           audit: {
              audit_store: auditStore,
              session_id: sessionId,
              workbook_id: workbookId,
              mode: "inline_edit",
              input: { prompt: params.prompt, selection: selectionRef, workbookId, sheetId: params.sheetId },
              model
            },
          require_approval: async (call) => {
            this.overlay.setRunning("Generating preview…");
            throwIfAborted(signal);
            const preview = await this.previewEngine.generatePreview(
              [{ name: call.name, parameters: call.arguments } as any],
              api,
              { default_sheet: params.sheetId }
            );
            throwIfAborted(signal);
            const approved = await this.overlay.requestApproval(preview);
            if (!approved) return false;

            if (!batchStarted) {
              this.options.document.beginBatch({ label: "AI Inline Edit" });
              batchStarted = true;
            }
            this.overlay.setRunning("Applying changes…");
            return true;
          }
        });

        if (batchStarted) {
          this.options.document.endBatch();
        }
        this.closeOverlayIfOpen();
        this.options.onApplied?.();
      } catch (error) {
        if (batchStarted) {
          this.options.document.cancelBatch();
        }

        if (isAbortError(error)) {
          this.closeOverlayIfOpen();
          return;
        }

        if (error instanceof Error && error.message.includes("was denied")) {
          this.closeOverlayIfOpen();
          return;
        }

        if (this.overlay.isOpen()) {
          this.overlay.showError(error instanceof Error ? error.message : String(error));
        }
      }
    } finally {
      this.isRunning = false;
      if (this.abortController === abortController) {
        this.abortController = null;
      }
    }
  }

  private closeOverlayIfOpen(): void {
    if (!this.overlay.isOpen()) return;
    this.overlay.close();
    this.options.onClosed?.();
  }

  private cancel(): void {
    if (this.isRunning) {
      try {
        this.abortController?.abort();
      } catch {
        // ignore
      }
      // Roll back any in-flight batch so subsequent user edits don't get swallowed.
      try {
        this.options.document.cancelBatch?.();
      } catch {
        // ignore
      }
    }
    this.closeOverlayIfOpen();
  }
}

function createSessionId(): string {
  if (typeof globalThis.crypto !== "undefined" && typeof (globalThis.crypto as any).randomUUID === "function") {
    return (globalThis.crypto as any).randomUUID();
  }
  return `inline-edit-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createDefaultInlineEditClient(opts: { model?: string } = {}): InlineEditLLMClient | null {
  const apiKey = loadOpenAIApiKeyFromRuntime();
  if (!apiKey) return null;
  try {
    return new OpenAIClient({ apiKey, model: opts.model });
  } catch {
    return null;
  }
}

function loadOpenAIApiKeyFromRuntime(): string | null {
  try {
    const stored = globalThis.localStorage?.getItem(OPENAI_API_KEY_STORAGE_KEY);
    if (stored) return stored;
  } catch {
    // ignore (storage may be disabled)
  }

  // Allow Vite devs to inject a key without touching localStorage.
  const envKey = (import.meta as any)?.env?.VITE_OPENAI_API_KEY;
  if (typeof envKey === "string" && envKey.length > 0) return envKey;

  return null;
}

function createAbortError(): Error {
  const error = new Error("Inline edit was cancelled.");
  error.name = "AbortError";
  return error;
}

function throwIfAborted(signal: AbortSignal): void {
  if (!signal.aborted) return;
  throw createAbortError();
}

function isAbortError(error: unknown): boolean {
  if (!error) return false;
  if (error instanceof DOMException) return error.name === "AbortError";
  if (error instanceof Error) return error.name === "AbortError";
  return false;
}

function createAbortableSpreadsheetApi(api: any, signal: AbortSignal): any {
  const guard = () => throwIfAborted(signal);
  return {
    listSheets: (...args: any[]) => api.listSheets(...args),
    listNonEmptyCells: (...args: any[]) => api.listNonEmptyCells(...args),
    getCell: (...args: any[]) => api.getCell(...args),
    setCell: (...args: any[]) => {
      guard();
      return api.setCell(...args);
    },
    readRange: (...args: any[]) => api.readRange(...args),
    writeRange: (...args: any[]) => {
      guard();
      return api.writeRange(...args);
    },
    applyFormatting: (...args: any[]) => {
      guard();
      return api.applyFormatting(...args);
    },
    getLastUsedRow: (...args: any[]) => api.getLastUsedRow(...args),
    clone: () => createAbortableSpreadsheetApi(api.clone(), signal),
    ...(typeof api.createChart === "function"
      ? {
          createChart: (...args: any[]) => {
            guard();
            return api.createChart(...args);
          }
        }
      : {})
  };
}

function buildMessages(options: {
  sheetId: string;
  selection: string;
  sampleRange: string;
  sampleValues: unknown;
  prompt: string;
}): Array<{ role: "system" | "user"; content: string }> {
  const system = [
    "You are an AI assistant embedded in a spreadsheet.",
    "The user has selected a range and wants to transform it.",
    "Use the provided spreadsheet tools to make the requested edit.",
    "",
    "Rules:",
    "- Prefer a single set_range call when possible.",
    "- Only modify cells within the selection unless the user explicitly asks otherwise.",
    "- Do not call apply_formatting unless explicitly asked.",
    "- If writing formulas, include the leading '='."
  ].join("\n");

  const user = [
    `Sheet: ${options.sheetId}`,
    `Selection: ${options.selection}`,
    `Selection sample (${options.sampleRange}):`,
    options.sampleValues != null ? JSON.stringify(options.sampleValues, null, 2) : "(unavailable)",
    "",
    `User request: ${options.prompt}`
  ].join("\n");

  return [
    { role: "system", content: system },
    { role: "user", content: user }
  ];
}

function buildSampleRange(sheetId: string, range: Range, limits: { maxRows: number; maxCols: number }): string {
  const rows = Math.max(1, range.endRow - range.startRow + 1);
  const cols = Math.max(1, range.endCol - range.startCol + 1);
  const endRow = range.startRow + Math.min(rows, limits.maxRows) - 1;
  const endCol = range.startCol + Math.min(cols, limits.maxCols) - 1;
  return `${sheetId}!${rangeToA1({ startRow: range.startRow, endRow, startCol: range.startCol, endCol })}`;
}
