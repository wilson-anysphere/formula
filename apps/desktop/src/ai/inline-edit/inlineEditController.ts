import { DocumentController } from "../../document/documentController.js";
import { rangeToA1 } from "../../selection/a1";
import type { Range } from "../../selection/types";

import { ToolExecutor, PreviewEngine, runChatWithToolsAudited } from "../../../../../packages/ai-tools/src/index.js";
import { SpreadsheetLLMToolExecutor } from "../../../../../packages/ai-tools/src/llm/integration.js";

import { OpenAIClient } from "../../../../../packages/llm/src/openai.js";

import type { AIAuditStore } from "../../../../../packages/ai-audit/src/store.js";
import { LocalStorageAIAuditStore } from "../../../../../packages/ai-audit/src/local-storage-store.js";

import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";
import { InlineEditOverlay } from "./inlineEditOverlay";

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
}

export class InlineEditController {
  private readonly overlay: InlineEditOverlay;
  private readonly previewEngine = new PreviewEngine();
  private isRunning = false;

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
        this.overlay.close();
        this.options.onClosed?.();
      },
      onRun: (prompt) => {
        void this.runInlineEdit({ sheetId, range, prompt });
      }
    });
  }

  close(): void {
    this.overlay.close();
    this.options.onClosed?.();
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
    try {
      const api = new DocumentControllerSpreadsheetApi(this.options.document);
      const executor = new ToolExecutor(api, { default_sheet: params.sheetId });

      const selectionRef = `${params.sheetId}!${rangeToA1(params.range)}`;
      const sampleRef = buildSampleRange(params.sheetId, params.range, { maxRows: 10, maxCols: 10 });

      this.overlay.setRunning("Reading selection…");
      const sampleResult = await executor.execute({
        name: "read_range",
        parameters: { range: sampleRef, include_formulas: false }
      });

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
        require_approval_for_mutations: true
      });

      const auditStore = this.options.auditStore ?? new LocalStorageAIAuditStore();
      const sessionId = createSessionId();
      const workbookId = this.options.workbookId ?? "local-workbook";

      this.options.document.beginBatch({ label: "AI Inline Edit" });

      try {
        this.overlay.setRunning("Running AI tools…");
        await runChatWithToolsAudited({
          client,
          tool_executor: toolExecutor as any,
          messages,
           audit: {
             audit_store: auditStore,
             session_id: sessionId,
             mode: "inline_edit",
             input: { prompt: params.prompt, selection: selectionRef, workbookId, sheetId: params.sheetId },
             model
           },
          require_approval: async (call) => {
            this.overlay.setRunning("Generating preview…");
            const preview = await this.previewEngine.generatePreview(
              [{ name: call.name, parameters: call.arguments } as any],
              api,
              { default_sheet: params.sheetId }
            );
            const approved = await this.overlay.requestApproval(preview);
            if (!approved) return false;
            this.overlay.setRunning("Applying changes…");
            return true;
          }
        });

        this.options.document.endBatch();
        this.overlay.close();
        this.options.onApplied?.();
        this.options.onClosed?.();
      } catch (error) {
        this.options.document.cancelBatch();
        if (error instanceof Error && error.message.includes("was denied")) {
          this.overlay.close();
          this.options.onClosed?.();
          return;
        }
        this.overlay.showError(error instanceof Error ? error.message : String(error));
      }
    } finally {
      this.isRunning = false;
    }
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
