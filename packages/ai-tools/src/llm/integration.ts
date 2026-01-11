import type { SpreadsheetApi } from "../spreadsheet/api.js";
import type { ToolExecutorOptions, ToolExecutionResult } from "../executor/tool-executor.js";
import { ToolExecutor } from "../executor/tool-executor.js";
import { PreviewEngine, type PreviewEngineOptions, type ToolPlanPreview } from "../preview/preview-engine.js";
import { SPREADSHEET_TOOL_DEFINITIONS, type ToolDefinition, type ToolName } from "../tool-schema.js";
import { enforceToolOutputDlp, type ToolOutputDlpOptions } from "../dlp/toolOutputDlp.js";

export interface LLMToolCall {
  id?: string;
  name: string;
  arguments: unknown;
}

export interface PreviewApprovalRequest {
  call: LLMToolCall;
  preview: ToolPlanPreview;
}

export interface PreviewApprovalHandlerOptions {
  spreadsheet: SpreadsheetApi;
  preview_engine?: PreviewEngine;
  preview_options?: PreviewEngineOptions;
  executor_options?: ToolExecutorOptions;
  on_approval_required?: (request: PreviewApprovalRequest) => Promise<boolean>;
}

/**
 * Create a `requireApproval` callback compatible with `packages/llm`'s `runChatWithTools`.
 *
 * The callback uses `PreviewEngine` to simulate the tool call on a clone of the spreadsheet and:
 * - auto-approves safe/no-op changes
 * - delegates to `on_approval_required` when the preview engine flags risk
 *
 * If approval is required but no `on_approval_required` handler is provided, the callback rejects
 * by returning `false` (safe default).
 */
export function createPreviewApprovalHandler(options: PreviewApprovalHandlerOptions): (call: LLMToolCall) => Promise<boolean> {
  const previewEngine = options.preview_engine ?? new PreviewEngine(options.preview_options);
  const executorOptions = options.executor_options ?? {};

  return async (call: LLMToolCall) => {
    const preview = await previewEngine.generatePreview(
      [{ name: call.name, parameters: call.arguments } as any],
      options.spreadsheet,
      executorOptions
    );

    if (!preview.requires_approval) return true;
    if (!options.on_approval_required) return false;
    return options.on_approval_required({ call, preview });
  };
}

export interface SpreadsheetLLMToolExecutorOptions extends ToolExecutorOptions {
  /**
   * When enabled, tools that mutate the workbook are marked `requiresApproval: true`
   * so higher-level orchestration (e.g. `runChatWithTools`) can gate them.
   *
   * Note: `fetch_external_data` is always marked as requiring approval because it performs
   * external network access.
   */
  require_approval_for_mutations?: boolean;
  /**
   * Optional DLP enforcement for tool results intended for cloud LLM processing.
   *
   * IMPORTANT: When provided, the *returned* tool result (not the raw workbook data)
   * is what will be persisted in the tool message history. This prevents sensitive
   * workbook contents from being sent to cloud providers via tool outputs.
   */
  dlp?: ToolOutputDlpOptions;
}

export interface LLMToolDefinition extends ToolDefinition {
  requiresApproval?: boolean;
}

export function isSpreadsheetMutationTool(name: ToolName): boolean {
  switch (name) {
    case "read_range":
    case "filter_range":
    case "detect_anomalies":
    case "compute_statistics":
      return false;
    default:
      return true;
  }
}

export function getSpreadsheetToolDefinitions(options: { require_approval_for_mutations?: boolean } = {}): LLMToolDefinition[] {
  const requireApprovalForMutations = options.require_approval_for_mutations ?? false;
  return SPREADSHEET_TOOL_DEFINITIONS.map((tool) => {
    const requiresApproval =
      tool.name === "fetch_external_data"
        ? true
        : requireApprovalForMutations
          ? isSpreadsheetMutationTool(tool.name)
          : undefined;
    return {
      ...tool,
      ...(requiresApproval !== undefined ? { requiresApproval } : {})
    };
  });
}

/**
 * Adapter to use `packages/ai-tools`' ToolExecutor with `packages/llm`'s tool-calling loop.
 */
export class SpreadsheetLLMToolExecutor {
  readonly tools: LLMToolDefinition[];
  private readonly executor: ToolExecutor;
  private readonly dlp?: ToolOutputDlpOptions;

  constructor(spreadsheet: SpreadsheetApi, options: SpreadsheetLLMToolExecutorOptions = {}) {
    this.executor = new ToolExecutor(spreadsheet, options);
    this.tools = getSpreadsheetToolDefinitions({ require_approval_for_mutations: options.require_approval_for_mutations });
    this.dlp = options.dlp;
  }

  async execute(call: LLMToolCall): Promise<ToolExecutionResult> {
    const result = await this.executor.execute({ name: call.name, parameters: call.arguments });
    if (!this.dlp) return result;
    return enforceToolOutputDlp({
      call,
      result,
      dlp: this.dlp,
      defaultSheet: this.executor.options.default_sheet
    });
  }
}

export type { ToolOutputDlpOptions };
