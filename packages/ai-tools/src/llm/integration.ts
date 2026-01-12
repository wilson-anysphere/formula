import type { SpreadsheetApi } from "../spreadsheet/api.ts";
import type { ToolExecutorOptions, ToolExecutionResult } from "../executor/tool-executor.ts";
import { ToolExecutor } from "../executor/tool-executor.ts";
import { PreviewEngine, type PreviewEngineOptions, type ToolPlanPreview } from "../preview/preview-engine.ts";
import {
  SPREADSHEET_TOOL_DEFINITIONS,
  TOOL_CAPABILITIES,
  ToolNameSchema,
  isToolAllowedByPolicy,
  type ToolDefinition,
  type ToolName,
  type ToolPolicy
} from "../tool-schema.ts";

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
   * Least-privilege tool policy applied to both:
   * - tool exposure (`SpreadsheetLLMToolExecutor.tools`)
   * - runtime enforcement (`SpreadsheetLLMToolExecutor.execute`)
   */
  toolPolicy?: ToolPolicy;

  /**
   * Optional allowlist for tools exposed to the model.
   *
   * This is enforced at two levels:
   * - `SpreadsheetLLMToolExecutor.tools` only includes allowed tool definitions
   * - `SpreadsheetLLMToolExecutor.execute` rejects disallowed tool calls even if a model attempts them
   */
  allowed_tools?: ToolName[] | ((name: ToolName) => boolean);
}

export interface LLMToolDefinition extends ToolDefinition {
  requiresApproval?: boolean;
}

export function isSpreadsheetMutationTool(name: ToolName): boolean {
  return TOOL_CAPABILITIES[name].mutates_workbook;
}

export function getSpreadsheetToolDefinitions(
  options: { require_approval_for_mutations?: boolean; toolPolicy?: ToolPolicy } = {}
): LLMToolDefinition[] {
  const requireApprovalForMutations = options.require_approval_for_mutations ?? false;
  const policy = options.toolPolicy;

  return SPREADSHEET_TOOL_DEFINITIONS.filter((tool) => isToolAllowedByPolicy(tool.name, policy)).map((tool) => {
    const defaultRequiresApproval = TOOL_CAPABILITIES[tool.name].requires_approval_by_default ?? false;
    const requiresApproval = defaultRequiresApproval
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
  private readonly toolPolicy?: ToolPolicy;
  private readonly isAllowedTool: (name: ToolName) => boolean;
  private readonly spreadsheet: SpreadsheetApi;
  private readonly options: SpreadsheetLLMToolExecutorOptions;

  constructor(spreadsheet: SpreadsheetApi, options: SpreadsheetLLMToolExecutorOptions = {}) {
    this.spreadsheet = spreadsheet;
    this.options = options;
    this.executor = new ToolExecutor(spreadsheet, options);
    this.toolPolicy = options.toolPolicy;
    this.isAllowedTool = createAllowedToolPredicate(options.allowed_tools);

    const allTools = getSpreadsheetToolDefinitions({
      require_approval_for_mutations: options.require_approval_for_mutations,
      toolPolicy: options.toolPolicy
    });
    const supported = allTools.filter((tool) => isToolSupported(tool.name, spreadsheet, options));
    this.tools = options.allowed_tools ? supported.filter((tool) => this.isAllowedTool(tool.name)) : supported;
  }

  async execute(call: LLMToolCall): Promise<ToolExecutionResult> {
    const startedAt = nowMs();

    const nameParse = ToolNameSchema.safeParse(call.name);
    if (!nameParse.success) {
      return {
        tool: "read_range",
        ok: false,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        error: { code: "not_implemented", message: `Tool "${call.name}" is not implemented.` }
      } as ToolExecutionResult;
    }

    const name = nameParse.data;

    const supportError = toolSupportError(name, this.spreadsheet, this.options);
    if (supportError) {
      return {
        tool: name,
        ok: false,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        error: supportError
      } as ToolExecutionResult;
    }

    if (!isToolAllowedByPolicy(name, this.toolPolicy)) {
      return {
        tool: name,
        ok: false,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        error: { code: "permission_denied", message: `Tool "${name}" is not allowed by the current policy.` }
      } as ToolExecutionResult;
    }

    if (!this.isAllowedTool(name)) {
      return {
        tool: name,
        ok: false,
        timing: { started_at_ms: startedAt, duration_ms: nowMs() - startedAt },
        error: { code: "permission_denied", message: `Tool "${name}" is not allowed in this context.` }
      } as ToolExecutionResult;
    }

    return this.executor.execute({ name: call.name, parameters: call.arguments });
  }
}

function createAllowedToolPredicate(
  allowedTools: SpreadsheetLLMToolExecutorOptions["allowed_tools"]
): (name: ToolName) => boolean {
  if (!allowedTools) return () => true;
  if (Array.isArray(allowedTools)) {
    const allowSet = new Set<ToolName>(allowedTools);
    return (name) => allowSet.has(name);
  }
  return allowedTools;
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function isToolSupported(
  name: ToolName,
  spreadsheet: SpreadsheetApi,
  options: SpreadsheetLLMToolExecutorOptions
): boolean {
  return toolSupportError(name, spreadsheet, options) == null;
}

function toolSupportError(
  name: ToolName,
  spreadsheet: SpreadsheetApi,
  options: SpreadsheetLLMToolExecutorOptions
): ToolExecutionResult["error"] | null {
  if (name === "create_chart") {
    if (typeof (spreadsheet as any).createChart !== "function") {
      return { code: "not_implemented", message: "create_chart requires chart support in SpreadsheetApi" };
    }
  }

  if (name === "fetch_external_data") {
    if (!options.allow_external_data) {
      return { code: "permission_denied", message: "fetch_external_data is disabled by host configuration." };
    }
    const allowedHosts = normalizeAllowedExternalHosts(options.allowed_external_hosts);
    if (allowedHosts.length === 0) {
      return {
        code: "permission_denied",
        message: "fetch_external_data requires an explicit host allowlist (allowed_external_hosts)."
      };
    }
  }

  return null;
}

function normalizeAllowedExternalHosts(hosts: SpreadsheetLLMToolExecutorOptions["allowed_external_hosts"]): string[] {
  return (hosts ?? [])
    .map((host) => String(host).trim().toLowerCase())
    .filter((host) => host.length > 0);
}
