export type { Role, ToolCall, ToolDefinition, LLMMessage, ChatStreamEvent, LLMClient, ToolExecutor } from "./types.js";

import type { LLMClient, LLMMessage, ToolCall, ToolExecutor } from "./types.js";

export function runChatWithTools(params: {
  client: LLMClient;
  toolExecutor: ToolExecutor;
  messages: LLMMessage[];
  maxIterations?: number;
  onToolCall?: (call: ToolCall, meta: { requiresApproval: boolean }) => void;
  onToolResult?: (call: ToolCall, result: unknown) => void;
  requireApproval?: (call: ToolCall) => Promise<boolean>;
  /**
   * When true, approval denials are returned to the model as a tool result
   * (`ok:false`) and the loop continues, allowing the model to re-plan.
   * Subsequent tool calls in the same assistant message are skipped.
   *
   * Default is false (throw on denial).
   */
  continueOnApprovalDenied?: boolean;
  /**
   * When true, tool execution failures are returned to the model as tool results
   * (`ok:false`) and the loop continues, allowing the model to re-plan.
   *
   * Default is false (rethrow tool execution errors).
   */
  continueOnToolError?: boolean;
  /**
   * Max size of a serialized tool result appended to the model context (in characters).
   *
   * Default is 20_000.
   */
  maxToolResultChars?: number;
  model?: string;
  temperature?: number;
  maxTokens?: number;
  signal?: AbortSignal;
}): Promise<{ messages: LLMMessage[]; final: string }>;
