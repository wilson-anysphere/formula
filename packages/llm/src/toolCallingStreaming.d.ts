import type { ChatStreamEvent, LLMClient, LLMMessage, ToolCall, ToolExecutor } from "./types.js";

export function runChatWithToolsStreaming(params: {
  client: LLMClient;
  toolExecutor: ToolExecutor;
  messages: LLMMessage[];
  maxIterations?: number;
  onStreamEvent?: (event: ChatStreamEvent) => void;
  onToolCall?: (call: ToolCall, meta: { requiresApproval: boolean }) => void;
  onToolResult?: (call: ToolCall, result: unknown) => void;
  requireApproval?: (call: ToolCall) => Promise<boolean>;
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
