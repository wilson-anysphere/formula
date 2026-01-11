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
  model?: string;
  temperature?: number;
  maxTokens?: number;
  signal?: AbortSignal;
}): Promise<{ messages: LLMMessage[]; final: string }>;
