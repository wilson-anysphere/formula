export type Role = "system" | "user" | "assistant" | "tool";

export interface ToolCall {
  id: string;
  name: string;
  arguments: any;
}

export interface ToolDefinition {
  name: string;
  description: string;
  parameters: any;
  requiresApproval?: boolean;
}

export type ChatStreamEvent =
  | { type: "text"; delta: string }
  | { type: "tool_call_start"; id: string; name: string }
  | { type: "tool_call_delta"; id: string; delta: string }
  | { type: "tool_call_end"; id: string }
  | { type: "done" };

export type LLMMessage =
  | {
      role: "system" | "user" | "assistant";
      content: string;
      toolCalls?: ToolCall[];
    }
  | {
      role: "tool";
      toolCallId: string;
      content: string;
    };

export interface LLMClient {
  chat: (request: any) => Promise<{ message: { role: "assistant"; content: string; toolCalls?: ToolCall[] } }>;
  streamChat?: (request: any) => AsyncIterable<ChatStreamEvent>;
}

export interface ToolExecutor {
  tools: ToolDefinition[];
  execute: (call: ToolCall) => Promise<any>;
}

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
  model?: string;
  temperature?: number;
  maxTokens?: number;
  signal?: AbortSignal;
}): Promise<{ messages: LLMMessage[]; final: string }>;

export function runChatWithToolsStreaming(params: {
  client: LLMClient;
  toolExecutor: ToolExecutor;
  messages: LLMMessage[];
  maxIterations?: number;
  onStreamEvent?: (event: ChatStreamEvent) => void;
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
  model?: string;
  temperature?: number;
  maxTokens?: number;
  signal?: AbortSignal;
}): Promise<{ messages: LLMMessage[]; final: string }>;
