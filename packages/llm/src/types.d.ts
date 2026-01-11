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

export interface SystemMessage {
  role: "system";
  content: string;
}

export interface UserMessage {
  role: "user";
  content: string;
}

export interface AssistantMessage {
  role: "assistant";
  content: string;
  toolCalls?: ToolCall[];
}

export interface ToolMessage {
  role: "tool";
  toolCallId: string;
  content: string;
}

export type LLMMessage = SystemMessage | UserMessage | AssistantMessage | ToolMessage;

export interface ChatRequest {
  messages: LLMMessage[];
  tools?: ToolDefinition[];
  toolChoice?: "auto" | "none";
  model?: string;
  temperature?: number;
  maxTokens?: number;
  signal?: AbortSignal;
}

export interface ChatUsage {
  promptTokens?: number;
  completionTokens?: number;
  totalTokens?: number;
}

export interface ChatResponse {
  message: AssistantMessage;
  usage?: ChatUsage;
  raw?: any;
}

export type ChatStreamEvent =
  | { type: "text"; delta: string }
  | { type: "tool_call_start"; id: string; name: string }
  | { type: "tool_call_delta"; id: string; delta: string }
  | { type: "tool_call_end"; id: string }
  | { type: "done"; usage?: ChatUsage };

export interface LLMClient {
  chat: (request: ChatRequest) => Promise<ChatResponse>;
  streamChat?: (request: ChatRequest) => AsyncIterable<ChatStreamEvent>;
}

export interface ToolExecutor {
  tools: ToolDefinition[];
  execute: (call: ToolCall) => Promise<any>;
}
