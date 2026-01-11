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
  message: Extract<LLMMessage, { role: "assistant" }>;
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

