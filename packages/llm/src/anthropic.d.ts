import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export interface AnthropicClientOptions {
  apiKey?: string;
  model?: string;
  baseUrl?: string;
  timeoutMs?: number;
  maxTokens?: number;
}

export class AnthropicClient {
  constructor(options?: AnthropicClientOptions);

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}

