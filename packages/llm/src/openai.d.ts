import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export interface OpenAIClientOptions {
  apiKey?: string;
  model?: string;
  baseUrl?: string;
  timeoutMs?: number;
}

export class OpenAIClient {
  constructor(options?: OpenAIClientOptions);

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}

