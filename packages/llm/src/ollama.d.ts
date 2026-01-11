import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export interface OllamaChatClientOptions {
  baseUrl?: string;
  model?: string;
  timeoutMs?: number;
}

export class OllamaChatClient {
  constructor(options?: OllamaChatClientOptions);

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}

