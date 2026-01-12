import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export class CursorLLMClient {
  constructor();

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}
