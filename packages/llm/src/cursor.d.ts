import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export interface CursorLLMClientOptions {
  /**
   * Base URL of the Cursor backend. If the value ends with `/v1`, requests are
   * sent to `${baseUrl}/chat/completions`.
   *
   * If omitted, the client may fall back to runtime configuration (for example
   * `CURSOR_AI_BASE_URL`) or same-origin relative URLs.
   */
  baseUrl?: string;
  model?: string;
  timeoutMs?: number;
  /**
   * Convenience for injecting `Authorization: Bearer <token>`.
   */
  authToken?: string;
  /**
   * Called before each request to inject authentication headers.
   */
  getAuthHeaders?: () => Record<string, string> | Promise<Record<string, string>>;
}

export class CursorLLMClient {
  constructor(options?: CursorLLMClientOptions);

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}
