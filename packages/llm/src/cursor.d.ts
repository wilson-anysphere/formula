import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export interface CursorClientOptions {
  /**
   * Optional Bearer token forwarded as `Authorization: Bearer <token>`.
   *
   * Prefer `getAuthHeaders()` when available so Cursor can manage authentication.
   */
  authToken?: string;

  /**
   * Preferred authentication hook.
   *
   * Return additional headers required by the Cursor backend. This can be used
   * to forward session cookies, Cursor-issued tokens, etc.
   */
  getAuthHeaders?: () => Promise<Record<string, string>> | Record<string, string>;

  /**
   * Optional model hint. Cursor controls routing and may ignore this.
   */
  model?: string;

  /**
   * Base URL for the Cursor backend (default: `https://api.cursor.sh/v1`).
   */
  baseUrl?: string;

  timeoutMs?: number;
}

export class CursorLLMClient {
  constructor(options?: CursorClientOptions);

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}

