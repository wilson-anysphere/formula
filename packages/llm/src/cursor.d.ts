import type { ChatRequest, ChatResponse, ChatStreamEvent } from "./types.js";

export interface CursorClientOptions {
  /**
   * Base URL of the Cursor backend.
   *
   * If omitted, the client defaults to same-origin `/v1/chat/completions` and may
   * additionally be configured via `CURSOR_AI_BASE_URL` / `VITE_CURSOR_AI_BASE_URL`.
   *
   * The client is lenient in accepted values:
   * - `https://cursor.test` => `https://cursor.test/v1/chat/completions`
   * - `https://cursor.test/v1` => `https://cursor.test/v1/chat/completions`
   * - `https://cursor.test/v1/chat` => `https://cursor.test/v1/chat/completions`
   */
  baseUrl?: string;

  /**
   * Optional model hint. Cursor controls routing and may ignore this.
   */
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

/**
 * Backwards-compatible alias.
 */
export type CursorLLMClientOptions = CursorClientOptions;

export class CursorLLMClient {
  constructor(options?: CursorClientOptions);

  chat(request: ChatRequest): Promise<ChatResponse>;

  streamChat(request: ChatRequest): AsyncIterable<ChatStreamEvent>;
}

