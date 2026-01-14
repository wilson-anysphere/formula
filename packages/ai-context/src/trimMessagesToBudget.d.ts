import type { TokenEstimator } from "./tokenBudget.js";

export const CONTEXT_SUMMARY_MARKER: string;

export function trimMessagesToBudget(params: {
  messages: unknown[];
  maxTokens: number;
  reserveForOutputTokens?: number;
  estimator?: TokenEstimator;
  keepLastMessages?: number;
  summaryMaxTokens?: number;
  summarize?:
    | null
    | ((messagesToSummarize: unknown[]) => string | unknown | null | undefined | Promise<string | unknown | null | undefined>);
  summaryRole?: "system" | "assistant";
  /**
   * When true (default), preserve tool-call coherence:
   * - Never keep a `role: "tool"` message without also keeping the assistant message that issued the tool call.
   * - Never keep an assistant message with `toolCalls` without also keeping the subsequent `role: "tool"` messages (when present).
   */
  preserveToolCallPairs?: boolean;
  /**
   * When true, prefer dropping completed tool-call groups (assistant(toolCalls)+tool*) before dropping other messages.
   */
  dropToolMessagesFirst?: boolean;
  signal?: AbortSignal;
}): Promise<unknown[]>;
