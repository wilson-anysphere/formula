import type { ToolCall } from "./types.js";

export function serializeToolResultForModel(params: {
  toolCall: ToolCall;
  result: unknown;
  maxChars?: number;
  maxTokens?: number;
}): string;
