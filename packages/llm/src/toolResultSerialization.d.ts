import type { ToolCall } from "./toolCalling.js";

export function serializeToolResultForModel(params: {
  toolCall: ToolCall;
  result: unknown;
  maxChars?: number;
  maxTokens?: number;
}): string;

