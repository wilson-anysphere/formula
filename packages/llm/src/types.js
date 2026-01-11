/**
 * @typedef {"system"|"user"|"assistant"|"tool"} Role
 *
 * @typedef {{ id: string, name: string, arguments: any }} ToolCall
 *
 * @typedef {{
 *   name: string,
 *   description: string,
 *   parameters: any,
 *   requiresApproval?: boolean
 * }} ToolDefinition
 *
 * @typedef {{
 *   role: "system"|"user"|"assistant",
 *   content: string,
 *   toolCalls?: ToolCall[]
 * } | {
 *   role: "tool",
 *   toolCallId: string,
 *   content: string
 * }} LLMMessage
 *
 * @typedef {{
 *   messages: LLMMessage[],
 *   tools?: ToolDefinition[],
 *   toolChoice?: "auto" | "none",
 *   model?: string,
 *   temperature?: number,
 *   maxTokens?: number,
 *   signal?: AbortSignal
 * }} ChatRequest
 *
 * @typedef {{
 *   message: Extract<LLMMessage, { role: "assistant" }>,
 *   usage?: { promptTokens?: number, completionTokens?: number },
 *   raw?: any
 * }} ChatResponse
 *
 * @typedef {{
 *   type: "text",
 *   delta: string
 * } | {
 *   type: "tool_call_start",
 *   id: string,
 *   name: string
 * } | {
 *   type: "tool_call_delta",
 *   id: string,
 *   delta: string
 * } | {
 *   type: "tool_call_end",
 *   id: string
 * } | {
 *   type: "done"
 * }} ChatStreamEvent
 *
 * @typedef {{
 *   chat: (request: ChatRequest) => Promise<ChatResponse>,
 *   streamChat?: (request: ChatRequest) => AsyncIterable<ChatStreamEvent>
 * }} LLMClient
 *
 * @typedef {{
 *   tools: ToolDefinition[],
 *   execute: (call: ToolCall) => Promise<any>
 * }} ToolExecutor
 */

export {};
