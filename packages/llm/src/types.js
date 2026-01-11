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
 *   role: "system",
 *   content: string
 * } | {
 *   role: "user",
 *   content: string
 * } | {
 *   role: "assistant",
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
 *   usage?: { promptTokens?: number, completionTokens?: number, totalTokens?: number },
 *   raw?: any
 * }} ChatResponse
 *
 * @typedef {{
 *   // Streamed text delta. Consumers should append `delta` to reconstruct the
 *   // assistant text for the current model call.
 *   type: "text",
 *   delta: string
 * } | {
 *   // Tool call start. Some providers only provide a final tool call id late
 *   // in the stream; clients may synthesize ids (e.g. `toolcall-0`) so tool
 *   // results can still be attached as `role: "tool"` messages.
 *   type: "tool_call_start",
 *   id: string,
 *   name: string
 * } | {
 *   // Incremental tool call arguments. Providers are inconsistent: some stream
 *   // true deltas, others repeat the full argument string. Client
 *   // implementations should normalize to emitting only incremental suffixes.
 *   type: "tool_call_delta",
 *   id: string,
 *   delta: string
 * } | {
 *   type: "tool_call_end",
 *   id: string
 * } | {
 *   type: "done",
 *   usage?: { promptTokens?: number, completionTokens?: number, totalTokens?: number }
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
