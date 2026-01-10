export type AIMode = "tab_completion" | "inline_edit" | "chat" | "agent";

export interface TokenUsage {
  prompt_tokens: number;
  completion_tokens: number;
  total_tokens?: number;
}

export interface ToolCallLog {
  name: string;
  parameters: unknown;
  ok?: boolean;
  duration_ms?: number;
  result?: unknown;
  error?: string;
}

export type UserFeedback = "accepted" | "rejected" | "modified";

export interface AIAuditEntry {
  id: string;
  timestamp_ms: number;
  session_id: string;
  user_id?: string;

  mode: AIMode;
  input: unknown;

  model: string;
  token_usage?: TokenUsage;
  latency_ms?: number;

  tool_calls: ToolCallLog[];

  user_feedback?: UserFeedback;
}

export interface AuditListFilters {
  session_id?: string;
  limit?: number;
}

