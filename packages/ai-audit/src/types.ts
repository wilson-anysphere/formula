export type AIMode = "tab_completion" | "inline_edit" | "chat" | "agent" | "cell_function";

export interface TokenUsage {
  prompt_tokens: number;
  completion_tokens: number;
  total_tokens?: number;
}

export interface ToolCallLog {
  name: string;
  parameters: unknown;
  requires_approval?: boolean;
  approved?: boolean;
  ok?: boolean;
  duration_ms?: number;
  result?: unknown;
  error?: string;
}

export type UserFeedback = "accepted" | "rejected" | "modified";

export interface AIVerificationResult {
  needs_tools: boolean;
  used_tools: boolean;
  verified: boolean;
  confidence: number;
  warnings: string[];
}

export interface AIAuditEntry {
  id: string;
  timestamp_ms: number;
  session_id: string;
  /**
   * Optional workbook identifier used for server-side filtering in audit stores.
   *
   * This is intentionally separate from `session_id` because sessions may be
   * regenerated within a workbook (or shared across workbooks in some hosts).
   */
  workbook_id?: string;
  user_id?: string;

  mode: AIMode;
  input: unknown;

  model: string;
  token_usage?: TokenUsage;
  latency_ms?: number;

  tool_calls: ToolCallLog[];

  verification?: AIVerificationResult;

  user_feedback?: UserFeedback;
}

export interface AuditListFilters {
  session_id?: string;
  workbook_id?: string;
  mode?: AIMode | AIMode[];
  limit?: number;
}
