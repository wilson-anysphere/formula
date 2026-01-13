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
  /**
   * Full tool result payload.
   *
   * Note: This may be omitted by higher-level integrations (e.g. `packages/ai-tools`
   * audited runs) to keep audit storage bounded, and may also be dropped by
   * `BoundedAIAuditStore` when enforcing per-entry size caps.
   */
  result?: unknown;
  /**
   * Compact summary of the tool result (bounded).
   *
   * Intended for audit stores with strict size limits (e.g. LocalStorage), and
   * as the preferred payload when `BoundedAIAuditStore` compacts oversized entries.
   */
  audit_result_summary?: unknown;
  /**
   * True when the full `result` was omitted or truncated in the audit entry.
   */
  result_truncated?: boolean;
  error?: string;
}

export type UserFeedback = "accepted" | "rejected" | "modified";

export interface AIVerificationResult {
  needs_tools: boolean;
  used_tools: boolean;
  verified: boolean;
  confidence: number;
  warnings: string[];
  /**
   * Optional claim-level verification details.
   *
   * When present, each entry represents a deterministic check comparing the
   * assistant's numeric statement (`expected`) against a spreadsheet computation
   * (`actual`) along with tool evidence used to derive it.
   */
  claims?: Array<{
    claim: string;
    verified: boolean;
    expected?: number | string | null;
    actual?: number | string | null;
    toolEvidence?: unknown;
  }>;
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
  /**
   * Host-provided input payload (often `{ prompt, attachments, ... }`).
   *
   * Consumers should treat this as sensitive data. In size-constrained audit
   * stores, this may be replaced with a truncated JSON string summary by
   * `BoundedAIAuditStore` (see `packages/ai-audit/src/bounded-store.ts`).
   */
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
  /**
   * Exclusive upper bound on `timestamp_ms`.
   */
  before_timestamp_ms?: number;
  /**
   * Inclusive lower bound on `timestamp_ms`.
   */
  after_timestamp_ms?: number;
  /**
   * Stable pagination cursor.
   *
   * Results are ordered newest-first. When provided, the store should return
   * entries strictly older than the cursor (older timestamp, or for equal
   * timestamps, an `id` tiebreaker when `before_id` is present).
   */
  cursor?: { before_timestamp_ms: number; before_id?: string };
  limit?: number;
}
