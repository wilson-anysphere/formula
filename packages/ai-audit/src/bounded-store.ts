import type { AIAuditStore } from "./store.ts";
import type { AIAuditEntry, AuditListFilters, ToolCallLog } from "./types.ts";
import { stableStringify } from "./stable-json.ts";

export interface BoundedAIAuditStoreOptions {
  /**
   * Maximum allowed size (in characters) of a single serialized audit entry.
   *
   * This is a defense-in-depth cap to keep LocalStorage/IndexedDB-backed stores
   * from failing writes due to quota overruns.
   *
   * Defaults to 200k characters.
   */
  max_entry_chars?: number;
}

/**
 * AIAuditStore wrapper that enforces an upper bound on the serialized size of
 * each entry.
 *
 * If an entry exceeds `max_entry_chars`, it is compacted by:
 * - Replacing `input` with a truncated JSON string summary (with metadata).
 * - Truncating tool call `parameters` and `audit_result_summary` similarly.
 * - Dropping full tool call `result` payloads.
 *
 * This wrapper is intended as defense-in-depth in addition to upstream audit
 * log compaction (e.g. `packages/ai-tools` already bounds tool parameters and
 * result summaries).
 */
export class BoundedAIAuditStore implements AIAuditStore {
  private readonly store: AIAuditStore;
  readonly maxEntryChars: number;

  constructor(store: AIAuditStore, options: BoundedAIAuditStoreOptions = {}) {
    this.store = store;
    const max = options.max_entry_chars ?? 200_000;
    // Ensure we have a finite positive integer; fall back to default if invalid.
    this.maxEntryChars = Number.isFinite(max) && max > 0 ? Math.floor(max) : 200_000;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    const maxChars = this.maxEntryChars;
    let serialized: string | null = null;
    try {
      serialized = JSON.stringify(entry);
    } catch {
      // If the entry isn't JSON-serializable for some reason, force compaction to
      // avoid surprising store failures.
    }
    if (!serialized) {
      try {
        serialized = JSON.stringify(entry, bigIntReplacer);
      } catch {
        // Still not serializable (e.g. circular references). Fall back to compaction.
      }
    }

    if (serialized && serialized.length <= maxChars) {
      await this.store.logEntry(entry);
      return;
    }

    const compacted = compactAuditEntry(entry, maxChars);
    await this.store.logEntry(compacted);
  }

  async listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]> {
    return this.store.listEntries(filters);
  }
}

function bigIntReplacer(_key: string, value: unknown): unknown {
  return typeof value === "bigint" ? value.toString() : value;
}

type AuditJsonSummary = {
  audit_truncated: true;
  audit_original_chars: number;
  audit_json: string;
};

function compactAuditEntry(entry: AIAuditEntry, maxChars: number): AIAuditEntry {
  const toolCalls = Array.isArray(entry.tool_calls) ? entry.tool_calls : [];
  const workbookId =
    normalizeNonEmptyString(entry.workbook_id) ??
    extractWorkbookIdFromInput(entry.input) ??
    extractWorkbookIdFromSessionId(entry.session_id);

  // Fast path: even if the initial JSON.stringify threw, it's still possible that
  // the entry fits when serialized. Try once defensively.
  try {
    if (JSON.stringify(entry, bigIntReplacer).length <= maxChars) return entry;
  } catch {
    // continue with compaction
  }

  // Try progressively more aggressive budgets until we fit under `maxChars`.
  let toolCallLimit = toolCalls.length;
  let valueBudget = computeInitialValueBudget(maxChars, toolCallLimit);

  // Toggle optional fields as a last resort.
  let includeOptionalFields = true;

  for (let attempt = 0; attempt < 12; attempt++) {
    const candidate = buildCompactedEntry(entry, {
      workbookId,
      toolCalls,
      toolCallLimit,
      valueBudget,
      includeOptionalFields
    });

    let serialized: string | null = null;
    try {
      serialized = JSON.stringify(candidate, bigIntReplacer);
    } catch {
      // Non-serializable candidate (e.g. circular optional fields). We'll retry
      // with optional fields removed.
    }
    if (serialized && serialized.length <= maxChars) return candidate;
    if (!serialized) {
      if (includeOptionalFields) {
        includeOptionalFields = false;
        valueBudget = Math.min(valueBudget, 64);
        continue;
      }
      break;
    }

    if (valueBudget > 64) {
      valueBudget = Math.max(32, Math.floor(valueBudget * 0.5));
      continue;
    }

    if (toolCallLimit > 1) {
      toolCallLimit = Math.max(1, Math.floor(toolCallLimit * 0.5));
      valueBudget = computeInitialValueBudget(maxChars, toolCallLimit);
      continue;
    }

    if (includeOptionalFields) {
      includeOptionalFields = false;
      valueBudget = Math.min(valueBudget, 64);
      continue;
    }

    break;
  }

  // Absolute fallback: store only the fields required for filtering and a minimal
  // truncation marker.
  return minimalStub(entry, workbookId);
}

function buildCompactedEntry(
  entry: AIAuditEntry,
  options: {
    workbookId: string | undefined;
    toolCalls: ToolCallLog[];
    toolCallLimit: number;
    valueBudget: number;
    includeOptionalFields: boolean;
  }
): AIAuditEntry {
  const { toolCalls, toolCallLimit, valueBudget, includeOptionalFields } = options;
  const includedCalls = toolCalls.slice(0, toolCallLimit);
  const droppedCalls = toolCalls.length - includedCalls.length;

  const compactedToolCalls: ToolCallLog[] = includedCalls.map((call) => compactToolCall(call, valueBudget));

  if (droppedCalls > 0) {
    compactedToolCalls.push({
      name: "audit_truncated_tool_calls",
      parameters: {
        audit_truncated: true,
        audit_note: `Dropped ${droppedCalls} tool calls to fit audit entry size limit.`
      }
    });
  }

  const base: AIAuditEntry = {
    id: entry.id,
    timestamp_ms: entry.timestamp_ms,
    session_id: entry.session_id,
    workbook_id: options.workbookId,
    user_id: entry.user_id,
    mode: entry.mode,
    input: summarizeAsJson(entry.input, valueBudget),
    model: entry.model,
    tool_calls: compactedToolCalls
  };

  if (!includeOptionalFields) return base;

  // These fields are typically small, but we keep them only while they fit in the budget.
  if (entry.token_usage !== undefined) base.token_usage = entry.token_usage;
  if (entry.latency_ms !== undefined) base.latency_ms = entry.latency_ms;
  if (entry.verification !== undefined) base.verification = entry.verification;
  if (entry.user_feedback !== undefined) base.user_feedback = entry.user_feedback;

  return base;
}

function compactToolCall(call: ToolCallLog, valueBudget: number): ToolCallLog {
  const compacted: ToolCallLog = {
    name: call.name,
    parameters: summarizeAsJson(call.parameters, valueBudget)
  };

  if (call.requires_approval !== undefined) compacted.requires_approval = call.requires_approval;
  if (call.approved !== undefined) compacted.approved = call.approved;
  if (call.ok !== undefined) compacted.ok = call.ok;
  if (call.duration_ms !== undefined) compacted.duration_ms = call.duration_ms;
  if (call.error !== undefined) compacted.error = call.error;

  if (call.audit_result_summary !== undefined) {
    compacted.audit_result_summary = summarizeAsJson(call.audit_result_summary, valueBudget);
  }

  // Drop the full tool result payload. If it was present, mark it as truncated.
  if (call.result_truncated !== undefined) {
    compacted.result_truncated = call.result_truncated || call.result !== undefined;
  } else if (call.result !== undefined) {
    compacted.result_truncated = true;
  }

  return compacted;
}

function summarizeAsJson(value: unknown, maxJsonChars: number): AuditJsonSummary {
  const budget = Number.isFinite(maxJsonChars) && maxJsonChars > 0 ? Math.floor(maxJsonChars) : 1_000;
  const json = safeJsonStringify(value);
  return {
    audit_truncated: true,
    audit_original_chars: json.length,
    audit_json: truncateString(json, budget)
  };
}

function truncateString(value: string, maxLength: number): string {
  const limit = Number.isFinite(maxLength) && maxLength > 0 ? Math.floor(maxLength) : 0;
  if (limit === 0) return "";
  if (value.length <= limit) return value;

  // Keep the marker short so we don't blow the budget.
  const marker = "…";
  if (limit <= marker.length) return marker.slice(0, limit);
  return `${value.slice(0, limit - marker.length)}${marker}`;
}

function safeJsonStringify(value: unknown): string {
  try {
    return stableStringify(value);
  } catch {
    try {
      const json = JSON.stringify(String(value));
      return typeof json === "string" ? json : String(value);
    } catch {
      return String(value);
    }
  }
}

function computeInitialValueBudget(maxChars: number, toolCallCount: number): number {
  const limit = Number.isFinite(maxChars) && maxChars > 0 ? Math.floor(maxChars) : 200_000;
  const heavyFields = 1 + toolCallCount * 2;
  // Reserve a chunk for fixed fields / JSON overhead so we don't start with an
  // unrealistically large per-field budget.
  const reserved = Math.min(10_000, Math.floor(limit * 0.25));
  const remaining = Math.max(128, limit - reserved);
  return Math.max(128, Math.floor(remaining / Math.max(1, heavyFields)));
}

function minimalStub(entry: AIAuditEntry, workbookId: string | undefined): AIAuditEntry {
  return {
    id: entry.id,
    timestamp_ms: entry.timestamp_ms,
    session_id: entry.session_id,
    workbook_id: workbookId,
    user_id: entry.user_id,
    mode: entry.mode,
    model: entry.model,
    input: { audit_truncated: true },
    tool_calls: []
  };
}

function extractWorkbookIdFromInput(input: unknown): string | undefined {
  if (!input || typeof input !== "object") return undefined;
  const obj = input as Record<string, unknown>;
  const workbookId = obj.workbook_id ?? obj.workbookId;
  const trimmed = typeof workbookId === "string" ? workbookId.trim() : "";
  return trimmed ? trimmed : undefined;
}

function extractWorkbookIdFromSessionId(sessionId: string): string | undefined {
  const match = sessionId.match(/^([^:]+):([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$/);
  if (!match) return undefined;
  const workbookId = match[1];
  const trimmed = workbookId?.trim() ?? "";
  return trimmed ? trimmed : undefined;
}

function normalizeNonEmptyString(value: unknown): string | undefined {
  const trimmed = typeof value === "string" ? value.trim() : "";
  return trimmed ? trimmed : undefined;
}
