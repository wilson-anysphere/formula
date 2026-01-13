import type { AIAuditEntry, ToolCallLog } from "./types.ts";

export type AuditExportFormat = "ndjson" | "json";

export interface SerializeAuditEntriesOptions {
  /**
   * Output format.
   *
   * - `ndjson`: One JSON object per line.
   * - `json`: A single JSON array.
   */
  format?: AuditExportFormat;
  /**
   * When true, remove `tool_calls[].result` (often large / sensitive) and optionally
   * truncate oversized `audit_result_summary` payloads.
   */
  redactToolResults?: boolean;
  /**
   * Maximum number of characters allowed for a serialized `audit_result_summary`
   * when `redactToolResults` is enabled.
   */
  maxToolResultChars?: number;
}

const DEFAULT_FORMAT: AuditExportFormat = "ndjson";
const DEFAULT_MAX_TOOL_RESULT_CHARS = 10_000;
const DEFAULT_REDACT_TOOL_RESULTS = true;

/**
 * Serialize audit entries deterministically for export / troubleshooting.
 *
 * Notes:
 * - Deterministic output is achieved by recursively sorting object keys before
 *   stringification.
 * - When `redactToolResults` is enabled, full tool `result` payloads are removed
 *   to avoid exporting large or sensitive data.
 */
export function serializeAuditEntries(entries: AIAuditEntry[], opts: SerializeAuditEntriesOptions = {}): string {
  const format = opts.format ?? DEFAULT_FORMAT;
  const redactToolResults = opts.redactToolResults ?? DEFAULT_REDACT_TOOL_RESULTS;
  const maxToolResultChars = opts.maxToolResultChars ?? DEFAULT_MAX_TOOL_RESULT_CHARS;

  const sanitizedEntries = entries.map((entry) =>
    redactToolResults ? redactEntryToolResults(entry, maxToolResultChars) : entry
  );

  if (format === "ndjson") {
    return sanitizedEntries.map((entry) => stableStringify(entry)).join("\n");
  }

  return stableStringify(sanitizedEntries);
}

function redactEntryToolResults(entry: AIAuditEntry, maxToolResultChars: number): AIAuditEntry {
  return {
    ...entry,
    tool_calls: (entry.tool_calls ?? []).map((call) => redactToolCall(call, maxToolResultChars))
  };
}

function redactToolCall(call: ToolCallLog, maxToolResultChars: number): ToolCallLog & { export_truncated?: true } {
  const { result: _result, ...rest } = call;
  const output: ToolCallLog & { export_truncated?: true } = { ...rest };

  if (output.audit_result_summary !== undefined) {
    const summaryStr =
      typeof output.audit_result_summary === "string" ? output.audit_result_summary : stableStringify(output.audit_result_summary);

    if (summaryStr.length > maxToolResultChars) {
      output.audit_result_summary = truncateTo(summaryStr, maxToolResultChars);
      output.export_truncated = true;
    }
  }

  return output;
}

function truncateTo(input: string, maxChars: number): string {
  if (input.length <= maxChars) return input;
  if (maxChars <= 0) return "";
  // Keep within the requested limit while still being human-readable.
  if (maxChars === 1) return "…";
  return `${input.slice(0, maxChars - 1)}…`;
}

/**
 * A tiny, dependency-free stable JSON stringify.
 *
 * This is equivalent to JSON.stringify for supported inputs, but guarantees
 * stable object key ordering across runs.
 */
function stableStringify(value: unknown): string {
  // JSON.stringify can return `undefined` for unsupported top-level inputs
  // (e.g. `undefined`). Since our public API returns `string`, normalize those
  // cases to `"null"` for deterministic output.
  return JSON.stringify(stableJsonValue(value, new WeakSet())) ?? "null";
}

function stableJsonValue(value: unknown, ancestors: WeakSet<object>): unknown {
  if (value === null) return null;

  const t = typeof value;
  if (t === "string" || t === "number" || t === "boolean") return value;
  if (t === "bigint") return value.toString();
  if (t === "undefined" || t === "function" || t === "symbol") return undefined;

  if (Array.isArray(value)) {
    if (ancestors.has(value)) return "[Circular]";
    ancestors.add(value);
    const out = value.map((item) => stableJsonValue(item, ancestors));
    ancestors.delete(value);
    return out;
  }

  if (t !== "object") return undefined;

  const obj = value as Record<string, unknown>;

  // Preserve JSON.stringify behavior for objects with toJSON (e.g. Date).
  if (typeof (obj as { toJSON?: unknown }).toJSON === "function") {
    return stableJsonValue((obj as { toJSON: () => unknown }).toJSON(), ancestors);
  }

  if (ancestors.has(obj)) return "[Circular]";
  ancestors.add(obj);

  const sorted: Record<string, unknown> = {};
  for (const key of Object.keys(obj).sort()) {
    sorted[key] = stableJsonValue(obj[key], ancestors);
  }
  ancestors.delete(obj);
  return sorted;
}
