import { serializeAuditEntries, type AIAuditEntry, type AuditExportFormat } from "@formula/ai-audit/browser";

export interface AuditLogExport {
  blob: Blob;
  fileName: string;
}

export interface CreateAuditLogExportOptions {
  fileName?: string;
  /**
   * Output format.
   *
   * Defaults to NDJSON (one JSON object per line), which scales better for large logs.
   */
  format?: AuditExportFormat;
  /**
   * When enabled, remove `tool_calls[].result` (often large/sensitive) and truncate
   * oversized `audit_result_summary` payloads.
   *
   * Defaults to true for safety.
   */
  redactToolResults?: boolean;
  /**
   * Maximum number of characters allowed for a serialized `audit_result_summary`
   * when `redactToolResults` is enabled.
   */
  maxToolResultChars?: number;
}

const DEFAULT_MAX_TOOL_RESULT_CHARS = 10_000;

function defaultExportFileName(format: AuditExportFormat): string {
  const ext = format === "json" ? "json" : "ndjson";
  return `ai-audit-log-${new Date().toISOString().replaceAll(":", "-")}.${ext}`;
}

export function createAuditLogExport(entries: AIAuditEntry[], options: CreateAuditLogExportOptions = {}): AuditLogExport {
  const format = options.format ?? "ndjson";
  const redactToolResults = options.redactToolResults ?? true;
  const maxToolResultChars = options.maxToolResultChars ?? DEFAULT_MAX_TOOL_RESULT_CHARS;

  const serialized = serializeAuditEntries(entries, { format, redactToolResults, maxToolResultChars });
  const type = format === "json" ? "application/json" : "application/x-ndjson";
  const blob = new Blob([serialized], { type });
  const fileName = options.fileName ?? defaultExportFileName(format);
  return { blob, fileName };
}

export function downloadAuditLogExport(exp: AuditLogExport): void {
  // Download behavior is browser-specific. Keep the core export as a pure Blob,
  // and make the download best-effort so this can run in non-browser contexts.
  if (typeof document === "undefined") return;
  if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return;

  const url = URL.createObjectURL(exp.blob);
  try {
    const a = document.createElement("a");
    a.href = url;
    a.download = exp.fileName;
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
  } finally {
    URL.revokeObjectURL(url);
  }
}
