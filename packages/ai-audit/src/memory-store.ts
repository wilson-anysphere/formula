import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import type { AIAuditStore } from "./store.ts";

export class MemoryAIAuditStore implements AIAuditStore {
  private readonly entries: AIAuditEntry[] = [];

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.entries.push(cloneAuditEntry(entry));
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, workbook_id, mode, limit } = filters;
    let results = session_id ? this.entries.filter((entry) => entry.session_id === session_id) : [...this.entries];
    if (workbook_id) {
      results = results.filter((entry) => matchesWorkbookFilter(entry, workbook_id));
    }
    if (mode) {
      const modes = Array.isArray(mode) ? mode : [mode];
      results = results.filter((entry) => modes.includes(entry.mode));
    }
    results.sort((a, b) => b.timestamp_ms - a.timestamp_ms);
    return typeof limit === "number" ? results.slice(0, limit) : results;
  }
}

function cloneAuditEntry(entry: AIAuditEntry): AIAuditEntry {
  const structuredCloneFn =
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    typeof (globalThis as any)?.structuredClone === "function" ? ((globalThis as any).structuredClone as any) : null;
  if (structuredCloneFn) {
    try {
      return structuredCloneFn(entry);
    } catch {
      // Fall back to JSON-based cloning.
    }
  }

  try {
    return JSON.parse(JSON.stringify(entry)) as AIAuditEntry;
  } catch {
    return entry;
  }
}

function matchesWorkbookFilter(entry: AIAuditEntry, workbookId: string): boolean {
  if (typeof entry.workbook_id === "string" && entry.workbook_id.trim()) {
    return entry.workbook_id === workbookId;
  }

  const input = entry.input as unknown;
  if (!input || typeof input !== "object") return false;
  const obj = input as Record<string, unknown>;
  const legacyWorkbookId = obj.workbook_id ?? obj.workbookId;
  return typeof legacyWorkbookId === "string" ? legacyWorkbookId === workbookId : false;
}
