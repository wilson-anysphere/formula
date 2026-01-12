import type { AIAuditStore } from "./store.ts";
import type { AIAuditEntry, AuditListFilters } from "./types.ts";

export interface LocalStorageAIAuditStoreOptions {
  /**
   * localStorage key used to persist entries. Defaults to `formula_ai_audit_log_entries`.
   */
  key?: string;
  /**
   * Cap the number of stored entries (oldest dropped). Defaults to 1000.
   */
  max_entries?: number;
}

export class LocalStorageAIAuditStore implements AIAuditStore {
  readonly key: string;
  readonly maxEntries: number;
  private readonly memoryFallback: AIAuditEntry[] = [];
  private localStorageUnavailable: boolean = false;

  constructor(options: LocalStorageAIAuditStoreOptions = {}) {
    this.key = options.key ?? "formula_ai_audit_log_entries";
    this.maxEntries = options.max_entries ?? 1000;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    const entries = this.loadEntries();
    entries.push(cloneAuditEntry(entry));
    entries.sort((a, b) => a.timestamp_ms - b.timestamp_ms);
    while (entries.length > this.maxEntries) entries.shift();
    this.saveEntries(entries);
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, workbook_id, mode, limit } = filters;
    const entries = this.loadEntries();
    let filtered = session_id ? entries.filter((entry) => entry.session_id === session_id) : entries.slice();
    if (workbook_id) {
      filtered = filtered.filter((entry) => matchesWorkbookFilter(entry, workbook_id));
    }
    if (mode) {
      const modes = Array.isArray(mode) ? mode : [mode];
      filtered = filtered.filter((entry) => modes.includes(entry.mode));
    }
    filtered.sort((a, b) => b.timestamp_ms - a.timestamp_ms);
    return typeof limit === "number" ? filtered.slice(0, limit) : filtered;
  }

  private loadEntries(): AIAuditEntry[] {
    const storage = this.getLocalStorage();
    if (!storage) return this.memoryFallback.slice();
    try {
      const raw = storage.getItem(this.key);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? (parsed as AIAuditEntry[]) : [];
    } catch {
      // localStorage can throw in some environments (e.g. Node webstorage without a file,
      // Safari private mode). Fall back to in-memory storage.
      this.localStorageUnavailable = true;
      return this.memoryFallback.slice();
    }
  }

  private saveEntries(entries: AIAuditEntry[]): void {
    const storage = this.getLocalStorage();
    const snapshot = entries === this.memoryFallback ? entries.slice() : entries;
    if (!storage) {
      this.memoryFallback.length = 0;
      this.memoryFallback.push(...snapshot);
      return;
    }
    try {
      storage.setItem(this.key, JSON.stringify(entries));
    } catch {
      // If persistence fails, at least keep the latest entries in memory.
      this.localStorageUnavailable = true;
      this.memoryFallback.length = 0;
      this.memoryFallback.push(...snapshot);
    }
  }

  private getLocalStorage(): Storage | null {
    if (this.localStorageUnavailable) return null;
    const storage = getLocalStorageOrNull();
    // If localStorage exists but is inaccessible (e.g. Node's experimental webstorage
    // without a file path), avoid retrying (and throwing) on every call.
    if (!storage && typeof globalThis !== "undefined" && "localStorage" in globalThis) {
      this.localStorageUnavailable = true;
    }
    return storage;
  }
}

function getLocalStorageOrNull(): Storage | null {
  // Prefer `window.localStorage` when available (standard browser case).
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
      return storage;
    } catch {
      // ignore
    }
  }

  try {
    if (typeof globalThis === "undefined") return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
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
    // Last resort: return as-is (audit entries should be JSON-serializable in practice).
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
