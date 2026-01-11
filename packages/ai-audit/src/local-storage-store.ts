import type { AIAuditStore } from "./store.js";
import type { AIAuditEntry, AuditListFilters } from "./types.js";

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
    entries.push(structuredClone(entry));
    entries.sort((a, b) => a.timestamp_ms - b.timestamp_ms);
    while (entries.length > this.maxEntries) entries.shift();
    this.saveEntries(entries);
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, limit } = filters;
    const entries = this.loadEntries();
    const filtered = session_id ? entries.filter((entry) => entry.session_id === session_id) : entries.slice();
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
    if (!storage) {
      this.memoryFallback.length = 0;
      this.memoryFallback.push(...entries);
      return;
    }

    try {
      storage.setItem(this.key, JSON.stringify(entries));
    } catch {
      // If persistence fails, at least keep the latest entries in memory.
      this.localStorageUnavailable = true;
      this.memoryFallback.length = 0;
      this.memoryFallback.push(...entries);
    }
  }

  private getLocalStorage(): Storage | null {
    if (this.localStorageUnavailable) return null;
    return getLocalStorageOrNull();
  }
}

function getLocalStorageOrNull(): Storage | null {
  try {
    if (typeof globalThis === "undefined") return null;
    const storage = globalThis.localStorage;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}
