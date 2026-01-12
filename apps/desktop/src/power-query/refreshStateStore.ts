import type { RefreshPolicy } from "../../../../packages/power-query/src/model.js";

export type RefreshStateEntry = { policy: RefreshPolicy; lastRunAtMs?: number };
export type RefreshState = Record<string, RefreshStateEntry>;

/**
 * Desktop persistence hooks for `RefreshManager` scheduling state.
 *
 * In the full desktop app we prefer a Tauri-backed encrypted store so schedules
 * survive app restarts. In non-Tauri environments (tests, previews) we fall back
 * to LocalStorage (stable per-workbook key) and finally an in-memory store.
 */
export type RefreshStateStore = {
  load(): Promise<RefreshState>;
  save(state: RefreshState): Promise<void>;
};

export type StorageLike = {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem?(key: string): void;
};

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvokeOrNull(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
}

class TauriPowerQueryRefreshStateStore implements RefreshStateStore {
  private readonly workbookId: string;
  private readonly invoke: TauriInvoke;

  constructor(opts: { workbookId: string; invoke: TauriInvoke }) {
    this.workbookId = opts.workbookId;
    this.invoke = opts.invoke;
  }

  async load(): Promise<RefreshState> {
    try {
      const payload = await this.invoke("power_query_refresh_state_get", { workbook_id: this.workbookId });
      if (!payload || typeof payload !== "object" || Array.isArray(payload)) return {};
      return sanitizeState(payload);
    } catch {
      return {};
    }
  }

  async save(state: RefreshState): Promise<void> {
    try {
      await this.invoke("power_query_refresh_state_set", { workbook_id: this.workbookId, state: sanitizeState(state) });
    } catch {
      // Best-effort: persistence should never break refresh.
    }
  }
}

function sanitizePolicy(policy: unknown): RefreshPolicy | null {
  if (!policy || typeof policy !== "object") return null;
  const obj = policy as any;
  const type = obj.type;
  if (type === "manual") return { type: "manual" };
  if (type === "on-open") return { type: "on-open" };
  if (type === "interval") {
    const intervalMs = obj.intervalMs;
    if (typeof intervalMs !== "number" || !Number.isFinite(intervalMs) || intervalMs <= 0) return null;
    return { type: "interval", intervalMs };
  }
  if (type === "cron") {
    const cron = obj.cron;
    if (typeof cron !== "string" || cron.trim().length === 0) return null;
    return { type: "cron", cron: cron.trim() };
  }
  return null;
}

function sanitizeEntry(entry: unknown): RefreshStateEntry | null {
  if (!entry || typeof entry !== "object") return null;
  const obj = entry as any;
  const policy = sanitizePolicy(obj.policy);
  if (!policy) return null;
  const out: RefreshStateEntry = { policy };
  const lastRunAtMs = obj.lastRunAtMs;
  if (typeof lastRunAtMs === "number" && Number.isFinite(lastRunAtMs)) out.lastRunAtMs = lastRunAtMs;
  return out;
}

function sanitizeState(input: unknown): RefreshState {
  if (!input || typeof input !== "object" || Array.isArray(input)) return {};
  const out: RefreshState = Object.create(null);
  for (const [queryId, value] of Object.entries(input as Record<string, unknown>)) {
    const entry = sanitizeEntry(value);
    if (!entry) continue;
    out[queryId] = entry;
  }
  return out;
}

function safeStorage(storage: StorageLike): StorageLike {
  return {
    getItem(key) {
      try {
        return storage.getItem(key);
      } catch {
        return null;
      }
    },
    setItem(key, value) {
      try {
        storage.setItem(key, value);
      } catch {
        // ignore
      }
    },
    removeItem(key) {
      try {
        storage.removeItem?.(key);
      } catch {
        // ignore
      }
    },
  };
}

function getLocalStorageOrNull(): StorageLike | null {
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage as any;
      if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
        return safeStorage(storage);
      }
    } catch {
      // ignore
    }
  }

  try {
    const storage = (globalThis as any)?.localStorage as any;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") {
      return safeStorage(storage);
    }
  } catch {
    // ignore
  }

  return null;
}

function storageKey(workbookId: string | undefined): string {
  return `formula.desktop.powerQuery.refreshState:${workbookId ?? "default"}`;
}

function safeParseState(raw: string | null): RefreshState {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return sanitizeState(parsed);
  } catch {
    return {};
  }
}

function cloneState(state: RefreshState): RefreshState {
  if (typeof globalThis.structuredClone === "function") return globalThis.structuredClone(state);
  return JSON.parse(JSON.stringify(state)) as RefreshState;
}

const FALLBACK_STATE_BY_KEY = new Map<string, RefreshState>();

/**
 * Create a `RefreshStateStore` suitable for the desktop app.
 *
 * Uses browser `localStorage` when available and falls back to an in-memory store in
 * non-browser environments (tests, storybook, etc).
 */
export function createPowerQueryRefreshStateStore(
  options: {
    workbookId?: string;
    storage?: StorageLike | null;
  } = {},
): RefreshStateStore {
  const workbookId = options.workbookId ?? "default";

  if (options.storage === undefined) {
    const invoke = getTauriInvokeOrNull();
    if (invoke) {
      return new TauriPowerQueryRefreshStateStore({ workbookId, invoke });
    }
  }

  const storage = options.storage === undefined ? getLocalStorageOrNull() : options.storage;
  const key = storageKey(workbookId);

  if (!storage) {
    return {
      async load() {
        return cloneState(FALLBACK_STATE_BY_KEY.get(key) ?? {});
      },
      async save(next) {
        FALLBACK_STATE_BY_KEY.set(key, cloneState(sanitizeState(next ?? {})));
      },
    };
  }

  return {
    async load() {
      return safeParseState(safeStorage(storage).getItem(key));
    },
    async save(state) {
      safeStorage(storage).setItem(key, JSON.stringify(sanitizeState(state ?? {})));
    },
  };
}

// Backwards-compatible alias (older call sites/tests).
export const createDesktopRefreshStateStore = createPowerQueryRefreshStateStore;
