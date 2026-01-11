import type { RefreshPolicy } from "../../../../packages/power-query/src/model.js";

export type RefreshStateEntry = { policy: RefreshPolicy; lastRunAtMs?: number };
export type RefreshState = Record<string, RefreshStateEntry>;

export type RefreshStateStore = {
  load(): Promise<RefreshState>;
  save(state: RefreshState): Promise<void>;
};

type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem(key: string): void };

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
        storage.removeItem(key);
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
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return {};
    return parsed as RefreshState;
  } catch {
    return {};
  }
}

function cloneState(state: RefreshState): RefreshState {
  if (typeof globalThis.structuredClone === "function") return globalThis.structuredClone(state);
  return JSON.parse(JSON.stringify(state)) as RefreshState;
}

/**
 * Create a `RefreshStateStore` suitable for the desktop app.
 *
 * Uses browser `localStorage` when available and falls back to an in-memory store in
 * non-browser environments (tests, storybook, etc).
 */
export function createPowerQueryRefreshStateStore(options: {
  workbookId?: string;
  storage?: StorageLike;
} = {}): RefreshStateStore {
  const storage = options.storage ?? getLocalStorageOrNull();
  const key = storageKey(options.workbookId);

  if (!storage) {
    let state: RefreshState = {};
    return {
      async load() {
        return cloneState(state);
      },
      async save(next) {
        state = cloneState(next ?? {});
      },
    };
  }

  return {
    async load() {
      return safeParseState(storage.getItem(key));
    },
    async save(state) {
      storage.setItem(key, JSON.stringify(state ?? {}));
    },
  };
}

