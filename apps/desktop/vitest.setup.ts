class MemoryLocalStorage implements Storage {
  private readonly store = new Map<string, string>();

  get length(): number {
    return this.store.size;
  }

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.get(String(key)) ?? null;
  }

  key(index: number): string | null {
    if (index < 0) return null;
    return Array.from(this.store.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(String(key));
  }

  setItem(key: string, value: string): void {
    this.store.set(String(key), String(value));
  }
}

function storageUsable(storage: Storage | null | undefined): boolean {
  try {
    if (!storage) return false;
    // Node ships an experimental `globalThis.localStorage` that can be present but unusable unless the
    // process is started with `--localstorage-file`. Some methods may throw even if reads appear to work,
    // so probe the API surface our tests rely on.
    const probeKey = "vitest-probe";
    storage.setItem(probeKey, "1");
    storage.getItem(probeKey);
    storage.removeItem(probeKey);
    storage.clear();
    return true;
  } catch {
    return false;
  }
}

function installLocalStorage(storage: Storage): void {
  try {
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
  } catch {
    try {
      // eslint-disable-next-line no-global-assign
      (globalThis as any).localStorage = storage;
    } catch {
      // ignore
    }
  }

  if (typeof window !== "undefined") {
    try {
      Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    } catch {
      try {
        // eslint-disable-next-line no-global-assign
        (window as any).localStorage = storage;
      } catch {
        // ignore
      }
    }
  }
}

// Node 25+ ships an experimental `globalThis.localStorage` accessor that throws unless Node is started
// with `--localstorage-file`. Desktop tests rely on localStorage; provide a stable in-memory shim when
// the built-in accessor is unusable.
const existing = (() => {
  try {
    return (globalThis as any).localStorage as Storage | undefined;
  } catch {
    return undefined;
  }
})();

if (!storageUsable(existing)) {
  installLocalStorage(new MemoryLocalStorage());
}
