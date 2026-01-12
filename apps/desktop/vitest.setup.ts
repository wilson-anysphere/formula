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

function localStorageUsable(): boolean {
  try {
    // Node ships an experimental `globalThis.localStorage` that can be present but unusable
    // unless the process is started with `--localstorage-file`. Some methods may throw even if
    // simple reads appear to work, so probe the full API surface we rely on in tests.
    const storage = globalThis.localStorage;
    if (!storage) return false;

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

if (!localStorageUsable()) {
  const storage = new MemoryLocalStorage();
  try {
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
  } catch {
    // eslint-disable-next-line no-global-assign
    globalThis.localStorage = storage;
  }
}
