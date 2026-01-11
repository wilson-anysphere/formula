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
    // Node 22+ ships an experimental `localStorage` that throws unless started with
    // `--localstorage-file`. We only care that the storage APIs are callable.
    globalThis.localStorage?.getItem("vitest-probe");
    return true;
  } catch {
    return false;
  }
}

if (!localStorageUsable()) {
  // eslint-disable-next-line no-global-assign
  globalThis.localStorage = new MemoryLocalStorage();
}

