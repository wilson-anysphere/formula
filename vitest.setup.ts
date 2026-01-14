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
    const storage = globalThis.localStorage;
    if (!storage) return false;
    const key = "vitest-probe";
    storage.getItem(key);
    storage.setItem(key, "1");
    storage.removeItem(key);
    storage.clear();
    return true;
  } catch {
    return false;
  }
}

if (!localStorageUsable()) {
  // eslint-disable-next-line no-global-assign
  globalThis.localStorage = new MemoryLocalStorage();
}

// JSDOM does not always expose PointerEvent. Some UI tests dispatch pointer events;
// provide a minimal shim backed by MouseEvent so `new PointerEvent(...)` works.
if (typeof (globalThis as any).PointerEvent === "undefined" && typeof (globalThis as any).MouseEvent === "function") {
  const Base = (globalThis as any).MouseEvent as typeof MouseEvent;
  class PointerEventShim extends Base {
    pointerId: number;

    constructor(type: string, init?: PointerEventInit) {
      super(type, init);
      this.pointerId = typeof init?.pointerId === "number" ? init.pointerId : 1;
    }
  }

  // eslint-disable-next-line no-global-assign
  (globalThis as any).PointerEvent = PointerEventShim;
}
