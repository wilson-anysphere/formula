import { afterEach } from "vitest";

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
    pointerType: string;

    constructor(type: string, init?: PointerEventInit) {
      // Many unit tests construct PointerEvents directly and assume they behave like
      // real browser pointer events (which bubble). `MouseEvent` defaults `bubbles`
      // to false unless explicitly provided, so normalize the init to better match
      // what app code expects from user-driven pointer interactions.
      const normalizedInit = { bubbles: true, cancelable: true, ...(init ?? {}) } as PointerEventInit;
      super(type, normalizedInit);
      this.pointerId = typeof normalizedInit?.pointerId === "number" ? normalizedInit.pointerId : 1;
      // Match the platform default: pointer events from a mouse have `pointerType="mouse"`.
      this.pointerType = typeof (normalizedInit as any)?.pointerType === "string" ? String((normalizedInit as any).pointerType) : "mouse";
    }
  }

  // eslint-disable-next-line no-global-assign
  (globalThis as any).PointerEvent = PointerEventShim;
}

// Several desktop features share state via globals owned by the desktop shell (e.g.
// `__formulaSpreadsheetIsEditing`, `__formulaSpreadsheetIsReadOnly`). Unit tests sometimes set these
// to emulate split-view editing / collab permissions. Clean them up after each test to avoid
// cross-test leakage (which can lead to extremely confusing, order-dependent failures).
afterEach(() => {
  delete (globalThis as any).__formulaSpreadsheetIsEditing;
  delete (globalThis as any).__formulaSpreadsheetIsReadOnly;
});
