// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const DISMISSED_VERSION_KEY = "formula.updater.dismissedVersion";
const DISMISSED_AT_KEY = "formula.updater.dismissedAt";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

async function loadUpdaterUi() {
  return await import("./updaterUi");
}

const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

beforeEach(() => {
  vi.resetModules();
  document.body.innerHTML = `<div id="toast-root"></div>`;

  // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless
  // started with `--localstorage-file`. Provide a stable in-memory implementation for tests.
  const storage = createInMemoryLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
  Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
});

afterEach(() => {
  if (originalGlobalLocalStorage) {
    Object.defineProperty(globalThis, "localStorage", originalGlobalLocalStorage);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).localStorage;
  }

  if (originalWindowLocalStorage) {
    Object.defineProperty(window, "localStorage", originalWindowLocalStorage);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (window as any).localStorage;
  }
});

describe("tauri/updaterUi dismissal persistence", () => {
  it("persists the user's 'Later' dismissal decision", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });

    const later = document.querySelector<HTMLButtonElement>('[data-testid="updater-later"]');
    expect(later).toBeTruthy();
    later!.click();

    expect(localStorage.getItem(DISMISSED_VERSION_KEY)).toBe("1.2.3");
    expect(Number(localStorage.getItem(DISMISSED_AT_KEY))).toBeGreaterThan(0);
  });

  it("suppresses startup prompts for a recently-dismissed version, but manual checks still show", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();

    localStorage.setItem(DISMISSED_VERSION_KEY, "1.2.3");
    localStorage.setItem(DISMISSED_AT_KEY, String(Date.now()));

    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Notes" });
    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeNull();

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });
    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeTruthy();
  });
});

