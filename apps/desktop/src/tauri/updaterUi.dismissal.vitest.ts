// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as notifications from "./notifications";

const DISMISSED_VERSION_KEY = "formula.updater.dismissedVersion";
const DISMISSED_AT_KEY = "formula.updater.dismissedAt";
const TEST_TIMEOUT_MS = 30_000;

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
const originalTauri = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");

beforeEach(async () => {
  document.body.innerHTML = `<div id="toast-root"></div>`;

  // Ensure any stale test stubs from other suites don't leak into these tests.
  // In particular, a mocked `__TAURI__.window.show()` that never resolves would cause
  // `handleUpdaterEvent(..., { source: "manual" })` to hang while awaiting `showMainWindowBestEffort()`.
  Object.defineProperty(globalThis, "__TAURI__", { configurable: true, value: undefined });

  // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless
  // started with `--localstorage-file`. Provide a stable in-memory implementation for tests.
  const storage = createInMemoryLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
  Object.defineProperty(window, "localStorage", { configurable: true, value: storage });

  const { __resetUpdaterUiStateForTests } = await loadUpdaterUi();
  __resetUpdaterUiStateForTests();
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

  if (originalTauri) {
    Object.defineProperty(globalThis, "__TAURI__", originalTauri);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).__TAURI__;
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
  }, TEST_TIMEOUT_MS);

  it("treats dialog cancel (Escape) as 'Later' and persists the dismissal", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });

    const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
    expect(dialog).toBeTruthy();

    dialog!.dispatchEvent(new Event("cancel", { cancelable: true }));

    expect(localStorage.getItem(DISMISSED_VERSION_KEY)).toBe("1.2.3");
    expect(Number(localStorage.getItem(DISMISSED_AT_KEY))).toBeGreaterThan(0);
  }, TEST_TIMEOUT_MS);

  it("clears stored dismissal when the user initiates an update download", async () => {
    vi.spyOn(console, "warn").mockImplementation(() => {});

    const { handleUpdaterEvent } = await loadUpdaterUi();

    // Pre-existing "Later" dismissal for the same version should remain until the user acts.
    localStorage.setItem(DISMISSED_VERSION_KEY, "1.2.3");
    localStorage.setItem(DISMISSED_AT_KEY, String(Date.now()));

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });

    expect(localStorage.getItem(DISMISSED_VERSION_KEY)).toBe("1.2.3");
    expect(localStorage.getItem(DISMISSED_AT_KEY)).toBeTruthy();

    // The updater API is intentionally not stubbed: clearing the suppression state should happen
    // immediately when the user clicks "Download" (even if the download can't start).
    const download = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
    expect(download).toBeTruthy();
    download!.click();

    expect(localStorage.getItem(DISMISSED_VERSION_KEY)).toBeNull();
    expect(localStorage.getItem(DISMISSED_AT_KEY)).toBeNull();
  }, TEST_TIMEOUT_MS);

  it("suppresses startup prompts for a recently-dismissed version, but manual checks still show", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();
    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

    localStorage.setItem(DISMISSED_VERSION_KEY, "1.2.3");
    localStorage.setItem(DISMISSED_AT_KEY, String(Date.now()));

    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Notes" });
    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeNull();
    expect(notifySpy).not.toHaveBeenCalled();

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });
    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeTruthy();
    expect(notifySpy).not.toHaveBeenCalled();
  }, TEST_TIMEOUT_MS);

  it("does not suppress when a manual check is waiting on an in-flight startup check", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();

    localStorage.setItem(DISMISSED_VERSION_KEY, "1.2.3");
    localStorage.setItem(DISMISSED_AT_KEY, String(Date.now()));

    await handleUpdaterEvent("update-check-already-running", { source: "manual" });
    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Notes" });

    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeTruthy();
  }, TEST_TIMEOUT_MS);

  it("sends a startup notification again once the dismissal TTL expires", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();
    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

    const eightDaysAgoMs = Date.now() - 8 * 24 * 60 * 60 * 1000;
    localStorage.setItem(DISMISSED_VERSION_KEY, "1.2.3");
    localStorage.setItem(DISMISSED_AT_KEY, String(eightDaysAgoMs));

    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Notes" });
    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeNull();
    expect(notifySpy).toHaveBeenCalledTimes(1);
    expect(localStorage.getItem(DISMISSED_VERSION_KEY)).toBeNull();
    expect(localStorage.getItem(DISMISSED_AT_KEY)).toBeNull();
  }, TEST_TIMEOUT_MS);

  it("clears stored dismissal when a different version becomes available at startup", async () => {
    const { handleUpdaterEvent } = await loadUpdaterUi();
    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

    localStorage.setItem(DISMISSED_VERSION_KEY, "1.2.3");
    localStorage.setItem(DISMISSED_AT_KEY, String(Date.now()));

    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.4", body: "Notes" });
    expect(document.querySelector('[data-testid="updater-dialog"]')).toBeNull();
    expect(notifySpy).toHaveBeenCalledTimes(1);
    expect(localStorage.getItem(DISMISSED_VERSION_KEY)).toBeNull();
    expect(localStorage.getItem(DISMISSED_AT_KEY)).toBeNull();
  }, TEST_TIMEOUT_MS);
});
