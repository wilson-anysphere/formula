/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => {
  return {
    shellOpen: vi.fn<[string], Promise<void>>().mockResolvedValue(undefined),
  };
});

vi.mock("../shellOpen", () => ({
  shellOpen: mocks.shellOpen,
}));

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

async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("updaterUi (dialog + download)", () => {
  const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

  beforeEach(() => {
    document.body.innerHTML = '<div id="toast-root"></div>';

    // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless
    // started with `--localstorage-file`. Provide a stable in-memory implementation for tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });

    vi.resetModules();
  });

  afterEach(() => {
    mocks.shellOpen.mockClear();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    document.body.replaceChildren();

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

  it("does not open a dialog when update-available is received during startup checks", async () => {
    const handlers = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      handlers.set(eventName, handler);
      return () => handlers.delete(eventName);
    });

    vi.stubGlobal("__TAURI__", { event: { listen } });

    const { installUpdaterUi } = await import("../updaterUi");
    await installUpdaterUi();

    handlers.get("update-available")?.({ payload: { source: "startup", version: "9.9.9", body: "Release notes\nLine 2" } });
    await flushMicrotasks();

    const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
    expect(dialog).toBeNull();
  }, 30_000);

  it("does not suppress an update dialog when the user clicks manual check during an in-flight startup check", async () => {
    // Simulate "Later" suppression for a specific version.
    window.localStorage.setItem("formula.updater.dismissedVersion", "9.9.9");
    window.localStorage.setItem("formula.updater.dismissedAt", String(Date.now()));

    const { handleUpdaterEvent } = await import("../updaterUi");

    // User clicks "Check for Updates" while a startup check is already running.
    await handleUpdaterEvent("update-check-already-running", { source: "manual" });

    // Startup-sourced update result arrives, but should still open a dialog because the user
    // explicitly requested a manual check.
    await handleUpdaterEvent("update-available", { source: "startup", version: "9.9.9", body: "Release notes\nLine 2" });
    await flushMicrotasks();

    const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
    expect(dialog).not.toBeNull();
    expect(dialog?.getAttribute("open") === "" || dialog?.open === true).toBe(true);
  });

  it("shows download progress and reveals the restart button once downloaded", async () => {
    const handlers = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      handlers.set(eventName, handler);
      return () => handlers.delete(eventName);
    });

    const download = vi.fn(async (onProgress?: any) => {
      onProgress?.({ downloaded: 50, total: 100 });
      await flushMicrotasks(1);
      onProgress?.({ downloaded: 100, total: 100 });
    });

    const install = vi.fn(async () => {});
    const check = vi.fn(async () => ({
      version: "1.2.3",
      body: "notes",
      download,
      install,
    }));

    vi.stubGlobal("__TAURI__", {
      event: { listen },
      updater: { check },
    });

    const { installUpdaterUi } = await import("../updaterUi");
    await installUpdaterUi();

    handlers.get("update-available")?.({ payload: { source: "manual", version: "1.2.3", body: "notes" } });
    await flushMicrotasks();

    const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
    expect(downloadBtn).not.toBeNull();

    downloadBtn?.click();
    await flushMicrotasks(12);

    expect(download).toHaveBeenCalledTimes(1);
    expect(check).toHaveBeenCalledTimes(1);

    const progress = document.querySelector<HTMLProgressElement>('[data-testid="updater-progress"]');
    const progressText = document.querySelector<HTMLElement>('[data-testid="updater-progress-text"]');
    expect(progress).not.toBeNull();
    expect(progress?.value).toBe(100);
    expect(progressText?.textContent).toBe("100%");

    const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
    expect(restartBtn).not.toBeNull();
    expect(restartBtn?.style.display).not.toBe("none");
  });

  it("promotes 'Download manually' when the update download fails", async () => {
    vi.spyOn(console, "error").mockImplementation(() => {});

    const handlers = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      handlers.set(eventName, handler);
      return () => handlers.delete(eventName);
    });

    const download = vi.fn(async () => {
      throw new Error("network down");
    });

    const install = vi.fn(async () => {});
    const check = vi.fn(async () => ({
      version: "1.2.3",
      body: "notes",
      download,
      install,
    }));

    vi.stubGlobal("__TAURI__", {
      event: { listen },
      updater: { check },
    });

    const { installUpdaterUi } = await import("../updaterUi");
    await installUpdaterUi();

    handlers.get("update-available")?.({ payload: { source: "manual", version: "1.2.3", body: "notes" } });
    await flushMicrotasks();

    const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
    expect(downloadBtn).not.toBeNull();

    downloadBtn?.click();
    await flushMicrotasks(12);

    const dialog = document.querySelector('[data-testid="updater-dialog"]') as HTMLElement | null;
    expect(dialog).not.toBeNull();

    const viewBtn = dialog?.querySelector<HTMLButtonElement>('[data-testid="updater-view-versions"]');
    expect(viewBtn).not.toBeNull();
    expect(viewBtn?.textContent).toBe("Download manually");
  });

  it("opens the GitHub Releases page from the update dialog (manual rollback path)", async () => {
    const { handleUpdaterEvent, FORMULA_RELEASES_URL } = await import("../updaterUi");

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "Notes" });
    await flushMicrotasks();

    const dialog = document.querySelector('[data-testid="updater-dialog"]') as HTMLElement | null;
    expect(dialog).not.toBeNull();

    const viewBtn = dialog?.querySelector<HTMLButtonElement>('[data-testid="updater-view-versions"]');
    expect(viewBtn).not.toBeNull();

    viewBtn?.click();

    expect(mocks.shellOpen).toHaveBeenCalledTimes(1);
    expect(mocks.shellOpen).toHaveBeenCalledWith(FORMULA_RELEASES_URL);
  });
});
