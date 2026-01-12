/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const TEST_TIMEOUT_MS = 15_000;

const mocks = vi.hoisted(() => {
  return {
    shellOpen: vi.fn<[string], Promise<void>>().mockResolvedValue(undefined),
  };
});

vi.mock("../shellOpen", () => ({
  shellOpen: mocks.shellOpen,
}));

async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

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

const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

describe("updaterUi (dialog + download)", () => {
  beforeEach(async () => {
    vi.useRealTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';
    const storage = createInMemoryLocalStorage();
    // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless started
    // with `--localstorage-file`. Ensure tests always have a stable in-memory implementation.
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });

    // `updaterUi` holds state in module-level singletons. Tests in other suites may have
    // imported it already; ensure we always start clean.
    const { __resetUpdaterUiStateForTests } = await import("../updaterUi");
    __resetUpdaterUiStateForTests();
  });

  afterEach(() => {
    mocks.shellOpen.mockClear();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    vi.useRealTimers();

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

    document.body.replaceChildren();
  });

  it(
    "does not open a dialog when update-available is received during startup checks",
    async () => {
      const handlers = new Map<string, (event: any) => void>();
      const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
        handlers.set(eventName, handler);
        return () => handlers.delete(eventName);
      });

      vi.stubGlobal("__TAURI__", { event: { listen } });

      const { installUpdaterUi } = await import("../updaterUi");
      await installUpdaterUi();

      handlers
        .get("update-available")
        ?.({ payload: { source: "startup", version: "9.9.9", body: "Release notes\nLine 2" } });
      await flushMicrotasks();

      const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
      expect(dialog).toBeNull();
    },
    30_000,
  );

  it(
    "does not suppress an update dialog when the user clicks manual check during an in-flight startup check",
    async () => {
      // Simulate "Later" suppression for a specific version.
      window.localStorage.setItem("formula.updater.dismissedVersion", "9.9.9");
      window.localStorage.setItem("formula.updater.dismissedAt", String(Date.now()));

      const notifications = await import("../notifications");
      const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

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

      // This update result is treated as "manual" UX (dialog) because the user requested a check,
      // so it should NOT also create a system notification.
      expect(notifySpy).not.toHaveBeenCalled();
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "shows download progress and reveals the restart button once downloaded",
    async () => {
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
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "shows a restart CTA immediately when the update was already downloaded in the background",
    async () => {
      const { handleUpdaterEvent } = await import("../updaterUi");

      await handleUpdaterEvent("update-downloaded", { source: "startup", version: "1.2.3" });
      await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "notes" });
      await flushMicrotasks();

      const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
      expect(restartBtn).not.toBeNull();
      expect(restartBtn?.hidden).toBe(false);

      const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
      expect(downloadBtn).not.toBeNull();
      expect(downloadBtn?.disabled).toBe(true);
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "disables the manual download button and shows progress when a background download is in flight",
    async () => {
      const { handleUpdaterEvent } = await import("../updaterUi");

      await handleUpdaterEvent("update-download-started", { source: "startup", version: "1.2.3" });
      await handleUpdaterEvent("update-download-progress", { source: "startup", version: "1.2.3", percent: 50 });

      await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "notes" });
      await flushMicrotasks();

      const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
      expect(downloadBtn).not.toBeNull();
      expect(downloadBtn?.disabled).toBe(true);

      const progressWrap = document.querySelector<HTMLElement>('[data-testid="updater-progress-wrap"]');
      expect(progressWrap).not.toBeNull();
      expect(progressWrap?.hidden).toBe(false);

      const progress = document.querySelector<HTMLProgressElement>('[data-testid="updater-progress"]');
      expect(progress).not.toBeNull();
      expect(progress?.value).toBe(50);
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "updates the open dialog when the background download completes (no extra toast)",
    async () => {
      const { handleUpdaterEvent } = await import("../updaterUi");

      await handleUpdaterEvent("update-download-started", { source: "startup", version: "1.2.3" });
      await handleUpdaterEvent("update-download-progress", { source: "startup", version: "1.2.3", percent: 50 });
      await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "notes" });
      await flushMicrotasks();

      const toast = document.querySelector('[data-testid="update-ready-toast"]');
      expect(toast).toBeNull();

      const restartBtnBefore = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
      expect(restartBtnBefore).not.toBeNull();
      expect(restartBtnBefore?.hidden).toBe(true);

      await handleUpdaterEvent("update-downloaded", { source: "startup", version: "1.2.3" });
      await flushMicrotasks();

      const restartBtnAfter = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
      expect(restartBtnAfter).not.toBeNull();
      expect(restartBtnAfter?.hidden).toBe(false);

      // The dialog is already the primary UX surface here; avoid a redundant toast.
      const toastAfter = document.querySelector('[data-testid="update-ready-toast"]');
      expect(toastAfter).toBeNull();
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "keeps the update dialog open if the user cancels the unsaved-changes prompt on restart",
    async () => {
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

      const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
      expect(dialog).not.toBeNull();

      const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
      expect(downloadBtn).not.toBeNull();
      downloadBtn?.click();
      await flushMicrotasks(12);

      const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
      expect(restartBtn).not.toBeNull();

      // Simulate a pre-existing dismissal (e.g. user previously hit "Later"). Clicking "Restart now"
      // should clear the suppression state immediately, even if the restart is ultimately cancelled.
      window.localStorage.setItem("formula.updater.dismissedVersion", "1.2.3");
      window.localStorage.setItem("formula.updater.dismissedAt", String(Date.now()));

      // Simulate the user canceling the quit prompt.
      vi.spyOn(window, "confirm").mockReturnValue(false);
      const { registerAppQuitHandlers } = await import("../appQuit");
      const restartApp = vi.fn().mockResolvedValue(undefined);
      const quitApp = vi.fn().mockResolvedValue(undefined);
      registerAppQuitHandlers({
        isDirty: () => true,
        drainBackendSync: vi.fn().mockResolvedValue(undefined),
        restartApp,
        quitApp,
      });

      restartBtn?.click();
      await flushMicrotasks(6);

      expect(install).not.toHaveBeenCalled();
      expect(restartApp).not.toHaveBeenCalled();
      expect(quitApp).not.toHaveBeenCalled();

      expect(window.localStorage.getItem("formula.updater.dismissedVersion")).toBeNull();
      expect(window.localStorage.getItem("formula.updater.dismissedAt")).toBeNull();

      // The dialog should still be open because the restart was cancelled.
      expect(dialog?.getAttribute("open") === "" || dialog?.open === true).toBe(true);

      registerAppQuitHandlers(null);
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "closes the update dialog after a successful restart-to-install flow",
    async () => {
      const handlers = new Map<string, (event: any) => void>();
      const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
        handlers.set(eventName, handler);
        return () => handlers.delete(eventName);
      });

      const download = vi.fn(async (onProgress?: any) => {
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

      const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
      expect(dialog).not.toBeNull();
      expect(dialog?.getAttribute("open") === "" || dialog?.open === true).toBe(true);

      const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
      expect(downloadBtn).not.toBeNull();
      downloadBtn?.click();

      await flushMicrotasks(8);
      expect(download).toHaveBeenCalledTimes(1);

      const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
      expect(restartBtn).not.toBeNull();

      const { registerAppQuitHandlers } = await import("../appQuit");
      const restartApp = vi.fn().mockResolvedValue(undefined);
      registerAppQuitHandlers({
        isDirty: () => false,
        drainBackendSync: vi.fn().mockResolvedValue(undefined),
        quitApp: vi.fn().mockResolvedValue(undefined),
        restartApp,
      });

      restartBtn?.click();
      await flushMicrotasks(8);

      expect(install).toHaveBeenCalledTimes(1);
      expect(restartApp).toHaveBeenCalledTimes(1);
      expect(dialog?.getAttribute("open") === "" || dialog?.open === true).toBe(false);

      registerAppQuitHandlers(null);
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "promotes 'Download manually' when the update download fails",
    async () => {
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
    },
    TEST_TIMEOUT_MS,
  );

  it(
    "opens the GitHub Releases page from the update dialog (manual rollback path)",
    async () => {
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
    },
    TEST_TIMEOUT_MS,
  );
});
