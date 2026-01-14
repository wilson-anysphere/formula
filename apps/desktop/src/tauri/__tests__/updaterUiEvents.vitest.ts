/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as ui from "../../extensions/ui.js";
import { setLocale, t, tWithVars } from "../../i18n/index.js";
import * as notifications from "../notifications";
import { __resetUpdaterUiStateForTests, handleUpdaterEvent, installUpdaterUi } from "../updaterUi";

async function flushMicrotasks(times = 4): Promise<void> {
  for (let idx = 0; idx < times; idx++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("updaterUi (events)", () => {
  beforeEach(() => {
    setLocale("en-US");
    __resetUpdaterUiStateForTests();
  });

  afterEach(() => {
    try {
      vi.runOnlyPendingTimers();
    } catch {
      // Timers weren't mocked.
    }
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it("shows + focuses the main window before rendering manual-check feedback", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-not-available", { source: "manual" });

    expect(show).toHaveBeenCalledTimes(1);
    expect(setFocus).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledTimes(1);

    const toast = document.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toBe(t("updater.upToDate"));

    expect(show.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
    expect(setFocus.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
  });

  it("renders an 'already checking' toast for repeated manual update checks", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    const handlers = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      handlers.set(eventName, handler);
      return () => {};
    });

    await installUpdaterUi(listen);

    expect(listen).toHaveBeenCalledWith("update-check-already-running", expect.any(Function));

    handlers.get("update-check-already-running")?.({ payload: { source: "manual" } });

    // `installUpdaterUi` wires handlers using `void handleUpdaterEvent(...)`; flush a few
    // microtasks so async window show/focus completes before we assert.
    await flushMicrotasks(10);

    expect(show).toHaveBeenCalledTimes(1);
    expect(setFocus).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledTimes(1);

    const toast = document.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toBe(t("updater.alreadyChecking"));

    expect(show.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
    expect(setFocus.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
  });

  it("ignores 'already running' events emitted during startup checks", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-check-already-running", { source: "startup" });

    expect(show).not.toHaveBeenCalled();
    expect(setFocus).not.toHaveBeenCalled();
    expect(toastSpy).not.toHaveBeenCalled();
  });

  it("does not show toasts or focus the window for startup update-not-available events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-not-available", { source: "startup" });

    expect(show).not.toHaveBeenCalled();
    expect(setFocus).not.toHaveBeenCalled();
    expect(toastSpy).not.toHaveBeenCalled();
  });

  it("does not show toasts or focus the window for startup update-check-error events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-check-error", { source: "startup", message: "network down" });

    expect(show).not.toHaveBeenCalled();
    expect(setFocus).not.toHaveBeenCalled();
    expect(toastSpy).not.toHaveBeenCalled();
  });

  it("does not show toasts or focus the window for startup update-check-started events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-check-started", { source: "startup" });

    expect(show).not.toHaveBeenCalled();
    expect(setFocus).not.toHaveBeenCalled();
    expect(toastSpy).not.toHaveBeenCalled();
  });

  it("shows toasts for manual update-check-started and update-check-error events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-check-started", { source: "manual" });
    await handleUpdaterEvent("update-check-error", { source: "manual", message: "network down" });

    expect(toastSpy).toHaveBeenCalledTimes(2);

    const toasts = document.querySelectorAll<HTMLElement>('[data-testid="toast"]');
    expect(toasts).toHaveLength(2);
    expect(toasts[0]?.textContent).toBe(t("updater.checking"));
    expect(toasts[1]?.dataset.type).toBe("error");
    expect(toasts[1]?.textContent).toBe(tWithVars("updater.errorWithMessage", { message: "network down" }));
  });

  it("surfaces startup completion events after the user clicks 'Check for Updates' during a startup check", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-check-already-running", { source: "manual" });
    await handleUpdaterEvent("update-not-available", { source: "startup" });

    // The "already running" event is manual (and can be triggered while the app is hidden),
    // so we show/focus the main window once before surfacing any user-visible feedback.
    expect(show).toHaveBeenCalledTimes(1);
    expect(setFocus).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledTimes(2);

    const toasts = document.querySelectorAll('[data-testid="toast"]');
    expect(toasts).toHaveLength(2);
    expect(toasts[1]?.textContent).toBe(t("updater.upToDate"));
  });

  it("sends a system notification for startup update-available events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);
    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Bug fixes" });

    expect(notifySpy).toHaveBeenCalledTimes(1);
    const appName = t("app.title");
    expect(notifySpy).toHaveBeenCalledWith({
      title: t("updater.updateAvailableTitle"),
      body: tWithVars("updater.systemNotificationBodyWithNotes", { appName, version: "1.2.3", notes: "Bug fixes" }),
    });
    expect(toastSpy).not.toHaveBeenCalled();
  });

  it("does not show/focus the window for startup update-available events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Bug fixes" });

    expect(notifySpy).toHaveBeenCalledTimes(1);
    expect(show).not.toHaveBeenCalled();
    expect(setFocus).not.toHaveBeenCalled();
  });

  it("does not send a system notification for manual update-available events", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

    await handleUpdaterEvent("update-available", { source: "manual", version: "9.9.9" });

    expect(notifySpy).not.toHaveBeenCalled();

    const dialog = document.querySelector('[data-testid="updater-dialog"]');
    expect(dialog).toBeTruthy();
    expect(document.querySelector('[data-testid="updater-version"]')?.textContent).toBe(
      tWithVars("updater.updateAvailableMessage", { version: "9.9.9" }),
    );
  });

  it("trims release notes text for manual update-available dialogs", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    await handleUpdaterEvent("update-available", { source: "manual", version: "1.2.3", body: "  Bug fixes  " });

    const dialog = document.querySelector('[data-testid="updater-dialog"]');
    expect(dialog).toBeTruthy();
    expect(document.querySelector('[data-testid="updater-body"]')?.textContent).toBe("Bug fixes");
  });

  it("shows an update dialog (and skips the system notification) when a manual check is queued behind an in-flight startup check", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const notifySpy = vi.spyOn(notifications, "notify").mockResolvedValue(undefined);

    await handleUpdaterEvent("update-check-already-running", { source: "manual" });
    await handleUpdaterEvent("update-available", { source: "startup", version: "1.2.3", body: "Bug fixes" });

    expect(notifySpy).not.toHaveBeenCalled();

    const dialog = document.querySelector('[data-testid="updater-dialog"]');
    expect(dialog).toBeTruthy();
    expect(document.querySelector('[data-testid="updater-version"]')?.textContent).toBe(
      tWithVars("updater.updateAvailableMessage", { version: "1.2.3" }),
    );
  });

  it("surfaces manual check results as toasts", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const listeners = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      listeners.set(eventName, handler);
      return () => listeners.delete(eventName);
    });

    vi.stubGlobal("__TAURI__", {
      event: { listen },
    });

    await installUpdaterUi();

    listeners.get("update-check-started")?.({ payload: { source: "manual" } });
    listeners.get("update-not-available")?.({ payload: { source: "manual" } });
    listeners.get("update-check-error")?.({ payload: { source: "manual", error: "network down" } });

    // `installUpdaterUi` wires handlers using `void handleUpdaterEvent(...)`; flush a few
    // microtasks so async window show/focus completes before we assert.
    await flushMicrotasks();

    const toastEls = Array.from(document.querySelectorAll('[data-testid="toast"]'));
    expect(toastEls).toHaveLength(3);

    const toasts = toastEls.map((el) => el.textContent ?? "");
    expect(toasts.join("\n")).toContain(t("updater.checking"));
    expect(toasts.join("\n")).toContain(t("updater.upToDate"));
    expect(toasts.join("\n")).toContain("network down");
  });

  it("trims update-check-error messages before surfacing toasts", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const toastSpy = vi.spyOn(ui, "showToast");
    await handleUpdaterEvent("update-check-error", { source: "manual", error: "  network down  " });

    expect(toastSpy).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledWith(tWithVars("updater.errorWithMessage", { message: "network down" }), "error");
  });

  it("updates progress UI when downloading an update", async () => {
    const listeners = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      listeners.set(eventName, handler);
      return () => listeners.delete(eventName);
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

    await installUpdaterUi();

    listeners.get("update-available")?.({ payload: { version: "1.2.3", body: "notes", source: "manual" } });

    // `installUpdaterUi` wires handlers using `void handleUpdaterEvent(...)`; flush a few
    // microtasks so async window show/focus completes before we query the dialog controls.
    await flushMicrotasks();

    const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
    expect(downloadBtn).not.toBeNull();

    downloadBtn?.click();
    // `startUpdateDownload()` yields to a timer so the dialog can render the "Downloading…"
    // state before potentially-fast updater checks/downloads complete.
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    await flushMicrotasks(8);

    const progress = document.querySelector<HTMLProgressElement>('[data-testid="updater-progress"]');
    const progressText = document.querySelector<HTMLElement>('[data-testid="updater-progress-text"]');

    expect(download).toHaveBeenCalledTimes(1);
    expect(progress?.value).toBe(100);
    expect(progressText?.textContent).toBe("100%");

    const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
    expect(restartBtn?.style.display === "none").toBe(false);
  });

  it("persists dismissal when the user clicks 'Later' on the update-ready toast", async () => {
    // Provide a stable in-memory localStorage for this test (Node can throw on globalThis.localStorage).
    const store = new Map<string, string>();
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      value: {
        getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
        setItem: (key: string, value: string) => {
          store.set(String(key), String(value));
        },
        removeItem: (key: string) => {
          store.delete(String(key));
        },
      },
    });
    Object.defineProperty(window, "localStorage", { configurable: true, value: (globalThis as any).localStorage });

    document.body.innerHTML = '<div id="toast-root"></div>';

    const { handleUpdaterEvent } = await import("../updaterUi");

    await handleUpdaterEvent("update-downloaded", { source: "startup", version: "9.9.9" });
    const laterBtn = document.querySelector<HTMLButtonElement>('[data-testid="update-ready-later"]');
    expect(laterBtn).not.toBeNull();
    laterBtn?.click();

    expect(window.localStorage.getItem("formula.updater.dismissedVersion")).toBe("9.9.9");
    expect(Number(window.localStorage.getItem("formula.updater.dismissedAt"))).toBeGreaterThan(0);
  });

  it("clears dismissal when the user clicks restart on the update-ready toast", async () => {
    const store = new Map<string, string>();
    const storage = {
      getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
      setItem: (key: string, value: string) => {
        store.set(String(key), String(value));
      },
      removeItem: (key: string) => {
        store.delete(String(key));
      },
    };
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });

    document.body.innerHTML = '<div id="toast-root"></div>';

    const { registerAppQuitHandlers } = await import("../appQuit");
    registerAppQuitHandlers({
      // Force a confirm prompt and then cancel it so we don't attempt to install.
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp: vi.fn().mockResolvedValue(undefined),
    });
    vi.spyOn(window, "confirm").mockReturnValue(false);

    const { handleUpdaterEvent } = await import("../updaterUi");
    await handleUpdaterEvent("update-downloaded", { source: "startup", version: "9.9.9" });

    // Simulate the user previously dismissing the update prompt.
    window.localStorage.setItem("formula.updater.dismissedVersion", "9.9.9");
    window.localStorage.setItem("formula.updater.dismissedAt", String(Date.now()));

    const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="update-ready-restart"]');
    expect(restartBtn).not.toBeNull();
    restartBtn?.click();

    expect(window.localStorage.getItem("formula.updater.dismissedVersion")).toBeNull();
    expect(window.localStorage.getItem("formula.updater.dismissedAt")).toBeNull();

    registerAppQuitHandlers(null);
  });
});
