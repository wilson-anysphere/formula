// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest";

import * as ui from "../../extensions/ui.js";
import { handleUpdaterEvent, installUpdaterUi } from "../updaterUi";

async function flushMicrotasks(times = 4): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("updaterUi", () => {
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
    expect(toast?.textContent).toContain("up to date");

    expect(show.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
    expect(setFocus.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
  });

  it("wires the updater event listeners and renders an 'already checking' toast for repeated manual checks", async () => {
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

    installUpdaterUi(listen);

    expect(listen).toHaveBeenCalledWith("update-check-already-running", expect.any(Function));

    handlers.get("update-check-already-running")?.({ payload: { source: "manual" } });

    // `installUpdaterUi` wires handlers using `void handleUpdaterEvent(...)`; flush a few
    // microtasks so async window show/focus completes before we assert.
    for (let idx = 0; idx < 10; idx++) {
      // eslint-disable-next-line no-await-in-loop
      await Promise.resolve();
    }

    expect(show).toHaveBeenCalledTimes(1);
    expect(setFocus).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledTimes(1);

    const toast = document.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("Already checking");

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

  it("installs updater event listeners and surfaces manual check results as toasts", async () => {
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

    installUpdaterUi();

    listeners.get("update-check-started")?.({ payload: { source: "manual" } });
    listeners.get("update-not-available")?.({ payload: { source: "manual" } });
    listeners.get("update-check-error")?.({ payload: { source: "manual", error: "network down" } });

    const toasts = Array.from(document.querySelectorAll('[data-testid="toast"]')).map((el) => el.textContent ?? "");
    expect(toasts.join("\n")).toContain("Checking for updates");
    expect(toasts.join("\n")).toContain("up to date");
    expect(toasts.join("\n")).toContain("network down");
  });

  it("opens an update dialog when update-available is emitted", async () => {
    const listeners = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      listeners.set(eventName, handler);
      return () => listeners.delete(eventName);
    });

    vi.stubGlobal("__TAURI__", {
      event: { listen },
    });

    installUpdaterUi();

    listeners.get("update-available")?.({ payload: { version: "9.9.9", body: "Release notes\nLine 2", source: "startup" } });

    const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
    expect(dialog).not.toBeNull();
    expect(dialog?.getAttribute("open") === "" || dialog?.open === true).toBe(true);

    expect(dialog?.querySelector('[data-testid="updater-version"]')?.textContent).toContain("9.9.9");
    expect(dialog?.querySelector('[data-testid="updater-body"]')?.textContent).toContain("Release notes");
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

    installUpdaterUi();

    listeners.get("update-available")?.({ payload: { version: "1.2.3", body: "notes", source: "manual" } });

    const downloadBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-download"]');
    expect(downloadBtn).not.toBeNull();

    downloadBtn?.click();
    await flushMicrotasks(8);

    const progress = document.querySelector<HTMLProgressElement>('[data-testid="updater-progress"]');
    const progressText = document.querySelector<HTMLElement>('[data-testid="updater-progress-text"]');

    expect(download).toHaveBeenCalledTimes(1);
    expect(progress?.value).toBe(100);
    expect(progressText?.textContent).toBe("100%");

    const restartBtn = document.querySelector<HTMLButtonElement>('[data-testid="updater-restart"]');
    expect(restartBtn?.style.display === "none").toBe(false);
  });
});

