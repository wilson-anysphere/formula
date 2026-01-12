/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("updaterUi (dialog + download)", () => {
  beforeEach(() => {
    document.body.innerHTML = '<div id="toast-root"></div>';
    vi.resetModules();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it("opens a dialog when an update-available event is received", async () => {
    const handlers = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      handlers.set(eventName, handler);
      return () => handlers.delete(eventName);
    });

    vi.stubGlobal("__TAURI__", { event: { listen } });

    const { installUpdaterUi } = await import("../updaterUi");
    installUpdaterUi();

    handlers.get("update-available")?.({ payload: { source: "startup", version: "9.9.9", body: "Release notes\nLine 2" } });
    await flushMicrotasks();

    const dialog = document.querySelector<HTMLDialogElement>('[data-testid="updater-dialog"]');
    expect(dialog).not.toBeNull();
    expect(dialog?.getAttribute("open") === "" || dialog?.open === true).toBe(true);

    expect(dialog?.querySelector('[data-testid="updater-version"]')?.textContent).toContain("9.9.9");
    expect(dialog?.querySelector('[data-testid="updater-body"]')?.textContent).toContain("Release notes");
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
    installUpdaterUi();

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
});

