// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { registerAppQuitHandlers, requestAppQuit, requestAppRestart } from "../appQuit";

describe("appQuit helpers", () => {
  beforeEach(() => {
    vi.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(() => {
    registerAppQuitHandlers(null);
    vi.restoreAllMocks();
    document.body.innerHTML = "";
  });

  it("continues quitting even if Workbook_BeforeClose fails", async () => {
    const runWorkbookBeforeClose = vi.fn(async () => {
      throw new Error("macro crash");
    });
    const drainBackendSync = vi.fn().mockResolvedValue(undefined);
    const quitApp = vi.fn().mockResolvedValue(undefined);

    registerAppQuitHandlers({
      isDirty: () => true,
      runWorkbookBeforeClose,
      drainBackendSync,
      quitApp,
    });

    vi.spyOn(window, "confirm").mockReturnValue(true);

    const ok = await requestAppQuit();
    expect(ok).toBe(true);
    expect(runWorkbookBeforeClose).toHaveBeenCalledTimes(1);
    expect(drainBackendSync).toHaveBeenCalledTimes(1);
    expect(quitApp).toHaveBeenCalledTimes(1);
  });

  it("falls back to quitApp if restartApp throws", async () => {
    const drainBackendSync = vi.fn().mockResolvedValue(undefined);
    const restartApp = vi.fn(async () => {
      throw new Error("restart failed");
    });
    const quitApp = vi.fn().mockResolvedValue(undefined);

    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync,
      quitApp,
      restartApp,
    });

    vi.spyOn(window, "confirm").mockReturnValue(true);

    const ok = await requestAppRestart({ beforeQuit: vi.fn().mockResolvedValue(undefined) });
    expect(ok).toBe(true);
    expect(restartApp).toHaveBeenCalledTimes(1);
    expect(quitApp).toHaveBeenCalledTimes(1);
  });

  it("does not restart if the unsaved-changes prompt is cancelled", async () => {
    const drainBackendSync = vi.fn().mockResolvedValue(undefined);
    const restartApp = vi.fn().mockResolvedValue(undefined);
    const quitApp = vi.fn().mockResolvedValue(undefined);

    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync,
      quitApp,
      restartApp,
    });

    vi.spyOn(window, "confirm").mockReturnValue(false);

    const beforeQuit = vi.fn().mockResolvedValue(undefined);
    const ok = await requestAppRestart({ beforeQuit });
    expect(ok).toBe(false);
    expect(beforeQuit).not.toHaveBeenCalled();
    expect(restartApp).not.toHaveBeenCalled();
    expect(quitApp).not.toHaveBeenCalled();
  });
});
