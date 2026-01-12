// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { registerAppQuitHandlers } from "../appQuit";
import { setLocale, t } from "../../i18n/index.js";

const mocks = vi.hoisted(() => {
  return {
    installUpdateAndRestart: vi.fn<[], Promise<void>>().mockResolvedValue(undefined),
  };
});

vi.mock("../updater", () => ({
  installUpdateAndRestart: mocks.installUpdateAndRestart,
}));

describe("updater restart", () => {
  beforeEach(() => {
    setLocale("en-US");
  });

  afterEach(() => {
    registerAppQuitHandlers(null);
    mocks.installUpdateAndRestart.mockReset();
    mocks.installUpdateAndRestart.mockResolvedValue(undefined);
    vi.restoreAllMocks();
    document.body.innerHTML = "";
  });

  it("does not install if the unsaved-changes prompt is cancelled", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;
    const restartApp = vi.fn().mockResolvedValue(undefined);
    const quitApp = vi.fn().mockResolvedValue(undefined);
    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp,
      restartApp,
    });

    vi.spyOn(window, "confirm").mockReturnValue(false);

    const { restartToInstallUpdate } = await import("../updaterUi");
    await restartToInstallUpdate();

    expect(mocks.installUpdateAndRestart).not.toHaveBeenCalled();
    expect(restartApp).not.toHaveBeenCalled();
    expect(quitApp).not.toHaveBeenCalled();
  });

  it("installs exactly once when the unsaved-changes prompt is confirmed", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;
    const restartApp = vi.fn().mockResolvedValue(undefined);
    const quitApp = vi.fn().mockResolvedValue(undefined);
    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp,
      restartApp,
    });

    vi.spyOn(window, "confirm").mockReturnValue(true);

    const { restartToInstallUpdate } = await import("../updaterUi");
    await restartToInstallUpdate();

    expect(mocks.installUpdateAndRestart).toHaveBeenCalledTimes(1);
    expect(restartApp).toHaveBeenCalledTimes(1);
    expect(quitApp).not.toHaveBeenCalled();
  });

  it("drains backend sync before installing and restarting", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const calls: string[] = [];
    const drainBackendSync = vi.fn(async () => {
      calls.push("drain");
    });
    const quitApp = vi.fn(async () => {
      calls.push("quit");
    });
    const restartApp = vi.fn(async () => {
      calls.push("restart");
    });

    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync,
      quitApp,
      restartApp,
    });

    mocks.installUpdateAndRestart.mockImplementation(async () => {
      calls.push("install");
    });

    vi.spyOn(window, "confirm").mockReturnValue(true);

    const { restartToInstallUpdate } = await import("../updaterUi");
    await restartToInstallUpdate();

    expect(calls).toEqual(["drain", "install", "restart"]);
    expect(quitApp).not.toHaveBeenCalled();
  });

  it("aborts restart and shows an error toast if install fails", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const quitApp = vi.fn().mockResolvedValue(undefined);
    const restartApp = vi.fn().mockResolvedValue(undefined);

    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp,
      restartApp,
    });

    mocks.installUpdateAndRestart.mockRejectedValue(new Error("boom"));
    vi.spyOn(window, "confirm").mockReturnValue(true);

    const { restartToInstallUpdate } = await import("../updaterUi");
    const ok = await restartToInstallUpdate();

    expect(ok).toBe(false);
    expect(quitApp).not.toHaveBeenCalled();
    expect(restartApp).not.toHaveBeenCalled();

    const toast = document.querySelector<HTMLElement>('[data-testid="toast"]');
    expect(toast).not.toBeNull();
    expect(toast?.dataset.type).toBe("error");
    expect(toast?.textContent).toBe(t("updater.restartFailed"));
  });
});
