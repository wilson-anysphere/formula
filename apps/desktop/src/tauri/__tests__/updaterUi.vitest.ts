// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest";

import { registerAppQuitHandlers } from "../appQuit";

const mocks = vi.hoisted(() => {
  return {
    installUpdateAndRestart: vi.fn<[], Promise<void>>().mockResolvedValue(undefined),
  };
});

vi.mock("../updater", () => ({
  installUpdateAndRestart: mocks.installUpdateAndRestart,
}));

describe("updater restart", () => {
  afterEach(() => {
    registerAppQuitHandlers(null);
    mocks.installUpdateAndRestart.mockReset();
    mocks.installUpdateAndRestart.mockResolvedValue(undefined);
    vi.restoreAllMocks();
    document.body.innerHTML = "";
  });

  it("does not install if the unsaved-changes prompt is cancelled", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;
    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp: vi.fn().mockResolvedValue(undefined),
    });

    vi.spyOn(window, "confirm").mockReturnValue(false);

    const { restartToInstallUpdate } = await import("../updaterUi");
    await restartToInstallUpdate();

    expect(mocks.installUpdateAndRestart).not.toHaveBeenCalled();
  });

  it("installs exactly once when the unsaved-changes prompt is confirmed", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;
    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp: vi.fn().mockResolvedValue(undefined),
    });

    vi.spyOn(window, "confirm").mockReturnValue(true);

    const { restartToInstallUpdate } = await import("../updaterUi");
    await restartToInstallUpdate();

    expect(mocks.installUpdateAndRestart).toHaveBeenCalledTimes(1);
  });

  it("drains backend sync before installing and quitting", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const calls: string[] = [];
    const drainBackendSync = vi.fn(async () => {
      calls.push("drain");
    });
    const quitApp = vi.fn(async () => {
      calls.push("quit");
    });

    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync,
      quitApp,
    });

    mocks.installUpdateAndRestart.mockImplementation(async () => {
      calls.push("install");
    });

    vi.spyOn(window, "confirm").mockReturnValue(true);

    const { restartToInstallUpdate } = await import("../updaterUi");
    await restartToInstallUpdate();

    expect(calls).toEqual(["drain", "install", "quit"]);
  });

  it("aborts restart and shows an error toast if install fails", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const quitApp = vi.fn().mockResolvedValue(undefined);

    registerAppQuitHandlers({
      isDirty: () => true,
      drainBackendSync: vi.fn().mockResolvedValue(undefined),
      quitApp,
    });

    mocks.installUpdateAndRestart.mockRejectedValue(new Error("boom"));
    vi.spyOn(window, "confirm").mockReturnValue(true);

    const { restartToInstallUpdate } = await import("../updaterUi");
    const ok = await restartToInstallUpdate();

    expect(ok).toBe(false);
    expect(quitApp).not.toHaveBeenCalled();

    const toast = document.querySelector<HTMLElement>('[data-testid="toast"]');
    expect(toast).not.toBeNull();
    expect(toast?.dataset.type).toBe("error");
    expect(toast?.textContent).toBe("Failed to restart to install the update.");
  });
});
