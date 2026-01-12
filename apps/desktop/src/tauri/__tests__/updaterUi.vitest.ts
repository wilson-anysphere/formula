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
    mocks.installUpdateAndRestart.mockClear();
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
});

