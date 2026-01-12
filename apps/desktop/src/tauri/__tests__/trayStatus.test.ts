import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { setTrayStatus } from "../trayStatus";

describe("trayStatus", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  beforeEach(() => {
    (globalThis as any).__TAURI__ = undefined;
  });

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  it("no-ops outside of Tauri", async () => {
    await expect(setTrayStatus("idle")).resolves.toBeUndefined();
  });

  it("invokes set_tray_status when __TAURI__.core.invoke is present", async () => {
    const invoke = vi.fn().mockResolvedValue(null);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await setTrayStatus("syncing");

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("set_tray_status", { status: "syncing" });
  });
});

