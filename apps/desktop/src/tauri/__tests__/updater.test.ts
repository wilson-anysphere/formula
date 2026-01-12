import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { checkForUpdatesFromCommandPalette } from "../updater";

describe("checkForUpdatesFromCommandPalette", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  beforeEach(() => {
    const invoke = vi.fn().mockResolvedValue(null);
    (globalThis as any).__TAURI__ = { core: { invoke } };
  });

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  it("invokes the backend update check command with a manual source", async () => {
    await checkForUpdatesFromCommandPalette();
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;
    expect(invoke).toHaveBeenCalledWith("check_for_updates", { source: "manual" });
  });
});

