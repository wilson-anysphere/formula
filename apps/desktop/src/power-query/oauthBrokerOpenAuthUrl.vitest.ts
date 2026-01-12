import { afterEach, describe, expect, it, vi } from "vitest";

import { DesktopOAuthBroker } from "./oauthBroker.js";

describe("DesktopOAuthBroker.openAuthUrl", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const originalWindow = (globalThis as any).window;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    (globalThis as any).window = originalWindow;
    vi.restoreAllMocks();
  });

  it("opens https auth URLs via the Tauri shell plugin when available", async () => {
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { plugin: { shell: { open: tauriOpen } } };
    // Guard against accidental webview navigation fallback.
    (globalThis as any).window = { open: vi.fn() };

    const broker = new DesktopOAuthBroker();
    await broker.openAuthUrl("https://example.com/auth");

    expect(tauriOpen).toHaveBeenCalledTimes(1);
    expect(tauriOpen).toHaveBeenCalledWith("https://example.com/auth");
  });

  it("rejects non-http(s) auth URLs", async () => {
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { plugin: { shell: { open: tauriOpen } } };

    const broker = new DesktopOAuthBroker();
    await expect(broker.openAuthUrl("ftp://example.com")).rejects.toThrow(/untrusted protocol/i);
    expect(tauriOpen).not.toHaveBeenCalled();
  });
});

