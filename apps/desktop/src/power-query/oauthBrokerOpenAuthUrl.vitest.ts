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

  it("opens https auth URLs via the Rust open_external_url command when running under Tauri", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke }, plugin: { shell: { open: tauriOpen } } };
    // Guard against accidental webview navigation fallback.
    const windowOpen = vi.fn();
    (globalThis as any).window = { open: windowOpen };

    const broker = new DesktopOAuthBroker();
    await broker.openAuthUrl("https://example.com/auth");

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: "https://example.com/auth" });
    expect(windowOpen).not.toHaveBeenCalled();
    expect(tauriOpen).not.toHaveBeenCalled();
  });

  it("rejects non-http(s) auth URLs", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    await expect(broker.openAuthUrl("ftp://example.com")).rejects.toThrow(/untrusted protocol/i);
    expect(invoke).not.toHaveBeenCalled();
  });
});
