import { afterEach, describe, expect, it, vi } from "vitest";

import { openExternalHyperlink } from "./openExternal.js";

describe("openExternalHyperlink", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  it("opens allowlisted https links without prompting", async () => {
    const shellOpen = vi.fn().mockResolvedValue(undefined);
    const confirmUntrustedProtocol = vi.fn().mockResolvedValue(true);
    const permissions = { request: vi.fn().mockResolvedValue(true) };

    await expect(
      openExternalHyperlink("https://example.com", { shellOpen, confirmUntrustedProtocol, permissions }),
    ).resolves.toBe(true);

    expect(confirmUntrustedProtocol).not.toHaveBeenCalled();
    expect(permissions.request).toHaveBeenCalledTimes(1);
    expect(permissions.request).toHaveBeenCalledWith("external_navigation", { uri: "https://example.com", protocol: "https" });
    expect(shellOpen).toHaveBeenCalledTimes(1);
    expect(shellOpen).toHaveBeenCalledWith("https://example.com");
  });

  it("blocks file: links", async () => {
    const shellOpen = vi.fn().mockResolvedValue(undefined);
    const confirmUntrustedProtocol = vi.fn().mockResolvedValue(true);
    const permissions = { request: vi.fn().mockResolvedValue(true) };

    await expect(openExternalHyperlink("file:///etc/passwd", { shellOpen, confirmUntrustedProtocol, permissions })).resolves.toBe(false);

    expect(confirmUntrustedProtocol).not.toHaveBeenCalled();
    expect(permissions.request).not.toHaveBeenCalled();
    expect(shellOpen).not.toHaveBeenCalled();
  });

  it("blocks http(s) links with userinfo", async () => {
    (globalThis as any).__TAURI__ = undefined;

    const shellOpen = vi.fn().mockResolvedValue(undefined);
    const confirmUntrustedProtocol = vi.fn().mockResolvedValue(true);
    const permissions = { request: vi.fn().mockResolvedValue(true) };

    await expect(
      openExternalHyperlink("https://user:pass@example.com", { shellOpen, confirmUntrustedProtocol, permissions }),
    ).resolves.toBe(false);

    expect(confirmUntrustedProtocol).not.toHaveBeenCalled();
    expect(permissions.request).not.toHaveBeenCalled();
    expect(shellOpen).not.toHaveBeenCalled();
  });

  it("blocks non-allowlisted schemes in Tauri builds without prompting", async () => {
    (globalThis as any).__TAURI__ = { core: { invoke: vi.fn() } };

    const shellOpen = vi.fn().mockResolvedValue(undefined);
    const confirmUntrustedProtocol = vi.fn().mockResolvedValue(true);

    await expect(openExternalHyperlink("ftp://example.com", { shellOpen, confirmUntrustedProtocol })).resolves.toBe(false);

    expect(confirmUntrustedProtocol).not.toHaveBeenCalled();
    expect(shellOpen).not.toHaveBeenCalled();
  });

  it("does not allow overriding the protocol allowlist in Tauri builds", async () => {
    (globalThis as any).__TAURI__ = { core: { invoke: vi.fn() } };

    const shellOpen = vi.fn().mockResolvedValue(undefined);
    const confirmUntrustedProtocol = vi.fn().mockResolvedValue(true);
    const allowedProtocols = new Set<string>(["ftp", "http", "https", "mailto"]);

    await expect(
      openExternalHyperlink("ftp://example.com", { shellOpen, confirmUntrustedProtocol, allowedProtocols }),
    ).resolves.toBe(false);

    expect(confirmUntrustedProtocol).not.toHaveBeenCalled();
    expect(shellOpen).not.toHaveBeenCalled();
  });

  it("prompts for non-allowlisted schemes in web builds", async () => {
    (globalThis as any).__TAURI__ = undefined;

    const shellOpen = vi.fn().mockResolvedValue(undefined);
    const confirmUntrustedProtocol = vi.fn().mockResolvedValue(true);

    await expect(openExternalHyperlink("ftp://example.com", { shellOpen, confirmUntrustedProtocol })).resolves.toBe(true);

    expect(confirmUntrustedProtocol).toHaveBeenCalledTimes(1);
    expect(shellOpen).toHaveBeenCalledTimes(1);
    expect(shellOpen).toHaveBeenCalledWith("ftp://example.com");
  });
});
