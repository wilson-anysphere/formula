import { afterEach, describe, expect, it, vi } from "vitest";

import { shellOpen } from "../shellOpen";

describe("shellOpen", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const originalWindow = (globalThis as any).window;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    (globalThis as any).window = originalWindow;
    vi.restoreAllMocks();
  });

  it("invokes the Rust open_external_url command in Tauri builds", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke }, plugin: { shell: { open: tauriOpen } } };
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await shellOpen("https://example.com");

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: "https://example.com" });
    expect(tauriOpen).not.toHaveBeenCalled();
    expect(winOpen).not.toHaveBeenCalled();
  });

  it("allows mailto: URLs in Tauri builds", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await shellOpen("mailto:test@example.com");

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: "mailto:test@example.com" });
  });

  it("falls back to window.open in web builds", async () => {
    (globalThis as any).__TAURI__ = undefined;
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await shellOpen("https://example.com");

    expect(winOpen).toHaveBeenCalledTimes(1);
    expect(winOpen).toHaveBeenCalledWith("https://example.com", "_blank", "noopener,noreferrer");
  });

  it("does not crash when __TAURI__.shell access throws and the shell plugin is available", async () => {
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    const tauri: any = { plugin: { shell: { open: tauriOpen } } };
    Object.defineProperty(tauri, "shell", {
      configurable: true,
      get() {
        throw new Error("Blocked shell access");
      },
    });
    (globalThis as any).__TAURI__ = tauri;
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await shellOpen("https://example.com");

    expect(tauriOpen).toHaveBeenCalledTimes(1);
    expect(tauriOpen).toHaveBeenCalledWith("https://example.com");
    expect(winOpen).not.toHaveBeenCalled();
  });

  it("blocks javascript: URLs even when the Tauri API is available", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await expect(shellOpen("javascript:alert(1)")).rejects.toThrow(/blocked protocol/i);
    expect(invoke).not.toHaveBeenCalled();
  });

  it("blocks data: URLs even when the Tauri API is available", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await expect(shellOpen("data:text/plain,hello")).rejects.toThrow(/blocked protocol/i);
    expect(invoke).not.toHaveBeenCalled();
  });

  it("blocks file: URLs even when the Tauri API is available", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await expect(shellOpen("file:///etc/passwd")).rejects.toThrow(/blocked protocol/i);
    expect(invoke).not.toHaveBeenCalled();
  });

  it("blocks http(s) URLs with userinfo even when the Tauri API is available", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await expect(shellOpen("https://user:pass@example.com")).rejects.toThrow(/username\/password/i);
    expect(invoke).not.toHaveBeenCalled();
  });

  it("blocks http(s) URLs with userinfo in web builds", async () => {
    (globalThis as any).__TAURI__ = undefined;
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await expect(shellOpen("https://user:pass@example.com")).rejects.toThrow(/username\/password/i);
    expect(winOpen).not.toHaveBeenCalled();
  });

  it("does not fall back to window.open when __TAURI__ is present but the invoke API is missing", async () => {
    (globalThis as any).__TAURI__ = { plugin: {} };
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await expect(shellOpen("https://example.com")).rejects.toThrow(/invoke api unavailable/i);
    expect(winOpen).not.toHaveBeenCalled();
  });
});
