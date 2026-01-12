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

  it("prefers the Tauri shell plugin when available", async () => {
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { plugin: { shell: { open: tauriOpen } } };
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await shellOpen("https://example.com");

    expect(tauriOpen).toHaveBeenCalledTimes(1);
    expect(tauriOpen).toHaveBeenCalledWith("https://example.com");
    expect(winOpen).not.toHaveBeenCalled();
  });

  it("falls back to window.open in web builds", async () => {
    (globalThis as any).__TAURI__ = undefined;
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await shellOpen("https://example.com");

    expect(winOpen).toHaveBeenCalledTimes(1);
    expect(winOpen).toHaveBeenCalledWith("https://example.com", "_blank", "noopener,noreferrer");
  });

  it("blocks javascript: URLs even when the shell API is available", async () => {
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { plugin: { shell: { open: tauriOpen } } };

    await expect(shellOpen("javascript:alert(1)")).rejects.toThrow(/blocked protocol/i);
    expect(tauriOpen).not.toHaveBeenCalled();
  });

  it("blocks data: URLs even when the shell API is available", async () => {
    const tauriOpen = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { plugin: { shell: { open: tauriOpen } } };

    await expect(shellOpen("data:text/plain,hello")).rejects.toThrow(/blocked protocol/i);
    expect(tauriOpen).not.toHaveBeenCalled();
  });

  it("does not fall back to window.open when __TAURI__ is present but the shell plugin is missing", async () => {
    (globalThis as any).__TAURI__ = { core: { invoke: vi.fn() } };
    const winOpen = vi.fn();
    (globalThis as any).window = { open: winOpen };

    await expect(shellOpen("https://example.com")).rejects.toThrow(/shell plugin unavailable/i);
    expect(winOpen).not.toHaveBeenCalled();
  });
});
