import { afterEach, describe, expect, it, vi } from "vitest";

import { DesktopOAuthBroker } from "./oauthBroker.js";

describe("DesktopOAuthBroker.openAuthUrl", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const originalWindow = (globalThis as any).window;

  function buildAuthUrl(redirectUri: string): string {
    return `https://example.com/oauth/authorize?redirect_uri=${encodeURIComponent(redirectUri)}`;
  }

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    (globalThis as any).window = originalWindow;
    vi.restoreAllMocks();
  });

  it("opens https auth URLs via the Rust open_external_url command when running under Tauri", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };
    // Guard against accidental webview navigation fallback.
    const windowOpen = vi.fn();
    (globalThis as any).window = { open: windowOpen };

    const broker = new DesktopOAuthBroker();
    await broker.openAuthUrl("https://example.com/auth");

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: "https://example.com/auth" });
    expect(windowOpen).not.toHaveBeenCalled();
  });

  it.each([
    ["http://127.0.0.1:1234/callback"],
    ["http://localhost:1234/callback"],
    ["http://[::1]:1234/callback"],
  ])("starts a loopback listener for %s redirect URIs when running under Tauri", async (redirectUri) => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };
    // Guard against accidental webview navigation fallback.
    const windowOpen = vi.fn();
    (globalThis as any).window = { open: windowOpen };

    const broker = new DesktopOAuthBroker();
    const authUrl = buildAuthUrl(redirectUri);
    await broker.openAuthUrl(authUrl);

    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke.mock.calls[0]).toEqual(["oauth_loopback_listen", { redirect_uri: redirectUri }]);
    expect(invoke.mock.calls[1]).toEqual(["open_external_url", { url: authUrl }]);
    expect(windowOpen).not.toHaveBeenCalled();
  });

  it("does not start a loopback listener for https redirect URIs", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    const authUrl = buildAuthUrl("https://127.0.0.1:1234/callback");
    await broker.openAuthUrl(authUrl);

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: authUrl });
  });

  it("does not start a loopback listener for http loopback redirect URIs without an explicit port", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    const authUrl = buildAuthUrl("http://127.0.0.1/callback");
    await broker.openAuthUrl(authUrl);

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: authUrl });
  });

  it("does not start a loopback listener for non-loopback http redirect URIs", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    const authUrl = buildAuthUrl("http://example.com:1234/callback");
    await broker.openAuthUrl(authUrl);

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: authUrl });
  });

  it("rejects non-http(s) auth URLs", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    await expect(broker.openAuthUrl("ftp://example.com")).rejects.toThrow(/untrusted protocol/i);
    expect(invoke).not.toHaveBeenCalled();
  });

  it("rejects auth URLs containing userinfo", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    await expect(broker.openAuthUrl("https://user:pass@example.com/auth")).rejects.toThrow(/username\\/password/i);
    expect(invoke).not.toHaveBeenCalled();
  });

  it.each([
    ["127.0.0.1", "http://127.0.0.1:4242/oauth/callback"],
    ["localhost", "http://localhost:4242/oauth/callback"],
    ["::1", "http://[::1]:4242/oauth/callback"],
  ] as const)("starts the Rust loopback listener for %s redirect URIs", async (_hostLabel, redirectUri) => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    // Guard against accidental webview navigation fallback.
    const windowOpen = vi.fn();
    (globalThis as any).window = { open: windowOpen };

    const broker = new DesktopOAuthBroker();
    const authUrl = new URL("https://example.com/oauth/authorize");
    authUrl.searchParams.set("redirect_uri", redirectUri);

    await broker.openAuthUrl(authUrl.toString());

    expect(invoke).toHaveBeenCalledTimes(2);
    expect(invoke).toHaveBeenNthCalledWith(1, "oauth_loopback_listen", { redirect_uri: redirectUri });
    expect(invoke).toHaveBeenNthCalledWith(2, "open_external_url", { url: authUrl.toString() });
    expect(windowOpen).not.toHaveBeenCalled();
  });

  it("does not start the loopback listener for non-loopback redirect URIs", async () => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    const authUrl = new URL("https://example.com/oauth/authorize");
    authUrl.searchParams.set("redirect_uri", "http://example.com:4242/oauth/callback");

    await broker.openAuthUrl(authUrl.toString());

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: authUrl.toString() });
    expect(invoke).not.toHaveBeenCalledWith("oauth_loopback_listen", expect.anything());
  });

  it.each([
    ["missing port", "http://localhost/oauth/callback"],
    ["https scheme", "https://localhost:4242/oauth/callback"],
    ["non-loopback IPv4", "http://127.0.0.2:4242/oauth/callback"],
  ] as const)("does not start the loopback listener when redirect_uri is %s", async (_label, redirectUri) => {
    const invoke = vi.fn().mockResolvedValue(undefined);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const broker = new DesktopOAuthBroker();
    const authUrl = new URL("https://example.com/oauth/authorize");
    authUrl.searchParams.set("redirect_uri", redirectUri);

    await broker.openAuthUrl(authUrl.toString());

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("open_external_url", { url: authUrl.toString() });
    expect(invoke).not.toHaveBeenCalledWith("oauth_loopback_listen", expect.anything());
  });
});
