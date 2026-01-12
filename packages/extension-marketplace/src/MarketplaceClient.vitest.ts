import { afterEach, describe, expect, it, vi } from "vitest";

import { MarketplaceClient, normalizeMarketplaceBaseUrl } from "./MarketplaceClient";

describe("normalizeMarketplaceBaseUrl", () => {
  it("defaults to /api when unset", () => {
    expect(normalizeMarketplaceBaseUrl("")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("   ")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl(undefined as any)).toBe("/api");
    expect(normalizeMarketplaceBaseUrl(null as any)).toBe("/api");
  });

  it("trims and strips trailing slashes", () => {
    expect(normalizeMarketplaceBaseUrl(" /api ")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("/api/")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("/api////")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("api")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("api/")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("/")).toBe("/api");
  });

  it("collapses multiple leading slashes for relative paths", () => {
    expect(normalizeMarketplaceBaseUrl("//api")).toBe("/api");
    expect(normalizeMarketplaceBaseUrl("///api")).toBe("/api");
  });

  it("supports absolute URLs (and treats bare origins as marketplace hosts)", () => {
    expect(normalizeMarketplaceBaseUrl("https://marketplace.formula.app")).toBe("https://marketplace.formula.app/api");
    expect(normalizeMarketplaceBaseUrl("https://marketplace.formula.app/")).toBe("https://marketplace.formula.app/api");
    expect(normalizeMarketplaceBaseUrl("https://marketplace.formula.app/api")).toBe("https://marketplace.formula.app/api");
    expect(normalizeMarketplaceBaseUrl("https://marketplace.formula.app/api/")).toBe("https://marketplace.formula.app/api");

    // Local stubs often run over http.
    expect(normalizeMarketplaceBaseUrl("http://127.0.0.1:8080")).toBe("http://127.0.0.1:8080/api");
    expect(normalizeMarketplaceBaseUrl("http://127.0.0.1:8080/api")).toBe("http://127.0.0.1:8080/api");
  });
});

describe("MarketplaceClient default baseUrl", () => {
  const original = process.env.VITE_FORMULA_MARKETPLACE_BASE_URL;

  afterEach(() => {
    if (original === undefined) {
      delete process.env.VITE_FORMULA_MARKETPLACE_BASE_URL;
    } else {
      process.env.VITE_FORMULA_MARKETPLACE_BASE_URL = original;
    }
  });

  it("uses VITE_FORMULA_MARKETPLACE_BASE_URL when available in process.env (node tooling/tests)", () => {
    process.env.VITE_FORMULA_MARKETPLACE_BASE_URL = "https://marketplace.formula.app";
    expect(new MarketplaceClient().baseUrl).toBe("https://marketplace.formula.app/api");
  });
});

describe("MarketplaceClient Tauri IPC integration", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const originalFetch = (globalThis as any).fetch;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    if (originalFetch === undefined) {
      delete (globalThis as any).fetch;
    } else {
      (globalThis as any).fetch = originalFetch;
    }
  });

  it("uses __TAURI__.core.invoke for absolute base URLs", async () => {
    const invoke = vi.fn(async () => ({ total: 0, results: [], nextCursor: null }));
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const client = new MarketplaceClient({ baseUrl: "https://example.com/api" });
    await client.search({ q: "test", limit: 10 });

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("marketplace_search", {
      baseUrl: "https://example.com/api",
      q: "test",
      category: undefined,
      tag: undefined,
      verified: undefined,
      featured: undefined,
      sort: undefined,
      limit: 10,
      offset: undefined,
      cursor: undefined,
    });
  });

  it("does not use invoke for relative base URLs (falls back to fetch)", async () => {
    const invoke = vi.fn();
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const fetch = vi.fn(async (url: string) => ({
      ok: true,
      json: async () => ({ total: 0, results: [], nextCursor: null }),
      url,
    }));
    (globalThis as any).fetch = fetch;

    const client = new MarketplaceClient({ baseUrl: "/api" });
    await client.search({ q: "hello" });

    expect(invoke).not.toHaveBeenCalled();
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(String(fetch.mock.calls[0]?.[0])).toContain("http://localhost/api/search?q=hello");
  });
});
