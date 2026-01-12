import { afterEach, describe, expect, it } from "vitest";

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
