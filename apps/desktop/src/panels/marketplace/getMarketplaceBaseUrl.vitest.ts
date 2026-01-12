import { describe, expect, it } from "vitest";

import { getMarketplaceBaseUrl } from "./getMarketplaceBaseUrl.js";

describe("getMarketplaceBaseUrl", () => {
  it("prefers localStorage override", () => {
    const storage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return " https://example.com/api/ ";
        return null;
      },
    };

    expect(
      getMarketplaceBaseUrl({
        storage,
        env: { PROD: true, VITE_FORMULA_MARKETPLACE_BASE_URL: "https://env.example/api" },
      }),
    ).toBe("https://example.com/api");
  });

  it("normalizes origin overrides to /api", () => {
    const storage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return "https://example.com";
        return null;
      },
    };

    expect(getMarketplaceBaseUrl({ storage, env: { PROD: true } })).toBe("https://example.com/api");
  });

  it("strips query/hash from overrides", () => {
    const storage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return " https://example.com/api?x=y#z ";
        return null;
      },
    };
    expect(getMarketplaceBaseUrl({ storage, env: { PROD: true } })).toBe("https://example.com/api");

    const relativeStorage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return " /api?x=y#z ";
        return null;
      },
    };
    expect(getMarketplaceBaseUrl({ storage: relativeStorage, env: { DEV: true } })).toBe("/api");
  });

  it("collapses multiple leading slashes for relative overrides", () => {
    const storage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return "//api";
        return null;
      },
    };
    expect(getMarketplaceBaseUrl({ storage, env: { DEV: true } })).toBe("/api");
  });

  it("ignores invalid absolute URL overrides and falls back to defaults", () => {
    const storage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return "https://";
        return null;
      },
    };

    expect(getMarketplaceBaseUrl({ storage, env: { PROD: true } })).toBe("https://marketplace.formula.app/api");
  });

  it("falls back to VITE_FORMULA_MARKETPLACE_BASE_URL", () => {
    const storage = {
      getItem() {
        return null;
      },
    };

    expect(getMarketplaceBaseUrl({ storage, env: { VITE_FORMULA_MARKETPLACE_BASE_URL: "https://env.example/api/" } })).toBe(
      "https://env.example/api",
    );
  });

  it("ignores invalid VITE_FORMULA_MARKETPLACE_BASE_URL absolute URLs", () => {
    const storage = {
      getItem() {
        return null;
      },
    };
    expect(getMarketplaceBaseUrl({ storage, env: { PROD: true, VITE_FORMULA_MARKETPLACE_BASE_URL: "https://" } })).toBe(
      "https://marketplace.formula.app/api",
    );
  });

  it("defaults to /api in dev/test", () => {
    expect(getMarketplaceBaseUrl({ env: { DEV: true, PROD: false } })).toBe("/api");
  });

  it("defaults to hosted marketplace in production builds", () => {
    expect(getMarketplaceBaseUrl({ env: { PROD: true } })).toBe("https://marketplace.formula.app/api");
  });
});
