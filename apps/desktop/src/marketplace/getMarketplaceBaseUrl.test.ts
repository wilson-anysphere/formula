import { describe, expect, test } from "vitest";

import { getMarketplaceBaseUrl } from "./getMarketplaceBaseUrl.js";

describe("getMarketplaceBaseUrl", () => {
  test("prefers localStorage override", () => {
    const storage = {
      getItem(key: string) {
        if (key === "formula:marketplace:baseUrl") return " https://example.com/api/ ";
        return null;
      },
    };

    expect(getMarketplaceBaseUrl({ storage, env: { PROD: true, VITE_FORMULA_MARKETPLACE_BASE_URL: "https://env.example/api" } })).toBe(
      "https://example.com/api",
    );
  });

  test("falls back to VITE_FORMULA_MARKETPLACE_BASE_URL", () => {
    const storage = {
      getItem() {
        return null;
      },
    };

    expect(getMarketplaceBaseUrl({ storage, env: { VITE_FORMULA_MARKETPLACE_BASE_URL: "https://env.example/api/" } })).toBe(
      "https://env.example/api",
    );
  });

  test("defaults to /api in dev/test", () => {
    expect(getMarketplaceBaseUrl({ env: { DEV: true, PROD: false } })).toBe("/api");
  });

  test("defaults to hosted marketplace in production builds", () => {
    expect(getMarketplaceBaseUrl({ env: { PROD: true } })).toBe("https://marketplace.formula.app/api");
  });
});

