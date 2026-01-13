import { describe, expect, it } from "vitest";

import { resolveCssVar } from "./cssVars.js";

function withStubbedGetComputedStyle<T>(getPropertyValue: (name: string) => string, fn: () => T): T {
  const hadGetComputedStyle = Object.prototype.hasOwnProperty.call(globalThis, "getComputedStyle");
  const originalGetComputedStyle = (globalThis as any).getComputedStyle;

  (globalThis as any).getComputedStyle = () => ({
    getPropertyValue,
  });

  try {
    return fn();
  } finally {
    if (hadGetComputedStyle) {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).getComputedStyle;
    }
  }
}

describe("theme/cssVars.resolveCssVar", () => {
  it("returns the explicit fallback when there is no DOM / getComputedStyle", () => {
    expect(resolveCssVar("--missing", { root: null, fallback: "hotpink" })).toBe("hotpink");
  });

  it("reads literal values from computed style", () => {
    const root = {};
    withStubbedGetComputedStyle((name) => (name === "--color" ? "rebeccapurple" : ""), () => {
      expect(resolveCssVar("--color", { root: root as any, fallback: "black" })).toBe("rebeccapurple");
    });
  });

  it("resolves var(--token) indirections", () => {
    const root = {};
    const vars: Record<string, string> = {
      "--a": "var(--b)",
      "--b": "rgb(var(--r), var(--g), var(--b))",
    };
    withStubbedGetComputedStyle((name) => vars[name] ?? "", () => {
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe(
        "rgb(var(--r), var(--g), var(--b))",
      );
    });
  });

  it("resolves var(--token, fallback) when the referenced token is missing", () => {
    const root = {};
    const vars: Record<string, string> = {
      "--a": "var(--missing, rgb(var(--r), var(--g), var(--b)))",
    };
    withStubbedGetComputedStyle((name) => vars[name] ?? "", () => {
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe(
        "rgb(var(--r), var(--g), var(--b))",
      );
    });
  });

  it("resolves nested fallbacks like var(--missing, var(--other))", () => {
    const root = {};
    const vars: Record<string, string> = {
      "--a": "var(--missing, var(--b))",
      "--b": "blue",
    };
    withStubbedGetComputedStyle((name) => vars[name] ?? "", () => {
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe("blue");
    });
  });

  it("falls back to the callsite fallback when the token is undefined", () => {
    const root = {};
    withStubbedGetComputedStyle(() => "", () => {
      expect(resolveCssVar("--missing", { root: root as any, fallback: "rebeccapurple" })).toBe("rebeccapurple");
    });
  });

  it("breaks cycles using the nearest var() fallback", () => {
    const root = {};
    const vars: Record<string, string> = {
      "--a": "var(--b)",
      "--b": "var(--a, red)",
    };
    withStubbedGetComputedStyle((name) => vars[name] ?? "", () => {
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe("red");
    });
  });
});
