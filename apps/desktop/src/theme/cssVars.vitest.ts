import { describe, expect, it } from "vitest";

import { resolveCssVar } from "./cssVars.js";

describe("theme/cssVars.resolveCssVar", () => {
  it("returns the explicit fallback when there is no DOM / getComputedStyle", () => {
    expect(resolveCssVar("--missing", { root: null, fallback: "hotpink" })).toBe("hotpink");
  });

  it("reads literal values from computed style", () => {
    const root = {};
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: (name: string) => (name === "--color" ? "#123456" : ""),
      });
      expect(resolveCssVar("--color", { root: root as any, fallback: "black" })).toBe("#123456");
    } finally {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    }
  });

  it("resolves var(--token) indirections", () => {
    const root = {};
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      const vars: Record<string, string> = {
        "--a": "var(--b)",
        "--b": "rgb(1, 2, 3)",
      };
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: (name: string) => vars[name] ?? "",
      });
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe("rgb(1, 2, 3)");
    } finally {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    }
  });

  it("resolves var(--token, fallback) when the referenced token is missing", () => {
    const root = {};
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      const vars: Record<string, string> = {
        "--a": "var(--missing, rgb(4, 5, 6))",
      };
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: (name: string) => vars[name] ?? "",
      });
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe("rgb(4, 5, 6)");
    } finally {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    }
  });

  it("resolves nested fallbacks like var(--missing, var(--other))", () => {
    const root = {};
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      const vars: Record<string, string> = {
        "--a": "var(--missing, var(--b))",
        "--b": "blue",
      };
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: (name: string) => vars[name] ?? "",
      });
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe("blue");
    } finally {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    }
  });

  it("falls back to the callsite fallback when the token is undefined", () => {
    const root = {};
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: () => "",
      });
      expect(resolveCssVar("--missing", { root: root as any, fallback: "rebeccapurple" })).toBe("rebeccapurple");
    } finally {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    }
  });

  it("breaks cycles using the nearest var() fallback", () => {
    const root = {};
    const originalGetComputedStyle = (globalThis as any).getComputedStyle;

    try {
      const vars: Record<string, string> = {
        "--a": "var(--b)",
        "--b": "var(--a, red)",
      };
      (globalThis as any).getComputedStyle = () => ({
        getPropertyValue: (name: string) => vars[name] ?? "",
      });
      expect(resolveCssVar("--a", { root: root as any, fallback: "black" })).toBe("red");
    } finally {
      (globalThis as any).getComputedStyle = originalGetComputedStyle;
    }
  });
});

