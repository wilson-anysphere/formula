import { afterEach, describe, expect, it } from "vitest";

import {
  getExplicitPyodideIndexURL,
  isSafePyodideIndexURLForDesktopOverride,
  normalizePyodideIndexURL,
  pickPyodideIndexURL,
} from "./pyodideIndexURL.js";

describe("pyodideIndexURL helpers", () => {
  afterEach(() => {
    delete (globalThis as any).__pyodideIndexURL;
    delete (globalThis as any).__TAURI__;
  });

  it("normalizes trailing slashes", () => {
    expect(normalizePyodideIndexURL("https://cdn.example.com/pyodide")).toBe("https://cdn.example.com/pyodide/");
    expect(normalizePyodideIndexURL("https://cdn.example.com/pyodide/")).toBe("https://cdn.example.com/pyodide/");
  });

  it("treats empty or non-string values as undefined", () => {
    expect(normalizePyodideIndexURL("")).toBeUndefined();
    expect(normalizePyodideIndexURL("   ")).toBeUndefined();
    expect(normalizePyodideIndexURL(null)).toBeUndefined();
    expect(normalizePyodideIndexURL(undefined)).toBeUndefined();
    expect(normalizePyodideIndexURL(123)).toBeUndefined();
  });

  it("prefers explicit indexURL over cached", () => {
    expect(pickPyodideIndexURL({ explicitIndexURL: "explicit/", cachedIndexURL: "cached/" })).toBe("explicit/");
    expect(pickPyodideIndexURL({ cachedIndexURL: "cached/" })).toBe("cached/");
    expect(pickPyodideIndexURL({})).toBeUndefined();
  });

  it("treats only local URLs as safe desktop overrides", () => {
    expect(isSafePyodideIndexURLForDesktopOverride("pyodide://v0.25.1/full/")).toBe(true);
    expect(isSafePyodideIndexURLForDesktopOverride("/pyodide/v0.25.1/full/")).toBe(true);
    expect(isSafePyodideIndexURLForDesktopOverride("https://cdn.jsdelivr.net/pyodide/v0.25.1/full/")).toBe(false);
    expect(isSafePyodideIndexURLForDesktopOverride("")).toBe(false);
  });

  it("allows global __pyodideIndexURL overrides outside the desktop app", () => {
    (globalThis as any).__pyodideIndexURL = "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/";
    expect(getExplicitPyodideIndexURL()).toBe("https://cdn.jsdelivr.net/pyodide/v0.25.1/full/");
  });

  it("ignores non-local __pyodideIndexURL overrides inside the desktop app", () => {
    (globalThis as any).__TAURI__ = { core: { invoke: async () => null } };
    (globalThis as any).__pyodideIndexURL = "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/";
    expect(getExplicitPyodideIndexURL()).toBeUndefined();
  });

  it("accepts local __pyodideIndexURL overrides inside the desktop app", () => {
    (globalThis as any).__TAURI__ = { core: { invoke: async () => null } };
    (globalThis as any).__pyodideIndexURL = "/pyodide/v0.25.1/full/";
    expect(getExplicitPyodideIndexURL()).toBe("/pyodide/v0.25.1/full/");
  });
});
