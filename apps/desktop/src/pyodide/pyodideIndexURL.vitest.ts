import { describe, expect, it } from "vitest";

import { normalizePyodideIndexURL, pickPyodideIndexURL } from "./pyodideIndexURL.js";

describe("pyodideIndexURL helpers", () => {
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
});

