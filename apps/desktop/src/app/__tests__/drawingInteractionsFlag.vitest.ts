import { describe, expect, it } from "vitest";

import { resolveEnableDrawingInteractions } from "../../drawings/drawingInteractionsFlag";

describe("resolveEnableDrawingInteractions", () => {
  it("defaults to false when there are no query/env overrides", () => {
    expect(resolveEnableDrawingInteractions("", null)).toBe(false);
  });

  it("honors env overrides", () => {
    expect(resolveEnableDrawingInteractions("", "1")).toBe(true);
    expect(resolveEnableDrawingInteractions("", "true")).toBe(true);
    expect(resolveEnableDrawingInteractions("", true)).toBe(true);

    expect(resolveEnableDrawingInteractions("", "0")).toBe(false);
    expect(resolveEnableDrawingInteractions("", "false")).toBe(false);
    expect(resolveEnableDrawingInteractions("", false)).toBe(false);
  });

  it("honors query string overrides over env", () => {
    expect(resolveEnableDrawingInteractions("?drawingInteractions=1", "0")).toBe(true);
    expect(resolveEnableDrawingInteractions("?drawingInteractions=false", "1")).toBe(false);
  });

  it("accepts query param aliases", () => {
    expect(resolveEnableDrawingInteractions("?drawings=1", null)).toBe(true);
    expect(resolveEnableDrawingInteractions("?enableDrawingInteractions=true", null)).toBe(true);
  });
});

