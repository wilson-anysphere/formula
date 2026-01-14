import { describe, expect, it } from "vitest";

import { emuToPx, pxToEmu } from "../overlay";

describe("drawings EMU ↔ px conversions", () => {
  it("converts EMU to px using zoom", () => {
    const emu = pxToEmu(10);
    expect(emuToPx(emu, 2)).toBeCloseTo(20, 6);
  });

  it("converts px to EMU using zoom", () => {
    expect(pxToEmu(20, 2)).toBeCloseTo(pxToEmu(10), 6);
  });

  it("treats invalid zoom values as zoom=1", () => {
    const emu = pxToEmu(10);
    expect(emuToPx(emu, 0)).toBeCloseTo(10, 6);
    expect(emuToPx(emu, Number.NaN)).toBeCloseTo(10, 6);
    expect(pxToEmu(10, 0)).toBeCloseTo(pxToEmu(10), 6);
  });

  it("rounds half values away from zero so pxToEmu is odd", () => {
    // `Math.round` is asymmetric for negative half values; pxToEmu normalizes so
    // opposite pixel deltas cancel exactly (important for reversible zoom nudges).
    expect(pxToEmu(-0.5)).toBe(-pxToEmu(0.5));
    expect(pxToEmu(-1, 2)).toBe(-pxToEmu(1, 2));
  });
});
