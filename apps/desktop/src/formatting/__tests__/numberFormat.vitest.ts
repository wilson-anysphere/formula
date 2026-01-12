import { describe, expect, it } from "vitest";

import { formatValueWithNumberFormat } from "../numberFormat.js";

describe("formatValueWithNumberFormat", () => {
  it("formats time-only hh:mm:ss formats (Excel-style)", () => {
    expect(formatValueWithNumberFormat(0, "hh:mm:ss")).toBe("00:00:00");
    expect(formatValueWithNumberFormat(0.5, "hh:mm:ss")).toBe("12:00:00");
  });
});

