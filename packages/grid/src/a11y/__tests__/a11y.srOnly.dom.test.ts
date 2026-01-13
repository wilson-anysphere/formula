// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { applySrOnlyStyle, SR_ONLY_STYLE } from "../a11y";

describe("a11y SR-only helpers", () => {
  it("applies SR-only inline styles to an element", () => {
    const div = document.createElement("div");

    applySrOnlyStyle(div);

    expect(div.style.position).toBe(SR_ONLY_STYLE.position);
    expect(div.style.width).toBe(SR_ONLY_STYLE.width);
    expect(div.style.height).toBe(SR_ONLY_STYLE.height);
    expect(div.style.padding).toBe(SR_ONLY_STYLE.padding);
    expect(div.style.margin).toBe(SR_ONLY_STYLE.margin);
    expect(div.style.overflow).toBe(SR_ONLY_STYLE.overflow);
    expect(div.style.clip).toBe(SR_ONLY_STYLE.clip);
    expect(div.style.whiteSpace).toBe(SR_ONLY_STYLE.whiteSpace);
    expect(div.style.border).toBe(SR_ONLY_STYLE.border);
  });
});

