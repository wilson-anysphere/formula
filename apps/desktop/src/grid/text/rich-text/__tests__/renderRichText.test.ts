import { beforeEach, describe, expect, it, vi } from "vitest";

const measure = vi.fn((text: string) => ({
  width: text.length * 10,
  ascent: 10,
  descent: 4,
}));

const layout = vi.fn(() => ({
  width: 0,
  height: 0,
  lineHeight: 0,
  lines: [],
}));

vi.mock("../../textLayout.js", () => ({
  getSharedTextLayoutEngine: () => ({
    measure,
    layout,
  }),
}));

function createMockCtx() {
  const fillTextCalls: Array<{ text: string; x: number; y: number }> = [];

  const ctx = new Proxy(
    {
      save: () => {},
      restore: () => {},
      beginPath: () => {},
      rect: () => {},
      clip: () => {},
      fillText: (text: string, x: number, y: number) => {
        fillTextCalls.push({ text, x, y });
      },
      moveTo: () => {},
      lineTo: () => {},
      stroke: () => {},
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return () => {};
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );

  return { ctx: ctx as unknown as CanvasRenderingContext2D, fillTextCalls };
}

describe("renderRichText", () => {
  beforeEach(() => {
    measure.mockClear();
    layout.mockClear();
  });

  it("renders plain text without allocating run slices", async () => {
    const { renderRichText } = await import("../render.js");
    const { ctx, fillTextCalls } = createMockCtx();

    renderRichText(
      ctx,
      { text: "A😀B" },
      { x: 0, y: 0, width: 200, height: 30 },
      { wrapMode: "none", align: "left", verticalAlign: "top", direction: "ltr", color: "#000" },
    );

    expect(fillTextCalls.map((c) => c.text)).toEqual(["A😀B"]);
  });

  it("treats an empty runs array as a single full-length run", async () => {
    const { renderRichText } = await import("../render.js");
    const { ctx, fillTextCalls } = createMockCtx();

    renderRichText(
      ctx,
      { text: "Hello", runs: [] },
      { x: 0, y: 0, width: 200, height: 30 },
      { wrapMode: "none", align: "left", verticalAlign: "top", direction: "ltr", color: "#000" },
    );

    expect(fillTextCalls.map((c) => c.text)).toEqual(["Hello"]);
  });

  it("splits styled runs by code point indices (emoji-safe)", async () => {
    const { renderRichText } = await import("../render.js");
    const { ctx, fillTextCalls } = createMockCtx();

    renderRichText(
      ctx,
      {
        text: "A😀B",
        runs: [
          { start: 0, end: 2, style: {} }, // "A😀"
          { start: 2, end: 3, style: {} }, // "B"
        ],
      },
      { x: 0, y: 0, width: 200, height: 30 },
      { wrapMode: "none", align: "left", verticalAlign: "top", direction: "ltr", color: "#000" },
    );

    expect(fillTextCalls.map((c) => c.text)).toEqual(["A😀", "B"]);
  });
});

