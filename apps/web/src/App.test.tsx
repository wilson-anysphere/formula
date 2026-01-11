// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@formula/text-layout", () => ({
  TextLayoutEngine: class TextLayoutEngine {
    measure(text: string, font: { sizePx: number }) {
      const size = font?.sizePx ?? 12;
      return { width: text.length * (size * 0.6), ascent: size * 0.8, descent: size * 0.2 };
    }

    layout() {
      return { lines: [], width: 0, height: 0, lineHeight: 0, direction: "ltr", maxWidth: 0, resolvedAlign: "left" };
    }
  },
  createCanvasTextMeasurer: () => ({ measure: () => ({ width: 0, ascent: 0, descent: 0 }) }),
  detectBaseDirection: () => "ltr",
  resolveAlign: (align: string, direction: "ltr" | "rtl") => {
    if (align === "start") return direction === "rtl" ? "right" : "left";
    if (align === "end") return direction === "rtl" ? "left" : "right";
    return align;
  },
  drawTextLayout: () => {},
  toCanvasFontString: (font: { family: string; sizePx: number; weight?: string }) =>
    `${font.weight ?? "400"} ${font.sizePx}px ${font.family}`
}));

const mocks = vi.hoisted(() => {
  const seeded = new Map<string, string | number | boolean>([
    ["A1", true],
    ["A2", 2],
    ["B1", 3],
    ["B2", 6],
    ["C1", "hello"]
  ]);

  const colToLetters = (col: number) => {
    let n = col;
    let out = "";
    while (n >= 0) {
      out = String.fromCharCode(65 + (n % 26)) + out;
      n = Math.floor(n / 26) - 1;
    }
    return out;
  };

  const parseA1 = (addr: string) => {
    const match = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!match) throw new Error(`invalid A1 address: ${addr}`);
    const [, letters, digits] = match;
    let col = 0;
    for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64);
    col -= 1;
    const row = Number.parseInt(digits, 10) - 1;
    return { row, col };
  };

  const parseRange = (range: string) => {
    const [start, end = start] = range.split(":");
    const a = parseA1(start);
    const b = parseA1(end);
    return {
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col),
      endCol: Math.max(a.col, b.col)
    };
  };

  const engine = {
    init: vi.fn(async () => {}),
    terminate: vi.fn(),
    newWorkbook: vi.fn(async () => {}),
    loadWorkbookFromJson: vi.fn(async () => {}),
    toJson: vi.fn(async () => "{}"),
    getCell: vi.fn(async (address: string) => ({
      sheet: "Sheet1",
      address,
      input: null,
      value: address === "B1" ? 3 : null
    })),
    getRange: vi.fn(async (range: string) => {
      const bounds = parseRange(range);
      const rows: Array<Array<{ value: string | number | boolean | null }>> = [];
      for (let r = bounds.startRow; r <= bounds.endRow; r++) {
        const cols: Array<{ value: string | number | boolean | null }> = [];
        for (let c = bounds.startCol; c <= bounds.endCol; c++) {
          const addr = `${colToLetters(c)}${r + 1}`;
          cols.push({ value: seeded.get(addr) ?? null });
        }
        rows.push(cols);
      }
      return rows;
    }),
    setCell: vi.fn(async () => {}),
    setRange: vi.fn(async () => {}),
    recalculate: vi.fn(async () => [])
  };

  return { engine };
});

vi.mock("@formula/engine", () => ({
  createEngineClient: () => mocks.engine
}));

function createMock2dContext(
  canvas: HTMLCanvasElement,
  drawnText: Array<{ text: string; x: number; y: number }>
): CanvasRenderingContext2D {
  const noop = () => {};
  return {
    canvas,
    fillStyle: "#000",
    strokeStyle: "#000",
    lineWidth: 1,
    font: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: noop,
    fillRect: noop,
    strokeRect: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    fill: noop,
    stroke: noop,
    moveTo: noop,
    lineTo: noop,
    closePath: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText: (text: string, x: number, y: number) => {
      drawnText.push({ text: String(text), x, y });
    },
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

async function flushMicrotasks(count = 5): Promise<void> {
  for (let i = 0; i < count; i++) {
    // `Promise.resolve()` is enough: the web preview boot sequence + provider prefetch
    // uses microtasks only (async/await).
    await Promise.resolve();
  }
}

describe("App (web preview)", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  let rectSpy: ReturnType<typeof vi.spyOn> | null = null;

  beforeEach(() => {
    // React 18 relies on this flag to suppress act() warnings in test runners.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

    rectSpy = vi
      .spyOn(Element.prototype, "getBoundingClientRect")
      .mockReturnValue({
        left: 0,
        top: 0,
        right: 300,
        bottom: 63,
        width: 300,
        height: 63,
        x: 0,
        y: 0,
        toJSON: () => {}
      } as DOMRect);
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    rectSpy?.mockRestore();
    rectSpy = null;
    vi.unstubAllGlobals();
    vi.clearAllMocks();
  });

  it("renders engine-backed values (B1=3) into the Canvas grid", async () => {
    const drawnText: Array<{ text: string; x: number; y: number }> = [];
    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this, drawnText);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const { App } = await import("./App");

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(App));
    });

    await act(async () => {
      await flushMicrotasks(10);
    });

    expect(host.textContent).toContain("ready (B1=3)");
    expect(mocks.engine.getRange).toHaveBeenCalled();
    expect(drawnText.some((call) => call.text === "3" && call.x > 100 && call.y > 21)).toBe(true);
    expect(drawnText.some((call) => call.text === "TRUE")).toBe(true);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
