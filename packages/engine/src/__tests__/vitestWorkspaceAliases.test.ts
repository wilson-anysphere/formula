import { describe, expect, it, vi } from "vitest";

// The `@formula/grid/node` entrypoint re-exports the renderer, which depends on `@formula/text-layout`.
// In environments where `node_modules` is cached/stale, we alias workspace entrypoints in
// `vitest.config.ts`. Mock `@formula/text-layout` so this test stays lightweight and doesn't require
// native/text shaping dependencies.
vi.mock("@formula/text-layout", () => ({
  TextLayoutEngine: class {},
  createCanvasTextMeasurer: () => null,
  detectBaseDirection: () => "ltr",
  drawTextLayout: () => {},
  resolveAlign: () => "start",
  toCanvasFontString: () => "",
}));

describe("Vitest workspace aliases", () => {
  it("can import @formula/fill-engine", async () => {
    const mod = await import("@formula/fill-engine");
    expect(typeof mod.computeFillEdits).toBe("function");
  });

  it("can import @formula/grid/node when @formula/text-layout is mocked", async () => {
    const mod = await import("@formula/grid/node");
    expect(typeof mod.DEFAULT_GRID_FONT_FAMILY).toBe("string");
    expect(typeof mod.LruCache).toBe("function");
  });
});

