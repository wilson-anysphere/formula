import { describe, expect, it } from "vitest";

import { createDemoScene, renderSceneToCanvas, renderSceneToSvg } from "../scene/index.js";
import type { Scene } from "../scene/index.js";

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: string[] } {
  const calls: string[] = [];
  const ctx: any = {
    save: () => calls.push("save"),
    restore: () => calls.push("restore"),
    translate: () => calls.push("translate"),
    scale: () => calls.push("scale"),
    rotate: () => calls.push("rotate"),
    beginPath: () => calls.push("beginPath"),
    moveTo: () => calls.push("moveTo"),
    lineTo: () => calls.push("lineTo"),
    quadraticCurveTo: () => calls.push("quadraticCurveTo"),
    bezierCurveTo: () => calls.push("bezierCurveTo"),
    arcTo: () => calls.push("arcTo"),
    arc: () => calls.push("arc"),
    closePath: () => calls.push("closePath"),
    rect: () => calls.push("rect"),
    fill: () => calls.push("fill"),
    stroke: () => calls.push("stroke"),
    clip: () => calls.push("clip"),
    setLineDash: () => calls.push("setLineDash"),
    fillText: () => calls.push("fillText"),
    measureText: () => ({ width: 0 }),
  };

  return { ctx: ctx as CanvasRenderingContext2D, calls };
}

describe("charts scene graph", () => {
  it("renders expected SVG tags for core primitives", () => {
    const scene = createDemoScene();
    const svg = renderSceneToSvg(scene, { width: 120, height: 80 });

    expect(svg).toContain("<svg");
    expect(svg).toContain("<rect");
    expect(svg).toContain("<line");
    expect(svg).toContain("<path");
    expect(svg).toContain("<text");
    expect(svg).toContain("<g");
    expect(svg).toContain("<clipPath");

    expect(svg).toContain('clip-path="url(#clip0)"');
  });

  it("renders deterministically", () => {
    const scene = createDemoScene();
    const a = renderSceneToSvg(scene, { width: 120, height: 80 });
    const b = renderSceneToSvg(scene, { width: 120, height: 80 });
    expect(a).toEqual(b);
  });

  it("renders to Canvas2D without throwing", () => {
    const scene = createDemoScene();
    const { ctx, calls } = createStubCanvasContext();
    expect(() => renderSceneToCanvas(scene, ctx)).not.toThrow();
    expect(calls.length).toBeGreaterThan(0);
  });

  it("supports rounded rects + clip-shape transforms on Canvas", () => {
    const scene: Scene = {
      nodes: [
        {
          kind: "clip",
          clip: {
            kind: "rect",
            x: 0,
            y: 0,
            width: 40,
            height: 40,
            rx: 6,
            transform: [{ kind: "translate", x: 5, y: 5 }],
          },
          children: [
            {
              kind: "rect",
              x: 0,
              y: 0,
              width: 100,
              height: 100,
              rx: 10,
              fill: { color: "#00ff00" },
            },
          ],
        },
      ],
    };

    const { ctx } = createStubCanvasContext();
    expect(() => renderSceneToCanvas(scene, ctx)).not.toThrow();
  });
});
