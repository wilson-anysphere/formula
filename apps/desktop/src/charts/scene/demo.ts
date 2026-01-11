import { path } from "./path.js";
import type { Scene } from "./types.js";

export function createDemoScene(): Scene {
  const zigzag = path()
    .moveTo(10, 60)
    .lineTo(30, 20)
    .lineTo(50, 60)
    .lineTo(70, 20)
    .lineTo(90, 60)
    .build();

  return {
    nodes: [
      {
        kind: "rect",
        x: 0,
        y: 0,
        width: 120,
        height: 80,
        fill: { color: "#ffffff" },
        stroke: { paint: { color: "#dddddd" }, width: 1 },
      },
      {
        kind: "group",
        transform: [{ kind: "translate", x: 10, y: 10 }],
        children: [
          {
            kind: "line",
            x1: 0,
            y1: 0,
            x2: 100,
            y2: 0,
            stroke: { paint: { color: "#000000" }, width: 2, dash: [4, 2] },
          },
          {
            kind: "clip",
            clip: { kind: "rect", x: 0, y: 10, width: 100, height: 60 },
            children: [
              {
                kind: "path",
                path: zigzag,
                stroke: { paint: { color: "#d0021b" }, width: 2 },
                fill: { color: "#d0021b", opacity: 0.1 },
              },
            ],
          },
          {
            kind: "text",
            x: 50,
            y: 75,
            text: "Demo",
            align: "center",
            baseline: "alphabetic",
            font: { family: "sans-serif", sizePx: 12, weight: "bold" },
            fill: { color: "#333333" },
          },
        ],
      },
    ],
  };
}

