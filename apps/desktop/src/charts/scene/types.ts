import type { PathData } from "./path.js";

export interface Scene {
  nodes: Node[];
}

export interface Paint {
  color: string;
  opacity?: number;
}

export interface Stroke {
  paint: Paint;
  width: number;
  dash?: number[];
}

export interface FontSpec {
  family: string;
  sizePx: number;
  weight?: number | "normal" | "bold";
  style?: "normal" | "italic" | "oblique";
}

export type Transform =
  | { kind: "translate"; x: number; y: number }
  | { kind: "scale"; x: number; y?: number }
  | { kind: "rotate"; radians: number; cx?: number; cy?: number };

export type FillRule = "nonzero" | "evenodd";

export type TextAlign = "left" | "right" | "center" | "start" | "end";

export type TextBaseline = "alphabetic" | "top" | "hanging" | "middle" | "ideographic" | "bottom";

export interface RectNode {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
  rx?: number;
  ry?: number;
  fill?: Paint;
  stroke?: Stroke;
  transform?: Transform[];
}

export interface LineNode {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  stroke: Stroke;
  transform?: Transform[];
}

export interface PathNode {
  kind: "path";
  path: PathData;
  fill?: Paint;
  stroke?: Stroke;
  fillRule?: FillRule;
  transform?: Transform[];
}

export interface TextNode {
  kind: "text";
  x: number;
  y: number;
  text: string;
  font: FontSpec;
  fill: Paint;
  align?: TextAlign;
  baseline?: TextBaseline;
  maxWidth?: number;
  transform?: Transform[];
}

export interface GroupNode {
  kind: "group";
  children: Node[];
  transform?: Transform[];
}

export type ClipShape = RectNode | PathNode;

export interface ClipNode {
  kind: "clip";
  clip: ClipShape;
  children: Node[];
  transform?: Transform[];
}

export type Node = RectNode | LineNode | PathNode | TextNode | GroupNode | ClipNode;

