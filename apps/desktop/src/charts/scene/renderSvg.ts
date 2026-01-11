import { paintToRgba, rgbaToHex } from "./color.js";
import { formatNumber } from "./format.js";
import { pathToSvgD } from "./path.js";
import { measureTextWidth } from "./text.js";
import { transformToSvg } from "./transform.js";
import type {
  ClipNode,
  ClipShape,
  CircleNode,
  FillRule,
  FontSpec,
  GroupNode,
  LineNode,
  Node,
  Paint,
  PathNode,
  PolylineNode,
  RectNode,
  Scene,
  Stroke,
  TextBaseline,
  TextNode,
} from "./types.js";

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function paintAttrs(paint: Paint | undefined, attr: "fill" | "stroke"): string[] {
  if (!paint) {
    if (attr === "fill") return ['fill="none"'];
    return [];
  }

  const rgba = paintToRgba(paint);
  if (!rgba) {
    const attrs: string[] = [];
    attrs.push(`${attr}="${escapeXml(paint.color)}"`);
    if (paint.opacity != null) {
      const opacity = Math.max(0, Math.min(1, paint.opacity));
      attrs.push(`${attr}-opacity="${formatNumber(opacity)}"`);
    }
    return attrs;
  }

  const attrs: string[] = [];
  attrs.push(`${attr}="${rgbaToHex(rgba)}"`);
  if (rgba.a < 1) attrs.push(`${attr}-opacity="${formatNumber(rgba.a)}"`);
  return attrs;
}

function strokeAttrs(stroke: Stroke | undefined): string[] {
  if (!stroke) return [];
  const attrs: string[] = [];
  attrs.push(...paintAttrs(stroke.paint, "stroke"));
  attrs.push(`stroke-width="${formatNumber(stroke.width)}"`);
  if (stroke.dash?.length) attrs.push(`stroke-dasharray="${stroke.dash.map(formatNumber).join(" ")}"`);
  return attrs;
}

function fontAttrs(font: FontSpec): string[] {
  const attrs: string[] = [];
  attrs.push(`font-family="${escapeXml(font.family)}"`);
  attrs.push(`font-size="${formatNumber(font.sizePx)}"`);
  if (font.weight != null) attrs.push(`font-weight="${escapeXml(String(font.weight))}"`);
  if (font.style != null) attrs.push(`font-style="${escapeXml(font.style)}"`);
  return attrs;
}

function textAlignToSvgAnchor(align: TextNode["align"]): "start" | "middle" | "end" | null {
  switch (align) {
    case "center":
      return "middle";
    case "right":
    case "end":
      return "end";
    case "left":
    case "start":
      return "start";
    default:
      return null;
  }
}

function baselineToDominant(baseline: TextBaseline | undefined): string | null {
  switch (baseline) {
    case "top":
      return "text-before-edge";
    case "hanging":
      return "hanging";
    case "middle":
      return "middle";
    case "bottom":
      return "text-after-edge";
    case "alphabetic":
    case "ideographic":
    default:
      return null;
  }
}

function fillRuleAttr(rule: FillRule | undefined): string[] {
  if (rule === "evenodd") return ['fill-rule="evenodd"'];
  return [];
}

function renderClipShape(shape: ClipShape): string {
  if (shape.kind === "rect") {
    const attrs = [
      `x="${formatNumber(shape.x)}"`,
      `y="${formatNumber(shape.y)}"`,
      `width="${formatNumber(shape.width)}"`,
      `height="${formatNumber(shape.height)}"`,
    ];
    if (shape.rx != null) attrs.push(`rx="${formatNumber(shape.rx)}"`);
    if (shape.ry != null) attrs.push(`ry="${formatNumber(shape.ry)}"`);
    const t = transformToSvg(shape.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<rect ${attrs.join(" ")} />`;
  }

  const attrs = [`d="${pathToSvgD(shape.path)}"`];
  attrs.push(...fillRuleAttr(shape.fillRule));
  const t = transformToSvg(shape.transform);
  if (t) attrs.push(`transform="${t}"`);
  return `<path ${attrs.join(" ")} />`;
}

export function renderSceneToSvg(scene: Scene, options: { width: number; height: number }): string {
  const defs: string[] = [];
  let clipCounter = 0;

  const renderNode = (node: Node): string => {
    switch (node.kind) {
      case "rect":
        return renderRect(node);
      case "line":
        return renderLine(node);
      case "polyline":
        return renderPolyline(node);
      case "circle":
        return renderCircle(node);
      case "path":
        return renderPath(node);
      case "text":
        return renderText(node);
      case "group":
        return renderGroup(node);
      case "clip":
        return renderClip(node);
    }
  };

  const renderRect = (node: RectNode): string => {
    const attrs: string[] = [
      `x="${formatNumber(node.x)}"`,
      `y="${formatNumber(node.y)}"`,
      `width="${formatNumber(node.width)}"`,
      `height="${formatNumber(node.height)}"`,
    ];
    if (node.rx != null) attrs.push(`rx="${formatNumber(node.rx)}"`);
    if (node.ry != null) attrs.push(`ry="${formatNumber(node.ry)}"`);
    attrs.push(...paintAttrs(node.fill, "fill"));
    attrs.push(...strokeAttrs(node.stroke));
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<rect ${attrs.join(" ")} />`;
  };

  const renderLine = (node: LineNode): string => {
    const attrs: string[] = [
      `x1="${formatNumber(node.x1)}"`,
      `y1="${formatNumber(node.y1)}"`,
      `x2="${formatNumber(node.x2)}"`,
      `y2="${formatNumber(node.y2)}"`,
    ];
    attrs.push(...strokeAttrs(node.stroke));
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<line ${attrs.join(" ")} />`;
  };

  const renderPolyline = (node: PolylineNode): string => {
    const attrs: string[] = [];
    const points = node.points
      .map((p) => `${formatNumber(p.x)},${formatNumber(p.y)}`)
      .join(" ");
    attrs.push(`points="${points}"`);
    attrs.push(...paintAttrs(node.fill, "fill"));
    attrs.push(...strokeAttrs(node.stroke));
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<polyline ${attrs.join(" ")} />`;
  };

  const renderCircle = (node: CircleNode): string => {
    const attrs: string[] = [
      `cx="${formatNumber(node.cx)}"`,
      `cy="${formatNumber(node.cy)}"`,
      `r="${formatNumber(node.r)}"`,
    ];
    attrs.push(...paintAttrs(node.fill, "fill"));
    attrs.push(...strokeAttrs(node.stroke));
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<circle ${attrs.join(" ")} />`;
  };

  const renderPath = (node: PathNode): string => {
    const attrs: string[] = [`d="${pathToSvgD(node.path)}"`];
    attrs.push(...paintAttrs(node.fill, "fill"));
    attrs.push(...strokeAttrs(node.stroke));
    attrs.push(...fillRuleAttr(node.fillRule));
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<path ${attrs.join(" ")} />`;
  };

  const renderText = (node: TextNode): string => {
    const attrs: string[] = [`x="${formatNumber(node.x)}"`, `y="${formatNumber(node.y)}"`];
    const anchor = textAlignToSvgAnchor(node.align);
    if (anchor) attrs.push(`text-anchor="${anchor}"`);
    const dominant = baselineToDominant(node.baseline);
    if (dominant) attrs.push(`dominant-baseline="${dominant}"`);
    attrs.push(...fontAttrs(node.font));
    attrs.push(...paintAttrs(node.fill, "fill"));

    if (node.maxWidth != null && node.maxWidth > 0) {
      const measured = measureTextWidth(node.text, node.font);
      if (measured > node.maxWidth + 0.01) {
        attrs.push(`textLength="${formatNumber(node.maxWidth)}"`, 'lengthAdjust="spacingAndGlyphs"');
      }
    }

    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    return `<text ${attrs.join(" ")}>${escapeXml(node.text)}</text>`;
  };

  const renderGroup = (node: GroupNode): string => {
    const attrs: string[] = [];
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    const inner = node.children.map(renderNode).join("");
    if (attrs.length) return `<g ${attrs.join(" ")}>${inner}</g>`;
    return `<g>${inner}</g>`;
  };

  const renderClip = (node: ClipNode): string => {
    const id = `clip${clipCounter++}`;
    defs.push(`<clipPath id="${id}" clipPathUnits="userSpaceOnUse">${renderClipShape(node.clip)}</clipPath>`);

    const attrs: string[] = [`clip-path="url(#${id})"`];
    const t = transformToSvg(node.transform);
    if (t) attrs.push(`transform="${t}"`);
    const inner = node.children.map(renderNode).join("");
    return `<g ${attrs.join(" ")}>${inner}</g>`;
  };

  const body = scene.nodes.map(renderNode).join("");
  const defsBlock = defs.length ? `<defs>${defs.join("")}</defs>` : "";

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${formatNumber(options.width)}" height="${formatNumber(options.height)}" viewBox="0 0 ${formatNumber(
      options.width
    )} ${formatNumber(options.height)}">`,
    defsBlock,
    body,
    `</svg>`,
  ].join("");
}
