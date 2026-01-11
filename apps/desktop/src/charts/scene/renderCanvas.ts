import { paintToRgba, rgbaToCss } from "./color.js";
import { applyPathToCanvas } from "./path.js";
import { applyInverseTransformToCanvas, applyTransformToCanvas } from "./transform.js";
import type { CircleNode, ClipShape, Node, Paint, PathNode, PolylineNode, RectNode, Scene, Stroke, TextNode } from "./types.js";
import { fontSpecToCss } from "./text.js";

function applyFill(ctx: CanvasRenderingContext2D, paint: Paint | undefined): boolean {
  if (!paint) return false;
  const rgba = paintToRgba(paint);
  if (!rgba || rgba.a <= 0) return false;
  ctx.fillStyle = rgbaToCss(rgba);
  return true;
}

function applyStroke(ctx: CanvasRenderingContext2D, stroke: Stroke | undefined): boolean {
  if (!stroke) return false;
  const rgba = paintToRgba(stroke.paint);
  if (!rgba || rgba.a <= 0) return false;
  ctx.strokeStyle = rgbaToCss(rgba);
  ctx.lineWidth = stroke.width;
  ctx.setLineDash(stroke.dash ?? []);
  return true;
}

function drawRoundedRect(ctx: CanvasRenderingContext2D, node: Pick<RectNode, "x" | "y" | "width" | "height" | "rx" | "ry">): void {
  const { x, y, width, height } = node;
  const rxRaw = node.rx ?? node.ry ?? 0;
  const ryRaw = node.ry ?? node.rx ?? 0;
  const rx = Math.max(0, Math.min(rxRaw, width / 2));
  const ry = Math.max(0, Math.min(ryRaw, height / 2));

  if (!rx || !ry) {
    ctx.rect(x, y, width, height);
    return;
  }

  // Use `ellipse` so rx/ry can differ (Canvas `arcTo` only supports circular radii).
  ctx.moveTo(x + rx, y);
  ctx.lineTo(x + width - rx, y);
  ctx.ellipse(x + width - rx, y + ry, rx, ry, 0, -Math.PI / 2, 0);
  ctx.lineTo(x + width, y + height - ry);
  ctx.ellipse(x + width - rx, y + height - ry, rx, ry, 0, 0, Math.PI / 2);
  ctx.lineTo(x + rx, y + height);
  ctx.ellipse(x + rx, y + height - ry, rx, ry, 0, Math.PI / 2, Math.PI);
  ctx.lineTo(x, y + ry);
  ctx.ellipse(x + rx, y + ry, rx, ry, 0, Math.PI, (Math.PI * 3) / 2);
  ctx.closePath();
}

function drawClipShape(ctx: CanvasRenderingContext2D, shape: ClipShape): void {
  if (shape.kind === "rect") {
    drawRoundedRect(ctx, shape);
    return;
  }
  applyPathToCanvas(ctx, shape.path);
}

export function renderSceneToCanvas(scene: Scene, ctx: CanvasRenderingContext2D): void {
  const renderNode = (node: Node): void => {
    switch (node.kind) {
      case "rect":
        renderRect(node);
        break;
      case "line":
        renderLine(node);
        break;
      case "polyline":
        renderPolyline(node);
        break;
      case "circle":
        renderCircle(node);
        break;
      case "path":
        renderPath(node);
        break;
      case "text":
        renderText(node);
        break;
      case "group":
        renderGroup(node);
        break;
      case "clip":
        renderClip(node);
        break;
    }
  };

  const renderRect = (node: RectNode): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    ctx.beginPath();
    drawRoundedRect(ctx, node);
    if (applyFill(ctx, node.fill)) ctx.fill();
    if (applyStroke(ctx, node.stroke)) ctx.stroke();
    ctx.restore();
  };

  const renderLine = (node: Extract<Node, { kind: "line" }>): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    ctx.beginPath();
    ctx.moveTo(node.x1, node.y1);
    ctx.lineTo(node.x2, node.y2);
    if (applyStroke(ctx, node.stroke)) ctx.stroke();
    ctx.restore();
  };

  const renderPolyline = (node: PolylineNode): void => {
    if (node.points.length === 0) return;
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    ctx.beginPath();
    ctx.moveTo(node.points[0].x, node.points[0].y);
    for (let i = 1; i < node.points.length; i += 1) {
      ctx.lineTo(node.points[i].x, node.points[i].y);
    }
    if (applyFill(ctx, node.fill)) ctx.fill();
    if (applyStroke(ctx, node.stroke)) ctx.stroke();
    ctx.restore();
  };

  const renderCircle = (node: CircleNode): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    ctx.beginPath();
    ctx.arc(node.cx, node.cy, node.r, 0, Math.PI * 2);
    if (applyFill(ctx, node.fill)) ctx.fill();
    if (applyStroke(ctx, node.stroke)) ctx.stroke();
    ctx.restore();
  };

  const renderPath = (node: PathNode): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    ctx.beginPath();
    applyPathToCanvas(ctx, node.path);
    if (applyFill(ctx, node.fill)) ctx.fill(node.fillRule);
    if (applyStroke(ctx, node.stroke)) ctx.stroke();
    ctx.restore();
  };

  const renderText = (node: TextNode): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    ctx.font = fontSpecToCss(node.font);
    if (node.align) ctx.textAlign = node.align === "center" ? "center" : node.align;
    if (node.baseline) ctx.textBaseline = node.baseline;
    if (applyFill(ctx, node.fill)) {
      if (node.maxWidth != null) ctx.fillText(node.text, node.x, node.y, node.maxWidth);
      else ctx.fillText(node.text, node.x, node.y);
    }
    ctx.restore();
  };

  const renderGroup = (node: Extract<Node, { kind: "group" }>): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    for (const child of node.children) renderNode(child);
    ctx.restore();
  };

  const renderClip = (node: Extract<Node, { kind: "clip" }>): void => {
    ctx.save();
    applyTransformToCanvas(ctx, node.transform);
    applyTransformToCanvas(ctx, node.clip.transform);
    ctx.beginPath();
    drawClipShape(ctx, node.clip);
    ctx.clip(node.clip.kind === "path" ? node.clip.fillRule : undefined);
    applyInverseTransformToCanvas(ctx, node.clip.transform);
    for (const child of node.children) renderNode(child);
    ctx.restore();
  };

  for (const node of scene.nodes) renderNode(node);
}
