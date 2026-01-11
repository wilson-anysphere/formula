import { paintToRgba, rgbaToCss } from "./color.js";
import { applyPathToCanvas } from "./path.js";
import { applyTransformToCanvas } from "./transform.js";
import type { ClipShape, Node, Paint, PathNode, RectNode, Scene, Stroke, TextNode } from "./types.js";
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

function drawClipShape(ctx: CanvasRenderingContext2D, shape: ClipShape): void {
  if (shape.kind === "rect") {
    ctx.rect(shape.x, shape.y, shape.width, shape.height);
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
    ctx.rect(node.x, node.y, node.width, node.height);
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
    ctx.beginPath();
    drawClipShape(ctx, node.clip);
    ctx.clip(node.clip.kind === "path" ? node.clip.fillRule : undefined);
    for (const child of node.children) renderNode(child);
    ctx.restore();
  };

  for (const node of scene.nodes) renderNode(node);
}

