export interface CellBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface CommentIndicatorStyle {
  color?: string;
  size?: number;
}

export function drawCommentIndicator(
  ctx: CanvasRenderingContext2D,
  bounds: CellBounds,
  style: CommentIndicatorStyle = {},
): void {
  const size = style.size ?? Math.max(6, Math.min(bounds.width, bounds.height) * 0.25);
  const color = style.color ?? "#F59E0B";

  ctx.save();
  ctx.beginPath();
  ctx.moveTo(bounds.x + bounds.width, bounds.y);
  ctx.lineTo(bounds.x + bounds.width - size, bounds.y);
  ctx.lineTo(bounds.x + bounds.width, bounds.y + size);
  ctx.closePath();
  ctx.fillStyle = color;
  ctx.fill();
  ctx.restore();
}

