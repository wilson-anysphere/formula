import { argbToCss } from "./colors.js";

/**
 * Draw conditional formatting for a set of cell rects.
 *
 * This module is intentionally decoupled from the spreadsheet model. The
 * expectation is that higher layers provide already-evaluated per-cell CF
 * results for the visible range.
 *
 * @param {CanvasRenderingContext2D|any} ctx Canvas 2D context (or recording ctx in tests)
 * @param {Array<{a1:string,x:number,y:number,width:number,height:number}>} cellRects
 * @param {Record<string, {style?:{fill?:string,font_color?:string}, data_bar?:{color:string,fill_ratio:number}, icon?:{set:string,index:number}}>} byCell
 */
export function drawConditionalFormattingLayer(ctx, cellRects, byCell) {
  for (const rect of cellRects) {
    const cf = byCell[rect.a1];
    if (!cf) continue;

    // Background fill / color scales / dxf fills.
    if (cf.style && cf.style.fill) {
      ctx.fillStyle = argbToCss(cf.style.fill);
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }

    // Data bars.
    if (cf.data_bar) {
      const ratio = Math.max(0, Math.min(1, cf.data_bar.fill_ratio ?? 0));
      ctx.fillStyle = argbToCss(cf.data_bar.color);
      const padX = 1;
      const padY = Math.max(1, Math.floor(rect.height * 0.25));
      const h = rect.height - padY * 2;
      const w = Math.max(0, Math.floor((rect.width - padX * 2) * ratio));
      ctx.fillRect(rect.x + padX, rect.y + padY, w, h);
    }

    // Icon sets (start with 3Arrows).
    if (cf.icon && cf.icon.set === "ThreeArrows") {
      drawThreeArrowsIcon(ctx, rect, cf.icon.index);
    }
  }
}

function drawThreeArrowsIcon(ctx, rect, index) {
  const size = Math.min(rect.width, rect.height) * 0.5;
  const cx = rect.x + rect.width - size * 0.75;
  const cy = rect.y + rect.height * 0.5;

  let color = "FF00FF00"; // up = green
  if (index === 1) color = "FFFFFF00"; // mid = yellow
  if (index === 0) color = "FFFF0000"; // down = red
  ctx.fillStyle = argbToCss(color);

  const half = size / 2;
  ctx.beginPath();
  if (index === 2) {
    // Up triangle
    ctx.moveTo(cx, cy - half);
    ctx.lineTo(cx - half, cy + half);
    ctx.lineTo(cx + half, cy + half);
  } else if (index === 1) {
    // Right triangle
    ctx.moveTo(cx + half, cy);
    ctx.lineTo(cx - half, cy - half);
    ctx.lineTo(cx - half, cy + half);
  } else {
    // Down triangle
    ctx.moveTo(cx, cy + half);
    ctx.lineTo(cx - half, cy - half);
    ctx.lineTo(cx + half, cy - half);
  }
  ctx.closePath();
  ctx.fill();
}
