/**
 * @typedef {import("./engine.js").TextLayout} TextLayout
 */

/**
 * Draw a previously computed layout at a given origin.
 *
 * Consumers are responsible for setting `ctx.fillStyle` etc. If you want per-run styling,
 * iterate `layout.lines[].runs` and draw each fragment manually.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {TextLayout} layout
 * @param {number} x
 * @param {number} y
 * @param {{ rotationRad?: number }} [opts]
 */
export function drawTextLayout(ctx, layout, x, y, opts = {}) {
  const rotation = opts.rotationRad ?? 0;

  ctx.save();
  if (rotation) {
    ctx.translate(x, y);
    ctx.rotate(rotation);
    ctx.translate(-x, -y);
  }

  for (let i = 0; i < layout.lines.length; i++) {
    const line = layout.lines[i];
    const baselineY = y + i * layout.lineHeight + line.ascent;
    ctx.fillText(line.text, x + line.x, baselineY);
  }

  ctx.restore();
}

