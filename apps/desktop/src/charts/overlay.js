import { emuToPx } from "../shared/emu.js";

export { emuToPx };

function sumUpTo(index, sizes, defaultSize) {
  if (!Number.isFinite(index) || index <= 0) return 0;
  if (!sizes || sizes.length === 0) return index * defaultSize;

  let sum = 0;
  for (let i = 0; i < index; i += 1) {
    sum += sizes[i] ?? defaultSize;
  }
  return sum;
}

export function anchorToRectPx(anchor, opts = {}) {
  const colWidthsPx = opts.colWidthsPx;
  const rowHeightsPx = opts.rowHeightsPx;
  const defaultColWidthPx = opts.defaultColWidthPx ?? 64;
  const defaultRowHeightPx = opts.defaultRowHeightPx ?? 20;

  if (!anchor || !anchor.kind) return null;

  if (anchor.kind === "absolute") {
    const left = emuToPx(anchor.xEmu);
    const top = emuToPx(anchor.yEmu);
    const width = emuToPx(anchor.cxEmu);
    const height = emuToPx(anchor.cyEmu);
    return { left, top, width, height };
  }

  if (anchor.kind === "oneCell") {
    const left =
      sumUpTo(anchor.fromCol, colWidthsPx, defaultColWidthPx) +
      emuToPx(anchor.fromColOffEmu);
    const top =
      sumUpTo(anchor.fromRow, rowHeightsPx, defaultRowHeightPx) +
      emuToPx(anchor.fromRowOffEmu);
    const width = emuToPx(anchor.cxEmu);
    const height = emuToPx(anchor.cyEmu);
    return { left, top, width, height };
  }

  if (anchor.kind === "twoCell") {
    const left =
      sumUpTo(anchor.fromCol, colWidthsPx, defaultColWidthPx) +
      emuToPx(anchor.fromColOffEmu);
    const top =
      sumUpTo(anchor.fromRow, rowHeightsPx, defaultRowHeightPx) +
      emuToPx(anchor.fromRowOffEmu);
    const right =
      sumUpTo(anchor.toCol, colWidthsPx, defaultColWidthPx) +
      emuToPx(anchor.toColOffEmu);
    const bottom =
      sumUpTo(anchor.toRow, rowHeightsPx, defaultRowHeightPx) +
      emuToPx(anchor.toRowOffEmu);
    return { left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
  }

  return null;
}
