export function emuToPx(emu: number): number;

export function anchorToRectPx(
  anchor: any,
  opts?: {
    colWidthsPx?: number[];
    rowHeightsPx?: number[];
    defaultColWidthPx?: number;
    defaultRowHeightPx?: number;
  }
): { left: number; top: number; width: number; height: number } | null;

