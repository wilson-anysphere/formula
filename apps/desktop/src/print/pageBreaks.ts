import type { CellRange, ManualPageBreaks, Page, PageSetup } from "./types";

const POINTS_PER_INCH = 72;

function paperSizeInches(code: number): { w: number; h: number } {
  switch (code) {
    case 1: // Letter
      return { w: 8.5, h: 11 };
    case 9: // A4
      return { w: 8.267716535, h: 11.69291339 };
    default:
      return { w: 8.5, h: 11 };
  }
}

function normalizeRange(r: CellRange): CellRange {
  return {
    startRow: Math.min(r.startRow, r.endRow),
    endRow: Math.max(r.startRow, r.endRow),
    startCol: Math.min(r.startCol, r.endCol),
    endCol: Math.max(r.startCol, r.endCol),
  };
}

function sumSliceRange(sizes: number[], start1: number, end1: number): number {
  let sum = 0;
  for (let i = start1; i <= end1; i++) {
    sum += sizes[i - 1] ?? 0;
  }
  return sum;
}

function scaleFactor(
  printArea: CellRange,
  colWidthsPts: number[],
  rowHeightsPts: number[],
  setup: PageSetup,
): number {
  const { w: wIn, h: hIn } = paperSizeInches(setup.paperSize);
  let pageW = wIn * POINTS_PER_INCH;
  let pageH = hIn * POINTS_PER_INCH;
  if (setup.orientation === "landscape") {
    [pageW, pageH] = [pageH, pageW];
  }

  const printableW =
    pageW - (setup.margins.left + setup.margins.right) * POINTS_PER_INCH;
  const printableH =
    pageH - (setup.margins.top + setup.margins.bottom) * POINTS_PER_INCH;

  if (setup.scaling.kind === "percent") {
    return setup.scaling.percent / 100;
  }

  const norm = normalizeRange(printArea);
  const contentW = sumSliceRange(colWidthsPts, norm.startCol, norm.endCol);
  const contentH = sumSliceRange(rowHeightsPts, norm.startRow, norm.endRow);

  const widthPages = setup.scaling.widthPages;
  const heightPages = setup.scaling.heightPages;
  const scaleW =
    widthPages > 0 && contentW > 0 ? (widthPages * printableW) / contentW : null;
  const scaleH =
    heightPages > 0 && contentH > 0
      ? (heightPages * printableH) / contentH
      : null;

  if (scaleW != null && scaleH != null) return Math.min(scaleW, scaleH);
  if (scaleW != null) return scaleW;
  if (scaleH != null) return scaleH;
  return 1;
}

function breakStarts(
  start: number,
  end: number,
  sizes: number[],
  capacity: number,
  manualStarts: number[],
): number[] {
  const starts: number[] = [start];
  let current = start;

  while (current <= end) {
    let acc = 0;
    let next = current;
    while (next <= end) {
      const size = sizes[next - 1] ?? 0;
      if (next > current && acc + size > capacity) break;
      acc += size;
      next++;
    }
    if (next === current) next++;
    current = next;
    if (current <= end) starts.push(current);
  }

  for (const ms of manualStarts) {
    if (ms > start && ms <= end) starts.push(ms);
  }

  starts.sort((a, b) => a - b);
  return Array.from(new Set(starts));
}

function startsToSegments(starts: number[], endInclusive: number): Array<[number, number]> {
  const segments: Array<[number, number]> = [];
  for (let i = 0; i < starts.length; i++) {
    const start = starts[i]!;
    const end = i + 1 < starts.length ? starts[i + 1]! - 1 : endInclusive;
    if (start <= end) segments.push([start, end]);
  }
  return segments;
}

export function calculatePages(
  printArea: CellRange,
  colWidthsPts: number[],
  rowHeightsPts: number[],
  setup: PageSetup,
  manualBreaks: ManualPageBreaks = { rowBreaksAfter: [], colBreaksAfter: [] },
): Page[] {
  const norm = normalizeRange(printArea);

  const { w: wIn, h: hIn } = paperSizeInches(setup.paperSize);
  let pageW = wIn * POINTS_PER_INCH;
  let pageH = hIn * POINTS_PER_INCH;
  if (setup.orientation === "landscape") {
    [pageW, pageH] = [pageH, pageW];
  }

  const printableW =
    pageW - (setup.margins.left + setup.margins.right) * POINTS_PER_INCH;
  const printableH =
    pageH - (setup.margins.top + setup.margins.bottom) * POINTS_PER_INCH;

  const scale = scaleFactor(norm, colWidthsPts, rowHeightsPts, setup);
  const effW = scale > 0 ? printableW / scale : printableW;
  const effH = scale > 0 ? printableH / scale : printableH;

  const colStarts = breakStarts(
    norm.startCol,
    norm.endCol,
    colWidthsPts,
    effW,
    manualBreaks.colBreaksAfter.map((c) => c + 1),
  );
  const rowStarts = breakStarts(
    norm.startRow,
    norm.endRow,
    rowHeightsPts,
    effH,
    manualBreaks.rowBreaksAfter.map((r) => r + 1),
  );

  const colSegs = startsToSegments(colStarts, norm.endCol);
  const rowSegs = startsToSegments(rowStarts, norm.endRow);

  const pages: Page[] = [];
  for (const [rs, re] of rowSegs) {
    for (const [cs, ce] of colSegs) {
      pages.push({ startRow: rs, endRow: re, startCol: cs, endCol: ce });
    }
  }
  return pages;
}

