import { parseRangeA1 } from "../document/coords.js";

function normalizeRanges(rangeOrRanges) {
  if (Array.isArray(rangeOrRanges)) return rangeOrRanges;
  return [rangeOrRanges];
}

function allCellsMatch(doc, sheetId, rangeOrRanges, predicate) {
  for (const range of normalizeRanges(rangeOrRanges)) {
    const r = typeof range === "string" ? parseRangeA1(range) : range;
    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const style = doc.getCellFormat(sheetId, { row, col });
        if (!predicate(style)) return false;
      }
    }
  }
  return true;
}

export function toggleBold(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean" ? options.next : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.bold));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { bold: next } }, { label: "Bold" });
  }
}

export function toggleItalic(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.italic));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { italic: next } }, { label: "Italic" });
  }
}

export function toggleUnderline(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.underline));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { underline: next } }, { label: "Underline" });
  }
}

export function setFontSize(doc, sheetId, range, sizePt) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { size: sizePt } }, { label: "Font size" });
  }
}

export function setFontColor(doc, sheetId, range, argb) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { color: argb } }, { label: "Font color" });
  }
}

export function setFillColor(doc, sheetId, range, argb) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(
      sheetId,
      r,
      { fill: { pattern: "solid", fgColor: argb } },
      { label: "Fill color" },
    );
  }
}

const DEFAULT_BORDER_ARGB = "FF000000";

export function applyAllBorders(doc, sheetId, range, { style = "thin", color } = {}) {
  const resolvedColor = color ?? `#${DEFAULT_BORDER_ARGB}`;
  const edge = { style, color: resolvedColor };
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(
      sheetId,
      r,
      { border: { left: edge, right: edge, top: edge, bottom: edge } },
      { label: "Borders" },
    );
  }
}

export function setHorizontalAlign(doc, sheetId, range, align) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(
      sheetId,
      r,
      { alignment: { horizontal: align } },
      { label: "Horizontal align" },
    );
  }
}

export function toggleWrap(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.alignment?.wrapText));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { alignment: { wrapText: next } }, { label: "Wrap" });
  }
}

export const NUMBER_FORMATS = {
  currency: "$#,##0.00",
  percent: "0%",
  date: "m/d/yyyy",
};

export function applyNumberFormatPreset(doc, sheetId, range, preset) {
  const code = NUMBER_FORMATS[preset];
  if (!code) throw new Error(`Unknown number format preset: ${preset}`);
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { numberFormat: code }, { label: "Number format" });
  }
}
