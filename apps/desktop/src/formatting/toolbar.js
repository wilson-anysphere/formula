import { parseRangeA1 } from "../document/coords.js";

function allCellsMatch(doc, sheetId, range, predicate) {
  const r = typeof range === "string" ? parseRangeA1(range) : range;
  for (let row = r.start.row; row <= r.end.row; row++) {
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = doc.getCell(sheetId, { row, col });
      const style = doc.styleTable.get(cell.styleId);
      if (!predicate(style)) return false;
    }
  }
  return true;
}

export function toggleBold(doc, sheetId, range) {
  const next = !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.bold));
  doc.setRangeFormat(sheetId, range, { font: { bold: next } }, { label: "Bold" });
}

export function toggleItalic(doc, sheetId, range) {
  const next = !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.italic));
  doc.setRangeFormat(sheetId, range, { font: { italic: next } }, { label: "Italic" });
}

export function toggleUnderline(doc, sheetId, range) {
  const next = !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.underline));
  doc.setRangeFormat(sheetId, range, { font: { underline: next } }, { label: "Underline" });
}

export function setFontSize(doc, sheetId, range, sizePt) {
  doc.setRangeFormat(sheetId, range, { font: { size: sizePt } }, { label: "Font size" });
}

export function setFontColor(doc, sheetId, range, argb) {
  doc.setRangeFormat(sheetId, range, { font: { color: argb } }, { label: "Font color" });
}

export function setFillColor(doc, sheetId, range, argb) {
  doc.setRangeFormat(
    sheetId,
    range,
    { fill: { pattern: "solid", fgColor: argb } },
    { label: "Fill color" },
  );
}

const DEFAULT_BORDER_ARGB = "FF000000";

export function applyAllBorders(doc, sheetId, range, { style = "thin", color } = {}) {
  const resolvedColor = color ?? `#${DEFAULT_BORDER_ARGB}`;
  const edge = { style, color: resolvedColor };
  doc.setRangeFormat(
    sheetId,
    range,
    { border: { left: edge, right: edge, top: edge, bottom: edge } },
    { label: "Borders" },
  );
}

export function setHorizontalAlign(doc, sheetId, range, align) {
  doc.setRangeFormat(
    sheetId,
    range,
    { alignment: { horizontal: align } },
    { label: "Horizontal align" },
  );
}

export function toggleWrap(doc, sheetId, range) {
  const next = !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.alignment?.wrapText));
  doc.setRangeFormat(sheetId, range, { alignment: { wrapText: next } }, { label: "Wrap" });
}

export const NUMBER_FORMATS = {
  currency: "$#,##0.00",
  percent: "0%",
  date: "m/d/yyyy",
};

export function applyNumberFormatPreset(doc, sheetId, range, preset) {
  const code = NUMBER_FORMATS[preset];
  if (!code) throw new Error(`Unknown number format preset: ${preset}`);
  doc.setRangeFormat(sheetId, range, { numberFormat: code }, { label: "Number format" });
}
