import type { DocumentController } from "../document/documentController.js";

export type CellCoord = { row: number; col: number };
export type CellRange = { start: CellCoord; end: CellCoord };
export type RangeInput = string | CellRange;
export type RangeInputOrRanges = RangeInput | RangeInput[];

export type HorizontalAlign = "left" | "center" | "right" | "justify" | (string & {});
export type NumberFormatPreset = "currency" | "percent" | "date";

export function toggleBold(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { next?: boolean },
): boolean;
export function toggleItalic(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { next?: boolean },
): boolean;
export function toggleUnderline(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { next?: boolean },
): boolean;
export function toggleStrikethrough(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { next?: boolean },
): boolean;

export function applyAllBorders(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { style?: string; color?: string },
): boolean;

export function applyOutsideBorders(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { style?: string; color?: string },
): boolean;

export function setHorizontalAlign(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  align: HorizontalAlign,
): boolean;

export function toggleWrap(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  options?: { next?: boolean },
): boolean;

export function setFontSize(doc: DocumentController, sheetId: string, range: RangeInputOrRanges, sizePt: number): boolean;
export function setFontColor(doc: DocumentController, sheetId: string, range: RangeInputOrRanges, argb: string): boolean;
export function setFillColor(doc: DocumentController, sheetId: string, range: RangeInputOrRanges, argb: string): boolean;

export const NUMBER_FORMATS: Record<NumberFormatPreset, string>;
export function applyNumberFormatPreset(
  doc: DocumentController,
  sheetId: string,
  range: RangeInputOrRanges,
  preset: NumberFormatPreset,
): boolean;
