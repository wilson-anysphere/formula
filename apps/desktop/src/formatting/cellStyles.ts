import type { DocumentController } from "../document/documentController.js";
import type { CellRange } from "./toolbar.js";

export type GoodBadNeutralCellStyleId = "good" | "bad" | "neutral";

export type CellStylePreset = {
  id: GoodBadNeutralCellStyleId;
  label: string;
  description: string;
  fillColor: string;
  fontColor: string;
  /**
   * When omitted, bold is left unchanged.
   */
  bold?: boolean;
};

export const GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS: Record<GoodBadNeutralCellStyleId, CellStylePreset> = {
  // These colors match Excel's commonly used "Good, Bad, and Neutral" palette.
  // (#AARRGGBB format)
  good: {
    id: "good",
    label: "Good",
    description: "Green fill, dark text",
    fillColor: "#FFC6EFCE",
    fontColor: "#FF006100",
    bold: true,
  },
  bad: {
    id: "bad",
    label: "Bad",
    description: "Red fill",
    fillColor: "#FFFFC7CE",
    fontColor: "#FF9C0006",
    bold: true,
  },
  neutral: {
    id: "neutral",
    label: "Neutral",
    description: "Yellow/gray fill",
    fillColor: "#FFFFEB9C",
    fontColor: "#FF9C6500",
    bold: true,
  },
};

export function getGoodBadNeutralCellStyleQuickPickItems(): Array<{
  label: string;
  value: GoodBadNeutralCellStyleId;
  description: string;
}> {
  return (Object.values(GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS) as CellStylePreset[]).map((preset) => ({
    label: preset.label,
    value: preset.id,
    description: preset.description,
  }));
}

export function applyGoodBadNeutralCellStyle(
  doc: DocumentController,
  sheetId: string,
  ranges: CellRange[],
  presetId: GoodBadNeutralCellStyleId,
): boolean {
  const preset = GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS[presetId];
  if (!preset) return false;

  const font: Record<string, unknown> = { color: preset.fontColor };
  if (typeof preset.bold === "boolean") font.bold = preset.bold;

  const patch = {
    fill: { pattern: "solid", fgColor: preset.fillColor },
    font,
  };

  let applied = true;
  for (const range of ranges) {
    const ok = doc.setRangeFormat(sheetId, range, patch, { label: "Cell style" });
    if (ok === false) applied = false;
  }
  return applied;
}

