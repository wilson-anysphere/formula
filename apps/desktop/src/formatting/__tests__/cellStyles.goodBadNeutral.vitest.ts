import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import {
  applyGoodBadNeutralCellStyle,
  GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS,
  type GoodBadNeutralCellStyleId,
} from "../cellStyles.js";

describe("cell styles: good/bad/neutral", () => {
  it.each([
    ["good", "#FFC6EFCE", "#FF006100"],
    ["bad", "#FFFFC7CE", "#FF9C0006"],
    ["neutral", "#FFFFEB9C", "#FF9C6500"],
  ] satisfies Array<[GoodBadNeutralCellStyleId, string, string]>)(
    "applies %s preset fill + font color",
    (presetId, expectedFillColor, expectedFontColor) => {
      const doc = new DocumentController();
      const preset = GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS[presetId];

      const applied = applyGoodBadNeutralCellStyle(
        doc,
        "Sheet1",
        [{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }],
        presetId,
      );
      expect(applied).toBe(true);

      const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 }) as any;
      expect(style?.fill?.fgColor).toBe(expectedFillColor);
      expect(style?.font?.color).toBe(expectedFontColor);

      // Ensure we used the expected constants (guard against accidental drift).
      expect(preset.fillColor).toBe(expectedFillColor);
      expect(preset.fontColor).toBe(expectedFontColor);
    },
  );
});
