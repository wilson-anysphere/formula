import type { DocumentController } from "../document/documentController.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { setFontSize, type CellRange } from "../formatting/toolbar.js";
import { t } from "../i18n/index.js";

export type ApplyFormattingToSelection = (
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
  options?: { forceBatch?: boolean },
) => void;

const FONT_NAME_PRESETS: Record<"calibri" | "arial" | "times" | "courier", string> = {
  calibri: "Calibri",
  arial: "Arial",
  times: "Times New Roman",
  courier: "Courier New",
};

export const FONT_SIZE_PRESETS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as const;

export const FORMAT_FONT_NAME_PRESET_COMMAND_IDS = Object.keys(FONT_NAME_PRESETS).map((id) => `format.fontName.${id}`);
export const FORMAT_FONT_SIZE_PRESET_COMMAND_IDS = FONT_SIZE_PRESETS.map((size) => `format.fontSize.${size}`);

export function registerBuiltinFormatFontCommands(params: {
  commandRegistry: CommandRegistry;
  applyFormattingToSelection: ApplyFormattingToSelection;
}): void {
  const { commandRegistry, applyFormattingToSelection } = params;
  const category = t("commandCategory.format");

  for (const [presetId, fontName] of Object.entries(FONT_NAME_PRESETS)) {
    const commandId = `format.fontName.${presetId}`;
    commandRegistry.registerBuiltinCommand(
      commandId,
      `Font: ${fontName}`,
      () =>
        applyFormattingToSelection("Font", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { font: { name: fontName } }, { label: "Font" });
            if (ok === false) applied = false;
          }
          return applied;
        }),
      { category },
    );
  }

  for (const size of FONT_SIZE_PRESETS) {
    const commandId = `format.fontSize.${size}`;
    commandRegistry.registerBuiltinCommand(
      commandId,
      `Font size: ${size}`,
      () =>
        applyFormattingToSelection("Font size", (doc, sheetId, ranges) => {
          return setFontSize(doc, sheetId, ranges, size);
        }),
      { category },
    );
  }
}
