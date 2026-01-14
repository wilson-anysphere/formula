import type { SpreadsheetApp } from "../app/spreadsheetApp";
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

function activeCellFontSizePt(app: SpreadsheetApp): number {
  // Match the more defensive `registerBuiltinCommands` font-size step helper:
  // - Keep safe when `SpreadsheetApp` is stubbed in tests.
  // - Fall back to the typical Excel default when we cannot resolve a size.
  try {
    const sheetId = app.getCurrentSheetId?.();
    const cell = app.getActiveCell?.();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const docAny = app.getDocument?.() as any;
    if (!sheetId || !cell || !docAny) return 11;
    const effectiveSize = docAny.getCellFormat?.(sheetId, cell)?.font?.size;
    const state = docAny.getCell?.(sheetId, cell);
    const style = docAny.styleTable?.get?.(state?.styleId ?? 0) ?? {};
    const size = typeof effectiveSize === "number" ? effectiveSize : style.font?.size;
    return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : 11;
  } catch {
    return 11;
  }
}

function stepFontSize(current: number, direction: "increase" | "decrease"): number {
  const value = Number(current);
  const resolved = Number.isFinite(value) && value > 0 ? value : 11;
  if (direction === "increase") {
    for (const step of FONT_SIZE_PRESETS) {
      if (step > resolved + 1e-6) return step;
    }
    return resolved;
  }

  for (let i = FONT_SIZE_PRESETS.length - 1; i >= 0; i -= 1) {
    const step = FONT_SIZE_PRESETS[i]!;
    if (step < resolved - 1e-6) return step;
  }
  return resolved;
}

export const FORMAT_FONT_NAME_PRESET_COMMAND_IDS = Object.keys(FONT_NAME_PRESETS).map((id) => `format.fontName.${id}`);
export const FORMAT_FONT_SIZE_PRESET_COMMAND_IDS = FONT_SIZE_PRESETS.map((size) => `format.fontSize.${size}`);
export const FORMAT_FONT_SIZE_STEP_COMMAND_IDS = ["format.increaseFontSize", "format.decreaseFontSize"] as const;

export function registerBuiltinFormatFontCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  applyFormattingToSelection: ApplyFormattingToSelection;
}): void {
  const { commandRegistry, app, applyFormattingToSelection } = params;
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

  commandRegistry.registerBuiltinCommand(
    "format.increaseFontSize",
    t("command.format.fontSize.increase"),
    () => {
      const current = activeCellFontSizePt(app);
      const next = stepFontSize(current, "increase");
      if (next === current) return;
      applyFormattingToSelection(t("command.format.fontSize.increase"), (doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
    },
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.decreaseFontSize",
    t("command.format.fontSize.decrease"),
    () => {
      const current = activeCellFontSizePt(app);
      const next = stepFontSize(current, "decrease");
      if (next === current) return;
      applyFormattingToSelection(t("command.format.fontSize.decrease"), (doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
    },
    { category },
  );
}
