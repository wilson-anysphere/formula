import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { t } from "../i18n/index.js";
import { setHorizontalAlign, type CellRange } from "../formatting/toolbar.js";

type ApplyFormattingToSelection = (
  label: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  fn: (doc: any, sheetId: string, ranges: CellRange[]) => void | boolean,
  options?: { forceBatch?: boolean },
) => void;

export function registerFormatAlignmentCommands(params: {
  commandRegistry: CommandRegistry;
  applyFormattingToSelection: ApplyFormattingToSelection;
  activeCellIndentLevel: () => number;
  openAlignmentDialog: () => void;
}): void {
  const { commandRegistry, applyFormattingToSelection, activeCellIndentLevel, openAlignmentDialog } = params;

  const category = t("commandCategory.format");

  // --- Horizontal alignment ---------------------------------------------------

  commandRegistry.registerBuiltinCommand(
    "format.alignLeft",
    t("command.format.alignLeft"),
    () => applyFormattingToSelection("Align left", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "left")),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.alignCenter",
    t("command.format.alignCenter"),
    () =>
      applyFormattingToSelection("Align center", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "center")),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.alignRight",
    t("command.format.alignRight"),
    () =>
      applyFormattingToSelection("Align right", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "right")),
    { category },
  );

  // --- Vertical alignment -----------------------------------------------------

  const applyVerticalAlign = (value: "top" | "center" | "bottom"): void => {
    applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
      let applied = true;
      for (const range of ranges) {
        const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: value } }, { label: "Vertical align" });
        if (ok === false) applied = false;
      }
      return applied;
    });
  };

  commandRegistry.registerBuiltinCommand("format.alignTop", t("command.format.alignTop"), () => applyVerticalAlign("top"), { category });

  commandRegistry.registerBuiltinCommand(
    "format.alignMiddle",
    t("command.format.alignMiddle"),
    () =>
      // Spreadsheet vertical alignment uses "center" (Excel/OOXML); the grid maps this to CSS middle.
      applyVerticalAlign("center"),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.alignBottom",
    t("command.format.alignBottom"),
    () => applyVerticalAlign("bottom"),
    { category },
  );

  // --- Indent -----------------------------------------------------------------

  commandRegistry.registerBuiltinCommand(
    "format.increaseIndent",
    t("command.format.increaseIndent"),
    () => {
      const current = activeCellIndentLevel();
      const next = Math.min(250, current + 1);
      if (next === current) return;
      applyFormattingToSelection("Indent", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
          if (ok === false) applied = false;
        }
        return applied;
      });
    },
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.decreaseIndent",
    t("command.format.decreaseIndent"),
    () => {
      const current = activeCellIndentLevel();
      const next = Math.max(0, current - 1);
      if (next === current) return;
      applyFormattingToSelection("Indent", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
          if (ok === false) applied = false;
        }
        return applied;
      });
    },
    { category },
  );

  // --- Text rotation ----------------------------------------------------------

  const applyTextRotation = (value: number): void => {
    applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
      let applied = true;
      for (const range of ranges) {
        const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: value } }, { label: "Text orientation" });
        if (ok === false) applied = false;
      }
      return applied;
    });
  };

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.angleCounterclockwise",
    t("command.format.textRotation.angleCounterclockwise"),
    () => applyTextRotation(45),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.angleClockwise",
    t("command.format.textRotation.angleClockwise"),
    () => applyTextRotation(-45),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.verticalText",
    t("command.format.textRotation.verticalText"),
    () =>
      // Excel/OOXML uses 255 as a sentinel for vertical text (stacked).
      applyTextRotation(255),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.rotateUp",
    t("command.format.textRotation.rotateUp"),
    () => applyTextRotation(90),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.rotateDown",
    t("command.format.textRotation.rotateDown"),
    () => applyTextRotation(-90),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.openAlignmentDialog",
    t("command.format.openAlignmentDialog"),
    () => openAlignmentDialog(),
    { category },
  );
}

