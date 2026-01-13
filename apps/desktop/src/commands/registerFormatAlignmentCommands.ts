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
    () =>
      applyFormattingToSelection(
        t("command.format.alignLeft"),
        (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "left"),
      ),
    { category, keywords: ["align", "alignment", "left", "horizontal"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.alignCenter",
    t("command.format.alignCenter"),
    () =>
      applyFormattingToSelection(
        t("command.format.alignCenter"),
        (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "center"),
      ),
    { category, keywords: ["align", "alignment", "center", "horizontal"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.alignRight",
    t("command.format.alignRight"),
    () =>
      applyFormattingToSelection(
        t("command.format.alignRight"),
        (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "right"),
      ),
    { category, keywords: ["align", "alignment", "right", "horizontal"] },
  );

  // --- Vertical alignment -----------------------------------------------------

  const applyVerticalAlign = (label: string, value: "top" | "center" | "bottom"): void => {
    applyFormattingToSelection(label, (doc, sheetId, ranges) => {
      let applied = true;
      for (const range of ranges) {
        const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: value } }, { label: "Vertical align" });
        if (ok === false) applied = false;
      }
      return applied;
    });
  };

  commandRegistry.registerBuiltinCommand("format.alignTop", t("command.format.alignTop"), () => applyVerticalAlign(t("command.format.alignTop"), "top"), {
    category,
    keywords: ["align", "alignment", "top", "vertical"],
  });

  commandRegistry.registerBuiltinCommand(
    "format.alignMiddle",
    t("command.format.alignMiddle"),
    () =>
      // Spreadsheet vertical alignment uses "center" (Excel/OOXML); the grid maps this to CSS middle.
      applyVerticalAlign(t("command.format.alignMiddle"), "center"),
    { category, keywords: ["align", "alignment", "middle", "center", "vertical"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.alignBottom",
    t("command.format.alignBottom"),
    () => applyVerticalAlign(t("command.format.alignBottom"), "bottom"),
    { category, keywords: ["align", "alignment", "bottom", "vertical"] },
  );

  // --- Indent -----------------------------------------------------------------

  commandRegistry.registerBuiltinCommand(
    "format.increaseIndent",
    t("command.format.increaseIndent"),
    () => {
      const current = activeCellIndentLevel();
      const next = Math.min(250, current + 1);
      if (next === current) return;
      applyFormattingToSelection(t("command.format.increaseIndent"), (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
          if (ok === false) applied = false;
        }
        return applied;
      });
    },
    { category, keywords: ["indent", "increase indent", "alignment"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.decreaseIndent",
    t("command.format.decreaseIndent"),
    () => {
      const current = activeCellIndentLevel();
      const next = Math.max(0, current - 1);
      if (next === current) return;
      applyFormattingToSelection(t("command.format.decreaseIndent"), (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
          if (ok === false) applied = false;
        }
        return applied;
      });
    },
    { category, keywords: ["indent", "decrease indent", "outdent", "alignment"] },
  );

  // --- Text rotation ----------------------------------------------------------

  const applyTextRotation = (label: string, value: number): void => {
    applyFormattingToSelection(label, (doc, sheetId, ranges) => {
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
    () => applyTextRotation(t("command.format.textRotation.angleCounterclockwise"), 45),
    { category, keywords: ["text rotation", "orientation", "angle", "counterclockwise", "rotate"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.angleClockwise",
    t("command.format.textRotation.angleClockwise"),
    () => applyTextRotation(t("command.format.textRotation.angleClockwise"), -45),
    { category, keywords: ["text rotation", "orientation", "angle", "clockwise", "rotate"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.verticalText",
    t("command.format.textRotation.verticalText"),
    () =>
      // Excel/OOXML uses 255 as a sentinel for vertical text (stacked).
      applyTextRotation(t("command.format.textRotation.verticalText"), 255),
    { category, keywords: ["text rotation", "orientation", "vertical text", "stacked"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.rotateUp",
    t("command.format.textRotation.rotateUp"),
    () => applyTextRotation(t("command.format.textRotation.rotateUp"), 90),
    { category, keywords: ["text rotation", "orientation", "rotate", "up"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.textRotation.rotateDown",
    t("command.format.textRotation.rotateDown"),
    () => applyTextRotation(t("command.format.textRotation.rotateDown"), -90),
    { category, keywords: ["text rotation", "orientation", "rotate", "down"] },
  );

  commandRegistry.registerBuiltinCommand(
    "format.openAlignmentDialog",
    t("command.format.openAlignmentDialog"),
    () => openAlignmentDialog(),
    { category, keywords: ["alignment", "format cells", "dialog", "cell alignment"] },
  );
}
