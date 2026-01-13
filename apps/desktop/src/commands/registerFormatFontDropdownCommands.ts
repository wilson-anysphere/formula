import type { DocumentController } from "../document/documentController.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import {
  applyAllBorders,
  setFillColor,
  setFontColor,
  type CellRange,
} from "../formatting/toolbar.js";

type ApplyFormattingToSelection = (
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
  options?: { forceBatch?: boolean },
) => void;

export function registerFormatFontDropdownCommands(params: {
  commandRegistry: CommandRegistry;
  category: string;
  applyFormattingToSelection: ApplyFormattingToSelection;
}): void {
  const { commandRegistry, category, applyFormattingToSelection } = params;

  commandRegistry.registerBuiltinCommand(
    "format.clearFormats",
    "Clear Formats",
    () =>
      applyFormattingToSelection("Clear formats", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, null, { label: "Clear formats" });
          if (ok === false) applied = false;
        }
        return applied;
      }),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.clearContents",
    "Clear Contents",
    () =>
      applyFormattingToSelection("Clear contents", (doc, sheetId, ranges) => {
        for (const range of ranges) {
          doc.clearRange(sheetId, range, { label: "Clear contents" });
        }
      }),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.clearAll",
    "Clear All",
    () =>
      applyFormattingToSelection(
        "Clear all",
        (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            doc.clearRange(sheetId, range, { label: "Clear all" });
            const ok = doc.setRangeFormat(sheetId, range, null, { label: "Clear all" });
            if (ok === false) applied = false;
          }
          return applied;
        },
        { forceBatch: true },
      ),
    { category },
  );

  const defaultBorderColor = ["#", "FF", "000000"].join("");

  commandRegistry.registerBuiltinCommand(
    "format.borders.none",
    "No Border",
    () =>
      applyFormattingToSelection("Borders", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { border: null }, { label: "Borders" });
          if (ok === false) applied = false;
        }
        return applied;
      }),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.borders.all",
    "All Borders",
    () => applyFormattingToSelection("Borders", (doc, sheetId, ranges) => applyAllBorders(doc, sheetId, ranges)),
    { category },
  );

  const registerBoxBorderCommand = (commandId: string, title: string, edgeStyle: "thin" | "thick"): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      () =>
        applyFormattingToSelection(
          "Borders",
          (doc, sheetId, ranges) => {
            let applied = true;
            const edge = { style: edgeStyle, color: defaultBorderColor };
            const applyBorder = (targetRange: CellRange, patch: Record<string, any> | null) => {
              const ok = doc.setRangeFormat(sheetId, targetRange, patch, { label: "Borders" });
              if (ok === false) applied = false;
            };
            for (const range of ranges) {
              const startRow = range.start.row;
              const endRow = range.end.row;
              const startCol = range.start.col;
              const endCol = range.end.col;

              // Top edge.
              applyBorder({ start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } }, { border: { top: edge } });

              // Bottom edge.
              applyBorder({ start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } }, { border: { bottom: edge } });

              // Left edge.
              applyBorder({ start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } }, { border: { left: edge } });

              // Right edge.
              applyBorder({ start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } }, { border: { right: edge } });
            }
            return applied;
          },
          { forceBatch: true },
        ),
      { category },
    );
  };

  registerBoxBorderCommand("format.borders.outside", "Outside Borders", "thin");
  registerBoxBorderCommand("format.borders.thickBox", "Thick Box Border", "thick");

  const registerEdgeBorderCommand = (kind: "top" | "bottom" | "left" | "right", title: string): void => {
    commandRegistry.registerBuiltinCommand(
      `format.borders.${kind}`,
      title,
      () =>
        applyFormattingToSelection(
          "Borders",
          (doc, sheetId, ranges) => {
            let applied = true;
            const edge = { style: "thin", color: defaultBorderColor };
            const borderPatch = { border: { [kind]: edge } } as any;
            for (const range of ranges) {
              const startRow = range.start.row;
              const endRow = range.end.row;
              const startCol = range.start.col;
              const endCol = range.end.col;

              const targetRange = (() => {
                switch (kind) {
                  case "bottom":
                    return { start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } };
                  case "top":
                    return { start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } };
                  case "left":
                    return { start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } };
                  case "right":
                    return { start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } };
                  default:
                    return range;
                }
              })();

              const ok = doc.setRangeFormat(sheetId, targetRange, borderPatch, { label: "Borders" });
              if (ok === false) applied = false;
            }
            return applied;
          },
          { forceBatch: true },
        ),
      { category },
    );
  };

  registerEdgeBorderCommand("top", "Top Border");
  registerEdgeBorderCommand("bottom", "Bottom Border");
  registerEdgeBorderCommand("left", "Left Border");
  registerEdgeBorderCommand("right", "Right Border");

  commandRegistry.registerBuiltinCommand(
    "format.fillColor.moreColors",
    "Fill Color: More Colors…",
    () => commandRegistry.executeCommand("format.fillColor"),
    { category },
  );

  const registerFillPreset = (commandId: string, title: string, argb: string | null): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      () => {
        if (argb == null) {
          applyFormattingToSelection("Fill color", (doc, sheetId, ranges) => {
            let applied = true;
            for (const range of ranges) {
              const ok = doc.setRangeFormat(sheetId, range, { fill: null }, { label: "Fill color" });
              if (ok === false) applied = false;
            }
            return applied;
          });
          return;
        }
        applyFormattingToSelection("Fill color", (doc, sheetId, ranges) => setFillColor(doc, sheetId, ranges, argb));
      },
      { category },
    );
  };

  registerFillPreset("format.fillColor.none", "Fill Color: No Fill", null);
  registerFillPreset("format.fillColor.lightGray", "Fill Color: Light Gray", ["#", "FF", "D9D9D9"].join(""));
  registerFillPreset("format.fillColor.yellow", "Fill Color: Yellow", ["#", "FF", "FFFF00"].join(""));
  registerFillPreset("format.fillColor.blue", "Fill Color: Blue", ["#", "FF", "0000FF"].join(""));
  registerFillPreset("format.fillColor.green", "Fill Color: Green", ["#", "FF", "00FF00"].join(""));
  registerFillPreset("format.fillColor.red", "Fill Color: Red", ["#", "FF", "FF0000"].join(""));

  commandRegistry.registerBuiltinCommand(
    "format.fontColor.moreColors",
    "Font Color: More Colors…",
    () => commandRegistry.executeCommand("format.fontColor"),
    { category },
  );

  commandRegistry.registerBuiltinCommand(
    "format.fontColor.automatic",
    "Font Color: Automatic",
    () =>
      applyFormattingToSelection("Font color", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { font: { color: null } }, { label: "Font color" });
          if (ok === false) applied = false;
        }
        return applied;
      }),
    { category },
  );

  const registerFontPreset = (commandId: string, title: string, argb: string): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      () => {
        applyFormattingToSelection("Font color", (doc, sheetId, ranges) => setFontColor(doc, sheetId, ranges, argb));
      },
      { category },
    );
  };

  registerFontPreset("format.fontColor.black", "Font Color: Black", ["#", "FF", "000000"].join(""));
  registerFontPreset("format.fontColor.blue", "Font Color: Blue", ["#", "FF", "0000FF"].join(""));
  registerFontPreset("format.fontColor.red", "Font Color: Red", ["#", "FF", "FF0000"].join(""));
  registerFontPreset("format.fontColor.green", "Font Color: Green", ["#", "FF", "00FF00"].join(""));
}
