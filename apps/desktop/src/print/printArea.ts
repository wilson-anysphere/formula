import type { CellRange, SheetPrintSettings } from "./types";

export function setPrintArea(settings: SheetPrintSettings, range: CellRange): SheetPrintSettings {
  return {
    ...settings,
    printArea: [range],
  };
}

export function clearPrintArea(settings: SheetPrintSettings): SheetPrintSettings {
  const { printArea: _ignored, ...rest } = settings;
  return rest;
}

