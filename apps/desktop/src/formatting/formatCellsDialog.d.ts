import type { DocumentController } from "../document/documentController.js";
import type { CellRange } from "../document/coords.js";

export function applyFormatCells(
  doc: DocumentController,
  sheetId: string,
  range: string | CellRange,
  changes: Record<string, any>,
): boolean;

