import type { CellRange } from "../../document/coords.js";
import type { DocumentController } from "../../document/documentController.js";

export type CsvExportOptions = {
  delimiter?: string;
  newline?: "\n" | "\r\n";
  maxCells?: number;
  dlp?: { documentId: string; classificationStore: any; policy: any };
};

export function exportCellGridToCsv(grid: any[][], options?: CsvExportOptions): string;

export function exportDocumentRangeToCsv(
  doc: DocumentController,
  sheetId: string,
  range: CellRange | string,
  options?: CsvExportOptions,
): string;

