import type { CellRange, PageSetup } from "./types";

export type PdfExportRequest = {
  sheetName: string;
  range: CellRange;
  pageSetup: PageSetup;
};

// The desktop app can provide a backend implementation via Tauri IPC.
export type PdfExportBackend = (req: PdfExportRequest) => Promise<Uint8Array>;

export async function exportToPdf(
  backend: PdfExportBackend,
  req: PdfExportRequest,
): Promise<Uint8Array> {
  return backend(req);
}

function decodeBase64ToBytes(data: string): Uint8Array {
  const binary = atob(data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

export async function exportSheetRangePdfViaTauri(args: {
  sheetId: string;
  range: CellRange;
  colWidthsPoints?: number[];
  rowHeightsPoints?: number[];
}): Promise<Uint8Array> {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as
    | ((cmd: string, args?: any) => Promise<any>)
    | undefined;

  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }

  const b64 = await invoke("export_sheet_range_pdf", {
    sheet_id: args.sheetId,
    range: {
      start_row: args.range.startRow,
      end_row: args.range.endRow,
      start_col: args.range.startCol,
      end_col: args.range.endCol,
    },
    col_widths_points: args.colWidthsPoints,
    row_heights_points: args.rowHeightsPoints,
  });

  return decodeBase64ToBytes(String(b64));
}
