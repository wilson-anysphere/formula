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

