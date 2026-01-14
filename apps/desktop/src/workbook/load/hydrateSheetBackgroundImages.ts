import type { SpreadsheetApp } from "../../app/spreadsheetApp";
import type { WorkbookSheetStore } from "../../sheets/workbookSheetStore";
import type { ImportedSheetBackgroundImageInfo } from "../../tauri/workbookBackend";

function decodeBase64ToBytes(data: string): Uint8Array {
  const trimmed = String(data ?? "").trim();
  if (!trimmed) return new Uint8Array();

  // Browser path.
  if (typeof atob === "function") {
    try {
      const binary = atob(trimmed);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      return bytes;
    } catch {
      // Fall through to Node fallback (useful for tests).
    }
  }

  // Node/test fallback (Vitest can run without a browser atob()).
  const BufferCtor = (globalThis as any).Buffer as
    | { from: (input: string, encoding: string) => { [Symbol.iterator](): Iterator<number> } }
    | undefined;
  if (BufferCtor?.from) {
    try {
      return Uint8Array.from(BufferCtor.from(trimmed, "base64"));
    } catch {
      // ignore
    }
  }

  return new Uint8Array();
}

export async function hydrateSheetBackgroundImagesFromBackend(opts: {
  app: SpreadsheetApp;
  workbookSheetStore: WorkbookSheetStore;
  backend: { listImportedSheetBackgroundImages: () => Promise<ImportedSheetBackgroundImageInfo[]> };
}): Promise<void> {
  const { app, workbookSheetStore, backend } = opts;

  // Always clear any stale sheet background mappings for the current workbook.
  // Background images are sourced from the opened XLSX package; if the new workbook does not
  // specify any, we must not keep prior background ids around.
  try {
    for (const sheet of workbookSheetStore.listAll()) {
      app.setSheetBackgroundImageId(sheet.id, null);
    }
  } catch {
    // ignore
  }

  let imported: ImportedSheetBackgroundImageInfo[] = [];
  try {
    imported = (await backend.listImportedSheetBackgroundImages()) ?? [];
  } catch {
    return;
  }

  if (!Array.isArray(imported) || imported.length === 0) {
    return;
  }

  const docAny = app.getDocument() as any;
  const supportsBatch =
    typeof docAny?.beginBatch === "function" &&
    typeof docAny?.endBatch === "function" &&
    typeof docAny?.cancelBatch === "function" &&
    typeof docAny?.setImage === "function";

  const loadedImageIds = new Set<string>();

  let batchStarted = false;
  if (supportsBatch) {
    try {
      docAny.beginBatch({ label: "Load worksheet background images" });
      batchStarted = true;
    } catch {
      // Fall through to non-batched path below.
    }
  }

  try {
    for (const entry of imported) {
      const sheetName = typeof (entry as any)?.sheet_name === "string" ? String((entry as any).sheet_name).trim() : "";
      if (!sheetName) continue;
      const sheetId = workbookSheetStore.resolveIdByName(sheetName);
      if (!sheetId) continue;

      const imageId = typeof (entry as any)?.image_id === "string" ? String((entry as any).image_id).trim() : "";
      if (!imageId) continue;

      const bytesBase64 =
        typeof (entry as any)?.bytes_base64 === "string" ? String((entry as any).bytes_base64).trim() : "";
      if (!bytesBase64) continue;

      const mimeType =
        typeof (entry as any)?.mime_type === "string" && String((entry as any).mime_type).trim() !== ""
          ? String((entry as any).mime_type).trim()
          : undefined;

      if (!loadedImageIds.has(imageId)) {
        loadedImageIds.add(imageId);
        try {
          const bytes = decodeBase64ToBytes(bytesBase64);
          if (typeof docAny?.setImage === "function") {
            docAny.setImage(imageId, { bytes, ...(mimeType ? { mimeType } : {}) });
          }
        } catch {
          // Skip corrupt/un-decodable image bytes.
          continue;
        }
      }

      try {
        app.setSheetBackgroundImageId(sheetId, imageId);
      } catch {
        // Ignore per-sheet failures; continue to other sheets.
      }
    }
  } finally {
    if (batchStarted && typeof docAny?.endBatch === "function") {
      try {
        docAny.endBatch();
      } catch {
        try {
          docAny.cancelBatch();
        } catch {
          // ignore
        }
      }
    }
  }
}
