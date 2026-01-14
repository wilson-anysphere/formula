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

  let imported: ImportedSheetBackgroundImageInfo[] = [];
  try {
    imported = (await backend.listImportedSheetBackgroundImages()) ?? [];
  } catch {
    imported = [];
  }

  const docAny = app.getDocument() as any;

  const supportsExternalImages = typeof docAny?.applyExternalImageDeltas === "function";
  const supportsExternalViews = typeof docAny?.applyExternalSheetViewDeltas === "function";

  /** @type {Map<string, string>} */
  const backgroundImageIdBySheetId = new Map();
  /** @type {Map<string, { bytes: Uint8Array, mimeType?: string }>} */
  const imagesById = new Map();

  if (Array.isArray(imported)) {
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

      if (!imagesById.has(imageId)) {
        try {
          imagesById.set(imageId, { bytes: decodeBase64ToBytes(bytesBase64), ...(mimeType ? { mimeType } : {}) });
        } catch {
          // Skip corrupt/un-decodable image bytes.
          continue;
        }
      }

      backgroundImageIdBySheetId.set(sheetId, imageId);
    }
  }

  // 1) Apply image bytes into the document image store (non-undoable) so loading a workbook does
  // not create undo history. (The file contents are the baseline state.)
  if (supportsExternalImages && imagesById.size > 0) {
    /** @type {any[]} */
    const imageDeltas: any[] = [];
    for (const [imageId, entry] of imagesById.entries()) {
      const before = typeof docAny.getImage === "function" ? docAny.getImage(imageId) : null;
      imageDeltas.push({ imageId, before, after: { bytes: entry.bytes, ...(entry.mimeType ? { mimeType: entry.mimeType } : {}) } });
    }
    try {
      docAny.applyExternalImageDeltas(imageDeltas, { source: "backend", markDirty: false });
    } catch {
      // ignore
    }
  } else if (typeof docAny?.setImage === "function") {
    // Fallback for older DocumentController builds.
    for (const [imageId, entry] of imagesById.entries()) {
      try {
        docAny.setImage(imageId, { bytes: entry.bytes, ...(entry.mimeType ? { mimeType: entry.mimeType } : {}) });
      } catch {
        // ignore
      }
    }
  }

  // 2) Apply sheet view background ids (also non-undoable during workbook load).
  const sheets = (() => {
    try {
      return workbookSheetStore.listAll();
    } catch {
      return [];
    }
  })();

  if (supportsExternalViews && Array.isArray(sheets) && sheets.length > 0) {
    /** @type {any[]} */
    const viewDeltas: any[] = [];
    for (const sheet of sheets) {
      const sheetId = typeof (sheet as any)?.id === "string" ? String((sheet as any).id).trim() : "";
      if (!sheetId) continue;
      const before = docAny.getSheetView(sheetId);
      const desired = backgroundImageIdBySheetId.get(sheetId) ?? null;
      const current = typeof before?.backgroundImageId === "string" && before.backgroundImageId.trim() !== "" ? before.backgroundImageId : null;
      if (current === desired) continue;
      const after = { ...(before ?? {}) };
      if (desired) after.backgroundImageId = desired;
      else delete after.backgroundImageId;
      viewDeltas.push({ sheetId, before, after });
    }

    if (viewDeltas.length > 0) {
      try {
        docAny.applyExternalSheetViewDeltas(viewDeltas, { source: "backend", markDirty: false });
      } catch {
        // ignore
      }
    }
  } else {
    // Fallback: best-effort set via the app helper (older builds).
    for (const [sheetId, imageId] of backgroundImageIdBySheetId.entries()) {
      try {
        app.setSheetBackgroundImageId(sheetId, imageId);
      } catch {
        // ignore
      }
    }
  }
}
