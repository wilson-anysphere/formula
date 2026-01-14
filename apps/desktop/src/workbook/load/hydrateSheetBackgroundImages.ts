import type { SpreadsheetApp } from "../../app/spreadsheetApp";
import { decodeBase64ToBytes, MAX_INSERT_IMAGE_BYTES } from "../../drawings/insertImage";
import { MAX_PNG_DIMENSION, MAX_PNG_PIXELS, readImageDimensions } from "../../drawings/pngDimensions";
import type { WorkbookSheetStore } from "../../sheets/workbookSheetStore";
import type { ImportedSheetBackgroundImageInfo } from "../../tauri/workbookBackend";

function isSheetKnownMissing(doc: any, sheetId: string): boolean {
  const id = String(sheetId ?? "").trim();
  if (!id) return true;

  const sheets: any = doc?.model?.sheets;
  const sheetMeta: any = doc?.sheetMeta;
  if (
    sheets &&
    typeof sheets.has === "function" &&
    typeof sheets.size === "number" &&
    sheetMeta &&
    typeof sheetMeta.has === "function" &&
    typeof sheetMeta.size === "number"
  ) {
    const workbookHasAnySheets = sheets.size > 0 || sheetMeta.size > 0;
    if (!workbookHasAnySheets) return false;
    return !sheets.has(id) && !sheetMeta.has(id);
  }

  return false;
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
    // Treat backend failures as "no imported backgrounds".
    imported = [];
  }

  const docAny = app.getDocument() as any;
  const supportsExternalImages = typeof docAny?.applyExternalImageDeltas === "function";
  const supportsExternalViews = typeof docAny?.applyExternalSheetViewDeltas === "function";

  const backgroundImageIdBySheetId = new Map<string, string>();
  const imagesById = new Map<string, { bytes: Uint8Array; mimeType?: string }>();

  if (Array.isArray(imported)) {
    for (const entry of imported) {
      const sheetName = typeof (entry as any)?.sheet_name === "string" ? String((entry as any).sheet_name).trim() : "";
      if (!sheetName) continue;
      const sheetId = workbookSheetStore.resolveIdByName(sheetName);
      if (!sheetId) continue;

      const imageId = typeof (entry as any)?.image_id === "string" ? String((entry as any).image_id).trim() : "";
      if (!imageId) continue;

      const bytesBase64 = typeof (entry as any)?.bytes_base64 === "string" ? String((entry as any).bytes_base64).trim() : "";
      if (!bytesBase64) continue;

      const mimeType =
        typeof (entry as any)?.mime_type === "string" && String((entry as any).mime_type).trim() !== ""
          ? String((entry as any).mime_type).trim()
          : undefined;

      // Prefer the last entry if multiple are returned for the same sheet.
      backgroundImageIdBySheetId.set(sheetId, imageId);

      if (!imagesById.has(imageId)) {
        try {
          const bytes = decodeBase64ToBytes(bytesBase64, { maxBytes: MAX_INSERT_IMAGE_BYTES });
          if (!bytes || bytes.byteLength === 0) continue;
          const dims = readImageDimensions(bytes);
          if (dims) {
            if (
              dims.width > MAX_PNG_DIMENSION ||
              dims.height > MAX_PNG_DIMENSION ||
              dims.width * dims.height > MAX_PNG_PIXELS
            ) {
              continue;
            }
          }
          imagesById.set(imageId, { bytes, ...(mimeType ? { mimeType } : {}) });
        } catch {
          // ignore invalid base64
        }
      }
    }
  }

  // 1) Apply image bytes into the document image store (non-undoable) so loading a workbook does
  // not create undo history. (The file contents are the baseline state.)
  if (supportsExternalImages && imagesById.size > 0) {
    /** @type {any[]} */
    const imageDeltas: any[] = [];
    for (const [imageId, entry] of imagesById.entries()) {
      const before = typeof docAny.getImage === "function" ? docAny.getImage(imageId) : null;
      imageDeltas.push({
        imageId,
        before,
        after: { bytes: entry.bytes, ...(entry.mimeType ? { mimeType: entry.mimeType } : {}) },
      });
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
      if (isSheetKnownMissing(docAny, sheetId)) continue;
      const before = typeof docAny.getSheetView === "function" ? docAny.getSheetView(sheetId) : null;
      const desired = backgroundImageIdBySheetId.get(sheetId) ?? null;
      const currentRaw = typeof (before as any)?.backgroundImageId === "string" ? String((before as any).backgroundImageId).trim() : "";
      const current = currentRaw ? currentRaw : null;
      if (current === desired) continue;
      const after = { ...(before ?? {}) } as any;
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
    for (const sheet of sheets) {
      const sheetId = typeof (sheet as any)?.id === "string" ? String((sheet as any).id).trim() : "";
      if (!sheetId) continue;
      if (isSheetKnownMissing(docAny, sheetId)) continue;
      const desired = backgroundImageIdBySheetId.get(sheetId) ?? null;
      try {
        app.setSheetBackgroundImageId(sheetId, desired);
      } catch {
        // ignore
      }
    }
  }
}
