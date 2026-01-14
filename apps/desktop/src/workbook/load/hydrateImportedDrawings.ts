import { convertModelDrawingObjectToUiDrawingObject } from "../../drawings/modelAdapters";
import { MAX_INSERT_IMAGE_BYTES } from "../../drawings/insertImageLimits.js";
import type { WorkbookSheetStore } from "../../sheets/workbookSheetStore";

import { coerceBase64StringWithinLimit } from "./base64.js";

export type ImportedDrawingLayerPayload = {
  drawings?: unknown;
  images?: unknown;
};

type SnapshotImageEntry = { id: string; bytesBase64: string; mimeType?: string | null };

/**
 * Convert the backend "imported drawing layer" payload into the DocumentController snapshot
 * schema (`snapshot.images` + `snapshot.drawingsBySheet`).
 *
 * Best-effort: malformed entries are skipped rather than throwing.
 */
export function buildImportedDrawingLayerSnapshotAdditions(
  imported: ImportedDrawingLayerPayload | null | undefined,
  workbookSheetStore: WorkbookSheetStore,
): { images: SnapshotImageEntry[]; drawingsBySheet: Record<string, any[]> } | null {
  const drawingsRaw = imported && typeof imported === "object" ? (imported as any).drawings : null;
  const imagesRaw = imported && typeof imported === "object" ? (imported as any).images : null;

  const drawingsEntries = Array.isArray(drawingsRaw) ? drawingsRaw : [];
  const imagesEntries = Array.isArray(imagesRaw) ? imagesRaw : [];

  const imagesById = new Map<string, SnapshotImageEntry>();
  for (const raw of imagesEntries) {
    const e = raw as any;
    const id = typeof e?.id === "string" ? e.id.trim() : "";
    if (!id) continue;
    if (imagesById.has(id)) continue;

    const bytesBase64 =
      typeof e?.bytesBase64 === "string"
        ? e.bytesBase64
        : typeof e?.bytes_base64 === "string"
          ? e.bytes_base64
          : "";
    if (!bytesBase64) continue;

    const normalizedBase64 = coerceBase64StringWithinLimit(bytesBase64, MAX_INSERT_IMAGE_BYTES);
    if (!normalizedBase64) continue;

    const mimeTypeRaw = Object.prototype.hasOwnProperty.call(e ?? {}, "mimeType")
      ? e.mimeType
      : Object.prototype.hasOwnProperty.call(e ?? {}, "mime_type")
        ? e.mime_type
        : undefined;
    const mimeType =
      mimeTypeRaw === undefined || mimeTypeRaw === null
        ? mimeTypeRaw
        : typeof mimeTypeRaw === "string"
          ? mimeTypeRaw
          : null;

    const entry: SnapshotImageEntry = { id, bytesBase64: normalizedBase64, ...(mimeTypeRaw !== undefined ? { mimeType } : {}) };
    imagesById.set(id, entry);
  }

  const images: SnapshotImageEntry[] = Array.from(imagesById.values()).sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  const drawingsBySheet: Record<string, any[]> = {};

  for (const rawSheet of drawingsEntries) {
    const entry = rawSheet as any;
    const sheetName =
      typeof entry?.sheet_name === "string"
        ? entry.sheet_name
        : typeof entry?.sheetName === "string"
          ? entry.sheetName
          : "";
    if (!sheetName) continue;
    const sheetId = workbookSheetStore.resolveIdByName(sheetName) ?? null;
    if (!sheetId) continue;

    const objects = Array.isArray(entry?.objects) ? entry.objects : [];
    if (objects.length === 0) continue;

    const list: any[] = drawingsBySheet[sheetId] ? [...drawingsBySheet[sheetId]!] : [];
    for (const obj of objects) {
      try {
        const ui = convertModelDrawingObjectToUiDrawingObject(obj, { sheetId });
        // DocumentController persists drawing ids as strings.
        list.push({ ...ui, id: String(ui.id) });
      } catch {
        // ignore malformed drawing objects (best-effort)
      }
    }
    if (list.length > 0) drawingsBySheet[sheetId] = list;
  }

  if (images.length === 0 && Object.keys(drawingsBySheet).length === 0) return null;
  return { images, drawingsBySheet };
}
