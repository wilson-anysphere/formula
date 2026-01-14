import type { SnapshotCell } from "../mergeFormattingIntoSnapshot.js";
import { MAX_INSERT_IMAGE_BYTES } from "../../drawings/insertImageLimits.js";

import { coerceBase64StringWithinLimit } from "./base64.js";

export type ImportedEmbeddedCellImage = {
  worksheet_part: string;
  sheet_name?: string | null;
  row: number;
  col: number;
  image_id: string;
  bytes_base64: string;
  mime_type: string;
  alt_text?: string | null;
};

export type SnapshotImageEntry = {
  id: string;
  bytesBase64: string;
  mimeType?: string | null;
};

function normalizeString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function warnDefault(message: string): void {
  // Avoid crashing in environments where `console` is unavailable.
  try {
    console.warn(message);
  } catch {
    // ignore
  }
}

function resolveSheetIdForEmbeddedImage(options: {
  entry: ImportedEmbeddedCellImage;
  resolveSheetIdByName: (name: string) => string | null;
  sheetIdsInOrder: string[];
}): string | null {
  const sheetName = normalizeString(options.entry.sheet_name);
  if (sheetName) {
    const resolved = options.resolveSheetIdByName(sheetName);
    if (resolved) return resolved;
    // Some backends may already return sheet ids in the `sheet_name` slot; accept it if it
    // matches a known sheet id.
    return options.sheetIdsInOrder.includes(sheetName) ? sheetName : null;
  }

  const part = normalizeString(options.entry.worksheet_part);
  if (!part) return null;

  // Best-effort fallback: attempt to map `xl/worksheets/sheet{N}.xml` -> Nth sheet.
  const match = /sheet(\d+)\.xml$/i.exec(part);
  if (!match) return null;
  const idx = Number(match[1]) - 1;
  if (!Number.isInteger(idx) || idx < 0 || idx >= options.sheetIdsInOrder.length) return null;
  return options.sheetIdsInOrder[idx] ?? null;
}

function coordKey(row: number, col: number): number {
  // Ensure the mapping is collision-free for the desktop default caps (10k rows, 200 cols) and
  // stays well below MAX_SAFE_INTEGER.
  return row * 1_000_000 + col;
}

/**
 * Merge embedded-in-cell images extracted from an XLSX package into the workbook snapshot
 * consumed by `SpreadsheetApp.restoreDocumentState`.
 *
 * This is intentionally best-effort and defensive: callers should be able to load workbooks even
 * when extraction fails or produces malformed entries.
 */
export function mergeEmbeddedCellImagesIntoSnapshot<
  TSheet extends { id: string; name: string; cells: SnapshotCell[] },
>(options: Readonly<{
  sheets: TSheet[];
  images?: SnapshotImageEntry[] | null;
  embeddedCellImages: ImportedEmbeddedCellImage[] | null | undefined;
  resolveSheetIdByName: (name: string) => string | null;
  sheetIdsInOrder: string[];
  maxRows: number;
  maxCols: number;
  warn?: ((message: string) => void) | null;
}>): { sheets: TSheet[]; images: SnapshotImageEntry[] } {
  const warn = options.warn ?? warnDefault;
  const maxRows = Number.isFinite(options.maxRows) && options.maxRows > 0 ? Math.floor(options.maxRows) : 0;
  const maxCols = Number.isFinite(options.maxCols) && options.maxCols > 0 ? Math.floor(options.maxCols) : 0;

  const sheetsById = new Map<string, TSheet>();
  for (const sheet of options.sheets) {
    sheetsById.set(sheet.id, sheet);
  }

  const imagesById = new Map<string, SnapshotImageEntry>();
  for (const img of options.images ?? []) {
    const id = normalizeString((img as any)?.id);
    const bytesBase64 = typeof (img as any)?.bytesBase64 === "string" ? (img as any).bytesBase64 : "";
    if (!id || !bytesBase64) continue;
    const normalizedBase64 = coerceBase64StringWithinLimit(bytesBase64, MAX_INSERT_IMAGE_BYTES);
    if (!normalizedBase64) continue;
    const mimeTypeRaw = (img as any)?.mimeType;
    const mimeType =
      mimeTypeRaw === null || typeof mimeTypeRaw === "string"
        ? mimeTypeRaw
        : typeof mimeTypeRaw === "undefined"
          ? undefined
          : null;
    const entry: SnapshotImageEntry = { id, bytesBase64: normalizedBase64 };
    if (mimeTypeRaw !== undefined) entry.mimeType = mimeType;
    imagesById.set(id, entry);
  }

  const cellIndexBySheet = new Map<string, Map<number, number>>();

  const embeddedImages = Array.isArray(options.embeddedCellImages) ? options.embeddedCellImages : [];
  for (const raw of embeddedImages) {
    if (!raw || typeof raw !== "object") continue;

    const entry = raw as ImportedEmbeddedCellImage;
    const row = Number(entry.row);
    const col = Number(entry.col);
    if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) continue;
    if (row >= maxRows || col >= maxCols) {
      warn(
        `[formula][desktop] Skipping embedded cell image outside workbook load bounds (row=${row}, col=${col}, maxRows=${maxRows}, maxCols=${maxCols}).`,
      );
      continue;
    }

    const imageId = normalizeString(entry.image_id);
    const bytesBase64 = normalizeString(entry.bytes_base64);
    if (!imageId || !bytesBase64) continue;

    const sheetId =
      resolveSheetIdForEmbeddedImage({
        entry,
        resolveSheetIdByName: options.resolveSheetIdByName,
        sheetIdsInOrder: options.sheetIdsInOrder,
      }) ?? null;
    if (!sheetId) {
      warn(
        `[formula][desktop] Skipping embedded cell image (imageId=${imageId}) because sheet could not be resolved (sheet_name=${String(
          entry.sheet_name ?? "",
        )}, worksheet_part=${String(entry.worksheet_part ?? "")}).`,
      );
      continue;
    }

    const sheet = sheetsById.get(sheetId);
    if (!sheet) continue;

    if (!imagesById.has(imageId)) {
      const mimeType = normalizeString(entry.mime_type);
      const normalizedBase64 = coerceBase64StringWithinLimit(bytesBase64, MAX_INSERT_IMAGE_BYTES);
      if (!normalizedBase64) {
        warn(`[formula][desktop] Skipping embedded cell image (imageId=${imageId}) because image bytes are too large.`);
        continue;
      }
      const imageEntry: SnapshotImageEntry = { id: imageId, bytesBase64: normalizedBase64 };
      if (mimeType) imageEntry.mimeType = mimeType;
      imagesById.set(imageId, imageEntry);
    }

    let indexByCoord = cellIndexBySheet.get(sheetId);
    if (!indexByCoord) {
      indexByCoord = new Map();
      for (let i = 0; i < sheet.cells.length; i += 1) {
        const cell = sheet.cells[i] as any;
        const r = Number(cell?.row);
        const c = Number(cell?.col);
        if (!Number.isInteger(r) || r < 0 || !Number.isInteger(c) || c < 0) continue;
        indexByCoord.set(coordKey(r, c), i);
      }
      cellIndexBySheet.set(sheetId, indexByCoord);
    }

    const key = coordKey(row, col);
    const existingIdx = indexByCoord.get(key);
    const existing = existingIdx != null ? sheet.cells[existingIdx] : null;
    const existingFormat = existing ? (existing as any).format ?? null : null;
    const existingFormula = existing && typeof (existing as any).formula === "string" ? (existing as any).formula : null;

    const altText = normalizeString(entry.alt_text);
    const imageValue: any = { type: "image", value: { imageId } };
    if (altText) imageValue.value.altText = altText;

    const nextCell: SnapshotCell = {
      row,
      col,
      value: imageValue,
      formula: existingFormula,
      format: existingFormat,
    };

    if (existingIdx != null) {
      sheet.cells[existingIdx] = nextCell;
    } else {
      indexByCoord.set(key, sheet.cells.length);
      sheet.cells.push(nextCell);
    }
  }

  const images = Array.from(imagesById.values());
  images.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  return { sheets: options.sheets, images };
}
