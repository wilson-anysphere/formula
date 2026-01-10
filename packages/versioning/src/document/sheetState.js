import { cellKey } from "../diff/semanticDiff.js";

const decoder = new TextDecoder();

/**
 * Convert a snapshot produced by `apps/desktop/src/document/DocumentController.encodeState()`
 * into the `SheetState` shape expected by `semanticDiff`.
 *
 * @param {Uint8Array} snapshot
 * @param {{ sheetId: string }} opts
 * @returns {{ cells: Map<string, { value?: any, formula?: string | null, format?: any }> }}
 */
export function sheetStateFromDocumentSnapshot(snapshot, opts) {
  const sheetId = opts?.sheetId;
  if (!sheetId) throw new Error("sheetId is required");

  let parsed;
  try {
    parsed = JSON.parse(decoder.decode(snapshot));
  } catch (err) {
    throw new Error("Invalid document snapshot: not valid JSON");
  }

  const sheets = Array.isArray(parsed?.sheets) ? parsed.sheets : [];
  /** @type {Map<string, any>} */
  const cells = new Map();

  const sheet = sheets.find((s) => s?.id === sheetId);
  if (!sheet) return { cells };

  const entries = Array.isArray(sheet?.cells) ? sheet.cells : [];
  for (const entry of entries) {
    const row = Number(entry?.row);
    const col = Number(entry?.col);
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    cells.set(cellKey(row, col), {
      value: entry?.value ?? null,
      formula: entry?.formula ?? null,
      format: entry?.format ?? null,
    });
  }

  return { cells };
}

