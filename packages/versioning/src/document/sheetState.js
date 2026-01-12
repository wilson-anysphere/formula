import { cellKey } from "../diff/semanticDiff.js";

const decoder = new TextDecoder();

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function deepMerge(base, patch) {
  if (!isPlainObject(base) || !isPlainObject(patch)) return patch;
  const out = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (isPlainObject(value) && isPlainObject(out[key])) {
      out[key] = deepMerge(out[key], value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

function mergeStyleLayers(base, layer) {
  if (layer == null) return base;
  if (!isPlainObject(base)) base = {};
  if (!isPlainObject(layer)) return deepMerge(base, layer);
  if (Object.keys(layer).length === 0) return base;
  return deepMerge(base, layer);
}

function normalizeAxisFormats(raw) {
  /** @type {Map<number, any>} */
  const out = new Map();
  if (!raw) return out;

  if (Array.isArray(raw)) {
    for (const entry of raw) {
      const index = Array.isArray(entry) ? entry[0] : entry?.index ?? entry?.col ?? entry?.row;
      const format = Array.isArray(entry) ? entry[1] : entry?.format ?? entry?.style ?? entry?.value;
      const idx = Number(index);
      if (!Number.isInteger(idx) || idx < 0) continue;
      out.set(idx, format ?? null);
    }
    return out;
  }

  if (typeof raw === "object") {
    for (const [key, value] of Object.entries(raw)) {
      const idx = Number(key);
      if (!Number.isInteger(idx) || idx < 0) continue;
      out.set(idx, value ?? null);
    }
  }

  return out;
}

function normalizeRangeRuns(raw) {
  /** @type {Array<{ startRow: number, startCol: number, endRow: number, endCol: number, format: any }>} */
  const out = [];
  if (!Array.isArray(raw)) return out;
  for (const run of raw) {
    if (!run || typeof run !== "object") continue;
    const startRow = Number(run.startRow ?? run.start?.row ?? run.sr);
    const startCol = Number(run.startCol ?? run.start?.col ?? run.sc);
    const endRow = Number(run.endRow ?? run.end?.row ?? run.er);
    const endCol = Number(run.endCol ?? run.end?.col ?? run.ec);
    if (!Number.isInteger(startRow) || startRow < 0) continue;
    if (!Number.isInteger(startCol) || startCol < 0) continue;
    if (!Number.isInteger(endRow) || endRow < 0) continue;
    if (!Number.isInteger(endCol) || endCol < 0) continue;
    const format = run.format ?? run.style ?? run.value ?? null;
    out.push({ startRow, startCol, endRow, endCol, format });
  }
  return out;
}

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

  // Layered formatting (Task 44):
  // - sheet default format
  // - column defaults
  // - row defaults
  // (Task 118) optional range format runs
  const sheetDefaultFormat =
    sheet?.sheetFormat ?? sheet?.defaultFormat ?? sheet?.format ?? sheet?.sheetDefaultFormat ?? null;
  const colFormats = normalizeAxisFormats(sheet?.colFormats ?? sheet?.columnFormats ?? sheet?.colFormat ?? null);
  const rowFormats = normalizeAxisFormats(sheet?.rowFormats ?? sheet?.rowsFormats ?? sheet?.rowFormat ?? null);
  const formatRuns = normalizeRangeRuns(
    sheet?.formatRuns ?? sheet?.rangeFormatRuns ?? sheet?.rangeRuns ?? sheet?.formattingRuns ?? null,
  );

  const entries = Array.isArray(sheet?.cells) ? sheet.cells : [];
  for (const entry of entries) {
    const row = Number(entry?.row);
    const col = Number(entry?.col);
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    let format = {};
    format = mergeStyleLayers(format, sheetDefaultFormat);
    format = mergeStyleLayers(format, colFormats.get(col) ?? null);
    format = mergeStyleLayers(format, rowFormats.get(row) ?? null);
    for (const run of formatRuns) {
      if (row < run.startRow || row > run.endRow) continue;
      if (col < run.startCol || col > run.endCol) continue;
      format = mergeStyleLayers(format, run.format);
    }
    format = mergeStyleLayers(format, entry?.format ?? null);

    const normalizedFormat = isPlainObject(format) && Object.keys(format).length === 0 ? null : format;

    cells.set(cellKey(row, col), {
      value: entry?.value ?? null,
      formula: entry?.formula ?? null,
      format: normalizedFormat,
    });
  }

  return { cells };
}
