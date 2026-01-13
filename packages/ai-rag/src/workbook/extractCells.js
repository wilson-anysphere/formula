import { getCellRaw, normalizeCell } from "./normalizeCell.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * @param {any} sheet
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @param {{ maxRows?: number, maxCols?: number, signal?: AbortSignal }} [opts]
 */
export function extractCells(sheet, rect, opts) {
  const signal = opts?.signal;
  throwIfAborted(signal);
  const maxRows = opts?.maxRows ?? Number.POSITIVE_INFINITY;
  const maxCols = opts?.maxCols ?? Number.POSITIVE_INFINITY;
  const rMax = Math.min(rect.r1, rect.r0 + maxRows - 1);
  const cMax = Math.min(rect.c1, rect.c0 + maxCols - 1);

  const out = [];
  for (let r = rect.r0; r <= rMax; r += 1) {
    throwIfAborted(signal);
    const row = [];
    // Avoid a signal check on every single cell, but still check frequently enough
    // to cancel long extractions promptly.
    let abortCountdown = 0;
    for (let c = rect.c0; c <= cMax; c += 1) {
      if (abortCountdown === 0) {
        throwIfAborted(signal);
        abortCountdown = 256;
      }
      abortCountdown -= 1;
      row.push(normalizeCell(getCellRaw(sheet, r, c)));
    }
    out.push(row);
  }
  return out;
}
