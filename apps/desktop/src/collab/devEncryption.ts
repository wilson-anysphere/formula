import { parseA1Range } from "../charts/a1.js";
import type { CellAddress, CollabSessionOptions } from "@formula/collab-session";
import type { CellEncryptionKey } from "@formula/collab-encryption";

/**
 * Dev-only URL-param toggle for exercising end-to-end cell encryption in collab sessions.
 *
 * This is intentionally *not* production key management. It exists so developers can
 * run two clients against the same doc and verify:
 *   - protected cells are written with `enc` payloads (no plaintext `value`/`formula`)
 *   - clients without the key see masked values ("###")
 *
 * Enable with:
 *   - `?collabEncrypt=1`
 *   - `?collabEncryptRange=Sheet1!A1:C10` (optional; defaults to `Sheet1!A1:C10`)
 */

export const DEV_COLLAB_ENCRYPT_PARAM = "collabEncrypt";
export const DEV_COLLAB_ENCRYPT_RANGE_PARAM = "collabEncryptRange";
export const DEFAULT_DEV_COLLAB_ENCRYPT_RANGE = "Sheet1!A1:C10";

const DEV_ENCRYPTION_SALT = "formula-dev-collab-cell-encryption-v1";

export type DevEncryptionRange = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

function parseBooleanFlag(raw: string | null): boolean {
  if (raw == null) return false;
  const normalized = raw.trim().toLowerCase();
  if (normalized === "") return true; // `?flag` / `?flag=`
  if (normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on") return true;
  return false;
}

export function parseDevEncryptionRange(
  rangeRef: string,
  options: { defaultSheetId?: string; resolveSheetIdByName?: (name: string) => string | null } = {}
): DevEncryptionRange | null {
  const parsed = parseA1Range(rangeRef);
  if (!parsed) return null;
  const rawSheetName = String(parsed.sheetName ?? options.defaultSheetId ?? "Sheet1").trim();
  if (!rawSheetName) return null;
  // `parseA1Range` yields a user-facing sheet name (the identifier used in formulas).
  // Collab cell keys use a stable sheet id. When a resolver is provided, map the
  // display name to the id so `collabEncryptRange=My Sheet!A1:C10` works even when
  // sheet ids differ from names.
  const resolvedSheetId =
    parsed.sheetName && typeof options.resolveSheetIdByName === "function"
      ? options.resolveSheetIdByName(rawSheetName)
      : null;
  const sheetId = String(resolvedSheetId ?? rawSheetName).trim();
  if (!sheetId) return null;
  return {
    sheetId,
    startRow: parsed.startRow,
    startCol: parsed.startCol,
    endRow: parsed.endRow,
    endCol: parsed.endCol,
  };
}

// Deterministic 32-bit FNV-1a hash for seeding dev key derivation.
function fnv1a32(text: string): number {
  let hash = 0x811c9dc5; // offset basis
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function xorshift32(seed: number): number {
  let x = seed >>> 0;
  x ^= x << 13;
  x ^= x >>> 17;
  x ^= x << 5;
  return x >>> 0;
}

export function deriveDevEncryptionKey(docId: string): CellEncryptionKey {
  const seedText = `${DEV_ENCRYPTION_SALT}:${docId}`;
  let seed = fnv1a32(seedText);
  if (seed === 0) seed = 0x9e3779b9; // avoid a stuck xorshift state

  const keyBytes = new Uint8Array(32);
  let state = seed;
  for (let i = 0; i < keyBytes.length; i += 1) {
    state = xorshift32(state);
    keyBytes[i] = state & 0xff;
  }

  const keyId = `dev:${fnv1a32(`keyId:${seedText}`).toString(16).padStart(8, "0")}`;
  return { keyId, keyBytes };
}

function cellInRange(cell: CellAddress, range: DevEncryptionRange): boolean {
  if (cell.sheetId !== range.sheetId) return false;
  if (cell.row < range.startRow || cell.row > range.endRow) return false;
  if (cell.col < range.startCol || cell.col > range.endCol) return false;
  return true;
}

export function createDevEncryptionConfig(opts: {
  docId: string;
  range: DevEncryptionRange;
}): NonNullable<CollabSessionOptions["encryption"]> {
  const key = deriveDevEncryptionKey(opts.docId);
  const range = opts.range;
  return {
    // Only provide the key for the configured demo range.
    //
    // This keeps the dev toggle from accidentally granting edit access to cells that
    // were encrypted with a different key id (e.g. production-managed encryption).
    //
    // CollabSession enforces that encrypted cell overwrites require a *matching*
    // `keyId`, but we still scope the resolver so the dev toggle only affects the
    // intended range and doesn't accidentally expose keys across unrelated cells.
    keyForCell: (cell: CellAddress) => (cellInRange(cell, range) ? key : null),
    shouldEncryptCell: (cell: CellAddress) => cellInRange(cell, range),
  };
}

export function resolveDevCollabEncryptionFromParams(opts: {
  params: URLSearchParams;
  docId: string;
  defaultSheetId?: string;
  defaultRangeRef?: string;
  resolveSheetIdByName?: (name: string) => string | null;
}): NonNullable<CollabSessionOptions["encryption"]> | null {
  const enabled = parseBooleanFlag(opts.params.get(DEV_COLLAB_ENCRYPT_PARAM));
  if (!enabled) return null;

  const defaultSheetId = opts.defaultSheetId ?? "Sheet1";
  const defaultRangeRef = opts.defaultRangeRef ?? DEFAULT_DEV_COLLAB_ENCRYPT_RANGE;
  const rangeRef = opts.params.get(DEV_COLLAB_ENCRYPT_RANGE_PARAM) ?? defaultRangeRef;

  let range = parseDevEncryptionRange(rangeRef, { defaultSheetId, resolveSheetIdByName: opts.resolveSheetIdByName });
  if (!range && rangeRef !== defaultRangeRef) {
    range = parseDevEncryptionRange(defaultRangeRef, { defaultSheetId, resolveSheetIdByName: opts.resolveSheetIdByName });
  }
  if (!range) return null;

  return createDevEncryptionConfig({ docId: opts.docId, range });
}

export function resolveDevCollabEncryptionFromSearch(opts: {
  search: string;
  docId: string;
  defaultSheetId?: string;
  defaultRangeRef?: string;
  resolveSheetIdByName?: (name: string) => string | null;
}): NonNullable<CollabSessionOptions["encryption"]> | null {
  const raw = opts.search.startsWith("?") ? opts.search.slice(1) : opts.search;
  const params = new URLSearchParams(raw);
  return resolveDevCollabEncryptionFromParams({
    params,
    docId: opts.docId,
    defaultSheetId: opts.defaultSheetId,
    defaultRangeRef: opts.defaultRangeRef,
    resolveSheetIdByName: opts.resolveSheetIdByName,
  });
}
