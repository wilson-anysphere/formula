import { resolveCssVar } from "../theme/cssVars.js";

export type CollabUserIdentity = { id: string; name: string; color: string };

export const COLLAB_USER_STORAGE_KEY = "formula:collab:user";

const COLOR_TOKEN_PALETTE: ReadonlyArray<string> = [
  "--formula-grid-remote-presence-1",
  "--formula-grid-remote-presence-2",
  "--formula-grid-remote-presence-3",
  "--formula-grid-remote-presence-4",
  "--formula-grid-remote-presence-5",
  "--formula-grid-remote-presence-6",
  "--formula-grid-remote-presence-7",
  "--formula-grid-remote-presence-8",
  "--formula-grid-remote-presence-9",
  "--formula-grid-remote-presence-10",
  "--formula-grid-remote-presence-11",
  "--formula-grid-remote-presence-12",
];

let inMemoryIdentity: CollabUserIdentity | null = null;

function tryGetDefaultStorage(): Storage | null {
  try {
    const storage = (globalThis as any)?.localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function") return null;
    if (typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}

function normalizeHexColor(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;
  const match = /^#?([0-9a-f]{6})$/i.exec(trimmed);
  if (!match) return null;
  return `#${match[1].toLowerCase()}`;
}

function hashStringToUInt32(value: string): number {
  // FNV-1a 32-bit
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function toHexByte(value: number): string {
  const clamped = Math.max(0, Math.min(255, Math.round(value)));
  return clamped.toString(16).padStart(2, "0");
}

function fallbackColorFromHash(hash: number): string {
  // When CSS tokens are unavailable (e.g. node-only unit tests), produce a stable
  // 6-digit hex color without embedding literal hex values in the source.
  // Keep colors away from pure black/white so they remain visible on common backgrounds.
  const channel = (offset: number) => 64 + ((hash >>> offset) & 0x7f);
  const r = channel(0);
  const g = channel(7);
  const b = channel(14);
  return `#${toHexByte(r)}${toHexByte(g)}${toHexByte(b)}`;
}

function colorFromId(id: string): string {
  const hash = hashStringToUInt32(id);
  // Presence renderers accept any CSS color, but for collab identity we standardize
  // on `#rrggbb` strings so downstream metadata (comments/conflicts) is portable and
  // consistent across environments.
  const fallbackFromHash = fallbackColorFromHash(hash);
  const fallbackFromCss = resolveCssVar("--formula-grid-remote-presence-default", { fallback: fallbackFromHash });
  const fallback = normalizeHexColor(fallbackFromCss) ?? fallbackFromHash;
  if (COLOR_TOKEN_PALETTE.length === 0) return fallback;
  const idx = hash % COLOR_TOKEN_PALETTE.length;
  const token = COLOR_TOKEN_PALETTE[idx]!;
  const resolved = resolveCssVar(token, { fallback });
  return normalizeHexColor(resolved) ?? fallback;
}

function defaultNameFromId(id: string): string {
  const suffix = id.slice(0, 4);
  return suffix ? `User ${suffix}` : "User";
}

function generateId(cryptoObj: Crypto | null): string {
  const randomUUID = (cryptoObj as any)?.randomUUID;
  if (typeof randomUUID === "function") {
    try {
      return String(randomUUID.call(cryptoObj));
    } catch {
      // fall through
    }
  }
  return `${Date.now()}-${Math.random()}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object";
}

function coerceIdentity(value: unknown): CollabUserIdentity | null {
  if (!isRecord(value)) return null;
  const id = typeof value.id === "string" ? value.id.trim() : "";
  if (!id) return null;

  const name = typeof value.name === "string" ? value.name.trim() : "";
  const colorRaw = typeof value.color === "string" ? value.color : "";
  const color = colorRaw ? normalizeHexColor(colorRaw) : null;

  const normalized: CollabUserIdentity = {
    id,
    name: name || defaultNameFromId(id),
    color: color ?? colorFromId(id),
  };

  return normalized;
}

function readStoredIdentity(storage: Storage): CollabUserIdentity | null {
  const raw = storage.getItem(COLLAB_USER_STORAGE_KEY);
  if (!raw) return null;
  try {
    return coerceIdentity(JSON.parse(raw));
  } catch {
    return null;
  }
}

function writeStoredIdentity(storage: Storage, identity: CollabUserIdentity): void {
  storage.setItem(COLLAB_USER_STORAGE_KEY, JSON.stringify(identity));
}

function parseUrlOverrides(search: string): Partial<CollabUserIdentity> {
  if (!search) return {};
  try {
    const params = new URLSearchParams(search.startsWith("?") ? search : `?${search}`);

    const read = (key: string): string => params.get(key)?.trim() ?? "";

    // Prefer the explicit `collabUser*` params (recommended), but also accept
    // legacy/simple `user*` params for backwards compatibility.
    const id = read("collabUserId") || read("userId");
    const name = read("collabUserName") || read("userName");
    const colorRaw = read("collabUserColor") || read("userColor");
    const color = colorRaw ? normalizeHexColor(colorRaw) : null;

    const out: Partial<CollabUserIdentity> = {};
    if (id) out.id = id;
    if (name) out.name = name;
    if (color) out.color = color;
    return out;
  } catch {
    return {};
  }
}

export function getCollabUserIdentity(options?: {
  storage?: Storage | null;
  /** `window.location.search` (e.g. `?collabUserName=Alice`). */
  search?: string;
  crypto?: Crypto | null;
}): CollabUserIdentity {
  const storage = options?.storage ?? tryGetDefaultStorage();
  const search =
    options?.search ??
    (typeof window !== "undefined" && typeof window.location?.search === "string" ? window.location.search : "");
  const cryptoObj = options?.crypto ?? (globalThis.crypto ?? null);

  const overrides = parseUrlOverrides(search);

  // 1) Load from storage (best-effort).
  let base: CollabUserIdentity | null = null;
  if (storage) {
    try {
      base = readStoredIdentity(storage);
    } catch {
      base = null;
    }
  }

  // 2) If storage is empty/corrupt/unavailable, fall back to an in-memory identity
  // (and attempt to persist when possible).
  if (!base) {
    base =
      inMemoryIdentity ??
      (() => {
        const id = generateId(cryptoObj);
        return { id, name: defaultNameFromId(id), color: colorFromId(id) };
      })();

    if (!inMemoryIdentity) inMemoryIdentity = base;

    if (storage) {
      try {
        writeStoredIdentity(storage, base);
      } catch {
        // Ignore persistence failures (e.g. storage disabled/quota).
      }
    }
  } else if (storage) {
    // If the stored value is missing derived fields, update it.
    // (E.g. `color` palette updates or legacy/malformed data.)
    try {
      writeStoredIdentity(storage, base);
    } catch {
      // Ignore.
    }
  }

  // 3) Apply URL overrides for this run (do not overwrite stored identity).
  const id = overrides.id ?? base.id;
  const name = overrides.name ?? (overrides.id ? defaultNameFromId(id) : base.name);
  const color = overrides.color ?? (overrides.id ? colorFromId(id) : base.color);

  return { id, name: name || defaultNameFromId(id), color };
}

/**
 * Create a new identity that preserves the original user's display name, but uses a
 * stable color derived from the provided id.
 *
 * This is useful when the server enforces a canonical user id (e.g. JWT `sub`) that may
 * differ from the locally generated identity.
 */
export function overrideCollabUserIdentityId(identity: CollabUserIdentity, id: string): CollabUserIdentity {
  const normalizedId = typeof id === "string" ? id.trim() : "";
  if (!normalizedId) return identity;
  if (identity.id === normalizedId) return identity;

  return {
    id: normalizedId,
    name: identity.name || defaultNameFromId(normalizedId),
    color: colorFromId(normalizedId),
  };
}

/** @internal (tests) */
export function __resetCollabUserIdentityForTests(): void {
  inMemoryIdentity = null;
}
