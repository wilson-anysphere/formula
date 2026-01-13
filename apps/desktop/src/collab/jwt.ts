import type { DocumentRole } from "@formula/collab-session";

const VALID_DOCUMENT_ROLES: ReadonlySet<string> = new Set(["owner", "admin", "editor", "commenter", "viewer"]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function base64UrlToBase64(input: string): string {
  const normalized = input.replace(/-/g, "+").replace(/_/g, "/");
  const pad = (4 - (normalized.length % 4)) % 4;
  return `${normalized}${"=".repeat(pad)}`;
}

function tryDecodeBase64ToBytes(base64: string): Uint8Array | null {
  const atobFn = (globalThis as any)?.atob as ((data: string) => string) | undefined;
  if (typeof atobFn === "function") {
    try {
      const binary = atobFn(base64);
      const out = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i += 1) {
        out[i] = binary.charCodeAt(i);
      }
      return out;
    } catch {
      // ignore
    }
  }

  const BufferCtor = (globalThis as any)?.Buffer as { from?: (data: string, encoding: string) => Uint8Array } | undefined;
  if (typeof BufferCtor?.from === "function") {
    try {
      return BufferCtor.from(base64, "base64");
    } catch {
      // ignore
    }
  }

  return null;
}

function tryDecodeBase64UrlToUtf8String(input: string): string | null {
  const bytes = tryDecodeBase64ToBytes(base64UrlToBase64(input));
  if (!bytes) return null;

  try {
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

/**
 * Best-effort decode of a JWT payload.
 *
 * Security note: This does **not** verify the signature. Callers must treat the output as
 * untrusted data and apply strict validation (or fall back to safe defaults).
 *
 * This intentionally never throws and never logs token contents.
 */
export function tryDecodeJwtPayload(token: string): Record<string, unknown> | null {
  const trimmed = typeof token === "string" ? token.trim() : "";
  if (!trimmed) return null;

  const parts = trimmed.split(".");
  if (parts.length !== 3) return null;
  const [headerSegment, payloadSegment, signatureSegment] = parts;
  if (!headerSegment || !payloadSegment || !signatureSegment) return null;

  const decoded = tryDecodeBase64UrlToUtf8String(payloadSegment);
  if (!decoded) return null;

  try {
    const parsed = JSON.parse(decoded) as unknown;
    return isRecord(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

function normalizeRole(value: unknown): DocumentRole | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  if (!normalized) return null;
  return VALID_DOCUMENT_ROLES.has(normalized) ? (normalized as DocumentRole) : null;
}

function normalizeUserId(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  return normalized ? normalized : null;
}

export type JwtDerivedCollabSessionPermissions = {
  /**
   * `sub` from the JWT, when present.
   *
   * This may be `null` even when the token decodes as a JWT (e.g. missing/invalid `sub`).
   */
  userId: string | null;
  role: DocumentRole;
  rangeRestrictions: unknown[];
};

/**
 * Best-effort extraction of collaboration permission claims from a JWT token.
 *
 * Returns `null` when the token doesn't decode as a JWT payload (or payload is malformed).
 * Otherwise returns defaults for missing/invalid claims:
 * - role: defaults to `"editor"`
 * - rangeRestrictions: defaults to `[]`
 */
export function tryDeriveCollabSessionPermissionsFromJwtToken(token: string): JwtDerivedCollabSessionPermissions | null {
  const payload = tryDecodeJwtPayload(token);
  if (!payload) return null;

  const userId = normalizeUserId(payload.sub);
  const role = normalizeRole(payload.role) ?? "editor";
  const rangeRestrictions = Array.isArray(payload.rangeRestrictions) ? payload.rangeRestrictions : [];

  return { userId, role, rangeRestrictions };
}

export function resolveCollabSessionPermissionsFromToken(options: {
  token?: string;
  fallbackUserId: string;
}): { role: DocumentRole; rangeRestrictions: unknown[]; userId: string } {
  const token = options.token;
  if (!token) {
    return { role: "editor", rangeRestrictions: [], userId: options.fallbackUserId };
  }

  const derived = tryDeriveCollabSessionPermissionsFromJwtToken(token);
  if (!derived) {
    return { role: "editor", rangeRestrictions: [], userId: options.fallbackUserId };
  }

  return {
    role: derived.role,
    rangeRestrictions: derived.rangeRestrictions,
    userId: derived.userId ?? options.fallbackUserId,
  };
}
