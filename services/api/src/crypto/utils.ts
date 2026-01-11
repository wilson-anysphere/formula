export function assertBufferLength(buf: Buffer, expected: number, name: string): void {
  if (!Buffer.isBuffer(buf)) {
    throw new TypeError(`${name} must be a Buffer`);
  }
  if (buf.length !== expected) {
    throw new RangeError(`${name} must be ${expected} bytes (got ${buf.length})`);
  }
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && (value as any).constructor === Object;
}

function sortJson(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(sortJson);
  }
  if (isPlainObject(value)) {
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(value).sort()) {
      out[key] = sortJson(value[key]);
    }
    return out;
  }
  return value;
}

/**
 * Deterministic JSON encoding suitable for use as AAD / encryption context.
 *
 * This exists so encryption context bytes are stable across runtime instances.
 * Do NOT use for security-sensitive canonicalization of untrusted input.
 */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(sortJson(value));
}

export function aadFromContext(context: unknown | null | undefined): Buffer | null {
  if (context === null || context === undefined) return null;
  return Buffer.from(canonicalJson(context), "utf8");
}

