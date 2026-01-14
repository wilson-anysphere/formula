/**
 * Fast, deterministic content hash intended for incremental indexing cache keys.
 *
 * Uses WebCrypto SHA-256 when available; falls back to FNV-1a 64-bit for
 * environments without WebCrypto.
 *
 * @returns lowercase hex digest
 */
export function contentHash(text: string): Promise<string>;

/**
 * Alias for {@link contentHash}.
 *
 * @returns lowercase hex digest
 */
export function sha256Hex(text: string): Promise<string>;

