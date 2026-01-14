/**
 * Deterministic JSON stringification with stable object key ordering.
 *
 * Intended for hashing/cache keys, not for human readability.
 */
export function stableJsonStringify(value: any): string;

