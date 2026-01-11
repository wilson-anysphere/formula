/**
 * Power Query (Excel)-compatible privacy level model.
 *
 * The engine uses privacy levels to implement the "formula firewall": prevent
 * accidental data leakage when combining sources and when pushing computation
 * down into external systems (query folding).
 */

/**
 * @typedef {"public" | "organizational" | "private" | "unknown"} PrivacyLevel
 */

/**
 * @param {PrivacyLevel} level
 * @returns {number}
 */
export function privacyRank(level) {
  // Higher means more restrictive / sensitive.
  switch (level) {
    case "public":
      return 0;
    case "organizational":
      return 1;
    case "private":
      return 2;
    case "unknown":
      // Treat unknown as "high" in the firewall (conservative).
      return 2;
    default: {
      /** @type {never} */
      const exhausted = level;
      throw new Error(`Unsupported privacy level '${exhausted}'`);
    }
  }
}

/**
 * @param {Record<string, PrivacyLevel> | undefined} levelsBySourceId
 * @param {string | null | undefined} sourceId
 * @returns {PrivacyLevel}
 */
export function getPrivacyLevel(levelsBySourceId, sourceId) {
  if (!levelsBySourceId || !sourceId) return "unknown";
  const level = levelsBySourceId[sourceId];
  return level ?? "unknown";
}

