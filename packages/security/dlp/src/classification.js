export const CLASSIFICATION_LEVEL = Object.freeze({
  PUBLIC: "Public",
  INTERNAL: "Internal",
  CONFIDENTIAL: "Confidential",
  RESTRICTED: "Restricted",
});

export const CLASSIFICATION_LEVELS = Object.freeze([
  CLASSIFICATION_LEVEL.PUBLIC,
  CLASSIFICATION_LEVEL.INTERNAL,
  CLASSIFICATION_LEVEL.CONFIDENTIAL,
  CLASSIFICATION_LEVEL.RESTRICTED,
]);

export const DEFAULT_CLASSIFICATION = Object.freeze({
  level: CLASSIFICATION_LEVEL.PUBLIC,
  labels: [],
});

/**
 * @param {string} level
 * @returns {number}
 */
export function classificationRank(level) {
  const idx = CLASSIFICATION_LEVELS.indexOf(level);
  if (idx === -1) {
    throw new Error(`Unknown classification level: ${level}`);
  }
  return idx;
}

/**
 * @param {{level: string, labels?: string[]}|null|undefined} classification
 * @returns {{level: string, labels: string[]}}
 */
export function normalizeClassification(classification) {
  if (!classification) return { ...DEFAULT_CLASSIFICATION };
  if (!CLASSIFICATION_LEVELS.includes(classification.level)) {
    throw new Error(`Invalid classification level: ${classification.level}`);
  }
  const labels = Array.isArray(classification.labels) ? classification.labels : [];
  const cleaned = [...new Set(labels.map((l) => String(l).trim()).filter(Boolean))].sort();
  return { level: classification.level, labels: cleaned };
}

/**
 * Returns the more restrictive of two classifications. Labels are unioned so
 * audit logs preserve all tags that contributed to the result.
 *
 * @param {{level: string, labels?: string[]} | null | undefined} a
 * @param {{level: string, labels?: string[]} | null | undefined} b
 */
export function maxClassification(a, b) {
  const na = normalizeClassification(a);
  const nb = normalizeClassification(b);
  const level = classificationRank(na.level) >= classificationRank(nb.level) ? na.level : nb.level;
  const labels = [...new Set([...(na.labels || []), ...(nb.labels || [])])].sort();
  return { level, labels };
}

/**
 * @param {{level: string, labels?: string[]}|null|undefined} a
 * @param {{level: string, labels?: string[]}|null|undefined} b
 * @returns {number} -1 if a < b, 0 if equal, 1 if a > b (more restrictive)
 */
export function compareClassification(a, b) {
  const na = normalizeClassification(a);
  const nb = normalizeClassification(b);
  const ra = classificationRank(na.level);
  const rb = classificationRank(nb.level);
  if (ra === rb) return 0;
  return ra > rb ? 1 : -1;
}

