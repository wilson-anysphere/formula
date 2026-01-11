import { getPrivacyLevel, privacyRank } from "./levels.js";

/**
 * @typedef {import("./levels.js").PrivacyLevel} PrivacyLevel
 * @typedef {"ignore" | "enforce" | "warn"} PrivacyMode
 */

/**
 * @typedef {{
 *   sourceId: string;
 *   level: PrivacyLevel;
 * }} SourcePrivacyInfo
 */

/**
 * @param {Iterable<string>} sourceIds
 * @param {Record<string, PrivacyLevel> | undefined} levelsBySourceId
 * @returns {SourcePrivacyInfo[]}
 */
export function collectSourcePrivacy(sourceIds, levelsBySourceId) {
  /** @type {SourcePrivacyInfo[]} */
  const out = [];
  for (const sourceId of sourceIds) {
    out.push({ sourceId, level: getPrivacyLevel(levelsBySourceId, sourceId) });
  }
  out.sort((a, b) => a.sourceId.localeCompare(b.sourceId));
  return out;
}

/**
 * @param {SourcePrivacyInfo[]} infos
 * @returns {Set<PrivacyLevel>}
 */
export function distinctPrivacyLevels(infos) {
  /** @type {Set<PrivacyLevel>} */
  const levels = new Set();
  for (const info of infos) levels.add(info.level);
  return levels;
}

/**
 * Conservative block policy for *local* data combination.
 *
 * The goal is to prevent obvious "high -> low" combinations in strict mode
 * (e.g. Private + Public). The engine can still allow other combinations by
 * executing them locally (buffered) instead of folding/pushdown.
 *
 * @param {SourcePrivacyInfo[]} infos
 * @returns {boolean}
 */
export function shouldBlockCombination(infos) {
  if (infos.length <= 1) return false;
  let min = Infinity;
  let max = -Infinity;
  for (const info of infos) {
    const rank = privacyRank(info.level);
    min = Math.min(min, rank);
    max = Math.max(max, rank);
  }
  // Block only when combining the least restrictive (public) with the most
  // restrictive (private/unknown). This matches the "obvious" exfiltration risk.
  return min === 0 && max >= 2;
}

