const NON_BREAKING_SPACES = new Set(["\u00A0", "\u202F", "\u2060"]);
const WHITESPACE_RE = /\s/u;

import GraphemeSplitter from "grapheme-splitter";

const GRAPHEME_SPLITTER = new GraphemeSplitter();

/** @type {Map<string, Intl.Segmenter>} */
const SEGMENTER_CACHE = new Map();

/**
 * @param {string | undefined} locale
 * @param {"grapheme" | "word"} granularity
 */
function getSegmenter(locale, granularity) {
  const normalizedLocale = locale ?? "und";
  const key = `${granularity}:${normalizedLocale}`;
  const cached = SEGMENTER_CACHE.get(key);
  if (cached) return cached;
  const segmenter = new Intl.Segmenter(normalizedLocale, { granularity });
  SEGMENTER_CACHE.set(key, segmenter);
  return segmenter;
}

/**
 * @param {string} ch
 * @returns {boolean}
 */
function isBreakableWhitespaceChar(ch) {
  return WHITESPACE_RE.test(ch) && !NON_BREAKING_SPACES.has(ch);
}

/**
 * @param {string} s
 * @returns {boolean}
 */
function isBreakableWhitespaceSegment(s) {
  for (const ch of s) {
    if (!isBreakableWhitespaceChar(ch)) return false;
  }
  return s.length > 0;
}

/**
 * Break opportunities at grapheme boundaries. Returns an array of indices in the input string
 * representing valid end positions for a line.
 *
 * @param {string} text
 * @param {string | undefined} locale
 * @returns {number[]}
 */
export function graphemeBreakPositions(text, locale) {
  if (!text) return [0];
  // Prefer a library implementation to avoid subtle Intl/ICU version differences across platforms.
  const graphemes = GRAPHEME_SPLITTER.splitGraphemes(text);
  const positions = [];
  let idx = 0;
  for (const g of graphemes) {
    idx += g.length;
    positions.push(idx);
  }
  return positions.length ? positions : [text.length];
}

/**
 * Break opportunities at the start of breakable whitespace runs (so the whitespace is excluded
 * from the line). Always includes `text.length` as a final break.
 *
 * @param {string} text
 * @param {string | undefined} locale
 * @returns {{ breaks: number[], wordBreakSet: Set<number> }}
 */
export function wordBreakPositions(text, locale) {
  const breaks = [];
  const wordBreakSet = new Set();

  if (typeof Intl === "undefined" || typeof Intl.Segmenter === "undefined") {
    // Best-effort fallback:
    // - break before runs of breakable whitespace (so the whitespace is excluded from the line),
    // - otherwise allow breaks at every grapheme boundary.
    for (let i = 0; i < text.length; i++) {
      if (!isBreakableWhitespaceChar(text[i])) continue;
      breaks.push(i);
      wordBreakSet.add(i);
      while (i + 1 < text.length && isBreakableWhitespaceChar(text[i + 1])) i++;
    }
    for (const b of graphemeBreakPositions(text, locale)) breaks.push(b);
  } else {
    const segmenter = getSegmenter(locale, "word");
    for (const seg of segmenter.segment(text)) {
      // For whitespace, break at the *start* of the whitespace segment so trailing whitespace is
      // excluded from the rendered line (and can be skipped when starting the next line).
      if (WHITESPACE_RE.test(seg.segment) && isBreakableWhitespaceSegment(seg.segment)) {
        breaks.push(seg.index);
        wordBreakSet.add(seg.index);
        continue;
      }

      // Otherwise break at segment boundaries (UAX #29 word boundaries).
      breaks.push(seg.index + seg.segment.length);
    }
  }

  // Ensure final break and make monotonic.
  breaks.push(text.length);
  breaks.sort((a, b) => a - b);
  const deduped = [];
  for (let i = 0; i < breaks.length; i++) {
    if (i === 0 || breaks[i] !== breaks[i - 1]) deduped.push(breaks[i]);
  }
  return { breaks: deduped, wordBreakSet };
}

/**
 * @param {string} text
 * @param {number} index
 * @param {number} end
 * @returns {number}
 */
export function skipBreakableWhitespace(text, index, end) {
  let i = index;
  while (i < end && isBreakableWhitespaceChar(text[i])) i++;
  return i;
}
