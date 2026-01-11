const NON_BREAKING_SPACES = new Set(["\u00A0", "\u202F", "\u2060"]);
const WHITESPACE_RE = /\s/u;

import GraphemeSplitter from "grapheme-splitter";
import LineBreaker from "linebreak";

const GRAPHEME_SPLITTER = new GraphemeSplitter();

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

  // Use the Unicode Line Breaking Algorithm (UAX #14). This provides high-quality word-ish break
  // opportunities for scripts with and without whitespace, and behaves well around punctuation/emoji.
  //
  // Note: The line breaker reports break opportunities *after* the character at `position - 1`. For
  // whitespace this would keep trailing spaces at the end of the line. We map any break that lands
  // after a run of breakable whitespace to the *start* of that whitespace run so lines exclude
  // trailing whitespace and the engine can skip it on the next line.
  const breaker = new LineBreaker(text);
  let lastPos = -1;
  for (let brk = breaker.nextBreak(); brk; brk = breaker.nextBreak()) {
    let pos = brk.position;
    if (pos <= 0) continue;
    if (pos > text.length) pos = text.length;

    if (pos > 0 && isBreakableWhitespaceChar(text[pos - 1])) {
      let start = pos - 1;
      while (start > 0 && isBreakableWhitespaceChar(text[start - 1])) start--;
      pos = start;
      wordBreakSet.add(pos);
    }

    if (pos !== lastPos) {
      breaks.push(pos);
      lastPos = pos;
    }
  }

  // Ensure final break.
  if (breaks.length === 0 || breaks[breaks.length - 1] !== text.length) breaks.push(text.length);
  return { breaks, wordBreakSet };
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
