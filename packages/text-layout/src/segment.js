const NON_BREAKING_SPACES = new Set(["\u00A0", "\u202F", "\u2060"]);
const WHITESPACE_RE = /\s/u;

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
  if (typeof Intl === "undefined" || typeof Intl.Segmenter === "undefined") {
    const positions = [];
    let idx = 0;
    for (const ch of Array.from(text)) {
      idx += ch.length;
      positions.push(idx);
    }
    return positions.length ? positions : [text.length];
  }

  const segmenter = new Intl.Segmenter(locale ?? "und", { granularity: "grapheme" });
  const positions = [];
  for (const seg of segmenter.segment(text)) {
    positions.push(seg.index + seg.segment.length);
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
    for (let i = 0; i < text.length; i++) {
      if (isBreakableWhitespaceChar(text[i])) {
        breaks.push(i);
        wordBreakSet.add(i);
        while (i + 1 < text.length && isBreakableWhitespaceChar(text[i + 1])) i++;
      }
    }
    breaks.push(text.length);
    return { breaks, wordBreakSet };
  }

  const segmenter = new Intl.Segmenter(locale ?? "und", { granularity: "word" });
  for (const seg of segmenter.segment(text)) {
    if (!WHITESPACE_RE.test(seg.segment)) continue;
    if (!isBreakableWhitespaceSegment(seg.segment)) continue;
    breaks.push(seg.index);
    wordBreakSet.add(seg.index);
  }

  breaks.push(text.length);
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

