import { LRUCache } from "./lru.js";
import { detectBaseDirection, resolveAlign } from "./direction.js";
import { fontKey } from "./font.js";
import { graphemeBreakPositions, skipBreakableWhitespace, wordBreakPositions } from "./segment.js";

/**
 * Stable serialization for cache keys.
 *
 * Layout decisions depend on `text` + `font` + width/wrap options, but consumers may attach
 * extra metadata to runs (e.g. color/underline). Because the engine returns those runs back,
 * the layout cache key must include that metadata to avoid returning a layout object with
 * stale run properties.
 *
 * @param {unknown} value
 * @returns {string}
 */
function stableValueKey(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return `[${value.map(stableValueKey).join(",")}]`;
  if (typeof value === "object") {
    const obj = /** @type {Record<string, unknown>} */ (value);
    const keys = Object.keys(obj).sort();
    return `{${keys.map((k) => `${k}:${stableValueKey(obj[k])}`).join(",")}}`;
  }
  return String(value);
}

/**
 * @param {Record<string, unknown>} run
 * @returns {string}
 */
function runExtrasKey(run) {
  const keys = Object.keys(run)
    .filter((k) => k !== "text" && k !== "font")
    .sort();
  if (keys.length === 0) return "";
  return keys.map((k) => `${k}=${stableValueKey(run[k])}`).join(",");
}

/**
 * @typedef {import("./font.js").FontSpec} FontSpec
 * @typedef {import("./measurer.js").TextMeasurer} TextMeasurer
 * @typedef {import("./measurer.js").TextMeasurement} TextMeasurement
 */

/**
 * @typedef {"none" | "word" | "char"} WrapMode
 * @typedef {"ltr" | "rtl" | "auto"} TextDirection
 * @typedef {"left" | "right" | "center" | "start" | "end"} TextAlign
 *
 * @typedef {Object} TextRun
 * @property {string} text
 * @property {FontSpec} [font]
 * @property {string} [color]
 *
 * @typedef {Object} LayoutOptions
 * @property {string} [text]
 * @property {TextRun[]} [runs]
 * @property {FontSpec} font
 * @property {number} maxWidth
 * @property {WrapMode} wrapMode
 * @property {TextAlign} align
 * @property {TextDirection} [direction]
 * @property {number} [lineHeightPx]
 * @property {number} [maxLines]
 * @property {string} [ellipsis]
 * @property {string} [locale]
 *
 * @typedef {Object} LayoutLine
 * @property {string} text
 * @property {TextRun[]} runs
 * @property {number} width
 * @property {number} ascent
 * @property {number} descent
 * @property {number} x
 *
 * @typedef {Object} TextLayout
 * @property {LayoutLine[]} lines
 * @property {number} width
 * @property {number} height
 * @property {number} lineHeight
 * @property {"ltr" | "rtl"} direction
 * @property {number} maxWidth
 * @property {"left" | "right" | "center"} resolvedAlign
 */

export class TextLayoutEngine {
  /**
   * @param {TextMeasurer} measurer
   * @param {{ maxMeasureCacheEntries?: number, maxLayoutCacheEntries?: number }} [opts]
   */
  constructor(measurer, opts = {}) {
    this.measurer = measurer;
    this.measureCache = new LRUCache(opts.maxMeasureCacheEntries ?? 10_000);
    this.layoutCache = new LRUCache(opts.maxLayoutCacheEntries ?? 2_000);
  }

  #measurerCacheKey() {
    const k = /** @type {any} */ (this.measurer).cacheKey;
    if (!k) return "";
    if (typeof k === "function") return k.call(this.measurer);
    return String(k);
  }

  /**
   * Cached measurement for a single text run.
   *
   * Useful for consumers that need fast width measurement but not full wrapping/layout.
   *
   * @param {string} text
   * @param {FontSpec} font
   * @returns {TextMeasurement}
   */
  measure(text, font) {
    return this.#measureCached(text, font);
  }

  /**
   * @param {LayoutOptions} options
   * @returns {TextLayout}
   */
  layout(options) {
    const normalized = this.#normalizeOptions(options);
    const cacheKey = this.#layoutCacheKey(normalized);
    const cached = this.layoutCache.get(cacheKey);
    if (cached) return cached;

    const layout = this.#layoutNormalized(normalized);
    this.layoutCache.set(cacheKey, layout);
    return layout;
  }

  /**
   * @param {string} text
   * @param {FontSpec} font
   * @returns {TextMeasurement}
   */
  #measureCached(text, font) {
    const fk = fontKey(font);
    const key = `${this.#measurerCacheKey()}\n${fk}\n${text}`;
    const cached = this.measureCache.get(key);
    if (cached) return cached;
    const measurement = this.measurer.measure(text, font);
    this.measureCache.set(key, measurement);
    return measurement;
  }

  /**
   * @param {LayoutOptions} options
   */
  #normalizeOptions(options) {
    const runs = options.runs
      ? options.runs.map((r) => ({ ...r, font: r.font ?? options.font }))
      : [{ text: options.text ?? "", font: options.font }];

    const text = runs.map((r) => r.text).join("");

    const direction =
      options.direction && options.direction !== "auto"
        ? options.direction
        : detectBaseDirection(text);

    const resolvedAlign = resolveAlign(options.align, direction);

    const maxFontSize = Math.max(...runs.map((r) => r.font.sizePx));
    const lineHeight = options.lineHeightPx ?? Math.ceil(maxFontSize * 1.2);

    return {
      runs,
      text,
      defaultFont: options.font,
      direction,
      resolvedAlign,
      maxWidth: options.maxWidth,
      wrapMode: options.wrapMode,
      lineHeight,
      maxLines: options.maxLines ?? Infinity,
      ellipsis: options.ellipsis ?? "…",
      locale: options.locale,
      align: options.align,
    };
  }

  /**
   * @param {ReturnType<TextLayoutEngine["#normalizeOptions"]>} options
   * @returns {string}
   */
  #layoutCacheKey(options) {
    const runsKey = options.runs
      .map((r) => {
        const extra = runExtrasKey(/** @type {any} */ (r));
        return extra ? `${fontKey(r.font)}:${r.text}:${extra}` : `${fontKey(r.font)}:${r.text}`;
      })
      .join("|");
    return [
      `v1`,
      `mk=${this.#measurerCacheKey()}`,
      `runs=${runsKey}`,
      `mw=${options.maxWidth}`,
      `wrap=${options.wrapMode}`,
      `lh=${options.lineHeight}`,
      `ml=${options.maxLines}`,
      `el=${options.ellipsis}`,
      `dir=${options.direction}`,
      `align=${options.align}`,
      `loc=${options.locale ?? ""}`,
    ].join(";");
  }

  /**
   * @param {ReturnType<TextLayoutEngine["#normalizeOptions"]>} options
   * @returns {TextLayout}
   */
  #layoutNormalized(options) {
    const fullText = options.text;
    const paragraphRanges = this.#splitParagraphs(fullText);

    /** @type {LayoutLine[]} */
    const lines = [];

    for (const range of paragraphRanges) {
      this.#layoutParagraph(lines, options, range.start, range.end);
    }

    if (lines.length === 0) {
      lines.push({
        text: "",
        runs: [],
        width: 0,
        ascent: 0,
        descent: 0,
        x: 0,
      });
    }

    const maxLineWidth = Math.max(0, ...lines.map((l) => l.width));
    const height = lines.length * options.lineHeight;

    return {
      lines,
      width: maxLineWidth,
      height,
      lineHeight: options.lineHeight,
      direction: options.direction,
      maxWidth: options.maxWidth,
      resolvedAlign: options.resolvedAlign,
    };
  }

  /**
   * @param {LayoutLine[]} linesOut
   * @param {ReturnType<TextLayoutEngine["#normalizeOptions"]>} options
   * @param {number} paragraphStart
   * @param {number} paragraphEnd
   */
  #layoutParagraph(linesOut, options, paragraphStart, paragraphEnd) {
    if (linesOut.length >= options.maxLines) return;

    if (paragraphStart === paragraphEnd) {
      const line = this.#buildLine(options.runs, paragraphStart, paragraphEnd, options);
      linesOut.push(line);
      return;
    }

    if (options.wrapMode === "none" || !Number.isFinite(options.maxWidth)) {
      const line = this.#buildLine(options.runs, paragraphStart, paragraphEnd, options);
      linesOut.push(line);
      return;
    }

    /** @type {number[]} */
    let wordBreaksGlobal = [];
    /** @type {Set<number>} */
    let wordBreakSetGlobal = new Set();
    if (options.wrapMode === "word") {
      const paragraphText = options.text.slice(paragraphStart, paragraphEnd);
      const { breaks: wordBreaks, wordBreakSet } = wordBreakPositions(paragraphText, options.locale);

      // Convert breaks to full-text indices.
      wordBreaksGlobal = wordBreaks.map((b) => paragraphStart + b);
      wordBreakSetGlobal = new Set([...wordBreakSet].map((b) => paragraphStart + b));
    }

    let lineStart = paragraphStart;

    while (lineStart < paragraphEnd && linesOut.length < options.maxLines) {
      const remainingLines = options.maxLines - linesOut.length;
      const isLastAllowedLine = remainingLines === 1;

      if (isLastAllowedLine) {
        const truncatedLine = this.#layoutLastLineWithEllipsis(
          options,
          lineStart,
          paragraphEnd,
        );
        linesOut.push(truncatedLine);
        return;
      }

      const { lineEnd, usedWordBreak } = this.#findLineEnd({
        runs: options.runs,
        fullText: options.text,
        paragraphStart,
        paragraphEnd,
        lineStart,
        maxWidth: options.maxWidth,
        wrapMode: options.wrapMode,
        wordBreaksGlobal,
        wordBreakSetGlobal,
        locale: options.locale,
      });

      const line = this.#buildLine(options.runs, lineStart, lineEnd, options);
      linesOut.push(line);

      if (lineEnd >= paragraphEnd) return;

      if (usedWordBreak && wordBreakSetGlobal.has(lineEnd)) {
        lineStart = skipBreakableWhitespace(options.text, lineEnd, paragraphEnd);
      } else {
        lineStart = lineEnd;
      }
    }
  }

  /**
   * @param {ReturnType<TextLayoutEngine["#normalizeOptions"]>} options
   * @param {number} start
   * @param {number} end
   */
  #layoutLastLineWithEllipsis(options, start, end) {
    const ellipsis = options.ellipsis;
    const fullText = options.text;

    const initialMetrics = this.#measureRunsSlice(options.runs, start, end);
    if (initialMetrics.width <= options.maxWidth) {
      return this.#buildLine(options.runs, start, end, options);
    }

    // Find the largest prefix that fits, then append ellipsis.
    const segmentText = fullText.slice(start, end);
    const charBreaks = graphemeBreakPositions(segmentText, options.locale).map((b) => start + b);

    let low = 0;
    let high = charBreaks.length - 1;
    let bestIdx = -1;

    while (low <= high) {
      const mid = (low + high) >> 1;
      const candidateEnd = charBreaks[mid];
      const candidateMetrics = this.#measureRunsSliceWithSuffix(
        options.runs,
        start,
        candidateEnd,
        ellipsis,
        options.defaultFont,
      );
      if (candidateMetrics.width <= options.maxWidth) {
        bestIdx = mid;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    const finalEnd = bestIdx === -1 ? start : charBreaks[bestIdx];
    const lineRuns = this.#sliceRuns(options.runs, start, finalEnd);
    const ellipsisRun = { text: ellipsis, font: options.defaultFont };
    lineRuns.push(ellipsisRun);

    const metrics = this.#measureRuns(lineRuns);
    const x = this.#computeLineX(metrics.width, options.maxWidth, options.resolvedAlign);

    return {
      text: lineRuns.map((r) => r.text).join(""),
      runs: lineRuns,
      width: metrics.width,
      ascent: metrics.ascent,
      descent: metrics.descent,
      x,
    };
  }

  /**
   * @param {{runs: any[], fullText: string, paragraphStart: number, paragraphEnd: number, lineStart: number, maxWidth: number, wrapMode: string, wordBreaksGlobal: number[], wordBreakSetGlobal: Set<number>, locale: string | undefined}} params
   * @returns {{ lineEnd: number, usedWordBreak: boolean }}
   */
  #findLineEnd(params) {
    const {
      runs,
      fullText,
      paragraphEnd,
      lineStart,
      maxWidth,
      wrapMode,
      wordBreaksGlobal,
      locale,
    } = params;

    if (wrapMode === "char") {
      const segmentText = fullText.slice(lineStart, paragraphEnd);
      const charBreaks = graphemeBreakPositions(segmentText, locale).map((b) => lineStart + b);
      const { end: lineEnd } = this.#findFarthestFittingBreak(runs, lineStart, charBreaks, maxWidth);
      return { lineEnd, usedWordBreak: false };
    }

    // Word wrap (with char fallback).
    const candidateBreaks = wordBreaksGlobal.filter((b) => b > lineStart && b <= paragraphEnd);
    if (candidateBreaks.length === 0 || candidateBreaks[candidateBreaks.length - 1] !== paragraphEnd) {
      candidateBreaks.push(paragraphEnd);
    }

    const wordResult = this.#findFarthestFittingBreak(runs, lineStart, candidateBreaks, maxWidth);
    if (wordResult.fits && wordResult.end > lineStart) return { lineEnd: wordResult.end, usedWordBreak: true };

    // Fallback to char wrap so we always make progress.
    const segmentText = fullText.slice(lineStart, paragraphEnd);
    const charBreaks = graphemeBreakPositions(segmentText, locale).map((b) => lineStart + b);
    const { end: charEnd } = this.#findFarthestFittingBreak(runs, lineStart, charBreaks, maxWidth);
    return { lineEnd: charEnd, usedWordBreak: false };
  }

  /**
   * @param {any[]} runs
   * @param {number} start
   * @param {number[]} candidateEnds
   * @param {number} maxWidth
   * @returns {number}
   */
  #findFarthestFittingBreak(runs, start, candidateEnds, maxWidth) {
    let low = 0;
    let high = candidateEnds.length - 1;
    let best = -1;

    while (low <= high) {
      const mid = (low + high) >> 1;
      const end = candidateEnds[mid];
      const metrics = this.#measureRunsSlice(runs, start, end);
      if (metrics.width <= maxWidth) {
        best = mid;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    if (best >= 0) return { end: candidateEnds[best], fits: true };
    return { end: candidateEnds[0] ?? start, fits: false };
  }

  /**
   * @param {any[]} runs
   * @param {number} start
   * @param {number} end
   * @returns {{ width: number, ascent: number, descent: number }}
   */
  #measureRunsSlice(runs, start, end) {
    if (start === end) return { width: 0, ascent: 0, descent: 0 };
    const slice = this.#sliceRuns(runs, start, end);
    return this.#measureRuns(slice);
  }

  /**
   * @param {any[]} runs
   * @param {number} start
   * @param {number} end
   * @param {string} suffixText
   * @param {FontSpec} suffixFont
   */
  #measureRunsSliceWithSuffix(runs, start, end, suffixText, suffixFont) {
    const slice = this.#sliceRuns(runs, start, end);
    slice.push({ text: suffixText, font: suffixFont });
    return this.#measureRuns(slice);
  }

  /**
   * @param {any[]} lineRuns
   * @returns {{ width: number, ascent: number, descent: number }}
   */
  #measureRuns(lineRuns) {
    let width = 0;
    let ascent = 0;
    let descent = 0;

    /** @type {string | null} */
    let currentFontKey = null;
    /** @type {FontSpec | null} */
    let currentFont = null;
    let currentText = "";

    const flush = () => {
      if (!currentFont || !currentText) return;
      const m = this.#measureCached(currentText, currentFont);
      width += m.width;
      ascent = Math.max(ascent, m.ascent);
      descent = Math.max(descent, m.descent);
      currentText = "";
    };

    for (const run of lineRuns) {
      if (!run.text) continue;
      const fk = fontKey(run.font);
      if (currentFontKey === null) {
        currentFontKey = fk;
        currentFont = run.font;
        currentText = run.text;
      } else if (fk === currentFontKey) {
        currentText += run.text;
      } else {
        flush();
        currentFontKey = fk;
        currentFont = run.font;
        currentText = run.text;
      }
    }

    flush();
    return { width, ascent, descent };
  }

  /**
   * @param {any[]} runs
   * @param {number} start
   * @param {number} end
   * @returns {any[]}
   */
  #sliceRuns(runs, start, end) {
    /** @type {any[]} */
    const out = [];
    let offset = 0;

    for (const run of runs) {
      const runStart = offset;
      const runEnd = offset + run.text.length;
      offset = runEnd;

      if (runEnd <= start) continue;
      if (runStart >= end) break;

      const localStart = Math.max(0, start - runStart);
      const localEnd = Math.min(run.text.length, end - runStart);
      if (localEnd <= localStart) continue;

      out.push({ ...run, text: run.text.slice(localStart, localEnd) });
    }

    return out;
  }

  /**
   * @param {any[]} runs
   * @param {number} start
   * @param {number} end
   * @param {ReturnType<TextLayoutEngine["#normalizeOptions"]>} options
   * @returns {LayoutLine}
   */
  #buildLine(runs, start, end, options) {
    const lineRuns = this.#sliceRuns(runs, start, end);
    const text = lineRuns.map((r) => r.text).join("");
    const metrics = this.#measureRuns(lineRuns);
    const x = this.#computeLineX(metrics.width, options.maxWidth, options.resolvedAlign);

    return {
      text,
      runs: lineRuns,
      width: metrics.width,
      ascent: metrics.ascent,
      descent: metrics.descent,
      x,
    };
  }

  /**
   * @param {number} lineWidth
   * @param {number} maxWidth
   * @param {"left" | "right" | "center"} align
   */
  #computeLineX(lineWidth, maxWidth, align) {
    if (!Number.isFinite(maxWidth)) return 0;
    if (align === "center") return (maxWidth - lineWidth) / 2;
    if (align === "right") return maxWidth - lineWidth;
    return 0;
  }

  /**
   * @param {string} text
   * @returns {{ start: number, end: number }[]}
   */
  #splitParagraphs(text) {
    /** @type {{ start: number, end: number }[]} */
    const ranges = [];

    const newlineRe = /\r\n|\r|\n/g;
    let lastIndex = 0;
    let match;

    while ((match = newlineRe.exec(text))) {
      ranges.push({ start: lastIndex, end: match.index });
      lastIndex = match.index + match[0].length;
    }

    ranges.push({ start: lastIndex, end: text.length });
    return ranges;
  }
}
