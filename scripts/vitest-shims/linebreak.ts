// Vitest shim for the `linebreak` dependency used by `@formula/text-layout`.
//
// The real `linebreak` package implements UAX #14 (Unicode Line Breaking Algorithm). For Vitest
// suites that don't validate the full text wrapping behavior, we only need a lightweight
// implementation that:
//   - exposes a default-exported class `LineBreaker`
//   - supports `new LineBreaker(text)` and `.nextBreak()`
//   - returns objects with `{ position: number }` or `null` when exhausted
//
// Break positions are derived from `Intl.Segmenter` word boundaries when available, falling back
// to per-code-unit breaks.

type BreakResult = { position: number };

function computeBreakPositions(text: string): number[] {
  if (!text) return [];

  /** @type {number[]} */
  const breaks: number[] = [];

  if (typeof Intl !== "undefined" && "Segmenter" in Intl) {
    const segmenter = new Intl.Segmenter(undefined, { granularity: "word" });
    for (const part of segmenter.segment(text)) {
      const end = part.index + part.segment.length;
      if (end > 0 && end <= text.length) breaks.push(end);
    }
  } else {
    for (let i = 1; i <= text.length; i++) breaks.push(i);
  }

  if (breaks.length === 0 || breaks[breaks.length - 1] !== text.length) breaks.push(text.length);

  // Dedupe + sort to be safe.
  const unique = Array.from(new Set(breaks));
  unique.sort((a, b) => a - b);
  return unique;
}

export default class LineBreaker {
  private readonly breaks: number[];
  private index = 0;

  constructor(text: string) {
    this.breaks = computeBreakPositions(text);
  }

  nextBreak(): BreakResult | null {
    if (this.index >= this.breaks.length) return null;
    const position = this.breaks[this.index++]!;
    return { position };
  }
}

