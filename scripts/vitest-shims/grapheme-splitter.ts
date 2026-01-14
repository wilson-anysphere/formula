// Vitest shim for the `grapheme-splitter` dependency used by `@formula/text-layout`.
//
// In some CI/dev environments, tests may run with cached/stale `node_modules` where workspace
// packages are resolved via `vitest.config.ts` aliases but transitive deps are missing.
//
// This shim provides the minimal API surface used by `@formula/text-layout`:
//   - `new GraphemeSplitter()`
//   - `.splitGraphemes(text) => string[]`
//
// We implement this using `Intl.Segmenter` when available (Node 16+/modern browsers), falling back
// to code-point splitting.
export default class GraphemeSplitter {
  private readonly segmenter: Intl.Segmenter | null;

  constructor() {
    this.segmenter =
      typeof Intl !== "undefined" && "Segmenter" in Intl
        ? new Intl.Segmenter(undefined, { granularity: "grapheme" })
        : null;
  }

  splitGraphemes(text: string): string[] {
    if (!text) return [];
    if (this.segmenter) {
      return Array.from(this.segmenter.segment(text), (seg) => seg.segment);
    }
    // Fallback: split by codepoints.
    return Array.from(text);
  }
}

