# `@formula/text-layout`
Canvas-friendly text layout for grid rendering.

## Goals
- Accurate wrapping/measurement for multilingual text (including RTL scripts).
- Avoid per-frame `CanvasRenderingContext2D.measureText()` calls via caching.
- Provide a layout API that can be reused by both on-screen canvas rendering and print/PDF rendering by swapping the measurement backend.

## Measurement strategies

### Canvas (platform-dependent)
`CanvasTextMeasurer` delegates shaping/measurement to the platform (browser canvas / Skia). This is fast and easy to integrate for on-screen rendering, but glyph advances can vary slightly across OS/browser/font versions.

### HarfBuzz (deterministic)
`HarfBuzzTextMeasurer` uses HarfBuzz (WASM) + explicitly loaded font bytes to produce deterministic advances across platforms. This is the recommended backend for Excel-fidelity layout and consistent print/PDF output.

## API overview
```js
import { TextLayoutEngine } from "@formula/text-layout";

const engine = new TextLayoutEngine(measurer);
const layout = engine.layout({
  text: "שלום world",
  font: { family: "Inter", sizePx: 13, weight: 400 },
  maxWidth: 120,
  wrapMode: "word",
  align: "start",
  direction: "auto"
});
```

## Using HarfBuzz

HarfBuzz bindings live under the `@formula/text-layout/harfbuzz` entrypoint so browser bundles that only use
canvas-based measurement don't pull in the HarfBuzz WASM + Node compatibility helpers.

```js
import {
  TextLayoutEngine,
} from "@formula/text-layout";
import { createHarfBuzzTextMeasurer } from "@formula/text-layout/harfbuzz";

// Load font bytes however you want (fetch, fs, bundler asset, etc).
const notoSansBytes = await fetch("/fonts/NotoSans-Regular.ttf").then((r) => r.arrayBuffer());
const notoSansHebrewBytes = await fetch("/fonts/NotoSansHebrew-Regular.ttf").then((r) => r.arrayBuffer());

const measurer = await createHarfBuzzTextMeasurer({
  fonts: [
    { family: "Noto Sans", weight: 400, style: "normal", data: notoSansBytes },
    { family: "Noto Sans Hebrew", weight: 400, style: "normal", data: notoSansHebrewBytes },
  ],
  // Optional global fallback order when glyphs are missing in the requested font.
  fallbackFamilies: ["Noto Sans Hebrew", "Noto Sans"],
});

const engine = new TextLayoutEngine(measurer);
```

## Caching
- Text metrics are cached by `(measurerKey, fontKey, text)`.
- Full layout results are cached by `(measurerKey, text/runs, font(s), width, wrapMode, lineHeight, maxLines, ellipsis, direction, align)`.

If a `TextMeasurer` exposes `cacheKey`, the engine will automatically include it in its cache keys. This is used by the HarfBuzz backend to invalidate cached measurements/layouts when fonts or fallback settings change.

## Segmentation & line breaking
- `wrapMode: "char"` wraps at grapheme cluster boundaries (UAX #29).
- `wrapMode: "word"` uses the Unicode Line Breaking Algorithm (UAX #14) and trims breakable whitespace at line boundaries.

Consumers should keep a single `TextLayoutEngine` instance alive for the lifetime of the renderer.
