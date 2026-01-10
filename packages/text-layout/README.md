# `@formula/text-layout`
Canvas-friendly text layout for grid rendering.

## Goals
- Accurate wrapping/measurement for multilingual text (including RTL scripts).
- Avoid per-frame `CanvasRenderingContext2D.measureText()` calls via caching.
- Provide a layout API that can be reused by both on-screen canvas rendering and print/PDF rendering by swapping the measurement backend.

## MVP strategy (current)
This package relies on the platform text shaper (browser canvas / Skia) for shaping, ligatures, combining marks, bidi, etc. Layout is computed by repeatedly measuring candidate substrings and choosing line breaks that fit the requested width. Results are cached so repeated renders of the same cell value are fast.

### Determinism trade-off
Relying on the platform shaper means glyph advances can vary slightly across OS/browser/font versions. This is generally acceptable for on-screen rendering but can cause subtle cross-platform differences.

## Advanced strategy (future)
For deterministic cross-platform layout (and consistent print/PDF output), integrate a HarfBuzz-based shaper (WASM in the browser or native in the desktop app). The layout engine is structured around a `TextMeasurer` interface so a HarfBuzz-backed measurer can be introduced without changing call sites.

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

## Caching
- Text metrics are cached by `(fontKey, text)`.
- Full layout results are cached by `(text/runs, font(s), width, wrapMode, lineHeight, maxLines, ellipsis, direction, align)`.

Consumers should keep a single `TextLayoutEngine` instance alive for the lifetime of the renderer.

