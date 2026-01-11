# `@formula/grid`

Canvas-based virtualized spreadsheet grid renderer.

## Theming

`CanvasGrid` / `CanvasGridRenderer` use a `GridTheme` token set (no hard-coded UI colors). You can theme the grid in two ways:

1. **CSS variables** on the grid container (recommended for app-wide theming).
2. **Explicit theme overrides** via `theme?: Partial<GridTheme>` (React) or `theme?: Partial<GridTheme>` in the `CanvasGridRenderer` constructor / `setTheme()`.

### CSS variables

The grid reads the following CSS variables (see `GRID_THEME_CSS_VAR_NAMES`):

- `--formula-grid-bg`
- `--formula-grid-line`
- `--formula-grid-header-bg`
- `--formula-grid-header-text`
- `--formula-grid-cell-text`
- `--formula-grid-error-text`
- `--formula-grid-selection-fill`
- `--formula-grid-selection-border`
- `--formula-grid-selection-handle`
- `--formula-grid-scrollbar-track`
- `--formula-grid-scrollbar-thumb`
- `--formula-grid-freeze-line`
- `--formula-grid-comment-indicator`
- `--formula-grid-comment-indicator-resolved`
- `--formula-grid-remote-presence-default`

Example:

```css
.myGridTheme {
  --formula-grid-bg: #ffffff;
  --formula-grid-line: #e6e6e6;
  --formula-grid-selection-border: #0e65eb;
}

@media (prefers-color-scheme: dark) {
  .myGridTheme {
    --formula-grid-bg: #0b1220;
    --formula-grid-line: rgba(255, 255, 255, 0.12);
  }
}

@media (prefers-contrast: more) {
  .myGridTheme {
    --formula-grid-bg: Canvas;
    --formula-grid-line: CanvasText;
    --formula-grid-selection-border: Highlight;
  }
}
```

Note: `CanvasGrid` resolves nested `var(...)` references (e.g. `--formula-grid-bg: var(--app-bg)`), and normalizes system colors (`Canvas`, `Highlight`, etc.) into computed `rgb(...)` strings before passing them to the canvas renderer.

If you build custom theme plumbing, `@formula/grid` also exports `resolveCssVarValue()` (best-effort resolver for simple `var(--token, fallback)` chains).

### Non-React usage

You can use the renderer directly:

```ts
import { CanvasGridRenderer, resolveGridThemeFromCssVars } from "@formula/grid";

const renderer = new CanvasGridRenderer({ provider, rowCount, colCount });
renderer.attach({ grid, content, selection });

// Theme from CSS vars on an element:
renderer.setTheme(resolveGridThemeFromCssVars(containerEl));
```

## Accessibility

The grid is canvas-rendered, but includes baseline accessibility scaffolding:

- Focusable container (`tabIndex=0`) with `role="grid"` and a default accessible name (`"Spreadsheet grid"`).
- Canvases are `aria-hidden`.
- The active cell is also exposed via an offscreen `role="gridcell"` element wired up with `aria-activedescendant` (including `aria-rowindex`/`aria-colindex`).
- A visually-hidden `role="status"` live region announces:
  - active cell address (A1-style when headers are enabled via `frozenRows/frozenCols`)
  - active cell value
  - active selection range
- Keyboard navigation is supported via arrow keys when the grid container is focused.
