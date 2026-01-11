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

## Autofill (fill handle)

When `interactionMode="default"`, the selection overlay draws a small **fill handle** at the bottom-right of the active selection range (Excel-like).

Dragging the fill handle:

- Extends the active selection range in the drag direction.
- Renders a dashed preview overlay while dragging.
- Fires callbacks so consumers can apply an autofill algorithm to their underlying data model:
  - `onFillHandleChange?: ({ source, target }) => void`
  - `onFillHandleCommit?: ({ source, target }) => void | Promise<void`

`target` is the full extended range **including** `source`.

If you need to compute pixel-accurate drag start points (for custom UI or tests), the imperative `GridApi` exposes:

- `getFillHandleRect(): { x; y; width; height } | null` (viewport coordinates relative to the grid canvases)

Note: `@formula/grid` only handles interaction + rendering. Consumers are responsible for actually writing filled values/formulas into their backing store.

## Selection + keyboard navigation

When the grid container is focused, `CanvasGrid` supports spreadsheet-like navigation and selection:

- Arrow keys move the active cell.
- Shift+arrows extends the active selection range.
- Ctrl/Cmd+arrows jump to the first/last row/col (data region, excluding header row/col when `frozenRows/frozenCols` are used as headers).
- PageUp/PageDown (and Alt+PageUp/PageDown for horizontal paging).
- Home/End (+Ctrl/Cmd for absolute edges).
- Tab/Enter move the active cell; when a multi-cell range is selected, Tab/Enter move *within* the selection range (wrapping) instead of collapsing it.
- Ctrl/Cmd+A selects all; Ctrl/Cmd+Space selects a column; Shift+Space selects a row.
- Header pointer selection (when headers are enabled): click corner header selects all; click row/col headers select entire row/column.
- Ctrl/Cmd+click adds a new selection range (multi-range selection); Shift+click/drag extends the active range without clearing others.

## `GridApi` (React)

Pass `apiRef` to `CanvasGrid` to obtain an imperative API:

```tsx
const apiRef = useRef<GridApi | null>(null);
<CanvasGrid apiRef={(api) => (apiRef.current = api)} {...props} />;
```

Notable helpers:

- `scrollToCell(row, col, { align, padding })` keeps a cell in view.
- `getCellRect(row, col)` returns a viewport-space rect.
  - Coordinates are relative to the top-left of the grid viewport (the canvases). To position an overlay in page coordinates, add the grid element’s `getBoundingClientRect().left/top`.
  - If the cell is part of a merged range, this returns the merged bounds when possible.
  - If a merge crosses frozen boundaries (frozen vs scrollable quadrants), the merged bounds cannot be represented as a single viewport rect; in that case this falls back to the merged anchor cell rect.
- `getViewportState()` returns the current scroll/viewport metrics (useful for overlay positioning without extra DOM reads).
- `getFillHandleRect()` returns the active selection fill-handle rect (also in viewport coordinates).
- `setZoom(zoom)` / `getZoom()` control the grid zoom level (scales cell sizes + text rendering).
- Multi-range selection helpers: `setSelectionRanges`, `getSelectionRanges`, `getActiveSelectionRangeIndex`.

## Zoom

`CanvasGrid` supports scaling the grid UI (cell sizes + text rendering) via a zoom factor.

- **Imperative:** use `GridApi.setZoom(zoom)` / `getZoom()`.
- **Controlled prop:** pass `zoom?: number` to `CanvasGrid` (optional; when provided, the grid treats zoom as controlled).
  - Use `onZoomChange?: (zoom) => void` to respond to user gestures (pinch / ctrl+wheel) in controlled mode.

User interactions:

- **Touch devices:** two-finger pinch-to-zoom.
- **Trackpads / browsers:** `ctrl+wheel` (trackpad pinch often surfaces as `ctrl+wheel`).

## Merged cells

To enable merged-cell rendering and interactions, implement one of these optional `CellProvider` methods:

- `getMergedRangeAt(row, col) -> MergedCellRange | null`
- `getMergedRangesInRange(range) -> MergedCellRange[]` (bulk API; recommended for performance)

Notes:

- Ranges use **exclusive end** coordinates (`endRow/endCol`).
- The merged “anchor” is always the top-left cell (`startRow/startCol`).
- The renderer only draws text for anchor cells and suppresses interior gridlines.
- Keyboard navigation and `scrollToCell` treat merged ranges as a single cell (jumping over interior merged cells).
