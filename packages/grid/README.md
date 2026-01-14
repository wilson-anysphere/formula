# `@formula/grid`

Canvas-based virtualized spreadsheet grid renderer.

## Dev harness

For manual verification while developing:

```bash
pnpm --dir packages/grid dev
```

Then open one of:

- `http://localhost:5173/?demo=style` — cell formatting demo (Excel-style rendering)
- `http://localhost:5173/?demo=merged` — merged cells + text overflow demo
- `http://localhost:5173/?demo=perf` — performance harness

## Rich text (inline runs)

Cells can optionally provide a `richText` payload to render inline formatting (bold/italic/underline/color/etc) within a single cell.

```ts
import type { CellData } from "@formula/grid";

const cell: CellData = {
  row: 0,
  col: 0,
  // `value` remains the plain-text fallback (used for a11y/status strings).
  value: "Hello world",
  richText: {
    text: "Hello world",
    runs: [
      { start: 0, end: 5, style: { bold: true } },
      { start: 5, end: 11, style: { italic: true, underline: "single", color: "#FFEF4444", size_100pt: 1200 } }
    ]
  }
};
```

Notes:

- `runs[].start/end` are **Unicode code point indexes** (not UTF-16 offsets).
- Runs do **not** have to cover the entire string; the renderer fills gaps with default/cell-level styling.
- Supported per-run style keys:
  - `bold?: boolean`
  - `italic?: boolean`
  - `underline?: string | boolean`
    - `true` / `"single"`: single underline
    - `"double"` / `"doubleAccounting"`: double underline (Excel-like)
    - `"none"` / `false`: no underline
    - Any other truthy string is treated as single underline
  - `strike?: boolean` / `strikethrough?: boolean`
  - `color?: string` (engine colors are serialized as `#AARRGGBB` and are converted to canvas `rgba(...)`)
  - `font?: string`
  - `size_100pt?: number` (font size in 1/100 points; converted at 96DPI)

Cell-level styling notes:

- `CellStyle.underlineStyle?: "single" | "double"` can be used to request a double underline without rich text.
- `CellStyle.strike?: boolean` enables strike-through for the whole cell.

## Borders

Cells can render Excel-style borders (on top of default gridlines) via `CellStyle.borders`:

```ts
import type { CellData } from "@formula/grid";

const cell: CellData = {
  row: 0,
  col: 0,
  value: "Bordered",
  style: {
    borders: {
      top: { width: 1, style: "solid", color: "#0f172a" },
      right: { width: 1, style: "dashed", color: "#ef4444" },
      bottom: { width: 2, style: "double", color: "#a855f7" },
      left: { width: 1, style: "dotted", color: "#3b82f6" }
    }
  }
};
```

Notes:

- Border widths are in **CSS pixels at zoom=1**. The renderer scales them by the current zoom.
- Supported styles: `"solid" | "dashed" | "dotted" | "double"`.
- When two adjacent cells specify borders on the same shared edge, the renderer deterministically picks the winner:
  - larger effective width wins, then
  - style rank (`double` > `solid` > `dashed` > `dotted`), then
  - right/bottom cell wins (stable tie-breaker).
- Borders inside merged regions are suppressed; merged range borders use the anchor cell’s border specs.

### Diagonal borders

You can draw diagonal borders (Excel-style) via `CellStyle.diagonalBorders`:

```ts
const cell: CellData = {
  row: 0,
  col: 0,
  value: "X",
  style: {
    diagonalBorders: {
      // Top-left → bottom-right
      down: { width: 1, style: "solid", color: "#0f172a" },
      // Bottom-left → top-right
      up: { width: 1, style: "dotted", color: "#0f172a" }
    }
  }
};
```

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

// Optional: treat the first row/col as headers for styling.
// When unset, the renderer uses legacy behavior and treats the first frozen
// row/col (if any) as the header region.
renderer.setHeaders(1, 1);
```

If you build custom scrollbars or overlays that depend on viewport metrics (total size, frozen extents, max scroll),
you can subscribe to *layout* viewport changes (axis sizes, frozen panes, resize, zoom) via:

```ts
const unsubscribe = renderer.subscribeViewport(
  ({ viewport, reason }) => {
    // Recompute scrollbar thumbs, overlay geometry, etc.
    // Note: this does NOT fire on scroll offset changes (to avoid per-frame work during scroll).
    console.log(reason, viewport.maxScrollX, viewport.maxScrollY);
  },
  { animationFrame: true } // throttle to 1 callback per frame
);
```

## Accessibility

The grid is canvas-rendered, but includes baseline accessibility scaffolding:

- Focusable container (`tabIndex=0`) with `role="grid"` and a default accessible name (`"Spreadsheet grid"`).
- Canvases are `aria-hidden`.
- The active cell is also exposed via an offscreen `role="gridcell"` element wired up with `aria-activedescendant` (including `aria-rowindex`/`aria-colindex`).
- A visually-hidden `role="status"` live region announces:
  - active cell address (A1-style when headers are enabled via `headerRows/headerCols`)
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

- `getFillHandleRect(): { x; y; width; height } | null` (viewport coordinates relative to the grid canvases, clipped to the visible viewport. Returns `null` when the handle is not visible (e.g. offscreen, behind frozen rows/cols, or when `interactionMode !== "default"`).)

Note: `@formula/grid` only handles interaction + rendering. Consumers are responsible for actually writing filled values/formulas into their backing store.

## Selection + keyboard navigation

When the grid container is focused, `CanvasGrid` supports spreadsheet-like navigation and selection:

- Arrow keys move the active cell.
- Shift+arrows extends the active selection range.
- Ctrl/Cmd+arrows jump to the first/last row/col (data region, excluding header rows/cols when `headerRows/headerCols` are configured).
- PageUp/PageDown (and Alt+PageUp/PageDown for horizontal paging).
- Home/End (+Ctrl/Cmd for absolute edges).
- Tab/Enter move the active cell; when a multi-cell range is selected, Tab/Enter move *within* the selection range (wrapping) instead of collapsing it.
- Ctrl/Cmd+A selects all; Ctrl/Cmd+Space selects a column; Shift+Space selects a row.
- Header pointer selection (when headers are enabled): click corner header selects all; click row/col headers select entire row/column.
- Ctrl/Cmd+click adds a new selection range (multi-range selection); Shift+click/drag extends the active range without clearing others.
- Header rows/cols (as defined by `headerRows`/`headerCols`) are styled using the `headerBg`/`headerText` theme tokens.

## Resizing + auto-fit (double-click)

When `enableResize` is enabled, `CanvasGrid` supports Excel-like row/column resizing:

- **Drag** a row/column header boundary to resize.
- **Double-click** (mouse) / **double-tap** (touch) a row/column header boundary to **auto-fit** to content.

To persist user-driven size changes (including auto-fit), provide:

- `onAxisSizeChange?: (change: GridAxisSizeChange) => void`

`GridAxisSizeChange.size` is reported in **CSS pixels at the current zoom** (matching `GridApi.setRowHeight` / `setColWidth`). If you persist sizes as zoom-independent “base” sizes, divide by `change.zoom` before storing.

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
  - `applyAxisSizeOverrides({ rows?, cols? }, { resetUnspecified? })` applies many row/column size overrides at once (single redraw).
  - Multi-range selection helpers: `setSelectionRanges`, `getSelectionRanges`, `getActiveSelectionRangeIndex`.

## Overlays / viewport helpers

Overlays (diffs, auditing, presence, etc.) often need to convert a sheet `CellRange` into pixel geometry in the current viewport.

`CanvasGridRenderer.getRangeRects(range)` returns an array of viewport-space `Rect` objects (relative to the grid canvases), clipped to the visible viewport and split across frozen-row/column quadrants when needed.

```ts
import type { CellRange, Rect } from "@formula/grid";

const range: CellRange = { startRow: 0, endRow: 10, startCol: 0, endCol: 5 };
const rects: Rect[] = renderer.getRangeRects(range);
```

To position DOM overlays in page coordinates, add the grid element’s `getBoundingClientRect().left/top` to each rect.

## Zoom

`CanvasGrid` supports scaling the grid UI (cell sizes + text rendering) via a zoom factor.

- **Imperative:** use `GridApi.setZoom(zoom)` / `getZoom()`.
- **Controlled prop:** pass `zoom?: number` to `CanvasGrid` (optional; when provided, the grid treats zoom as controlled).
  - Use `onZoomChange?: (zoom) => void` to respond to user gestures (pinch / ctrl+wheel) in controlled mode.
- Zoom is clamped to the range **0.25–4.0**.

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
- Keyboard navigation and `scrollToCell` treat merged ranges as a single cell (jumping over interior merged cells). If a merge crosses frozen boundaries, `scrollToCell` will still try to reveal the scrollable portion of the merge.
