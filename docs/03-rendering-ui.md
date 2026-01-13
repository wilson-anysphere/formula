# Rendering & UI Architecture

## Overview

The rendering layer must achieve 60fps scrolling with millions of rows while maintaining full visual fidelity. This is only possible with **Canvas-based rendering** and aggressive **virtualization**. DOM-based grids cannot scale.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  RENDER LOOP (requestAnimationFrame)                                        │
│  ├── Process input events                                                   │
│  ├── Update viewport state                                                  │
│  ├── Render dirty regions                                                   │
│  └── Update overlay elements                                                │
├─────────────────────────────────────────────────────────────────────────────┤
│  LAYERS (bottom to top)                                                     │
│  ├── Grid Canvas: Cell backgrounds, borders, gridlines                      │
│  ├── Content Canvas: Text, numbers, formulas, images                        │
│  ├── Selection Canvas: Selection highlight, fill handle                     │
│  └── DOM Overlays: Cell editor, context menus, tooltips, dropdowns         │
├─────────────────────────────────────────────────────────────────────────────┤
│  VIEWPORT MANAGER                                                           │
│  ├── Scroll position (virtual coordinates)                                  │
│  ├── Visible range calculation                                              │
│  ├── Row/column size cache                                                  │
│  └── Frozen rows/columns handling                                           │
├─────────────────────────────────────────────────────────────────────────────┤
│  DATA INTERFACE                                                             │
│  ├── Request visible cells from engine                                      │
│  ├── Format values for display                                              │
│  └── Cache rendered cell content                                            │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Canvas-Based Grid Rendering

### Why Canvas?

| Approach | Cells Limit | Memory | Scroll FPS |
|----------|-------------|--------|------------|
| DOM (one div per cell) | ~10K | High | 20-30fps |
| DOM Virtualized | ~100K | Medium | 40-50fps |
| Canvas Virtualized | 1M+ | Low | 60fps |

Google Sheets uses Canvas for the entire grid. DOM elements only appear for:
- Text input (cell editor)
- Selection overlays
- Menus and dialogs

### Canvas Setup

```typescript
class GridRenderer {
  private gridCanvas: HTMLCanvasElement;
  private gridCtx: CanvasRenderingContext2D;
  private contentCanvas: HTMLCanvasElement;
  private contentCtx: CanvasRenderingContext2D;
  private selectionCanvas: HTMLCanvasElement;
  private selectionCtx: CanvasRenderingContext2D;
  
  constructor(container: HTMLElement) {
    // Create layered canvases
    this.gridCanvas = this.createCanvas(container, 0);
    this.contentCanvas = this.createCanvas(container, 1);
    this.selectionCanvas = this.createCanvas(container, 2);
    
    // Get contexts
    this.gridCtx = this.gridCanvas.getContext("2d", { alpha: false })!;
    this.contentCtx = this.contentCanvas.getContext("2d")!;
    this.selectionCtx = this.selectionCanvas.getContext("2d")!;
    
    // Handle device pixel ratio
    this.setupHiDPI();
  }
  
  private setupHiDPI(): void {
    const dpr = window.devicePixelRatio || 1;
    const rect = this.gridCanvas.getBoundingClientRect();
    
    for (const canvas of [this.gridCanvas, this.contentCanvas, this.selectionCanvas]) {
      canvas.width = rect.width * dpr;
      canvas.height = rect.height * dpr;
      canvas.style.width = `${rect.width}px`;
      canvas.style.height = `${rect.height}px`;
      
      const ctx = canvas.getContext("2d")!;
      ctx.scale(dpr, dpr);
    }
  }
}
```

### Coordinate Systems

```
┌────────────────────────────────────────────────────────────────┐
│  COORDINATE SPACES                                              │
├────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Data Space (row, col)     Scroll Space (scrollX, scrollY)     │
│  ├── Row: 0 to ∞           ├── Pixels from origin               │
│  └── Col: 0 to ∞           └── Limited by browser (~33M px)     │
│         │                            │                          │
│         ▼                            ▼                          │
│  ┌─────────────────┐       ┌─────────────────┐                 │
│  │ Position Cache  │──────▶│ Virtual Scroll  │                 │
│  │ (row→y, col→x)  │       │    Manager      │                 │
│  └─────────────────┘       └─────────────────┘                 │
│         │                            │                          │
│         ▼                            ▼                          │
│  Canvas Space (x, y)        Screen Space (clientX, clientY)    │
│  ├── Relative to canvas    ├── Relative to viewport            │
│  └── After DPR scaling     └── Mouse/touch input               │
│                                                                 │
└────────────────────────────────────────────────────────────────┘
```

---

## Virtualization

### The Browser Scroll Limit Problem

Browsers limit scroll containers to approximately **33 million pixels** (varies by browser):
- Chrome/Safari: ~33,554,432 px (2^25)
- Firefox: ~17,895,697 px

At 30px row height: ~1.1 million rows max with native scrolling.

### Virtual Scrolling Solution

Instead of a giant scrollable container, we:
1. Keep a fixed-size viewport
2. Translate scroll position to data indices
3. Render only visible cells
4. Fake scrollbar with calculated thumb position

```typescript
interface VirtualScrollState {
  // Data indices (unlimited)
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  
  // Pixel offsets within first visible cell
  offsetX: number;
  offsetY: number;
  
  // Total dimensions (for scrollbar thumb sizing)
  totalRows: number;
  totalCols: number;
  estimatedTotalHeight: number;
  estimatedTotalWidth: number;
}

class VirtualScrollManager {
  private rowHeights: Map<number, number> = new Map();
  private colWidths: Map<number, number> = new Map();
  private defaultRowHeight = 21;
  private defaultColWidth = 100;
  
  // Cumulative position cache for fast lookup
  private rowPositions: number[] = [];
  private colPositions: number[] = [];
  
  scrollToPixel(scrollY: number, viewportHeight: number): VirtualScrollState {
    // Binary search for start row
    const startRow = this.findRowAtPosition(scrollY);
    const offsetY = scrollY - this.getRowPosition(startRow);
    
    // Find end row
    let endRow = startRow;
    let accumulatedHeight = -offsetY;
    while (accumulatedHeight < viewportHeight && endRow < this.totalRows) {
      accumulatedHeight += this.getRowHeight(endRow);
      endRow++;
    }
    
    return {
      startRow,
      endRow,
      offsetY,
      // ... similar for columns
    };
  }
  
  private findRowAtPosition(y: number): number {
    // Binary search through rowPositions
    let low = 0, high = this.rowPositions.length - 1;
    while (low < high) {
      const mid = Math.floor((low + high + 1) / 2);
      if (this.rowPositions[mid] <= y) {
        low = mid;
      } else {
        high = mid - 1;
      }
    }
    return low;
  }
}
```

### Custom Scrollbar Implementation

```typescript
class CustomScrollbar {
  private thumb: HTMLDivElement;
  private track: HTMLDivElement;
  private isDragging = false;
  
  updateThumb(scrollPos: number, viewportSize: number, contentSize: number): void {
    const trackSize = this.track.getBoundingClientRect().height;
    const thumbSize = Math.max(30, (viewportSize / contentSize) * trackSize);
    const thumbPos = (scrollPos / (contentSize - viewportSize)) * (trackSize - thumbSize);
    
    this.thumb.style.height = `${thumbSize}px`;
    this.thumb.style.transform = `translateY(${thumbPos}px)`;
  }
  
  setupDrag(): void {
    this.thumb.addEventListener("mousedown", (e) => {
      this.isDragging = true;
      const startY = e.clientY;
      const startScroll = this.scrollPos;
      
      const onMove = (e: MouseEvent) => {
        const deltaY = e.clientY - startY;
        const scrollDelta = deltaY * (this.contentSize / this.track.offsetHeight);
        this.onScroll(startScroll + scrollDelta);
      };
      
      const onUp = () => {
        this.isDragging = false;
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
      };
      
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });
  }
}
```

---

## Cell Rendering

### Render Pipeline

```typescript
class CellRenderer {
  render(
    ctx: CanvasRenderingContext2D,
    cell: CellData,
    bounds: Rect,
    style: CellStyle
  ): void {
    // 1. Background
    if (style.fill) {
      ctx.fillStyle = style.fill;
      ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
    }
    
    // 2. Conditional formatting overlays
    if (cell.conditionalFormat) {
      this.renderConditionalFormat(ctx, cell.conditionalFormat, bounds);
    }
    
    // 3. Content
    const displayValue = this.formatValue(cell.value, style.numberFormat);
    this.renderText(ctx, displayValue, bounds, style);
    
    // 4. Borders (rendered separately for efficiency)
    // Borders are batched and drawn in a single pass
    
    // 5. Icons (data validation dropdown, comment indicator, etc.)
    if (cell.hasComment) {
      this.renderCommentIndicator(ctx, bounds);
    }
    if (cell.hasDropdown) {
      this.renderDropdownArrow(ctx, bounds);
    }
  }
  
  private renderText(
    ctx: CanvasRenderingContext2D,
    text: string,
    bounds: Rect,
    style: CellStyle
  ): void {
    // Use the shared text layout engine to avoid per-frame `measureText` calls and
    // to ensure wrapping works for complex scripts (RTL, combining marks, ligatures).
    const engine = getSharedTextLayoutEngine();
    const font = {
      family: style.fontFamily,
      sizePx: style.fontSize,
      weight: style.fontWeight,
    };
    
    const layout = engine.layout({
      text,
      font,
      maxWidth: bounds.width - 8,
      wrapMode: style.wrap ? "word" : "none",
      align: style.horizontalAlign === "left"
        ? "start"
        : style.horizontalAlign === "right"
          ? "end"
          : "center",
      direction: "auto",
      lineHeightPx: Math.ceil(style.fontSize * 1.2),
    });
    
    // Vertical centering.
    const originX = bounds.x + 4;
    const originY = bounds.y + (bounds.height - layout.height) / 2;
    
    ctx.save();
    ctx.fillStyle = style.color || "#000000";
    ctx.beginPath();
    ctx.rect(bounds.x, bounds.y, bounds.width, bounds.height);
    ctx.clip();
    
    drawTextLayout(ctx, layout, originX, originY);
    ctx.restore();
  }
}
```

### Shared-grid axis sizing (including Hide/Unhide)

In **shared-grid mode**, row heights and column widths are driven by *sheet view metadata* and applied to the renderer as **axis size overrides** (batched, not per-index setters). This same mechanism is used to support **Hide / Unhide**:

- **Hide**: apply an override that collapses the target row/column (typically to a minimal size)
- **Unhide**: remove/restore the override so the axis returns to its prior/default size

This keeps the canvas renderer, scroll model, and any secondary panes in sync without needing legacy outline/visibility caches.

> Note: Excel-style **outline grouping controls** (Data → Outline: Group/Ungroup/Show Detail/Hide Detail) may still be implemented only in the legacy renderer even when basic Hide/Unhide is available in shared-grid mode.

### Batched Drawing

Minimize context state changes by batching similar operations:

```typescript
class BatchedRenderer {
  private fillBatches: Map<string, Rect[]> = new Map();
  private strokeBatches: Map<string, Line[]> = new Map();
  
  queueFill(color: string, rect: Rect): void {
    if (!this.fillBatches.has(color)) {
      this.fillBatches.set(color, []);
    }
    this.fillBatches.get(color)!.push(rect);
  }
  
  queueStroke(color: string, line: Line): void {
    if (!this.strokeBatches.has(color)) {
      this.strokeBatches.set(color, []);
    }
    this.strokeBatches.get(color)!.push(line);
  }
  
  flush(ctx: CanvasRenderingContext2D): void {
    // Batch fills by color
    for (const [color, rects] of this.fillBatches) {
      ctx.fillStyle = color;
      ctx.beginPath();
      for (const rect of rects) {
        ctx.rect(rect.x, rect.y, rect.width, rect.height);
      }
      ctx.fill();
    }
    
    // Batch strokes by color
    for (const [color, lines] of this.strokeBatches) {
      ctx.strokeStyle = color;
      ctx.beginPath();
      for (const line of lines) {
        ctx.moveTo(line.x1, line.y1);
        ctx.lineTo(line.x2, line.y2);
      }
      ctx.stroke();
    }
    
    this.fillBatches.clear();
    this.strokeBatches.clear();
  }
}
```

---

## Selection Rendering

### Selection States

```typescript
type SelectionType =
  | "cell"           // Single cell
  | "range"          // Contiguous range
  | "multi"          // Multiple non-contiguous selections (Ctrl+click)
  | "column"         // Entire column(s)
  | "row"            // Entire row(s)
  | "all";           // Entire sheet

interface Selection {
  type: SelectionType;
  ranges: Range[];
  activeCell: CellRef;
  anchor: CellRef;  // For shift+click expansion
}
```

### Selection Rendering

```typescript
class SelectionRenderer {
  render(ctx: CanvasRenderingContext2D, selection: Selection, viewport: Viewport): void {
    // 1. Fill selected ranges with semi-transparent highlight
    ctx.fillStyle = "rgba(14, 101, 235, 0.1)";
    for (const range of selection.ranges) {
      const bounds = this.rangeToBounds(range, viewport);
      if (bounds) {
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
      }
    }
    
    // 2. Draw range border
    ctx.strokeStyle = "#0e65eb";
    ctx.lineWidth = 2;
    for (const range of selection.ranges) {
      const bounds = this.rangeToBounds(range, viewport);
      if (bounds) {
        ctx.strokeRect(bounds.x + 1, bounds.y + 1, bounds.width - 2, bounds.height - 2);
      }
    }
    
    // 3. Draw active cell highlight (darker border)
    const activeCell = selection.activeCell;
    const activeBounds = this.cellToBounds(activeCell, viewport);
    if (activeBounds) {
      ctx.strokeStyle = "#0e65eb";
      ctx.lineWidth = 3;
      ctx.strokeRect(activeBounds.x, activeBounds.y, activeBounds.width, activeBounds.height);
    }
    
    // 4. Draw fill handle (small square in bottom-right corner)
    if (activeBounds && selection.type !== "column" && selection.type !== "row") {
      const handleSize = 8;
      ctx.fillStyle = "#0e65eb";
      ctx.fillRect(
        activeBounds.x + activeBounds.width - handleSize / 2,
        activeBounds.y + activeBounds.height - handleSize / 2,
        handleSize,
        handleSize
      );
    }
  }
}
```

---

## Overlay System

### DOM Overlays for Interactive Elements

Certain elements require DOM for accessibility and standard behavior:

```typescript
class OverlayManager {
  private overlayContainer: HTMLDivElement;
  private cellEditor: HTMLDivElement;
  private contextMenu: HTMLDivElement;
  private dropdown: HTMLDivElement;
  
  showCellEditor(cell: CellRef, bounds: Rect, initialValue: string): void {
    // Position editor over cell
    this.cellEditor.style.display = "block";
    this.cellEditor.style.left = `${bounds.x}px`;
    this.cellEditor.style.top = `${bounds.y}px`;
    this.cellEditor.style.width = `${bounds.width}px`;
    this.cellEditor.style.height = `${bounds.height}px`;
    
    // Initialize content
    this.cellEditor.textContent = initialValue;
    this.cellEditor.focus();
    
    // Select all if replacing, place cursor at end if editing
    const range = document.createRange();
    range.selectNodeContents(this.cellEditor);
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }
  
  hideCellEditor(): string {
    const value = this.cellEditor.textContent || "";
    this.cellEditor.style.display = "none";
    return value;
  }
}
```

In shared-grid mode, overlay-driven interactions that affect layout (row/column sizing, Hide/Unhide) should flow through the same **axis size override** path used by the canvas renderer so that visual state is consistent across the grid canvas and any DOM overlays.

### Expanding Cell Editor

Cell editor should expand as user types:

```typescript
class ExpandingCellEditor {
  private editor: HTMLDivElement;
  private minWidth: number;
  private minHeight: number;
  
  constructor(private gridRenderer: GridRenderer) {
    this.editor = document.createElement("div");
    this.editor.contentEditable = "true";
    this.editor.className = "cell-editor";
    
    this.editor.addEventListener("input", () => this.adjustSize());
  }
  
  private adjustSize(): void {
    // Reset to measure
    this.editor.style.width = "auto";
    this.editor.style.height = "auto";
    
    // Measure content
    const contentWidth = this.editor.scrollWidth;
    const contentHeight = this.editor.scrollHeight;
    
    // Apply with minimum (original cell size)
    this.editor.style.width = `${Math.max(this.minWidth, contentWidth)}px`;
    this.editor.style.height = `${Math.max(this.minHeight, contentHeight)}px`;
  }
}
```

---

## Performance Optimizations

### Dirty Region Tracking

Only repaint what changed:

```typescript
class DirtyRegionTracker {
  private dirtyRegions: Rect[] = [];
  
  markDirty(region: Rect): void {
    // Merge overlapping regions
    for (let i = 0; i < this.dirtyRegions.length; i++) {
      if (this.overlaps(region, this.dirtyRegions[i])) {
        this.dirtyRegions[i] = this.union(region, this.dirtyRegions[i]);
        return;
      }
    }
    this.dirtyRegions.push(region);
  }
  
  getDirtyRegions(): Rect[] {
    const regions = this.dirtyRegions;
    this.dirtyRegions = [];
    return regions;
  }
  
  renderDirtyRegions(ctx: CanvasRenderingContext2D): void {
    for (const region of this.getDirtyRegions()) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(region.x, region.y, region.width, region.height);
      ctx.clip();
      
      // Render only cells in this region
      this.renderRegion(ctx, region);
      
      ctx.restore();
    }
  }
}
```

### Off-Screen Buffer

Render to off-screen canvas, then composite:

```typescript
class BufferedRenderer {
  private buffer: OffscreenCanvas;
  private bufferCtx: OffscreenCanvasRenderingContext2D;
  
  constructor(width: number, height: number) {
    this.buffer = new OffscreenCanvas(width, height);
    this.bufferCtx = this.buffer.getContext("2d")!;
  }
  
  render(targetCtx: CanvasRenderingContext2D): void {
    // Render to buffer
    this.renderContent(this.bufferCtx);
    
    // Copy to target in single operation
    targetCtx.drawImage(this.buffer, 0, 0);
  }
}
```

### Cell Content Cache

Cache rendered text for cells that haven't changed:

```typescript
interface CachedCell {
  value: CellValue;
  formattedText: string;
  style: CellStyle;
  textMetrics: TextMetrics;
}

class CellCache {
  private cache = new Map<string, CachedCell>();
  private maxSize = 10000;
  
  get(row: number, col: number): CachedCell | undefined {
    return this.cache.get(`${row},${col}`);
  }
  
  set(row: number, col: number, cell: CachedCell): void {
    if (this.cache.size >= this.maxSize) {
      // LRU eviction
      const firstKey = this.cache.keys().next().value;
      this.cache.delete(firstKey);
    }
    this.cache.set(`${row},${col}`, cell);
  }
  
  invalidate(row: number, col: number): void {
    this.cache.delete(`${row},${col}`);
  }
  
  invalidateRange(range: Range): void {
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        this.invalidate(r, c);
      }
    }
  }
}
```

---

## Frozen Rows and Columns

### Split Pane Rendering

```typescript
class FrozenPaneRenderer {
  render(
    ctx: CanvasRenderingContext2D,
    viewport: Viewport,
    frozenRows: number,
    frozenCols: number
  ): void {
    // Four quadrants:
    // ┌────────┬────────────────┐
    // │ Frozen │ Frozen Rows    │
    // │ Corner │ (scrolls X)    │
    // ├────────┼────────────────┤
    // │ Frozen │ Main Area      │
    // │ Cols   │ (scrolls X+Y)  │
    // │(scr Y) │                │
    // └────────┴────────────────┘
    
    const frozenWidth = this.getColumnsWidth(0, frozenCols);
    const frozenHeight = this.getRowsHeight(0, frozenRows);
    
    // 1. Render main scrolling area
    ctx.save();
    ctx.beginPath();
    ctx.rect(frozenWidth, frozenHeight, 
             viewport.width - frozenWidth, 
             viewport.height - frozenHeight);
    ctx.clip();
    this.renderCells(ctx, viewport.startRow, viewport.endRow,
                     viewport.startCol, viewport.endCol, 
                     frozenWidth - viewport.offsetX,
                     frozenHeight - viewport.offsetY);
    ctx.restore();
    
    // 2. Render frozen columns (scrolls vertically only)
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, frozenHeight, frozenWidth, viewport.height - frozenHeight);
    ctx.clip();
    this.renderCells(ctx, viewport.startRow, viewport.endRow,
                     0, frozenCols - 1,
                     0, frozenHeight - viewport.offsetY);
    ctx.restore();
    
    // 3. Render frozen rows (scrolls horizontally only)
    ctx.save();
    ctx.beginPath();
    ctx.rect(frozenWidth, 0, viewport.width - frozenWidth, frozenHeight);
    ctx.clip();
    this.renderCells(ctx, 0, frozenRows - 1,
                     viewport.startCol, viewport.endCol,
                     frozenWidth - viewport.offsetX, 0);
    ctx.restore();
    
    // 4. Render frozen corner (never scrolls)
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, frozenWidth, frozenHeight);
    ctx.clip();
    this.renderCells(ctx, 0, frozenRows - 1, 0, frozenCols - 1, 0, 0);
    ctx.restore();
    
    // 5. Draw freeze lines
    ctx.strokeStyle = "#c0c0c0";
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(frozenWidth, 0);
    ctx.lineTo(frozenWidth, viewport.height);
    ctx.moveTo(0, frozenHeight);
    ctx.lineTo(viewport.width, frozenHeight);
    ctx.stroke();
  }
}
```

---

## Responsive Design

### Window Resize Handling

```typescript
class ResponsiveGrid {
  private resizeObserver: ResizeObserver;
  
  constructor(container: HTMLElement) {
    this.resizeObserver = new ResizeObserver((entries) => {
      for (const entry of entries) {
        this.handleResize(entry.contentRect);
      }
    });
    
    this.resizeObserver.observe(container);
  }
  
  private handleResize(rect: DOMRect): void {
    // Resize canvases
    this.resizeCanvases(rect.width, rect.height);
    
    // Recalculate visible range
    this.viewport.update(rect.width, rect.height);
    
    // Re-render
    this.requestRender();
  }
  
  private resizeCanvases(width: number, height: number): void {
    const dpr = window.devicePixelRatio || 1;
    
    for (const canvas of this.canvases) {
      canvas.width = width * dpr;
      canvas.height = height * dpr;
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
      
      const ctx = canvas.getContext("2d")!;
      ctx.scale(dpr, dpr);
    }
  }
}
```

---

## Accessibility

### Screen Reader Support

```typescript
class AccessibleGrid {
  private liveRegion: HTMLDivElement;
  private virtualTable: HTMLTableElement;
  
  constructor() {
    // Live region for announcements
    this.liveRegion = document.createElement("div");
    this.liveRegion.setAttribute("role", "status");
    this.liveRegion.setAttribute("aria-live", "polite");
    this.liveRegion.className = "sr-only";
    
    // Hidden table for screen reader navigation
    this.virtualTable = document.createElement("table");
    this.virtualTable.setAttribute("role", "grid");
    this.virtualTable.className = "sr-only";
  }
  
  announceCell(cell: CellRef, value: string): void {
    const address = this.cellToAddress(cell);
    this.liveRegion.textContent = `${address}: ${value}`;
  }
  
  announceSelection(selection: Selection): void {
    if (selection.ranges.length === 1) {
      const range = selection.ranges[0];
      if (range.startRow === range.endRow && range.startCol === range.endCol) {
        this.announceCell(selection.activeCell, this.getCellValue(selection.activeCell));
      } else {
        this.liveRegion.textContent = `Selected range ${this.rangeToAddress(range)}`;
      }
    } else {
      this.liveRegion.textContent = `${selection.ranges.length} ranges selected`;
    }
  }
}
```

### Keyboard Navigation

| Key | Action |
|-----|--------|
| Arrow keys | Move selection |
| Tab | Move right, wrap to next row |
| Enter | Move down, confirm edit |
| Ctrl+Arrow | Jump to edge of data |
| Ctrl+Shift+Arrow | Extend selection to edge |
| Ctrl+Home | Go to A1 |
| Ctrl+End | Go to last used cell |
| F2 | Edit cell |
| Escape | Cancel edit |
| Delete | Clear cell content |

---

## Testing Strategy

### Visual Regression Testing

```typescript
// Use tools like Percy or Chromatic
describe("Grid Rendering", () => {
  it("renders basic cells correctly", async () => {
    const grid = new GridRenderer(container);
    grid.setData(testData);
    grid.render();
    
    await expect(container).toMatchSnapshot();
  });
  
  it("renders selection correctly", async () => {
    const grid = new GridRenderer(container);
    grid.setSelection({ type: "range", ranges: [{ startRow: 0, endRow: 5, startCol: 0, endCol: 3 }] });
    grid.render();
    
    await expect(container).toMatchSnapshot();
  });
});
```

### Performance Benchmarks

```typescript
describe("Performance", () => {
  it("maintains 60fps while scrolling", async () => {
    const grid = new GridRenderer(container);
    grid.setData(generateLargeDataset(1000000, 100)); // 1M rows, 100 cols
    
    const frames: number[] = [];
    const startTime = performance.now();
    
    // Simulate smooth scroll
    for (let i = 0; i < 60; i++) {
      grid.scrollTo(0, i * 100);
      frames.push(performance.now());
      await new Promise(r => requestAnimationFrame(r));
    }
    
    // Calculate frame times
    const frameTimes = frames.slice(1).map((t, i) => t - frames[i]);
    const avgFrameTime = frameTimes.reduce((a, b) => a + b) / frameTimes.length;
    
    expect(avgFrameTime).toBeLessThan(16.67); // 60fps = 16.67ms per frame
  });
});
```
