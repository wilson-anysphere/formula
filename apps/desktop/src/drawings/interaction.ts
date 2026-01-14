import type { AnchorPoint, DrawingObject, DrawingTransform, Rect } from "./types";
import type { GridGeometry, Viewport } from "./overlay";
import { anchorToRectPx, emuToPx, pxToEmu } from "./overlay";
import { buildHitTestIndex, hitTestDrawings, hitTestDrawingsObject, type HitTestIndex, type HitTestResult } from "./hitTest";
import {
  cursorForResizeHandleWithTransform,
  cursorForRotationHandle,
  hitTestResizeHandle,
  hitTestRotationHandle,
  type ResizeHandle,
} from "./selectionHandles";
import {
  extractXfrmOff,
  patchAnchorExt,
  patchAnchorPoint,
  patchAnchorPos,
  patchXfrmExt,
  patchXfrmOff,
  patchXfrmRot,
} from "./drawingml/patch";
import { applyTransformVectorInto, inverseTransformVectorInto, normalizeRotationDeg } from "./transform";

const A1_CELL = { row: 0, col: 0 };
const CELL_SCRATCH = { row: 0, col: 0 };
const TRANSFORM_VEC_SCRATCH = { x: 0, y: 0 };
const IDENTITY_DRAWING_TRANSFORM: DrawingTransform = { rotationDeg: 0, flipH: false, flipV: false };

export type DrawingInteractionCommitKind = "move" | "resize" | "rotate";

export type DrawingInteractionCommit = {
  kind: DrawingInteractionCommitKind;
  /** UI id (numeric). */
  id: number;
  /** Drawing state at gesture start. */
  before: DrawingObject;
  /** Final drawing state at gesture commit time (after any DrawingML patching). */
  after: DrawingObject;
  /**
   * Full object list at commit time (after any commit-time patching).
   *
   * This allows persistence layers to write the final list without re-reading
   * `getObjects()` snapshots taken during pointermove.
   */
  objects: DrawingObject[];
};

export interface DrawingInteractionCallbacks {
  getViewport(): Viewport;
  getObjects(): DrawingObject[];
  setObjects(next: DrawingObject[]): void;
  /**
   * Commit a single drawing interaction gesture (move/resize/rotate).
   *
   * This is called once per gesture (pointerup) and includes the before/after
   * values for the edited object. The commit payload may also include the full
   * objects list at commit time (`commit.objects`) after any DrawingML patching
   * has been applied.
   */
  onInteractionCommit?(commit: DrawingInteractionCommit): void;
  /**
   * Commit the final drawing state to the backing document/store.
   *
   * This is called once per gesture (pointerup) so implementations can avoid
   * spamming document/collaboration updates during pointermove.
   */
  commitObjects?(next: DrawingObject[]): void;
  /**
   * Begin an undo batch for an interaction gesture.
   *
   * Implementations should call `DocumentController.beginBatch` (or equivalent).
   */
  beginBatch?(options: { label: string }): void;
  /** End the current undo batch. */
  endBatch?(): void;
  /** Cancel the current undo batch (Esc). */
  cancelBatch?(): void;
  onSelectionChange?(selectedId: number | null): void;
  /**
   * Optional focus request hook for integrations that want drawing selection to
   * keep keyboard focus on the grid root (so Delete/Ctrl+D shortcuts work).
   */
  requestFocus?(): void;
  /**
   * Optional scroll hook used for Excel-like auto-scroll while dragging/resizing.
   *
   * Returns true when scroll offsets changed.
   */
  scrollBy?(dx: number, dy: number): boolean;
  /**
   * Return false to skip handling the pointer down entirely.
   *
   * This is useful for cases where the grid should "win" even when the pointer
   * lands on a drawing (e.g. formula-bar range selection mode).
   */
  shouldHandlePointerDown?(event: PointerEvent): boolean;
  /**
   * Called when a pointer down hits a drawing, before selection/drag state is set.
   *
   * Return false to cancel drawing handling and allow the event to propagate.
   */
  onPointerDownHit?(event: PointerEvent, hit: HitTestResult): boolean | void;
}

export interface DrawingInteractionControllerOptions {
  /**
   * Register pointer listeners in capture phase.
   *
   * This is useful when the controller is attached to a grid root element and
   * needs to intercept events before a child canvas (e.g. the shared-grid
   * selection canvas).
   */
  capture?: boolean;
}

/**
 * Minimal MVP interactions: click-to-select and drag to move.
 */
export class DrawingInteractionController {
  private readonly scratchRect: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly scratchPaneLayout: PaneLayout = {
    frozenRows: 0,
    frozenCols: 0,
    headerOffsetX: 0,
    headerOffsetY: 0,
    frozenBoundaryX: 0,
    frozenBoundaryY: 0,
  };
  private hitTestIndex: HitTestIndex | null = null;
  private hitTestIndexObjects: readonly DrawingObject[] | null = null;
  private hitTestIndexZoom: number = 1;
  private dragging:
    | {
        id: number;
        startIndex: number;
        startSheetX: number;
        startSheetY: number;
        startObjects: DrawingObject[];
        pointerId: number;
      }
    | null = null;
  private resizing:
    | {
        id: number;
        startIndex: number;
        handle: ResizeHandle;
        startSheetX: number;
        startSheetY: number;
        startObjects: DrawingObject[];
        pointerId: number;
        transform?: DrawingTransform;
        startWidthPx: number;
        startHeightPx: number;
        /** Only set for image objects; used when Shift is held during resize. */
        aspectRatio: number | null;
      }
    | null = null;
  private rotating:
    | {
        id: number;
        startIndex: number;
        startAngleRad: number;
        centerX: number;
        centerY: number;
        startRotationDeg: number;
        startObjects: DrawingObject[];
        pointerId: number;
        transform?: DrawingTransform;
      }
    | null = null;
  private selectedId: number | null = null;
  private readonly isMacPlatform: boolean = (() => {
    try {
      const platform = typeof navigator !== "undefined" ? navigator.platform : "";
      return /Mac|iPhone|iPad|iPod/.test(platform);
    } catch {
      return false;
    }
  })();
  private escapeListenerAttached = false;
  private autoScrollRaf: number | null = null;
  private readonly localPointScratch: { x: number; y: number } = { x: 0, y: 0 };
  private readonly lastPointerScratch: { x: number; y: number; shiftKey: boolean } = { x: 0, y: 0, shiftKey: false };
  private lastPointer: { x: number; y: number; shiftKey: boolean } | null = null;
  private readonly sheetPointScratch: { x: number; y: number } = { x: 0, y: 0 };
  private readonly autoScrollDeltaScratch: { dx: number; dy: number } = { dx: 0, dy: 0 };
  private readonly transformVecScratch: { x: number; y: number } = { x: 0, y: 0 };
  private readonly aspectRatioDeltaScratch: { dx: number; dy: number } = { dx: 0, dy: 0 };

  /**
   * Mark a pointer event as a context-click that hit a drawing object.
   *
   * This is used to coordinate with other pointer handlers (e.g. the shared-grid
   * selection canvas) without stopping propagation: downstream listeners can
   * detect this flag and avoid treating the click as a grid cell context-click.
   */
  private markDrawingContextClick(e: PointerEvent): void {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (e as any).__formulaDrawingContextClick = true;
  }

  constructor(
    private readonly element: HTMLElement,
    private readonly geom: GridGeometry,
    private readonly callbacks: DrawingInteractionCallbacks,
    options: DrawingInteractionControllerOptions = {},
  ) {
    this.listenerOptions = { capture: options.capture ?? false };
    this.element.addEventListener("pointerdown", this.onPointerDown, this.listenerOptions);
    this.element.addEventListener("pointermove", this.onPointerMove, this.listenerOptions);
    this.element.addEventListener("pointerleave", this.onPointerLeave, this.listenerOptions);
    this.element.addEventListener("pointerup", this.onPointerUp, this.listenerOptions);
    this.element.addEventListener("pointercancel", this.onPointerCancel, this.listenerOptions);
  }

  dispose(options: { revert?: boolean } = {}): void {
    // If the app/view is torn down mid-gesture (e.g. hot reload, workbook switch),
    // ensure we release pointer capture and close any pending undo batch.
    try {
      if (options.revert === true) {
        this.cancelActiveGesture();
      } else {
        // Important: do *not* call `setObjects` during teardown. The in-memory object
        // state is owned by the host app; if it is being disposed, reverting objects
        // is unnecessary and can trigger unwanted UI work/tests (the view is going away).
        this.abandonActiveGesture();
      }
    } catch {
      // Best-effort: teardown should still remove listeners even if the integration
      // callbacks throw (e.g. partially-initialized test harnesses).
    }
    this.stopAutoScroll();
    this.lastPointer = null;
    this.element.removeEventListener("pointerdown", this.onPointerDown, this.listenerOptions);
    this.element.removeEventListener("pointermove", this.onPointerMove, this.listenerOptions);
    this.element.removeEventListener("pointerleave", this.onPointerLeave, this.listenerOptions);
    this.element.removeEventListener("pointerup", this.onPointerUp, this.listenerOptions);
    this.element.removeEventListener("pointercancel", this.onPointerCancel, this.listenerOptions);
    this.detachEscapeListener();
  }

  setSelectedId(id: number | null): void {
    this.selectedId = id;
  }

  /**
   * Clears cached hit test state so the next pointer interaction recomputes geometry.
   *
   * Call this when the underlying grid geometry changes (e.g. row/col resize) while
   * the `GridGeometry` object reference is stable.
   */
  invalidateHitTestIndex(): void {
    this.hitTestIndex = null;
    this.hitTestIndexObjects = null;
    this.hitTestIndexZoom = 1;
  }

  /**
   * Reset interaction state (drag/resize/rotate + selection).
   *
   * This is intended for integrations that swap out the underlying drawing layer
   * (e.g. switching sheets) while a gesture is in progress. We cancel any active
   * gesture before the integration changes the active sheet so that gesture
   * cleanup (`setObjects`, undo batching) is applied to the correct sheet.
   */
  reset(options?: { clearSelection?: boolean }): void {
    // Best-effort: if an interaction is in progress, cancel it and release any
    // pointer capture so sheet switches / teardown do not leave stale state.
    try {
      this.cancelActiveGesture();
    } catch {
      // ignore
    }
    if (options?.clearSelection) {
      this.selectedId = null;
    }
    // Cursor best-effort: avoid leaving resize/move cursors stuck when the
    // sheet changes mid-hover.
    this.element.style.cursor = "default";
  }

  /**
   * Cached bounding rect (client-space) used to convert `clientX/Y` → local
   * coordinates without doing per-pointermove layout reads.
   *
   * This is set on pointerdown when a drag/resize starts and cleared on
   * pointerup/cancel.
   */
  private activeRect: DOMRect | null = null;
  private readonly listenerOptions: AddEventListenerOptions;

  private getLocalPoint(e: PointerEvent, rect: DOMRect): { x: number; y: number } {
    // Perf: this is called on the pointermove hot path; avoid allocating a new `{x,y}` object.
    const out = this.localPointScratch;
    out.x = e.clientX - rect.left;
    out.y = e.clientY - rect.top;
    return out;
  }

  private stopPointerEvent(e: PointerEvent): void {
    const anyEvent = e as any;
    if (typeof anyEvent.preventDefault === "function") anyEvent.preventDefault();
    // Stop both bubbling to parents and any subsequent listeners on the same element.
    if (typeof anyEvent.stopPropagation === "function") anyEvent.stopPropagation();
    // `stopImmediatePropagation` isn't strictly required for the grid-root capture use case,
    // but it makes arbitration resilient when multiple listeners are attached to the same element.
    if (typeof anyEvent.stopImmediatePropagation === "function") anyEvent.stopImmediatePropagation();
  }

  private readonly onKeyDown = (e: KeyboardEvent) => {
    if (e.key !== "Escape") return;
    if (!this.dragging && !this.resizing && !this.rotating) return;
    e.preventDefault();
    // Ensure the spreadsheet/grid key handlers do not interpret Escape as "deselect"
    // while we're actively dragging/resizing/rotating a drawing.
    e.stopPropagation();
    this.cancelActiveGesture();
  };

  private attachEscapeListener(): void {
    if (this.escapeListenerAttached) return;
    if (typeof window === "undefined") return;
    // Capture phase so Escape cancels the gesture before SpreadsheetApp's root keydown
    // handler can consume it (and stop propagation).
    window.addEventListener("keydown", this.onKeyDown, { capture: true });
    this.escapeListenerAttached = true;
  }

  private detachEscapeListener(): void {
    if (!this.escapeListenerAttached) return;
    if (typeof window === "undefined") return;
    window.removeEventListener("keydown", this.onKeyDown, { capture: true });
    this.escapeListenerAttached = false;
  }

  private trySetPointerCapture(pointerId: number): void {
    const fn = (this.element as any)?.setPointerCapture;
    if (typeof fn !== "function") return;
    try {
      fn.call(this.element, pointerId);
    } catch {
      // Best-effort: some environments (jsdom) may not fully implement pointer capture.
    }
  }

  private tryReleasePointerCapture(pointerId: number): void {
    const fn = (this.element as any)?.releasePointerCapture;
    if (typeof fn !== "function") return;
    try {
      fn.call(this.element, pointerId);
    } catch {
      // ignore
    }
  }

  private readonly onPointerDown = (e: PointerEvent) => {
    if (this.callbacks.shouldHandlePointerDown && this.callbacks.shouldHandlePointerDown(e) === false) {
      return;
    }

    const pointerType = e.pointerType ?? "";
    const button = typeof e.button === "number" ? e.button : 0;
    const isMouse = pointerType === "mouse";
    // On macOS, Ctrl+click is commonly treated as a right click and fires the
    // `contextmenu` event. Ensure we treat it as a context-click (not a drag/resize).
    const isMacContextClick = isMouse && this.isMacPlatform && button === 0 && e.ctrlKey && !e.metaKey;
    const isNonPrimaryMouseButton = isMouse && button !== 0;
    const isContextClick = isNonPrimaryMouseButton || isMacContextClick;

    const rect = this.element.getBoundingClientRect();
    const { x, y } = this.getLocalPoint(e, rect);
    const viewport = this.callbacks.getViewport();
    const zoom = sanitizeZoom(viewport.zoom);
    const objects = this.callbacks.getObjects();
    const startObjects = ensureZOrderSorted(objects);
    const index = this.getHitTestIndex(objects, zoom);
    const paneLayout = resolveViewportPaneLayout(viewport, this.geom, this.scratchPaneLayout);
    const inHeader = x < paneLayout.headerOffsetX || y < paneLayout.headerOffsetY;
    // Pointer coordinates are reported for the full overlay element, including the row/column
    // header areas. Drawings live in the cell area under the headers, so clamp pointer
    // coordinates to the cell-area boundary before converting to sheet-space (avoids jumps
    // when a drag crosses into the header region while scroll offsets are non-zero).
    const clampedX = paneLayout.headerOffsetX > 0 ? Math.max(x, paneLayout.headerOffsetX) : x;
    const clampedY = paneLayout.headerOffsetY > 0 ? Math.max(y, paneLayout.headerOffsetY) : y;
    const pointInFrozenCols = clampedX < paneLayout.frozenBoundaryX;
    const pointInFrozenRows = clampedY < paneLayout.frozenBoundaryY;
    const startSheetPoint = this.sheetPointFromLocal(clampedX, clampedY, viewport, paneLayout);
    const startSheetX = startSheetPoint.x;
    const startSheetY = startSheetPoint.y;

    // Allow grabbing a resize handle for the current selection even when the
    // pointer is slightly outside the object's bounds (handles are centered on
    // the outline and extend half their size beyond the rect).
    const selectedIndex = this.selectedId != null ? index.byId.get(this.selectedId) : undefined;
    const selectedObject = selectedIndex != null ? index.ordered[selectedIndex] : undefined;
    if (selectedObject && !inHeader) {
      const anchor = selectedObject.anchor;
      const objInFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < paneLayout.frozenRows;
      const objInFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < paneLayout.frozenCols;
      if (objInFrozenCols === pointInFrozenCols && objInFrozenRows === pointInFrozenRows) {
        const selectedBounds = objectToScreenRect(
          selectedObject,
          viewport,
          this.geom,
          index.bounds[selectedIndex!],
          this.scratchRect,
        );
        const canRotate = selectedObject.kind.type !== "chart";
        if (canRotate && hitTestRotationHandle(selectedBounds, x, y, selectedObject.transform)) {
          if (
            this.callbacks.onPointerDownHit &&
            this.callbacks.onPointerDownHit(e, { object: selectedObject, bounds: selectedBounds }) === false
          ) {
            return;
          }
          if (isContextClick) {
            // Keep the current selection but allow the event to bubble so the app
            // can show a context menu.
            this.markDrawingContextClick(e);
            this.callbacks.requestFocus?.();
            return;
          }
          const centerX = selectedBounds.x + selectedBounds.width / 2;
          const centerY = selectedBounds.y + selectedBounds.height / 2;
          const startAngleRad = Math.atan2(y - centerY, x - centerX);
          const startRotationDeg = selectedObject.transform?.rotationDeg ?? 0;
          const startIndex = findObjectIndex(startObjects, selectedObject.id);
          this.stopPointerEvent(e);
          this.callbacks.requestFocus?.();
          this.activeRect = rect;
          this.trySetPointerCapture(e.pointerId);
          this.callbacks.beginBatch?.({ label: "Rotate Picture" });
          this.attachEscapeListener();
          this.rotating = {
            id: selectedObject.id,
            startIndex,
            startAngleRad,
            centerX,
            centerY,
            startRotationDeg,
            startObjects,
            pointerId: e.pointerId,
            transform: selectedObject.transform,
          };
          this.element.style.cursor = cursorForRotationHandle(true);
          return;
        }

        const handle = hitTestResizeHandle(selectedBounds, x, y, selectedObject.transform);
        if (handle) {
          if (
            this.callbacks.onPointerDownHit &&
            this.callbacks.onPointerDownHit(e, { object: selectedObject, bounds: selectedBounds }) === false
          ) {
            return;
          }
          if (isContextClick) {
            // Keep the current selection but allow the event to bubble so the app
            // can show a context menu.
            this.markDrawingContextClick(e);
            this.callbacks.requestFocus?.();
            return;
          }
          this.stopPointerEvent(e);
          this.activeRect = rect;
          this.trySetPointerCapture(e.pointerId);
          this.callbacks.beginBatch?.({ label: "Resize Picture" });
          this.attachEscapeListener();
          const startIndex = findObjectIndex(startObjects, selectedObject.id);
          this.resizing = {
            id: selectedObject.id,
            startIndex,
            handle,
            startSheetX,
            startSheetY,
            startObjects,
            pointerId: e.pointerId,
            transform: selectedObject.transform,
            startWidthPx: selectedBounds.width,
            startHeightPx: selectedBounds.height,
            aspectRatio:
              selectedObject.kind.type === "image" && selectedBounds.width > 0 && selectedBounds.height > 0
                ? selectedBounds.width / selectedBounds.height
                : null,
          };
          this.element.style.cursor = cursorForResizeHandleWithTransform(handle, selectedObject.transform);
          return;
        }
      }
    }

    const hit = hitTestDrawings(index, viewport, x, y, this.geom, paneLayout);
    if (hit && this.callbacks.onPointerDownHit && this.callbacks.onPointerDownHit(e, hit) === false) {
      return;
    }

    const nextSelectedId = hit?.object.id ?? null;
    // Excel-like behavior: context-clicks (right-click / Ctrl+click on macOS) should not drop the
    // current selection when the click misses all objects. They should still select the object
    // under the cursor when there is a hit.
    const shouldUpdateSelection = !isContextClick || hit != null;
    if (shouldUpdateSelection && nextSelectedId !== this.selectedId) {
      this.selectedId = nextSelectedId;
      this.callbacks.onSelectionChange?.(this.selectedId);
    }
    if (!hit) {
      this.element.style.cursor = "default";
      return;
    }

    if (isContextClick) {
      this.markDrawingContextClick(e);
      // Allow the downstream `contextmenu` handler to open without initiating a
      // drag/resize. (Don't stop propagation: right-click selection should still
      // bubble so consumers can open their own context menus.)
      this.callbacks.requestFocus?.();
      return;
    }

    this.stopPointerEvent(e);
    this.callbacks.requestFocus?.();
    this.activeRect = rect;
    this.trySetPointerCapture(e.pointerId);
    const handle = hitTestResizeHandle(hit.bounds, x, y, hit.object.transform);
    if (handle) {
      this.callbacks.beginBatch?.({ label: "Resize Picture" });
      this.attachEscapeListener();
      const startIndex = findObjectIndex(startObjects, hit.object.id);
      this.resizing = {
        id: hit.object.id,
        startIndex,
        handle,
        startSheetX,
        startSheetY,
        startObjects,
        pointerId: e.pointerId,
        transform: hit.object.transform,
        startWidthPx: hit.bounds.width,
        startHeightPx: hit.bounds.height,
        aspectRatio:
          hit.object.kind.type === "image" && hit.bounds.width > 0 && hit.bounds.height > 0
            ? hit.bounds.width / hit.bounds.height
            : null,
      };
      this.element.style.cursor = cursorForResizeHandleWithTransform(handle, hit.object.transform);
    } else {
      this.callbacks.beginBatch?.({ label: "Move Picture" });
      this.attachEscapeListener();
      const startIndex = findObjectIndex(startObjects, hit.object.id);
      this.dragging = {
        id: hit.object.id,
        startIndex,
        startSheetX,
        startSheetY,
        startObjects,
        pointerId: e.pointerId,
      };
      this.element.style.cursor = "move";
    }
  };

  private readonly onPointerMove = (e: PointerEvent) => {
    if (this.rotating) {
      if (e.pointerId !== this.rotating.pointerId) return;
      this.stopPointerEvent(e);
      const rect = this.activeRect ?? this.element.getBoundingClientRect();
      const { x, y } = this.getLocalPoint(e, rect);

      const angle = Math.atan2(y - this.rotating.centerY, x - this.rotating.centerX);
      const deltaDeg = ((angle - this.rotating.startAngleRad) * 180) / Math.PI;
      const rotationDeg = normalizeRotationDeg(this.rotating.startRotationDeg + deltaDeg);
      const baseTransform = this.rotating.transform ?? IDENTITY_DRAWING_TRANSFORM;
      const flipH = baseTransform.flipH;
      const flipV = baseTransform.flipV;
      const needsTransform = rotationDeg !== 0 || flipH || flipV;
      const nextTransform: DrawingTransform | null = needsTransform ? { rotationDeg, flipH, flipV } : null;

      const startObjects = this.rotating.startObjects;
      const startIndex = this.rotating.startIndex;
      const base = startObjects[startIndex];
      const next = startObjects.slice();
      if (base && base.id === this.rotating.id) {
        if (!needsTransform) {
          const { transform: _old, ...rest } = base;
          next[startIndex] = rest;
        } else {
          next[startIndex] = { ...base, transform: nextTransform! };
        }
      } else {
        // Defensive fallback: if the cached start index is invalid, fall back to a
        // linear scan to locate the edited object.
        const fallbackIndex = findObjectIndex(startObjects, this.rotating.id);
        if (fallbackIndex >= 0) {
          const fallbackBase = startObjects[fallbackIndex]!;
          if (!needsTransform) {
            const { transform: _old, ...rest } = fallbackBase;
            next[fallbackIndex] = rest;
          } else {
            next[fallbackIndex] = { ...fallbackBase, transform: nextTransform! };
          }
        }
      }
      this.callbacks.setObjects(next);
      this.element.style.cursor = cursorForRotationHandle(true);
      return;
    }

    if (this.resizing || this.dragging) {
      const active = this.resizing ?? this.dragging;
      if (!active) return;
      if (e.pointerId !== active.pointerId) return;
      this.stopPointerEvent(e);
      const rect = this.activeRect ?? this.element.getBoundingClientRect();
      const { x, y } = this.getLocalPoint(e, rect);
      const pointer = this.lastPointerScratch;
      pointer.x = x;
      pointer.y = y;
      pointer.shiftKey = Boolean(e.shiftKey);
      this.lastPointer = pointer;
      this.applyInteractionAtPointer(pointer);
      this.maybeStartAutoScroll();
      return;
    }

    const rect = this.element.getBoundingClientRect();
    const { x, y } = this.getLocalPoint(e, rect);
    this.updateCursor(x, y);
  };

  private readonly onPointerUp = (e: PointerEvent) => {
    const dragging = this.dragging;
    const resizing = this.resizing;
    const rotating = this.rotating;
    if (!dragging && !resizing && !rotating) return;
    const active = dragging ?? resizing ?? rotating;
    if (e.pointerId !== active.pointerId) return;

    this.stopPointerEvent(e);
    this.stopAutoScroll();
    this.lastPointer = null;

    const kind: "move" | "resize" | "rotate" = dragging ? "move" : resizing ? "resize" : "rotate";

    // Commit-time patching only: pointermove updates anchors for live previews,
    // while pointerup updates preserved DrawingML fragments (`rawXml`, `xlsx.pic_xml`)
    // so inner `<a:xfrm>` values (when present) stay consistent with the new anchor.
    const objects = this.callbacks.getObjects();
    let finalObjects = objects;
    const startObj =
      active.startObjects[active.startIndex]?.id === active.id
        ? active.startObjects[active.startIndex]
        : active.startObjects.find((o) => o.id === active.id);
    const currentObj =
      objects[active.startIndex]?.id === active.id ? objects[active.startIndex] : objects.find((o) => o.id === active.id);

    if (startObj && currentObj) {
      const zoom = sanitizeZoom(this.callbacks.getViewport().zoom);
      // Compute deltas/sizes in EMU directly so we preserve exact DrawingML units
      // (avoids float drift from px<->emu round-trips).
      const startPos = anchorTopLeftEmu(startObj.anchor, this.geom, zoom);
      const endPos = anchorTopLeftEmu(currentObj.anchor, this.geom, zoom);
      const endSize = anchorSizeEmu(currentObj.anchor, this.geom, zoom);

      const dxEmu = endPos.xEmu - startPos.xEmu;
      const dyEmu = endPos.yEmu - startPos.yEmu;
      const cxEmu = endSize.cxEmu;
      const cyEmu = endSize.cyEmu;

      let patched = currentObj;
      if (resizing) {
        patched = patchDrawingXmlForResize(patched, cxEmu, cyEmu);
      }
      if (rotating) {
        const rotationDeg = currentObj.transform?.rotationDeg ?? 0;
        patched = patchDrawingXmlForRotate(patched, rotationDeg);
      }
      if (dxEmu !== 0 || dyEmu !== 0) {
        patched = patchDrawingXmlForMove(patched, dxEmu, dyEmu);
      }

      if (patched !== currentObj) {
        finalObjects = objects.map((obj) => (obj.id === active.id ? patched : obj));
        this.callbacks.setObjects(finalObjects);
      }
    }
    const finalObj = finalObjects.find((o) => o.id === active.id);
    const rect = this.activeRect ?? this.element.getBoundingClientRect();
    const { x, y } = this.getLocalPoint(e, rect);

    this.dragging = null;
    this.resizing = null;
    this.rotating = null;
    this.activeRect = null;
    this.detachEscapeListener();
    this.tryReleasePointerCapture(active.pointerId);
    this.updateCursor(x, y);

    try {
      const onCommit = this.callbacks.onInteractionCommit;
      let committed = false;
      if (typeof onCommit === "function" && startObj && finalObj) {
        try {
          onCommit({ kind, id: active.id, before: startObj, after: finalObj, objects: finalObjects });
          committed = true;
        } catch {
          // Best-effort: do not crash interaction cleanup if the commit hook throws.
          committed = false;
        }
      }
      if (!committed) {
        try {
          this.callbacks.commitObjects?.(finalObjects);
        } catch {
          // Best-effort: persistence hooks should not break interaction cleanup.
        }
      }
    } finally {
      try {
        this.callbacks.endBatch?.();
      } catch {
        // Best-effort: never throw from cleanup.
      }
    }
  };

  private sheetPointFromLocal(x: number, y: number, viewport: Viewport, paneLayout: PaneLayout): { x: number; y: number } {
    // Sheet-space coordinates are relative to the A1 origin. Screen-space points
    // include `headerOffsetX/Y` and are shifted by scroll offsets except when the
    // pointer is inside a frozen pane.
    const clampedX = paneLayout.headerOffsetX > 0 ? Math.max(x, paneLayout.headerOffsetX) : x;
    const clampedY = paneLayout.headerOffsetY > 0 ? Math.max(y, paneLayout.headerOffsetY) : y;
    const effectiveScrollX = clampedX < paneLayout.frozenBoundaryX ? 0 : viewport.scrollX;
    const effectiveScrollY = clampedY < paneLayout.frozenBoundaryY ? 0 : viewport.scrollY;
    const out = this.sheetPointScratch;
    out.x = effectiveScrollX + (clampedX - paneLayout.headerOffsetX);
    out.y = effectiveScrollY + (clampedY - paneLayout.headerOffsetY);
    return out;
  }

  private applyInteractionAtPointer(pointer: { x: number; y: number; shiftKey: boolean }): void {
    const viewport = this.callbacks.getViewport();
    const zoom = sanitizeZoom(viewport.zoom);
    const paneLayout = resolveViewportPaneLayout(viewport, this.geom, this.scratchPaneLayout);
    const sheetPoint = this.sheetPointFromLocal(pointer.x, pointer.y, viewport, paneLayout);

    if (this.resizing) {
      let dx = sheetPoint.x - this.resizing.startSheetX;
      let dy = sheetPoint.y - this.resizing.startSheetY;

      const handle = this.resizing.handle;
      const isCornerHandle = handle === "nw" || handle === "ne" || handle === "se" || handle === "sw";
      if (isCornerHandle && pointer.shiftKey && this.resizing.aspectRatio != null) {
        const transform = this.resizing.transform;
        const startWidthPx = this.resizing.startWidthPx;
        const startHeightPx = this.resizing.startHeightPx;
        const aspectRatio = this.resizing.aspectRatio;
        const minSizePx = 8;

        if (hasNonIdentityTransform(transform)) {
          const local = inverseTransformVectorInto(dx, dy, transform, this.transformVecScratch);
          const lockedLocal = lockAspectRatioResizeInto(
            handle,
            local.x,
            local.y,
            startWidthPx,
            startHeightPx,
            aspectRatio,
            minSizePx,
            this.aspectRatioDeltaScratch,
          );
          const world = applyTransformVectorInto(lockedLocal.dx, lockedLocal.dy, transform, this.transformVecScratch);
          dx = world.x;
          dy = world.y;
        } else {
          const locked = lockAspectRatioResizeInto(
            handle,
            dx,
            dy,
            startWidthPx,
            startHeightPx,
            aspectRatio,
            minSizePx,
            this.aspectRatioDeltaScratch,
          );
          dx = locked.dx;
          dy = locked.dy;
        }
      }

      const startObjects = this.resizing.startObjects;
      const startIndex = this.resizing.startIndex;
      const base = startObjects[startIndex];
      const next = startObjects.slice();
      if (base && base.id === this.resizing.id) {
        next[startIndex] = {
          ...base,
          anchor: resizeAnchor(base.anchor, this.resizing.handle, dx, dy, this.geom, base.transform, zoom),
        };
      } else {
        // Defensive fallback: if the cached start index is invalid, fall back to a
        // linear scan to locate the edited object.
        const fallbackIndex = findObjectIndex(startObjects, this.resizing.id);
        const fallbackBase = fallbackIndex >= 0 ? startObjects[fallbackIndex]! : null;
        if (fallbackBase) {
          next[fallbackIndex] = {
            ...fallbackBase,
            anchor: resizeAnchor(fallbackBase.anchor, this.resizing.handle, dx, dy, this.geom, fallbackBase.transform, zoom),
          };
        }
      }
      this.callbacks.setObjects(next);
      this.element.style.cursor = cursorForResizeHandleWithTransform(this.resizing.handle, this.resizing.transform);
      return;
    }

    if (this.dragging) {
      const dx = sheetPoint.x - this.dragging.startSheetX;
      const dy = sheetPoint.y - this.dragging.startSheetY;

      const startObjects = this.dragging.startObjects;
      const startIndex = this.dragging.startIndex;
      const base = startObjects[startIndex];
      const next = startObjects.slice();
      if (base && base.id === this.dragging.id) {
        next[startIndex] = {
          ...base,
          anchor: shiftAnchor(base.anchor, dx, dy, this.geom, zoom),
        };
      } else {
        const fallbackIndex = findObjectIndex(startObjects, this.dragging.id);
        const fallbackBase = fallbackIndex >= 0 ? startObjects[fallbackIndex]! : null;
        if (fallbackBase) {
          next[fallbackIndex] = {
            ...fallbackBase,
            anchor: shiftAnchor(fallbackBase.anchor, dx, dy, this.geom, zoom),
          };
        }
      }
      this.callbacks.setObjects(next);
      this.element.style.cursor = "move";
    }
  }

  private computeAutoScrollDelta(pointer: { x: number; y: number }, viewport: Viewport, paneLayout: PaneLayout): { dx: number; dy: number } {
    const out = this.autoScrollDeltaScratch;
    const inHeader = pointer.x < paneLayout.headerOffsetX || pointer.y < paneLayout.headerOffsetY;
    if (inHeader) {
      out.dx = 0;
      out.dy = 0;
      return out;
    }

    const threshold = 24;
    const maxSpeed = 20;

    const pointInFrozenCols = pointer.x < paneLayout.frozenBoundaryX;
    const pointInFrozenRows = pointer.y < paneLayout.frozenBoundaryY;

    const leftEdge = paneLayout.frozenBoundaryX;
    const topEdge = paneLayout.frozenBoundaryY;
    const rightEdge = Math.max(0, viewport.width);
    const bottomEdge = Math.max(0, viewport.height);

    let dx = 0;
    if (!pointInFrozenCols) {
      if (pointer.x >= leftEdge && pointer.x < leftEdge + threshold) {
        let t = (leftEdge + threshold - pointer.x) / threshold;
        if (t < 0) t = 0;
        else if (t > 1) t = 1;
        dx = -Math.round(t * maxSpeed);
      } else if (pointer.x > rightEdge - threshold) {
        let t = (pointer.x - (rightEdge - threshold)) / threshold;
        if (t < 0) t = 0;
        else if (t > 1) t = 1;
        dx = Math.round(t * maxSpeed);
      }
    }

    let dy = 0;
    if (!pointInFrozenRows) {
      if (pointer.y >= topEdge && pointer.y < topEdge + threshold) {
        let t = (topEdge + threshold - pointer.y) / threshold;
        if (t < 0) t = 0;
        else if (t > 1) t = 1;
        dy = -Math.round(t * maxSpeed);
      } else if (pointer.y > bottomEdge - threshold) {
        let t = (pointer.y - (bottomEdge - threshold)) / threshold;
        if (t < 0) t = 0;
        else if (t > 1) t = 1;
        dy = Math.round(t * maxSpeed);
      }
    }

    out.dx = dx;
    out.dy = dy;
    return out;
  }

  private maybeStartAutoScroll(): void {
    if (this.autoScrollRaf != null) return;
    if (!this.dragging && !this.resizing) return;
    if (!this.lastPointer) return;
    if (!this.callbacks.scrollBy) return;

    const viewport = this.callbacks.getViewport();
    const paneLayout = resolveViewportPaneLayout(viewport, this.geom, this.scratchPaneLayout);
    const { dx, dy } = this.computeAutoScrollDelta(this.lastPointer, viewport, paneLayout);
    if (dx === 0 && dy === 0) return;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 16);

    const tick = () => {
      this.autoScrollRaf = null;
      if (!this.dragging && !this.resizing) return;
      if (!this.lastPointer) return;
      const viewport = this.callbacks.getViewport();
      const paneLayout = resolveViewportPaneLayout(viewport, this.geom, this.scratchPaneLayout);
      const { dx, dy } = this.computeAutoScrollDelta(this.lastPointer, viewport, paneLayout);
      if (dx === 0 && dy === 0) return;

      const didScroll = this.callbacks.scrollBy?.(dx, dy) ?? false;
      if (!didScroll) return;

      // Apply drag/resize again so the object tracks the changing scroll offsets
      // even when the pointer stays stationary near the edge.
      this.applyInteractionAtPointer(this.lastPointer);

      this.autoScrollRaf = schedule(tick) as unknown as number;
    };

    this.autoScrollRaf = schedule(tick) as unknown as number;
  }

  private stopAutoScroll(): void {
    if (this.autoScrollRaf == null) return;
    if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.autoScrollRaf);
    else globalThis.clearTimeout(this.autoScrollRaf);
    this.autoScrollRaf = null;
  }

  private readonly onPointerCancel = (e: PointerEvent) => {
    const active = this.dragging ?? this.resizing ?? this.rotating;
    if (!active) return;
    if (e.pointerId !== active.pointerId) return;
    // Treat `pointercancel` as a best-effort gesture commit (like pointerup). This matches
    // the behavior of other drag interactions in the app (e.g. scrollbars) and avoids
    // snapping the drawing back unexpectedly when the browser cancels the pointer stream.
    //
    // Explicit user cancellation is handled via Escape, which calls `cancelActiveGesture()`.
    this.onPointerUp(e);
  };

  private cancelActiveGesture(options?: { revertObjects?: boolean }): void {
    const active = this.dragging ?? this.resizing ?? this.rotating;
    if (!active) return;

    const revertObjects = options?.revertObjects !== false;
    const startObjects = active.startObjects;
    const pointerId = active.pointerId;
    this.stopAutoScroll();
    this.lastPointer = null;

    this.dragging = null;
    this.resizing = null;
    this.rotating = null;
    this.activeRect = null;
    this.detachEscapeListener();
    this.tryReleasePointerCapture(pointerId);

    // Revert the live in-memory state and cancel the undo batch.
    try {
      if (revertObjects) {
        this.callbacks.setObjects(startObjects);
      }
    } finally {
      this.callbacks.cancelBatch?.();
    }

    // Cursor best-effort: we may not have a meaningful point after cancel.
    this.element.style.cursor = "default";
  }

  /**
   * Clear any active gesture state without invoking `callbacks.setObjects`.
   *
   * This is useful for integration teardown paths where the host is being disposed anyway
   * and doesn't want cleanup to trigger additional rendering/state updates.
   */
  private abandonActiveGesture(): void {
    const active = this.dragging ?? this.resizing ?? this.rotating;
    if (!active) return;

    const pointerId = active.pointerId;

    this.dragging = null;
    this.resizing = null;
    this.rotating = null;
    this.activeRect = null;
    this.detachEscapeListener();
    this.tryReleasePointerCapture(pointerId);

    try {
      this.callbacks.cancelBatch?.();
    } catch {
      // ignore
    }

    this.element.style.cursor = "default";
  }

  private readonly onPointerLeave = () => {
    // Avoid leaving the resize/move cursor stuck when the pointer leaves the overlay canvas.
    if (this.dragging || this.resizing || this.rotating) return;
    this.element.style.cursor = "default";
  };

  private updateCursor(x: number, y: number): void {
    const viewport = this.callbacks.getViewport();
    const zoom = sanitizeZoom(viewport.zoom);
    const objects = this.callbacks.getObjects();
    const index = this.getHitTestIndex(objects, zoom);
    const paneLayout = resolveViewportPaneLayout(viewport, this.geom, this.scratchPaneLayout);
    if (x < paneLayout.headerOffsetX || y < paneLayout.headerOffsetY) {
      this.element.style.cursor = "default";
      return;
    }
    const pointInFrozenCols = x < paneLayout.frozenBoundaryX;
    const pointInFrozenRows = y < paneLayout.frozenBoundaryY;

    if (this.selectedId != null) {
      const selectedIndex = index.byId.get(this.selectedId);
      if (selectedIndex != null) {
        const selected = index.ordered[selectedIndex]!;
        const anchor = selected.anchor;
        const selectedInFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < paneLayout.frozenRows;
        const selectedInFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < paneLayout.frozenCols;
        if (
          selectedInFrozenCols === pointInFrozenCols &&
          selectedInFrozenRows === pointInFrozenRows
        ) {
          const bounds = objectToScreenRect(selected, viewport, this.geom, index.bounds[selectedIndex], this.scratchRect);
          if (selected.kind.type !== "chart" && hitTestRotationHandle(bounds, x, y, selected.transform)) {
            this.element.style.cursor = cursorForRotationHandle(false);
            return;
          }
          const handle = hitTestResizeHandle(bounds, x, y, selected.transform);
          if (handle) {
            this.element.style.cursor = cursorForResizeHandleWithTransform(handle, selected.transform);
            return;
          }
          if (pointInRect(x, y, bounds)) {
            this.element.style.cursor = "move";
            return;
          }
        }
      }
    }

    const hit = hitTestDrawingsObject(index, viewport, x, y, this.geom, paneLayout);
    if (hit) {
      this.element.style.cursor = "move";
      return;
    }

    this.element.style.cursor = "default";
  }

  private getHitTestIndex(objects: readonly DrawingObject[], zoom: number): HitTestIndex {
    const z = sanitizeZoom(zoom);
    const cached = this.hitTestIndex;
    // Use an epsilon comparison to avoid rebuilding the index for tiny floating-point
    // differences in zoom (e.g. when zoom comes from a scaled scroll/renderer state).
    //
    // This keeps the cache behavior aligned with `hitTestDrawings`' zoom-mismatch fallback
    // threshold (1e-6) so we don't accidentally fall back to O(N) scans.
    if (cached && this.hitTestIndexObjects === objects && Math.abs(this.hitTestIndexZoom - z) < 1e-6) return cached;
    const built = buildHitTestIndex(objects, this.geom, { zoom: z });
    this.hitTestIndex = built;
    this.hitTestIndexObjects = objects;
    this.hitTestIndexZoom = z;
    return built;
  }
}

function sanitizeZoom(zoom: number | undefined): number {
  return typeof zoom === "number" && Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
}

function ensureZOrderSorted(objects: DrawingObject[]): DrawingObject[] {
  if (objects.length <= 1) return objects;
  for (let i = 1; i < objects.length; i += 1) {
    if (objects[i - 1]!.zOrder > objects[i]!.zOrder) {
      return [...objects].sort((a, b) => a.zOrder - b.zOrder);
    }
  }
  return objects;
}

function findObjectIndex(objects: readonly DrawingObject[], id: number): number {
  for (let i = 0; i < objects.length; i += 1) {
    if (objects[i]!.id === id) return i;
  }
  return -1;
}

function anchorTopLeftEmu(
  anchor: DrawingObject["anchor"],
  geom: GridGeometry,
  zoom: number,
): { xEmu: number; yEmu: number } {
  const z = sanitizeZoom(zoom);
  switch (anchor.type) {
    case "absolute":
      return { xEmu: anchor.pos.xEmu, yEmu: anchor.pos.yEmu };
    case "oneCell": {
      const origin = geom.cellOriginPx(anchor.from.cell);
      return {
        xEmu: pxToEmu(origin.x / z) + anchor.from.offset.xEmu,
        yEmu: pxToEmu(origin.y / z) + anchor.from.offset.yEmu,
      };
    }
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);
      const x1 = pxToEmu(fromOrigin.x / z) + anchor.from.offset.xEmu;
      const y1 = pxToEmu(fromOrigin.y / z) + anchor.from.offset.yEmu;
      const x2 = pxToEmu(toOrigin.x / z) + anchor.to.offset.xEmu;
      const y2 = pxToEmu(toOrigin.y / z) + anchor.to.offset.yEmu;
      return { xEmu: Math.min(x1, x2), yEmu: Math.min(y1, y2) };
    }
  }
}

function anchorSizeEmu(anchor: DrawingObject["anchor"], geom: GridGeometry, zoom: number): { cxEmu: number; cyEmu: number } {
  const z = sanitizeZoom(zoom);
  switch (anchor.type) {
    case "absolute":
      return { cxEmu: anchor.size.cx, cyEmu: anchor.size.cy };
    case "oneCell":
      return { cxEmu: anchor.size.cx, cyEmu: anchor.size.cy };
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);
      const x1 = pxToEmu(fromOrigin.x / z) + anchor.from.offset.xEmu;
      const y1 = pxToEmu(fromOrigin.y / z) + anchor.from.offset.yEmu;
      const x2 = pxToEmu(toOrigin.x / z) + anchor.to.offset.xEmu;
      const y2 = pxToEmu(toOrigin.y / z) + anchor.to.offset.yEmu;
      return { cxEmu: Math.abs(x2 - x1), cyEmu: Math.abs(y2 - y1) };
    }
  }
}

export function patchDrawingXmlForResize(obj: DrawingObject, cxEmu: number, cyEmu: number): DrawingObject {
  return patchDrawingInnerXml(obj, (xml) => patchXfrmExt(xml, cxEmu, cyEmu));
}

export function patchDrawingXmlForMove(obj: DrawingObject, dxEmu: number, dyEmu: number): DrawingObject {
  // Only patch a:xfrm/a:off when the existing representation uses non-zero off
  // values. When off is already 0, we keep it at 0 and rely on anchors.
  return patchDrawingInnerXml(obj, (xml) => {
    const off = extractXfrmOff(xml);
    if (!off) return xml;
    if (off.xEmu === 0 && off.yEmu === 0) return xml;
    return patchXfrmOff(xml, off.xEmu + dxEmu, off.yEmu + dyEmu);
  });
}

function patchDrawingXmlForRotate(obj: DrawingObject, rotationDeg: number): DrawingObject {
  return patchDrawingInnerXml(obj, (xml) => patchXfrmRot(xml, rotationDeg));
}

function patchDrawingInnerXml(obj: DrawingObject, patch: (xml: string) => string): DrawingObject {
  if (obj.kind.type === "image") {
    const picXml = obj.preserved?.["xlsx.pic_xml"];
    if (typeof picXml !== "string") return obj;
    const patched = patch(picXml);
    if (patched === picXml) return obj;
    return {
      ...obj,
      preserved: {
        ...(obj.preserved ?? {}),
        "xlsx.pic_xml": patched,
      },
    };
  }

  const kindAny = obj.kind as any;
  const rawXml: unknown = kindAny.rawXml ?? kindAny.raw_xml;
  if (typeof rawXml !== "string") return obj;
  let patched = patch(rawXml);

  // Some DrawingML payloads (e.g. `DrawingObjectKind::Unknown` in the Rust model)
  // preserve the *entire* anchor subtree (`<xdr:twoCellAnchor>…</xdr:twoCellAnchor>`).
  // If the UI edits the anchor, we must patch those anchor fields too; otherwise
  // export will keep the stale wrapper.
  const isFullAnchorXml =
    /^\s*<(?:[A-Za-z_][\w.-]*:)?(?:oneCellAnchor|twoCellAnchor|absoluteAnchor)\b/.test(patched);
  if (isFullAnchorXml) {
    switch (obj.anchor.type) {
      case "oneCell":
        patched = patchAnchorPoint(patched, "from", {
          col: obj.anchor.from.cell.col,
          row: obj.anchor.from.cell.row,
          colOffEmu: obj.anchor.from.offset.xEmu,
          rowOffEmu: obj.anchor.from.offset.yEmu,
        });
        patched = patchAnchorExt(patched, obj.anchor.size.cx, obj.anchor.size.cy);
        break;
      case "twoCell":
        patched = patchAnchorPoint(patched, "from", {
          col: obj.anchor.from.cell.col,
          row: obj.anchor.from.cell.row,
          colOffEmu: obj.anchor.from.offset.xEmu,
          rowOffEmu: obj.anchor.from.offset.yEmu,
        });
        patched = patchAnchorPoint(patched, "to", {
          col: obj.anchor.to.cell.col,
          row: obj.anchor.to.cell.row,
          colOffEmu: obj.anchor.to.offset.xEmu,
          rowOffEmu: obj.anchor.to.offset.yEmu,
        });
        break;
      case "absolute":
        patched = patchAnchorPos(patched, obj.anchor.pos.xEmu, obj.anchor.pos.yEmu);
        patched = patchAnchorExt(patched, obj.anchor.size.cx, obj.anchor.size.cy);
        break;
    }
  }

  if (patched === rawXml) return obj;
  return {
    ...obj,
    kind: { ...kindAny, rawXml: patched, raw_xml: patched },
  };
}

function deltaPxToEmu(deltaPx: number, zoom: number): number {
  if (!Number.isFinite(deltaPx) || deltaPx === 0) return 0;
  // `Math.round` is asymmetric for negative half values (e.g. `Math.round(-0.5) === -0`).
  // Convert using the absolute value so `deltaPx` and `-deltaPx` always negate cleanly.
  const abs = pxToEmu(Math.abs(deltaPx), zoom);
  if (abs === 0) return 0;
  return deltaPx < 0 ? -abs : abs;
}

export function shiftAnchor(
  anchor: DrawingObject["anchor"],
  dxPx: number,
  dyPx: number,
  geom: GridGeometry,
  zoom: number = 1,
): DrawingObject["anchor"] {
  const z = sanitizeZoom(zoom);
  switch (anchor.type) {
    case "oneCell":
      return {
        ...anchor,
        from: shiftAnchorPoint(anchor.from, dxPx, dyPx, geom, z),
      };
    case "twoCell":
      return {
        ...anchor,
        from: shiftAnchorPoint(anchor.from, dxPx, dyPx, geom, z),
        to: shiftAnchorPoint(anchor.to, dxPx, dyPx, geom, z),
      };
    case "absolute":
      return {
        ...anchor,
        pos: {
          // Convert screen-pixel deltas to EMUs using symmetric rounding so opposite
          // nudges cancel cleanly (avoids off-by-one drift at fractional zoom steps).
          xEmu: anchor.pos.xEmu + deltaPxToEmu(dxPx, z),
          yEmu: anchor.pos.yEmu + deltaPxToEmu(dyPx, z),
        },
      };
  }
}

export function resizeAnchor(
  anchor: DrawingObject["anchor"],
  handle: ResizeHandle,
  dxPx: number,
  dyPx: number,
  geom: GridGeometry,
  transform?: DrawingTransform,
  zoom: number = 1,
): DrawingObject["anchor"] {
  const z = sanitizeZoom(zoom);
  let originA1x = 0;
  let originA1y = 0;
  if (anchor.type === "absolute") {
    try {
      const origin = geom.cellOriginPx(A1_CELL);
      originA1x = origin.x;
      originA1y = origin.y;
    } catch {
      originA1x = 0;
      originA1y = 0;
    }
  }

  let left = 0;
  let top = 0;
  let right = 0;
  let bottom = 0;
  if (anchor.type === "absolute") {
    left = originA1x + emuToPx(anchor.pos.xEmu) * z;
    top = originA1y + emuToPx(anchor.pos.yEmu) * z;
    right = originA1x + emuToPx(anchor.pos.xEmu + anchor.size.cx) * z;
    bottom = originA1y + emuToPx(anchor.pos.yEmu + anchor.size.cy) * z;
  } else if (anchor.type === "oneCell") {
    const p = anchorPointToSheetPx(anchor.from, geom, z);
    left = p.x;
    top = p.y;
    right = p.x + emuToPx(anchor.size.cx) * z;
    bottom = p.y + emuToPx(anchor.size.cy) * z;
  } else {
    const from = anchorPointToSheetPx(anchor.from, geom, z);
    const to = anchorPointToSheetPx(anchor.to, geom, z);
    left = from.x;
    top = from.y;
    right = to.x;
    bottom = to.y;
  }

  const movesLeftEdge = handle === "nw" || handle === "w" || handle === "sw";
  const movesTopEdge = handle === "nw" || handle === "n" || handle === "ne";
  const movesRightEdge = handle === "ne" || handle === "e" || handle === "se";
  const movesBottomEdge = handle === "sw" || handle === "s" || handle === "se";

  if (hasNonIdentityTransform(transform)) {
    // Convert pointer movement into the shape's local coordinate system (pre-rotation).
    const localDelta = inverseTransformVectorInto(dxPx, dyPx, transform, TRANSFORM_VEC_SCRATCH);
    let localDx = localDelta.x;
    let localDy = localDelta.y;

    // Edge handles resize along a single local axis: ignore perpendicular movement.
    if (handle === "e" || handle === "w") {
      localDy = 0;
    } else if (handle === "n" || handle === "s") {
      localDx = 0;
    }

    const width = right - left;
    const height = bottom - top;
    const hw = width / 2;
    const hh = height / 2;
    const cx = left + hw;
    const cy = top + hh;

    let localLeft = -hw;
    let localRight = hw;
    let localTop = -hh;
    let localBottom = hh;

    if (movesLeftEdge) localLeft += localDx;
    if (movesRightEdge) localRight += localDx;
    if (movesTopEdge) localTop += localDy;
    if (movesBottomEdge) localBottom += localDy;

    // Clamp against negative widths/heights while keeping the opposite edge stationary.
    if (localRight < localLeft) {
      if (movesLeftEdge) {
        localLeft = localRight;
      } else {
        localRight = localLeft;
      }
    }
    if (localBottom < localTop) {
      if (movesTopEdge) {
        localTop = localBottom;
      } else {
        localBottom = localTop;
      }
    }

    const nextWidth = Math.max(0, localRight - localLeft);
    const nextHeight = Math.max(0, localBottom - localTop);
    const localCenterShiftX = (localLeft + localRight) / 2;
    const localCenterShiftY = (localTop + localBottom) / 2;
    const worldCenterShift = applyTransformVectorInto(localCenterShiftX, localCenterShiftY, transform, TRANSFORM_VEC_SCRATCH);

    const nextCx = cx + worldCenterShift.x;
    const nextCy = cy + worldCenterShift.y;

    left = nextCx - nextWidth / 2;
    right = nextCx + nextWidth / 2;
    top = nextCy - nextHeight / 2;
    bottom = nextCy + nextHeight / 2;
  } else {
    switch (handle) {
      case "se":
        right += dxPx;
        bottom += dyPx;
        break;
      case "nw":
        left += dxPx;
        top += dyPx;
        break;
      case "ne":
        right += dxPx;
        top += dyPx;
        break;
      case "sw":
        left += dxPx;
        bottom += dyPx;
        break;
      case "e":
        right += dxPx;
        break;
      case "w":
        left += dxPx;
        break;
      case "s":
        bottom += dyPx;
        break;
      case "n":
        top += dyPx;
        break;
    }

    // Prevent negative widths/heights by clamping the moved edges against the
    // fixed ones. This keeps the opposite edge stationary.
    if (right < left) {
      if (movesLeftEdge) {
        left = right;
      } else {
        right = left;
      }
    }
    if (bottom < top) {
      if (movesTopEdge) {
        top = bottom;
      } else {
        bottom = top;
      }
    }
  }

  const widthPx = Math.max(0, right - left);
  const heightPx = Math.max(0, bottom - top);

  switch (anchor.type) {
    case "oneCell": {
      const start = anchorPointToSheetPx(anchor.from, geom, z);
      const nextFrom = shiftAnchorPoint(anchor.from, left - start.x, top - start.y, geom, z);
      return {
        ...anchor,
        from: nextFrom,
        size: { cx: Math.round(pxToEmu(widthPx / z)), cy: Math.round(pxToEmu(heightPx / z)) },
      };
    }
    case "absolute": {
      return {
        ...anchor,
        pos: { xEmu: Math.round(pxToEmu((left - originA1x) / z)), yEmu: Math.round(pxToEmu((top - originA1y) / z)) },
        size: { cx: Math.round(pxToEmu(widthPx / z)), cy: Math.round(pxToEmu(heightPx / z)) },
      };
    }
    case "twoCell": {
      const startFrom = anchorPointToSheetPx(anchor.from, geom, z);
      const startTo = anchorPointToSheetPx(anchor.to, geom, z);
      const nextFrom = shiftAnchorPoint(anchor.from, left - startFrom.x, top - startFrom.y, geom, z);
      const nextTo = shiftAnchorPoint(anchor.to, right - startTo.x, bottom - startTo.y, geom, z);
      return { ...anchor, from: nextFrom, to: nextTo };
    }
  }
}

function anchorPointToSheetPx(point: AnchorPoint, geom: GridGeometry, zoom: number = 1): { x: number; y: number } {
  const z = sanitizeZoom(zoom);
  const origin = geom.cellOriginPx(point.cell);
  return { x: origin.x + emuToPx(point.offset.xEmu) * z, y: origin.y + emuToPx(point.offset.yEmu) * z };
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): transform is DrawingTransform {
  if (!transform) return false;
  return transform.rotationDeg !== 0 || transform.flipH || transform.flipV;
}

const MAX_CELL_STEPS = 10_000;

export function shiftAnchorPoint(
  point: AnchorPoint,
  dxPx: number,
  dyPx: number,
  geom: GridGeometry,
  zoom: number = 1,
): AnchorPoint {
  const z = sanitizeZoom(zoom);
  let row = Number.isFinite(point.cell.row) ? Math.max(0, Math.trunc(point.cell.row)) : 0;
  let col = Number.isFinite(point.cell.col) ? Math.max(0, Math.trunc(point.cell.col)) : 0;
  // Keep offsets in EMUs and apply the delta in EMUs as well. Converting EMU -> px -> EMU
  // introduces rounding error at fractional zoom steps (e.g. 1px at 2x zoom => 0.5px),
  // which can accumulate into off-by-one drift after repeated nudges.
  let xEmu = (Number.isFinite(point.offset.xEmu) ? point.offset.xEmu : 0) + deltaPxToEmu(dxPx, z);
  let yEmu = (Number.isFinite(point.offset.yEmu) ? point.offset.yEmu : 0) + deltaPxToEmu(dyPx, z);

  // Normalize X across column boundaries.
  for (let i = 0; i < MAX_CELL_STEPS && xEmu < 0; i++) {
    if (col <= 0) {
      col = 0;
      xEmu = 0;
      break;
    }
    col -= 1;
    CELL_SCRATCH.row = row;
    CELL_SCRATCH.col = col;
    const wEmu = pxToEmu(geom.cellSizePx(CELL_SCRATCH).width, z);
    if (wEmu <= 0) {
      xEmu = 0;
      break;
    }
    xEmu += wEmu;
  }
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    CELL_SCRATCH.row = row;
    CELL_SCRATCH.col = col;
    const wEmu = pxToEmu(geom.cellSizePx(CELL_SCRATCH).width, z);
    if (wEmu <= 0) {
      xEmu = 0;
      break;
    }
    if (xEmu < wEmu) break;
    xEmu -= wEmu;
    col += 1;
  }

  // Normalize Y across row boundaries.
  for (let i = 0; i < MAX_CELL_STEPS && yEmu < 0; i++) {
    if (row <= 0) {
      row = 0;
      yEmu = 0;
      break;
    }
    row -= 1;
    CELL_SCRATCH.row = row;
    CELL_SCRATCH.col = col;
    const hEmu = pxToEmu(geom.cellSizePx(CELL_SCRATCH).height, z);
    if (hEmu <= 0) {
      yEmu = 0;
      break;
    }
    yEmu += hEmu;
  }
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    CELL_SCRATCH.row = row;
    CELL_SCRATCH.col = col;
    const hEmu = pxToEmu(geom.cellSizePx(CELL_SCRATCH).height, z);
    if (hEmu <= 0) {
      yEmu = 0;
      break;
    }
    if (yEmu < hEmu) break;
    yEmu -= hEmu;
    row += 1;
  }

  // Best-effort clamp to avoid tiny drift.
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    CELL_SCRATCH.row = row;
    CELL_SCRATCH.col = col;
    const wEmu = pxToEmu(geom.cellSizePx(CELL_SCRATCH).width, z);
    if (wEmu <= 0) {
      xEmu = 0;
      break;
    }
    if (xEmu < 0) xEmu = 0;
    if (xEmu < wEmu) break;
    xEmu -= wEmu;
    col += 1;
  }
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    CELL_SCRATCH.row = row;
    CELL_SCRATCH.col = col;
    const hEmu = pxToEmu(geom.cellSizePx(CELL_SCRATCH).height, z);
    if (hEmu <= 0) {
      yEmu = 0;
      break;
    }
    if (yEmu < 0) yEmu = 0;
    if (yEmu < hEmu) break;
    yEmu -= hEmu;
    row += 1;
  }

  // Normalize -0 so strict equality checks behave as expected.
  if (xEmu === 0) xEmu = 0;
  if (yEmu === 0) yEmu = 0;

  return {
    ...point,
    cell: { row, col },
    offset: { xEmu, yEmu },
  };
}

function objectToScreenRect(
  obj: DrawingObject,
  viewport: Viewport,
  geom: GridGeometry,
  sheetRect?: Rect,
  out?: Rect,
): Rect {
  const zoom = typeof viewport.zoom === "number" && Number.isFinite(viewport.zoom) && viewport.zoom > 0 ? viewport.zoom : 1;
  const rect = sheetRect ?? anchorToRectPx(obj.anchor, geom, zoom);
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  const anchor = obj.anchor;
  const inFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < frozenRows;
  const inFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < frozenCols;
  const scrollX = inFrozenCols ? 0 : viewport.scrollX;
  const scrollY = inFrozenRows ? 0 : viewport.scrollY;
  const target = out ?? { x: 0, y: 0, width: 0, height: 0 };
  target.x = rect.x - scrollX + headerOffsetX;
  target.y = rect.y - scrollY + headerOffsetY;
  target.width = rect.width;
  target.height = rect.height;
  return target;
}

function pointInRect(
  x: number,
  y: number,
  rect: { x: number; y: number; width: number; height: number },
): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}

type PaneLayout = {
  frozenRows: number;
  frozenCols: number;
  headerOffsetX: number;
  headerOffsetY: number;
  frozenBoundaryX: number;
  frozenBoundaryY: number;
};

const PANE_CELL_SCRATCH = { row: 0, col: 0 };

function clampNumber(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function resolveViewportPaneLayout(viewport: Viewport, geom: GridGeometry, out: PaneLayout): PaneLayout {
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  let frozenBoundaryX = headerOffsetX;
  let frozenBoundaryY = headerOffsetY;

  if (frozenCols > 0) {
    let raw = viewport.frozenWidthPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        PANE_CELL_SCRATCH.row = 0;
        PANE_CELL_SCRATCH.col = frozenCols;
        derived = geom.cellOriginPx(PANE_CELL_SCRATCH).x;
      } catch {
        derived = 0;
      }
      raw = headerOffsetX + derived;
    }
    frozenBoundaryX = clampNumber(raw as number, headerOffsetX, viewport.width);
  }

  if (frozenRows > 0) {
    let raw = viewport.frozenHeightPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        PANE_CELL_SCRATCH.row = frozenRows;
        PANE_CELL_SCRATCH.col = 0;
        derived = geom.cellOriginPx(PANE_CELL_SCRATCH).y;
      } catch {
        derived = 0;
      }
      raw = headerOffsetY + derived;
    }
    frozenBoundaryY = clampNumber(raw as number, headerOffsetY, viewport.height);
  }

  out.frozenRows = frozenRows;
  out.frozenCols = frozenCols;
  out.headerOffsetX = headerOffsetX;
  out.headerOffsetY = headerOffsetY;
  out.frozenBoundaryX = frozenBoundaryX;
  out.frozenBoundaryY = frozenBoundaryY;
  return out;
}

// NOTE: Call sites avoid allocating pane objects by computing frozen-row/col membership inline.
function lockAspectRatioResizeInto(
  handle: ResizeHandle,
  dx: number,
  dy: number,
  startWidthPx: number,
  startHeightPx: number,
  aspectRatio: number,
  minSizePx: number,
  out: { dx: number; dy: number },
): { dx: number; dy: number } {
  out.dx = dx;
  out.dy = dy;

  // Only lock corner-handle resizes (edge handles remain unconstrained).
  if (handle === "n" || handle === "e" || handle === "s" || handle === "w") return out;
  if (!Number.isFinite(aspectRatio) || aspectRatio <= 0) return out;
  if (!Number.isFinite(startWidthPx) || !Number.isFinite(startHeightPx)) return out;
  if (startWidthPx <= 0 || startHeightPx <= 0) return out;

  const sx = handle === "ne" || handle === "se" ? 1 : -1;
  const sy = handle === "sw" || handle === "se" ? 1 : -1;

  // Use the original bounds (captured once on resize start) as the single source of truth for the
  // aspect ratio. Avoid recomputing from intermediate sizes to prevent drift.
  //
  // Prefer width-driven scaling when the user is changing width more (relative to the starting
  // width). Otherwise, preserve the user's height change and derive width.
  const proposedWidth = startWidthPx + sx * dx;
  const proposedHeight = startHeightPx + sy * dy;

  const scaleW = proposedWidth / startWidthPx;
  const scaleH = proposedHeight / startHeightPx;

  const widthDriven = Math.abs(scaleW - 1) >= Math.abs(scaleH - 1);

  const minScale = Math.max(
    startWidthPx > minSizePx ? minSizePx / startWidthPx : 0,
    startHeightPx > minSizePx ? minSizePx / startHeightPx : 0,
    0,
  );

  if (widthDriven) {
    let scale = scaleW;
    if (!Number.isFinite(scale)) scale = 1;
    // Prevent flipping, and enforce a minimum visual size for stable ratio math.
    scale = Math.max(scale, minScale);
    const nextWidth = startWidthPx * scale;
    const nextHeight = nextWidth / aspectRatio;
    out.dx = (nextWidth - startWidthPx) * sx;
    out.dy = (nextHeight - startHeightPx) * sy;
    return out;
  }

  let scale = scaleH;
  if (!Number.isFinite(scale)) scale = 1;
  scale = Math.max(scale, minScale);
  const nextHeight = startHeightPx * scale;
  const nextWidth = nextHeight * aspectRatio;
  out.dx = (nextWidth - startWidthPx) * sx;
  out.dy = (nextHeight - startHeightPx) * sy;
  return out;
}
