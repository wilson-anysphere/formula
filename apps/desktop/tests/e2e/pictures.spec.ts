import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

type RectPx = { x: number; y: number; width: number; height: number };
type DrawingDebug = { id: number | string; kind: unknown; rectPx: RectPx | null };
type DrawingsDebugState = { selectedId: number | string | null; drawings: DrawingDebug[] };

const MODES: Array<{ name: string; path: string }> = [
  // Enable drawing interactions via the `?drawings=1` feature flag so drag/resize gestures are active.
  { name: "legacy", path: "/?grid=legacy&drawings=1" },
  { name: "shared-grid", path: "/?grid=shared&drawings=1" },
];

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n));
}

function normalizeKind(kind: unknown): string {
  if (typeof kind === "string") return kind;
  if (kind && typeof kind === "object" && "type" in kind && typeof (kind as any).type === "string") {
    return String((kind as any).type);
  }
  return String(kind);
}

function expectApprox(actual: number, expected: number, tolerance = 6): void {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(tolerance);
}

function pointInRect(point: { x: number; y: number }, rect: RectPx): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.width && point.y >= rect.y && point.y <= rect.y + rect.height;
}

function applyOffset(rect: RectPx, offset: { x: number; y: number }): RectPx {
  return { x: rect.x + offset.x, y: rect.y + offset.y, width: rect.width, height: rect.height };
}

function findImageDrawings(state: DrawingsDebugState): DrawingDebug[] {
  return state.drawings.filter((drawing) => normalizeKind(drawing.kind) === "image");
}

function findDrawingById(state: DrawingsDebugState, drawingId: number | string): DrawingDebug | null {
  const target = String(drawingId);
  return state.drawings.find((drawing) => String(drawing.id) === target) ?? null;
}

async function waitForDrawingsReady(page: Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.getDrawingsDebugState), undefined, { timeout: 10_000 });
}

async function getDrawingsDebugState(page: Page): Promise<DrawingsDebugState> {
  return await page.evaluate(() => {
    const app = window.__formulaApp as any;
    if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
    if (typeof app.getDrawingsDebugState !== "function") {
      throw new Error("Missing window.__formulaApp.getDrawingsDebugState()");
    }
    return app.getDrawingsDebugState();
  });
}

async function getBottomRightHandlePx(page: Page, drawingId: number | string, rect: RectPx): Promise<{ x: number; y: number }> {
  const maybePoint = await page.evaluate((id) => {
    const app = window.__formulaApp as any;
    if (!app || typeof app.getDrawingHandlePointsPx !== "function") return null;
    try {
      return app.getDrawingHandlePointsPx(id);
    } catch {
      return null;
    }
  }, drawingId);

  if (maybePoint && typeof maybePoint === "object") {
    const candidates: unknown[] = [
      (maybePoint as any).bottomRight,
      (maybePoint as any).se,
      (maybePoint as any).southEast,
      (maybePoint as any).bottom_right,
      (maybePoint as any)["bottom-right"],
    ];
    for (const candidate of candidates) {
      if (
        candidate &&
        typeof candidate === "object" &&
        typeof (candidate as any).x === "number" &&
        typeof (candidate as any).y === "number"
      ) {
        return { x: (candidate as any).x, y: (candidate as any).y };
      }
    }
  }

  // Fall back to the bottom-right corner of the rect (viewport-relative).
  return { x: rect.x + rect.width, y: rect.y + rect.height };
}

for (const mode of MODES) {
  test.describe(`pictures (${mode.name})`, () => {
    test("insert/select/drag/resize + undo/redo", async ({ page }) => {
      await gotoDesktop(page, mode.path);
      await waitForDrawingsReady(page);

      // Capture runtime errors after the app is booted (startup errors are handled by `gotoDesktop`).
      const runtimeConsoleErrors: string[] = [];
      const runtimePageErrors: string[] = [];
      const onConsole = (msg: any): void => {
        try {
          if (msg?.type?.() !== "error") return;
          runtimeConsoleErrors.push(msg.text());
        } catch {
          // ignore
        }
      };
      const onPageError = (err: any): void => {
        runtimePageErrors.push(err?.message ?? String(err));
      };
      page.on("console", onConsole);
      page.on("pageerror", onPageError);

      try {
        await page.evaluate(async () => {
          const app = window.__formulaApp as any;
          if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
          if (typeof app.insertPicturesFromFiles !== "function") {
            throw new Error("Missing window.__formulaApp.insertPicturesFromFiles()");
          }

          // 1x1 PNG (transparent) to avoid external assets.
          const base64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR42mP8z/C/HwAF/gL+F8m0lAAAAABJRU5ErkJggg==";
          const binary = atob(base64);
          const bytes = new Uint8Array(binary.length);
          for (let i = 0; i < binary.length; i += 1) {
            bytes[i] = binary.charCodeAt(i);
          }
          const file = new File([bytes], "tiny.png", { type: "image/png" });

          await app.insertPicturesFromFiles([file]);
          await app.whenIdle?.();
        });

        await expect
          .poll(async () => {
            const state = await getDrawingsDebugState(page);
            return findImageDrawings(state).length;
          }, { timeout: 10_000 })
          .toBe(1);
        await expect
          .poll(async () => {
            const state = await getDrawingsDebugState(page);
            const image = findImageDrawings(state)[0] ?? null;
            const rect = image?.rectPx ?? null;
            if (!rect) return null;
            if (!Number.isFinite(rect.width) || !Number.isFinite(rect.height)) return null;
            if (rect.width <= 0 || rect.height <= 0) return null;
            return rect;
          }, { timeout: 10_000 })
          .not.toBeNull();

        const stateAfterInsert = await getDrawingsDebugState(page);
        const drawing = findImageDrawings(stateAfterInsert)[0];
        if (!drawing) throw new Error("Expected an inserted image drawing");
        const drawingId = drawing.id;

        expect(normalizeKind(drawing.kind)).toBe("image");
        if (!drawing.rectPx) throw new Error("Expected drawing.rectPx to be non-null");
        expect(drawing.rectPx.width).toBeGreaterThan(0);
        expect(drawing.rectPx.height).toBeGreaterThan(0);

        // Clear selection by clicking outside the image.
        const gridBox = await page.locator("#grid").boundingBox();
        if (!gridBox) throw new Error("Missing grid bounding box");

        // The drawings debug hook is expected to return viewport-relative coordinates, but in some
        // renderers the returned rects are relative to the grid root. Detect and compensate so the
        // test stays robust in both legacy + shared-grid modes.
        const rectCenter = { x: drawing.rectPx.x + drawing.rectPx.width / 2, y: drawing.rectPx.y + drawing.rectPx.height / 2 };
        const gridOffset = pointInRect(rectCenter, gridBox as RectPx) ? { x: 0, y: 0 } : { x: gridBox.x, y: gridBox.y };

        // Aim near the bottom-right of the grid so we don't accidentally re-hit the drawing.
        const clearX = clamp(gridBox.x + gridBox.width - 40, gridBox.x + 10, gridBox.x + gridBox.width - 10);
        const clearY = clamp(gridBox.y + gridBox.height - 40, gridBox.y + 10, gridBox.y + gridBox.height - 10);
        await page.mouse.click(clearX, clearY);

        await expect
          .poll(async () => (await getDrawingsDebugState(page)).selectedId, { timeout: 10_000 })
          .not.toBe(drawingId);

        // Select.
        const rectPage = applyOffset(drawing.rectPx, gridOffset);
        const centerX = rectPage.x + rectPage.width / 2;
        const centerY = rectPage.y + rectPage.height / 2;
        await page.mouse.click(centerX, centerY);

        await expect.poll(async () => (await getDrawingsDebugState(page)).selectedId, { timeout: 10_000 }).toBe(drawingId);

        // Drag move.
        const rectBeforeMoveDebug = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
        if (!rectBeforeMoveDebug) throw new Error("Expected drawing rect before move");
        const rectBeforeMove = applyOffset(rectBeforeMoveDebug, gridOffset);
        const startX = rectBeforeMove.x + rectBeforeMove.width / 2;
        const startY = rectBeforeMove.y + rectBeforeMove.height / 2;
        const moveDx = 40;
        const moveDy = 20;
        await page.mouse.move(startX, startY);
        await page.mouse.down();
        await page.mouse.move(startX + moveDx, startY + moveDy, { steps: 6 });
        await page.mouse.up();

        await page.evaluate(() => (window.__formulaApp as any).whenIdle?.());

        await expect
          .poll(async () => {
            const rect = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
            return rect ? applyOffset(rect, gridOffset).x : null;
          }, { timeout: 10_000 })
          .toBeGreaterThan(rectBeforeMove.x + 10);
        await expect
          .poll(async () => {
            const rect = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
            return rect ? applyOffset(rect, gridOffset).y : null;
          }, { timeout: 10_000 })
          .toBeGreaterThan(rectBeforeMove.y + 5);

        const rectAfterMoveDebug = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
        if (!rectAfterMoveDebug) throw new Error("Expected drawing rect after move");
        const rectAfterMove = applyOffset(rectAfterMoveDebug, gridOffset);
        expectApprox(rectAfterMove.x, rectBeforeMove.x + moveDx);
        expectApprox(rectAfterMove.y, rectBeforeMove.y + moveDy);

        // Resize from the bottom-right handle (or bottom-right rect corner).
        const handleStartDebug = await getBottomRightHandlePx(page, drawingId, rectAfterMoveDebug);
        // Avoid hitting exactly on the rect boundary, which can be sensitive to rounding when the
        // drawing coordinates are fractional. Nudge slightly inward while staying within the handle
        // hit target (handle squares extend half their size beyond the rect bounds).
        const handleStart = { x: handleStartDebug.x + gridOffset.x - 2, y: handleStartDebug.y + gridOffset.y - 2 };
        const resizeDx = 30;
        const resizeDy = 30;
        await page.mouse.move(handleStart.x, handleStart.y);
        await page.mouse.down();
        // Make a small initial move to ensure we cross the controller's "dragging" threshold before
        // the main resize delta.
        await page.mouse.move(handleStart.x + 6, handleStart.y + 6, { steps: 2 });
        await page.mouse.move(handleStart.x + resizeDx, handleStart.y + resizeDy, { steps: 6 });
        await page.mouse.up();

        await page.evaluate(() => (window.__formulaApp as any).whenIdle?.());

        await expect
          .poll(async () => {
            const rect = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
            return rect?.width ?? null;
          }, { timeout: 10_000 })
          .toBeGreaterThan(rectAfterMove.width + 10);
        await expect
          .poll(async () => {
            const rect = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
            return rect?.height ?? null;
          }, { timeout: 10_000 })
          .toBeGreaterThan(rectAfterMove.height + 10);

        const rectAfterResizeDebug = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
        if (!rectAfterResizeDebug) throw new Error("Expected drawing rect after resize");
        const rectAfterResize = applyOffset(rectAfterResizeDebug, gridOffset);
        expect(rectAfterResize.width).toBeGreaterThan(rectAfterMove.width + 10);
        expect(rectAfterResize.height).toBeGreaterThan(rectAfterMove.height + 10);

        // Undo should revert the resize back to the moved rect.
        const quickAccessToolbar = page.getByRole("toolbar", { name: "Quick access toolbar" });
        await expect(quickAccessToolbar.getByRole("button", { name: "Undo" })).toBeEnabled({ timeout: 10_000 });
        await quickAccessToolbar.getByRole("button", { name: "Undo" }).click();
        await page.evaluate(() => (window.__formulaApp as any)?.whenIdle?.());

        await expect
          .poll(async () => {
            const rect = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
            return rect?.width ?? null;
          }, { timeout: 10_000 })
          .toBeLessThan(rectAfterResize.width - 5);

        const rectAfterUndoDebug = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
        if (!rectAfterUndoDebug) throw new Error("Expected drawing rect after undo");
        const rectAfterUndo = applyOffset(rectAfterUndoDebug, gridOffset);
        expectApprox(rectAfterUndo.x, rectAfterMove.x);
        expectApprox(rectAfterUndo.y, rectAfterMove.y);
        expectApprox(rectAfterUndo.width, rectAfterMove.width);
        expectApprox(rectAfterUndo.height, rectAfterMove.height);

        // Redo should re-apply the resize.
        await expect(quickAccessToolbar.getByRole("button", { name: "Redo" })).toBeEnabled({ timeout: 10_000 });
        await quickAccessToolbar.getByRole("button", { name: "Redo" }).click();
        await page.evaluate(() => (window.__formulaApp as any)?.whenIdle?.());

        await expect
          .poll(async () => {
            const rect = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
            return rect?.width ?? null;
          }, { timeout: 10_000 })
          .toBeGreaterThan(rectAfterUndo.width + 10);

        const rectAfterRedoDebug = findDrawingById(await getDrawingsDebugState(page), drawingId)?.rectPx ?? null;
        if (!rectAfterRedoDebug) throw new Error("Expected drawing rect after redo");
        const rectAfterRedo = applyOffset(rectAfterRedoDebug, gridOffset);
        expectApprox(rectAfterRedo.x, rectAfterResize.x);
        expectApprox(rectAfterRedo.y, rectAfterResize.y);
        expectApprox(rectAfterRedo.width, rectAfterResize.width);
        expectApprox(rectAfterRedo.height, rectAfterResize.height);
      } finally {
        page.off("console", onConsole);
        page.off("pageerror", onPageError);

        if (runtimeConsoleErrors.length > 0 || runtimePageErrors.length > 0) {
          const uniqueConsole = [...new Set(runtimeConsoleErrors)];
          const uniquePage = [...new Set(runtimePageErrors)];
          const parts: string[] = [];
          if (uniqueConsole.length > 0) parts.push(`Console errors:\n${uniqueConsole.join("\n")}`);
          if (uniquePage.length > 0) parts.push(`Page errors:\n${uniquePage.join("\n")}`);
          throw new Error(`Unexpected runtime errors while interacting with pictures.\n\n${parts.join("\n\n")}`);
        }
      }
    });
  });
}
