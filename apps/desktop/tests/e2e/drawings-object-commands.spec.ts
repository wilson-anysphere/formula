import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

const EMU_PER_PX = 914_400 / 96;
const DEFAULT_DRAWING_ANCHOR = {
  type: "oneCell",
  from: { cell: { row: 0, col: 0 }, offset: { xEmu: EMU_PER_PX * 18, yEmu: EMU_PER_PX * 18 } },
  size: { cx: EMU_PER_PX * 240, cy: EMU_PER_PX * 120 },
} as const;

async function getDrawingObjects(page: import("@playwright/test").Page): Promise<any[]> {
  return await page.evaluate(() => {
    const app = window.__formulaApp as any;
    const objects = Array.isArray(app?.getDrawingObjects?.()) ? (app.getDrawingObjects() as any[]) : [];
    // Canvas charts are enabled by default; ChartStore charts are surfaced as drawing objects.
    // This suite focuses on workbook drawings, so filter out ChartStore chart objects.
    return objects.filter((obj) => obj?.kind?.type !== "chart");
  });
}

async function getSelectedDrawingId(page: import("@playwright/test").Page): Promise<number | null> {
  return await page.evaluate(() => (window.__formulaApp as any).getSelectedDrawingId());
}

async function clickGridAt(page: import("@playwright/test").Page, pos: { x: number; y: number }): Promise<void> {
  const grid = page.locator("#grid");
  const box = await grid.boundingBox();
  if (!box) throw new Error("Missing #grid bounding box");
  await page.mouse.click(box.x + pos.x, box.y + pos.y);
}

async function openContextMenuAt(page: import("@playwright/test").Page, pos: { x: number; y: number }): Promise<void> {
  await page.evaluate(
    ({ x, y }) => {
      const root = document.querySelector("#grid") as HTMLElement | null;
      if (!root) throw new Error("Missing #grid");
      const rect = root.getBoundingClientRect();
      root.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + x,
          clientY: rect.top + y,
        }),
      );
    },
    { x: pos.x, y: pos.y },
  );
}

async function seedSingleShapeDrawing(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate((anchor) => {
    const app = window.__formulaApp as any;
    if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
    const doc = app.getDocument?.();
    if (!doc) throw new Error("Missing __formulaApp.getDocument()");
    const sheetId = app.getCurrentSheetId?.();
    if (!sheetId) throw new Error("Missing __formulaApp.getCurrentSheetId()");
    if (typeof doc.setSheetDrawings !== "function") {
      throw new Error("Missing DocumentController.setSheetDrawings()");
    }

    // Replace the drawings list with a single deterministic shape so this suite does not rely on
    // the `?drawings=1` demo fallback (which is intentionally non-persisted).
    doc.setSheetDrawings(
      sheetId,
      [
        {
          id: "1",
          kind: { type: "shape", label: "E2E Drawing" },
          anchor,
          zOrder: 0,
        },
      ],
      { label: "Seed Drawing" },
    );

    // Clear SpreadsheetApp caches + re-render so hit testing sees the new object immediately.
    app.syncSheetDrawings?.();
    app.focus?.();
  }, DEFAULT_DRAWING_ANCHOR);
}

test.describe("Drawing object commands", () => {
  test("duplicate (Ctrl/Cmd+D) and delete (Del) operate on selected drawing", async ({ page }) => {
    // `drawingInteractions=1` enables the interactions controller without opting into the
    // `?drawings=1` demo mode (which seeds non-persisted drawings + ChartStore charts).
    await gotoDesktop(page, "/?grid=shared&drawingInteractions=1");
    await seedSingleShapeDrawing(page);

    await expect.poll(() => getDrawingObjects(page)).toHaveLength(1);
    const [initial] = await getDrawingObjects(page);
    const initialId = initial?.id as number;
    expect(typeof initialId).toBe("number");

    // Click inside the seeded drawing (oneCell anchor near A1).
    await clickGridAt(page, { x: 100, y: 100 });
    await expect.poll(() => getSelectedDrawingId(page)).toBe(initialId);

    // Ensure keyboard focus is on the grid root so the window-level selected-object
    // shortcut handler can see the event target inside `#grid` (required for Ctrl/Cmd+D).
    await page.evaluate(() => (window as any).__formulaApp.focus());

    await page.keyboard.press("ControlOrMeta+D");

    await expect.poll(() => getDrawingObjects(page)).toHaveLength(2);
    const afterDup = await getDrawingObjects(page);
    const ids = afterDup.map((o) => o.id as number);
    // Duplicates should get a new globally-unique id (not necessarily `max+1`), so assert only
    // uniqueness + inclusion of the original id.
    expect(ids).toContain(initialId);
    expect(new Set(ids).size).toBe(2);
    const duplicatedId = ids.find((id) => id !== initialId);
    if (duplicatedId == null) throw new Error("Failed to find duplicated drawing id");

    const orig = afterDup.find((o) => o.id === initialId);
    const dup = afterDup.find((o) => o.id === duplicatedId);
    expect(dup).toBeTruthy();
    expect(orig).toBeTruthy();
    expect(dup.anchor.type).toBe(orig.anchor.type);
    if (dup.anchor.type === "absolute") {
      expect(dup.anchor.pos.xEmu - orig.anchor.pos.xEmu).toBe(EMU_PER_PX * 10);
      expect(dup.anchor.pos.yEmu - orig.anchor.pos.yEmu).toBe(EMU_PER_PX * 10);
    } else if (dup.anchor.type === "oneCell") {
      expect(dup.anchor.from.offset.xEmu - orig.anchor.from.offset.xEmu).toBe(EMU_PER_PX * 10);
      expect(dup.anchor.from.offset.yEmu - orig.anchor.from.offset.yEmu).toBe(EMU_PER_PX * 10);
    } else if (dup.anchor.type === "twoCell") {
      expect(dup.anchor.from.offset.xEmu - orig.anchor.from.offset.xEmu).toBe(EMU_PER_PX * 10);
      expect(dup.anchor.from.offset.yEmu - orig.anchor.from.offset.yEmu).toBe(EMU_PER_PX * 10);
      expect(dup.anchor.to.offset.xEmu - orig.anchor.to.offset.xEmu).toBe(EMU_PER_PX * 10);
      expect(dup.anchor.to.offset.yEmu - orig.anchor.to.offset.yEmu).toBe(EMU_PER_PX * 10);
    } else {
      throw new Error(`Unexpected anchor type: ${(dup.anchor as any).type}`);
    }

    // Duplicate selects the cloned object; Delete should remove it.
    await expect.poll(() => getSelectedDrawingId(page)).toBe(duplicatedId);
    await page.keyboard.press("Delete");

    await expect.poll(() => getDrawingObjects(page)).toHaveLength(1);
    const remaining = await getDrawingObjects(page);
    expect(remaining[0]?.id).toBe(initialId);
    await expect.poll(() => getSelectedDrawingId(page)).toBe(null);
  });

  test("right-clicking a drawing opens a drawing context menu", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared&drawingInteractions=1");
    await seedSingleShapeDrawing(page);

    await expect.poll(() => getDrawingObjects(page)).toHaveLength(1);

    await openContextMenuAt(page, { x: 100, y: 100 });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Delete" })).toBeVisible();
    await expect(menu.getByRole("button", { name: "Duplicate" })).toBeVisible();
    await expect(menu.getByRole("button", { name: "Bring Forward" })).toBeVisible();
    await expect(menu.getByRole("button", { name: "Send Backward" })).toBeVisible();
  });
});
