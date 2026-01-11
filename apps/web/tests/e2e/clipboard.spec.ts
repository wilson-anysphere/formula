import { expect, test, type Page } from "@playwright/test";

async function getGridGeometry(page: Page) {
  const selectionCanvas = page.getByTestId("canvas-grid-selection");
  await expect(selectionCanvas).toBeVisible({ timeout: 30_000 });

  const box = await selectionCanvas.boundingBox();
  expect(box).not.toBeNull();

  // Defaults from `VirtualScrollManager`: col width = 100, row height = 21.
  const headerWidth = 100;
  const headerHeight = 21;
  const colWidth = 100;
  const rowHeight = 21;

  const a1X = box!.x + headerWidth + colWidth / 2;
  const a1Y = box!.y + headerHeight + rowHeight / 2;

  return { selectionCanvas, box: box!, colWidth, rowHeight, a1X, a1Y };
}

async function dragSelect(page: Page, from: { x: number; y: number }, to: { x: number; y: number }) {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y);
  await page.mouse.up();
}

test("copies and pastes a rectangular grid selection via TSV clipboard payload", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  await page.evaluate(() => {
    (window as any).__lastCopy = { text: "", html: "" };
    document.addEventListener("copy", (event) => {
      const clipboard = (event as ClipboardEvent).clipboardData;
      (window as any).__lastCopy = {
        text: clipboard?.getData("text/plain") ?? "",
        html: clipboard?.getData("text/html") ?? ""
      };
    });
  });

  const { selectionCanvas, box, colWidth, rowHeight, a1X, a1Y } = await getGridGeometry(page);
  const a2Y = a1Y + rowHeight;

  // Select A1:A2 (workbook initializes A1=1, A2=2).
  await dragSelect(page, { x: a1X, y: a1Y }, { x: a1X, y: a2Y });

  await page.keyboard.press("ControlOrMeta+C");

  const clipboard = await page.evaluate(() => (window as any).__lastCopy as { text: string; html: string });
  expect(clipboard.text).toBe("1\n2");
  expect(clipboard.html).toContain("<table>");
  expect(clipboard.html).toContain("<td>1</td>");

  const c1X = a1X + colWidth * 2;
  await selectionCanvas.click({ position: { x: c1X - box.x, y: a1Y - box.y } });
  await expect(page.getByTestId("active-address")).toHaveText("C1");

  await page.keyboard.press("ControlOrMeta+V");

  await expect(page.getByTestId("formula-bar-value")).toHaveText("1");

  await selectionCanvas.click({ position: { x: c1X - box.x, y: a2Y - box.y } });
  await expect(page.getByTestId("active-address")).toHaveText("C2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("2");
});

test("cut clears the source range and preserves the clipboard for pasting", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const { selectionCanvas, box, colWidth, rowHeight, a1X, a1Y } = await getGridGeometry(page);
  const a2Y = a1Y + rowHeight;

  // Select A1:A2 and cut.
  await dragSelect(page, { x: a1X, y: a1Y }, { x: a1X, y: a2Y });
  await page.keyboard.press("ControlOrMeta+X");

  // Verify source cleared.
  await selectionCanvas.click({ position: { x: a1X - box.x, y: a1Y - box.y } });
  await expect(page.getByTestId("active-address")).toHaveText("A1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("");

  await selectionCanvas.click({ position: { x: a1X - box.x, y: a2Y - box.y } });
  await expect(page.getByTestId("active-address")).toHaveText("A2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("");

  // Paste into C1 and verify values.
  const c1X = a1X + colWidth * 2;
  await selectionCanvas.click({ position: { x: c1X - box.x, y: a1Y - box.y } });
  await expect(page.getByTestId("active-address")).toHaveText("C1");
  await page.keyboard.press("ControlOrMeta+V");

  await expect(page.getByTestId("formula-bar-value")).toHaveText("1");

  await selectionCanvas.click({ position: { x: c1X - box.x, y: a2Y - box.y } });
  await expect(page.getByTestId("active-address")).toHaveText("C2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("2");
});
