import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForGridFocus(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => (document.activeElement as HTMLElement | null)?.id === "grid", null, { timeout: 5_000 });
}

test.describe("ribbon clipboard (Home → Clipboard)", () => {
  test("Cut/Copy/Paste buttons operate on the spreadsheet selection", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");

    // Ensure we're on the Home tab (future-proofing if another test changes default tab persistence).
    await ribbon.getByRole("tab", { name: "Home" }).click();

    const marker = `__formula_ribbon_clipboard__${Math.random().toString(16).slice(2)}`;
    await page.evaluate(async (text) => {
      await navigator.clipboard.writeText(text);
    }, marker);

    // Seed A1 = Hello, A2 = World.
    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed ribbon clipboard cells" });
      doc.setCellValue(sheetId, "A1", "Hello");
      doc.setCellValue(sheetId, "A2", "World");
      doc.endBatch();
      app.refresh();
    });

    // Select A1:A2 via drag.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");
    await page.mouse.move(gridBox.x + 60, gridBox.y + 40); // A1
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 60, gridBox.y + 64); // A2
    await page.mouse.up();
    await expect(page.getByTestId("selection-range")).toHaveText("A1:A2");

    // Copy via ribbon.
    await ribbon.getByRole("button", { name: "Copy" }).click();
    await waitForGridFocus(page);

    // Assert the clipboard changed from the marker value.
    await expect
      .poll(
        () =>
          page.evaluate(async () => {
            return await navigator.clipboard.readText();
          }),
        { timeout: 5_000 },
      )
      .not.toBe(marker);

    // Paste into C1 via ribbon dropdown.
    await page.click("#grid", { position: { x: 260, y: 40 } }); // C1
    await expect(page.getByTestId("active-cell")).toHaveText("C1");
    await ribbon.getByTestId("ribbon-paste").click();
    await ribbon.getByRole("menuitem", { name: "Paste", exact: true }).click();
    await waitForGridFocus(page);

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("C1"))).toBe("Hello");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("C2"))).toBe("World");

    // Cut the original A1:A2 via ribbon, then paste into E1 to verify cut+paste behavior.
    await page.click("#grid", { position: { x: 53, y: 29 } }); // A1
    await page.mouse.move(gridBox.x + 60, gridBox.y + 40); // A1
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 60, gridBox.y + 64); // A2
    await page.mouse.up();
    await expect(page.getByTestId("selection-range")).toHaveText("A1:A2");

    await ribbon.getByRole("button", { name: "Cut" }).click();
    await waitForGridFocus(page);

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"))).toBe("");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A2"))).toBe("");

    await page.click("#grid", { position: { x: 460, y: 40 } }); // E1
    await expect(page.getByTestId("active-cell")).toHaveText("E1");
    await ribbon.getByTestId("ribbon-paste").click();
    await ribbon.getByRole("menuitem", { name: "Paste", exact: true }).click();
    await waitForGridFocus(page);

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("E1"))).toBe("Hello");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("E2"))).toBe("World");
  });

  test("Paste Values from the Paste dropdown uses paste-special values semantics", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "Home" }).click();

    // Seed A1 = 1, B1 = =A1+1 (-> 2) with bold formatting.
    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed ribbon paste values" });
      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellInput(sheetId, "B1", "=A1+1");
      doc.setRangeFormat(sheetId, "B1", { font: { bold: true } }, { label: "Bold" });
      doc.endBatch();
      app.refresh();
    });

    // Copy B1 via ribbon.
    await page.click("#grid", { position: { x: 160, y: 40 } }); // B1
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await ribbon.getByRole("button", { name: "Copy" }).click();
    await waitForGridFocus(page);

    // Paste Values into C1 using the Paste dropdown.
    await page.click("#grid", { position: { x: 260, y: 40 } }); // C1
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await ribbon.getByTestId("ribbon-paste").click();
    await ribbon.getByRole("menuitem", { name: "Paste Values" }).click();
    await waitForGridFocus(page);

    const c1Value = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("C1"));
    expect(c1Value).toBe("2");

    const { formula, styleId } = await page.evaluate(() => {
      const app = window.__formulaApp as any;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, "C1");
      return { formula: cell.formula, styleId: cell.styleId };
    });
    expect(formula).toBeNull();
    // Paste Values should not paste formats; C1 should keep the default styleId.
    expect(styleId).toBe(0);
  });
});
