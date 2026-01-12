import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("load");
        continue;
      }
      throw err;
    }
  }
}

test.describe("clipboard shortcuts (copy/cut/paste)", () => {
  test("Ctrl/Cmd+C copies selection and Ctrl/Cmd+V pastes starting at active cell", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed A1 = Hello, A2 = World.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed clipboard cells" });
      doc.setCellValue(sheetId, "A1", "Hello");
      doc.setCellValue(sheetId, "A2", "World");
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Select A1:A2 via drag.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");
    await page.mouse.move(gridBox.x + 60, gridBox.y + 40); // A1
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 60, gridBox.y + 64); // A2
    await page.mouse.up();

    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste into C1.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    // Paste updates the selection to match the pasted dimensions.
    await expect(page.getByTestId("selection-range")).toHaveText("C1:C2");

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("Hello");
    const c2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"));
    expect(c2Value).toBe("World");

    // Paste should be undoable as a single history entry.
    await page.keyboard.press(`${modifier}+Z`);
    await waitForIdle(page);
    const c1AfterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1AfterUndo).toBe("");
    const c2AfterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"));
    expect(c2AfterUndo).toBe("");

    // Redo should restore the pasted values.
    await page.keyboard.press(`${modifier}+Shift+Z`);
    await waitForIdle(page);
    const c1AfterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1AfterRedo).toBe("Hello");
    const c2AfterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"));
    expect(c2AfterRedo).toBe("World");

    // Cut A1 and paste to B1.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await page.keyboard.press(`${modifier}+X`);
    await waitForIdle(page);

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")))
      .toBe("");

    await page.click("#grid", { position: { x: 160, y: 40 } });
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")))
      .toBe("Hello");
  });

  test("copy/paste shifts relative references inside formulas (Excel-style)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed a simple scenario where shifting is observable:
    // B1 = A1 + 1, and A2 has a different value so pasting down should change the result.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed clipboard formula shift" });
      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 10);
      doc.setCellInput(sheetId, "B1", "=A1+1");
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Copy B1.
    await page.click("#grid", { position: { x: 160, y: 40 } }); // B1
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste into B2: formula should shift A1 -> A2, so computed value becomes 11.
    await page.click("#grid", { position: { x: 160, y: 64 } }); // B2
    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    const b2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B2"));
    expect(b2Value).toBe("11");
  });

  test("copy/paste preserves internal styleId for DocumentController formats", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed clipboard styles" });
      doc.setCellValue(sheetId, "A1", "Styled");
      doc.setRangeFormat(sheetId, "A1", { font: { bold: true } }, { label: "Bold" });
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Copy A1 and paste to B1.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    await page.click("#grid", { position: { x: 160, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    const { a1StyleId, b1StyleId } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const a1 = doc.getCell(sheetId, "A1");
      const b1 = doc.getCell(sheetId, "B1");
      return { a1StyleId: a1.styleId, b1StyleId: b1.styleId };
    });

    expect(b1StyleId).toBe(a1StyleId);
  });
});
