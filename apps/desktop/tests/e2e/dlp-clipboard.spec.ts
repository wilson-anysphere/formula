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

test.describe("DLP clipboard enforcement", () => {
  test("DLP-blocked copy/cut shows a toast and does not modify clipboard or sheet", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed A1:A2 with some content.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed DLP clipboard cells" });
      doc.setCellValue(sheetId, "A1", "RestrictedA1");
      doc.setCellValue(sheetId, "A2", "RestrictedA2");
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Mark the whole document as Restricted so clipboard.copy is blocked by the default policy.
    await page.evaluate(() => {
      const documentId = "local-workbook";
      const key = `dlp:classifications:${documentId}`;
      const record = {
        selector: { scope: "document", documentId },
        classification: { level: "Restricted", labels: [] },
        updatedAt: new Date().toISOString(),
      };
      localStorage.setItem(key, JSON.stringify([record]));
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

    const marker = `DLP_CLIPBOARD_MARKER_${Date.now()}`;
    await page.evaluate(async ({ marker }) => {
      await navigator.clipboard.writeText(marker);
    }, { marker });
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(marker);

    const toastLocator = page.getByTestId("toast");
    const toastCountBefore = await toastLocator.count();

    // Copy should be blocked: toast should appear and clipboard should remain unchanged.
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    await expect(toastLocator).toHaveCount(toastCountBefore + 1);
    await expect(page.getByTestId("toast-root")).toContainText("Clipboard copy is blocked");
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(marker);

    // Cut should also be blocked and should not clear the cells.
    await page.keyboard.press(`${modifier}+X`);
    await waitForIdle(page);

    await expect(toastLocator).toHaveCount(toastCountBefore + 2);
    await expect(page.getByTestId("toast-root")).toContainText("Clipboard copy is blocked");
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(marker);

    const { a1, a2 } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return { a1: app.getCellValueA1("A1"), a2: app.getCellValueA1("A2") };
    });
    expect(a1).toBe("RestrictedA1");
    expect(a2).toBe("RestrictedA2");
  });
});

