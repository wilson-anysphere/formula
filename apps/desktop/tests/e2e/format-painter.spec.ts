import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function clickCellA1(page: import("@playwright/test").Page, addr: string): Promise<void> {
  await page.waitForFunction(
    (a1) => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.(a1);
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    },
    addr,
    { timeout: 10_000 },
  );
  const rect = await page.evaluate((a1) => (window as any).__formulaApp.getCellRectA1(a1), addr);
  await page.click("#grid", { position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } });
}

test.describe("Format Painter (Home → Clipboard)", () => {
  test("applies formatting to the next selection once (one-shot)", async ({ page }) => {
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "Home" }).click();

    // Seed A1 with visible formatting (bold + fill color).
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed format painter" });
      doc.setRangeFormat(
        sheetId,
        "A1",
        { font: { bold: true }, fill: { pattern: "solid", fgColor: "FFFF0000" } },
        { label: "Seed format painter" },
      );
      doc.endBatch();
      app.refresh();
    });

    // Select A1 and arm format painter.
    await clickCellA1(page, "A1");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await ribbon.getByRole("button", { name: "Format Painter" }).click();

    // Select B2 to apply.
    await clickCellA1(page, "B2");
    await expect(page.getByTestId("active-cell")).toHaveText("B2");

    await expect
      .poll(
        () =>
          page.evaluate(() => {
            const app = (window as any).__formulaApp;
            const doc = app.getDocument();
            const sheetId = app.getCurrentSheetId();
            const style = doc.getCellFormat(sheetId, "B2") as any;
            const cell = doc.getCell(sheetId, "B2") as any;
            return {
              styleId: cell?.styleId ?? 0,
              bold: Boolean(style?.font?.bold),
              fill: style?.fill?.fgColor ?? null,
            };
          }),
        { timeout: 5_000 },
      )
      .toEqual({ styleId: expect.any(Number), bold: true, fill: "FFFF0000" });

    const b2StyleId = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      return (doc.getCell(sheetId, "B2") as any)?.styleId ?? 0;
    });
    expect(b2StyleId).toBeGreaterThan(0);

    // Move selection again; Format Painter should have disarmed (one-shot).
    await clickCellA1(page, "C3");
    await expect(page.getByTestId("active-cell")).toHaveText("C3");

    // Wait longer than the Format Painter selection debounce. If the mode wasn't disarmed,
    // C3 would have received the formatting by now.
    await page.waitForTimeout(250);

    const c3 = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const style = doc.getCellFormat(sheetId, "C3") as any;
      const cell = doc.getCell(sheetId, "C3") as any;
      return {
        styleId: cell?.styleId ?? 0,
        bold: Boolean(style?.font?.bold),
        fill: style?.fill?.fgColor ?? null,
      };
    });

    expect(c3.styleId).toBe(0);
    expect(c3.bold).toBe(false);
    expect(c3.fill).toBeNull();
  });
});

