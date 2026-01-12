import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Built-in commands", () => {
  test("view.toggleShowFormulas toggles between computed value and formula text", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellFormula(sheetId, { row: 0, col: 0 }, "=1+1");
    });

    const before = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(before).toBe("2");

    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const registry: any = (window as any).__formulaCommandRegistry;
      if (!registry) throw new Error("Missing window.__formulaCommandRegistry (desktop e2e harness)");
      await registry.executeCommand("view.toggleShowFormulas");
    });

    const after = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(after).toBe("=1+1");

    // Toggle back to computed values.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const registry: any = (window as any).__formulaCommandRegistry;
      await registry.executeCommand("view.toggleShowFormulas");
    });

    const final = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(final).toBe("2");
  });
});

