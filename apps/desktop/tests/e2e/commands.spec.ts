import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Built-in commands", () => {
  test("view.toggleShowFormulas toggles between computed value and formula text", async ({ page }) => {
    await gotoDesktop(page);

    // `__formulaCommandRegistry` is assigned later in `main.ts` than `__formulaApp`.
    // Wait explicitly so the test doesn't race startup.
    await page.waitForFunction(() => Boolean((window as any).__formulaCommandRegistry), undefined, { timeout: 10_000 });

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellFormula(sheetId, { row: 0, col: 0 }, "=1+1");
      app.refresh();
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

  test("edit.undo/edit.redo execute through the command registry", async ({ page }) => {
    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__formulaCommandRegistry), undefined, { timeout: 10_000 });

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, "A1", "UndoRedoTest", { label: "Set A1" });
      app.refresh();
    });

    const edited = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(edited).toBe("UndoRedoTest");

    await page.evaluate(async () => {
      const registry = (window as any).__formulaCommandRegistry;
      await registry.executeCommand("edit.undo");
      await (window as any).__formulaApp.whenIdle();
    });

    const afterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(afterUndo).toBe(before);

    await page.evaluate(async () => {
      const registry = (window as any).__formulaCommandRegistry;
      await registry.executeCommand("edit.redo");
      await (window as any).__formulaApp.whenIdle();
    });

    const afterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(afterRedo).toBe("UndoRedoTest");
  });
});
