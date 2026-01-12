import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Built-in commands", () => {
  test("view.toggleShowFormulas toggles between computed value and formula text", async ({ page }) => {
    await gotoDesktop(page);

    // `__formulaCommandRegistry` is assigned later in `main.ts` than `__formulaApp`.
    // Wait explicitly so the test doesn't race startup.
    await page.waitForFunction(() => Boolean(window.__formulaCommandRegistry), undefined, { timeout: 10_000 });

    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellFormula(sheetId, { row: 0, col: 0 }, "=1+1");
      app.refresh();
    });

    const before = await page.evaluate(async () => {
      const app = window.__formulaApp as any;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(before).toBe("2");

    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      if (!registry) throw new Error("Missing window.__formulaCommandRegistry (desktop e2e harness)");
      await registry.executeCommand("view.toggleShowFormulas");
    });

    const after = await page.evaluate(async () => {
      const app = window.__formulaApp as any;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(after).toBe("=1+1");

    // Toggle back to computed values.
    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("view.toggleShowFormulas");
    });

    const final = await page.evaluate(async () => {
      const app = window.__formulaApp as any;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(final).toBe("2");
  });

  test("edit.undo/edit.redo execute through the command registry", async ({ page }) => {
    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean(window.__formulaCommandRegistry), undefined, { timeout: 10_000 });

    const before = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));

    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, "A1", "UndoRedoTest", { label: "Set A1" });
      app.refresh();
    });

    const edited = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));
    expect(edited).toBe("UndoRedoTest");

    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("edit.undo");
      await (window.__formulaApp as any).whenIdle();
    });

    const afterUndo = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));
    expect(afterUndo).toBe(before);

    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("edit.redo");
      await (window.__formulaApp as any).whenIdle();
    });

    const afterRedo = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));
    expect(afterRedo).toBe("UndoRedoTest");
  });

  test("view.togglePanel.versionHistory + view.togglePanel.branchManager open/close panels", async ({ page }) => {
    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean(window.__formulaCommandRegistry), undefined, { timeout: 10_000 });

    // Version History.
    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("view.togglePanel.versionHistory");
    });
    await expect(page.getByTestId("panel-versionHistory")).toBeVisible();

    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("view.togglePanel.versionHistory");
    });
    await expect(page.getByTestId("panel-versionHistory")).toHaveCount(0);

    // Branch Manager.
    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("view.togglePanel.branchManager");
    });
    await expect(page.getByTestId("panel-branchManager")).toBeVisible();

    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("view.togglePanel.branchManager");
    });
    await expect(page.getByTestId("panel-branchManager")).toHaveCount(0);
  });
});
