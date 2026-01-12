import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("fill shortcuts (Ctrl/Cmd+D, Ctrl/Cmd+R)", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Ctrl+D fills down from top row (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } });
        app.focus();
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+D");
      await waitForIdle(page);

      const [a2, a3] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3")),
      ]);
      expect(a2).toBe("1");
      expect(a3).toBe("1");
    });

    test(`Ctrl+D fills down for each selected range and is a single undo step (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      const beforeUndoDepth = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        // Seed values in one batch so the undo depth baseline is stable.
        doc.beginBatch({ label: "Seed" });
        doc.setCellValue(sheetId, "E10", 1);
        doc.setCellValue(sheetId, "G10", 3);
        doc.endBatch();

        // Create a multi-range selection: E10:E12 and G10:G12 (outside the default seeded demo data).
        (app as any).selection = {
          type: "multi",
          ranges: [
            { startRow: 9, endRow: 11, startCol: 4, endCol: 4 },
            { startRow: 9, endRow: 11, startCol: 6, endCol: 6 },
          ],
          active: { row: 9, col: 4 },
          anchor: { row: 9, col: 4 },
          activeRangeIndex: 0,
        };

        if ((app as any).sharedGrid) {
          (app as any).syncSharedGridSelectionFromState();
        }
        app.refresh();
        app.focus();

        return doc.getStackDepths().undo;
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+D");
      await waitForIdle(page);

      const [e11, e12, g11, g12] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E11")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E12")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("G11")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("G12")),
      ]);
      expect(e11).toBe("1");
      expect(e12).toBe("1");
      expect(g11).toBe("3");
      expect(g12).toBe("3");

      const [undoDepthAfterFill, undoLabel] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getDocument().getStackDepths().undo),
        page.evaluate(() => (window as any).__formulaApp.getDocument().undoLabel),
      ]);
      expect(undoDepthAfterFill - beforeUndoDepth).toBe(1);
      expect(undoLabel).toBe("Fill Down");

      await page.evaluate(() => (window as any).__formulaApp.undo());
      await waitForIdle(page);

      const [e11AfterUndo, e12AfterUndo, g11AfterUndo, g12AfterUndo, e10AfterUndo, g10AfterUndo] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E11")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E12")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("G11")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("G12")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E10")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("G10")),
      ]);

      expect(e11AfterUndo).toBe("");
      expect(e12AfterUndo).toBe("");
      expect(g11AfterUndo).toBe("");
      expect(g12AfterUndo).toBe("");
      // Seeded source row remains.
      expect(e10AfterUndo).toBe("1");
      expect(g10AfterUndo).toBe("3");
    });

    test(`Ctrl+R fills right from leftmost column (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "B1", 2);
        app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 2 } });
        app.focus();
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+R");
      await waitForIdle(page);

      const [b1, c1] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1")),
      ]);
      expect(b1).toBe("1");
      expect(c1).toBe("1");
    });

    test(`Ctrl+R fills right for each selected range and is a single undo step (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      const beforeUndoDepth = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        doc.beginBatch({ label: "Seed" });
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "B1", 2);
        doc.setCellValue(sheetId, "C1", 3);
        doc.setCellValue(sheetId, "A3", 4);
        doc.setCellValue(sheetId, "B3", 5);
        doc.setCellValue(sheetId, "C3", 6);
        doc.endBatch();

        // Multi-range selection: A1:C1 and A3:C3.
        (app as any).selection = {
          type: "multi",
          ranges: [
            { startRow: 0, endRow: 0, startCol: 0, endCol: 2 },
            { startRow: 2, endRow: 2, startCol: 0, endCol: 2 },
          ],
          active: { row: 0, col: 0 },
          anchor: { row: 0, col: 0 },
          activeRangeIndex: 0,
        };

        if ((app as any).sharedGrid) {
          (app as any).syncSharedGridSelectionFromState();
        }
        app.refresh();
        app.focus();

        return doc.getStackDepths().undo;
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+R");
      await waitForIdle(page);

      const [b1, c1, b3, c3] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B3")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C3")),
      ]);
      expect(b1).toBe("1");
      expect(c1).toBe("1");
      expect(b3).toBe("4");
      expect(c3).toBe("4");

      const [undoDepthAfterFill, undoLabel] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getDocument().getStackDepths().undo),
        page.evaluate(() => (window as any).__formulaApp.getDocument().undoLabel),
      ]);
      expect(undoDepthAfterFill - beforeUndoDepth).toBe(1);
      expect(undoLabel).toBe("Fill Right");

      await page.evaluate(() => (window as any).__formulaApp.undo());
      await waitForIdle(page);

      const [b1AfterUndo, c1AfterUndo, b3AfterUndo, c3AfterUndo] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B3")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C3")),
      ]);
      expect(b1AfterUndo).toBe("2");
      expect(c1AfterUndo).toBe("3");
      expect(b3AfterUndo).toBe("5");
      expect(c3AfterUndo).toBe("6");
    });

    test(`Ctrl+D shifts formulas (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "B1", 2);
        doc.setCellInput(sheetId, "C1", "=A1+B1");
        app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } });
        app.focus();
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+D");
      await waitForIdle(page);

      const [c2Formula, c3Formula] = await Promise.all([
        page.evaluate(
          () =>
            (window as any).__formulaApp.getDocument().getCell((window as any).__formulaApp.getCurrentSheetId(), "C2").formula
        ),
        page.evaluate(
          () =>
            (window as any).__formulaApp.getDocument().getCell((window as any).__formulaApp.getCurrentSheetId(), "C3").formula
        ),
      ]);

      expect(c2Formula).toBe("=A2+B2");
      expect(c3Formula).toBe("=A3+B3");
    });
  }
});
