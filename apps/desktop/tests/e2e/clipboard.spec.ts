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
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

type RichClipboardItem = {
  types: readonly string[];
  html: string | null;
  rtf: string | null;
};

async function readRichClipboardItem(
  page: import("@playwright/test").Page,
  {
    timeoutMs = 5_000,
    expectedHtmlSubstrings = [],
  }: { timeoutMs?: number; expectedHtmlSubstrings?: string[] } = {}
): Promise<RichClipboardItem> {
  const start = Date.now();
  let last: RichClipboardItem | null = null;
  let lastError: unknown;

  while (Date.now() - start < timeoutMs) {
    try {
      last = await page.evaluate(async () => {
        const items = await navigator.clipboard.read();
        const item = items[0];
        if (!item) return null;

        const types = item.types;
        const readType = async (type: string): Promise<string | null> => {
          if (!types.includes(type)) return null;
          const blob = await item.getType(type);
          return await blob.text();
        };

        return {
          types,
          html: await readType("text/html"),
          rtf: await readType("text/rtf"),
        };
      });

      if (
        last?.types.includes("text/html") &&
        last.html &&
        expectedHtmlSubstrings.every((substring) => last!.html!.includes(substring))
      ) {
        return last;
      }
    } catch (err) {
      lastError = err;
    }

    await page.waitForTimeout(100);
  }

  const errorMessage =
    lastError instanceof Error ? lastError.message : lastError ? String(lastError) : "Unknown clipboard read error";
  const types = last?.types?.length ? last.types.join(", ") : "none";
  throw new Error(`Timed out waiting for rich clipboard data. Last error: ${errorMessage}. Last types: ${types}`);
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

  test("copy/paste preserves inherited (effective) formatting from column defaults", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed clipboard column-default style" });
      doc.setCellValue(sheetId, "A1", "X");
      // Apply bold to the entire column A. With layered formats this should be stored
      // as a column default (so individual cells may still have styleId=0).
      doc.setRangeFormat(sheetId, "A1:A1048576", { font: { bold: true } }, { label: "Bold column" });
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Sanity check: A1 should *not* have a per-cell styleId (it inherits via column default),
    // otherwise this test wouldn't catch the regression.
    const { a1StyleId, a1EffectiveBold } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const a1 = doc.getCell(sheetId, "A1");
      const effective = doc.getCellFormat(sheetId, "A1");
      return { a1StyleId: a1.styleId, a1EffectiveBold: effective?.font?.bold === true };
    });
    expect(a1StyleId).toBe(0);
    expect(a1EffectiveBold).toBe(true);

    // Copy A1 and paste to B1.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    await page.click("#grid", { position: { x: 160, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    const b1Bold = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const b1 = doc.getCell(sheetId, "B1");
      const style = doc.styleTable.get(b1.styleId);
      return style?.font?.bold === true;
    });

    expect(b1Bold).toBe(true);
  });

  test("desktop copy writes rich clipboard formats (HTML + optional RTF)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    const seededValue = "RichClipboardA1";

    // Seed a 2x2 range with mixed value types.
    await page.evaluate(
      ({ seededValue }) => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();
        doc.beginBatch({ label: "Seed rich clipboard cells" });
        doc.setCellValue(sheetId, "A1", seededValue);
        doc.setCellValue(sheetId, "B1", 123);
        doc.setCellValue(sheetId, "A2", "RichClipboardA2");
        doc.setCellValue(sheetId, "B2", 456);
        doc.endBatch();
        app.refresh();
      },
      { seededValue }
    );
    await waitForIdle(page);

    // Select A1:B2 via drag and copy.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");
    await page.mouse.move(gridBox.x + 60, gridBox.y + 40); // A1
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 160, gridBox.y + 64); // B2
    await page.mouse.up();
    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");

    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    const { types, html, rtf } = await readRichClipboardItem(page, {
      expectedHtmlSubstrings: ["<table", seededValue],
    });

    expect(types).toContain("text/html");
    expect(html).toBeTruthy();
    expect(html!).toContain("<table");
    expect(html!).toContain(seededValue);

    if (types.includes("text/rtf")) {
      expect(rtf).toBeTruthy();
      expect(rtf!).toContain("\\rtf1");
      expect(rtf!).toContain(seededValue);
    } else {
      console.log(`[clipboard] text/rtf missing from clipboard types: ${types.join(", ")}`);
    }
  });

  test("Paste Special Values pastes the computed value (not the formula)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed A1 = 1, B1 = =A1+1 (-> 2), and format B1 so we can copy a formula+format cell.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed paste special values" });
      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellInput(sheetId, "B1", "=A1+1");
      doc.setRangeFormat(sheetId, "B1", { font: { bold: true } }, { label: "Bold" });
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Copy B1.
    await page.click("#grid", { position: { x: 160, y: 40 } }); // B1
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste Special Values into C1.
    await page.click("#grid", { position: { x: 260, y: 40 } }); // C1
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.keyboard.press(`${modifier}+Shift+V`);
    await expect(page.getByTestId("quick-pick")).toBeVisible();
    await page.getByRole("button", { name: "Paste Values" }).click();
    await waitForIdle(page);

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("2");

    const { formula, styleId } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, "C1");
      return { formula: cell.formula, styleId: cell.styleId };
    });
    expect(formula).toBeNull();
    // Paste Values should not paste formats; C1 should keep the default styleId.
    expect(styleId).toBe(0);
  });

  test("Paste Special Formulas pastes the formula (not formats)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed:
    // - A1 = 1
    // - B1 = =$A$1+1 (-> 2) + bold formatting (source)
    // - C1 = KeepStyle + italic formatting (destination)
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed paste special formulas" });
      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellInput(sheetId, "B1", "=$A$1+1");
      doc.setRangeFormat(sheetId, "B1", { font: { bold: true } }, { label: "Bold" });
      doc.setCellValue(sheetId, "C1", "KeepStyle");
      doc.setRangeFormat(sheetId, "C1", { font: { italic: true } }, { label: "Italic" });
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Copy B1.
    await page.click("#grid", { position: { x: 160, y: 40 } }); // B1
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste Special Formulas into C1.
    await page.click("#grid", { position: { x: 260, y: 40 } }); // C1
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.keyboard.press(`${modifier}+Shift+V`);
    await expect(page.getByTestId("quick-pick")).toBeVisible();
    await page.getByRole("button", { name: "Paste Formulas" }).click();
    await waitForIdle(page);

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("2");

    const { formula, bold, italic } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, "C1");
      const style = doc.styleTable.get(cell.styleId);
      return { formula: cell.formula, bold: style?.font?.bold === true, italic: style?.font?.italic === true };
    });

    expect(formula).toBe("=$A$1+1");
    // Paste Formulas should not paste formats; destination italic should remain,
    // and source bold should not be applied.
    expect(italic).toBe(true);
    expect(bold).toBe(false);
  });

  test("Paste Special Formats pastes formats (not values)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed:
    // - B1 = Styled + bold formatting (source)
    // - C1 = KeepValue + italic formatting (destination, should keep value but lose italic)
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed paste special formats" });
      doc.setCellValue(sheetId, "B1", "Styled");
      doc.setRangeFormat(sheetId, "B1", { font: { bold: true } }, { label: "Bold" });
      doc.setCellValue(sheetId, "C1", "KeepValue");
      doc.setRangeFormat(sheetId, "C1", { font: { italic: true } }, { label: "Italic" });
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Copy B1.
    await page.click("#grid", { position: { x: 160, y: 40 } }); // B1
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste Special Formats into C1.
    await page.click("#grid", { position: { x: 260, y: 40 } }); // C1
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.keyboard.press(`${modifier}+Shift+V`);
    await expect(page.getByTestId("quick-pick")).toBeVisible();
    await page.getByRole("button", { name: "Paste Formats" }).click();
    await waitForIdle(page);

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("KeepValue");

    const { formula, bold, italic } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, "C1");
      const style = doc.styleTable.get(cell.styleId);
      return { formula: cell.formula, bold: style?.font?.bold === true, italic: style?.font?.italic === true };
    });

    expect(formula).toBeNull();
    expect(bold).toBe(true);
    expect(italic).toBe(false);
  });

  test("DLP blocks spreadsheet copy/cut for Restricted ranges (toast shown, clipboard unchanged)", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    // Use a unique sentinel in case multiple Playwright workers run clipboard tests concurrently.
    const sentinel = `sentinel-${Date.now()}`;

    // Ensure a clean slate in case previous tests left transient UI messages around.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    // Seed A1 = Secret.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed DLP clipboard cells" });
      doc.setCellValue(sheetId, "A1", "Secret");
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Mark A1:A1 as Restricted in the local classification store (keyed by workbook id).
    await page.evaluate(() => {
      const docIdParam = new URL(window.location.href).searchParams.get("docId");
      const workbookId = typeof docIdParam === "string" && docIdParam.trim() !== "" ? docIdParam.trim() : "local-workbook";
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();

      const record = {
        selector: {
          scope: "range",
          documentId: workbookId,
          sheetId,
          range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }, // A1
        },
        classification: { level: "Restricted", labels: [] },
        updatedAt: new Date().toISOString(),
      };

      localStorage.setItem(`dlp:classifications:${workbookId}`, JSON.stringify([record]));
    });

    // Seed clipboard with sentinel value so we can verify copy/cut do not overwrite it.
    // Some clipboard implementations require a user gesture even with permissions granted,
    // so fall back to legacy DOM copy if `navigator.clipboard.writeText` fails.
    await page.evaluate(async (text) => {
      try {
        await navigator.clipboard.writeText(text);
        return;
      } catch {
        // Fall back to legacy DOM copy.
      }

      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.style.position = "fixed";
      textarea.style.left = "-9999px";
      textarea.style.top = "0";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      const ok = document.execCommand("copy");
      textarea.remove();
      if (!ok) throw new Error("Failed to seed clipboard with sentinel text");
    }, sentinel);
    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe(sentinel);

    // Select A1 and attempt copy (should be blocked).
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      });
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1");

    await page.keyboard.press(`${modifier}+C`);

    const toastRoot = page.getByTestId("toast-root");
    const copyToast = toastRoot.getByTestId("toast").last();
    await expect(copyToast).toBeVisible();
    await expect(copyToast).toHaveAttribute("data-type", "warning");
    await expect(copyToast).toContainText(/clipboard copy is blocked|data loss prevention/i);
    await expect(copyToast).toContainText("Restricted");
    await expect(copyToast).toContainText("Confidential");

    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe(sentinel);

    // Clear the existing toast so we can assert cut creates its own message.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    // Attempt cut (should be blocked: clipboard unchanged and cell not cleared).
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      });
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press(`${modifier}+X`);

    const cutToast = toastRoot.getByTestId("toast").last();
    await expect(cutToast).toBeVisible();
    await expect(cutToast).toHaveAttribute("data-type", "warning");
    await expect(cutToast).toContainText(/clipboard copy is blocked|data loss prevention/i);
    await expect(cutToast).toContainText("Restricted");
    await expect(cutToast).toContainText("Confidential");

    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe(sentinel);

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")))
      .toBe("Secret");
  });
});
