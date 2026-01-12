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

async function writeClipboardText(page: import("@playwright/test").Page, text: string): Promise<void> {
  await page.evaluate(async (raw) => {
    try {
      await navigator.clipboard.writeText(raw);
      return;
    } catch {
      // Fall back to legacy DOM copy (some environments still require a user gesture).
    }

    const textarea = document.createElement("textarea");
    textarea.value = raw;
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    const ok = document.execCommand("copy");
    textarea.remove();
    if (!ok) throw new Error("Failed to write clipboard text (fallback copy)");
  }, text);
}

test.describe("keybindings: input scoping policy", () => {
  test("command palette opens while formula bar input is focused", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Focus the (hidden) formula bar textarea via the highlight affordance.
    await page.getByTestId("formula-highlight").click();
    await expect(page.getByTestId("formula-input")).toBeFocused();

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
  });

  test("clipboard shortcuts still operate inside the formula bar input (not intercepted by spreadsheet commands)", async ({
    page,
  }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const text = `Hello-${Date.now()}`;

    // Seed a couple cells so we can prove nothing changes as a side-effect of Ctrl/Cmd+C/V
    // while the formula bar is focused.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed keybinding-input clipboard scenario" });
      doc.setCellValue(sheetId, "A1", "SeedA1");
      doc.setCellValue(sheetId, "B1", "SeedB1");
      doc.endBatch();
      app.refresh();
    });
    await waitForIdle(page);

    // Pick C1 as the active cell so we can detect unintended spreadsheet paste into the grid.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    // Focus formula bar textarea and type some text.
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeFocused();
    await input.fill(text);

    // Select the textarea contents via DOM APIs (more reliable than Ctrl/Cmd+A in headless).
    await page.evaluate(() => {
      const el = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
      if (!el) throw new Error("Missing formula bar input");
      el.focus();
      el.setSelectionRange(0, el.value.length);
    });

    // Ctrl/Cmd+C should copy the textarea selection (native browser behavior), not the sheet selection.
    await page.keyboard.press(`${modifier}+C`);

    await expect
      .poll(() => page.evaluate(async () => (await navigator.clipboard.readText()).trim()), { timeout: 10_000 })
      .toBe(text);

    // Move the caret to the end and ensure clipboard contains the expected text before pasting
    // (avoids flakiness when Playwright workers share the OS clipboard).
    await page.evaluate(() => {
      const el = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
      if (!el) throw new Error("Missing formula bar input");
      const end = el.value.length;
      el.focus();
      el.setSelectionRange(end, end);
    });
    await writeClipboardText(page, text);

    await page.keyboard.press(`${modifier}+V`);
    await expect(input).toHaveValue(`${text}${text}`);

    // Cancel formula editing so we can assert against committed cell values.
    await page.keyboard.press("Escape");
    await waitForIdle(page);

    const { a1, b1, c1 } = await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      return {
        a1: await app.getCellValueA1("A1"),
        b1: await app.getCellValueA1("B1"),
        c1: await app.getCellValueA1("C1"),
      };
    });

    expect(a1).toBe("SeedA1");
    expect(b1).toBe("SeedB1");
    // Spreadsheet clipboard paste should *not* run while the formula bar is focused.
    expect(c1).toBe("");
  });
});
