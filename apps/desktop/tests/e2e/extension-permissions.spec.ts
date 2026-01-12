import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

const PERMISSIONS_KEY = "formula.extensionHost.permissions";
const SAMPLE_HELLO_ID = "formula.sample-hello";

async function setExtensionPermissionGrants(page: Page, grants: Record<string, any>): Promise<void> {
  await page.evaluate(
    ({ key, grants }) => {
      localStorage.setItem(key, JSON.stringify(grants));
    },
    { key: PERMISSIONS_KEY, grants },
  );
}

test.describe("Extension permission prompts (desktop UI)", () => {
  test("denying clipboard permission shows a prompt and surfaces an error toast", async ({ page }) => {
    await gotoDesktop(page);

    // Pre-grant only what's needed to activate + read selection so the first prompt we see is
    // specifically for `clipboard`.
    await setExtensionPermissionGrants(page, {
      [SAMPLE_HELLO_ID]: {
        "ui.commands": true,
        "cells.read": true,
      },
    });

    await page.getByTestId("open-extensions-panel").click();
    const runButton = page.getByTestId("run-command-sampleHello.copySumToClipboard");
    await expect(runButton).toBeVisible({ timeout: 30_000 });
    // Avoid hit-target flakiness from fixed overlays (status bar/sheet tabs) by
    // dispatching a click directly.
    await runButton.dispatchEvent("click");

    const prompt = page.getByTestId("extension-permission-prompt");
    await expect(prompt).toBeVisible({ timeout: 30_000 });
    await expect(prompt).toContainText("Sample Hello");
    await expect(prompt).toContainText("clipboard");

    await page.getByTestId("extension-permission-deny").click();
    await expect(prompt).toHaveCount(0);

    await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
  });
});
