import { expect, test } from "@playwright/test";
import http from "node:http";

import { gotoDesktop } from "./helpers";

test.describe("Extensions permissions UI", () => {
  test("can view and revoke extension network permission from the Extensions panel", async ({ page }) => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello");
    });

    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : null;
    if (!port) throw new Error("Failed to allocate test port");

    const url = `http://127.0.0.1:${port}/`;
    const extensionId = "formula.sample-hello";

    try {
      await page.addInitScript(() => {
        // Start with a clean permission store so this test exercises the allow/deny UI prompt flow.
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await gotoDesktop(page);

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-ui.commands")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId(`permission-${extensionId}-ui.commands`)).toBeVisible();
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("mode: allowlist");
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("127.0.0.1");

      // Revoke only network permission; ensure other grants remain.
      await page.getByTestId(`revoke-permission-${extensionId}-network`).click();
      await expect(page.getByTestId(`permission-${extensionId}-ui.commands`)).toBeVisible();
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toHaveCount(0);

      // Re-run and deny only the network prompt.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permission-${extensionId}-ui.commands`)).toBeVisible();
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toHaveCount(0);

      await page.getByTestId(`revoke-all-permissions-${extensionId}`).click();
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });
});
