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
      await page.addInitScript(({ extensionId }) => {
        // Seed granted permissions so the extension can activate without prompts.
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();
        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "ui.panels": true,
          "cells.read": true,
          "cells.write": true,
          network: { mode: "full" },
        };
        localStorage.setItem(key, JSON.stringify(existing));
      }, { extensionId });

      await gotoDesktop(page);

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("mode: full");

      page.once("dialog", async (dialog) => {
        expect(dialog.type()).toBe("prompt");
        await dialog.accept(JSON.stringify([url]));
      });
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");

      await page.getByTestId(`revoke-all-permissions-${extensionId}`).click();
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

      page.once("dialog", async (dialog) => {
        expect(dialog.type()).toBe("prompt");
        await dialog.accept(JSON.stringify([url]));
      });
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });
});
