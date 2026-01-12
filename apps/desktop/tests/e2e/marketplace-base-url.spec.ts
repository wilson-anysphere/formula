import { expect, test } from "@playwright/test";
import http from "node:http";
import type { AddressInfo } from "node:net";

import { gotoDesktop } from "./helpers";

test.describe("Marketplace base URL configuration", () => {
  test("can override the marketplace base URL (origin) and requests normalize to /api", async ({ page }) => {
    let sawSearch: ((value: string) => void) | null = null;
    const sawSearchPromise = new Promise<string>((resolve) => {
      sawSearch = resolve;
    });
    let sawSearchOnce = false;

    const server = http.createServer((req, res) => {
      const method = req.method || "GET";
      const url = new URL(req.url || "/", "http://localhost");

      const allowCors = () => {
        res.setHeader("Access-Control-Allow-Origin", "*");
        res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
        res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
      };

      if (method === "OPTIONS") {
        allowCors();
        res.writeHead(204);
        res.end();
        return;
      }

      if (method === "GET" && url.pathname === "/api/search") {
        if (!sawSearchOnce) {
          sawSearchOnce = true;
          sawSearch?.(req.url || "");
        }

        allowCors();
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(
          JSON.stringify(
            {
              total: 1,
              results: [
                {
                  id: "test.publisher",
                  name: "publisher",
                  displayName: "Test Publisher",
                  publisher: "test",
                  description: "stub extension",
                  latestVersion: "1.0.0",
                  verified: true,
                  featured: false,
                  categories: [],
                  tags: [],
                  screenshots: [],
                  downloadCount: 0,
                  updatedAt: new Date().toISOString(),
                },
              ],
              nextCursor: null,
            },
            null,
            2,
          ),
        );
        return;
      }

      allowCors();
      res.writeHead(404, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ error: "Not found" }));
    });

    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const port = (server.address() as AddressInfo).port;
    const origin = `http://127.0.0.1:${port}`;

    try {
      // Apply the override before the app's scripts run so the panel wiring picks it up.
      await page.addInitScript(({ origin }) => {
        try {
          localStorage.setItem("formula:marketplace:baseUrl", origin);
          // The built-in `formula.e2e-events` extension activates during startup and writes
          // event traces into extension storage. Pre-grant its `storage` permission so the
          // permission prompt doesn't block unrelated marketplace assertions.
          localStorage.setItem(
            "formula.extensionHost.permissions",
            JSON.stringify({
              "formula.e2e-events": { storage: true },
            }),
          );
        } catch {
          // ignore
        }
      }, { origin });

      await gotoDesktop(page);

      await page.getByRole("tab", { name: "View", exact: true }).click();
      await page.getByTestId("open-marketplace-panel").click();

      const panel = page.getByTestId("panel-marketplace");
      await expect(panel).toBeVisible();

      await panel.locator('input[type="search"]').fill("test");
      await panel.getByRole("button", { name: "Search" }).click();

      const requestUrl = await Promise.race([
        sawSearchPromise,
        new Promise<string>((_resolve, reject) =>
          setTimeout(() => reject(new Error("Timed out waiting for /api/search request")), 10_000),
        ),
      ]);
      expect(requestUrl).toContain("/api/search");

      // Prove the response was actually read by the UI.
      await expect(panel).toContainText("Test Publisher (test.publisher)");
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });
});
