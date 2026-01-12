import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("grid theme css vars", () => {
  test("desktop defines --formula-grid-* tokens (and they respond to theme changes)", async ({ page }) => {
    await gotoDesktop(page);

    const selectionBorderVar = await page.evaluate(() =>
      getComputedStyle(document.documentElement).getPropertyValue("--formula-grid-selection-border").trim(),
    );
    expect(selectionBorderVar).not.toEqual("");

    const resolveSelectionBorderColor = async () =>
      page.evaluate(() => {
        const probe = document.createElement("div");
        probe.style.color = "var(--formula-grid-selection-border)";
        document.body.appendChild(probe);
        const color = getComputedStyle(probe).color;
        probe.remove();
        return color;
      });

    // Force light -> dark by toggling the theme attribute (desktop shells set this at runtime).
    await page.evaluate(() => document.documentElement.removeAttribute("data-theme"));
    const light = await resolveSelectionBorderColor();

    await page.evaluate(() => document.documentElement.setAttribute("data-theme", "dark"));
    await page.waitForFunction(
      (prev) => {
        const probe = document.createElement("div");
        probe.style.color = "var(--formula-grid-selection-border)";
        document.body.appendChild(probe);
        const color = getComputedStyle(probe).color;
        probe.remove();
        return color !== prev;
      },
      light,
      { timeout: 5_000 },
    );

    const dark = await resolveSelectionBorderColor();
    expect(dark).not.toEqual(light);
  });
});

