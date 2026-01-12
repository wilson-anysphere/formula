import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

const GRID_THEME_CSS_VAR_NAMES = [
  "--formula-grid-bg",
  "--formula-grid-line",
  "--formula-grid-header-bg",
  "--formula-grid-header-text",
  "--formula-grid-cell-text",
  "--formula-grid-error-text",
  "--formula-grid-selection-fill",
  "--formula-grid-selection-border",
  "--formula-grid-selection-handle",
  "--formula-grid-scrollbar-track",
  "--formula-grid-scrollbar-thumb",
  "--formula-grid-freeze-line",
  "--formula-grid-comment-indicator",
  "--formula-grid-comment-indicator-resolved",
  "--formula-grid-remote-presence-default",
] as const;

test.describe("grid theme css vars", () => {
  test("defines all @formula/grid theme vars and maps core values to desktop tokens", async ({ page }) => {
    await gotoDesktop(page);

    const { vars, resolved } = await page.evaluate((names) => {
      const style = getComputedStyle(document.documentElement);
      const values: Record<string, string> = {};
      for (const name of names) {
        values[name] = style.getPropertyValue(name).trim();
      }

      const resolveColor = (cssVar: string): string => {
        const probe = document.createElement("div");
        probe.style.position = "absolute";
        probe.style.width = "0";
        probe.style.height = "0";
        probe.style.overflow = "hidden";
        probe.style.pointerEvents = "none";
        probe.style.visibility = "hidden";
        probe.style.backgroundColor = `var(${cssVar})`;
        document.body.appendChild(probe);
        const resolved = getComputedStyle(probe).backgroundColor.trim();
        probe.remove();
        return resolved;
      };

      return {
        vars: values,
        resolved: {
          raw: {
            formulaGridBg: style.getPropertyValue("--formula-grid-bg").trim(),
            bgPrimary: style.getPropertyValue("--bg-primary").trim(),
            formulaGridLine: style.getPropertyValue("--formula-grid-line").trim(),
            gridLine: style.getPropertyValue("--grid-line").trim(),
          },
          gridBg: resolveColor("--formula-grid-bg"),
          bgPrimary: resolveColor("--bg-primary"),
          gridLine: resolveColor("--formula-grid-line"),
          tokenGridLine: resolveColor("--grid-line"),
        },
      };
    }, GRID_THEME_CSS_VAR_NAMES);

    for (const [name, value] of Object.entries(vars)) {
      expect(value, `${name} should be defined`).not.toBe("");
    }

    // Mapping verification (token parity).
    expect([resolved.raw.bgPrimary, "var(--bg-primary)"]).toContain(resolved.raw.formulaGridBg);
    expect([resolved.raw.gridLine, "var(--grid-line)"]).toContain(resolved.raw.formulaGridLine);

    expect(resolved.gridBg).toBe(resolved.bgPrimary);
    expect(resolved.gridLine).toBe(resolved.tokenGridLine);
  });
});
