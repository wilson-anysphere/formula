// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { resolveGridThemeFromCssVars } from "../resolveThemeFromCssVars";

describe("resolveGridThemeFromCssVars (DOM)", () => {
  it("normalizes var() indirection to a computed color string", () => {
    const host = document.createElement("div");
    host.style.setProperty("--app-bg", "rgb(10, 20, 30)");
    host.style.setProperty("--formula-grid-bg", "var(--app-bg)");
    document.body.appendChild(host);

    try {
      const theme = resolveGridThemeFromCssVars(host);
      expect(theme.gridBg).toBe("rgb(10, 20, 30)");
    } finally {
      host.remove();
    }
  });

  it("does not leak a previous token's computed color when a token is invalid", () => {
    const host = document.createElement("div");
    host.style.setProperty("--formula-grid-bg", "#111111");
    host.style.setProperty("--formula-grid-line", "not-a-color");
    document.body.appendChild(host);

    try {
      const theme = resolveGridThemeFromCssVars(host);
      expect(theme.gridBg).toBe("rgb(17, 17, 17)");
      expect(theme.gridLine).toBe("not-a-color");
    } finally {
      host.remove();
    }
  });
});
