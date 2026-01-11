import { describe, expect, it } from "vitest";
import { DEFAULT_GRID_THEME, resolveGridTheme } from "../GridTheme";
import { readGridThemeFromCssVars, resolveCssVarValue } from "../resolveThemeFromCssVars";

describe("grid theme resolution", () => {
  it("returns defaults when no overrides are provided", () => {
    expect(resolveGridTheme()).toEqual(DEFAULT_GRID_THEME);
  });

  it("merges overrides on top of defaults (ignoring empty strings)", () => {
    const theme = resolveGridTheme({ gridBg: "#000000", gridLine: "" });
    expect(theme.gridBg).toBe("#000000");
    expect(theme.gridLine).toBe(DEFAULT_GRID_THEME.gridLine);
  });

  it("reads theme tokens from CSS variables and trims whitespace", () => {
    const style = {
      getPropertyValue: (name: string) => {
        if (name === "--formula-grid-bg") return "  #101010 ";
        if (name === "--formula-grid-selection-border") return "\n#ff00ff\t";
        return "";
      }
    };

    const partial = readGridThemeFromCssVars(style);
    expect(partial).toEqual({
      gridBg: "#101010",
      selectionBorder: "#ff00ff"
    });
  });

  it("resolves simple var() indirection between custom properties", () => {
    const style = {
      getPropertyValue: (name: string) => {
        if (name === "--formula-grid-bg") return "var(--app-bg)";
        if (name === "--app-bg") return "rgb(10, 20, 30)";
        return "";
      }
    };

    expect(resolveCssVarValue("var(--app-bg)", style)).toBe("rgb(10, 20, 30)");
    expect(resolveCssVarValue("var(--missing, #fff)", style)).toBe("#fff");
  });

  it("uses the nearest available fallback when resolving missing vars or cycles", () => {
    const styleMissing = {
      getPropertyValue: (name: string) => {
        if (name === "--a") return "var(--b)";
        if (name === "--b") return "";
        return "";
      }
    };

    expect(resolveCssVarValue("var(--a, #fff)", styleMissing)).toBe("#fff");

    const styleCycle = {
      getPropertyValue: (name: string) => {
        if (name === "--a") return "var(--b)";
        if (name === "--b") return "var(--a)";
        return "";
      }
    };

    expect(resolveCssVarValue("var(--a, #fff)", styleCycle)).toBe("#fff");
  });

  it("applies later sources last (prop overrides css)", () => {
    const style = {
      getPropertyValue: (name: string) => {
        if (name === "--formula-grid-bg") return "#111111";
        return "";
      }
    };

    const cssTheme = readGridThemeFromCssVars(style);
    const resolved = resolveGridTheme(cssTheme, { gridBg: "#222222" });
    expect(resolved.gridBg).toBe("#222222");
  });
});
