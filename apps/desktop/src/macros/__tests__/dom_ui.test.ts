/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import type { MacroBackend, MacroCellUpdate } from "../types";
import { renderMacroRunner } from "../dom_ui";

describe("renderMacroRunner", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("invokes onApplyUpdates when the backend returns updates", async () => {
    vi.stubGlobal("confirm", vi.fn(() => true));

    const updates: MacroCellUpdate[] = [
      { sheetId: "Sheet1", row: 0, col: 0, value: 42, formula: null, displayValue: "42" },
    ];

    const backend: MacroBackend = {
      listMacros: vi.fn(async (_workbookId: string) => [{ id: "macro-1", name: "Macro1", language: "vba" as const }]),
      getMacroSecurityStatus: vi.fn(async (_workbookId: string) => ({
        hasMacros: true,
        trust: "trusted_once" as const,
        signature: { status: "unsigned" as const },
      })),
      setMacroTrust: vi.fn(async (_workbookId: string) => ({
        hasMacros: true,
        trust: "trusted_once" as const,
        signature: { status: "unsigned" as const },
      })),
      runMacro: vi.fn(async () => ({ ok: true, output: ["Hello"], updates })),
    };

    const onApplyUpdates = vi.fn(async () => {});

    const container = document.createElement("div");
    document.body.appendChild(container);
    await renderMacroRunner(container, backend, "workbook-1", { onApplyUpdates });

    const banner = container.querySelector('[data-testid="macro-security-banner"]');
    expect(banner?.textContent).toContain("Trust Center = trusted_once");

    const runButton = container.querySelectorAll("button")[1];
    expect(runButton).toBeTruthy();

    await (runButton as any).onclick?.();

    expect(onApplyUpdates).toHaveBeenCalledTimes(1);
    expect(onApplyUpdates).toHaveBeenCalledWith(updates);

    const output = container.querySelector("pre");
    expect(output?.textContent).toContain("Applied 1 updates.");

    container.remove();
  });

  it("renders Trust Center blocked state and blocked errors", async () => {
    vi.stubGlobal("confirm", vi.fn(() => true));
    vi.stubGlobal("prompt", vi.fn(() => "trusted_once"));

    const backend: MacroBackend = {
      listMacros: vi.fn(async (_workbookId: string) => [{ id: "macro-1", name: "Macro1", language: "vba" as const }]),
      getMacroSecurityStatus: vi.fn(async (_workbookId: string) => ({
        hasMacros: true,
        trust: "blocked" as const,
        signature: { status: "unsigned" as const },
      })),
      setMacroTrust: vi.fn(async (_workbookId: string, decision) => ({
        hasMacros: true,
        trust: decision as any,
        signature: { status: "unsigned" as const },
      })),
      runMacro: vi.fn(async () => ({
        ok: false,
        output: [],
        error: {
          message: "Macros are blocked by Trust Center policy.",
          code: "macro_blocked",
          blocked: {
            reason: "not_trusted" as const,
            status: { hasMacros: true, trust: "blocked" as const, signature: { status: "unsigned" as const } },
          },
        },
      })),
    };

    const container = document.createElement("div");
    document.body.appendChild(container);
    await renderMacroRunner(container, backend, "workbook-1");

    const banner = container.querySelector('[data-testid="macro-security-banner"]');
    expect(banner?.textContent).toContain("Macros blocked by Trust Center");

    const runButton = container.querySelectorAll("button")[1];
    await (runButton as any).onclick?.();

    const output = container.querySelector("pre");
    expect(output?.textContent).toContain("Blocked by Trust Center (not_trusted)");
    expect(output?.textContent).toContain("Macros are blocked by Trust Center policy.");

    container.remove();
  });
});
