// @vitest-environment jsdom

import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { createPanelBodyRenderer } from "../panelBodyRenderer.js";
import { PanelIds } from "../panelRegistry.js";
import { VbaMigratePanel } from "./VbaMigratePanel.js";
import { VbaMigrator } from "../../../../../packages/vba-migrate/src/converter.js";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

function installTauriInvoke(invoke: (cmd: string, args?: any) => Promise<any>) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).__TAURI__ = { core: { invoke } };
}

function clearTauriInvoke() {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  delete (globalThis as any).__TAURI__;
}

describe("VbaMigratePanel", () => {
  afterEach(() => {
    document.body.innerHTML = "";
    clearTauriInvoke();
  });

  it("loads the VBA project via get_vba_project and renders module buttons", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd !== "get_vba_project") throw new Error(`Unexpected command: ${cmd}`);
      return {
        name: "TestProject",
        constants: "Const Foo = 1",
        references: [],
        modules: [
          {
            name: "Module1",
            module_type: "Standard",
            code: 'Sub Main()\n  Range("A1").Value = 1\nEnd Sub\n',
          },
        ],
      };
    });

    installTauriInvoke(invoke);

    const renderer = createPanelBodyRenderer({ getDocumentController: () => ({}), workbookId: "workbook-1" });

    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.VBA_MIGRATE, body);
    });

    await act(async () => {
      await flushPromises();
    });

    expect(invoke).toHaveBeenCalledWith("get_vba_project", { workbook_id: "workbook-1" });
    expect(body.querySelector('[data-testid="vba-project-name"]')?.textContent).toContain("TestProject");
    expect(body.querySelector('[data-testid="vba-module-Module1"]')).toBeTruthy();

    act(() => {
      renderer.cleanup([]);
    });
    expect(body.childElementCount).toBe(0);
  });

  it("shows an analyzer report with risk + object model usage", async () => {
    installTauriInvoke(async () => {
      return {
        name: "RiskyProject",
        constants: null,
        references: [],
        modules: [
          {
            name: "ModuleRisk",
            module_type: "Standard",
            code: [
              "Sub Main()",
              '  Range("A1").Value = 1',
              '  Shell("calc.exe")',
              '  Execute "MsgBox 1"',
              "End Sub",
              "",
            ].join("\n"),
          },
        ],
      };
    });

    const renderer = createPanelBodyRenderer({ getDocumentController: () => ({}), workbookId: "workbook-2" });
    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.VBA_MIGRATE, body);
      await flushPromises();
    });

    const risk = body.querySelector('[data-testid="vba-analysis-risk"]');
    expect(risk?.textContent).toContain("55");
    expect(risk?.textContent).toContain("(medium)");

    const rangeUsage = body.querySelector('[data-testid="vba-analysis-usage-Range"]');
    expect(rangeUsage?.textContent).toContain("Range:");
    expect(rangeUsage?.textContent).toContain("1");
  });

  it(
    "can convert a selected module via a mocked migrator/LLM",
    async () => {
    installTauriInvoke(async () => {
      return {
        name: "ConvertProject",
        constants: null,
        references: [],
        modules: [
          {
            name: "Module1",
            module_type: "Standard",
            code: ['Sub Main()', '  Range("A1").Value = 1', "End Sub"].join("\n"),
          },
        ],
      };
    });

    const llmComplete = vi.fn(async () => {
      return ['const sheet = ctx.activeSheet;', 'sheet.range("A1").value = 1;'].join("\n");
    });

    const createMigrator = () => new VbaMigrator({ llm: { complete: llmComplete } as any });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(VbaMigratePanel, { workbookId: "workbook-3", createMigrator }));
    });

    await act(async () => {
      await flushPromises();
    });

    const convertBtn = host.querySelector('[data-testid="vba-convert-typescript"]') as HTMLButtonElement | null;
    expect(convertBtn).toBeInstanceOf(HTMLButtonElement);
    expect(convertBtn?.disabled).toBe(false);

    await act(async () => {
      convertBtn?.click();
    });

    const started = Date.now();
    while (Date.now() - started < 5_000) {
      const output = host.querySelector('[data-testid="vba-converted-code"]') as HTMLTextAreaElement | null;
      const error = host.querySelector('[data-testid="vba-conversion-error"]');
      if (output?.value || error) break;
      await act(async () => {
        await flushPromises();
      });
    }

    expect(llmComplete).toHaveBeenCalled();
    expect(host.querySelector('[data-testid="vba-conversion-error"]')).toBeFalsy();
    const output = host.querySelector('[data-testid="vba-converted-code"]') as HTMLTextAreaElement | null;
    expect(output?.value).toContain("export default async function main");
    expect(output?.value).toContain('sheet.range("A1").value = 1');
    },
    10_000,
  );
});
