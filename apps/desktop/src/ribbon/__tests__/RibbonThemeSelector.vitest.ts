// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import { setRibbonUiState } from "../ribbonUiState";
import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { registerBuiltinCommands } from "../../commands/registerBuiltinCommands.js";

afterEach(() => {
  act(() => {
    setRibbonUiState({
      pressedById: Object.create(null),
      labelById: Object.create(null),
      disabledById: Object.create(null),
      shortcutById: Object.create(null),
      ariaKeyShortcutsById: Object.create(null),
    });
  });
  document.body.innerHTML = "";
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function renderRibbon(actions: any, initialTabId: string) {
  // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
  // https://react.dev/reference/react/act#configuring-your-test-environment
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);

  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions, initialTabId }));
  });

  return { container, root };
}

describe("Ribbon theme selector", () => {
  it("renders in the View tab and emits theme command ids", async () => {
    const onCommand = vi.fn();
    const { container, root } = renderRibbon({ onCommand }, "view");

    // Let effects run (e.g. density + event listeners).
    await act(async () => {
      await Promise.resolve();
    });

    const trigger = container.querySelector<HTMLButtonElement>('[data-testid="theme-selector"]');
    expect(trigger).toBeTruthy();

    act(() => {
      trigger!.click();
    });

    const dark = container.querySelector<HTMLButtonElement>('[data-testid="theme-option-dark"]');
    expect(dark).toBeTruthy();

    act(() => {
      dark!.click();
    });

    expect(onCommand).toHaveBeenCalledWith("view.appearance.theme.dark");

    act(() => root.unmount());
  });

  it("executes theme commands via CommandRegistry and updates the theme selector label override", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    const commandRegistry = new CommandRegistry();
    const layoutController = {} as any;
    const app = { focus: vi.fn() } as any;

    // Default theme preference is Light (not System) for new users.
    let preference: "system" | "light" | "dark" | "high-contrast" = "light";
    const themeController = {
      setThemePreference: vi.fn((next: string) => {
        preference = next as typeof preference;
      }),
    } as any;

    const refreshRibbonUiState = vi.fn(() => {
      const label = (() => {
        switch (preference) {
          case "system":
            return "System";
          case "light":
            return "Light";
          case "dark":
            return "Dark";
          case "high-contrast":
            return "High Contrast";
          default:
            return "Light";
        }
      })();

      act(() => {
        setRibbonUiState({
          pressedById: Object.create(null),
          labelById: { "view.appearance.theme": `Theme: ${label}` },
          disabledById: Object.create(null),
          shortcutById: Object.create(null),
          ariaKeyShortcutsById: Object.create(null),
        });
      });
    });

    registerBuiltinCommands({
      commandRegistry,
      app,
      layoutController,
      themeController,
      refreshRibbonUiState,
    });

    // Seed the initial label override the same way main.ts does.
    refreshRibbonUiState();

    let lastCommandPromise: Promise<unknown> | null = null;
    const onCommand = (id: string) => {
      lastCommandPromise = commandRegistry.executeCommand(id);
    };

    const { container, root } = renderRibbon({ onCommand }, "view");

    // Let effects run (e.g. density + event listeners).
    await act(async () => {
      await Promise.resolve();
    });

    const trigger = container.querySelector<HTMLButtonElement>('[data-testid="theme-selector"]');
    expect(trigger).toBeTruthy();
    const themeLabel = () =>
      container.querySelector<HTMLButtonElement>('[data-testid="theme-selector"]')?.querySelector(".ribbon-button__label")?.textContent?.trim() ??
      "";
    expect(themeLabel()).toBe("Theme: Light");

    act(() => {
      trigger!.click();
    });

    const dark = container.querySelector<HTMLButtonElement>('[data-testid="theme-option-dark"]');
    expect(dark).toBeTruthy();

    act(() => {
      dark!.click();
    });

    await act(async () => {
      await (lastCommandPromise ?? Promise.resolve());
    });

    expect(themeController.setThemePreference).toHaveBeenCalledWith("dark");
    expect(refreshRibbonUiState).toHaveBeenCalled();
    expect(themeLabel()).toBe("Theme: Dark");

    act(() => root.unmount());
  });
});
