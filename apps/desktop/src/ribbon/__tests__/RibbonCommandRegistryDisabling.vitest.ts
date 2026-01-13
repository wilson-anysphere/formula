// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { Ribbon } from "../Ribbon";
import type { RibbonSchema } from "../ribbonSchema";
import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbonCommandRegistryDisabling";
import { setRibbonUiState } from "../ribbonUiState";

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
  vi.restoreAllMocks();
});

describe("CommandRegistry-backed ribbon disabling", () => {
  it("renders an unknown command id as disabled when the baseline override is applied", () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const schema: RibbonSchema = {
      tabs: [
        {
          id: "home",
          label: "Home",
          groups: [
            {
              id: "group",
              label: "Group",
              buttons: [{ id: "unknown.command", label: "Unknown", ariaLabel: "Unknown" }],
            },
          ],
        },
      ],
    };

    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema });

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: baselineDisabledById,
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    act(() => {
      root.render(<Ribbon actions={{}} schema={schema} />);
    });

    const button = container.querySelector<HTMLButtonElement>('[data-command-id="unknown.command"]');
    expect(button).toBeInstanceOf(HTMLButtonElement);
    expect(button?.disabled).toBe(true);

    act(() => root.unmount());
  });
});
