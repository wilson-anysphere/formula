/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { CollaboratorsListUiController } from "./collaborators-list-ui-controller.js";

describe("CollaboratorsListUiController", () => {
  it("renders N collaborators", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const ui = new CollaboratorsListUiController({ container });
    ui.setCollaborators([
      { key: "u1:1", name: "Ada", color: "var(--accent)" },
      { key: "u2:2", name: "Grace", color: "var(--success)" },
      { key: "u3:3", name: "Linus", color: "var(--warning)" },
    ]);

    const items = container.querySelectorAll('[data-testid="presence-collaborator"]');
    expect(items.length).toBe(3);

    ui.destroy();
    container.remove();
  });

  it("caps visible collaborators and shows an overflow pill", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const ui = new CollaboratorsListUiController({ container, maxVisible: 2 });
    ui.setCollaborators([
      { key: "u1:1", name: "Ada", color: "var(--accent)" },
      { key: "u2:2", name: "Grace", color: "var(--success)" },
      { key: "u3:3", name: "Linus", color: "var(--warning)" },
    ]);

    expect(container.querySelector('[data-testid="presence-collaborators"]')?.getAttribute("aria-label")).toBe(
      "Collaborators",
    );

    const items = container.querySelectorAll('[data-testid="presence-collaborator"]');
    expect(items.length).toBe(2);

    const overflow = container.querySelector('[data-testid="presence-collaborator-overflow"]') as HTMLElement | null;
    expect(overflow?.textContent).toBe("+1");

    ui.destroy();
    container.remove();
  });
});
