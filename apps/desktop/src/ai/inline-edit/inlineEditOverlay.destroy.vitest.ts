/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import type { ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";
import { InlineEditOverlay } from "./inlineEditOverlay";

describe("InlineEditOverlay destroy", () => {
  it("resolves pending approval and removes the element from the DOM", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new InlineEditOverlay(container);
    overlay.open("Sheet1!A1:A1", { onCancel: vi.fn(), onRun: vi.fn() });

    const preview: ToolPlanPreview = {
      timing_ms: 0,
      tool_results: [],
      changes: [],
      summary: { total_changes: 0, creates: 0, modifies: 0, deletes: 0 },
      warnings: [],
      requires_approval: true,
      approval_reasons: ["unit-test"],
    };

    const approval = overlay.requestApproval(preview);

    overlay.destroy();

    await expect(approval).resolves.toBe(false);
    expect(overlay.isOpen()).toBe(false);
    expect(overlay.element.isConnected).toBe(false);

    // No-op after destroy (should not reattach or show UI).
    overlay.open("Sheet1!A1:A1", { onCancel: vi.fn(), onRun: vi.fn() });
    expect(overlay.isOpen()).toBe(false);
    expect(overlay.element.isConnected).toBe(false);

    container.remove();
  });
});

