// @vitest-environment jsdom

import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";

import { ApprovalModal } from "./ApprovalModal.js";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

describe("ApprovalModal", () => {
  it("renders summary + change preview", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const root = createRoot(host);

    await act(async () => {
      root.render(
        React.createElement(ApprovalModal, {
          request: {
            call: { name: "write_cell", arguments: { cell: "Sheet1!A1", value: 123 } },
            preview: {
              timing_ms: 1,
              tool_results: [],
              changes: [{ cell: "Sheet1!A1", type: "modify", before: { value: null }, after: { value: 123 } }],
              summary: { total_changes: 1, creates: 0, modifies: 1, deletes: 0 },
              warnings: [],
              requires_approval: true,
              approval_reasons: ["unit-test"],
            },
          },
          onApprove: () => {},
          onReject: () => {},
        }),
      );
    });

    expect(host.textContent).toContain("Approve AI changes?");
    expect(host.textContent).toContain("write_cell");
    expect(host.textContent).toContain("Sheet1!A1");

    act(() => {
      root.unmount();
    });
    host.remove();
  });
});
