// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { setLocale } from "../../../i18n/index.js";
import { DocumentController } from "../../../document/documentController.js";

// Keep this test focused on sheet-id handling rather than UI rendering details.
vi.mock("../PivotBuilderPanel.js", () => ({
  PivotBuilderPanel: () => null,
}));

import { PivotBuilderPanelContainer } from "../PivotBuilderPanelContainer";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

async function flushMicrotasks(count = 8): Promise<void> {
  for (let i = 0; i < count; i += 1) {
    await Promise.resolve();
  }
}

describe("PivotBuilderPanelContainer sheet-id safety", () => {
  let host: HTMLDivElement | null = null;
  let root: ReturnType<typeof createRoot> | null = null;

  beforeEach(() => {
    setLocale("en-US");
    host = document.createElement("div");
    document.body.appendChild(host);
    root = createRoot(host);
  });

  afterEach(async () => {
    if (root) {
      await act(async () => {
        root?.unmount();
      });
    }
    host?.remove();
    host = null;
    root = null;
    vi.restoreAllMocks();
  });

  it("does not resurrect a deleted sheet when the initial selection references a stale sheet id", async () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", "A1", 1);
    doc.setCellValue("Sheet2", "A1", 2);
    doc.deleteSheet("Sheet2");

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();

    const staleSelection = {
      sheetId: "Sheet2",
      // Must be at least 2 rows so the container attempts to read header values.
      range: { startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
    };

    await act(async () => {
      root?.render(
        <PivotBuilderPanelContainer
          getDocumentController={() => doc}
          getActiveSheetId={() => "Sheet1"}
          getSelection={() => staleSelection}
        />,
      );
      await flushMicrotasks(10);
    });

    // The container should treat stale sheet ids as missing rather than probing via
    // `DocumentController.getCell()` (which would create a phantom sheet).
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(doc.getSheetMeta("Sheet2")).toBeNull();
  });
});

