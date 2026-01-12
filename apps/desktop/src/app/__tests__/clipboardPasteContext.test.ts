import { describe, expect, it } from "vitest";

import { reconcileClipboardCopyContextForPaste } from "../clipboardPasteContext";

describe("clipboardPasteContext", () => {
  it("clears stale internal clipboard context when the clipboard no longer matches", () => {
    const initialContext = {
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      payload: { text: "A\tB", html: "<table><tr><td>A</td><td>B</td></tr></table>" },
      cells: [[{ value: "A", formula: null, styleId: 0 }]],
    };

    const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(initialContext, { text: "X\tY" });

    expect(isInternalPaste).toBe(false);
    expect(nextContext).toBeNull();
  });
});

