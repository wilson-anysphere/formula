// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest";

import { parseDrawingMLShapeText } from "../drawingml/shapeText";

describe("parseDrawingMLShapeText (DOMParser path)", () => {
  it("does not generate invalid xmlns:xmlns wrapper attributes when rawXml includes xmlns declarations", () => {
    const rawXml = `
      <xdr:sp
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      >
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Hello</a:t></a:r>
            <a:br/>
            <a:r><a:t>World</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const OriginalDOMParser = DOMParser;
    class InspectingDOMParser extends OriginalDOMParser {
      parseFromString(str: string, type: any): Document {
        // If we accidentally treat `xmlns:` declarations as namespaced attributes, we can end up
        // emitting invalid XML like `xmlns:xmlns="..."` in the synthetic wrapper root.
        // Guard against regressions by asserting we never hand such markup to DOMParser.
        if (str.includes("xmlns:xmlns=")) {
          throw new Error(`Invalid wrapper XML contained xmlns:xmlns=:\n${str}`);
        }
        return super.parseFromString(str, type);
      }
    }
    vi.stubGlobal("DOMParser", InspectingDOMParser as any);
    try {
      const parsed = parseDrawingMLShapeText(rawXml);
      expect(parsed).not.toBeNull();
      expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("Hello\nWorld");
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

