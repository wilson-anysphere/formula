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

  it("preserves <a:tab/> placeholders as tab characters", () => {
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
            <a:tab/>
            <a:r><a:t>World</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("Hello\tWorld");
  });

  it("inherits bullet characters from <a:lstStyle> based on paragraph lvl", () => {
    const rawXml = `
      <xdr:sp
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      >
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle>
            <a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr>
            <a:lvl2pPr><a:buChar char="◦"/></a:lvl2pPr>
          </a:lstStyle>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r><a:t>Top</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="1"/>
            <a:r><a:t>Nested</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("• Top\n  ◦ Nested");
  });

  it("indents lvl paragraphs even when bullets are disabled via <a:buNone/>", () => {
    const rawXml = `
      <xdr:sp
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      >
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle>
            <a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr>
            <a:lvl2pPr><a:buChar char="◦"/></a:lvl2pPr>
          </a:lstStyle>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r><a:t>Top</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="1"><a:buNone/></a:pPr>
            <a:r><a:t>No bullet but indented</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("• Top\n  No bullet but indented");
  });

  it("applies default run styles with per-run overrides", () => {
    const rawXml = `
      <xdr:sp
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      >
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="1400">
                <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
              </a:defRPr>
            </a:pPr>
            <a:r>
              <a:rPr b="1"/>
              <a:t>Styled</a:t>
            </a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns).toHaveLength(1);
    expect(parsed?.textRuns[0]).toMatchObject({
      text: "Styled",
      bold: true,
      fontSizePt: 14,
      color: "#00FF00",
    });
  });
});
