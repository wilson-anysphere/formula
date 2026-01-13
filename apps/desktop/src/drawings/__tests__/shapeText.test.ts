import { describe, expect, it } from "vitest";

import { parseDrawingMLShapeText } from "../drawingml/shapeText";

describe("parseDrawingMLShapeText", () => {
  it("extracts text from a single paragraph", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Hello</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("Hello");
  });

  it("preserves line breaks between paragraphs", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Line 1</a:t></a:r></a:p>
          <a:p><a:r><a:t>Line 2</a:t></a:r></a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("Line 1\nLine 2");
  });

  it("applies default run styles with per-run overrides", () => {
    const rawXml = `
      <xdr:sp>
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

