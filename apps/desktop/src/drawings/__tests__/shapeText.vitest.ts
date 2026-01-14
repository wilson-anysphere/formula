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

  it("preserves <a:tab/> placeholders as tab characters", () => {
    const rawXml = `
      <xdr:sp>
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

  it("prepends <a:buChar> bullet characters to paragraph text", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:buChar char="•"/></a:pPr>
            <a:r><a:t>Item 1</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr><a:buChar char="•"/></a:pPr>
            <a:r><a:t>Item 2</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("• Item 1\n• Item 2");
  });

  it("inherits bullet characters from <a:lstStyle> based on paragraph lvl", () => {
    const rawXml = `
      <xdr:sp>
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

  it("inherits default run properties from <a:lstStyle> lvlNpPr", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle>
            <a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr>
            <a:lvl2pPr><a:defRPr sz="1000"/></a:lvl2pPr>
          </a:lstStyle>
          <a:p><a:pPr lvl="0"/><a:r><a:t>Top</a:t></a:r></a:p>
          <a:p><a:pPr lvl="1"/><a:r><a:t>Nested</a:t></a:r></a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.find((r) => r.text === "Top")?.fontSizePt).toBe(20);
    expect(parsed?.textRuns.find((r) => r.text === "Nested")?.fontSizePt).toBe(10);
  });

  it("indents lvl paragraphs even when bullets are disabled via <a:buNone/>", () => {
    const rawXml = `
      <xdr:sp>
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

  it("prepends <a:buAutoNum> numbering to paragraph text", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr>
            <a:r><a:t>First</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr>
            <a:r><a:t>Second</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("1. First\n2. Second");
  });

  it("numbers nested <a:buAutoNum> paragraphs independently by lvl", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle>
            <a:lvl1pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:lvl1pPr>
            <a:lvl2pPr><a:buAutoNum type="alphaLcPeriod" startAt="1"/></a:lvl2pPr>
          </a:lstStyle>
          <a:p><a:pPr lvl="0"/><a:r><a:t>Item 1</a:t></a:r></a:p>
          <a:p><a:pPr lvl="1"/><a:r><a:t>Sub 1</a:t></a:r></a:p>
          <a:p><a:pPr lvl="1"/><a:r><a:t>Sub 2</a:t></a:r></a:p>
          <a:p><a:pPr lvl="0"/><a:r><a:t>Item 2</a:t></a:r></a:p>
          <a:p><a:pPr lvl="1"/><a:r><a:t>Sub again</a:t></a:r></a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe(
      ["1. Item 1", "  a. Sub 1", "  b. Sub 2", "2. Item 2", "  a. Sub again"].join("\n"),
    );
  });

  it("supports alpha/roman buAutoNum formats (parenBoth)", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:buAutoNum type="alphaLcParenBoth" startAt="1"/></a:pPr>
            <a:r><a:t>First</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr><a:buAutoNum type="alphaLcParenBoth" startAt="1"/></a:pPr>
            <a:r><a:t>Second</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr><a:buAutoNum type="romanUcParenBoth" startAt="3"/></a:pPr>
            <a:r><a:t>Third</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("(a) First\n(b) Second\n(III) Third");
  });

  it("decodes numeric XML entities (including code points > 0xFFFF)", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Hello &#x1F600;</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("Hello 😀");
  });

  it("parses bodyPr insets (lIns/tIns/rIns/bIns) as EMU values", () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr lIns="95250" tIns="0" rIns="19050" bIns="38100"/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.insetLeftEmu).toBe(95250);
    expect(parsed?.insetTopEmu).toBe(0);
    expect(parsed?.insetRightEmu).toBe(19050);
    expect(parsed?.insetBottomEmu).toBe(38100);
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

  it('treats "yes"/"no" boolean values as true/false for run style attrs', () => {
    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="yes" i="no"/>
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
      italic: false,
    });
  });

  it("handles namespace declarations and non-self-closing line breaks", () => {
    // When raw snippets include `xmlns:` attributes, our DOMParser wrapper must avoid
    // generating invalid `xmlns:xmlns="..."` declarations. This case also uses an explicit
    // `<a:br></a:br>` pair (valid XML), which the regex fallback does not treat as a break.
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
            <a:br></a:br>
            <a:r><a:t>World</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const parsed = parseDrawingMLShapeText(rawXml);
    expect(parsed).not.toBeNull();
    expect(parsed?.textRuns.map((r) => r.text).join("")).toBe("Hello\nWorld");
  });
});
