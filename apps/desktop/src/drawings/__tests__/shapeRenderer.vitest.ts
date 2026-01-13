import { describe, expect, it } from "vitest";

import { parseShapeRenderSpec } from "../shapeRenderer";

describe("parseShapeRenderSpec", () => {
  it("parses preset geometry from a real-world <xdr:spPr> snippet (image fixture)", () => {
    // Snippet copied from `fixtures/xlsx/basic/image.xlsx` -> `xl/drawings/drawing1.xml`.
    const raw = `
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    `;

    expect(parseShapeRenderSpec(raw)).toEqual({
      geometry: { type: "rect" },
      fill: { type: "none" },
      stroke: { color: "#000000", widthEmu: 9525 },
      label: undefined,
    });
  });

  it("parses solid fill + stroke and extracts the first txBody line as label", () => {
    const raw = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Hello world</a:t></a:r>
          </a:p>
        </xdr:txBody>
        <xdr:spPr>
          <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="ff0000"/></a:solidFill>
          <a:ln w="12700">
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
          </a:ln>
        </xdr:spPr>
      </xdr:sp>
    `;

    expect(parseShapeRenderSpec(raw)).toEqual({
      geometry: { type: "ellipse" },
      fill: { type: "solid", color: "#FF0000" },
      stroke: { color: "#00FF00", widthEmu: 12700 },
      label: "Hello world",
    });
  });

  it("returns null for unsupported preset geometries", () => {
    const raw = `
      <xdr:spPr>
        <a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    `;
    expect(parseShapeRenderSpec(raw)).toBeNull();
  });

  it("parses roundRect adjust values from <a:avLst>", () => {
    const raw = `
      <xdr:spPr>
        <a:prstGeom prst="roundRect">
          <a:avLst>
            <a:gd name="adj" fmla="val 50000"/>
          </a:avLst>
        </a:prstGeom>
      </xdr:spPr>
    `;

    expect(parseShapeRenderSpec(raw)).toEqual({
      geometry: { type: "roundRect", adj: 50000 },
      fill: { type: "none" },
      stroke: { color: "#000000", widthEmu: 9525 },
      label: undefined,
    });
  });

  it("emits rgba() colors when srgbClr includes an alpha child", () => {
    const raw = `
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill>
          <a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>
        </a:solidFill>
      </xdr:spPr>
    `;

    expect(parseShapeRenderSpec(raw)?.fill).toEqual({ type: "solid", color: "rgba(255,0,0,0.5)" });
  });
});
