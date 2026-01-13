use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use formula_model::charts::ChartAnchor;
use formula_model::drawings::{Anchor, AnchorPoint, CellOffset, EmuSize};
use formula_model::CellRef;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::XlsxPackage;

fn chart_anchor_to_drawing_anchor(anchor: &ChartAnchor) -> Anchor {
    match anchor {
        ChartAnchor::TwoCell {
            from_col,
            from_row,
            from_col_off_emu,
            from_row_off_emu,
            to_col,
            to_row,
            to_col_off_emu,
            to_row_off_emu,
        } => Anchor::TwoCell {
            from: AnchorPoint::new(
                CellRef::new(*from_row, *from_col),
                CellOffset::new(*from_col_off_emu, *from_row_off_emu),
            ),
            to: AnchorPoint::new(
                CellRef::new(*to_row, *to_col),
                CellOffset::new(*to_col_off_emu, *to_row_off_emu),
            ),
        },
        ChartAnchor::OneCell {
            from_col,
            from_row,
            from_col_off_emu,
            from_row_off_emu,
            cx_emu,
            cy_emu,
        } => Anchor::OneCell {
            from: AnchorPoint::new(
                CellRef::new(*from_row, *from_col),
                CellOffset::new(*from_col_off_emu, *from_row_off_emu),
            ),
            ext: EmuSize::new(*cx_emu, *cy_emu),
        },
        ChartAnchor::Absolute {
            x_emu,
            y_emu,
            cx_emu,
            cy_emu,
        } => Anchor::Absolute {
            pos: CellOffset::new(*x_emu, *y_emu),
            ext: EmuSize::new(*cx_emu, *cy_emu),
        },
    }
}

fn wrap_wsdr(anchor_xml: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
{anchor_xml}
</xdr:wsDr>"#
    )
}

fn chart_frame_xml() -> &'static str {
    r#"<xdr:graphicFrame>
  <xdr:nvGraphicFramePr>
    <xdr:cNvPr id="1" name="Chart 1"/>
  </xdr:nvGraphicFramePr>
  <a:graphic>
    <a:graphicData>
      <c:chart r:id="rId1"/>
    </a:graphicData>
  </a:graphic>
</xdr:graphicFrame>"#
}

fn zip_with_drawing_part(drawing_xml: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start file");
    zip.write_all(drawing_xml.as_bytes())
        .expect("write drawing xml");

    zip.finish().expect("finish zip").into_inner()
}

fn assert_anchor_parses_consistently(drawing_xml: &str, expected: Anchor) {
    // 1) Legacy chart ref extraction (`ChartAnchor`).
    let chart_refs =
        formula_xlsx::drawingml::extract_chart_refs(drawing_xml.as_bytes(), "drawing1.xml")
            .expect("extract_chart_refs");
    assert_eq!(chart_refs.len(), 1, "expected one chart ref");
    assert_eq!(chart_refs[0].rel_id, "rId1");
    assert_eq!(
        chart_anchor_to_drawing_anchor(&chart_refs[0].anchor),
        expected,
        "legacy ChartAnchor parse mismatch"
    );

    // 2) ChartObject extraction (`drawings::Anchor`) via the public package API.
    let package_bytes = zip_with_drawing_part(drawing_xml);
    let package = XlsxPackage::from_bytes(&package_bytes).expect("read package");
    let chart_objects = package
        .extract_chart_objects()
        .expect("extract_chart_objects");
    assert_eq!(chart_objects.len(), 1, "expected one chart object");
    assert_eq!(
        chart_objects[0].anchor, expected,
        "chart object anchor parse mismatch"
    );

    // 3) DrawingPart parsing (`drawings::Anchor`).
    let mut parts = BTreeMap::<String, Vec<u8>>::new();
    parts.insert(
        "xl/drawings/drawing1.xml".to_string(),
        drawing_xml.as_bytes().to_vec(),
    );
    parts.insert(
        "xl/drawings/_rels/drawing1.xml.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_vec(),
    );
    let mut workbook = formula_model::Workbook::new();
    let drawing_part =
        DrawingPart::parse_from_parts(0, "xl/drawings/drawing1.xml", &parts, &mut workbook)
            .expect("parse drawing part");
    assert_eq!(drawing_part.objects.len(), 1, "expected one drawing object");
    assert_eq!(
        drawing_part.objects[0].anchor, expected,
        "drawing part anchor parse mismatch"
    );
}

#[test]
fn parses_two_cell_anchor_with_whitespace_and_missing_offsets() {
    let anchor_xml = format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from>
    <xdr:col>
      2
    </xdr:col>
    <xdr:row> 3 </xdr:row>
  </xdr:from>
  <xdr:to>
    <xdr:col> 4 </xdr:col>
    <xdr:row>
      5
    </xdr:row>
  </xdr:to>
  {}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#,
        chart_frame_xml()
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(3, 2), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(5, 4), CellOffset::new(0, 0)),
    };
    assert_anchor_parses_consistently(&drawing_xml, expected);
}

#[test]
fn parses_one_cell_anchor_with_whitespace_and_missing_offsets() {
    let anchor_xml = format!(
        r#"<xdr:oneCellAnchor>
  <xdr:from>
    <xdr:col> 6 </xdr:col>
    <xdr:row>
      7
    </xdr:row>
  </xdr:from>
  <xdr:ext cx=" 111 " cy="222"/>
  {}
  <xdr:clientData/>
</xdr:oneCellAnchor>"#,
        chart_frame_xml()
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(7, 6), CellOffset::new(0, 0)),
        ext: EmuSize::new(111, 222),
    };
    assert_anchor_parses_consistently(&drawing_xml, expected);
}

#[test]
fn parses_absolute_anchor_and_trims_attribute_values() {
    let anchor_xml = format!(
        r#"<xdr:absoluteAnchor>
  <xdr:pos x=" 100 " y=" 200 "/>
  <xdr:ext cx=" 300 " cy="400"/>
  {}
  <xdr:clientData/>
</xdr:absoluteAnchor>"#,
        chart_frame_xml()
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::Absolute {
        pos: CellOffset::new(100, 200),
        ext: EmuSize::new(300, 400),
    };
    assert_anchor_parses_consistently(&drawing_xml, expected);
}
