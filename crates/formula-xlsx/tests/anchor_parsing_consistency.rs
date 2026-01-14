use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use formula_model::charts::ChartAnchor;
use formula_model::drawings::{Anchor, AnchorPoint, CellOffset, DrawingObjectKind, EmuSize};
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

fn assert_anchor_parses_consistently(drawing_xml: &str, expected: Anchor) -> DrawingPart {
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

    drawing_part
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

#[test]
fn parses_anchor_wrapped_in_mc_alternate_content_choice_branch() {
    // Some producers wrap anchors in `mc:AlternateContent`, with both Choice + Fallback branches.
    // We should pick the first Choice branch that contains an anchor and avoid duplicating refs.
    let choice_anchor = format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
  {}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#,
        chart_frame_xml()
    );

    let fallback_anchor = choice_anchor.replace("rId1", "rId2");

    let anchor_xml = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="c14">{choice_anchor}</mc:Choice>
  <mc:Fallback>{fallback_anchor}</mc:Fallback>
</mc:AlternateContent>"#
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(4, 3), CellOffset::new(0, 0)),
    };
    assert_anchor_parses_consistently(&drawing_xml, expected);
}

#[test]
fn parses_anchor_wrapped_in_mc_alternate_content_fallback_branch() {
    // When Choice branches do not contain anchors, fall back to the first Fallback branch that does.
    let fallback_anchor = format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
  {}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#,
        chart_frame_xml()
    );

    let anchor_xml = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="c14"><xdr:sp/></mc:Choice>
  <mc:Fallback>{fallback_anchor}</mc:Fallback>
</mc:AlternateContent>"#
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(4, 3), CellOffset::new(0, 0)),
    };
    assert_anchor_parses_consistently(&drawing_xml, expected);
}

#[test]
fn parses_chart_frame_wrapped_in_mc_alternate_content_choice_branch() {
    let choice_frame = chart_frame_xml().to_string();
    let fallback_frame = choice_frame.replace("rId1", "rId2");

    let object_xml = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="c14">{choice_frame}</mc:Choice>
  <mc:Fallback>{fallback_frame}</mc:Fallback>
</mc:AlternateContent>"#
    );

    let anchor_xml = format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
  {object_xml}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(4, 3), CellOffset::new(0, 0)),
    };

    let drawing_part = assert_anchor_parses_consistently(&drawing_xml, expected);
    assert!(
        matches!(
            &drawing_part.objects[0].kind,
            DrawingObjectKind::ChartPlaceholder { rel_id, raw_xml } if rel_id == "rId1" && raw_xml.contains("rId1") && !raw_xml.contains("rId2")
        ),
        "expected DrawingPart to materialize a single chart placeholder from the Choice branch"
    );
}

#[test]
fn parses_chart_frame_wrapped_in_mc_alternate_content_fallback_branch() {
    let fallback_frame = chart_frame_xml().to_string();

    // Choice branch contains no chart; we should fall back to the first Fallback branch that does.
    let object_xml = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="c14"><xdr:graphicFrame/></mc:Choice>
  <mc:Fallback>{fallback_frame}</mc:Fallback>
</mc:AlternateContent>"#
    );

    let anchor_xml = format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
  {object_xml}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(4, 3), CellOffset::new(0, 0)),
    };

    let drawing_part = assert_anchor_parses_consistently(&drawing_xml, expected);
    assert!(
        matches!(
            &drawing_part.objects[0].kind,
            DrawingObjectKind::ChartPlaceholder { rel_id, raw_xml } if rel_id == "rId1" && raw_xml.contains("rId1")
        ),
        "expected DrawingPart to materialize a chart placeholder from the Fallback branch"
    );
}

#[test]
fn drawing_part_prefers_xdr_cnvpr_over_non_canonical_a_cnvpr() {
    // Some producers place a non-canonical `<a:cNvPr>` inside `a:graphicData` ahead of the
    // canonical `xdr:nvGraphicFramePr/xdr:cNvPr` block. DrawingPart parsing should prefer the
    // `xdr:*` node so object IDs remain stable and match chart extraction.
    let frame_xml = r#"<xdr:graphicFrame>
  <!-- non-canonical `a:cNvPr` before `nvGraphicFramePr` -->
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <a:cNvPr id="999" name="Wrong"/>
      <c:chart r:id="rId1"/>
    </a:graphicData>
  </a:graphic>
  <xdr:nvGraphicFramePr>
    <xdr:cNvPr id="7" name="Chart 7"/>
    <xdr:cNvGraphicFramePr/>
  </xdr:nvGraphicFramePr>
  <xdr:xfrm/>
</xdr:graphicFrame>"#;

    let anchor_xml = format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
  {frame_xml}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#
    );
    let drawing_xml = wrap_wsdr(&anchor_xml);

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(4, 3), CellOffset::new(0, 0)),
    };

    let drawing_part = assert_anchor_parses_consistently(&drawing_xml, expected);
    assert_eq!(
        drawing_part.objects[0].id.0, 7,
        "DrawingPart should prefer the canonical xdr:cNvPr id"
    );
    assert!(
        matches!(
            &drawing_part.objects[0].kind,
            DrawingObjectKind::ChartPlaceholder { rel_id, .. } if rel_id == "rId1"
        ),
        "expected a chart placeholder object"
    );
}
