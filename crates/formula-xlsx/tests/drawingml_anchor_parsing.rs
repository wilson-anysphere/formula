use std::collections::BTreeMap;
use std::io::Write;

use formula_model::charts::ChartAnchor;
use formula_model::drawings::{Anchor, AnchorPoint, CellOffset, EmuSize};
use formula_model::{CellRef, Workbook};
use formula_xlsx::drawingml::extract_chart_refs;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::XlsxPackage;

const DRAWING_PATH: &str = "xl/drawings/drawing1.xml";
const DRAWING_RELS_PATH: &str = "xl/drawings/_rels/drawing1.xml.rels";

const EMPTY_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

fn wrap_wsdr(anchor_xml: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
{anchor_xml}
</xdr:wsDr>"#
    )
}

fn zip_bytes_with_part(path: &str, bytes: &[u8]) -> Vec<u8> {
    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file(path, options).expect("start_file");
    zip.write_all(bytes).expect("write_all");
    zip.finish().expect("finish").into_inner()
}

fn chart_anchor_to_drawing_anchor(anchor: ChartAnchor) -> Anchor {
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
                CellRef::new(from_row, from_col),
                CellOffset::new(from_col_off_emu, from_row_off_emu),
            ),
            to: AnchorPoint::new(
                CellRef::new(to_row, to_col),
                CellOffset::new(to_col_off_emu, to_row_off_emu),
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
                CellRef::new(from_row, from_col),
                CellOffset::new(from_col_off_emu, from_row_off_emu),
            ),
            ext: EmuSize::new(cx_emu, cy_emu),
        },
        ChartAnchor::Absolute {
            x_emu,
            y_emu,
            cx_emu,
            cy_emu,
        } => Anchor::Absolute {
            pos: CellOffset::new(x_emu, y_emu),
            ext: EmuSize::new(cx_emu, cy_emu),
        },
    }
}

fn assert_anchor_parses_consistently(drawing_xml: &str, expected: Anchor) {
    // New chart object extraction path (uses `extract_chart_object_refs` internally).
    let zip_bytes = zip_bytes_with_part(DRAWING_PATH, drawing_xml.as_bytes());
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("package from bytes");
    let chart_objects = pkg.extract_chart_objects().expect("extract chart objects");
    assert_eq!(chart_objects.len(), 1, "expected one chart object");
    assert_eq!(chart_objects[0].anchor, expected);

    // Drawing part parsing.
    let mut parts = BTreeMap::<String, Vec<u8>>::new();
    parts.insert(DRAWING_PATH.to_string(), drawing_xml.as_bytes().to_vec());
    parts.insert(
        DRAWING_RELS_PATH.to_string(),
        EMPTY_RELS_XML.as_bytes().to_vec(),
    );
    let mut workbook = Workbook::default();
    let drawing_part = DrawingPart::parse_from_parts(0, DRAWING_PATH, &parts, &mut workbook)
        .expect("parse drawing");
    assert_eq!(drawing_part.objects.len(), 1, "expected one drawing object");
    assert_eq!(drawing_part.objects[0].anchor, expected);

    // Legacy chart extraction path (`drawingml::extract_chart_refs`).
    let chart_refs = extract_chart_refs(drawing_xml.as_bytes(), DRAWING_PATH).expect("chart refs");
    assert_eq!(chart_refs.len(), 1, "expected one chart ref");
    let legacy_anchor = chart_anchor_to_drawing_anchor(chart_refs[0].anchor.clone());
    assert_eq!(legacy_anchor, expected);
}

#[test]
fn one_cell_anchor_defaults_missing_offsets_to_zero() {
    let drawing_xml = wrap_wsdr(
        r#"<xdr:oneCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:ext cx="111" cy="222"/>
  <xdr:graphicFrame>
    <a:graphic><a:graphicData><chart r:id="rId1"/></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:oneCellAnchor>"#,
    );

    let expected = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        ext: EmuSize::new(111, 222),
    };

    assert_anchor_parses_consistently(&drawing_xml, expected);
}

#[test]
fn two_cell_anchor_defaults_missing_offsets_to_zero() {
    let drawing_xml = wrap_wsdr(
        r#"<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
  <xdr:graphicFrame>
    <a:graphic><a:graphicData><chart r:id="rId1"/></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>"#,
    );

    let expected = Anchor::TwoCell {
        from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
        to: AnchorPoint::new(CellRef::new(4, 3), CellOffset::new(0, 0)),
    };

    assert_anchor_parses_consistently(&drawing_xml, expected);
}

#[test]
fn absolute_anchor_parses_without_from_to() {
    let drawing_xml = wrap_wsdr(
        r#"<xdr:absoluteAnchor>
  <xdr:pos x="10" y="20"/>
  <xdr:ext cx="30" cy="40"/>
  <xdr:graphicFrame>
    <a:graphic><a:graphicData><chart r:id="rId1"/></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:absoluteAnchor>"#,
    );

    let expected = Anchor::Absolute {
        pos: CellOffset::new(10, 20),
        ext: EmuSize::new(30, 40),
    };

    assert_anchor_parses_consistently(&drawing_xml, expected);
}
