use std::collections::BTreeMap;

use formula_model::drawings::Anchor;
use formula_model::CellRef;

use formula_xlsx::drawings::DrawingPart;

#[test]
fn drawing_anchor_missing_offsets_defaults_to_zero() {
    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:row>2</xdr:row>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col>
      <xdr:row>4</xdr:row>
    </xdr:to>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"#
    .to_vec();

    let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"#
    .to_vec();

    let mut parts = BTreeMap::new();
    parts.insert("xl/drawings/drawing1.xml".to_string(), drawing_xml);
    parts.insert(
        "xl/drawings/_rels/drawing1.xml.rels".to_string(),
        rels_xml,
    );

    let mut workbook = formula_model::Workbook::new();
    let part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("drawing part should parse even when colOff/rowOff are missing");

    assert_eq!(part.objects.len(), 1);
    match part.objects[0].anchor {
        Anchor::TwoCell { from, to } => {
            assert_eq!(from.cell, CellRef::new(2, 1));
            assert_eq!(to.cell, CellRef::new(4, 3));
            assert_eq!(from.offset.x_emu, 0);
            assert_eq!(from.offset.y_emu, 0);
            assert_eq!(to.offset.x_emu, 0);
            assert_eq!(to.offset.y_emu, 0);
        }
        other => panic!("unexpected anchor: {other:?}"),
    }
}

#[test]
fn drawing_anchor_absolute_anchor_parses_without_from_to() {
    // `xdr:absoluteAnchor` does not include <from>/<to>; it uses absolute EMU positioning.
    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x=" 100 " y=" 200 "/>
    <xdr:ext cx=" 300 " cy=" 400 "/>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>
"#
    .to_vec();

    let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"#
    .to_vec();

    let mut parts = BTreeMap::new();
    parts.insert("xl/drawings/drawing1.xml".to_string(), drawing_xml);
    parts.insert(
        "xl/drawings/_rels/drawing1.xml.rels".to_string(),
        rels_xml,
    );

    let mut workbook = formula_model::Workbook::new();
    let part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("absoluteAnchor should parse without <from>/<to>");

    assert_eq!(part.objects.len(), 1);
    assert_eq!(
        part.objects[0].anchor,
        Anchor::Absolute {
            pos: formula_model::drawings::CellOffset::new(100, 200),
            ext: formula_model::drawings::EmuSize::new(300, 400),
        }
    );
}
