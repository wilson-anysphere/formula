use std::collections::BTreeMap;

use roxmltree::Document;

use formula_xlsx::drawings::{load_media_parts, DrawingPart};
use formula_xlsx::XlsxPackage;

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");

#[test]
fn drawing_part_root_xmlns_are_preserved_for_smartart() {
    let mut pkg = XlsxPackage::from_bytes(FIXTURE).expect("load smartart.xlsx fixture");
    let mut workbook = formula_model::Workbook::new();
    load_media_parts(&mut workbook, pkg.parts_map());

    let mut part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        pkg.parts_map(),
        &mut workbook,
    )
    .expect("parse drawing part");

    part.write_into_parts(pkg.parts_map_mut(), &workbook)
        .expect("write drawing part");

    let drawing_xml = std::str::from_utf8(
        pkg.part("xl/drawings/drawing1.xml")
            .expect("drawing1.xml should exist"),
    )
    .expect("drawing xml should be utf-8");

    assert!(
        drawing_xml.contains(r#"xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram""#),
        "expected rewritten drawing root to preserve SmartArt `xmlns:dgm` declaration"
    );

    Document::parse(drawing_xml).expect("rewritten drawing xml should be parseable");
}

#[test]
fn drawing_part_root_attrs_are_preserved() {
    // Markup-compatibility attributes like `mc:Ignorable` must be preserved when round-tripping;
    // they affect how Excel interprets extension elements.
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="dgm">
</xdr:wsDr>"#;

    let drawing_path = "xl/drawings/drawing1.xml";
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(drawing_path.to_string(), xml.as_bytes().to_vec());

    let mut workbook = formula_model::Workbook::new();
    let mut part =
        DrawingPart::parse_from_parts(0, drawing_path, &parts, &mut workbook).expect("parse");
    part.write_into_parts(&mut parts, &workbook).expect("write");

    let out = std::str::from_utf8(parts.get(drawing_path).expect("drawing part"))
        .expect("output xml should be utf-8");
    assert!(
        out.contains(r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#),
        "expected rewritten drawing root to preserve xmlns:mc, got:\n{out}"
    );
    assert!(
        out.contains(r#"mc:Ignorable="dgm""#),
        "expected rewritten drawing root to preserve mc:Ignorable, got:\n{out}"
    );
    Document::parse(out).expect("rewritten drawing xml should be parseable");
}
