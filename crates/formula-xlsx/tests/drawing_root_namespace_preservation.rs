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
