use std::collections::BTreeMap;

use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::{openxml, XlsxPackage};

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");

fn rels_by_id(rels: Vec<openxml::Relationship>) -> BTreeMap<String, openxml::Relationship> {
    rels.into_iter().map(|r| (r.id.clone(), r)).collect()
}

#[test]
fn drawing_part_roundtrip_preserves_unknown_relationship_types() {
    let mut pkg = XlsxPackage::from_bytes(FIXTURE).expect("load smartart.xlsx fixture");

    let original_rels = pkg
        .part("xl/drawings/_rels/drawing1.xml.rels")
        .expect("fixture must contain drawing1.xml.rels");
    let original_rels = rels_by_id(openxml::parse_relationships(original_rels).expect("parse rels"));
    assert!(
        original_rels.values().any(|rel| rel.type_uri.contains("diagramData")),
        "expected fixture rels to include diagram relationships"
    );

    let mut workbook = formula_model::Workbook::new();
    let part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        pkg.parts_map(),
        &mut workbook,
    )
    .expect("parse DrawingPart");

    let mut part = part;
    part.write_into_parts(pkg.parts_map_mut(), &workbook)
        .expect("write DrawingPart");

    let written_rels_bytes = pkg
        .parts_map()
        .get("xl/drawings/_rels/drawing1.xml.rels")
        .expect("written package should contain drawing rels");
    let written_rels =
        rels_by_id(openxml::parse_relationships(written_rels_bytes).expect("parse written rels"));

    assert_eq!(
        original_rels.len(),
        written_rels.len(),
        "relationship count should be preserved"
    );

    for (id, original) in &original_rels {
        let Some(written) = written_rels.get(id) else {
            panic!("expected relationship {id} to be preserved");
        };
        assert_eq!(
            original.type_uri, written.type_uri,
            "relationship type should be preserved for {id}"
        );
        assert_eq!(
            original.target, written.target,
            "relationship target should be preserved for {id}"
        );
        assert_eq!(
            original.target_mode, written.target_mode,
            "relationship TargetMode should be preserved for {id}"
        );
    }
}

