use std::collections::BTreeMap;

use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
    ImageId,
};
use formula_model::CellRef;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::openxml;

const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

#[test]
fn drawing_part_from_objects_repairs_missing_image_relationship_type() {
    // Simulate a producer that omits `Type` on an image relationship.
    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="../media/image1.png"/>
</Relationships>"#;

    let mut preserved = std::collections::HashMap::new();
    preserved.insert("xlsx.embed_rel_id".to_string(), "rId1".to_string());

    let ext = EmuSize::new(914_400, 914_400);
    let objects = vec![DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image {
            image_id: ImageId::new("image1.png"),
        },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext,
        },
        z_order: 0,
        size: Some(ext),
        preserved,
    }];

    let mut drawing_part = DrawingPart::from_objects(
        0,
        "xl/drawings/drawing1.xml".to_string(),
        objects,
        Some(rels_xml),
    )
    .expect("build DrawingPart");

    let mut parts = BTreeMap::<String, Vec<u8>>::new();
    // `DrawingPart::write_into_parts` validates that image relationships point at an existing
    // media part (either already present in `parts` or supplied via `workbook.images`).
    // This test focuses on repairing the relationship metadata, so we provide a dummy payload.
    parts.insert("xl/media/image1.png".to_string(), vec![0u8; 8]);
    let workbook = formula_model::Workbook::new();
    drawing_part
        .write_into_parts(&mut parts, &workbook)
        .expect("write parts");

    let rels_bytes = parts
        .get("xl/drawings/_rels/drawing1.xml.rels")
        .expect("drawing rels written");
    let rels = openxml::parse_relationships(rels_bytes).expect("parse rels");
    let rel = rels
        .iter()
        .find(|rel| rel.id == "rId1")
        .expect("rId1 relationship preserved");

    assert_eq!(rel.type_uri, REL_TYPE_IMAGE);
    assert_eq!(rel.target, "../media/image1.png");
    assert_eq!(rel.target_mode, None);
}
