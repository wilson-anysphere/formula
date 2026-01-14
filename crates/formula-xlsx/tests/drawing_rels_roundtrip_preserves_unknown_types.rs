use std::collections::BTreeMap;

use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::{openxml, XlsxPackage};

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");
const DRAWING_RELS: &str = "xl/drawings/_rels/drawing1.xml.rels";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

fn rels_by_id(rels: Vec<openxml::Relationship>) -> BTreeMap<String, openxml::Relationship> {
    rels.into_iter().map(|r| (r.id.clone(), r)).collect()
}

#[test]
fn drawing_part_roundtrip_preserves_unknown_relationship_types_and_target_mode() {
    let mut pkg = XlsxPackage::from_bytes(FIXTURE).expect("load smartart.xlsx fixture");

    // Inject a TargetMode attribute into one of the SmartArt diagram relationships so we can
    // assert it round-trips through `DrawingPart::parse_from_parts` -> `write_into_parts`.
    let rels_xml = std::str::from_utf8(
        pkg.part(DRAWING_RELS)
            .expect("fixture must contain drawing1.xml.rels"),
    )
    .expect("rels xml must be utf-8");
    let rels_xml = rels_xml.replace(
        r#"Target="../diagrams/data1.xml"/>"#,
        r#"Target="../diagrams/data1.xml" TargetMode="External"/>"#,
    );
    pkg.parts_map_mut()
        .insert(DRAWING_RELS.to_string(), rels_xml.as_bytes().to_vec());

    let original_rels = rels_by_id(
        openxml::parse_relationships(pkg.part(DRAWING_RELS).unwrap()).expect("parse rels"),
    );
    assert!(
        original_rels.values().any(|rel| rel.type_uri.contains("diagramData")),
        "expected fixture rels to include diagram relationships"
    );
    assert!(
        original_rels
            .values()
            .any(|rel| rel.target_mode.as_deref() == Some("External")),
        "expected test setup to inject a TargetMode=\"External\" relationship"
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
        .get(DRAWING_RELS)
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

#[test]
fn drawing_part_from_objects_preserves_existing_relationships() {
    use formula_model::drawings::{
        Anchor, AnchorPoint, CellOffset, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize,
        ImageId,
    };
    use formula_model::CellRef;

    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("load smartart.xlsx fixture");
    let rels_xml = std::str::from_utf8(pkg.part(DRAWING_RELS).unwrap()).unwrap();
    let rels_xml = rels_xml.replace(
        r#"Target="../diagrams/data1.xml"/>"#,
        r#"Target="../diagrams/data1.xml" TargetMode="External"/>"#,
    );

    let objects = vec![DrawingObject {
        id: DrawingObjectId(1),
        kind: DrawingObjectKind::Image {
            image_id: ImageId::new("image1.png"),
        },
        anchor: Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
            ext: EmuSize::new(914_400, 914_400),
        },
        z_order: 0,
        size: Some(EmuSize::new(914_400, 914_400)),
        preserved: Default::default(),
    }];

    let mut drawing_part = DrawingPart::from_objects(
        0,
        "xl/drawings/drawing1.xml".to_string(),
        objects,
        Some(&rels_xml),
    )
    .expect("build DrawingPart from objects");

    let mut parts = BTreeMap::<String, Vec<u8>>::new();
    // `DrawingPart::write_into_parts` validates that image relationships point at an existing
    // media part (either already present in `parts` or supplied via `workbook.images`).
    // This test focuses on relationship preservation, so we provide a dummy media payload.
    parts.insert("xl/media/image1.png".to_string(), vec![0u8; 8]);
    let workbook = formula_model::Workbook::new();
    drawing_part
        .write_into_parts(&mut parts, &workbook)
        .expect("write drawing parts");

    let expected_rels = rels_by_id(openxml::parse_relationships(rels_xml.as_bytes()).unwrap());
    assert!(
        expected_rels
            .values()
            .any(|rel| rel.target_mode.as_deref() == Some("External")),
        "expected test setup to inject a TargetMode=\"External\" relationship"
    );
    let written_rels = rels_by_id(
        openxml::parse_relationships(parts.get(DRAWING_RELS).unwrap()).expect("parse written rels"),
    );

    // Existing (unknown) relationships from the source `.rels` should be preserved.
    for (id, original) in &expected_rels {
        let Some(written) = written_rels.get(id) else {
            panic!("expected relationship {id} to be preserved");
        };
        assert_eq!(original.type_uri, written.type_uri);
        assert_eq!(original.target, written.target);
        assert_eq!(original.target_mode, written.target_mode);
    }

    // And relationships added for new objects should be appended without clobbering existing IDs.
    assert!(
        written_rels.values().any(|rel| {
            rel.type_uri == REL_TYPE_IMAGE && rel.target == "../media/image1.png"
        }),
        "expected an image relationship to be added, got: {written_rels:?}"
    );
}
