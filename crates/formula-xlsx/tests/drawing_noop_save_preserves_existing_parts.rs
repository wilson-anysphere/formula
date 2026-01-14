use formula_xlsx::{load_from_bytes, XlsxPackage};

#[test]
fn noop_save_preserves_drawing_parts_and_media_bytes() {
    let original_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let doc = load_from_bytes(original_bytes).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    assert_eq!(
        before.part("xl/drawings/drawing1.xml").unwrap(),
        after.part("xl/drawings/drawing1.xml").unwrap(),
        "drawing XML should be preserved byte-for-byte on no-op save"
    );
    assert_eq!(
        before.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        after.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        "drawing relationship XML should be preserved byte-for-byte on no-op save"
    );
    assert_eq!(
        before.part("xl/media/image1.png").unwrap(),
        after.part("xl/media/image1.png").unwrap(),
        "image media bytes should be preserved byte-for-byte on no-op save"
    );
}

