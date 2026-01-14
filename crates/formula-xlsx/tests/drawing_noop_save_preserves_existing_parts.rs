use formula_xlsx::{load_from_bytes, XlsxPackage};
use formula_model::{CellRef, CellValue};

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

#[test]
fn editing_cells_does_not_rewrite_drawing_parts_or_media() {
    let original_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut doc = load_from_bytes(original_bytes).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(CellRef::from_a1("A1").expect("valid A1"), CellValue::Number(123.0));
    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    assert_eq!(
        before.part("xl/drawings/drawing1.xml").unwrap(),
        after.part("xl/drawings/drawing1.xml").unwrap(),
        "drawing XML should be preserved byte-for-byte when editing unrelated cells"
    );
    assert_eq!(
        before.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        after.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        "drawing relationship XML should be preserved byte-for-byte when editing unrelated cells"
    );
    assert_eq!(
        before.part("xl/media/image1.png").unwrap(),
        after.part("xl/media/image1.png").unwrap(),
        "image media bytes should be preserved byte-for-byte when editing unrelated cells"
    );
}
