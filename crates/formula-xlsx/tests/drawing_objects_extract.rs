use formula_model::drawings::DrawingObjectKind;
use formula_xlsx::XlsxPackage;

#[test]
fn extract_drawing_objects_finds_image() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/image.xlsx"
    ))
    .expect("fixture exists");

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read fixture package");
    let drawings = pkg
        .extract_drawing_objects()
        .expect("extract drawing objects");

    let image_count = drawings
        .iter()
        .flat_map(|entry| entry.objects.iter())
        .filter(|obj| matches!(obj.kind, DrawingObjectKind::Image { .. }))
        .count();

    assert_eq!(image_count, 1, "expected one image object in fixture");
}

