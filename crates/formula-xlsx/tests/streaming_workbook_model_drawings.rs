use formula_model::drawings::DrawingObjectKind;

#[test]
fn streaming_reader_loads_drawing_images() {
    let bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(bytes).expect("read workbook model");

    let sheet = workbook
        .sheets
        .first()
        .expect("expected at least one sheet");
    let img_obj = sheet
        .drawings
        .iter()
        .find(|obj| matches!(obj.kind, DrawingObjectKind::Image { .. }))
        .expect("expected an image drawing object");

    let DrawingObjectKind::Image { image_id } = &img_obj.kind else {
        unreachable!("just matched image");
    };

    let data = workbook
        .images
        .get(image_id)
        .expect("expected image bytes to be loaded into workbook.images");
    assert!(
        !data.bytes.is_empty(),
        "expected non-empty image bytes for {image_id:?}"
    );
}
