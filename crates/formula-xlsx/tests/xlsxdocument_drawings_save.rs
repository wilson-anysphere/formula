use base64::Engine;
use formula_model::drawings::{
    Anchor, AnchorPoint, CellOffset, DrawingObjectKind, EmuSize, ImageData,
};
use formula_model::CellRef;
use formula_xlsx::drawings::DrawingPart;
use formula_xlsx::{load_from_bytes, XlsxDocument, XlsxPackage};

#[test]
fn xlsxdocument_roundtrip_preserves_floating_images_fixture() {
    let fixture = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let doc = load_from_bytes(fixture).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    assert!(
        pkg.part("xl/drawings/drawing1.xml").is_some(),
        "expected drawing part to be present after save"
    );
    assert!(
        pkg.part("xl/media/image1.png").is_some(),
        "expected image media part to be present after save"
    );

    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet = &reloaded.workbook.sheets[0];
    assert!(
        sheet
            .drawings
            .iter()
            .any(|o| matches!(o.kind, DrawingObjectKind::Image { .. })),
        "expected worksheet.drawings to contain an image object"
    );
}

#[test]
fn xlsxdocument_writes_newly_inserted_sheet_drawing_image() {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let image_id = workbook.images.ensure_unique_name("image", "png");
    workbook.images.insert(
        image_id.clone(),
        ImageData {
            bytes: png_bytes.clone(),
            content_type: Some("image/png".to_string()),
        },
    );

    // Build a drawing object via the existing `DrawingPart::insert_image_object` API.
    let (mut drawing_part, _sheet_drawing_rid) = DrawingPart::create_new(0).expect("create part");
    let anchor = Anchor::OneCell {
        from: AnchorPoint::new(CellRef::new(0, 0), CellOffset::new(0, 0)),
        ext: EmuSize::new(914_400, 914_400),
    };
    drawing_part.insert_image_object(&image_id, anchor);

    workbook.sheet_mut(sheet_id).expect("sheet exists").drawings = drawing_part.objects.clone();

    let doc = XlsxDocument::new(workbook);
    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved workbook");
    assert!(
        pkg.part("xl/drawings/drawing1.xml").is_some(),
        "expected drawing part to be emitted for newly inserted image"
    );

    let media_part_name = format!("xl/media/{}", image_id.as_str());
    assert_eq!(
        pkg.part(&media_part_name),
        Some(png_bytes.as_slice()),
        "expected image bytes to be present in {media_part_name}"
    );

    let reloaded = load_from_bytes(&saved).expect("reload");
    let sheet = &reloaded.workbook.sheets[0];
    assert!(
        sheet.drawings.iter().any(|o| matches!(
            &o.kind,
            DrawingObjectKind::Image { image_id: got } if got == &image_id
        )),
        "expected worksheet.drawings to contain the inserted image object"
    );
}
