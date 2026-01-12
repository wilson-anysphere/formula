use base64::Engine;
use formula_model::{CellRef, HyperlinkTarget};
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Image, Workbook};

#[test]
fn embedded_cell_image_includes_hyperlink_target() {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let image = Image::new_from_buffer(&png_bytes)
        .expect("image from buffer")
        .set_url("http://example.com")
        .expect("set image url");
    worksheet
        .embed_image(0, 0, &image)
        .expect("embed image into A1");

    let bytes = workbook.save_to_buffer().expect("save workbook");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("load xlsx package");

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    assert_eq!(images.len(), 1, "expected one embedded cell image");
    let key = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let image = images.get(&key).expect("expected image at Sheet1!A1");
    assert_eq!(image.image_bytes, png_bytes);
    assert_eq!(
        image.hyperlink_target,
        Some(HyperlinkTarget::ExternalUrl {
            uri: "http://example.com".to_string()
        })
    );
}
