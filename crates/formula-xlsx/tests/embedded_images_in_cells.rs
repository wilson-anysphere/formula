use base64::Engine as _;
use formula_model::CellRef;
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Image, Workbook};

const ONE_BY_ONE_PNG_BASE64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/w8AAgMBApZ9xO4AAAAASUVORK5CYII=";

fn one_by_one_png_bytes() -> Vec<u8> {
    base64::engine::general_purpose::STANDARD
        .decode(ONE_BY_ONE_PNG_BASE64)
        .expect("decode png base64")
}

fn write_temp_png(bytes: &[u8]) -> (tempfile::TempDir, std::path::PathBuf) {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = dir.path().join("image.png");
    std::fs::write(&path, bytes).expect("write png");
    (dir, path)
}

fn build_workbook_with_embedded_image(
    alt_text: Option<&str>,
    include_dynamic_array: bool,
) -> Vec<u8> {
    let png = one_by_one_png_bytes();
    let (_dir, image_path) = write_temp_png(&png);

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    if include_dynamic_array {
        // Force Excel to emit XLDAPR metadata alongside rich value metadata.
        // `SEQUENCE()` is a dynamic array function.
        worksheet
            .write_dynamic_array_formula(0, 1, 2, 1, "=SEQUENCE(3)")
            .unwrap();
    }

    let mut image = Image::new(&image_path).expect("create image");
    if let Some(text) = alt_text {
        image = image.set_alt_text(text);
    }

    worksheet.embed_image(0, 0, &image).unwrap();
    workbook.save_to_buffer().unwrap()
}

#[test]
fn extracts_single_embedded_image_in_cell() {
    let bytes = build_workbook_with_embedded_image(None, false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    assert_eq!(images.len(), 1);

    let key = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let cell_img = images.get(&key).expect("expected A1 embedded image");
    let stored = &cell_img.image_bytes;
    assert_eq!(
        stored,
        &one_by_one_png_bytes(),
        "expected extracted image bytes to match the inserted PNG"
    );
}

#[test]
fn extracts_alt_text_from_embedded_image() {
    let bytes = build_workbook_with_embedded_image(Some("hello alt text"), false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    assert_eq!(images.len(), 1);

    let key = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let cell_img = images.get(&key).expect("expected A1 image");
    assert_eq!(cell_img.alt_text.as_deref(), Some("hello alt text"));
    assert_eq!(cell_img.calc_origin, 5);
}

#[test]
fn dynamic_array_metadata_does_not_break_embedded_image_extraction() {
    let bytes = build_workbook_with_embedded_image(None, true);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let metadata_xml = std::str::from_utf8(pkg.part("xl/metadata.xml").unwrap()).unwrap();
    assert!(
        metadata_xml.contains("XLDAPR"),
        "expected dynamic array metadata type in xl/metadata.xml"
    );
    assert!(
        metadata_xml.contains("XLRICHVALUE"),
        "expected rich value metadata type in xl/metadata.xml"
    );

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    let addr = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    assert!(
        images.contains_key(&addr),
        "expected embedded image mapping even with dynamic array metadata present"
    );
}
