use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::XlsxPackage;

#[test]
fn extracts_embedded_cell_image_from_real_excel_fixture() {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/rich-data/images-in-cell.xlsx");
    let bytes = std::fs::read(&fixture_path).expect("read fixture");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("open xlsx");

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let key = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let img = images.get(&key).expect("expected embedded image at Sheet1!A1");
    assert_eq!(img.image_part, "xl/media/image1.png");
    assert_eq!(img.image_bytes.as_slice(), pkg.part("xl/media/image1.png").unwrap());
}

