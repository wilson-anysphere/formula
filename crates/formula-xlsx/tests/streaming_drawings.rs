use std::path::Path;

use formula_xlsx::read_workbook_model_from_bytes;

#[test]
fn streaming_reader_populates_worksheet_drawings_and_workbook_images() {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/image.xlsx");
    let bytes =
        std::fs::read(&fixture_path).unwrap_or_else(|_| panic!("read {}", fixture_path.display()));

    let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    let sheet = workbook.sheets.first().expect("expected at least one sheet");

    assert!(
        !sheet.drawings.is_empty(),
        "expected worksheet.drawings to be populated"
    );
    assert!(
        !workbook.images.is_empty(),
        "expected workbook.images to contain at least one image"
    );
}

