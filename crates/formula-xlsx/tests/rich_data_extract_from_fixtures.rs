use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::{rich_data, XlsxPackage};

#[test]
fn images_in_cell_fixture_extracts_image_bytes() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/rich-data/images-in-cell.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let extracted = rich_data::extract_rich_cell_images(&pkg)?;

    assert_eq!(
        extracted.len(),
        1,
        "expected exactly one extracted rich cell image"
    );

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1")?);
    let expected = pkg
        .part("xl/media/image1.png")
        .expect("fixture expected to contain xl/media/image1.png");
    let actual = extracted
        .get(&key)
        .expect("expected (Sheet1, A1) to resolve to an image");

    assert_eq!(actual.as_slice(), expected);

    Ok(())
}

#[test]
fn excel_image_in_cell_fixture_extracts_two_images() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/image-in-cell.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let extracted = rich_data::extract_rich_cell_images(&pkg)?;

    assert_eq!(extracted.len(), 3, "expected three extracted rich cell images");

    let image1 = pkg
        .part("xl/media/image1.png")
        .expect("fixture expected to contain xl/media/image1.png");
    let image2 = pkg
        .part("xl/media/image2.png")
        .expect("fixture expected to contain xl/media/image2.png");

    for cell in ["B2", "B3"] {
        let key = ("Sheet1".to_string(), CellRef::from_a1(cell)?);
        let actual = extracted
            .get(&key)
            .unwrap_or_else(|| panic!("expected (Sheet1, {cell}) to resolve to an image"));
        assert_eq!(
            actual.as_slice(),
            image1,
            "expected (Sheet1, {cell}) to resolve to xl/media/image1.png"
        );
    }

    let key = ("Sheet1".to_string(), CellRef::from_a1("B4")?);
    let actual = extracted
        .get(&key)
        .expect("expected (Sheet1, B4) to resolve to an image");
    assert_eq!(
        actual.as_slice(),
        image2,
        "expected (Sheet1, B4) to resolve to xl/media/image2.png"
    );

    Ok(())
}

