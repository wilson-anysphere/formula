use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::drawings::ImageId;
use zip::ZipArchive;

#[test]
fn cellimages_fixture_imports_blip_embed_without_pic_wrapper(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/cellimages.xlsx");
    let bytes = std::fs::read(&fixture_path)
        .unwrap_or_else(|e| panic!("read fixture {}: {e}", fixture_path.display()));

    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    let image_id = ImageId::new("image1.png");
    let image = doc
        .workbook
        .images
        .get(&image_id)
        .expect("expected Workbook.images to contain image1.png");

    let mut archive = ZipArchive::new(Cursor::new(&bytes)).expect("open xlsx as zip");
    let mut expected_bytes = Vec::new();
    archive
        .by_name("xl/media/image1.png")
        .expect("fixture contains xl/media/image1.png")
        .read_to_end(&mut expected_bytes)
        .expect("read xl/media/image1.png");

    assert_eq!(image.bytes, expected_bytes);

    Ok(())
}

