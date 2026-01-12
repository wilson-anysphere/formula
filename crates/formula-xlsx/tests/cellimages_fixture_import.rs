use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::drawings::ImageId;
use formula_xlsx::XlsxPackage;
use zip::ZipArchive;

#[test]
fn cellimages_fixture_cell_images_part_info_resolves_blip_embed_target(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/cellimages.xlsx");
    let bytes = std::fs::read(&fixture_path)
        .unwrap_or_else(|e| panic!("read fixture {}: {e}", fixture_path.display()));

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let part_info = pkg
        .cell_images_part_info()?
        .expect("expected fixture to contain a cellimages part");

    assert_eq!(part_info.part_path, "xl/cellimages.xml");
    assert_eq!(part_info.rels_path, "xl/_rels/cellimages.xml.rels");
    assert_eq!(part_info.embeds.len(), 1);
    assert_eq!(part_info.embeds[0].embed_rid, "rId1");
    assert_eq!(part_info.embeds[0].target_part, "xl/media/image1.png");

    let mut archive = ZipArchive::new(Cursor::new(&bytes)).expect("open xlsx as zip");
    let mut expected_bytes = Vec::new();
    archive
        .by_name("xl/media/image1.png")
        .expect("fixture contains xl/media/image1.png")
        .read_to_end(&mut expected_bytes)
        .expect("read xl/media/image1.png");

    assert_eq!(
        part_info.embeds[0].target_bytes, expected_bytes,
        "expected part_info to resolve embedded image bytes"
    );

    // Ensure the resolved image bytes match what the document reader exports too.
    let doc = formula_xlsx::load_from_bytes(&bytes)?;
    let image_id = ImageId::new("image1.png");
    let image = doc
        .workbook
        .images
        .get(&image_id)
        .expect("expected Workbook.images to contain image1.png");
    assert_eq!(image.bytes, part_info.embeds[0].target_bytes);

    Ok(())
}
