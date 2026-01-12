use std::io::{Cursor, Read, Write};

use formula_model::drawings::ImageId;
use zip::ZipArchive;

fn build_minimal_cellimages_xlsx(
    cellimages_rel_target: &str,
    image_bytes: Option<&[u8]>,
) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    // Minimal `xl/cellimages.xml` that references an image relationship ID.
    let cellimages_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2020/11/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <etc:cellImage r:id="rId1"/>
</etc:cellImages>"#;

    let cellimages_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{cellimages_rel_target}"/>
</Relationships>"#
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/cellimages.xml", options).unwrap();
    zip.write_all(cellimages_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/cellimages.xml.rels", options)
        .unwrap();
    zip.write_all(cellimages_rels.as_bytes()).unwrap();

    if let Some(image_bytes) = image_bytes {
        zip.start_file("xl/media/image1.png", options).unwrap();
        zip.write_all(image_bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn load_from_bytes_populates_workbook_images_from_cellimages() -> Result<(), Box<dyn std::error::Error>>
{
    let expected = b"png-bytes";
    let bytes = build_minimal_cellimages_xlsx("media/image1.png", Some(expected));
    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    let image = doc
        .workbook
        .images
        .get(&ImageId::new("image1.png"))
        .expect("expected Workbook.images to contain image1.png");
    assert_eq!(image.bytes.as_slice(), expected);

    Ok(())
}

#[test]
fn cellimages_load_from_bytes_tolerates_parent_dir_media_target_for_lightweight_schema(
) -> Result<(), Box<dyn std::error::Error>> {
    // Some producers (including older/newer Excel variants) appear to emit `../media/*` targets
    // from workbook-level `xl/cellimages.xml.rels`, even though the actual media parts live under
    // `xl/media/*`. We should resolve this best-effort rather than dropping the image.
    let expected = b"png-bytes-from-parent-dir-target";
    let bytes = build_minimal_cellimages_xlsx("../media/image1.png", Some(expected));
    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    let image = doc
        .workbook
        .images
        .get(&ImageId::new("image1.png"))
        .expect("expected Workbook.images to contain image1.png");
    assert_eq!(image.bytes.as_slice(), expected);

    Ok(())
}

#[test]
fn cellimages_missing_media_is_best_effort() -> Result<(), Box<dyn std::error::Error>> {
    // Missing referenced media should not fail workbook load; we should just skip that image.
    let bytes = build_minimal_cellimages_xlsx("media/image1.png", None);
    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    assert!(doc.workbook.images.get(&ImageId::new("image1.png")).is_none());
    Ok(())
}

#[test]
fn load_from_bytes_extracts_images_from_fixture_cellimages_blip_schema(
) -> Result<(), Box<dyn std::error::Error>> {
    // The **synthetic** fixture `fixtures/xlsx/basic/cellimages.xlsx` uses a lightweight
    // `<cellImage><a:blip r:embed="rId1"/></cellImage>` shape (no `<xdr:pic>` and no `r:id` on
    // `<cellImage>`). Ensure we can still discover and load the referenced image into
    // `Workbook.images`.
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/cellimages.xlsx");
    let bytes = std::fs::read(&fixture_path)
        .unwrap_or_else(|e| panic!("read fixture {}: {e}", fixture_path.display()));

    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    // Read the expected bytes from the fixture's media part.
    let mut archive = ZipArchive::new(Cursor::new(&bytes))?;
    let mut file = archive.by_name("xl/media/image1.png")?;
    let mut expected = Vec::new();
    file.read_to_end(&mut expected)?;

    let image = doc
        .workbook
        .images
        .get(&ImageId::new("image1.png"))
        .expect("expected Workbook.images to contain image1.png");
    assert_eq!(image.bytes.as_slice(), expected.as_slice());
    Ok(())
}
