use std::collections::BTreeSet;
use std::io::Read;
use std::path::Path;

use roxmltree::Document;
use zip::ZipArchive;

#[test]
fn cell_images_fixture_has_expected_parts_and_uris() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/cell-images.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let cursor = std::io::Cursor::new(bytes);
    let mut zip = ZipArchive::new(cursor)?;

    let mut part_names = BTreeSet::<String>::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_file() {
            part_names.insert(file.name().to_string());
        }
    }

    assert!(
        part_names.contains("xl/cellImages.xml"),
        "missing xl/cellImages.xml"
    );
    assert!(
        part_names.contains("xl/_rels/cellImages.xml.rels"),
        "missing xl/_rels/cellImages.xml.rels"
    );
    assert!(
        part_names.contains("xl/media/image1.png"),
        "missing xl/media/image1.png"
    );

    // Confirm documented [Content_Types].xml override.
    let mut content_types = String::new();
    zip.by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(
        content_types.contains("PartName=\"/xl/cellImages.xml\""),
        "expected [Content_Types].xml override for /xl/cellImages.xml"
    );
    assert!(
        content_types.contains("ContentType=\"application/vnd.ms-excel.cellimages+xml\""),
        "expected [Content_Types].xml ContentType application/vnd.ms-excel.cellimages+xml"
    );

    // Confirm documented workbook relationship type for cellImages.xml.
    let mut workbook_rels = String::new();
    zip.by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut workbook_rels)?;
    assert!(
        workbook_rels.contains("Target=\"cellImages.xml\""),
        "expected workbook.xml.rels to reference cellImages.xml"
    );
    assert!(
        workbook_rels.contains("Type=\"http://schemas.microsoft.com/office/2023/02/relationships/cellImage\""),
        "expected workbook.xml.rels relationship type for cellImages.xml"
    );

    // Confirm documented namespace in xl/cellImages.xml.
    let mut cell_images_xml = String::new();
    zip.by_name("xl/cellImages.xml")?
        .read_to_string(&mut cell_images_xml)?;
    let doc = Document::parse(&cell_images_xml)?;
    let root = doc.root_element();
    assert_eq!(root.tag_name().name(), "cellImages");
    assert_eq!(
        root.tag_name().namespace(),
        Some("http://schemas.microsoft.com/office/spreadsheetml/2023/02/main")
    );

    Ok(())
}

