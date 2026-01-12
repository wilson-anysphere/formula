use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::drawings::ImageId;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    extract_embedded_images, load_from_bytes, parse_value_metadata_vm_to_rich_value_index_map,
    XlsxPackage,
};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn zip_part_exists(zip_bytes: &[u8], name: &str) -> bool {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    // `ZipFile` borrows the archive, so ensure the result is dropped before `archive`.
    let exists = archive.by_name(name).is_ok();
    exists
}

#[test]
fn real_excel_images_in_cell_roundtrip_preserves_richdata_and_loads_media(
) -> Result<(), Box<dyn std::error::Error>> {
    // This fixture was created to represent a modern Excel 365 “Place in Cell” image workbook and
    // is expected to contain the full RichData + cellImages wiring:
    //
    //   unzip -Z1 fixtures/xlsx/rich-data/images-in-cell.xlsx | sort
    //
    // Must include at least:
    // - xl/cellimages.xml + xl/_rels/cellimages.xml.rels
    // - xl/metadata.xml + xl/_rels/metadata.xml.rels
    // - xl/richData/richValue.xml
    // - xl/richData/richValueRel.xml + xl/richData/_rels/richValueRel.xml.rels
    // - xl/richData/richValueTypes.xml
    // - xl/richData/richValueStructure.xml
    // - xl/media/image*.png
    // - c/@vm present on the image cell in xl/worksheets/sheet1.xml
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/rich-data/images-in-cell.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    // Ensure the embedded-image extractor can resolve the in-cell image for this real Excel file.
    let pkg = XlsxPackage::from_bytes(&fixture_bytes)?;
    let embedded = extract_embedded_images(&pkg)?;
    assert_eq!(embedded.len(), 1, "expected one embedded in-cell image");
    assert_eq!(embedded[0].sheet_part, "xl/worksheets/sheet1.xml");
    assert_eq!(embedded[0].cell, CellRef::from_a1("A1")?);
    assert_eq!(embedded[0].image_target, "xl/media/image1.png");
    assert_eq!(
        embedded[0].bytes,
        zip_part(&fixture_bytes, "xl/media/image1.png")
    );

    let app_xml = String::from_utf8(zip_part(&fixture_bytes, "docProps/app.xml"))?;
    assert!(
        app_xml.contains("<Application>Microsoft Excel</Application>"),
        "expected docProps/app.xml to indicate the workbook was saved by Excel, got: {app_xml}"
    );

    for part in [
        "xl/cellimages.xml",
        "xl/_rels/cellimages.xml.rels",
        "xl/metadata.xml",
        "xl/_rels/metadata.xml.rels",
        "xl/richData/richValue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/richValueTypes.xml",
        "xl/richData/richValueStructure.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert!(
            zip_part_exists(&fixture_bytes, part),
            "expected fixture to contain {part}"
        );
    }

    // Fixture-backed confirmations used by docs/20-images-in-cells.md.
    //
    // Workbook relationship type:
    // - Type="http://schemas.microsoft.com/office/2019/relationships/cellimages"
    // - Target="cellimages.xml"
    let workbook_rels_xml = String::from_utf8(zip_part(&fixture_bytes, "xl/_rels/workbook.xml.rels"))?;
    let workbook_rels_doc = roxmltree::Document::parse(&workbook_rels_xml)?;
    let cellimages_rel = workbook_rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Target") == Some("cellimages.xml")
        })
        .expect("expected workbook.xml.rels Relationship targeting cellimages.xml");
    assert_eq!(
        cellimages_rel.attribute("Type"),
        Some("http://schemas.microsoft.com/office/2019/relationships/cellimages"),
        "unexpected workbook relationship Type for cellimages.xml: {workbook_rels_xml}"
    );

    // cellimages.xml root namespace.
    let cellimages_xml = String::from_utf8(zip_part(&fixture_bytes, "xl/cellimages.xml"))?;
    let cellimages_doc = roxmltree::Document::parse(&cellimages_xml)?;
    let cellimages_root = cellimages_doc.root_element();
    assert_eq!(cellimages_root.tag_name().name(), "cellImages");
    assert_eq!(
        cellimages_root.tag_name().namespace(),
        Some("http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"),
        "unexpected cellimages.xml root namespace: {cellimages_xml}"
    );

    // [Content_Types].xml override for cellimages.xml.
    let content_types_xml =
        String::from_utf8(zip_part(&fixture_bytes, "[Content_Types].xml"))?;
    assert!(
        content_types_xml.contains("PartName=\"/xl/cellimages.xml\""),
        "expected [Content_Types].xml to contain an Override for /xl/cellimages.xml, got: {content_types_xml}"
    );
    assert!(
        content_types_xml.contains(
            "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml\""
        ),
        "expected [Content_Types].xml to use the cellimages content type override observed in Excel, got: {content_types_xml}"
    );

    let sheet_xml = String::from_utf8(zip_part(&fixture_bytes, "xl/worksheets/sheet1.xml"))?;
    assert!(
        sheet_xml.contains(r#"<c r="A1" vm=""#),
        "expected Sheet1!A1 cell to include a vm=\"...\" attribute, got: {sheet_xml}"
    );

    let metadata_bytes = zip_part(&fixture_bytes, "xl/metadata.xml");
    let vm_map = parse_value_metadata_vm_to_rich_value_index_map(&metadata_bytes)?;
    assert_eq!(
        vm_map.get(&1),
        Some(&0),
        "expected vm=1 to resolve to rich value index 0 via xl/metadata.xml"
    );

    // Capture original bytes for all rich-data parts we need to preserve byte-for-byte.
    let original_parts: Vec<(&str, Vec<u8>)> = [
        "xl/cellimages.xml",
        "xl/_rels/cellimages.xml.rels",
        "xl/metadata.xml",
        "xl/_rels/metadata.xml.rels",
        "xl/richData/richValue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/richValueTypes.xml",
        "xl/richData/richValueStructure.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ]
    .into_iter()
    .map(|name| (name, zip_part(&fixture_bytes, name)))
    .collect();

    let mut doc = load_from_bytes(&fixture_bytes)?;

    // The document loader should import media referenced by xl/cellimages.xml into workbook.images.
    assert!(
        !doc.workbook.images.is_empty(),
        "expected workbook.images to be non-empty"
    );
    let image_id = ImageId::new("image1.png");
    let stored = doc
        .workbook
        .images
        .get(&image_id)
        .expect("expected workbook.images to contain image1.png");
    assert_eq!(
        stored.bytes,
        zip_part(&fixture_bytes, "xl/media/image1.png"),
        "expected workbook.images bytes for image1.png to match the xl/media/image1.png part"
    );

    // Edit an unrelated cell to exercise worksheet patching while preserving rich-data parts.
    let sheet_id = doc.workbook.sheets[0].id;
    assert!(doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("B1")?,
        CellValue::Number(42.0)
    ));

    let saved = doc.save_to_vec()?;

    // Ensure the rich-data parts and image bytes survive byte-for-byte after write.
    for (name, original) in original_parts {
        assert_eq!(
            zip_part(&saved, name),
            original,
            "expected {name} to be preserved byte-for-byte"
        );
    }

    // Ensure the vm attribute on the image cell is still present after editing another cell.
    let saved_sheet_xml = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml"))?;
    let parsed = roxmltree::Document::parse(&saved_sheet_xml)?;
    let cell_a1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(
        cell_a1.attribute("vm"),
        Some("1"),
        "expected vm attribute to be preserved, got: {saved_sheet_xml}"
    );
    assert_eq!(
        cell_a1.attribute("cm"),
        Some("1"),
        "expected cm attribute to be preserved, got: {saved_sheet_xml}"
    );

    Ok(())
}
