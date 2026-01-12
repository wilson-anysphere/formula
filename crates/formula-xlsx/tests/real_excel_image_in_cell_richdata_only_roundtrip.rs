use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn real_excel_image_in_cell_richdata_only_roundtrip_preserves_parts_and_vm() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/image-in-cell.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    // Confirm this was saved by Excel (fixture provenance).
    let app_xml = String::from_utf8(zip_part(&fixture_bytes, "docProps/app.xml"))?;
    assert!(
        app_xml.contains("<Application>Microsoft Excel</Application>"),
        "expected docProps/app.xml to indicate Microsoft Excel, got: {app_xml}"
    );

    // Capture original bytes for rich-data parts that should round-trip byte-for-byte when we
    // edit a different cell.
    let original_parts: Vec<(&str, Vec<u8>)> = [
        "xl/metadata.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/media/image1.png",
        "xl/media/image2.png",
    ]
    .into_iter()
    .map(|name| (name, zip_part(&fixture_bytes, name)))
    .collect();

    let mut doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    // Edit a different cell (A2) to exercise the patch/write pipeline while ensuring the in-cell
    // images and their supporting parts remain preserved.
    assert!(doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("A2")?,
        CellValue::Number(42.0)
    ));

    let saved = doc.save_to_vec()?;

    // Rich-data parts should survive byte-for-byte.
    for (name, original) in original_parts {
        assert_eq!(
            zip_part(&saved, name),
            original,
            "expected {name} to be preserved byte-for-byte"
        );
    }

    // The image cells should still carry `vm` attributes.
    let saved_sheet_xml = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml"))?;
    let parsed = roxmltree::Document::parse(&saved_sheet_xml)?;
    for (addr, expected_vm) in [("B2", "1"), ("B3", "1"), ("B4", "2")] {
        let cell = parsed
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(addr))
            .unwrap_or_else(|| panic!("expected {addr} cell in sheet1.xml"));
        assert_eq!(
            cell.attribute("vm"),
            Some(expected_vm),
            "expected {addr} to preserve vm={expected_vm}, got: {saved_sheet_xml}"
        );
    }

    Ok(())
}

