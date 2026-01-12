use std::collections::BTreeMap;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use roxmltree::Document;
use zip::ZipArchive;

const REQUIRED_PARTS: &[&str] = &[
    "xl/metadata.xml",
    "xl/richData/richValue.xml",
    "xl/richData/richValueRel.xml",
    "xl/richData/_rels/richValueRel.xml.rels",
    "xl/media/image1.png",
];

const XML_PARTS: &[&str] = &[
    "xl/metadata.xml",
    "xl/richData/richValue.xml",
    "xl/richData/richValueRel.xml",
    "xl/richData/_rels/richValueRel.xml.rels",
];

fn read_zip_parts(bytes: &[u8], part_names: &[&str]) -> Result<BTreeMap<String, Vec<u8>>, zip::result::ZipError> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut out = BTreeMap::new();
    for name in part_names {
        let mut f = archive.by_name(name)?;
        let mut buf = Vec::with_capacity(f.size() as usize);
        f.read_to_end(&mut buf)?;
        out.insert((*name).to_string(), buf);
    }
    Ok(out)
}

fn assert_xml_bytes_equal_or_semantic(part_name: &str, expected: &[u8], actual: &[u8]) {
    if expected == actual {
        return;
    }

    let expected_norm = formula_xlsx::normalize_xml(expected).expect("normalize expected xml");
    let actual_norm = formula_xlsx::normalize_xml(actual).expect("normalize actual xml");
    assert_eq!(
        expected_norm, actual_norm,
        "XML part changed after round-trip: {part_name}"
    );
}

fn assert_rich_data_parts_preserved(original: &[u8], roundtripped: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let original_parts = read_zip_parts(original, REQUIRED_PARTS)?;
    let out_parts = read_zip_parts(roundtripped, REQUIRED_PARTS)?;

    for part_name in REQUIRED_PARTS {
        assert!(
            out_parts.contains_key(*part_name),
            "missing expected part in output zip: {part_name}"
        );
    }

    for part_name in XML_PARTS {
        let expected = original_parts
            .get(*part_name)
            .unwrap_or_else(|| panic!("missing part in fixture zip: {part_name}"));
        let actual = out_parts
            .get(*part_name)
            .unwrap_or_else(|| panic!("missing part in output zip: {part_name}"));
        assert_xml_bytes_equal_or_semantic(part_name, expected, actual);
    }

    assert_eq!(
        out_parts
            .get("xl/media/image1.png")
            .expect("missing xl/media/image1.png in output zip"),
        original_parts
            .get("xl/media/image1.png")
            .expect("missing xl/media/image1.png in fixture zip"),
        "expected image payload to be preserved byte-for-byte"
    );

    Ok(())
}

fn assert_a1_vm_is_zero(xlsx_bytes: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(xlsx_bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let doc = Document::parse(&sheet_xml)?;
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell in sheet1.xml");
    assert_eq!(
        cell.attribute("vm"),
        Some("0"),
        "expected A1 cell to preserve vm=\"0\""
    );

    Ok(())
}

#[test]
fn roundtrip_preserves_richdata_parts_for_image_in_cell() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/image-in-cell-richdata.xlsx");

    let original_bytes = std::fs::read(&fixture)?;

    let mut doc = formula_xlsx::load_from_path(&fixture)?;
    let out_bytes = doc.save_to_vec()?;
    assert_rich_data_parts_preserved(&original_bytes, &out_bytes)?;
    assert_a1_vm_is_zero(&out_bytes)?;

    // Regression: editing an unrelated cell should not delete or modify richData parts.
    let sheet_id = doc.workbook.sheets[0].id;
    assert!(
        doc.set_cell_value(sheet_id, CellRef::from_a1("B1")?, CellValue::Number(1.0)),
        "expected set_cell_value to succeed"
    );
    let out_bytes = doc.save_to_vec()?;
    assert_rich_data_parts_preserved(&original_bytes, &out_bytes)?;
    assert_a1_vm_is_zero(&out_bytes)?;

    Ok(())
}
