use std::collections::BTreeSet;
use std::io::{Cursor, Read};
use std::path::Path;

use zip::ZipArchive;

fn zip_entry_names(zip_bytes: &[u8]) -> Result<Vec<String>, zip::result::ZipError> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor)?;
    let mut out = Vec::with_capacity(archive.len());
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        out.push(file.name().to_string());
    }
    Ok(out)
}

fn zip_part_bytes(zip_bytes: &[u8], name: &str) -> Result<Vec<u8>, zip::result::ZipError> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor)?;
    let mut file = archive.by_name(name)?;
    let mut buf = Vec::with_capacity(file.size() as usize);
    file.read_to_end(&mut buf)?;
    Ok(buf)
}

#[test]
fn linked_data_types_fixture_roundtrips_richdata_parts() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/richdata/linked-data-types.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let names = zip_entry_names(&fixture_bytes)?;
    let name_set: BTreeSet<&str> = names.iter().map(|s| s.as_str()).collect();

    assert!(
        name_set.contains("xl/metadata.xml"),
        "fixture missing xl/metadata.xml; zip entries: {names:?}"
    );
    assert!(
        name_set.contains("xl/_rels/metadata.xml.rels"),
        "fixture missing xl/_rels/metadata.xml.rels; zip entries: {names:?}"
    );

    let richdata_parts: Vec<String> = names
        .iter()
        .filter(|n| n.starts_with("xl/richData/"))
        .cloned()
        .collect();
    assert!(
        !richdata_parts.is_empty(),
        "fixture missing xl/richData/* parts; zip entries: {names:?}"
    );

    // Parse the worksheet XML and ensure at least one cell has vm/cm attributes.
    let sheet_xml_bytes = zip_part_bytes(&fixture_bytes, "xl/worksheets/sheet1.xml")?;
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;
    let has_vm_or_cm = parsed.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "c"
            && (n.attribute("vm").is_some() || n.attribute("cm").is_some())
    });
    assert!(
        has_vm_or_cm,
        "expected at least one <c> element with vm/cm attributes; sheet1.xml: {sheet_xml}"
    );

    // Round-trip through XlsxDocument and ensure richdata-related parts remain present.
    let doc = formula_xlsx::load_from_bytes(&fixture_bytes)?;
    let saved = doc.save_to_vec()?;

    let saved_names = zip_entry_names(&saved)?;
    let saved_set: BTreeSet<&str> = saved_names.iter().map(|s| s.as_str()).collect();

    let mut expected_parts: BTreeSet<String> = BTreeSet::new();
    expected_parts.insert("xl/metadata.xml".to_string());
    expected_parts.insert("xl/_rels/metadata.xml.rels".to_string());
    expected_parts.extend(richdata_parts.iter().cloned());

    for part in expected_parts {
        assert!(
            saved_set.contains(part.as_str()),
            "round-tripped workbook missing {part}; richData parts in fixture: {richdata_parts:?}; saved zip entries: {saved_names:?}"
        );
    }

    Ok(())
}

