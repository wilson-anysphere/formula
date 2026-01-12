use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, parse_value_metadata_vm_to_rich_value_index_map};
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

fn zip_part_names(zip_bytes: &[u8]) -> Vec<String> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut out = Vec::new();
    for i in 0..archive.len() {
        let file = archive.by_index(i).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        out.push(file.name().to_string());
    }
    out
}

#[test]
fn real_excel_linked_data_types_roundtrip_preserves_richdata_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    // This fixture is expected to be a modern Excel workbook that stores linked data types via
    // `xl/metadata.xml` + `xl/richData/*` and uses worksheet cell `vm=` / `cm=` attributes.
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/rich-data/linked-data-types-excel.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let app_xml = String::from_utf8(zip_part(&fixture_bytes, "docProps/app.xml"))?;
    assert!(
        app_xml.contains("<Application>Microsoft Excel</Application>"),
        "expected docProps/app.xml to indicate the workbook was saved by Excel, got: {app_xml}"
    );

    for part in [
        "xl/metadata.xml",
        "xl/_rels/metadata.xml.rels",
        "xl/richData/richValue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/richValueTypes.xml",
        "xl/richData/richValueStructure.xml",
        "xl/worksheets/sheet1.xml",
    ] {
        assert!(
            zip_part_exists(&fixture_bytes, part),
            "expected fixture to contain {part}"
        );
    }

    // Confirm at least one worksheet cell has vm/c m and it resolves to a rich value index via
    // xl/metadata.xml.
    let sheet_xml = String::from_utf8(zip_part(&fixture_bytes, "xl/worksheets/sheet1.xml"))?;
    let parsed = roxmltree::Document::parse(&sheet_xml)?;
    let cell_a1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(
        cell_a1.attribute("vm"),
        Some("1"),
        "expected Sheet1!A1 to have vm=\"1\", got: {sheet_xml}"
    );
    assert_eq!(
        cell_a1.attribute("cm"),
        Some("1"),
        "expected Sheet1!A1 to have cm=\"1\", got: {sheet_xml}"
    );

    let metadata_bytes = zip_part(&fixture_bytes, "xl/metadata.xml");
    let vm_map = parse_value_metadata_vm_to_rich_value_index_map(&metadata_bytes)?;
    assert_eq!(
        vm_map.get(&1),
        Some(&0),
        "expected vm=1 to resolve to rich value index 0 via xl/metadata.xml"
    );
    assert_eq!(
        vm_map.get(&2),
        Some(&1),
        "expected vm=2 to resolve to rich value index 1 via xl/metadata.xml"
    );

    // Capture original bytes for all rich-data parts we need to preserve byte-for-byte.
    let mut rich_part_names: Vec<String> = zip_part_names(&fixture_bytes)
        .into_iter()
        .filter(|name| {
            name == "xl/metadata.xml"
                || name == "xl/_rels/metadata.xml.rels"
                || name.starts_with("xl/richData/")
        })
        .collect();
    rich_part_names.sort();
    assert!(
        rich_part_names.iter().any(|p| p.starts_with("xl/richData/")),
        "expected fixture to contain at least one xl/richData/* part"
    );

    let original_parts: Vec<(String, Vec<u8>)> = rich_part_names
        .iter()
        .map(|name| (name.clone(), zip_part(&fixture_bytes, name)))
        .collect();

    let mut doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheet_by_name("Sheet1").expect("Sheet1").id;

    // Task 483 baseline: rich values should decode into entity values in-memory.
    let sheet = doc.workbook.sheet(sheet_id).expect("Sheet1");
    let a1 = sheet.cell(CellRef::from_a1("A1")?).expect("A1 cell");
    let a2 = sheet.cell(CellRef::from_a1("A2")?).expect("A2 cell");
    match &a1.value {
        CellValue::Entity(entity) => assert_eq!(entity.display_value, "MSFT"),
        other => panic!("expected Sheet1!A1 to decode to CellValue::Entity, got: {other:?}"),
    }
    match &a2.value {
        CellValue::Entity(entity) => assert_eq!(entity.display_value, "Seattle"),
        other => panic!("expected Sheet1!A2 to decode to CellValue::Entity, got: {other:?}"),
    }

    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A1")?),
        Some(0),
        "expected rich value index 0 for Sheet1!A1"
    );
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A2")?),
        Some(1),
        "expected rich value index 1 for Sheet1!A2"
    );

    // Edit an unrelated cell to exercise worksheet patching while preserving rich-data parts.
    assert!(doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("B1")?,
        CellValue::Number(42.0)
    ));
    let saved = doc.save_to_vec()?;

    for (name, original) in original_parts {
        assert_eq!(
            zip_part(&saved, &name),
            original,
            "expected {name} to be preserved byte-for-byte"
        );
    }

    // Ensure vm/cm attributes on the linked data type cells are still present after editing B1.
    let saved_sheet_xml = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml"))?;
    let parsed = roxmltree::Document::parse(&saved_sheet_xml)?;
    for (coord, expected_vm, expected_cm) in [("A1", "1", "1"), ("A2", "2", "2")] {
        let cell = parsed
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(coord))
            .unwrap_or_else(|| panic!("expected {coord} cell"));
        assert_eq!(
            cell.attribute("vm"),
            Some(expected_vm),
            "expected {coord} vm attribute to be preserved, got: {saved_sheet_xml}"
        );
        assert_eq!(
            cell.attribute("cm"),
            Some(expected_cm),
            "expected {coord} cm attribute to be preserved, got: {saved_sheet_xml}"
        );
    }

    Ok(())
}
