use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches,
    XlsxPackage,
};
use zip::ZipArchive;

fn fixture_bytes() -> Vec<u8> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/rich-data/richdata-minimal.xlsx");
    std::fs::read(&fixture_path).expect("richdata fixture exists")
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn assert_sheet_a1_preserves_vm_and_cm(sheet_xml: &str) {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(
        cell.attribute("vm"),
        Some("1"),
        "expected vm attribute to be preserved (worksheet: {sheet_xml})"
    );
    assert_eq!(
        cell.attribute("cm"),
        Some("1"),
        "expected cm attribute to be preserved (worksheet: {sheet_xml})"
    );
}

fn assert_richdata_content_types_present(content_types_xml: &str) {
    for part in [
        r#"PartName="/xl/metadata.xml""#,
        r#"PartName="/xl/richData/richValue.xml""#,
        r#"PartName="/xl/richData/richValueRel.xml""#,
        r#"PartName="/xl/richData/richValueTypes.xml""#,
        r#"PartName="/xl/richData/richValueStructure.xml""#,
    ] {
        assert!(
            content_types_xml.contains(part),
            "expected [Content_Types].xml to include {part}, got: {content_types_xml}"
        );
    }
    assert!(
        content_types_xml.contains(r#"Extension="png""#),
        "expected [Content_Types].xml to include a png Default entry, got: {content_types_xml}"
    );
}

#[test]
fn streaming_patcher_preserves_richdata_parts_and_metadata_rels() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = fixture_bytes();

    let expected_metadata_xml = zip_part(&bytes, "xl/metadata.xml");
    let expected_metadata_rels = zip_part(&bytes, "xl/_rels/metadata.xml.rels");
    let expected_rich_value = zip_part(&bytes, "xl/richData/richValue.xml");
    let expected_rich_value_rel = zip_part(&bytes, "xl/richData/richValueRel.xml");
    let expected_rich_value_types = zip_part(&bytes, "xl/richData/richValueTypes.xml");
    let expected_rich_value_structure = zip_part(&bytes, "xl/richData/richValueStructure.xml");
    let expected_rich_value_rel_rels = zip_part(&bytes, "xl/richData/_rels/richValueRel.xml.rels");
    let expected_image = zip_part(&bytes, "xl/media/image1.png");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        // Include a material formula so the recalc-policy path executes and rewrites
        // workbook.xml + workbook.xml.rels + [Content_Types].xml.
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;
    let out_bytes = out.into_inner();

    assert_eq!(
        zip_part(&out_bytes, "xl/metadata.xml"),
        expected_metadata_xml,
        "expected xl/metadata.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/_rels/metadata.xml.rels"),
        expected_metadata_rels,
        "expected xl/_rels/metadata.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/richData/richValue.xml"),
        expected_rich_value,
        "expected xl/richData/richValue.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/richData/richValueRel.xml"),
        expected_rich_value_rel,
        "expected xl/richData/richValueRel.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/richData/richValueTypes.xml"),
        expected_rich_value_types,
        "expected xl/richData/richValueTypes.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/richData/richValueStructure.xml"),
        expected_rich_value_structure,
        "expected xl/richData/richValueStructure.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/richData/_rels/richValueRel.xml.rels"),
        expected_rich_value_rel_rels,
        "expected xl/richData/_rels/richValueRel.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&out_bytes, "xl/media/image1.png"),
        expected_image,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    let sheet_xml = String::from_utf8(zip_part(&out_bytes, "xl/worksheets/sheet1.xml"))?;
    assert_sheet_a1_preserves_vm_and_cm(&sheet_xml);

    let workbook_xml = String::from_utf8(zip_part(&out_bytes, "xl/workbook.xml"))?;
    assert!(
        workbook_xml.contains("fullCalcOnLoad=\"1\""),
        "expected workbook.xml to be rewritten with fullCalcOnLoad=1 after formula edits, got: {workbook_xml}"
    );

    let content_types = String::from_utf8(zip_part(&out_bytes, "[Content_Types].xml"))?;
    assert_richdata_content_types_present(&content_types);

    Ok(())
}

#[test]
fn package_patcher_preserves_richdata_parts_and_metadata_rels() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = fixture_bytes();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let expected_metadata_xml = pkg.part("xl/metadata.xml").unwrap().to_vec();
    let expected_metadata_rels = pkg.part("xl/_rels/metadata.xml.rels").unwrap().to_vec();
    let expected_rich_value = pkg.part("xl/richData/richValue.xml").unwrap().to_vec();
    let expected_rich_value_rel = pkg.part("xl/richData/richValueRel.xml").unwrap().to_vec();
    let expected_rich_value_types = pkg.part("xl/richData/richValueTypes.xml").unwrap().to_vec();
    let expected_rich_value_structure = pkg.part("xl/richData/richValueStructure.xml").unwrap().to_vec();
    let expected_rich_value_rel_rels = pkg
        .part("xl/richData/_rels/richValueRel.xml.rels")
        .unwrap()
        .to_vec();
    let expected_image = pkg.part("xl/media/image1.png").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );
    pkg.apply_cell_patches(&patches)?;

    assert_eq!(
        pkg.part("xl/metadata.xml").unwrap(),
        expected_metadata_xml.as_slice(),
        "expected xl/metadata.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/_rels/metadata.xml.rels").unwrap(),
        expected_metadata_rels.as_slice(),
        "expected xl/_rels/metadata.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/richData/richValue.xml").unwrap(),
        expected_rich_value.as_slice(),
        "expected xl/richData/richValue.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/richData/richValueRel.xml").unwrap(),
        expected_rich_value_rel.as_slice(),
        "expected xl/richData/richValueRel.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/richData/richValueTypes.xml").unwrap(),
        expected_rich_value_types.as_slice(),
        "expected xl/richData/richValueTypes.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/richData/richValueStructure.xml").unwrap(),
        expected_rich_value_structure.as_slice(),
        "expected xl/richData/richValueStructure.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/richData/_rels/richValueRel.xml.rels").unwrap(),
        expected_rich_value_rel_rels.as_slice(),
        "expected xl/richData/_rels/richValueRel.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/media/image1.png").unwrap(),
        expected_image.as_slice(),
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_sheet_a1_preserves_vm_and_cm(sheet_xml);

    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap())?;
    assert!(
        workbook_xml.contains("fullCalcOnLoad=\"1\""),
        "expected workbook.xml to be rewritten with fullCalcOnLoad=1 after formula edits, got: {workbook_xml}"
    );

    let content_types = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())?;
    assert_richdata_content_types_present(content_types);

    Ok(())
}

#[test]
fn document_roundtrip_preserves_richdata_parts_and_metadata_rels() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = fixture_bytes();

    let expected_metadata_xml = zip_part(&bytes, "xl/metadata.xml");
    let expected_metadata_rels = zip_part(&bytes, "xl/_rels/metadata.xml.rels");
    let expected_rich_value = zip_part(&bytes, "xl/richData/richValue.xml");
    let expected_rich_value_rel = zip_part(&bytes, "xl/richData/richValueRel.xml");
    let expected_rich_value_types = zip_part(&bytes, "xl/richData/richValueTypes.xml");
    let expected_rich_value_structure = zip_part(&bytes, "xl/richData/richValueStructure.xml");
    let expected_rich_value_rel_rels = zip_part(&bytes, "xl/richData/_rels/richValueRel.xml.rels");
    let expected_image = zip_part(&bytes, "xl/media/image1.png");

    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 exists")
        .id;

    assert!(
        doc.set_cell_value(sheet_id, CellRef::from_a1("A1")?, CellValue::Number(2.0)),
        "expected set_cell_value to succeed"
    );
    assert!(
        doc.set_cell_formula(sheet_id, CellRef::from_a1("A1")?, Some("=1+1".to_string())),
        "expected set_cell_formula to succeed"
    );

    let saved = doc.save_to_vec()?;

    assert_eq!(
        zip_part(&saved, "xl/metadata.xml"),
        expected_metadata_xml,
        "expected xl/metadata.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/_rels/metadata.xml.rels"),
        expected_metadata_rels,
        "expected xl/_rels/metadata.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValue.xml"),
        expected_rich_value,
        "expected xl/richData/richValue.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueRel.xml"),
        expected_rich_value_rel,
        "expected xl/richData/richValueRel.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueTypes.xml"),
        expected_rich_value_types,
        "expected xl/richData/richValueTypes.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueStructure.xml"),
        expected_rich_value_structure,
        "expected xl/richData/richValueStructure.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/_rels/richValueRel.xml.rels"),
        expected_rich_value_rel_rels,
        "expected xl/richData/_rels/richValueRel.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/media/image1.png"),
        expected_image,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    let sheet_xml = String::from_utf8(zip_part(&saved, "xl/worksheets/sheet1.xml"))?;
    assert_sheet_a1_preserves_vm_and_cm(&sheet_xml);

    Ok(())
}
