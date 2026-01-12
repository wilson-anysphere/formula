use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::{CellRef, CellValue, ErrorValue};
use formula_xlsx::{
    patch_xlsx_streaming, patch_xlsx_streaming_with_recalc_policy, RecalcPolicy, WorksheetCellPatch,
};
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::{CompressionMethod, ZipWriter};

fn read_worksheet_xml(bytes: &[u8], part_name: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut xml = String::new();
    archive.by_name(part_name)?.read_to_string(&mut xml)?;
    Ok(xml)
}

fn build_minimal_xlsx_with_vm_error_cell() -> Vec<u8> {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

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

    // Simulate an images-in-cell / rich value placeholder cell.
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e" vm="1" cm="2"><f>1+1</f><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types_xml.as_bytes()).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(sheet1_xml.as_bytes()).unwrap();

        zip.finish().unwrap();
    }
    cursor.into_inner()
}

#[test]
fn streaming_recalc_policy_can_clear_cached_values_for_formula_patches(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/rt_simple.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("C1")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );

    // When enabled, a formula patch should omit the cached `<v>` value.
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_with_recalc_policy(
        Cursor::new(bytes.clone()),
        &mut out,
        &[patch.clone()],
        RecalcPolicy {
            clear_cached_values_on_formula_change: true,
            ..Default::default()
        },
    )?;
    let sheet_xml = read_worksheet_xml(out.get_ref(), "xl/worksheets/sheet1.xml")?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("patched cell C1 should exist");
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "f"),
        "patched cell should contain a <f> element"
    );
    assert!(
        !patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "cached value should be omitted when clear_cached_values_on_formula_change is enabled"
    );

    // Default behavior should continue writing cached `<v>` values.
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;
    let sheet_xml = read_worksheet_xml(out.get_ref(), "xl/worksheets/sheet1.xml")?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("patched cell C1 should exist");
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "f"),
        "patched cell should contain a <f> element"
    );
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "default streaming patcher should preserve cached <v> values"
    );

    Ok(())
}

#[test]
fn streaming_recalc_policy_preserves_vm_for_placeholder_error_when_clearing_cached_values(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_xlsx_with_vm_error_cell();

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Error(ErrorValue::Value),
        Some("=1+1".to_string()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_with_recalc_policy(
        Cursor::new(bytes),
        &mut out,
        &[patch],
        RecalcPolicy {
            clear_cached_values_on_formula_change: true,
            ..Default::default()
        },
    )?;

    let sheet_xml = read_worksheet_xml(out.get_ref(), "xl/worksheets/sheet1.xml")?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("patched cell A1 should exist");

    assert_eq!(
        patched_cell.attribute("vm"),
        Some("1"),
        "expected vm to be preserved when patching to the rich-value placeholder error, got: {sheet_xml}"
    );
    assert_eq!(patched_cell.attribute("cm"), Some("2"));
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "f"),
        "patched cell should contain a <f> element"
    );
    assert!(
        !patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "cached value should be omitted when clear_cached_values_on_formula_change is enabled"
    );

    Ok(())
}
