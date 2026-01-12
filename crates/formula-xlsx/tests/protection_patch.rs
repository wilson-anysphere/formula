use std::io::{Cursor, Read, Write};

use formula_model::SheetProtection;
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};
use zip::ZipArchive;

fn build_minimal_protected_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:workbookProtection lockStructure="1" workbookPassword="83AF"/>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
  <x:sheetProtection sheet="1" password="CBEB"/>
</x:worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_part(bytes: &[u8], part: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut text = String::new();
    archive.by_name(part)?.read_to_string(&mut text)?;
    Ok(text)
}

#[test]
fn patches_workbook_and_sheet_protection_on_save() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_protected_xlsx();
    let mut doc = load_from_bytes(&bytes)?;

    // Patch workbook protection.
    doc.workbook.workbook_protection.lock_structure = false;
    doc.workbook.workbook_protection.lock_windows = true;
    doc.workbook.workbook_protection.password_hash = Some(0x1234);

    // Disable sheet protection (removes the `<sheetProtection>` element).
    doc.workbook.sheets[0].sheet_protection = SheetProtection::default();

    let expected_workbook_protection = doc.workbook.workbook_protection.clone();
    let expected_sheet_protection = doc.workbook.sheets[0].sheet_protection.clone();

    let out = doc.save_to_vec()?;
    let roundtrip = read_workbook_model_from_bytes(&out)?;

    assert_eq!(roundtrip.workbook_protection, expected_workbook_protection);
    assert_eq!(roundtrip.sheets[0].sheet_protection, expected_sheet_protection);

    // Ensure `sheetProtection` is removed in the written XML when disabled.
    let sheet_xml = read_part(&out, "xl/worksheets/sheet1.xml")?;
    assert!(
        !sheet_xml.contains("sheetProtection"),
        "expected `<sheetProtection>` removal, got:\n{sheet_xml}"
    );

    // Workbook uses a prefix-only SpreadsheetML namespace; ensure the patch keeps the prefix.
    let workbook_xml = read_part(&out, "xl/workbook.xml")?;
    roxmltree::Document::parse(&workbook_xml)?;
    assert!(
        workbook_xml.contains("<x:workbookProtection"),
        "expected prefixed `<x:workbookProtection>` after patching, got:\n{workbook_xml}"
    );

    Ok(())
}

