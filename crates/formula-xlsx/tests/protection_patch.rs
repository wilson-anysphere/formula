use std::io::{Cursor, Read, Write};

use formula_model::{SheetProtection, WorkbookProtection};
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};
use zip::ZipArchive;

fn build_minimal_xlsx(workbook_xml: &str, sheet_xml: &str) -> Vec<u8> {
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

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

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
  <x:sheetProtection sheet="1" password="CBEB"/>
</x:worksheet>"#;

    build_minimal_xlsx(workbook_xml, sheet_xml)
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

#[test]
fn removes_workbook_protection_when_cleared() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_protected_xlsx();
    let mut doc = load_from_bytes(&bytes)?;

    doc.workbook.workbook_protection = WorkbookProtection::default();

    let out = doc.save_to_vec()?;
    let roundtrip = read_workbook_model_from_bytes(&out)?;
    assert_eq!(roundtrip.workbook_protection, WorkbookProtection::default());

    let workbook_xml = read_part(&out, "xl/workbook.xml")?;
    roxmltree::Document::parse(&workbook_xml)?;
    assert!(
        !workbook_xml.contains("workbookProtection"),
        "expected `<workbookProtection>` removal, got:\n{workbook_xml}"
    );

    Ok(())
}

#[test]
fn inserts_workbook_and_sheet_protection_when_missing() -> Result<(), Box<dyn std::error::Error>> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
  <x:autoFilter ref="A1:A1"/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    doc.workbook.workbook_protection.lock_structure = true;
    doc.workbook.workbook_protection.password_hash = Some(0x1234);

    let sheet = &mut doc.workbook.sheets[0];
    sheet.sheet_protection.enabled = true;
    sheet.sheet_protection.password_hash = Some(0xABCD);
    sheet.sheet_protection.select_locked_cells = false;

    let expected_workbook_protection = doc.workbook.workbook_protection.clone();
    let expected_sheet_protection = doc.workbook.sheets[0].sheet_protection.clone();

    let out = doc.save_to_vec()?;
    let roundtrip = read_workbook_model_from_bytes(&out)?;

    assert_eq!(roundtrip.workbook_protection, expected_workbook_protection);
    assert_eq!(roundtrip.sheets[0].sheet_protection, expected_sheet_protection);

    let workbook_xml = read_part(&out, "xl/workbook.xml")?;
    roxmltree::Document::parse(&workbook_xml)?;
    let wb_prot = workbook_xml
        .find("<x:workbookProtection")
        .expect("workbookProtection inserted");
    let sheets = workbook_xml.find("<x:sheets").expect("sheets exists");
    assert!(wb_prot < sheets, "expected workbookProtection before sheets");

    let sheet_xml = read_part(&out, "xl/worksheets/sheet1.xml")?;
    roxmltree::Document::parse(&sheet_xml)?;
    let sheet_prot = sheet_xml
        .find("<x:sheetProtection")
        .expect("sheetProtection inserted");
    let sheet_data = sheet_xml.find("<x:sheetData").expect("sheetData exists");
    let auto_filter = sheet_xml.find("<x:autoFilter").expect("autoFilter exists");
    assert!(
        sheet_data < sheet_prot && sheet_prot < auto_filter,
        "expected sheetProtection after sheetData and before autoFilter, got:\n{sheet_xml}"
    );

    Ok(())
}

#[test]
fn inserts_workbook_and_sheet_protection_when_missing_with_default_namespace(
) -> Result<(), Box<dyn std::error::Error>> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <autoFilter ref="A1:A1"/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    doc.workbook.workbook_protection.lock_structure = true;
    doc.workbook.workbook_protection.password_hash = Some(0x1234);

    let sheet = &mut doc.workbook.sheets[0];
    sheet.sheet_protection.enabled = true;
    sheet.sheet_protection.password_hash = Some(0xABCD);
    sheet.sheet_protection.select_locked_cells = false;

    let expected_workbook_protection = doc.workbook.workbook_protection.clone();
    let expected_sheet_protection = doc.workbook.sheets[0].sheet_protection.clone();

    let out = doc.save_to_vec()?;
    let roundtrip = read_workbook_model_from_bytes(&out)?;

    assert_eq!(roundtrip.workbook_protection, expected_workbook_protection);
    assert_eq!(roundtrip.sheets[0].sheet_protection, expected_sheet_protection);

    let workbook_xml = read_part(&out, "xl/workbook.xml")?;
    roxmltree::Document::parse(&workbook_xml)?;
    assert!(
        workbook_xml.contains("<workbookProtection"),
        "expected unprefixed `<workbookProtection>` insertion, got:\n{workbook_xml}"
    );
    assert!(
        !workbook_xml.contains("<x:workbookProtection"),
        "should not introduce a SpreadsheetML prefix in a default-namespace workbook, got:\n{workbook_xml}"
    );

    let wb_prot = workbook_xml
        .find("<workbookProtection")
        .expect("workbookProtection inserted");
    let sheets = workbook_xml.find("<sheets").expect("sheets exists");
    assert!(wb_prot < sheets, "expected workbookProtection before sheets");

    let sheet_xml = read_part(&out, "xl/worksheets/sheet1.xml")?;
    roxmltree::Document::parse(&sheet_xml)?;
    assert!(
        sheet_xml.contains("<sheetProtection"),
        "expected unprefixed `<sheetProtection>` insertion, got:\n{sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("<x:sheetProtection"),
        "should not introduce a SpreadsheetML prefix in a default-namespace worksheet, got:\n{sheet_xml}"
    );

    let sheet_prot = sheet_xml
        .find("<sheetProtection")
        .expect("sheetProtection inserted");
    let sheet_data = sheet_xml.find("<sheetData").expect("sheetData exists");
    let auto_filter = sheet_xml.find("<autoFilter").expect("autoFilter exists");
    assert!(
        sheet_data < sheet_prot && sheet_prot < auto_filter,
        "expected sheetProtection after sheetData and before autoFilter, got:\n{sheet_xml}"
    );

    Ok(())
}

#[test]
fn preserves_unmodeled_workbook_protection_attributes() -> Result<(), Box<dyn std::error::Error>> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <workbookPr/>
  <workbookProtection lockStructure="1" workbookPassword="83AF" x14ac:dummy="1"/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    doc.workbook.workbook_protection.lock_structure = false;
    doc.workbook.workbook_protection.lock_windows = true;
    doc.workbook.workbook_protection.password_hash = Some(0x1234);

    let expected = doc.workbook.workbook_protection.clone();
    let out = doc.save_to_vec()?;

    let roundtrip = read_workbook_model_from_bytes(&out)?;
    assert_eq!(roundtrip.workbook_protection, expected);

    let workbook_xml = read_part(&out, "xl/workbook.xml")?;
    roxmltree::Document::parse(&workbook_xml)?;
    assert!(
        workbook_xml.contains(r#"x14ac:dummy="1""#),
        "expected unmodeled workbookProtection attributes to be preserved, got:\n{workbook_xml}"
    );

    Ok(())
}

#[test]
fn patches_non_empty_protection_elements() -> Result<(), Box<dyn std::error::Error>> {
    // Some producers write protection elements as explicit start/end tags instead of self-closing.
    // Ensure our patching logic still replaces them correctly.
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:workbookProtection lockStructure="1" workbookPassword="83AF"></x:workbookProtection>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData>
    <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>
  </x:sheetData>
  <x:sheetProtection sheet="1" password="CBEB"></x:sheetProtection>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    doc.workbook.workbook_protection.lock_structure = false;
    doc.workbook.workbook_protection.lock_windows = true;
    doc.workbook.workbook_protection.password_hash = Some(0x1234);

    let sheet = &mut doc.workbook.sheets[0];
    sheet.sheet_protection.enabled = true;
    sheet.sheet_protection.password_hash = Some(0x5678);
    sheet.sheet_protection.select_locked_cells = false;
    sheet.sheet_protection.format_cells = true;

    let expected_workbook_protection = doc.workbook.workbook_protection.clone();
    let expected_sheet_protection = doc.workbook.sheets[0].sheet_protection.clone();

    let out = doc.save_to_vec()?;
    let roundtrip = read_workbook_model_from_bytes(&out)?;

    assert_eq!(roundtrip.workbook_protection, expected_workbook_protection);
    assert_eq!(roundtrip.sheets[0].sheet_protection, expected_sheet_protection);

    Ok(())
}

#[test]
fn removes_sheet_protection_when_present_but_disabled() -> Result<(), Box<dyn std::error::Error>> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // Some writers store a disabled protection element as `<sheetProtection sheet="0"/>`.
    // We canonicalize disabled protection by removing the element entirely.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <sheetProtection sheet="0"/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);
    let doc = load_from_bytes(&bytes)?;
    assert_eq!(doc.workbook.sheets[0].sheet_protection, SheetProtection::default());

    let out = doc.save_to_vec()?;
    let roundtrip = read_workbook_model_from_bytes(&out)?;
    assert_eq!(roundtrip.sheets[0].sheet_protection, SheetProtection::default());

    let sheet_xml = read_part(&out, "xl/worksheets/sheet1.xml")?;
    roxmltree::Document::parse(&sheet_xml)?;
    assert!(
        !sheet_xml.contains("sheetProtection"),
        "expected `<sheetProtection>` removal when disabled, got:\n{sheet_xml}"
    );

    Ok(())
}

#[test]
fn inserts_sheet_protection_after_non_empty_sheet_data() -> Result<(), Box<dyn std::error::Error>> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData>
    <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>
  </x:sheetData>
  <x:autoFilter ref="A1:A1"/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(workbook_xml, sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    let sheet = &mut doc.workbook.sheets[0];
    sheet.sheet_protection.enabled = true;
    sheet.sheet_protection.password_hash = Some(0xABCD);

    let out = doc.save_to_vec()?;

    let sheet_xml = read_part(&out, "xl/worksheets/sheet1.xml")?;
    roxmltree::Document::parse(&sheet_xml)?;

    let sheet_data_end = sheet_xml
        .find("</x:sheetData>")
        .expect("expected closing sheetData tag");
    let sheet_prot = sheet_xml
        .find("<x:sheetProtection")
        .expect("sheetProtection inserted");
    let auto_filter = sheet_xml.find("<x:autoFilter").expect("autoFilter exists");

    assert!(
        sheet_data_end < sheet_prot && sheet_prot < auto_filter,
        "expected sheetProtection after sheetData and before autoFilter, got:\n{sheet_xml}"
    );

    Ok(())
}
