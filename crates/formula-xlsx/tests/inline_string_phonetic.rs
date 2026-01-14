use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming, CellPatch, WorkbookCellPatches, WorksheetCellPatch,
    XlsxPackage,
};
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::{CompressionMethod, ZipWriter};

const PHONETIC_TEXT: &str = "PHONETIC";

fn build_xlsx_with_worksheet_xml(worksheet_xml: &str) -> Vec<u8> {
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

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_inline_string_phonetic_fixture_xlsx() -> Vec<u8> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Base</t>
          <phoneticPr fontId="0" type="noConversion"/>
          <rPh sb="0" eb="2"><t>PHO</t></rPh>
          <rPh sb="2" eb="4"><t>NETIC</t></rPh>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
    build_xlsx_with_worksheet_xml(worksheet_xml)
}

fn build_inline_rich_string_phonetic_fixture_xlsx() -> Vec<u8> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r><t>Ba</t></r>
          <r><t>se</t></r>
          <phoneticPr fontId="0" type="noConversion"/>
          <rPh sb="0" eb="2"><t>PHO</t></rPh>
          <rPh sb="2" eb="4"><t>NETIC</t></rPh>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;
    build_xlsx_with_worksheet_xml(worksheet_xml)
}

#[test]
fn load_inline_string_imports_phonetic_text() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_inline_string_phonetic_fixture_xlsx();
    let doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");
    let cell_ref = CellRef::from_a1("A1")?;
    assert_eq!(
        sheet.value(cell_ref),
        CellValue::String("Base".to_string())
    );
    let cell = sheet.cell(cell_ref).expect("cell exists");
    assert_eq!(
        cell.phonetic_text(),
        Some(PHONETIC_TEXT),
        "expected inline string phonetic guide text to be imported into the model cell"
    );

    Ok(())
}

#[test]
fn streaming_patch_preserves_inline_string_phonetic_subtree_on_style_only_patch(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_inline_string_phonetic_fixture_xlsx();

    // Apply a style-only patch (value is unchanged) to force rewriting the `<c>` start tag.
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::String("Base".to_string()),
        None,
    )
    .with_xf_index(Some(1));

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;
    let out_bytes = out.into_inner();

    let mut archive = ZipArchive::new(Cursor::new(out_bytes.clone()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    assert!(
        sheet_xml.contains(r#"s="1""#),
        "expected patched cell to contain s=\"1\" style attribute:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<phoneticPr"),
        "expected worksheet XML to preserve <phoneticPr> subtree:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<rPh"),
        "expected worksheet XML to preserve <rPh> subtree:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("PHO"),
        "expected worksheet XML to preserve phonetic marker text:\n{sheet_xml}"
    );

    let doc = load_from_bytes(&out_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1")?),
        CellValue::String("Base".to_string())
    );
    let cell = sheet.cell(CellRef::from_a1("A1")?).expect("cell exists");
    assert_eq!(
        cell.phonetic_text(),
        Some(PHONETIC_TEXT),
        "expected inline string phonetic guide text to survive a style-only streaming patch"
    );

    Ok(())
}

#[test]
fn streaming_patch_preserves_inline_rich_string_phonetic_subtree_on_style_only_patch(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_inline_rich_string_phonetic_fixture_xlsx();

    // Apply a style-only patch (value is unchanged) to force rewriting the `<c>` start tag.
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::String("Base".to_string()),
        None,
    )
    .with_xf_index(Some(1));

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;
    let out_bytes = out.into_inner();

    let mut archive = ZipArchive::new(Cursor::new(out_bytes.clone()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    assert!(
        sheet_xml.contains(r#"s="1""#),
        "expected patched cell to contain s=\"1\" style attribute:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<phoneticPr"),
        "expected worksheet XML to preserve <phoneticPr> subtree:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<rPh"),
        "expected worksheet XML to preserve <rPh> subtree:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("PHO"),
        "expected worksheet XML to preserve phonetic marker text:\n{sheet_xml}"
    );

    let doc = load_from_bytes(&out_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1")?),
        CellValue::String("Base".to_string())
    );
    let cell = sheet.cell(CellRef::from_a1("A1")?).expect("cell exists");
    assert_eq!(
        cell.phonetic_text(),
        Some(PHONETIC_TEXT),
        "expected inline string phonetic guide text to survive a style-only streaming patch"
    );

    Ok(())
}

#[test]
fn in_memory_patch_preserves_inline_string_phonetic_subtree_on_style_only_patch(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_inline_string_phonetic_fixture_xlsx();

    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_style(CellValue::String("Base".to_string()), 1),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"s="1""#),
        "expected patched cell to contain s=\"1\" style attribute:\n{out_xml}"
    );
    assert!(
        out_xml.contains("<phoneticPr"),
        "expected worksheet XML to preserve <phoneticPr> subtree:\n{out_xml}"
    );
    assert!(
        out_xml.contains("<rPh"),
        "expected worksheet XML to preserve <rPh> subtree:\n{out_xml}"
    );
    assert!(
        out_xml.contains("PHO"),
        "expected worksheet XML to preserve phonetic marker text:\n{out_xml}"
    );

    let out_bytes = pkg.write_to_bytes()?;
    let doc = load_from_bytes(&out_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");
    let cell_ref = CellRef::from_a1("A1")?;
    assert_eq!(sheet.value(cell_ref), CellValue::String("Base".to_string()));
    let cell = sheet.cell(cell_ref).expect("cell exists");
    assert_eq!(cell.phonetic.as_deref(), Some(PHONETIC_TEXT));

    Ok(())
}
