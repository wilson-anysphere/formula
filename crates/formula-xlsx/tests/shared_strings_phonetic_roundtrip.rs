use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, CellPatch, WorkbookCellPatches, XlsxPackage};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const PHONETIC_MARKER: &str = "PHO_MARKER_123";
const EXTLST_MARKER: &str = "EXT_MARKER_456";

fn build_fixture_xlsx() -> Vec<u8> {
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let shared_strings_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Base</t>
    <rPh sb="0" eb="4"><t>{PHONETIC_MARKER}</t></rPh>
  </si>
  <extLst>
    <ext uri="{{{EXTLST_MARKER}}}">
      <marker>{EXTLST_MARKER}</marker>
    </ext>
  </extLst>
</sst>"#
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all(shared_strings_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = String::new();
    file.read_to_string(&mut buf).expect("read part");
    buf
}

#[test]
fn shared_strings_roundtrip_preserves_phonetic_and_unknown_subtrees() {
    let bytes = build_fixture_xlsx();
    let doc = load_from_bytes(&bytes).expect("load fixture");

    // Ensure the phonetic `<t>` does not pollute the visible string value.
    let sheet = doc.workbook.sheets.first().expect("sheet exists");
    let value = sheet.value_a1("A1").expect("A1");
    assert_eq!(value, CellValue::String("Base".to_string()));

    let saved = doc.save_to_vec().expect("save");
    let ss_xml = zip_part(&saved, "xl/sharedStrings.xml");
    assert!(
        ss_xml.contains(PHONETIC_MARKER),
        "expected sharedStrings.xml to preserve phonetic subtree"
    );
    assert!(
        ss_xml.contains(EXTLST_MARKER),
        "expected sharedStrings.xml to preserve <extLst> subtree"
    );
}

#[test]
fn patch_pipeline_preserves_unknown_shared_strings_and_appends_new_entry() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_fixture_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellPatch::set_value(CellValue::String("Patched".to_string())),
    );
    pkg.apply_cell_patches(&patches)?;

    let saved = pkg.write_to_bytes()?;
    let ss_xml = zip_part(&saved, "xl/sharedStrings.xml");
    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");

    assert!(
        ss_xml.contains(PHONETIC_MARKER),
        "expected patched sharedStrings.xml to preserve phonetic subtree"
    );
    assert!(
        ss_xml.contains(EXTLST_MARKER),
        "expected patched sharedStrings.xml to preserve <extLst> subtree"
    );
    assert!(
        ss_xml.contains("Patched"),
        "expected patched sharedStrings.xml to include newly inserted string"
    );
    assert!(
        sheet_xml.contains(r#"<c r="A2" t="s""#) || sheet_xml.contains(r#"r="A2" t="s""#),
        "expected patched cell A2 to use shared strings, got: {sheet_xml}"
    );

    Ok(())
}
