use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn build_richdata_xlsx() -> Vec<u8> {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"></Override>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.richValueTypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.richValues+xml"/>
</Types>"#;

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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml">
  </Relationship>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>
  </sheetData>
</worksheet>"#;

    // Arbitrary but stable bytes (not necessarily Excel-valid).
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <test>richdata-metadata</test>
</metadata>"#;

    let metadata_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/richValueTypes" Target="richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://example.com/richValues" Target="richData/richValues.xml"/>
</Relationships>"#;

    let rich_value_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://example.com/richValueTypes">
  <a>types</a>
</rvTypes>"#;

    let rich_values_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv xmlns="http://example.com/richValues">
  <a>values</a>
</rv>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types_xml.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(b"<calcChain/>").unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml).unwrap();

    zip.start_file("xl/_rels/metadata.xml.rels", options).unwrap();
    zip.write_all(metadata_rels_xml).unwrap();

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(rich_value_types_xml).unwrap();

    zip.start_file("xl/richData/richValues.xml", options)
        .unwrap();
    zip.write_all(rich_values_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn apply_cell_patches_preserves_richdata_parts_and_relationships() {
    let bytes = build_richdata_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let metadata_before = pkg.part("xl/metadata.xml").unwrap().to_vec();
    let metadata_rels_before = pkg.part("xl/_rels/metadata.xml.rels").unwrap().to_vec();
    let rich_value_types_before = pkg.part("xl/richData/richValueTypes.xml").unwrap().to_vec();
    let rich_values_before = pkg.part("xl/richData/richValues.xml").unwrap().to_vec();

    // Patch a string value, removing the existing formula and triggering workbook-level
    // relationship/content-type rewrites (calcChain drop + full-calc-on-load).
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::String("hi".to_string())),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    assert_eq!(
        pkg.part("xl/metadata.xml").unwrap(),
        metadata_before.as_slice(),
        "expected xl/metadata.xml bytes to be preserved"
    );
    assert_eq!(
        pkg.part("xl/_rels/metadata.xml.rels").unwrap(),
        metadata_rels_before.as_slice(),
        "expected xl/_rels/metadata.xml.rels bytes to be preserved"
    );
    assert_eq!(
        pkg.part("xl/richData/richValueTypes.xml").unwrap(),
        rich_value_types_before.as_slice(),
        "expected xl/richData/richValueTypes.xml bytes to be preserved"
    );
    assert_eq!(
        pkg.part("xl/richData/richValues.xml").unwrap(),
        rich_values_before.as_slice(),
        "expected xl/richData/richValues.xml bytes to be preserved"
    );

    let workbook_rels = std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap())
        .expect("utf8 workbook rels");
    assert!(
        workbook_rels.contains(r#"Id="rId9""#) && workbook_rels.contains(r#"Target="metadata.xml""#),
        "expected workbook.xml.rels to preserve metadata relationship, got:\n{workbook_rels}"
    );

    let ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).expect("utf8 content types");
    for part in [
        r#"PartName="/xl/metadata.xml""#,
        r#"PartName="/xl/richData/richValueTypes.xml""#,
        r#"PartName="/xl/richData/richValues.xml""#,
    ] {
        assert!(
            ct.contains(part),
            "expected [Content_Types].xml to preserve override for {part}, got:\n{ct}"
        );
    }
}

