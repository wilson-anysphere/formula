use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn build_richdata_xlsx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.richValueTypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.richValues+xml"/>
</Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <calcPr fullCalcOnLoad="0"/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    // Payloads are not parsed by the patcher; they only need to be stable so we can assert
    // byte-for-byte preservation.
    let calc_chain_xml = b"<calcChain/>\n";
    let metadata_xml = b"<metadata>stable-metadata</metadata>\n";
    let metadata_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/richDataTypes" Target="richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://example.com/richDataValues" Target="richData/richValues.xml"/>
</Relationships>"#;
    let rich_value_types_xml = b"<richValueTypes>stable-types</richValueTypes>\n";
    let rich_values_xml = b"<richValues>stable-values</richValues>\n";

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(calc_chain_xml).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml).unwrap();

    zip.start_file("xl/_rels/metadata.xml.rels", options)
        .unwrap();
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
fn recalc_policy_calcchain_drop_preserves_richdata_infrastructure() {
    let bytes = build_richdata_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    // Capture original bytes for rich-data parts; they should remain byte-identical even after
    // workbook-level recalc policy rewrites.
    let metadata_xml_before = pkg.part("xl/metadata.xml").unwrap().to_vec();
    let metadata_rels_before = pkg.part("xl/_rels/metadata.xml.rels").unwrap().to_vec();
    let rich_value_types_before = pkg.part("xl/richData/richValueTypes.xml").unwrap().to_vec();
    let rich_values_before = pkg.part("xl/richData/richValues.xml").unwrap().to_vec();

    // Apply a formula patch (not just a value) so the default recalc policy triggers calcChain
    // removal and workbook-level metadata rewrites.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("C1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::Number(43.0), "B1+1"),
    );

    pkg.apply_cell_patches(&patches).expect("apply patches");

    assert!(
        pkg.part("xl/calcChain.xml").is_none(),
        "calcChain.xml should be removed after formula edits"
    );

    let workbook_rels =
        std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(
        !workbook_rels.contains("calcChain.xml"),
        "workbook.xml.rels should no longer reference calcChain.xml, got: {workbook_rels}"
    );
    assert!(
        workbook_rels.contains(r#"Id="rId9""#)
            && workbook_rels.contains(r#"Target="metadata.xml""#),
        "workbook.xml.rels should preserve metadata relationship, got: {workbook_rels}"
    );

    let content_types = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
    assert!(
        !content_types.contains("/xl/calcChain.xml"),
        "[Content_Types].xml should drop calcChain override, got: {content_types}"
    );
    assert!(
        content_types.contains("/xl/metadata.xml"),
        "[Content_Types].xml should preserve metadata override, got: {content_types}"
    );
    assert!(
        content_types.contains("/xl/richData/richValueTypes.xml"),
        "[Content_Types].xml should preserve richValueTypes override, got: {content_types}"
    );
    assert!(
        content_types.contains("/xl/richData/richValues.xml"),
        "[Content_Types].xml should preserve richValues override, got: {content_types}"
    );

    assert_eq!(
        pkg.part("xl/metadata.xml").unwrap(),
        metadata_xml_before.as_slice(),
        "metadata.xml bytes should be preserved"
    );
    assert_eq!(
        pkg.part("xl/_rels/metadata.xml.rels").unwrap(),
        metadata_rels_before.as_slice(),
        "metadata.xml.rels bytes should be preserved"
    );
    assert_eq!(
        pkg.part("xl/richData/richValueTypes.xml").unwrap(),
        rich_value_types_before.as_slice(),
        "richValueTypes.xml bytes should be preserved"
    );
    assert_eq!(
        pkg.part("xl/richData/richValues.xml").unwrap(),
        rich_values_before.as_slice(),
        "richValues.xml bytes should be preserved"
    );
}
