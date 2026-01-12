use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches,
};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_rich_data_fixture_xlsx(metadata_bytes: &[u8], rich_data_bytes: &[u8]) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // Include a relationship to `metadata.xml` and a calcChain relationship so the streaming
    // patcher will rewrite workbook.xml.rels (dropping calcChain) when formulas change. The test
    // asserts the metadata relationship is preserved after rewriting.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
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
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/rdrichvalues.xml" ContentType="application/vnd.ms-excel.rdrichvalues+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"#;

    // `vm="1"` is used by Excel when rich data types are present; include it to ensure the
    // streaming patcher doesn't drop the attribute while patching.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let calc_chain = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

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

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_bytes).unwrap();

    zip.start_file("xl/richData/rdrichvalues.xml", options).unwrap();
    zip.write_all(rich_data_bytes).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(calc_chain.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn streaming_patcher_preserves_rich_data_parts_and_workbook_rels(
) -> Result<(), Box<dyn std::error::Error>> {
    let metadata_bytes = b"<metadata>FORMULA_TEST_METADATA</metadata>";
    let rich_data_bytes = b"<rdrichvalues>FORMULA_TEST_RICH_DATA</rdrichvalues>";
    let input_bytes = build_rich_data_fixture_xlsx(metadata_bytes, rich_data_bytes);

    // Apply a formula patch (not just a value patch) so the streaming patcher will rewrite
    // `xl/_rels/workbook.xml.rels` + `[Content_Types].xml` as part of calcChain removal logic.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(input_bytes), &mut out, &patches)?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;

    // Rich-data parts must be preserved byte-for-byte (images-in-cells depend on these).
    let mut out_metadata = Vec::new();
    archive
        .by_name("xl/metadata.xml")?
        .read_to_end(&mut out_metadata)?;
    assert_eq!(out_metadata, metadata_bytes);

    let mut out_rich_data = Vec::new();
    archive
        .by_name("xl/richData/rdrichvalues.xml")?
        .read_to_end(&mut out_rich_data)?;
    assert_eq!(out_rich_data, rich_data_bytes);

    // workbook.xml.rels must still contain the relationship to metadata.xml after rewriting.
    let mut rels_xml = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    let rels_doc = roxmltree::Document::parse(&rels_xml)?;
    let has_metadata_rel = rels_doc.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Target")
                .is_some_and(|t| t == "metadata.xml" || t.ends_with("/metadata.xml"))
    });
    assert!(
        has_metadata_rel,
        "expected workbook.xml.rels to preserve metadata relationship; got {rels_xml}"
    );

    Ok(())
}
