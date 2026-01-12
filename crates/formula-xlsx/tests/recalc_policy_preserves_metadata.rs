use std::io::{Cursor, Read, Seek, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, patch_xlsx_streaming, WorksheetCellPatch};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn read_zip_part_to_string<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<String, Box<dyn std::error::Error>> {
    let mut file = archive.by_name(name)?;
    let mut xml = String::new();
    file.read_to_string(&mut xml)?;
    Ok(xml)
}

fn build_rich_values_fixture() -> Vec<u8> {
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
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    // Keep both calcChain and metadata overrides; the regression under test is that metadata must
    // survive when recalc policy drops calcChain.
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
 <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
 <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData>
  <row r="1">
   <c r="A1">
    <f>1+1</f>
    <v>2</v>
   </c>
  </row>
 </sheetData>
</worksheet>"#;

    let calc_chain_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(calc_chain_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn assert_calc_chain_dropped_but_metadata_preserved(
    saved: &[u8],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(saved))?;

    assert!(
        archive.by_name("xl/calcChain.xml").is_err(),
        "expected calcChain.xml to be removed after formula edit"
    );
    archive
        .by_name("xl/metadata.xml")
        .expect("expected metadata.xml to be preserved");

    let rels_xml = read_zip_part_to_string(&mut archive, "xl/_rels/workbook.xml.rels")?;
    assert!(
        !rels_xml.contains("calcChain.xml"),
        "workbook.xml.rels relationship targeting calcChain.xml should be removed"
    );
    assert!(
        rels_xml.contains("metadata.xml") && rels_xml.contains("relationships/metadata"),
        "workbook.xml.rels should preserve the metadata relationship"
    );

    let ct_xml = read_zip_part_to_string(&mut archive, "[Content_Types].xml")?;
    assert!(
        !ct_xml.contains("/xl/calcChain.xml"),
        "[Content_Types].xml override for calcChain.xml should be removed"
    );
    assert!(
        ct_xml.contains("/xl/metadata.xml"),
        "[Content_Types].xml should preserve the metadata override"
    );

    Ok(())
}

#[test]
fn xlsx_document_formula_edit_preserves_metadata_relationships_and_overrides(
) -> Result<(), Box<dyn std::error::Error>> {
    let input = build_rich_values_fixture();
    let mut doc = load_from_bytes(&input)?;
    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .expect("fixture should have a sheet")
        .id;

    assert!(
        doc.set_cell_formula(
            sheet_id,
            CellRef::from_a1("A1")?,
            Some("=1+2".to_string()),
        ),
        "expected formula edit to succeed"
    );

    let saved = doc.save_to_vec()?;
    assert_calc_chain_dropped_but_metadata_preserved(&saved)?;
    Ok(())
}

#[test]
fn streaming_formula_patch_preserves_metadata_relationships_and_overrides(
) -> Result<(), Box<dyn std::error::Error>> {
    let input = build_rich_values_fixture();
    let mut output = Cursor::new(Vec::new());

    let patches = [WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(3.0),
        Some("=1+2".to_string()),
    )];

    patch_xlsx_streaming(Cursor::new(input), &mut output, &patches)?;
    let saved = output.into_inner();

    assert_calc_chain_dropped_but_metadata_preserved(&saved)?;
    Ok(())
}

