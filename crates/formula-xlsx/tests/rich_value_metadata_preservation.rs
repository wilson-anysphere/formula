use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches, XlsxPackage,
};
use zip::ZipArchive;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn build_rich_value_fixture_xlsx() -> Vec<u8> {
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
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"></Override>
</Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <metadata r:id="rId2"/>
</workbook>"#;

    // Intentionally encode the calcChain relationship as a non-empty element so the calc-chain
    // removal patch path exercises the "skipping" state machine (as opposed to only handling
    // `<Relationship ... />`).
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"></Relationship>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Minimal sheet metadata payload. Patching should preserve the part byte-for-byte.
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="0" maxSupportedVersion="0"/>
  </metadataTypes>
</metadata>"#;

    let calc_chain_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

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

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(calc_chain_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn zip_part_exists(zip_bytes: &[u8], name: &str) -> bool {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).expect("open zip");
    // `ZipFile` borrows the archive, so ensure the result is dropped before `archive`.
    let exists = archive.by_name(name).is_ok();
    exists
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

fn assert_workbook_rels_has_metadata_relationship(rels_xml: &str) {
    let doc = roxmltree::Document::parse(rels_xml).expect("parse workbook rels");
    let rel = doc.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Type")
                .is_some_and(|t| t.ends_with("/metadata") || t.ends_with("/relationships/metadata"))
    });
    let Some(rel) = rel else {
        panic!("expected workbook.xml.rels to contain a metadata relationship, got: {rels_xml}");
    };
    assert!(
        rel.attribute("Target")
            .is_some_and(|t| t.ends_with("metadata.xml")),
        "expected metadata relationship to target metadata.xml, got: {rels_xml}"
    );
}

fn metadata_rel_id_from_workbook_rels(rels_xml: &str) -> String {
    let doc = roxmltree::Document::parse(rels_xml).expect("parse workbook rels");
    let rel = doc.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Type")
                .is_some_and(|t| t.ends_with("/metadata") || t.ends_with("/relationships/metadata"))
            && n.attribute("Target")
                .is_some_and(|t| t.ends_with("metadata.xml"))
    });
    let Some(rel) = rel else {
        panic!("expected workbook.xml.rels to contain a metadata relationship, got: {rels_xml}");
    };
    rel.attribute("Id")
        .or_else(|| rel.attribute("id"))
        .expect("metadata relationship should have Id attribute")
        .to_string()
}

fn metadata_rid_from_workbook_xml(workbook_xml: &str) -> String {
    let doc = roxmltree::Document::parse(workbook_xml).expect("parse workbook xml");
    let metadata = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "metadata");
    let Some(metadata) = metadata else {
        panic!("expected xl/workbook.xml to contain <metadata r:id=\"...\"/>, got: {workbook_xml}");
    };
    metadata
        .attribute((REL_NS, "id"))
        .or_else(|| metadata.attribute("r:id"))
        .or_else(|| metadata.attribute("id"))
        .expect("metadata element should have r:id attribute")
        .to_string()
}

fn assert_workbook_xml_has_metadata_reference(workbook_xml: &str, rels_xml: &str) {
    let rel_id = metadata_rel_id_from_workbook_rels(rels_xml);
    let rid = metadata_rid_from_workbook_xml(workbook_xml);
    assert_eq!(
        rid, rel_id,
        "expected xl/workbook.xml <metadata r:id> to reference the metadata relationship id (rels: {rels_xml}, workbook.xml: {workbook_xml})"
    );
}

#[test]
fn streaming_patcher_preserves_rich_value_metadata_parts_and_drops_vm_on_edit(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_rich_value_fixture_xlsx();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        // Include a material formula so the streaming patcher takes the recalc-policy path that
        // rewrites workbook.xml, workbook.xml.rels, and [Content_Types].xml (where bugs have
        // historically dropped metadata relationships/overrides).
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let out_bytes = out.into_inner();

    let sheet_xml = String::from_utf8(zip_part(&out_bytes, "xl/worksheets/sheet1.xml"))?;
    assert_sheet_a1_preserves_vm_and_cm(&sheet_xml);

    assert!(
        zip_part_exists(&out_bytes, "xl/metadata.xml"),
        "expected xl/metadata.xml to be preserved"
    );

    let workbook_rels = String::from_utf8(zip_part(&out_bytes, "xl/_rels/workbook.xml.rels"))?;
    assert_workbook_rels_has_metadata_relationship(&workbook_rels);

    let workbook_xml = String::from_utf8(zip_part(&out_bytes, "xl/workbook.xml"))?;
    assert_workbook_xml_has_metadata_reference(&workbook_xml, &workbook_rels);

    assert!(
        !zip_part_exists(&out_bytes, "xl/calcChain.xml"),
        "expected xl/calcChain.xml to be dropped after formula edits"
    );
    assert!(
        !workbook_rels.contains("relationships/calcChain"),
        "expected workbook.xml.rels to drop calcChain relationship after formula edits, got: {workbook_rels}"
    );

    let content_types = String::from_utf8(zip_part(&out_bytes, "[Content_Types].xml"))?;
    assert!(
        content_types.contains("/xl/metadata.xml"),
        "expected [Content_Types].xml to preserve metadata override, got: {content_types}"
    );
    assert!(
        !content_types.contains("calcChain.xml"),
        "expected [Content_Types].xml to drop calcChain override after formula edits, got: {content_types}"
    );

    Ok(())
}

#[test]
fn package_patcher_preserves_rich_value_metadata_parts_and_drops_vm_on_edit(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_rich_value_fixture_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())?;
    assert_sheet_a1_preserves_vm_and_cm(sheet_xml);

    assert!(
        pkg.part("xl/metadata.xml").is_some(),
        "expected xl/metadata.xml to be preserved"
    );

    let workbook_rels = std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap())?;
    assert_workbook_rels_has_metadata_relationship(workbook_rels);

    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap())?;
    assert_workbook_xml_has_metadata_reference(workbook_xml, workbook_rels);

    assert!(
        pkg.part("xl/calcChain.xml").is_none(),
        "expected xl/calcChain.xml to be dropped after formula edits"
    );
    assert!(
        !workbook_rels.contains("relationships/calcChain"),
        "expected workbook.xml.rels to drop calcChain relationship after formula edits, got: {workbook_rels}"
    );

    let content_types = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())?;
    assert!(
        content_types.contains("/xl/metadata.xml"),
        "expected [Content_Types].xml to preserve metadata override, got: {content_types}"
    );
    assert!(
        !content_types.contains("calcChain.xml"),
        "expected [Content_Types].xml to drop calcChain override after formula edits, got: {content_types}"
    );

    Ok(())
}
