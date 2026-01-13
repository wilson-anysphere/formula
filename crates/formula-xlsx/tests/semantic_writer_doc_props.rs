use std::io::{Cursor, Read};

use formula_model::Workbook;
use formula_xlsx::write_workbook_to_writer;
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn relationship_target_from_rels_xml(rels_xml: &str, rel_type: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(rels_xml).ok()?;
    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Relationship" {
            continue;
        }
        if node.attribute("Type") != Some(rel_type) {
            continue;
        }
        return node.attribute("Target").map(|s| s.to_string());
    }
    None
}

fn content_type_from_content_types_xml(ct_xml: &str, part_name: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(ct_xml).ok()?;
    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Override" {
            continue;
        }
        if node.attribute("PartName") != Some(part_name) {
            continue;
        }
        return node.attribute("ContentType").map(|s| s.to_string());
    }
    None
}

#[test]
fn semantic_writer_emits_doc_props_parts_and_relationships() {
    const REL_TYPE_CORE: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    const REL_TYPE_APP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
    const CT_APP: &str = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1").expect("add sheet");

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();

    // Zip contains the parts.
    zip_part(&bytes, "docProps/core.xml");
    zip_part(&bytes, "docProps/app.xml");

    // Root relationships reference the parts.
    let rels_xml = zip_part(&bytes, "_rels/.rels");
    let rels_xml = std::str::from_utf8(&rels_xml).expect("utf8 rels");
    assert_eq!(
        relationship_target_from_rels_xml(rels_xml, REL_TYPE_CORE).as_deref(),
        Some("docProps/core.xml"),
        "missing/incorrect core-properties relationship; rels xml:\n{rels_xml}"
    );
    assert_eq!(
        relationship_target_from_rels_xml(rels_xml, REL_TYPE_APP).as_deref(),
        Some("docProps/app.xml"),
        "missing/incorrect extended-properties relationship; rels xml:\n{rels_xml}"
    );

    // Content types declare the parts.
    let ct_xml = zip_part(&bytes, "[Content_Types].xml");
    let ct_xml = std::str::from_utf8(&ct_xml).expect("utf8 content types");
    assert_eq!(
        content_type_from_content_types_xml(ct_xml, "/docProps/core.xml").as_deref(),
        Some(CT_CORE),
        "missing/incorrect core.xml content type override; ct xml:\n{ct_xml}"
    );
    assert_eq!(
        content_type_from_content_types_xml(ct_xml, "/docProps/app.xml").as_deref(),
        Some(CT_APP),
        "missing/incorrect app.xml content type override; ct xml:\n{ct_xml}"
    );
}

