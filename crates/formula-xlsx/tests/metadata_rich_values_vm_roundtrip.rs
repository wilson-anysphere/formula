use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::ZipArchive;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn preserves_metadata_xml_and_vm_attrs_on_edit_roundtrip() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/metadata/rich-values-vm.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let mut doc = load_from_bytes(&fixture_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    // Edit a different cell so we exercise "preserve unrelated sheet XML + metadata parts".
    assert!(doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("B2")?,
        CellValue::Number(99.0)
    ));

    let saved = doc.save_to_vec()?;

    // Ensure `xl/metadata.xml` survives byte-for-byte.
    assert_eq!(
        zip_part(&fixture_bytes, "xl/metadata.xml"),
        zip_part(&saved, "xl/metadata.xml")
    );

    // Ensure the workbook relationship to the metadata part is still present.
    let rels_bytes = zip_part(&saved, "xl/_rels/workbook.xml.rels");
    let rels = std::str::from_utf8(&rels_bytes)?;
    assert!(rels.contains("relationships/metadata"));
    assert!(rels.contains("Target=\"metadata.xml\""));

    // Ensure the `<metadata r:id="..."/>` element inside `xl/workbook.xml` is retained and still
    // points at the metadata relationship in `workbook.xml.rels`.
    let workbook_xml_bytes = zip_part(&saved, "xl/workbook.xml");
    let workbook_xml = std::str::from_utf8(&workbook_xml_bytes)?;
    let rels_doc = roxmltree::Document::parse(rels)?;
    let metadata_rel = rels_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name().eq_ignore_ascii_case("Relationship")
                && n.attribute("Type").is_some_and(|t| {
                    t.ends_with("/metadata") || t.ends_with("/relationships/metadata")
                })
                && n.attribute("Target")
                    .is_some_and(|t| t.ends_with("metadata.xml"))
        })
        .expect("workbook.xml.rels must contain metadata relationship");
    let metadata_rel_id = metadata_rel
        .attribute("Id")
        .expect("metadata relationship should have Id");

    let workbook_doc = roxmltree::Document::parse(workbook_xml)?;
    let metadata_node = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("metadata"))
        .expect("expected <metadata> element in xl/workbook.xml");
    let metadata_rid = metadata_node
        .attribute((REL_NS, "id"))
        .or_else(|| metadata_node.attribute("r:id"))
        .or_else(|| metadata_node.attribute("id"))
        .expect("<metadata> element should have r:id attribute");
    assert_eq!(
        metadata_rid, metadata_rel_id,
        "expected workbook.xml <metadata r:id> to match metadata relationship Id (rels: {rels}, workbook.xml: {workbook_xml})"
    );

    // Ensure the `vm="..."` attribute on the original cell is preserved.
    let sheet_xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;
    let cell_a1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("cell A1 exists");
    assert_eq!(cell_a1.attribute("vm"), Some("1"));

    Ok(())
}
