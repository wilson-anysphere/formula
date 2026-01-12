use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn fixture_bytes() -> Vec<u8> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/multi-sheet.xlsx");
    std::fs::read(&fixture_path).expect("fixture exists")
}

fn sheet_part_for_name(saved: &[u8], name: &str) -> String {
    let workbook_xml = zip_part(saved, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");
    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");

    let sheet = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheet" && n.attribute("name") == Some(name))
        .expect("expected sheet in workbook.xml");

    // `r:id` is a namespaced attribute, so roxmltree exposes it as `(namespace-uri, local-name)`.
    const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let rid = sheet
        .attribute((REL_NS, "id"))
        .or_else(|| sheet.attribute("r:id"))
        .expect("expected r:id on sheet");

    let workbook_rels = zip_part(saved, "xl/_rels/workbook.xml.rels");
    let rels = formula_xlsx::openxml::parse_relationships(&workbook_rels).expect("parse rels");
    let rel = rels
        .into_iter()
        .find(|rel| rel.id == rid)
        .expect("expected sheet relationship");

    formula_xlsx::openxml::resolve_target("xl/workbook.xml", &rel.target)
}

#[test]
fn writer_emits_vm_cm_on_added_sheet() {
    let fixture = fixture_bytes();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    // This sheet has no backing `xl/worksheets/sheet*.xml` in the original package, so it will be
    // generated from scratch via `render_sheet_data`/`append_cell_xml` on save.
    let sheet_id = doc.workbook.add_sheet("Added").expect("add sheet");
    let cell = CellRef::from_a1("A1").expect("valid A1");
    doc.set_cell_value(
        sheet_id,
        cell,
        CellValue::String("MSFT".to_string()),
    );

    // Inject value/cell metadata that should be serialized onto the `<c>` element.
    let meta = doc.meta.cell_meta.entry((sheet_id, cell)).or_default();
    meta.vm = Some("1".to_string());
    meta.cm = Some("2".to_string());

    let saved = doc.save_to_vec().expect("save xlsx");
    let sheet_part = sheet_part_for_name(&saved, "Added");

    let sheet_xml = zip_part(&saved, &sheet_part);
    let sheet_xml_str = std::str::from_utf8(&sheet_xml).expect("sheet xml utf-8");
    let parsed = roxmltree::Document::parse(sheet_xml_str).expect("parse sheet xml");

    let cell_node = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    assert_eq!(
        cell_node.attribute("vm"),
        Some("1"),
        "expected vm attribute to be written, got: {sheet_xml_str}"
    );
    assert_eq!(
        cell_node.attribute("cm"),
        Some("2"),
        "expected cm attribute to be written, got: {sheet_xml_str}"
    );
}
