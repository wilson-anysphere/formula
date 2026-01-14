use std::collections::BTreeSet;
use std::io::Cursor;

use formula_xlsx::load_from_bytes;
use roxmltree::Document;
use zip::ZipArchive;

const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

fn sheet_drawing_relationship(xml: &str) -> Option<(String, String)> {
    let doc = Document::parse(xml).ok()?;
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        if node.attribute("Type")? != REL_TYPE_DRAWING {
            continue;
        }
        let id = node.attribute("Id")?.to_string();
        let target = node.attribute("Target")?.to_string();
        return Some((id, target));
    }
    None
}

fn drawing_part_names(bytes: &[u8]) -> BTreeSet<String> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut out = BTreeSet::new();
    for idx in 0..archive.len() {
        let Ok(file) = archive.by_index(idx) else {
            continue;
        };
        if file.is_dir() {
            continue;
        }
        let name = file.name();
        let name = name.strip_prefix('/').unwrap_or(name);
        if name.starts_with("xl/drawings/") && name.ends_with(".xml") && !name.contains("/_rels/") {
            out.insert(name.to_string());
        }
    }
    out
}

#[test]
fn xlsx_document_roundtrip_preserves_sheet_drawing_relationship_mapping() {
    let input = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let doc = load_from_bytes(input).expect("load fixture");

    let orig_rels_xml = std::str::from_utf8(
        doc.parts()
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("fixture sheet1 rels"),
    )
    .expect("utf8 sheet rels");
    let (orig_id, orig_target) =
        sheet_drawing_relationship(orig_rels_xml).expect("fixture drawing relationship");

    let saved = doc.save_to_vec().expect("save");
    let doc2 = load_from_bytes(&saved).expect("reload");

    let out_rels_xml = std::str::from_utf8(
        doc2.parts()
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("output sheet1 rels"),
    )
    .expect("utf8 sheet rels");
    let (out_id, out_target) =
        sheet_drawing_relationship(out_rels_xml).expect("output drawing relationship");

    assert_eq!(
        out_id, orig_id,
        "drawing relationship id should be preserved"
    );
    assert_eq!(
        out_target, orig_target,
        "drawing relationship target should be preserved"
    );

    let in_parts = drawing_part_names(input);
    let out_parts = drawing_part_names(&saved);
    assert_eq!(
        out_parts, in_parts,
        "drawing part names should be preserved"
    );
}
