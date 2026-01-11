use formula_xlsx::XlsxPackage;
use roxmltree::Document;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_TYPE_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_TYPE_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

#[test]
fn drawing_preservation_renumbers_relationship_ids_on_conflict() {
    let preserved_source_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let preserved_source_pkg = XlsxPackage::from_bytes(preserved_source_bytes).expect("load source");
    let preserved = preserved_source_pkg
        .preserve_drawing_parts()
        .expect("preserve drawing parts");
    assert!(!preserved.is_empty(), "expected fixture to preserve drawings");

    let dest_bytes = include_bytes!("../../../fixtures/xlsx/hyperlinks/hyperlinks.xlsx");
    let mut dest_pkg = XlsxPackage::from_bytes(dest_bytes).expect("load destination");
    dest_pkg
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply drawing parts");

    // Round-trip through zip writer to mirror production output.
    let merged_bytes = dest_pkg.write_to_bytes().expect("write merged pkg");
    let merged_pkg = XlsxPackage::from_bytes(&merged_bytes).expect("read merged pkg");

    let sheet_xml =
        std::str::from_utf8(merged_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet_doc = Document::parse(sheet_xml).unwrap();
    let drawing = sheet_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "drawing")
        .expect("sheet should contain <drawing>");
    let drawing_rid = drawing
        .attribute((REL_NS, "id"))
        .or_else(|| drawing.attribute("r:id"))
        .or_else(|| drawing.attribute("id"))
        .expect("<drawing> should have r:id");

    assert_ne!(
        drawing_rid, "rId1",
        "drawing relationship should not reuse destination hyperlink rId1"
    );

    let sheet_rels_xml = std::str::from_utf8(
        merged_pkg
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    let (drawing_rel_type, drawing_rel_target) =
        find_relationship(sheet_rels_xml, drawing_rid).expect("drawing relationship exists");
    assert_eq!(drawing_rel_type, REL_TYPE_DRAWING);
    assert_eq!(drawing_rel_target, "../drawings/drawing1.xml");

    // Original hyperlink relationships should remain untouched.
    let (link1_type, _link1_target) =
        find_relationship(sheet_rels_xml, "rId1").expect("hyperlink rId1 exists");
    assert_eq!(link1_type, REL_TYPE_HYPERLINK);
    let (link2_type, _link2_target) =
        find_relationship(sheet_rels_xml, "rId2").expect("hyperlink rId2 exists");
    assert_eq!(link2_type, REL_TYPE_HYPERLINK);
}

fn find_relationship(xml: &str, id: &str) -> Option<(String, String)> {
    let doc = Document::parse(xml).ok()?;
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        if node.attribute("Id")? != id {
            continue;
        }
        return Some((
            node.attribute("Type")?.to_string(),
            node.attribute("Target")?.to_string(),
        ));
    }
    None
}

