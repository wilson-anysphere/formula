use std::path::Path;

use formula_xlsx::XlsxPackage;
use roxmltree::Document;
use rust_xlsxwriter::Workbook;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn attr_rel_id(node: roxmltree::Node<'_, '_>) -> Option<String> {
    node.attribute((REL_NS, "id"))
        .or_else(|| node.attribute("r:id"))
        .or_else(|| node.attribute("id"))
        .map(|s| s.to_string())
}

fn relationship_target(rels_xml: &str, rel_id: &str) -> Option<String> {
    let doc = Document::parse(rels_xml).ok()?;
    doc.descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Id") == Some(rel_id)
        })
        .and_then(|n| n.attribute("Target"))
        .map(|s| s.to_string())
}

#[test]
fn preserves_activex_controls_across_regeneration() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/activex-control.xlsx");
    let fixture_bytes = std::fs::read(&fixture).expect("read fixture");
    let pkg = XlsxPackage::from_bytes(&fixture_bytes).expect("load fixture package");
    let preserved = pkg
        .preserve_drawing_parts()
        .expect("preserve drawing parts");
    assert!(
        !preserved.is_empty(),
        "fixture should preserve at least one part"
    );

    let mut workbook = Workbook::new();
    workbook.add_worksheet();
    let regenerated_bytes = workbook
        .save_to_buffer()
        .expect("save regenerated workbook");
    let mut regenerated_pkg =
        XlsxPackage::from_bytes(&regenerated_bytes).expect("load regenerated package");

    regenerated_pkg
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved parts");
    let merged_bytes = regenerated_pkg
        .write_to_bytes()
        .expect("write merged workbook");
    let merged_pkg = XlsxPackage::from_bytes(&merged_bytes).expect("load merged package");

    // Preserve all control-chain parts byte-for-byte.
    for part in [
        "xl/ctrlProps/ctrlProp1.xml",
        "xl/ctrlProps/_rels/ctrlProp1.xml.rels",
        "xl/activeX/activeX1.xml",
        "xl/activeX/_rels/activeX1.xml.rels",
        "xl/activeX/activeX1.bin",
    ] {
        assert_eq!(
            pkg.part(part),
            merged_pkg.part(part),
            "mismatch for preserved part {part}",
        );
    }

    assert!(
        merged_pkg.part("xl/ctrlProps/ctrlProp1.xml").is_some(),
        "missing ctrlProps part",
    );
    assert!(
        merged_pkg.part("xl/activeX/activeX1.xml").is_some(),
        "missing activeX XML part",
    );
    assert!(
        merged_pkg.part("xl/activeX/activeX1.bin").is_some(),
        "missing activeX binary part",
    );

    let sheet_xml = std::str::from_utf8(
        merged_pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )
    .expect("sheet1.xml is utf-8");
    assert!(
        sheet_xml.contains("<controls"),
        "sheet1.xml missing <controls>"
    );
    assert!(
        sheet_xml.contains("r:id"),
        "sheet1.xml missing control relationship id",
    );

    let control_rel_id = {
        let doc = Document::parse(sheet_xml).expect("parse sheet xml");
        let control = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "control")
            .expect("sheet must contain a control element");
        attr_rel_id(control).expect("control missing r:id")
    };

    let sheet_rels = std::str::from_utf8(
        merged_pkg
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet1.xml.rels exists"),
    )
    .expect("sheet1.xml.rels is utf-8");
    assert!(
        sheet_rels.contains("ctrlProps/ctrlProp1.xml"),
        "worksheet rels missing ctrlProps relationship",
    );

    let sheet_rels_target =
        relationship_target(sheet_rels, &control_rel_id).expect("control relationship exists");
    assert!(
        sheet_rels_target.ends_with("ctrlProps/ctrlProp1.xml"),
        "control relationship should target ctrlProps: got {sheet_rels_target}",
    );

    let ctrl_rels = std::str::from_utf8(
        merged_pkg
            .part("xl/ctrlProps/_rels/ctrlProp1.xml.rels")
            .expect("ctrlProp1.xml.rels exists"),
    )
    .expect("ctrlProp1.xml.rels is utf-8");
    assert!(
        ctrl_rels.contains("activeX/activeX1.xml"),
        "ctrlProps rels missing activeX relationship",
    );

    let activex_rels = std::str::from_utf8(
        merged_pkg
            .part("xl/activeX/_rels/activeX1.xml.rels")
            .expect("activeX1.xml.rels exists"),
    )
    .expect("activeX1.xml.rels is utf-8");
    assert!(
        activex_rels.contains("activeX1.bin"),
        "activeX rels missing binary relationship",
    );

    let content_types = std::str::from_utf8(
        merged_pkg
            .part("[Content_Types].xml")
            .expect("content types exists"),
    )
    .expect("content types is utf-8");
    assert!(
        content_types.contains("Extension=\"bin\""),
        "content types missing default for .bin",
    );
}
