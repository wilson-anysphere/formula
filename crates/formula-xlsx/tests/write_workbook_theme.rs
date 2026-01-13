use std::io::Cursor;

use formula_model::{ArgbColor, Workbook};
use formula_xlsx::{write_workbook_to_writer, XlsxPackage};

#[test]
fn write_workbook_emits_theme_part_and_wires_rels_and_content_types() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1").unwrap();

    // Pick non-default accent colors so we can assert they appear in `theme1.xml`.
    workbook.theme.accent1 = ArgbColor::new(0xFF1A2B3C);
    workbook.theme.accent2 = ArgbColor::new(0xFF4D5E6F);
    workbook.theme.accent3 = ArgbColor::new(0xFF778899);
    workbook.theme.accent4 = ArgbColor::new(0xFFAABBCC);
    workbook.theme.accent5 = ArgbColor::new(0xFFDDEEFF);
    workbook.theme.accent6 = ArgbColor::new(0xFF102030);

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let theme_xml = std::str::from_utf8(pkg.part("xl/theme/theme1.xml").expect("theme part"))
        .expect("theme xml utf8");
    for expected in [
        "1A2B3C", "4D5E6F", "778899", "AABBCC", "DDEEFF", "102030",
    ] {
        assert!(
            theme_xml.contains(expected),
            "expected theme1.xml to contain {expected}, got:\n{theme_xml}"
        );
    }

    let rels_xml =
        std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").expect("workbook rels part"))
            .expect("rels utf8");
    let rels_doc = roxmltree::Document::parse(rels_xml).expect("parse workbook rels");
    assert!(
        rels_doc.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type")
                    == Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
                && n.attribute("Target") == Some("theme/theme1.xml")
        }),
        "expected workbook relationships to include theme, got:\n{rels_xml}"
    );

    let ct_xml = std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types part"))
        .expect("ct utf8");
    let ct_doc = roxmltree::Document::parse(ct_xml).expect("parse content types");
    assert!(
        ct_doc.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/theme/theme1.xml")
                && n.attribute("ContentType")
                    == Some("application/vnd.openxmlformats-officedocument.theme+xml")
        }),
        "expected [Content_Types].xml to include theme override, got:\n{ct_xml}"
    );
}

