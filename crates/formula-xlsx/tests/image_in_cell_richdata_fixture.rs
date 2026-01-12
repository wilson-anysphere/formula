use std::path::Path;

use formula_xlsx::XlsxPackage;

#[test]
fn image_in_cell_richdata_fixture_contains_expected_parts() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/image-in-cell-richdata.xlsx");
    let bytes = std::fs::read(&fixture).expect("read image-in-cell-richdata fixture");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("load fixture package");

    for part in [
        "xl/metadata.xml",
        "xl/richData/richValue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert!(pkg.part(part).is_some(), "expected fixture to contain {part}");
    }

    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )
    .expect("sheet1.xml is utf-8");
    assert!(
        sheet_xml.contains(r#"vm="0""#),
        "expected sheet1.xml to contain a cell with vm=\"0\""
    );

    let workbook_rels = std::str::from_utf8(
        pkg.part("xl/_rels/workbook.xml.rels")
            .expect("workbook.xml.rels exists"),
    )
    .expect("workbook.xml.rels is utf-8");
    assert!(
        workbook_rels.contains(r#"Id="rId99""#) && workbook_rels.contains(r#"Target="metadata.xml""#),
        "expected workbook.xml.rels to contain rId99 -> metadata.xml, got: {workbook_rels}"
    );

    let content_types = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())
        .expect("[Content_Types].xml is utf-8");
    for needle in [
        r#"PartName="/xl/metadata.xml""#,
        r#"PartName="/xl/richData/richValue.xml""#,
        r#"PartName="/xl/richData/richValueRel.xml""#,
        r#"Extension="png""#,
    ] {
        assert!(
            content_types.contains(needle),
            "expected [Content_Types].xml to contain {needle}, got: {content_types}"
        );
    }
}

