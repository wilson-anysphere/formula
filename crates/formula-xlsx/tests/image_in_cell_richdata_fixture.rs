use std::path::Path;

use formula_model::CellRef;
use formula_xlsx::extract_embedded_images;
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

    let embedded = extract_embedded_images(&pkg).expect("extract embedded images");
    assert_eq!(embedded.len(), 1, "expected one embedded image");
    let entry = &embedded[0];

    assert_eq!(entry.sheet_part, "xl/worksheets/sheet1.xml");
    assert_eq!(entry.cell, CellRef::from_a1("A1").unwrap());
    assert_eq!(entry.image_target, "xl/media/image1.png");

    let image_bytes = pkg
        .part("xl/media/image1.png")
        .expect("fixture image bytes exist")
        .to_vec();
    assert_eq!(entry.bytes, image_bytes);

    // This fixture uses `richValue.xml` and doesn't include the `_localImage` fields used by the
    // alternate `rdrichvalue.xml` schema, so we don't expect alt text / decorative metadata.
    assert_eq!(entry.alt_text, None);
    assert!(!entry.decorative);
}
