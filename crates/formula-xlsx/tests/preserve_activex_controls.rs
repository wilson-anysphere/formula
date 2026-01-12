use std::path::Path;

use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::Workbook;

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
}
