use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::Workbook;

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");

#[test]
fn preserved_drawing_parts_include_smartart_diagrams() {
    let package = XlsxPackage::from_bytes(FIXTURE).expect("load smartart.xlsx fixture");
    let preserved = package
        .preserve_drawing_parts()
        .expect("preserve drawing parts");

    assert!(
        preserved.parts.contains_key("xl/diagrams/data1.xml"),
        "expected drawing preservation to include SmartArt diagram parts (xl/diagrams/*)"
    );

    // Simulate a write pipeline that regenerates the workbook XML (dropping drawings) and then
    // re-applies the preserved drawing parts.
    let mut workbook = Workbook::new();
    workbook.add_worksheet();
    let regenerated_bytes = workbook.save_to_buffer().expect("save regenerated workbook");

    let mut regenerated = XlsxPackage::from_bytes(&regenerated_bytes).expect("load regenerated workbook");
    regenerated
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved drawing parts");
    let merged_bytes = regenerated.write_to_bytes().expect("write merged workbook");

    let merged = XlsxPackage::from_bytes(&merged_bytes).expect("load merged workbook");

    for part in [
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/diagrams/data1.xml",
        "xl/diagrams/layout1.xml",
        "xl/diagrams/quickStyle1.xml",
        "xl/diagrams/colors1.xml",
    ] {
        assert!(merged.part(part).is_some(), "expected {part} to be present after applying preserved drawing parts");
    }
}

