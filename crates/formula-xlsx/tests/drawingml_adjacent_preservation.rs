use std::path::Path;

use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::Workbook;

fn load_fixture(path: &Path) -> Vec<u8> {
    std::fs::read(path).unwrap_or_else(|_| panic!("failed to read fixture {}", path.display()))
}

fn build_minimal_workbook_bytes() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Sheet1").unwrap();
    workbook.save_to_buffer().unwrap()
}

#[test]
fn preserve_and_apply_background_picture() {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/background-image.xlsx");
    let fixture_bytes = load_fixture(&fixture_path);
    let pkg = XlsxPackage::from_bytes(&fixture_bytes).expect("load background-image.xlsx");
    let preserved = pkg
        .preserve_drawing_parts()
        .expect("preserve background image parts");

    let mut out_pkg =
        XlsxPackage::from_bytes(&build_minimal_workbook_bytes()).expect("load target workbook");
    out_pkg
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved parts");

    let written = out_pkg.write_to_bytes().expect("write workbook");
    let written_pkg = XlsxPackage::from_bytes(&written).expect("reload written workbook");

    assert!(
        written_pkg.part("xl/media/image1.png").is_some(),
        "expected background image media part to be present"
    );

    let sheet_xml =
        std::str::from_utf8(written_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        sheet_xml.contains("<picture"),
        "expected <picture> in sheet XML"
    );
    assert!(
        sheet_xml.contains("r:id"),
        "expected r:id in <picture> element"
    );
}

#[test]
fn preserve_and_apply_ole_object_embedding() {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/ole-object.xlsx");
    let fixture_bytes = load_fixture(&fixture_path);
    let pkg = XlsxPackage::from_bytes(&fixture_bytes).expect("load ole-object.xlsx");
    let preserved = pkg
        .preserve_drawing_parts()
        .expect("preserve ole object parts");

    let mut out_pkg =
        XlsxPackage::from_bytes(&build_minimal_workbook_bytes()).expect("load target workbook");
    out_pkg
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved parts");

    let written = out_pkg.write_to_bytes().expect("write workbook");
    let written_pkg = XlsxPackage::from_bytes(&written).expect("reload written workbook");

    assert!(
        written_pkg.part("xl/embeddings/oleObject1.bin").is_some(),
        "expected oleObject1.bin embedding part to be present"
    );

    let sheet_xml =
        std::str::from_utf8(written_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        sheet_xml.contains("<oleObjects"),
        "expected <oleObjects> in sheet XML"
    );
}

#[test]
fn preserve_and_apply_chart_sheet() {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/charts/chart-sheet.xlsx");
    let fixture_bytes = load_fixture(&fixture_path);
    let pkg = XlsxPackage::from_bytes(&fixture_bytes).expect("load chart-sheet.xlsx");
    let preserved = pkg
        .preserve_drawing_parts()
        .expect("preserve chart sheet parts");

    let mut out_pkg =
        XlsxPackage::from_bytes(&build_minimal_workbook_bytes()).expect("load target workbook");
    out_pkg
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved parts");

    let written = out_pkg.write_to_bytes().expect("write workbook");
    let written_pkg = XlsxPackage::from_bytes(&written).expect("reload written workbook");

    assert!(
        written_pkg.part("xl/chartsheets/sheet1.xml").is_some(),
        "expected chartsheet part to be present"
    );

    let workbook_xml = std::str::from_utf8(written_pkg.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(
        workbook_xml.contains("name=\"Chart1\""),
        "expected workbook.xml to contain Chart1 sheet entry"
    );

    let workbook_rels =
        std::str::from_utf8(written_pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(
        workbook_rels.contains("relationships/chartsheet"),
        "expected workbook.xml.rels to contain chartsheet relationship"
    );
}
