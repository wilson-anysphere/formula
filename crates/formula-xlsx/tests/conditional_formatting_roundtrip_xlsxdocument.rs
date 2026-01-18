use formula_xlsx::{assert_xml_semantic_eq, load_from_bytes, XlsxPackage};

const OFFICE2007_FIXTURE: &[u8] = include_bytes!("fixtures/conditional_formatting_2007.xlsx");
const X14_FIXTURE: &[u8] = include_bytes!("fixtures/conditional_formatting_x14.xlsx");

fn assert_roundtrip_preserves_conditional_formatting_parts(fixture: &[u8]) {
    let doc = load_from_bytes(fixture).expect("load fixture via XlsxDocument");
    let saved = doc.save_to_vec().expect("save XlsxDocument");

    let original_pkg = XlsxPackage::from_bytes(fixture).expect("open original fixture package");
    let saved_pkg = XlsxPackage::from_bytes(&saved).expect("open saved package");

    assert_xml_semantic_eq(
        original_pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("fixture must contain xl/worksheets/sheet1.xml"),
        saved_pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("saved file must contain xl/worksheets/sheet1.xml"),
    )
    .unwrap();

    assert_xml_semantic_eq(
        original_pkg
            .part("xl/styles.xml")
            .expect("fixture must contain xl/styles.xml"),
        saved_pkg
            .part("xl/styles.xml")
            .expect("saved file must contain xl/styles.xml"),
    )
    .unwrap();
}

#[test]
fn xlsxdocument_roundtrip_preserves_office2007_conditional_formatting_and_dxfs() {
    assert_roundtrip_preserves_conditional_formatting_parts(OFFICE2007_FIXTURE);
}

#[test]
fn xlsxdocument_roundtrip_preserves_x14_conditional_formatting_and_dxfs() {
    let doc = load_from_bytes(X14_FIXTURE).expect("load x14 fixture via XlsxDocument");
    let saved = doc.save_to_vec().expect("save XlsxDocument");

    let original_pkg = XlsxPackage::from_bytes(X14_FIXTURE).expect("open original fixture package");
    let saved_pkg = XlsxPackage::from_bytes(&saved).expect("open saved package");

    let original_sheet = original_pkg
        .part("xl/worksheets/sheet1.xml")
        .expect("fixture must contain xl/worksheets/sheet1.xml");
    let saved_sheet = saved_pkg
        .part("xl/worksheets/sheet1.xml")
        .expect("saved file must contain xl/worksheets/sheet1.xml");
    assert_xml_semantic_eq(original_sheet, saved_sheet).unwrap();

    let original_styles = original_pkg
        .part("xl/styles.xml")
        .expect("fixture must contain xl/styles.xml");
    let saved_styles = saved_pkg
        .part("xl/styles.xml")
        .expect("saved file must contain xl/styles.xml");
    assert_xml_semantic_eq(original_styles, saved_styles).unwrap();

    let saved_sheet_xml = std::str::from_utf8(saved_sheet).expect("sheet1.xml should be utf-8");
    assert!(
        saved_sheet_xml.contains("78C0D931-6437-407d-A8EE-F0AAD7539E65")
            || saved_sheet_xml.contains("78C0D931-6437-407D-A8EE-F0AAD7539E65"),
        "expected x14 conditional formatting extLst entry to be preserved; sheet1.xml: {saved_sheet_xml}"
    );
    assert!(
        saved_sheet_xml.contains("x14:negativeFillColor"),
        "expected x14:negativeFillColor to be preserved; sheet1.xml: {saved_sheet_xml}"
    );
    assert!(
        saved_sheet_xml.contains("x14:axisColor"),
        "expected x14:axisColor to be preserved; sheet1.xml: {saved_sheet_xml}"
    );
}

