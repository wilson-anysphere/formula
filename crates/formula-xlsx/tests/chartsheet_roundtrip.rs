use formula_xlsx::{load_from_bytes, XlsxPackage};

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/charts/chart-sheet.xlsx");

const CHARTSHEET_PARTS: &[&str] = &[
    "xl/chartsheets/sheet1.xml",
    "xl/chartsheets/_rels/sheet1.xml.rels",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
    "xl/charts/chart1.xml",
    "xl/_rels/workbook.xml.rels",
    "[Content_Types].xml",
];

#[test]
fn chartsheet_roundtrip_preserves_chartsheet_related_parts_byte_for_byte() {
    let doc = load_from_bytes(FIXTURE).expect("load chartsheet fixture");
    let roundtripped_bytes = doc.save_to_vec().expect("roundtrip save");

    let original = XlsxPackage::from_bytes(FIXTURE).expect("parse original");
    let roundtripped = XlsxPackage::from_bytes(&roundtripped_bytes).expect("parse roundtripped");

    for part in CHARTSHEET_PARTS {
        let original_bytes = original
            .part(part)
            .unwrap_or_else(|| panic!("original missing part {part}"));
        let roundtripped_bytes = roundtripped
            .part(part)
            .unwrap_or_else(|| panic!("roundtripped missing part {part}"));
        assert_eq!(
            roundtripped_bytes,
            original_bytes,
            "part {part} changed during no-op round-trip"
        );
    }
}

