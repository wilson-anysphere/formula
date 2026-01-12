use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::Workbook;

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/charts/chart-sheet.xlsx");

fn build_regenerated_workbook() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Category").unwrap();
    worksheet.write_string(0, 1, "Value").unwrap();

    let categories = ["A", "B", "C"];
    let values = [10.0, 20.0, 30.0];

    for (i, (cat, val)) in categories.iter().zip(values).enumerate() {
        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, *cat).unwrap();
        worksheet.write_number(row, 1, val).unwrap();
    }

    workbook.save_to_buffer().unwrap()
}

#[test]
fn preserved_chart_sheets_can_be_reapplied_to_regenerated_workbook() {
    let source = XlsxPackage::from_bytes(FIXTURE).expect("load chartsheet fixture");
    let preserved = source
        .preserve_drawing_parts()
        .expect("preserve drawing parts");
    assert!(
        !preserved.chart_sheets.is_empty(),
        "expected preserved payload to include chart sheet metadata"
    );

    let regenerated_bytes = build_regenerated_workbook();
    let mut dest = XlsxPackage::from_bytes(&regenerated_bytes).expect("load regenerated workbook");

    dest.apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved parts");

    let merged_bytes = dest.write_to_bytes().expect("write merged workbook");
    let merged = XlsxPackage::from_bytes(&merged_bytes).expect("parse merged workbook");

    assert!(
        merged.part("xl/chartsheets/sheet1.xml").is_some(),
        "expected chart sheet part to be present"
    );

    let workbook_xml = std::str::from_utf8(merged.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(
        workbook_xml.contains("name=\"Chart1\""),
        "expected workbook.xml to contain chartsheet entry"
    );

    let workbook_rels =
        std::str::from_utf8(merged.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(
        workbook_rels.contains("relationships/chartsheet"),
        "expected workbook rels to contain chartsheet relationship"
    );
    assert!(
        workbook_rels.contains("Target=\"chartsheets/sheet1.xml\""),
        "expected workbook rels to target chartsheet part"
    );
}
