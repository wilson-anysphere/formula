use std::path::Path;

#[test]
fn reads_workbook_date_system_1904() {
    let fixture = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/metadata/date-system-1904.xlsx"
    ));

    let doc = formula_xlsx::load_from_path(fixture).expect("load xlsx fixture");
    assert_eq!(doc.workbook.date_system, formula_model::DateSystem::Excel1904);
}
