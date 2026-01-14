use formula_engine::{Engine, Value};

#[test]
fn cell_address_uses_sheet_display_name_after_rename() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    engine.set_workbook_file_metadata(None, Some("Book1.xlsx"));
    engine.set_sheet_display_name("Sheet1", "Budget");

    engine
        .set_cell_formula("Sheet2", "A1", r#"=CELL("address",Sheet1!A1)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "B1", r#"=CELL("filename",Sheet1!A1)"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Text("Budget!$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "B1"),
        Value::Text("[Book1.xlsx]Budget".to_string())
    );

    engine.set_sheet_display_name("Sheet1", "Data");
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Text("Data!$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "B1"),
        Value::Text("[Book1.xlsx]Data".to_string())
    );
}
