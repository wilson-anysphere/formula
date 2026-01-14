use formula_engine::Engine;

#[test]
fn delete_sheet_rewrites_local_refs_but_not_external_workbook_refs() {
    let mut engine = Engine::new();

    // Ensure both sheets exist so `Sheet1!A1` is compiled as an internal sheet reference.
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    engine
        .set_cell_formula("Sheet2", "A1", "=[Book.xlsx]Sheet1!A1+Sheet1!A1")
        .unwrap();

    engine.delete_sheet("Sheet1").unwrap();

    // Local references to the deleted sheet should become `#REF!`, but external workbook refs with
    // the same sheet name must remain intact.
    assert_eq!(
        engine.get_cell_formula("Sheet2", "A1").unwrap(),
        "=[Book.xlsx]Sheet1!A1+#REF!"
    );
}

