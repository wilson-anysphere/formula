use formula_engine::calc_settings::{CalcSettings, CalculationMode};
use formula_engine::{Engine, SheetLifecycleError, Value};
use formula_model::{CellRef, Range, Table, TableColumn};
use pretty_assertions::assert_eq;

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
}

fn table_with_sheet_refs() -> Table {
    let formula = "Sheet1!A1+'[Book.xlsx]Sheet1'!A1".to_string();
    Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: range("A1:B3"),
        header_row_count: 1,
        totals_row_count: 1,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".to_string(),
                formula: Some(formula.clone()),
                totals_formula: Some(formula),
            },
            TableColumn {
                id: 2,
                name: "Col2".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

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

#[test]
fn delete_sheet_rewrites_refs_when_display_name_differs_from_stable_key() {
    let mut engine = Engine::new();

    // Create sheets with stable keys that differ from their user-visible display names.
    engine.ensure_sheet("sheet1_key");
    engine.ensure_sheet("sheet2_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");
    engine.set_sheet_display_name("sheet2_key", "Sheet2");

    engine
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A1")
        .unwrap();

    engine.delete_sheet("Sheet1").unwrap();

    // Deleting a sheet should rewrite local references using the *display name* to `#REF!`
    // (Excel semantics), even when the stable key differs.
    assert_eq!(engine.get_cell_formula("Sheet2", "A1"), Some("=#REF!"));

    // Recreating a new sheet with the same display name must not resurrect the reference.
    engine.ensure_sheet("sheet1_new");
    engine.set_sheet_display_name("sheet1_new", "Sheet1");
    assert_eq!(engine.get_cell_formula("Sheet2", "A1"), Some("=#REF!"));
}

#[test]
fn delete_sheet_matches_sheet_names_case_insensitively_across_unicode() {
    // Engine sheet lookups + delete rewrite logic should match Excel: sheet names compare using
    // NFKC + Unicode uppercasing (e.g. `ß` -> `SS`).
    let mut engine = Engine::new();

    engine.ensure_sheet("Straße");
    engine.ensure_sheet("Other");

    engine
        .set_cell_formula("Other", "A1", "='Straße'!A1")
        .unwrap();

    // Delete using a Unicode-case-equivalent spelling of the sheet name.
    engine.delete_sheet("STRASSE").unwrap();

    assert_eq!(engine.get_cell_formula("Other", "A1"), Some("=#REF!"));
}

#[test]
fn delete_sheet_matches_sheet_names_nfkc_case_insensitively_and_adjusts_3d_spans() {
    // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'.
    // When deleting a 3D span boundary, Excel shifts it one sheet inward based on tab order.
    let mut engine = Engine::new();

    engine.ensure_sheet("Kelvin");
    engine.ensure_sheet("Middle");
    engine.ensure_sheet("Sheet3");
    engine.ensure_sheet("Summary");

    engine
        .set_cell_formula("Summary", "A1", "=SUM(KELVIN:Sheet3!A1)")
        .unwrap();

    engine.delete_sheet("KELVIN").unwrap();

    assert_eq!(
        engine.get_cell_formula("Summary", "A1"),
        Some("=SUM(Middle:Sheet3!A1)")
    );
}

#[test]
fn rename_sheet_matches_old_name_case_insensitively_across_unicode() {
    // Renames should accept sheet names in a Unicode-case-insensitive form (Excel-like).
    let mut engine = Engine::new();

    engine.ensure_sheet("Straße");
    engine.ensure_sheet("Other");

    // Reference the sheet using a different spelling that should still resolve (`ß` -> `SS`).
    engine
        .set_cell_formula("Other", "A1", "=STRASSE!A1")
        .unwrap();

    assert!(engine.rename_sheet("STRASSE", "Renamed"));

    assert_eq!(engine.get_cell_formula("Other", "A1"), Some("=Renamed!A1"));
}

#[test]
fn delete_sheet_adjusts_3d_boundaries_when_display_name_differs_from_stable_key() {
    let mut engine = Engine::new();

    for (key, display) in [
        ("sheet1_key", "Sheet1"),
        ("sheet2_key", "Sheet2"),
        ("sheet3_key", "Sheet3"),
    ] {
        engine.ensure_sheet(key);
        engine.set_sheet_display_name(key, display);
    }

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "B1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(6.0));

    engine.delete_sheet("Sheet1").unwrap();

    // If a deleted sheet was a 3D boundary, Excel shifts the boundary inward by one sheet.
    assert_eq!(
        engine.get_cell_formula("Sheet2", "B1"),
        Some("=SUM(Sheet2:Sheet3!A1)")
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(5.0));
}

#[test]
fn delete_sheet_adjusts_mixed_3d_boundaries_with_stable_keys_and_display_names() {
    let mut engine = Engine::new();

    for (key, display) in [
        ("sheet1_key", "Sheet1"),
        ("sheet2_key", "Sheet2"),
        ("sheet3_key", "Sheet3"),
    ] {
        engine.ensure_sheet(key);
        engine.set_sheet_display_name(key, display);
    }

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    // Mixed 3D span boundaries: start uses display name, end uses stable key.
    engine
        .set_cell_formula("Sheet2", "B1", "=SUM(Sheet1:sheet3_key!A1)")
        .unwrap();
    // Mixed 3D span boundaries: start uses stable key, end uses display name.
    engine
        .set_cell_formula("Sheet2", "B2", "=SUM(sheet1_key:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B2"), Value::Number(6.0));

    engine.delete_sheet("Sheet1").unwrap();
    assert_eq!(
        engine.get_cell_formula("Sheet2", "B1"),
        Some("=SUM(Sheet2:sheet3_key!A1)")
    );
    assert_eq!(
        engine.get_cell_formula("Sheet2", "B2"),
        Some("=SUM(sheet2_key:Sheet3!A1)")
    );
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B2"), Value::Number(5.0));
}

#[test]
fn delete_sheet_rewrites_references_that_mix_stable_key_and_display_name_spellings() {
    let mut engine = Engine::new();

    engine.ensure_sheet("sheet1_key");
    engine.ensure_sheet("sheet2_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");
    engine.set_sheet_display_name("sheet2_key", "Sheet2");

    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", "=sheet1_key!A1+Sheet1!A1")
        .unwrap();

    engine.delete_sheet("Sheet1").unwrap();
    assert_eq!(
        engine.get_cell_formula("Sheet2", "A1"),
        Some("=#REF!+#REF!")
    );
}

#[test]
fn rename_sheet_rewrites_table_column_formulas_but_not_external_refs() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.set_sheet_tables("Sheet2", vec![table_with_sheet_refs()]);

    assert!(engine.rename_sheet("Sheet1", "Renamed"));

    let tables = engine.sheet_tables("Sheet2").unwrap();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];

    assert_eq!(
        table.columns[0].formula.as_deref(),
        Some("Renamed!A1+'[Book.xlsx]Sheet1'!A1")
    );
    assert_eq!(
        table.columns[0].totals_formula.as_deref(),
        Some("Renamed!A1+'[Book.xlsx]Sheet1'!A1")
    );
}

#[test]
fn rename_sheet_marks_formulatext_dependents_dirty_in_manual_mode() {
    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Manual,
        ..CalcSettings::default()
    });

    engine.set_cell_value("Sheet1", "A1", 5.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A1")
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "B1", "=FORMULATEXT(A1)")
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet2", "B1"),
        Value::Text("=Sheet1!A1".to_string())
    );

    assert!(engine.rename_sheet("Sheet1", "Renamed"));

    // Manual mode: rename should mark FORMULATEXT cells dirty; values update on next recalc tick.
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet2", "B1"),
        Value::Text("=Renamed!A1".to_string())
    );
}

#[test]
fn rename_sheet_refreshes_formulatext_dependents_in_automatic_mode() {
    let mut engine = Engine::new();
    engine.set_calc_settings(CalcSettings {
        calculation_mode: CalculationMode::Automatic,
        ..CalcSettings::default()
    });
    engine.set_cell_value("Sheet1", "A1", 5.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A1")
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "B1", "=FORMULATEXT(A1)")
        .unwrap();

    assert_eq!(
        engine.get_cell_value("Sheet2", "B1"),
        Value::Text("=Sheet1!A1".to_string())
    );

    assert!(engine.rename_sheet("Sheet1", "Renamed"));

    // Automatic mode: rename triggers a recalc tick, so FORMULATEXT updates immediately.
    assert_eq!(
        engine.get_cell_value("Sheet2", "B1"),
        Value::Text("=Renamed!A1".to_string())
    );
}

#[test]
fn delete_sheet_invalidates_table_column_formulas_but_not_external_refs() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.set_sheet_tables("Sheet2", vec![table_with_sheet_refs()]);

    engine.delete_sheet("Sheet1").unwrap();

    let tables = engine.sheet_tables("Sheet2").unwrap();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];

    assert_eq!(
        table.columns[0].formula.as_deref(),
        Some("#REF!+[Book.xlsx]Sheet1!A1")
    );
    assert_eq!(
        table.columns[0].totals_formula.as_deref(),
        Some("#REF!+[Book.xlsx]Sheet1!A1")
    );
}

#[test]
fn delete_sheet_adjusts_table_3d_boundary() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let formula = "SUM(Sheet1:Sheet3!A1)".to_string();
    let table = Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: Range::new(cell("A1"), cell("A3")),
        header_row_count: 1,
        totals_row_count: 1,
        columns: vec![TableColumn {
            id: 1,
            name: "Col1".to_string(),
            formula: Some(formula.clone()),
            totals_formula: Some(formula),
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    };

    engine.set_sheet_tables("Sheet3", vec![table]);

    engine.delete_sheet("Sheet1").unwrap();

    let tables = engine.sheet_tables("Sheet3").unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(
        tables[0].columns[0].formula.as_deref(),
        Some("SUM(Sheet2:Sheet3!A1)")
    );
}

#[test]
fn delete_sheet_refuses_to_delete_last_sheet() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Only");

    assert!(engine.delete_sheet("Only").is_err());
    assert!(engine.sheet_id("Only").is_some());
    assert_eq!(engine.sheet_ids_in_order().len(), 1);
}

#[test]
fn delete_sheet_succeeds_when_multiple_sheets_exist() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    let sheet2 = engine.sheet_id("Sheet2").unwrap();

    assert!(engine.delete_sheet("Sheet1").is_ok());
    assert_eq!(engine.sheet_id("Sheet1"), None);
    assert_eq!(engine.sheet_ids_in_order(), vec![sheet2]);
}

#[test]
fn delete_sheet_by_id_refuses_to_delete_last_sheet() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Only");
    let id = engine.sheet_id("Only").unwrap();

    assert!(matches!(
        engine.delete_sheet_by_id(id),
        Err(SheetLifecycleError::CannotDeleteLastSheet)
    ));
    assert!(engine.sheet_id("Only").is_some());
    assert_eq!(engine.sheet_ids_in_order().len(), 1);
}

#[test]
fn delete_sheet_by_id_succeeds_when_multiple_sheets_exist() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let sheet2 = engine.sheet_id("Sheet2").unwrap();
    assert!(engine.delete_sheet_by_id(sheet2).is_ok());
    assert_eq!(engine.sheet_id("Sheet2"), None);
    assert_eq!(
        engine.sheet_names_in_order(),
        vec!["Sheet1".to_string(), "Sheet3".to_string()]
    );
}
